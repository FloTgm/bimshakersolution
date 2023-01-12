Imports M4D.Production.Core
Imports M4D.Treegram.Core.Entities
Imports M4D.Treegram.Core.Kernel
Imports M4D.Treegram.Core.Extensions.Kernel
Imports M4D.Treegram.Core.Extensions.Entities
Imports DevExpress.Spreadsheet
Imports Treegram.Bam.Libraries.AuditFile

Public Class Check_WarningsScript
    Inherits ProdScript
    Public Sub New()
        MyBase.New("RVT :: Check : Warnings")
        AddAction(New Check_Warnings())
    End Sub
End Class
Public Class Check_Warnings
    Inherits ProdAction
    Public Sub New()
        Name = "RVT :: Check : Warnings (ProdAction)"
        PartOfScript = True
    End Sub

    Public Overrides Sub CreateGuiInputTree()
        InputTree = TempWorkspace.AddTree("--- Temporary Input Tree ---")
        Dim fileNode As Node = InputTree.AddNode("object.type", "File")
        InputTree.SmartAddNode("object.type", "File").Description = "Files"

        SelectYourSetsOfInput.Add("Fichiers", {fileNode}.ToList())

    End Sub

    Public Overrides Function GuiUpdateLauncherTree(inputs As Dictionary(Of String, IEnumerable(Of PersistEntity))) As Tree

        Dim launchTree = TempWorkspace.SmartAddTree("LauncherTree")

        ' Suppression des anciens noeuds
        While launchTree.Nodes.Count > 0
            launchTree.RemoveNode(launchTree.Nodes.First)
        End While

        If inputs.Values.SelectMany(Function(i) i).ToList.Count > 0 Then
            For Each inputObj In inputs.Values.SelectMany(Function(i) i).ToList
                Dim inputmObj As MetaObject = inputObj
                If inputmObj.GetTgmType <> "File" Then Continue For 'Throw New Exception("You must drag and drop file objects")
                launchTree.SmartAddNode("object.id", inputmObj.Id.ToString,,, inputmObj.Name)
            Next
        End If

        Return launchTree
    End Function

    <ActionMethod()>
    Public Function MyMethod(Fichiers As MultipleElements) As ActionResult
        If Fichiers.Source.Count = 0 Then
            Return New IgnoredActionResult("No input")
        End If
        ' Récupération des inputs
        Dim tgmFiles = GetInputAsMetaObjects(Fichiers.Source).ToHashSet 'To get rid of duplicates
        If tgmFiles.Count = 0 Then Throw New Exception("Inputs missing")

        'CHECK
        CheckWarning(tgmFiles(0))

        Return New SucceededActionResult("Succeeded", System.Drawing.Color.Green, New List(Of Element) From {Fichiers})
    End Function

    Private Function CheckWarning(fileTgm As MetaObject)
        Dim scanWs As Workspace = fileTgm.Container
        Dim myAudit As New FileObject(fileTgm)

        'Recupération du classeur warnings
        Dim warningWsheet = myAudit.GetOrInsertSettingsWSheet("AVERTISSEMENTS")

        'Récupération de la liste des warnings
        'Selon deux catégories : rules or extensive
        Dim extensiveDico As New Dictionary(Of Integer, List(Of (warningTxt As String, warningVisa As String)))
        Dim rulesToCheck As New List(Of Double)

        Dim i = 0
        While warningWsheet.Cells(i, 0).Value.ToString <> ""
            'détecter les règles en place dans l'onglet settings
            If warningWsheet.Cells(i, 1).Value.ToString.Contains("Par analyse") Then
                Dim namedRange = warningWsheet.Cells(i, 0).Name
                Dim numberRange = namedRange.Replace("Rule", "")
                Dim numberDB = Convert.ToDouble(numberRange)
                rulesToCheck.Add(numberDB)

            ElseIf warningWsheet.Cells(i, 1).Value.ToString.Contains("Avertissement précis") Then
                Dim myList As New List(Of (warningTxt As String, warningVisa As String))
                myList.Add((warningWsheet.Cells(i, 0).Value.ToString, warningWsheet.Cells(i, 2).Value.ToString))
                extensiveDico.Add(i, myList)

            End If
            i += 1
        End While

        'Récupération du classeur d'audit
        Dim auditWb = myAudit.AuditWorkbook
        If auditWb Is Nothing Then Throw New Exception("Launch initial audit first")
        Dim auditWsheet As DevExpress.Spreadsheet.Worksheet = auditWb.Worksheets.Item("AVERTISSEMENTS")
        If auditWsheet Is Nothing Then Throw New Exception("Worksheet ""AVERTISSEMENTS"" missing in audit workbook")

        'On récupère les avertissements de l'audit
        Dim wngLine As Integer = 2
        Dim wngCol = 1
        Dim auditList As New Dictionary(Of String, Integer)
        Dim j = 0
        While auditWsheet.Cells(wngLine, wngCol + j).Value.ToString <> ""
            auditList.Add(auditWsheet.Cells(wngLine, wngCol + j).Value.ToString, wngCol + j)
            j += 1
        End While

        'Comparaison des avertissements
        Dim statutWarning As New Dictionary(Of String, String)

        For Each localWarning In auditList

            'Rules
            For Each ruleNB In rulesToCheck
                Dim myStatut = CheckRule(localWarning.Key, ruleNB)
                If Not String.IsNullOrWhiteSpace(myStatut) Then
                    Dim myColor = MatchColor(myStatut)
                    auditWsheet.Cells(2, localWarning.Value).Font.Color = myColor
                    'If myStatut = "Préoccupant" Then
                    '    Dim badWarningCount = auditWsheet.Cells(4, localWarning.Value).Value.ToString
                    '    Dim badWarningDCount As Double = Convert.ToDouble(badWarningCount)
                    '    myAudit.ReportInfos.nbBadWarnings += badWarningDCount
                    'End If
                End If
            Next


            Dim myCount = auditWsheet.Cells(4, localWarning.Value).Value.ToString
            If myCount = "" Then
                GoTo nextWarning
            End If
            Dim doubleCount As Double = Convert.ToDouble(myCount)

            'Avertissements précis : écrasent les Rules
            For Each extensiveW In extensiveDico
                Dim txt = extensiveW.Value.FirstOrDefault.warningTxt
                Dim visa = extensiveW.Value.FirstOrDefault.warningVisa

                Dim restant As String = FindCommonText(localWarning.Key, txt, 0, localWarning.Key.Count - 1, 0, txt.Count - 1)

                If doubleCount = 1 Then
                    'warning exact
                    If restant = txt Then
                        Dim myColor = MatchColor(visa)
                        auditWsheet.Cells(2, localWarning.Value).Font.Color = myColor
                        'If visa = "Préoccupant" Then
                        '    Dim badWarningCount = auditWsheet.Cells(4, localWarning.Value).Value.ToString
                        '    Dim badWarningDCount As Double = Convert.ToDouble(badWarningCount)
                        '    myAudit.ReportInfos.nbBadWarnings += badWarningDCount
                        'End If

                        GoTo nextWarning
                    End If


                ElseIf doubleCount > 1 Then
                    'warning avec "..."
                    If localWarning.Key = txt Then
                        Dim myColor = MatchColor(visa)
                        auditWsheet.Cells(2, localWarning.Value).Font.Color = myColor
                        'If visa = "Préoccupant" Then
                        '    Dim badWarningCount = auditWsheet.Cells(4, localWarning.Value).Value.ToString
                        '    Dim badWarningDCount As Double = Convert.ToDouble(badWarningCount)
                        '    myAudit.ReportInfos.nbBadWarnings += badWarningDCount
                        'End If

                        GoTo nextWarning
                    End If

                End If
            Next
nextWarning:
        Next

        'on repasse sur l'audit pour récupérer le count de mauvais avertissements
        Dim c = 0
        While auditWsheet.Cells(wngLine, wngCol + c).Value.ToString <> ""
            If auditWsheet.Cells(wngLine, wngCol + c).Font.Color = Drawing.Color.Red Then
                Dim badWarningCount = auditWsheet.Cells(4, wngCol + c).Value.ToString
                Dim badWarningDCount As Double = Convert.ToDouble(badWarningCount)
                myAudit.ReportInfos.nbBadWarnings += badWarningDCount
            End If
            c += 1
        End While

        'LEGENDE
        Dim richText As RichTextString = New RichTextString()

        richText.AddTextRun("PEUT ETRE IGNORE;", New RichTextRunFont("Calibri", 11, Drawing.Color.Green))
        richText.AddTextRun("A CORRIGER;", New RichTextRunFont("Calibri", 11, Drawing.Color.Orange))
        richText.AddTextRun("PREOCCUPANT;", New RichTextRunFont("Calibri", 11, Drawing.Color.Red))
        richText.AddTextRun("NSP", New RichTextRunFont("Calibri", 11, Drawing.Color.Black))

        auditWsheet.Cells(0, 2).SetRichText(richText)

        'Critère
        myAudit.ReportInfos.CompleteCriteria(28, True)


        'Dim avertissementWS = getwork
        myAudit.ProjWs.PushAllModifiedEntities()
        myAudit.AuditWorkbook.Calculate()
        myAudit.AuditWorkbook.Worksheets.ActiveWorksheet = myAudit.ReportWorksheet
        myAudit.SaveAuditWorkbook()

    End Function

    Private Function MatchColor(statutW As String) As Drawing.Color

        Dim colorTh As Drawing.Color

        If statutW = "Peut être ignoré" Then
            colorTh = Drawing.Color.Green
        ElseIf statutW = "A corriger" Then
            colorTh = Drawing.Color.Orange
        ElseIf statutW = "Préoccupant" Then
            colorTh = Drawing.Color.Red
        End If

        Return colorTh
    End Function
    Private Function CheckRule(oWarning As String, ruleNumber As Double) As String
        Select Case ruleNumber
            Case 1
                If oWarning.Contains("Le léger décalage de la ligne par rapport à l'axe") AndAlso oWarning.Contains("quadrillage") = False Then
                    Return "Peut être ignoré"
                End If

            Case 2
                If oWarning.Contains("Le léger décalage de la ligne par rapport à l'axe") AndAlso oWarning.Contains("quadrillage") Then
                    Return "Préoccupant"
                End If

            Case 3
                If (oWarning.Contains("se chevauchent") OrElse oWarning.Contains("se chevauche")) AndAlso oWarning.Contains("pièce") = False Then
                    Return "A corriger"
                End If

            Case 4
                If (oWarning.Contains("se chevauchent") OrElse oWarning.Contains("se chevauche")) AndAlso oWarning.Contains("pièce") Then
                    Return "Préoccupant"
                End If

            Case 5
                If oWarning.Contains("doublon") OrElse oWarning.Contains("dupliqué") OrElse oWarning.Contains("occurrences identiques") Then
                    Return "Préoccupant"
                End If

            Case 6
                If oWarning.Contains("conflit") AndAlso (oWarning.Contains("pièce") = False OrElse oWarning.Contains("surface") = False) Then
                    Return "A corriger"
                End If

            Case 7
                If oWarning.Contains("conflit") AndAlso (oWarning.Contains("pièce") OrElse oWarning.Contains("surface")) Then
                    Return "Préoccupant"
                End If

            Case Else
                Return ""
        End Select

    End Function
    Public Shared Function FindCommonText(ByVal string1 As String, ByVal string2 As String, st1 As Integer, end1 As Integer, st2 As Integer, end2 As Integer) As String
        Dim lString1 = string1.ToLower
        Dim lstring2 = string2.ToLower

        Dim res = ""
        If Not (st1 >= end1 OrElse st2 >= end2 OrElse st1 < 0 OrElse st2 < 0 OrElse end1 < 0 OrElse end2 < 0 OrElse (end1 - st1) <= 3 OrElse (end2 - st2) <= 3) Then
            Dim ns1 = 0
            Dim ns2 = 0
            Dim i As Integer
            Dim max = 0

            For c1 = st1 To end1
                For c2 = st2 To end2
                    i = 0
                    Do Until lString1(c1 + i) <> lstring2(c2 + i)
                        i = i + 1
                        If i > max Then
                            ns1 = c1
                            ns2 = c2
                            max = i
                        End If
                        If c1 + i > end1 Or c2 + i > end2 Then Exit Do
                    Loop
                Next c2
            Next c1
            res = string1.Substring(ns1, max)
            If max = 0 Then
                Return res
            End If
            If res.Count <= 3 Then
                res = ""
            Else
                If ns1 > 0 Then
                    res = "..." + res
                End If
                If ns1 + max < end1 Then
                    res = res + "..."
                End If
            End If
            res = FindCommonText(string1, string2, st1, ns1 - 1, st2, ns2 - 1) + res + FindCommonText(string1, string2, ns1 + max, end1, ns2 + max, end2)
            res = res.Replace("......", "...")
        End If
        Return res
    End Function

End Class
