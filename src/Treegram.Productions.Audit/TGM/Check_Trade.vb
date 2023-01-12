Imports M4D.Production.Core
Imports M4D.Treegram.Core.Entities
Imports M4D.Treegram.Core.Kernel
Imports M4D.Treegram.Core.Extensions.Kernel
Imports M4D.Treegram.Core.Extensions.Entities
Imports DevExpress.Spreadsheet
Imports Treegram.Bam.Libraries.AuditFile

'Public Class Check_Trade_RVT
'    Inherits ProdScript
'    Public Sub New()
'        MyBase.New("RVT :: Check : Trade") 'Useless ProdAction, to replace with "TGM :: Check : Trade"
'        AddAction(New Check_Trade(Nothing))
'    End Sub
'End Class
'Public Class Check_Trade_IFC
'    Inherits ProdScript
'    Public Sub New()
'        MyBase.New("IFC :: Check : Trade") 'Useless ProdAction, to replace with "TGM :: Check : Trade"
'        AddAction(New Check_Trade(Nothing))
'    End Sub
'End Class
'Public Class Check_TradeScript
'    Inherits ProdScript
'    Public Sub New()
'        MyBase.New("TGM :: Check : Trade")
'        AddAction(New Check_Trade(Nothing))
'    End Sub
'End Class


Public Class Check_TradeARC
    Inherits ProdScript
    Public Sub New()
        MyBase.New("TGM :: Check : Trade ARC")
        AddAction(New Check_Trade() With {._MyTrade = Trade.ARC})
    End Sub
End Class
Public Class Check_TradeFAC
    Inherits ProdScript
    Public Sub New()
        MyBase.New("TGM :: Check : Trade FAC")
        AddAction(New Check_Trade() With {._MyTrade = Trade.FAC})
    End Sub
End Class
Public Class Check_TradeSTR
    Inherits ProdScript
    Public Sub New()
        MyBase.New("TGM :: Check : Trade STR")
        AddAction(New Check_Trade() With {._MyTrade = Trade.STR})
    End Sub
End Class
Public Class Check_TradeFLU
    Inherits ProdScript
    Public Sub New()
        MyBase.New("TGM :: Check : Trade FLU")
        AddAction(New Check_Trade() With {._MyTrade = Trade.FLU})
    End Sub
End Class


Public Enum Trade
    ARC
    FAC
    STR
    FLU
End Enum


Public Class Check_Trade
    Inherits ProdAction

    Public _MyTrade As Trade?
    Public Sub New()
        Name = "TGM :: Check : Trade (ProdAction)"
        PartOfScript = True
    End Sub

    Public Overrides Sub CreateGuiInputTree()
        InputTree = TempWorkspace.AddTree("--- Temporary Input Tree ---")
        Dim fileNode As Node = InputTree.AddNode("object.type", "File")
        InputTree.SmartAddNode("object.type", "File").Description = "Files"

        SelectYourSetsOfInput.Add("Fichiers", {fileNode}.ToList)
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
            Return New IgnoredActionResult("No Input File")
        End If
        ' Récupération des inputs
        Dim tgmFiles = GetInputAsMetaObjects(Fichiers.Source).ToHashSet 'To get rid of duplicates

        'CHECK
        CheckMetier(tgmFiles(0))


        Return New SucceededActionResult("Succeeded", System.Drawing.Color.Green, New List(Of Element) From {Fichiers})
    End Function


    Private Shared TradesDico As New Dictionary(Of String, String) From {{"ARC", "Architecture"}, {"FAC", "Facade"}, {"STR", "Structure"}, {"FLU", "Fluides"}}


    Private Function CheckMetier(fileTgm As MetaObject)
        Dim scanWs As Workspace = fileTgm.Container
        Dim myAudit As New FileObject(fileTgm)

        'Récupération du classeur d'audit
        Dim auditWb = myAudit.AuditWorkbook
        If auditWb Is Nothing Then Throw New Exception("Launch initial audit first")
        Dim tradeWsheet As Worksheet
        If myAudit.ExtensionType = "ifc" Then
            tradeWsheet = myAudit.InsertProdActionTemplateInAuditWbook("ProdActions_Template", "APPARTENANCE METIER IFC")
        Else
            tradeWsheet = myAudit.InsertProdActionTemplateInAuditWbook("ProdActions_Template", "APPARTENANCE METIER REVIT")
        End If

        Dim typeColumn As Integer = 1
        Dim bimshakerSettings = fileTgm.GetAttribute("BimShakerSettings")
        If IsNothing(bimshakerSettings.GetAttribute("Trade")) Then
            bimshakerSettings.SmartAddAttribute("Trade", _MyTrade.ToString)
        End If

        'Discipline de la maquette
        Dim trade As String
        If _MyTrade IsNot Nothing Then
            trade = TradesDico(_MyTrade.ToString)
        ElseIf fileTgm.GetAttribute("BimShakerSettings") IsNot Nothing Then 'Bientôt obsolète !!!!!!!!!!!!!!!!!
            trade = fileTgm.GetAttributes("Trade", True, False).FirstOrDefault.Value
        Else
            Throw New Exception("You must specify with which Trade you want to analyse your file")
        End If

        Dim tradeRange = tradeWsheet.Range(trade)
        Dim tradeColumn = tradeRange.LeftColumnIndex

        Dim lastTrRange = tradeWsheet.GetUsedRange()
        Dim lastTrRow = lastTrRange.BottomRowIndex

        Dim tradeList = New List(Of String) From {{"Architecture"}, {"Facade"}, {"Structure"}, {"Fluides"}}
        tradeList.Remove(trade)

        For Each tradeToHide In tradeList
            Dim rangeToHide = tradeWsheet.Range(tradeToHide)
            rangeToHide.ColumnWidth = 0
        Next

        'Catégories à vérifier
        Dim toCheckDico As New Dictionary(Of Integer, List(Of (categoryTxt As String, categoryStatus As String)))
        Dim existingN As Boolean = False
        For i = 4 To lastTrRow
            Dim categoryName = tradeWsheet.Cells(i, typeColumn).Value.ToString
            Dim checking = tradeWsheet.Cells(i, tradeColumn).Value.ToString

            If Not String.IsNullOrWhiteSpace(checking) Then
                Dim myList As New List(Of (categoryTxt As String, categoryStatus As String))

                If checking.ToLower = "x" Then
                    myList.Add((categoryName, "Expected"))
                    toCheckDico.Add(i, myList)
                Else
                    existingN = True
                    myList.Add((categoryName, "Necessary"))
                    toCheckDico.Add(i, myList)
                End If

            End If
        Next

        'Récupération de la feuille stats
        Dim statWsheet As DevExpress.Spreadsheet.Worksheet
        Dim statColumn As Integer
        Dim objCountColumn As Integer

        If myAudit.ExtensionType = "ifc" Then
            'Try
            statWsheet = auditWb.Worksheets.Item("TYPES IFC")
            'Catch ex As Exception
            '    Throw New Exception("Launch initial audit first")
            'End Try
            statColumn = 1
            objCountColumn = 3
        Else
            statWsheet = auditWb.Worksheets.Item("FAMILLES")
            Dim statRange = tradeWsheet.Range("STATS")
            statColumn = statRange.LeftColumnIndex
            objCountColumn = 4
        End If

        Dim lastStRange = statWsheet.GetUsedRange()
        Dim lastStRow = lastStRange.BottomRowIndex


        'Catégories présentes dans la maquette
        Dim categObj = scanWs.MetaObjects.Item(4)
        Dim categList As New List(Of MetaObject)
        categList = categObj.GetChildren.ToList
        Dim categoryCount As New Dictionary(Of String, Integer)
        For Each categ In categList
            Dim localCategory = categ.Name

            If Not String.IsNullOrWhiteSpace(localCategory) Then
                Dim myReq = From obj In scanWs.MetaObjects Where obj.IsActive AndAlso Not IsNothing(obj?.GetAttribute("object.containedIn")) AndAlso obj?.GetAttribute("object.containedIn").Value = "3D" AndAlso Not IsNothing(obj?.GetAttribute("RevitCategory")) AndAlso obj.GetAttribute("RevitCategory").Value = localCategory
                Dim myReq2 = From obj In scanWs.MetaObjects Where obj.IsActive AndAlso Not IsNothing(obj?.GetAttribute("RevitCategory")) AndAlso obj.GetAttribute("RevitCategory").Value = localCategory
                Dim localCount = myReq.Count 'Convert.ToInt32(statWsheet.Cells(j, objCountColumn).Value.ToString)
                categoryCount.Add(localCategory, localCount)
            End If

        Next

        'Total objets maquette
        Dim allObjects As Integer = Convert.ToInt32(statWsheet.Cells(4, objCountColumn).Value.ToString)

        'Tourner sur toCheckDico pour remplir nbre d'objets et ratio
        Dim countingObj As Integer = 0
        Dim failedCondition As Boolean

        For h = 0 To toCheckDico.Count - 1
            Dim rightRow = toCheckDico.ElementAt(h).Key
            Dim myList = toCheckDico.ElementAt(h).Value

            Dim txt = myList.FirstOrDefault.categoryTxt
            Dim status = myList.FirstOrDefault.categoryStatus

            'part du principe que la catégorie nécessaire est absente de l'audit jusqu'à preuve du contraire
            Dim necessaryBool As Boolean = False

            For Each auditElement In categoryCount
                If auditElement.Key = txt Then
                    'nb
                    tradeWsheet.Cells(rightRow, 6).Value = auditElement.Value
                    'ratio
                    tradeWsheet.Cells(rightRow, 7).Value = (auditElement.Value / allObjects)
                    tradeWsheet.Cells(rightRow, 7).NumberFormat = "0.0%"

                    'mise en page
                    For l = typeColumn To 7
                        tradeWsheet.Cells(rightRow, l).FillColor = Drawing.Color.Gainsboro
                    Next

                    'indicateurs
                    countingObj += auditElement.Value
                    necessaryBool = True

                    Exit For

                End If
            Next

            If existingN AndAlso status = "Necessary" AndAlso necessaryBool = False Then
                failedCondition = True

                'surligner la ligne en orange
                For g = typeColumn To 8
                    tradeWsheet.Cells(rightRow, g).FillColor = Drawing.Color.Orange
                Next
            End If

        Next

        'Les totaux à la fin
        Dim totalRange = tradeWsheet.Range("Total_" + trade)
        Dim totalColumn = totalRange.LeftColumnIndex
        Dim totalRow = totalRange.BottomRowIndex

        Dim endingValue = countingObj / allObjects
        tradeWsheet.Cells(totalRow + 1, totalColumn).Value = endingValue
        tradeWsheet.Cells(totalRow + 1, totalColumn).NumberFormat = "0.0%"

        If failedCondition Then
            tradeWsheet.Cells(totalRow + 2, totalColumn).Value = "False"
        ElseIf existingN AndAlso failedCondition = False Then
            tradeWsheet.Cells(totalRow + 2, totalColumn).Value = "True"
        End If


        'Critère
        Dim endingValuePER = endingValue * 100

        If failedCondition Then
            myAudit.ReportInfos.appropriateTrade = False
        Else
            If endingValuePER < 70% Then
                myAudit.ReportInfos.appropriateTrade = False
            Else
                myAudit.ReportInfos.appropriateTrade = True

            End If
        End If

        myAudit.ReportInfos.CompleteCriteria(67, True)


        myAudit.ProjWs.PushAllModifiedEntities()
        myAudit.AuditWorkbook.Calculate()
        myAudit.AuditWorkbook.Worksheets.ActiveWorksheet = myAudit.ReportWorksheet
        myAudit.SaveAuditWorkbook()

    End Function
End Class
