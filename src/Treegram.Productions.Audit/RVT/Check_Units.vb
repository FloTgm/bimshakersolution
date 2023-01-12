Imports M4D.Production.Core
Imports M4D.Treegram.Core.Entities
Imports M4D.Treegram.Core.Kernel
Imports M4D.Treegram.Core.Extensions.Kernel
Imports M4D.Treegram.Core.Extensions.Entities
Imports Treegram.Bam.Libraries.AuditFile
Imports DevExpress.Spreadsheet
Imports M4D.Treegram.Dashboard.Widgets.Welcome.BimShacker.AuditWindow

Public Class CheckUnitsProdScript
    Inherits ProdScript
    Public Sub New()
        MyBase.New("RVT :: Check : Units")
        AddAction(New CheckUnits())
    End Sub
End Class
Public Class CheckUnits

    Inherits ProdAction
    Public Sub New()
        Name = "RVT :: Check : Units (Prodaction)"
        PartOfScript = True
    End Sub

    ''' <summary>
    ''' Méthode appelée lors de la sélection de l'action
    ''' Elle n'est pas obligatoire
    ''' </summary>
    Public Overrides Sub OnActionSelected()

    End Sub

    ''' <summary>
    ''' Si jamais vous avez besoin de créer un arbre input, vous pouvez utiliser cette méthode.
    ''' Il n'est pas obligatoire de la rajouter
    ''' </summary>
    Public Overrides Sub CreateGuiInputTree()
        InputTree = TempWorkspace.AddTree("--- Temporary Input Tree ---")
        Dim fileNode As Node = InputTree.AddNode("object.type", "File")
        'InputTree.AddNode("object.type", "File").Description = "Files"

        Dim mandatName As String = ""
        SelectYourSetsOfInput.Add("FichiersRevit", {fileNode}.ToList())
        'SelectYourSetsOfInput.Add("MandatoryName", {mandatName}.ToList())

    End Sub

    ''' <summary>
    ''' Initialise l'Output Tree
    ''' Pas obligatoire non plus
    ''' </summary>
    Public Overrides Sub CreateGuiOutputTree()
        'OutputTree = TempWorkspace.AddTree("Arbre output temporaire")
    End Sub

    ''' <param name="inputs">La collection d'un paramètre et de ses PersistEntity(pour le moment uniquement des metaObjects</param>
    ''' <returns></returns>
    Public Overrides Function GuiUpdateLauncherTree(inputs As Dictionary(Of String, IEnumerable(Of PersistEntity))) As Tree

        Dim launchTree = TempWorkspace.SmartAddTree("LauncherTree")

        ' Suppression des anciens noeuds
        While launchTree.Nodes.Count > 0
            launchTree.RemoveNode(launchTree.Nodes.First)
        End While

        If inputs.Values.SelectMany(Function(i) i).ToList.Count > 0 Then
            For Each inputObj In inputs.Values.SelectMany(Function(i) i).ToList
                Dim inputmObj As MetaObject = inputObj
                'POUR UN FILTRAGE DE TYPE INFORMATIQUE : WORKSPACE, FILE, METAOBJET
                If inputmObj.GetTgmType <> "File" OrElse inputmObj.GetAttribute("Extension").Value.ToString.ToLower <> "rvt" Then Continue For 'Throw New Exception("You must drag and drop file objects")
                launchTree.SmartAddNode("object.id", inputmObj.Id.ToString,,, inputmObj.Name)
            Next
        End If

        Return launchTree
    End Function
    <ActionMethod()>
    Public Function MyMethod(FichiersRevit As MultipleElements) As ActionResult
        ' Récupération des inputs
        If FichiersRevit.Source.Count = 0 Then
            Return New IgnoredActionResult("No input")
        End If
        Dim tgmFiles = GetInputAsMetaObjects(FichiersRevit.Source).ToHashSet 'To get rid of duplicates
        Dim sortedDict As New Dictionary(Of Tuple(Of String, String), Dictionary(Of Tuple(Of String, String), List(Of Tuple(Of String, String))))
        'AUDIT
        For Each fileTgm In tgmFiles
            OutputTree = checkUnits(fileTgm, sortedDict)
        Next

        Return New SucceededActionResult("Succeeded", System.Drawing.Color.Green, New List(Of Element) From {FichiersRevit})

    End Function

    Private Function checkUnits(RevitFileTgmObj As MetaObject, ByRef sortedDict As Dictionary(Of Tuple(Of String, String), Dictionary(Of Tuple(Of String, String), List(Of Tuple(Of String, String))))) As Tree
        Dim myAudit As New FileObject(RevitFileTgmObj)
        Dim auditWb = myAudit.AuditWorkbook

        Dim unitsAtt = myAudit.FileTgm.GetAttribute("Units")

        Dim fichierWsheet As Worksheet = auditWb.Worksheets.Item("METRIQUES")
        Dim unitsStartLine = fichierWsheet.Range("UNITS").BottomRowIndex
        Dim unitsCol = fichierWsheet.Range("UNITS").LeftColumnIndex

        Dim k As Integer
        myAudit.ReportInfos.uncomparedUnits = New List(Of String)
        myAudit.ReportInfos.invalidUnits = New List(Of String)
        myAudit.ReportInfos.validUnits = New List(Of String)

        'Get References

        For k = 0 To unitsAtt.Attributes.Count - 1
            Dim unitAtt = unitsAtt.Attributes(k)
            Dim myUnitRow As Integer
            Dim maxNameRange As CellRange = Nothing
            Dim maxValueRange As CellRange = Nothing
            Dim valueRange, refRange, comparisonRange As CellRange

            'LENGTH
            If unitAtt.Name.ToLower = "length" Or unitAtt.Name.ToLower = "longueur" Then
                myUnitRow = unitsStartLine + 2
                'Value
                valueRange = fichierWsheet.Cells(myUnitRow, unitsCol + 1)
                valueRange.Value = unitAtt.Value?.ToString
                valueRange.Font.Italic = False

                Dim refLength As String = String.Empty
                If myAudit.ReportInfos.GetMaxRanges(maxNameRange, maxValueRange, "refLength") Then refLength = maxValueRange.Value.TextValue
                If String.IsNullOrEmpty(refLength) Then
                    myAudit.ReportInfos.uncomparedUnits.Add("longueur")
                Else
                    'Ref
                    refRange = fichierWsheet.Cells(myUnitRow, unitsCol + 2)
                    refRange.Value = refLength
                    refRange.Font.Italic = False

                    'Comparison
                    comparisonRange = fichierWsheet.Cells(myUnitRow, unitsCol + 3)
                    If unitAtt.Value.ToLower.Replace("tre", "ter") <> refLength.ToLower.Replace("tre", "ter") Then 'English/French
                        myAudit.ReportInfos.invalidUnits.Add("longueur")
                        comparisonRange.Value = "NOK"
                    Else
                        myAudit.ReportInfos.validUnits.Add("longueur")
                        comparisonRange.Value = "OK"
                    End If
                End If

                'ANGLE
            ElseIf unitAtt.Name.ToLower = "angle" Then
                myUnitRow = unitsStartLine + 5
                'Value
                valueRange = fichierWsheet.Cells(myUnitRow, unitsCol + 1)
                valueRange.Value = unitAtt.Value?.ToString
                valueRange.Font.Italic = False

                Dim refAngle As String = String.Empty
                If myAudit.ReportInfos.GetMaxRanges(maxNameRange, maxValueRange, "refAngle") Then refAngle = maxValueRange.Value.TextValue
                If String.IsNullOrEmpty(refAngle) Then
                    myAudit.ReportInfos.uncomparedUnits.Add("angle")
                Else
                    'Ref
                    refRange = fichierWsheet.Cells(myUnitRow, unitsCol + 2)
                    refRange.Value = refAngle
                    refRange.Font.Italic = False

                    'Comparison
                    comparisonRange = fichierWsheet.Cells(myUnitRow, unitsCol + 3)
                    If unitAtt.Value.ToLower <> refAngle.ToLower Then
                        myAudit.ReportInfos.invalidUnits.Add("angle")
                        comparisonRange.Value = "NOK"
                    Else
                        myAudit.ReportInfos.validUnits.Add("angle")
                        comparisonRange.Value = "OK"
                    End If
                End If

                'AREA
            ElseIf unitAtt.Name.ToLower = "area" Or unitAtt.Name.ToLower = "aire" Then
                myUnitRow = unitsStartLine + 3
                'Value
                valueRange = fichierWsheet.Cells(myUnitRow, unitsCol + 1)
                valueRange.Value = unitAtt.Value?.ToString
                valueRange.Font.Italic = False

                Dim refSurf As String = String.Empty
                If myAudit.ReportInfos.GetMaxRanges(maxNameRange, maxValueRange, "refSurf") Then refSurf = maxValueRange.Value.TextValue
                If String.IsNullOrEmpty(refSurf) Then
                    myAudit.ReportInfos.uncomparedUnits.Add("aire")
                Else
                    'Ref
                    refRange = fichierWsheet.Cells(myUnitRow, unitsCol + 2)
                    refRange.Value = refSurf
                    refRange.Font.Italic = False

                    'Comparison
                    comparisonRange = fichierWsheet.Cells(myUnitRow, unitsCol + 3)
                    If unitAtt.Value.ToLower.Replace("tre", "ter") <> refSurf.ToLower.Replace("tre", "ter") Then 'English/French
                        myAudit.ReportInfos.invalidUnits.Add("aire")
                        comparisonRange.Value = "NOK"
                    Else
                        myAudit.ReportInfos.validUnits.Add("aire")
                        comparisonRange.Value = "OK"
                    End If
                End If

                'VOLUME
            ElseIf unitAtt.Name.ToLower = "volume" Then
                myUnitRow = unitsStartLine + 4
                'Value
                valueRange = fichierWsheet.Cells(myUnitRow, unitsCol + 1)
                valueRange.Value = unitAtt.Value?.ToString
                valueRange.Font.Italic = False

                Dim refVol As String = String.Empty
                If myAudit.ReportInfos.GetMaxRanges(maxNameRange, maxValueRange, "refVol") Then refVol = maxValueRange.Value.TextValue
                If String.IsNullOrEmpty(refVol) Then
                    myAudit.ReportInfos.uncomparedUnits.Add("volume")
                Else
                    'Ref
                    refRange = fichierWsheet.Cells(myUnitRow, unitsCol + 2)
                    refRange.Value = refVol
                    refRange.Font.Italic = False

                    'Comparison
                    comparisonRange = fichierWsheet.Cells(myUnitRow, unitsCol + 3)
                    If unitAtt.Value.ToLower.Replace("tre", "ter") <> refVol.ToLower.Replace("tre", "ter") Then 'English/French
                        myAudit.ReportInfos.invalidUnits.Add("volume")
                        comparisonRange.Value = "NOK"
                    Else
                        myAudit.ReportInfos.validUnits.Add("volume")
                        comparisonRange.Value = "OK"
                    End If
                End If

            Else
                Continue For
            End If
        Next
        myAudit.ReportInfos.CompleteCriteria(54, True)
        myAudit.SaveAuditWorkbook()

        Return Nothing
    End Function
End Class
