Imports M4D.Production.Core
Imports M4D.Treegram.Core.Entities
Imports M4D.Treegram.Core.Kernel
Imports M4D.Treegram.Core.Extensions.Entities
Imports M4D.Treegram.Core.Extensions.Kernel
Imports Treegram.Bam.Libraries.AuditFile
Imports DevExpress.Spreadsheet
Imports DevExpress.Export.Xl
Imports Treegram.Bam.Libraries.AEC
Imports Treegram.Bam.Functions.AEC
Imports TGM.Deixi.Services
Imports M4D.Treegram.Core.Extensions.Models
Imports Treegram.GeomFunctions

Public Class Analyse_ColumnsProdScript
    Inherits ProdScript
    Public Sub New()
        MyBase.New("IFC :: Analyse : Columns")
        AddAction(New Analyse_Columns())
    End Sub
End Class
Public Class Analyse_Columns
    Inherits ProdAction
    Public Sub New()
        Name = "IFC :: Analyse : Columns (ProdAction)"
        PartOfScript = True
    End Sub
    'Private Shared TempWorkspace As Workspace = Workspace.GetTemp

    Public Overrides Sub CreateGuiInputTree()
        InputTree = TempWorkspace.AddTree("--- Temporary Input Tree ---")
        Dim objectNode = InputTree.SmartAddNode("object.containedIn", "3D", , , "3D Objects")
        Dim columnNode = objectNode.SmartAddNode("object.type", "IfcColumn")
        SelectYourSetsOfInput.Add("Objects3D", {columnNode}.ToList())
    End Sub

    Public Overrides Function GuiUpdateLauncherTree(inputs As Dictionary(Of String, IEnumerable(Of PersistEntity))) As Tree
        Dim launchTree As Tree = TempWorkspace.SmartAddTree("LauncherTree")
        launchTree.DuplicateObjectsWhileFiltering = False
        CommonAecLauncherTrees.FilesLaunchTree(launchTree, inputs)
        Return launchTree
    End Function

    <ActionMethod()>
    Public Function MyMethod(Objects3D As MultipleElements) As ActionResult

        ' Récupération des inputs
        Dim tgmObjsList = GetInputAsMetaObjects(Objects3D.Source).ToHashSet 'To get rid of duplicates
        If tgmObjsList.Count = 0 Then Throw New Exception("Inputs missing")

        'AUDIT
        Try
            OutputTree = GetColumnAnalysisFromIfc(tgmObjsList.ToList)
        Catch ex As Exception
            Return New FailedActionResult(ex.Message)
        End Try

#Region "VIEW CREATION"
        Dim scanWs As Workspace = tgmObjsList(0).Container

        Dim chargedWs As New List(Of Workspace) From {scanWs}
        Dim checkedNodes As New List(Of Node)
        'Dim savedColors As New Dictionary(Of Node, Tuple(Of System.Windows.Media.Color, Single))
        Dim colorNodes As New List(Of NodeColorTransparency)

        'Récuperation du file
        Dim fileTgm As MetaObject = scanWs.GetMetaObjects(, "File").FirstOrDefault

        Dim myTree = OutputTree
        Dim fileNode = myTree.SmartAddNode("File", fileTgm.Name)
        Dim columnNode = fileNode.Nodes(0)
        Dim lengthNode = columnNode.Nodes.Where(Function(o) o.Name = "ColumnAnalysisLength").FirstOrDefault
        Dim ratioNode = columnNode.Nodes.Where(Function(o) o.Name = "ColumnAnalysisRatio").FirstOrDefault

        checkedNodes.AddRange({fileNode, columnNode, lengthNode, ratioNode})

        'savedColors.Add(fileNode, New Tuple(Of System.Windows.Media.Color, Single)(System.Windows.Media.Color.FromRgb(191, 191, 191), 0.1F)) 'grey 10%
        'savedColors.Add(columnNode, New Tuple(Of Windows.Media.Color, Single)(Windows.Media.Color.FromRgb(182, 250, 182), 1.0F)) 'green 100%
        'savedColors.Add(lengthNode, New Tuple(Of Windows.Media.Color, Single)(Windows.Media.Color.FromRgb(229, 117, 114), 1.0F)) 'red 100%
        'savedColors.Add(ratioNode, New Tuple(Of Windows.Media.Color, Single)(Windows.Media.Color.FromRgb(165, 137, 199), 1.0F)) 'purple 100%

        colorNodes.Add(New NodeColorTransparency(fileNode, System.Windows.Media.Color.FromRgb(191, 191, 191), 0.1F))
        colorNodes.Add(New NodeColorTransparency(columnNode, System.Windows.Media.Color.FromRgb(182, 250, 182), 1F))
        colorNodes.Add(New NodeColorTransparency(lengthNode, System.Windows.Media.Color.FromRgb(229, 117, 114), 1F))
        colorNodes.Add(New NodeColorTransparency(ratioNode, System.Windows.Media.Color.FromRgb(165, 137, 199), 1F))

        TgmView.CreateView(fileTgm.GetProject, fileTgm.Name & " - METIER POTEAUX", chargedWs, myTree, checkedNodes, colorNodes)
#End Region

        Return New SucceededActionResult("Succeeded", System.Drawing.Color.Green, New List(Of Element) From {Objects3D})
    End Function

    Private Function GetColumnAnalysisFromIfc(tgmObjList As List(Of MetaObject)) As Tree

        Dim myFile = CommonAecFunctions.GetAuditIfcFileFromInputs(tgmObjList)
        If myFile.AuditWorkbook Is Nothing Then Create_IfcInitialAudit.GetAuditFromFile(myFile)
        Dim analysisWsheetName = "POTEAUX"
        Dim colAnaSettingsWsheet = myFile.GetOrInsertSettingsWSheet(analysisWsheetName)
        If colAnaSettingsWsheet Is Nothing Then Throw New Exception("Could not find """ + analysisWsheetName + """ worksheet in settings workbook")


        myFile.ReportInfos.ColumnGeoCriteria_columnWidthLimit = CDbl(colAnaSettingsWsheet.Range("largeurmax").Value.NumericValue)
        myFile.ReportInfos.ColumnGeoCriteria_columnRatioLimit = 5.0
        Try
            myFile.ReportInfos.ColumnGeoCriteria_columnRatioLimit = CDbl(colAnaSettingsWsheet.Range("ratiomax").Value.NumericValue)
        Catch ex As Exception
        End Try

        'FILTER AND LOAD - NEW WAY TO LOAD ADDITIONAL GEOM
        Dim listToLoad As New List(Of Workspace) From {myFile.ScanWs}
        listToLoad.AddRange(myFile.ScanWs.GetExtensionWorspaces(M4D.Treegram.Core.Extensions.Enums.ExtensionType.Analysis))
        listToLoad.AddRange(myFile.ScanWs.GetExtensionWorspaces(M4D.Treegram.Core.Extensions.Enums.ExtensionType.AdditionalGeometries))
        TempWorkspace.SmartAddTree("ALL").Filter(False, listToLoad).RunSynchronously()
        Treegram.GeomFunctions.Models.Reset3dDictionnary()
        Treegram.GeomFunctions.Models.LoadGeometryModels(tgmObjList, True, True)

#Region "ANALYSE"
        Dim maxLengthDico As New Dictionary(Of String, List(Of (objTgm As MetaObject, maxLength As Double, ratio As Double)))
        Dim maxLengthList As New List(Of (objTgm As MetaObject, maxLength As Double, ratio As Double))
        Dim aecColumns = tgmObjList.Select(Function(o) New AecObject(o)).ToList

        '---EXCLUDE OBJECTS WITHOUT GEOMETRY
        aecColumns = CommonAecFunctions.ExcludeNoGeomObjects(aecColumns).ToList

        '---LOCALPLACEMENT ANNALYSIS
        For Each aecCol In aecColumns
            CommonAecGeomFunctions.SmartAddLocalPlacement(aecCol, myFile)
        Next

        '---PROFILE ANALYSIS
        Dim objProfiles = CommonAecFunctions.SmartAddOrientedBBoxProfile(aecColumns, myFile)


        'DIMENSIONS ANALYSIS
        For Each columnProfile In objProfiles
            If columnProfile.Value Is Nothing OrElse columnProfile.Value.Points.Count = 0 Then
                maxLengthList.Add((columnProfile.Key.Metaobject, Nothing, Nothing))
                Continue For
            End If
            Dim minLength = Double.PositiveInfinity 'Min length pas fiable si le poteau a une forme spéciale !!
            Dim maxLength = Double.NegativeInfinity
            For Each line In columnProfile.Value.ToBasicModelerBoundary
                If line.Length > maxLength Then maxLength = line.Length
                If line.Length < minLength Then minLength = line.Length
            Next
            Dim ratioLgLarg = Math.Round(CDbl(maxLength / minLength), 3)
            AecObject.CompleteTgmPset(columnProfile.Key.Metaobject, "ColumnAnalysisLength", maxLength) 'Pour visualisation dans Tgm
            AecObject.CompleteTgmPset(columnProfile.Key.Metaobject, "ColumnAnalysisRatio", ratioLgLarg) 'Pour visualisation dans Tgm
            maxLengthList.Add((columnProfile.Key.Metaobject, maxLength, ratioLgLarg))
        Next
        maxLengthDico.Add("IfcColumn", maxLengthList)
#End Region

#Region "FILL EXCEL"
        Dim columnWsheet As Worksheet = myFile.InsertProdActionTemplateInAuditWbook("ProdActions_Template", analysisWsheetName)
        Dim typeStartLine = 5
        columnWsheet.Cells(typeStartLine - 1, 1).Value = maxLengthDico.Keys.Count
        columnWsheet.Cells(typeStartLine - 1, 2).Value = maxLengthDico.SelectMany(Function(o) o.Value).Count
        columnWsheet.Cells(typeStartLine - 2, 3).Value = "Largeur < " + myFile.ReportInfos.ColumnGeoCriteria_columnWidthLimit.ToString + "m :"
        columnWsheet.Cells(typeStartLine - 2, 4).Value = "Long/Larg < " + myFile.ReportInfos.ColumnGeoCriteria_columnRatioLimit.ToString + " :"

        'FILL ARRAY
        Dim indice_next_line = 0
        Dim indice_layer_start_line As Integer
        Dim totalRows = maxLengthDico.SelectMany(Function(o) o.Value).Count + maxLengthDico.Keys.Count

        '---Prepare array
        Dim arrayLinesNb = totalRows
        Dim arrayColsNb = 3 ' <<<<-----<<<<<---- ADD 1 TO ADD A CRITERIA

        For Each typeSt In maxLengthDico.Keys
            Dim columnAnalysis = maxLengthDico(typeSt)

            'Fill TYPES
            columnWsheet.Range("B" + (typeStartLine + indice_next_line + 1).ToString).Value = typeSt

            '---Mise en forme
            Dim typeRange = columnWsheet.Range("B" + (typeStartLine + indice_next_line + 1).ToString + ":" + FileObject.GetSpreadsheetColumnName(arrayColsNb + 1) + (typeStartLine + indice_next_line + 1).ToString)
            typeRange.FillColor = XlColor.FromTheme(XlThemeColor.Dark1, 0.7)
            typeRange.Font.Bold = True
            typeRange.Borders.SetOutsideBorders(System.Drawing.Color.Black, BorderLineStyle.Thin)
            typeRange.Borders.InsideVerticalBorders.LineStyle = BorderLineStyle.Thin
            typeRange.Borders.InsideHorizontalBorders.LineStyle = BorderLineStyle.Thin

            indice_next_line += 1
            indice_layer_start_line = indice_next_line
            Dim badColumnGeoCounter = 0
            For Each colAna In columnAnalysis
                Dim trustedColumn = True
                columnWsheet.Range("C" + (typeStartLine + indice_next_line + 1).ToString).Value = colAna.objTgm.GetAttribute("GlobalId").Value.ToString

                '<<<<<-------FILL CRITERIAS HERE !!!!------->>>>>>
                If colAna.maxLength > myFile.ReportInfos.ColumnGeoCriteria_columnWidthLimit Then
                    columnWsheet.Range("D" + (typeStartLine + indice_next_line + 1).ToString).Value = "FAUX"
                    trustedColumn = False
                Else
                    columnWsheet.Range("D" + (typeStartLine + indice_next_line + 1).ToString).Value = "VRAI"
                End If
                If colAna.ratio > myFile.ReportInfos.ColumnGeoCriteria_columnRatioLimit Then
                    columnWsheet.Range("E" + (typeStartLine + indice_next_line + 1).ToString).Value = "FAUX"
                    trustedColumn = False
                Else
                    columnWsheet.Range("E" + (typeStartLine + indice_next_line + 1).ToString).Value = "VRAI"
                End If
                '<<<<<-------FIN------->>>>>>
                If Not trustedColumn Then badColumnGeoCounter += 1
                indice_next_line += 1
            Next
            myFile.ReportInfos.ColumnGeoCriteria_badColumnGeoCount = badColumnGeoCounter

            '---Mise en forme
            If indice_next_line <> indice_layer_start_line Then
                Dim namesRange = columnWsheet.Range("B" + (typeStartLine + indice_layer_start_line + 1).ToString + ":" + FileObject.GetSpreadsheetColumnName(arrayColsNb + 1) + (typeStartLine + indice_next_line).ToString)
                namesRange.Borders.SetOutsideBorders(System.Drawing.Color.Black, BorderLineStyle.Thin)
                namesRange.Borders.InsideVerticalBorders.LineStyle = BorderLineStyle.Thin
                namesRange.GroupRows(False)
            End If
        Next

        'FILL EXCEL
        columnWsheet.Columns("C:" + FileObject.GetSpreadsheetColumnName(arrayColsNb + 1)).AutoFit()

#End Region

#Region "CREATE TREE"
        Dim myTree = myFile.AuditTreesWs.SmartAddTree("Audit - METIER POTEAUX")
        Dim fileNode = myTree.SmartAddNode("File", myFile.FileTgm.Name)
        myTree.RemoveNode(fileNode)
        fileNode = myTree.SmartAddNode("File", myFile.FileTgm.Name)

        For Each typeSt In maxLengthDico.Keys
            Dim typeNode = fileNode.SmartAddNode("object.type", typeSt)
            Dim columnAnalysis = maxLengthDico(typeSt)

            typeNode.SmartAddNode("ColumnAnalysisLength", myFile.ReportInfos.ColumnGeoCriteria_columnWidthLimit, AttributeType.Double, NodeOperator.Greater, "Lg > " + myFile.ReportInfos.ColumnGeoCriteria_columnWidthLimit.ToString + "m")
            typeNode.SmartAddNode("ColumnAnalysisRatio", myFile.ReportInfos.ColumnGeoCriteria_columnRatioLimit, AttributeType.Double, NodeOperator.Greater, "Lg/Larg > " + myFile.ReportInfos.ColumnGeoCriteria_columnRatioLimit.ToString)
        Next
#End Region

        'FILL REPORT
        myFile.ReportInfos.CompleteCriteria(61)

        myFile.ProjWs.PushAllModifiedEntities()
        myFile.AuditWorkbook.Calculate()
        myFile.AuditWorkbook.Worksheets.ActiveWorksheet = myFile.ReportWorksheet
        myFile.SaveAuditWorkbook()

        Return myTree

    End Function
End Class
