Imports M4D.Production.Core
Imports M4D.Treegram.Core.Entities
Imports M4D.Treegram.Core.Kernel
Imports M4D.Treegram.Core.Extensions.Entities
Imports M4D.Treegram.Core.Extensions.Kernel
Imports DevExpress.Spreadsheet
Imports DevExpress.Export.Xl
Imports Treegram.Bam.Libraries.AEC
Imports Treegram.Bam.Functions.AEC
Imports M4D.Treegram.Core.Extensions.Models
Imports Treegram.Bam.Functions
Imports Treegram.Bam.Libraries.AuditFile
Imports TGM.Deixi.Services

Public Class Analyse_WallsProdScript
    Inherits ProdScript
    Public Sub New()
        MyBase.New("IFC :: Analyse : Walls")
        AddAction(New Analyse_Walls() With {.Name = Name, .PartOfScript = True})
    End Sub
End Class
Public Class Analyse_Walls
    Inherits ProdAction
    Public Sub New()
    End Sub

    Public Overrides Sub CreateGuiInputTree()
        'Création d'une zone de filtrage (arbre) temporaire
        InputTree = TempWorkspace.AddTree("--- Temporary Input Tree ---")

        'Création d'un filtre : murs
        Dim objectNode = InputTree.SmartAddNode("object.containedIn", "3D",,, "3D Objects")
        Dim wallNode = objectNode.SmartAddNode("object.type", "IfcWall",,, DicoObjType("IfcWall"))

        'Drag and Drop automatique des objets filtrés vers l'algo (Optionnel)
        SelectYourSetsOfInput.Add("Murs", {wallNode}.ToList())
    End Sub

    Public Overrides Function GuiUpdateLauncherTree(inputs As Dictionary(Of String, IEnumerable(Of PersistEntity))) As Tree
        'Séparation des lancements par fichier
        Dim launchTree As Tree = TempWorkspace.SmartAddTree("LauncherTree")
        launchTree.DuplicateObjectsWhileFiltering = False
        CommonAecLauncherTrees.FilesLaunchTree(launchTree, inputs)
        If launchTree.Nodes.Count = 0 Then
            launchTree.AddNode("", Nothing)
        End If
        Return launchTree
    End Function

    <ActionMethod()>
    Public Function MyMethod(Murs As MultipleElements) As ActionResult

        'PREPARATION ENVIRONNEMENT TGM
        '--Sortie de l'algo s'il n'y a pas d'inputs
        If Murs.Source.Count = 0 Then
            Return New IgnoredActionResult("No input")
        End If

        '--Récupération des métaobjets murs sous formé de liste
        Dim tgmObjsList = GetInputAsMetaObjects(Murs.Source).ToHashSet.ToList 'To get rid of duplicates

        '--Récupération du fichier source
        Dim myFile = CommonAecFunctions.GetAuditFileFromInputs(tgmObjsList)

        '--Chargement des données BAMs
        Dim listToLoad As New List(Of Workspace) From {myFile.ScanWs}
        listToLoad.AddRange(myFile.ScanWs.GetExtensionWorspaces(M4D.Treegram.Core.Extensions.Enums.ExtensionType.Analysis))
        listToLoad.AddRange(myFile.ScanWs.GetExtensionWorspaces(M4D.Treegram.Core.Extensions.Enums.ExtensionType.AdditionalGeometries))
        CommonFunctions.RunFilter(TempWorkspace, listToLoad)
        Treegram.GeomFunctions.Models.Reset3dDictionnary()
        Treegram.GeomFunctions.Models.LoadGeometryModels(tgmObjsList, True, True)


        'ANALYSE MÉTIER
        Dim analysisWsheetName = "MURS"
        If myFile.AuditWorkbook Is Nothing Then
            Return New FailedActionResult("No Audit found")
        End If
        Dim vertSepSettingsWsheet = myFile.GetOrInsertSettingsWSheet(analysisWsheetName)
        If vertSepSettingsWsheet Is Nothing Then Throw New Exception("Could not find """ + analysisWsheetName + """ worksheet in settings workbook")
        Dim minThick, maxThick As Double
        If vertSepSettingsWsheet.DefinedNames.Contains("murepmin") Then 'New settings
            Try
                minThick = CDbl(vertSepSettingsWsheet.Range("murepmin").Value.NumericValue) 'm
                maxThick = CDbl(vertSepSettingsWsheet.Range("murepmax").Value.NumericValue) 'm
            Catch ex As Exception
                Return New FailedActionResult("Wrong settings values, try to modify them or delete 'MURS' excel tab and launch this script again")
            End Try
        Else 'Old settings
            Try
                minThick = CDbl(vertSepSettingsWsheet.Range("sepvertepmin").Value.NumericValue) 'm
                maxThick = CDbl(vertSepSettingsWsheet.Range("sepvertepmax").Value.NumericValue) 'm
            Catch ex As Exception
                Return New FailedActionResult("Wrong settings values, try to modify them or delete 'MURS' excel tab and launch this script again")
            End Try
        End If


#Region "ANALYSE"
        Dim aecWalls = tgmObjsList.Select(Function(o) New AecWall(o)).ToList

        '---EXCLUDE OBJECTS WITHOUT GEOMETRY
        aecWalls = CommonAecFunctions.ExcludeNoGeomObjects(aecWalls).Cast(Of AecWall).ToList

        '---LOCALPLACEMENT ANALYSIS
        For Each aecWall In aecWalls
            CommonAecGeomFunctions.SmartAddLocalPlacement(aecWall, myFile)
        Next

        '---GET AXIS
        Dim wallAxis = CommonAecFunctions.SmartAddAxis(aecWalls.AsEnumerable, myFile)

        '---PROFILE ANALYSIS
        Dim anaObjProfiles = CommonAecFunctions.SmartAddOrientedBBoxProfile(aecWalls.AsEnumerable, myFile)

        '---GET DIMENSIONS
        Dim noneDbl = -1.0
        Dim thicknessDico As New Dictionary(Of Double, List(Of (objTgm As AecWall, length As Double)))
        For Each aecWall In aecWalls
            aecWall.CompleteTgmPset("CategoryByTgm", Category.Wall.ToString)
            aecWall.CompleteTgmPset("IsVerticalSeparator", True)

            If aecWall.VectorU.IsZero Or Not wallAxis.ContainsKey(aecWall) Then
                'aecWall.CompleteTgmPset("IsVerticalSeparator", False)
                aecWall.CompleteTgmPset("LodByTgm", 0)

                'Add to dico
                If thicknessDico.ContainsKey(noneDbl) Then
                    thicknessDico(noneDbl).Add((aecWall, 0.0))
                Else
                    thicknessDico.Add(noneDbl, New List(Of (objTgm As AecWall, length As Double)) From {(aecWall, 0.0)})
                End If


            Else
                'thickness
                Dim thick As Double = Math.Round(aecWall.MaxV - aecWall.MinV, 5) 'rounded to 10-5 m
                aecWall.CompleteTgmPset("ThicknessByTgm", thick)
                'length
                Dim length As Double = Math.Round(aecWall.MaxU - aecWall.MinU, 5) 'rounded to 10-5 m
                aecWall.CompleteTgmPset("LengthByTgm", length)
                'height
                Dim minMaxZ = aecWall.SmartAddExtremumZ
                aecWall.CompleteTgmPset("HeightByTgm", Math.Round(minMaxZ.Item2 - minMaxZ.Item1, 5))

                If thick < minThick OrElse thick > maxThick Then
                    'aecWall.CompleteTgmPset("IsVerticalSeparator", False)
                    aecWall.CompleteTgmPset("LodByTgm", 0)
                Else
                    'aecWall.CompleteTgmPset("IsVerticalSeparator", True)
                    aecWall.CompleteTgmPset("LodByTgm", 100)
                End If

                'Add to dico
                Dim length_cm = Math.Round(length, 2) 'rounded to cm
                Dim thick_mm = Math.Round(thick, 3)
                aecWall.CompleteTgmPset("ThicknessByTgm_mm", thick_mm)

                If thicknessDico.ContainsKey(thick_mm) Then
                    thicknessDico(thick_mm).Add((aecWall, length_cm))
                Else
                    thicknessDico.Add(thick_mm, New List(Of (objTgm As AecWall, length As Double)) From {(aecWall, length_cm)})
                End If
            End If
        Next
#End Region

#Region "FILL EXCEL"
        Dim vertSepWsheet As Worksheet = myFile.InsertProdActionTemplateInAuditWbook("ProdActions_Template", analysisWsheetName)
        Dim typeStartLine = 6
        Dim typeStartCol = 4

        'FILL ARRAY
        Dim indice_next_line = 0
        Dim indice_layer_start_line As Integer
        Dim arrayColsNb = 4 ' <<<<-----<<<<<---- ADD 1 TO ADD A CRITERIA

        'FILL TYPES
        vertSepWsheet.Cells(typeStartLine + indice_next_line, typeStartCol).Value = ""
        vertSepWsheet.Cells(typeStartLine + indice_next_line, typeStartCol + 2).Value = thicknessDico.SelectMany(Function(o) o.Value).Count

        '---Mise en forme
        Dim typeRange = vertSepWsheet.Cells(typeStartLine + indice_next_line, typeStartCol).Resize(1, arrayColsNb)
        typeRange.FillColor = XlColor.FromTheme(XlThemeColor.Dark1, 0.7)
        typeRange.Font.Bold = True
        typeRange.Borders.SetOutsideBorders(System.Drawing.Color.Black, BorderLineStyle.Thin)
        typeRange.Borders.InsideVerticalBorders.LineStyle = BorderLineStyle.Thin
        typeRange.Borders.InsideHorizontalBorders.LineStyle = BorderLineStyle.Thin

        indice_next_line += 1
        indice_layer_start_line = indice_next_line

        'FILL THICKNESSES
        Dim sortedThickness = thicknessDico.Keys.ToList
        sortedThickness.Sort()
        Dim nbFalse As Integer = 0
        For Each thickness In sortedThickness
            Dim thickObjs = thicknessDico(thickness)
            If thickness = noneDbl Then
                vertSepWsheet.Cells(typeStartLine + indice_next_line, typeStartCol + 1).Value = "<NONE>"
            Else
                vertSepWsheet.Cells(typeStartLine + indice_next_line, typeStartCol + 1).Value = (thickness * 100)
            End If
            vertSepWsheet.Cells(typeStartLine + indice_next_line, typeStartCol + 2).Value = thickObjs.Count
            vertSepWsheet.Cells(typeStartLine + indice_next_line, typeStartCol + 3).Value = Math.Round(thickObjs.Select(Function(o) o.length).Sum, 1)
            'Set in red wrong thicknesses
            If thickness < minThick Or thickness > maxThick Then
                vertSepWsheet.Cells(typeStartLine + indice_next_line, typeStartCol + 1).Resize(1, 3).Font.Color = System.Drawing.Color.Red
                nbFalse = nbFalse + thickObjs.Count
            End If

            indice_next_line += 1
        Next

        '---Mise en forme
        Dim formatRange As CellRange = Nothing
        If indice_next_line <> indice_layer_start_line Then
            Dim namesRange = vertSepWsheet.Cells(typeStartLine + indice_layer_start_line, typeStartCol).Resize(indice_next_line - indice_layer_start_line, arrayColsNb)
            namesRange.Borders.SetOutsideBorders(System.Drawing.Color.Black, BorderLineStyle.Thin)
            namesRange.Borders.InsideVerticalBorders.LineStyle = BorderLineStyle.Thin
            namesRange.GroupRows(False)

            'colorFormatRanges.Add(namesRange)
            Dim lengthRange = vertSepWsheet.Cells(namesRange.TopRowIndex, namesRange.RightColumnIndex).Resize(namesRange.BottomRowIndex - namesRange.TopRowIndex + 1, 1)
            If formatRange Is Nothing Then
                formatRange = lengthRange
            Else
                formatRange = formatRange.Union(lengthRange)
            End If
        End If

        '---Mise en forme
        Dim minPoint As ConditionalFormattingValue = vertSepWsheet.ConditionalFormattings.CreateValue(ConditionalFormattingValueType.MinMax)
        Dim minColor = System.Drawing.Color.FromArgb(243, 229, 171) 'Vanilla
        Dim maxPoint As ConditionalFormattingValue = vertSepWsheet.ConditionalFormattings.CreateValue(ConditionalFormattingValueType.MinMax)
        Dim maxColor = System.Drawing.Color.FromArgb(65, 169, 76) 'Green
        Dim colorScale As ColorScale2ConditionalFormatting = vertSepWsheet.ConditionalFormattings.AddColorScale2ConditionalFormatting(formatRange, minPoint, minColor, maxPoint, maxColor)

        'FILL CRITERIAS
        'vertSepWsheet.Columns().AutoFit(typeStartCol, typeStartCol + arrayColsNb - 1)
        myFile.ReportInfos.nbWalls = tgmObjsList.Count
        If thicknessDico.ContainsKey(noneDbl) Then
            myFile.ReportInfos.nbWallswithAxis = tgmObjsList.Count - thicknessDico.Item(noneDbl).Count
        Else
            myFile.ReportInfos.nbWallswithAxis = tgmObjsList.Count
        End If
        myFile.ReportInfos.nbWallsWithWidth = myFile.ReportInfos.nbWallswithAxis - nbFalse

        vertSepWsheet.Cells(1, 2).Value = myFile.ReportInfos.nbWalls
        vertSepWsheet.Cells(2, 2).Value = myFile.ReportInfos.nbWallswithAxis
        vertSepWsheet.Cells(2, 3).Value = $"({Math.Floor(myFile.ReportInfos.nbWallswithAxis / myFile.ReportInfos.nbWalls * 100)}%)"
        vertSepWsheet.Cells(4, 1).Value = "Critère d'épaisseur :  < " + maxThick.ToString + "m  et  > " + minThick.ToString + "m"
        vertSepWsheet.Cells(5, 2).Value = nbFalse
#End Region

#Region "FILL REPORT"
        'Implementation des infos
        myFile.ReportInfos.nbWalls = tgmObjsList.Count
        If thicknessDico.ContainsKey(noneDbl) Then
            myFile.ReportInfos.nbWallswithAxis = tgmObjsList.Count - thicknessDico.Item(noneDbl).Count
        Else
            myFile.ReportInfos.nbWallswithAxis = tgmObjsList.Count
        End If
        myFile.ReportInfos.nbWallsWithWidth = myFile.ReportInfos.nbWallswithAxis - nbFalse

        'FILL REPORT
        myFile.ReportInfos.CompleteCriteria(68)
        myFile.ReportInfos.CompleteCriteria(69)
#End Region


        'SAVE
        myFile.ProjWs.PushAllModifiedEntities()
        myFile.AuditWorkbook.Calculate()
        myFile.AuditWorkbook.Worksheets.ActiveWorksheet = myFile.ReportWorksheet
        myFile.SaveAuditWorkbook()


        'OUTPUT TREE
        OutputTree = WallAnalysisOutputTree(minThick, maxThick, myFile, thicknessDico)

        myFile.ProjWs.PushAllModifiedEntities()
        Return New SucceededActionResult("Succeeded", System.Drawing.Color.Green, New List(Of Element) From {Murs})
    End Function

    Private Function WallAnalysisOutputTree(minThick As Double, maxThick As Double, myFile As FileObject, thicknessDico As Dictionary(Of Double, List(Of (objTgm As AecWall, length As Double)))) As Tree
        'Présentation des résultats grâce à une zone de filtrage

        'Definition des couleurs
        Dim colorNodes As New List(Of NodeColorTransparency)
        Dim grey20 As New Tuple(Of System.Windows.Media.Color, Single)(System.Windows.Media.Color.FromRgb(191, 191, 191), 0.2)
        Dim red100 As New Tuple(Of System.Windows.Media.Color, Single)(System.Windows.Media.Color.FromRgb(255, 0, 0), 1.0)
        Dim Violet100 As New Tuple(Of System.Windows.Media.Color, Single)(System.Windows.Media.Color.FromRgb(127, 0, 255), 1.0)

        Dim checkedNodes As New List(Of Node)

        'creation de l'arbre
        Dim myTree = myFile.FileTgm.GetProject.SmartAddTree("Analyse - Murs")
        Dim fileNode = myTree.SmartAddNode("File", myFile.FileTgm.Name)
        'myTree.RemoveNode(fileNode)
        'fileNode = myTree.SmartAddNode("File", myFile.FileTgm.Name)
        Dim typeNode = fileNode.SmartAddNode("CategoryByTgm", "Wall",,, "Murs")

        'Murs sans axes
        Dim NoAxisNode As Node = typeNode.SmartAddNode("LodByTgm", 0,,, "Murs sans axe")
        colorNodes.Add(New NodeColorTransparency(NoAxisNode, red100.Item1, red100.Item2))
        checkedNodes.Add(NoAxisNode)
        Dim hasAxisNode As Node = typeNode.SmartAddNode("LodByTgm", 100, AttributeType.Number, NodeOperator.GreaterOrEqual, "Murs avec axe : épaisseur")

        'Murs avec axes
        Dim sortedThickness = thicknessDico.Keys.ToList
        sortedThickness.Sort()
        For Each thickness In sortedThickness
            Dim thickNode As Node = hasAxisNode.SmartAddNode("ThicknessByTgm_mm", thickness, AttributeType.Number)
            checkedNodes.Add(thickNode)
            If thickness < minThick Or thickness > maxThick Then
                colorNodes.Add(New NodeColorTransparency(thickNode, Violet100.Item1, Violet100.Item2))
            Else
                colorNodes.Add(New NodeColorTransparency(thickNode, grey20.Item1, grey20.Item2))
            End If
        Next

        myFile.ProjWs.PushAllModifiedEntities()
        myTree.WriteColors(colorNodes)
        myFile.ProjWs.PushAllModifiedEntities()

        TgmView.CreateView(myFile.FileTgm.GetProject, myFile.FileTgm.Name & " - ANALYSE MURS", New List(Of Workspace) From {myFile.ScanWs}, myTree, checkedNodes, colorNodes)
        Return myTree
    End Function

End Class
