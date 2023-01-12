Imports M4D.Production.Core
Imports M4D.Treegram.Core.Entities
Imports M4D.Treegram.Core.Kernel
Imports M4D.Treegram.Core.Extensions.Entities
Imports M4D.Treegram.Core.Extensions.Kernel
Imports DevExpress.Spreadsheet
Imports DevExpress.Export.Xl
Imports Treegram.Bam.Libraries.AEC
Imports Treegram.Bam.Functions.AEC
Imports TGM.Deixi.Services
Imports M4D.Treegram.Core.Extensions.Models
Imports Treegram.Bam.Functions
Imports Treegram.Bam.Libraries.AuditFile

Public Class Analyse_CurtainWallsProdScript
    Inherits ProdScript
    Public Sub New()
        MyBase.New("IFC :: Analyse : Curtain Walls")
        AddAction(New Analyse_CurtainWalls() With {.Name = Name, .PartOfScript = True})
    End Sub
End Class
Public Class Analyse_CurtainWalls
    Inherits ProdAction
    Public Sub New()
    End Sub

    Public Overrides Sub CreateGuiInputTree()
        'Création d'une zone de filtrage (arbre) temporaire
        InputTree = TempWorkspace.AddTree("--- Temporary Input Tree ---")

        'Création d'un filtre : murs
        Dim objectNode = InputTree.SmartAddNode("object.containedIn", "3D", , , "3D Objects")
        Dim wallNode = objectNode.SmartAddNode("object.type", "IfcCurtainWall", , , DicoObjType("IfcCurtainWall"))

        'Drag and Drop automatique des objets filtrés vers l'algo (Optionnel)
        SelectYourSetsOfInput.Add("Objects3D", {wallNode}.ToList())
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
    Public Function MyMethod(Objects3D As MultipleElements) As ActionResult

        'PREPARATION ENVIRONNEMENT TGM
        '--Sortie de l'algo s'il n'y a pas d'inputs
        If Objects3D.Source.Count = 0 Then
            Return New IgnoredActionResult("No input")
        End If

        '--Récupération des métaobjets murs sous formé de liste
        Dim tgmObjsList = GetInputAsMetaObjects(Objects3D.Source).ToHashSet.ToList 'To get rid of duplicates

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
        Dim analysisWsheetName = "MURS RIDEAUX"
        If myFile.AuditWorkbook Is Nothing Then Create_IfcInitialAudit.GetAuditFromFile(myFile)
        Dim vertSepSettingsWsheet = myFile.GetOrInsertSettingsWSheet("MURS")
        If vertSepSettingsWsheet Is Nothing Then Throw New Exception("Could not find """ + analysisWsheetName + """ worksheet in settings workbook")
        Dim minThick, maxThick As Double
        If vertSepSettingsWsheet.DefinedNames.Contains("murepmin") Then 'New settings
            Try
                minThick = CDbl(vertSepSettingsWsheet.Range("murrideauepmin").Value.NumericValue) 'm
                maxThick = CDbl(vertSepSettingsWsheet.Range("murrideauepmax").Value.NumericValue) 'm
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

        Try 'Classer les nouveaux settings chronologiquement, le plus ancien en premier
            myFile.ReportInfos.OpeningAnalysis_widthAtt = myFile.GetOrInsertSettingsWSheet("METIER OUVERTURES").Range("widthAtt").Value.ToString
            myFile.ReportInfos.OpeningAnalysis_heightAtt = myFile.GetOrInsertSettingsWSheet("METIER OUVERTURES").Range("heightAtt").Value.ToString
            myFile.ReportInfos.OpeningAnalysis_glazAreaAtt = myFile.GetOrInsertSettingsWSheet("METIER OUVERTURES").Range("glazAreaAtt").Value.ToString
            myFile.ReportInfos.OpeningAnalysis_widthTolerance = CDbl(myFile.GetOrInsertSettingsWSheet("METIER OUVERTURES").Range("widthTol").Value.NumericValue)
            myFile.ReportInfos.OpeningAnalysis_heightTolerance = CDbl(myFile.GetOrInsertSettingsWSheet("METIER OUVERTURES").Range("heightTol").Value.NumericValue)
            myFile.ReportInfos.OpeningAnalysis_glazAreaTolerance = CDbl(myFile.GetOrInsertSettingsWSheet("METIER OUVERTURES").Range("glazAreaTol").Value.NumericValue)
        Catch ex As Exception
        End Try

#Region "ANALYSE"
        Dim aecCurtWalls = tgmObjsList.Select(Function(o) New AecCurtainWall(o)).ToList

        '---EXCLUDE OBJECTS WITHOUT GEOMETRY
        aecCurtWalls = CommonAecFunctions.ExcludeNoGeomObjects(aecCurtWalls).Cast(Of AecCurtainWall).ToList

        '---LOCALPLACEMENT ANNALYSIS
        For Each aecWall In aecCurtWalls
            CommonAecGeomFunctions.SmartAddLocalPlacement(aecWall, myFile)
        Next

        '---GET AXIS
        Dim wallAxis = CommonAecFunctions.SmartAddAxis(aecCurtWalls, myFile)

        '---PROFILE ANALYSIS
        Dim anaObjProfiles = CommonAecFunctions.SmartAddOrientedBBoxProfile(aecCurtWalls, myFile)

        '---GET DIMENSIONS
        Dim noneDbl = -1.0
        Dim thicknessDico As New Dictionary(Of Double, List(Of (objTgm As AecCurtainWall, length As Double)))
        For Each aecCurtWall In aecCurtWalls
            aecCurtWall.CompleteTgmPset("CategoryByTgm", Category.CurtainWall.ToString)
            aecCurtWall.CompleteTgmPset("IsVerticalSeparator", True)

            If aecCurtWall.VectorU.IsZero Or Not wallAxis.ContainsKey(aecCurtWall) Then
                'aecCurtWall.CompleteTgmPset("IsVerticalSeparator", False)
                aecCurtWall.CompleteTgmPset("LodByTgm", 0)

                'Add to dico
                If thicknessDico.ContainsKey(noneDbl) Then
                    thicknessDico(noneDbl).Add((aecCurtWall, 0.0))
                Else
                    thicknessDico.Add(noneDbl, New List(Of (objTgm As AecCurtainWall, length As Double)) From {(aecCurtWall, 0.0)})
                End If


            Else
                'thickness
                Dim thick As Double = Math.Round(aecCurtWall.MaxV - aecCurtWall.MinV, 5) 'rounded to 10-5 m
                aecCurtWall.CompleteTgmPset("ThicknessByTgm", thick)
                'length
                Dim length As Double = Math.Round(aecCurtWall.MaxU - aecCurtWall.MinU, 5) 'rounded to 10-5 m
                aecCurtWall.CompleteTgmPset("LengthByTgm", length)
                'height
                Dim minMaxZ = aecCurtWall.SmartAddExtremumZ
                aecCurtWall.CompleteTgmPset("HeightByTgm", Math.Round(minMaxZ.Item2 - minMaxZ.Item1, 5))

                If thick < minThick OrElse thick > maxThick Then
                    'aecCurtWall.CompleteTgmPset("IsVerticalSeparator", False)
                    aecCurtWall.CompleteTgmPset("LodByTgm", 0)
                Else
                    'aecCurtWall.CompleteTgmPset("IsVerticalSeparator", True)
                    aecCurtWall.CompleteTgmPset("LodByTgm", 100)
                End If

                'Add to dico
                Dim length_cm = Math.Round(length, 2) 'rounded to cm
                Dim thick_mm = Math.Round(thick, 3)
                aecCurtWall.CompleteTgmPset("ThicknessByTgm_mm", thick_mm)

                If thicknessDico.ContainsKey(thick_mm) Then
                    thicknessDico(thick_mm).Add((aecCurtWall, length_cm))
                Else
                    thicknessDico.Add(thick_mm, New List(Of (objTgm As AecCurtainWall, length As Double)) From {(aecCurtWall, length_cm)})
                End If
            End If
        Next

        '---OPENING PROFILES
        'Vraiment utile ?
        Analyse_Openings.GetOpeningOverallDimensions(aecCurtWalls, myFile)
        'Vraiment utile ?
        Analyse_Openings.GetReservationDimensions(aecCurtWalls, myFile)
        'Ok !
        Dim glazingProfiles = Analyse_Openings.GetGlazingDimensions(aecCurtWalls, myFile)
        For Each aecCurtWall In aecCurtWalls
            If glazingProfiles.ContainsKey(aecCurtWall) Then
                aecCurtWall.CompleteTgmPset("IsGlazed", True)
            Else
                aecCurtWall.CompleteTgmPset("IsGlazed", False)
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
        vertSepWsheet.Cells(4, 1).Value = "Critère d'épaisseur :  < " + maxThick.ToString + "m  et  > " + minThick.ToString + "m"
        vertSepWsheet.Cells(5, 2).Value = nbFalse

#End Region

        'SAVE
        myFile.ProjWs.PushAllModifiedEntities()
        myFile.AuditWorkbook.Calculate()
        myFile.AuditWorkbook.Worksheets.ActiveWorksheet = myFile.ReportWorksheet
        myFile.SaveAuditWorkbook()


        'OUTPUT TREE
        OutputTree = CurtainWallAnalysisOutputTree(minThick, maxThick, myFile, thicknessDico)

        Return New SucceededActionResult("Succeeded", System.Drawing.Color.Green, New List(Of Element) From {Objects3D})
    End Function

    Private Function CurtainWallAnalysisOutputTree(minThick As Double, maxThick As Double, myFile As FileObject, thicknessDico As Dictionary(Of Double, List(Of (objTgm As AecCurtainWall, length As Double)))) As Tree
        'Présentation des résultats grâce à une zone de filtrage
        'Dim savedColors As New Dictionary(Of Node, Tuple(Of System.Windows.Media.Color, Single))
        Dim colorNodes As New List(Of NodeColorTransparency)
        Dim grey20 As New Tuple(Of System.Windows.Media.Color, Single)(System.Windows.Media.Color.FromRgb(191, 191, 191), 0.2F)
        Dim red100 As New Tuple(Of System.Windows.Media.Color, Single)(System.Windows.Media.Color.FromRgb(255, 0, 0), 1.0F)
        Dim checkedNodes As New List(Of Node)

        Dim myTree = myFile.FileTgm.GetProject.SmartAddTree("Analyse - Murs Rideaux")
        Dim fileNode = myTree.SmartAddNode("File", myFile.FileTgm.Name)
        myTree.RemoveNode(fileNode)
        fileNode = myTree.SmartAddNode("File", myFile.FileTgm.Name)

        Dim typeNode = fileNode.SmartAddNode("CategoryByTgm", "CurtainWall")

        Dim sortedThickness = thicknessDico.Keys.ToList
        sortedThickness.Sort()
        For Each thickness In sortedThickness
            Dim thickNode = typeNode.SmartAddNode("ThicknessByTgm_mm", thickness, AttributeType.Number)
            checkedNodes.Add(thickNode)
            If thickness < minThick Or thickness > maxThick Then
                colorNodes.Add(New NodeColorTransparency(thickNode, red100.Item1, red100.Item2))
                'savedColors.Add(thickNode, red100)
            Else
                colorNodes.Add(New NodeColorTransparency(thickNode, grey20.Item1, grey20.Item2))
                'savedColors.Add(thickNode, grey20)
            End If
        Next

        myFile.ProjWs.PushAllModifiedEntities()
        myTree.WriteColors(colorNodes)
        TgmView.CreateView(myFile.ProjWs, myFile.FileTgm.Name & " - Validation des Murs Rideaux", {myFile.ScanWs}.ToList, myTree, checkedNodes, colorNodes)
        myFile.ProjWs.PushAllModifiedEntities()
        Return myTree
    End Function

End Class
