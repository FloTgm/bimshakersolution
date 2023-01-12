Imports System.Windows.Forms
Imports DevExpress.Export.Xl
Imports DevExpress.Spreadsheet
Imports DevExpress.Spreadsheet.Charts
Imports M4D.Production.Core
Imports M4D.Treegram.Core.Entities
Imports M4D.Treegram.Core.Extensions.Entities
Imports M4D.Treegram.Core.Extensions.Kernel
Imports M4D.Treegram.Core.Extensions.Models
Imports M4D.Treegram.Core.Kernel
Imports SharpDX
Imports TGM.Deixi.Services
Imports Treegram.Bam.Functions
Imports Treegram.Bam.Functions.AEC
Imports Treegram.GeomLibrary
Imports Treegram.Bam.Libraries.AEC
Imports Treegram.Bam.Libraries.AuditFile
Imports Treegram.GeomKernel.BasicModeler.Surface
Imports Treegram.GeomFunctions

Public Class Analyse_OpeningsProdScript
    Inherits ProdScript
    Public Sub New()
        MyBase.New("IFC :: Analyse : Openings")
        AddAction(New Analyse_Openings())
    End Sub
End Class
Public Class Analyse_Openings
    Inherits ProdAction
    Public Sub New()
        Name = "IFC :: Analyse : Openings (ProdAction)"
        PartOfScript = True
    End Sub

    Public Overrides Sub CreateGuiInputTree()

        InputTree = TempWorkspace.AddTree("--- Temporary Input Tree ---")
        Dim objectNode = InputTree.SmartAddNode("object.containedIn", "3D",,, "Objets 3D")
        '---IFC objects
        Dim ifcNode = objectNode.SmartAddNode("object.type", "ifc",,, "Objets IFC",, True)

        Dim doorNode = ifcNode.SmartAddNode("object.type", "IfcDoor")
        Dim windNode = ifcNode.SmartAddNode("object.type", "IfcWindow")
        ifcNode.SmartAddNode("object.type", "IfcCurtainWall") 'Some of them have to be INCLUDED in some projects (ex : Charenton)
        ifcNode.SmartAddNode("object.type", "IfcOpeningElement") 'Some of them have to be INCLUDED in some projects

        SelectYourSetsOfInput.Add("Doors", {doorNode}.ToList)
        SelectYourSetsOfInput.Add("Windows", {windNode}.ToList)

    End Sub

    Public Overrides Function GuiUpdateLauncherTree(inputs As Dictionary(Of String, IEnumerable(Of PersistEntity))) As Tree
        Dim launchTree = TempWorkspace.SmartAddTree("LauncherTree")
        CommonAecLauncherTrees.FilesLaunchTree(launchTree, inputs)
        If launchTree.Nodes.Count = 0 Then
            launchTree.AddNode("", Nothing)
        End If
        Return launchTree
    End Function

    <ActionMethod()>
    Public Function MyMethod(Doors As MultipleElements, Windows As MultipleElements) As ActionResult

        'PREPARATION ENVIRONNEMENT TGM
        '--Sortie de l'algo s'il n'y a pas d'inputs
        If Doors.Source.Count = 0 And Windows.Source.Count = 0 Then
            Return New IgnoredActionResult("Inputs missing")
        End If

        '--Récupération des métaobjets sous forme de liste
        Dim doorsTgmList As New List(Of MetaObject)
        If Doors.Source.Count > 0 Then
            doorsTgmList = GetInputAsMetaObjects(Doors.Source).ToHashSet.ToList 'To get rid of duplicates
        End If
        Dim windsTgmList As New List(Of MetaObject)
        If Windows.Source.Count > 0 Then
            windsTgmList = GetInputAsMetaObjects(Windows.Source).ToHashSet.ToList 'To get rid of duplicates
        End If

        '--Récupération du fichier source
        Dim tgmObjsList = doorsTgmList.Concat(windsTgmList).ToList
        Dim myFile = CommonAecFunctions.GetAuditFileFromInputs(tgmObjsList)
        If myFile.AuditWorkbook Is Nothing Then
            Create_IfcInitialAudit.GetAuditFromFile(myFile)
        End If

        '--Chargement des données BAMs
        Dim listToLoad As New List(Of Workspace) From {myFile.ScanWs}
        listToLoad.AddRange(myFile.ScanWs.GetExtensionWorspaces(M4D.Treegram.Core.Extensions.Enums.ExtensionType.AdditionalGeometries))
        listToLoad.AddRange(myFile.ScanWs.GetExtensionWorspaces(M4D.Treegram.Core.Extensions.Enums.ExtensionType.AdditionalAttributes))
        listToLoad.AddRange(myFile.ScanWs.GetExtensionWorspaces(M4D.Treegram.Core.Extensions.Enums.ExtensionType.Analysis))
        TempWorkspace.SmartAddTree("ALL").Filter(False, listToLoad).RunSynchronously()

        Treegram.GeomFunctions.Models.Reset3dDictionnary()
        Treegram.GeomFunctions.Models.LoadGeometryModels(tgmObjsList, True)
        For Each addGeomWs In myFile.ScanWs.GetExtensionWorspaces(M4D.Treegram.Core.Extensions.Enums.ExtensionType.AdditionalGeometries)
            Treegram.GeomFunctions.Models.LoadGeometryModels(addGeomWs)
        Next


        'ANALYSE MÉTIER
        Dim analysisWsheetName = "OUVERTURES"
        Try 'Classer les nouveaux settings chronologiquement, le plus ancien en premier
            myFile.ReportInfos.OpeningAnalysis_widthAtt = myFile.GetOrInsertSettingsWSheet(analysisWsheetName).Range("widthAtt").Value.ToString
            myFile.ReportInfos.OpeningAnalysis_heightAtt = myFile.GetOrInsertSettingsWSheet(analysisWsheetName).Range("heightAtt").Value.ToString
            myFile.ReportInfos.OpeningAnalysis_glazAreaAtt = myFile.GetOrInsertSettingsWSheet(analysisWsheetName).Range("glazAreaAtt").Value.ToString
            myFile.ReportInfos.OpeningAnalysis_widthTolerance = CDbl(myFile.GetOrInsertSettingsWSheet(analysisWsheetName).Range("widthTol").Value.NumericValue)
            myFile.ReportInfos.OpeningAnalysis_heightTolerance = CDbl(myFile.GetOrInsertSettingsWSheet(analysisWsheetName).Range("heightTol").Value.NumericValue)
            myFile.ReportInfos.OpeningAnalysis_glazAreaTolerance = CDbl(myFile.GetOrInsertSettingsWSheet(analysisWsheetName).Range("glazAreaTol").Value.NumericValue)

            'Doors Dimensions
            myFile.ReportInfos.OpeningAnalysis_DoorMaxHeight = CDbl(myFile.GetOrInsertSettingsWSheet(analysisWsheetName).Range("DoorMaxHeight").Value.NumericValue)
            myFile.ReportInfos.OpeningAnalysis_DoorMinHeight = CDbl(myFile.GetOrInsertSettingsWSheet(analysisWsheetName).Range("DoorMinHeight").Value.NumericValue)
            myFile.ReportInfos.OpeningAnalysis_DoorMaxWidth = CDbl(myFile.GetOrInsertSettingsWSheet(analysisWsheetName).Range("DoorMaxWidth").Value.NumericValue)
            myFile.ReportInfos.OpeningAnalysis_DoorMinWidth = CDbl(myFile.GetOrInsertSettingsWSheet(analysisWsheetName).Range("DoorMinWidth").Value.NumericValue)
            'Windows Dimensions
            myFile.ReportInfos.OpeningAnalysis_WindowMaxHeight = CDbl(myFile.GetOrInsertSettingsWSheet(analysisWsheetName).Range("WindowMaxHeight").Value.NumericValue)
            myFile.ReportInfos.OpeningAnalysis_WindowMinHeight = CDbl(myFile.GetOrInsertSettingsWSheet(analysisWsheetName).Range("WindowMinHeight").Value.NumericValue)
            myFile.ReportInfos.OpeningAnalysis_WindowMaxWidth = CDbl(myFile.GetOrInsertSettingsWSheet(analysisWsheetName).Range("WindowMaxWidth").Value.NumericValue)
            myFile.ReportInfos.OpeningAnalysis_WindowMinWidth = CDbl(myFile.GetOrInsertSettingsWSheet(analysisWsheetName).Range("WindowMinWidth").Value.NumericValue)
        Catch ex As Exception
        End Try

#Region "ANALYSE"
        Dim newAecObjs As New List(Of AecOpening)
        Dim categOpeningsDico As New Dictionary(Of Category, List(Of AecOpening)) From {
            {Category.Door, New List(Of AecOpening)},
            {Category.Window, New List(Of AecOpening)}
        }
        For Each openingTgm In tgmObjsList
            If doorsTgmList.Contains(openingTgm) Then
                Dim aecDoor = New AecDoor(openingTgm)
                If CommonAecFunctions.ExcludeNoGeomObjects({aecDoor}.ToList).Count > 0 Then
                    newAecObjs.Add(aecDoor)
                    categOpeningsDico(Category.Door).Add(aecDoor)
                End If
            Else
                Dim aecWind = New AecWindow(openingTgm)
                If CommonAecFunctions.ExcludeNoGeomObjects({aecWind}.ToList).Count > 0 Then
                    newAecObjs.Add(aecWind)
                    categOpeningsDico(Category.Window).Add(aecWind)
                End If
            End If
        Next
        myFile.ReportInfos.OpeningsNb = newAecObjs.Count
        myFile.ReportInfos.doorsNb = categOpeningsDico(Category.Door).Count
        myFile.ReportInfos.windowsNb = categOpeningsDico(Category.Window).Count

        Dim dicoDimensions As New Dictionary(Of Category, List(Of String))


        For Each aecOpening In newAecObjs

            'Category
            If aecOpening.GetType.Equals(GetType(AecDoor)) Then
                aecOpening.CompleteTgmPset("CategoryByTgm", Category.Door.ToString)
            Else
                aecOpening.CompleteTgmPset("CategoryByTgm", Category.Window.ToString)
            End If

            'Local placement
            If CommonAecGeomFunctions.SmartAddLocalPlacement(aecOpening, myFile) Then
                aecOpening.CompleteTgmPset("WidthByTgm", Math.Round(aecOpening.MaxU - aecOpening.MinU, 5))
                aecOpening.CompleteTgmPset("WidthByTgm_mm", Math.Round(aecOpening.MaxU - aecOpening.MinU, 3))
            End If

            'MinZ and MaxZ
            Dim minMaxZ = aecOpening.SmartAddExtremumZ
            aecOpening.CompleteTgmPset("HeightByTgm", Math.Round(minMaxZ.Item2 - minMaxZ.Item1, 5))
            aecOpening.CompleteTgmPset("HeightByTgm_mm", Math.Round(minMaxZ.Item2 - minMaxZ.Item1, 3))
            Dim tgmDimension = Math.Round((aecOpening.MaxU - aecOpening.MinU) * 100, 1).ToString + "x" + Math.Round((minMaxZ.Item2 - minMaxZ.Item1) * 100, 1).ToString
            aecOpening.CompleteTgmPset("DimensionsByTgm", tgmDimension)

            'Complete dimensions dico
            If aecOpening.GetType.Equals(GetType(AecDoor)) Then
                If dicoDimensions.ContainsKey(Category.Door) Then
                    dicoDimensions(Category.Door).Add(tgmDimension)
                Else
                    dicoDimensions.Add(Category.Door, New List(Of String) From {tgmDimension})
                End If
            Else
                If dicoDimensions.ContainsKey(Category.Window) Then
                    dicoDimensions(Category.Window).Add(tgmDimension)
                Else
                    dicoDimensions.Add(Category.Window, New List(Of String) From {tgmDimension})
                End If
            End If
        Next

        'Overall profiles
        Dim overallProfiles = GetOpeningOverallDimensions(newAecObjs, myFile)
        For Each aecOpen In overallProfiles.Keys
            aecOpen.CompleteTgmPset("OverallAreaByTgm", Math.Round(overallProfiles(aecOpen).ToBasicModelerSurface.Area3D, 4)) 'cm2 <-- à voir si c'est mieux dans la méthode !?
        Next

        'Reservation profiles
        Dim reservationProfiles = GetReservationDimensions(newAecObjs, myFile)
        Dim linkedOpenings As Integer = 0
        For Each aecOpening In newAecObjs
            Dim typesBelonging = ""
            Dim parents = aecOpening.Metaobject.GetParents("IfcOpeningElement").ToList
            If parents.Count = 1 Then
                Dim tgmOpening = parents.ElementAt(0)
                For Each parent In tgmOpening.GetParents().ToList
                    If CType(parent.Container, Workspace).Equals(myFile.ScanWs) And Not parent.Name = "3D" And Not parent.GetTgmType = "IfcGroup" Then
                        If typesBelonging = "" Then
                            typesBelonging = parent.GetTgmType
                        Else
                            typesBelonging = typesBelonging + ";" + parent.GetTgmType
                        End If
                    End If
                Next
                If Not String.IsNullOrEmpty(typesBelonging) Then
                    linkedOpenings += 1
                End If
            End If
            aecOpening.CompleteTgmPset("OpeningContainment", typesBelonging)
        Next
        myFile.ReportInfos.reservationNb = linkedOpenings

        'LodByTgm
        For Each aecOpening In newAecObjs
            If overallProfiles.ContainsKey(aecOpening) Or reservationProfiles.ContainsKey(aecOpening) Then
                aecOpening.CompleteTgmPset("IsOpenable", True)
                aecOpening.CompleteTgmPset("LodByTgm", 100)
            Else
                aecOpening.CompleteTgmPset("IsOpenable", False)
                aecOpening.CompleteTgmPset("LodByTgm", 0)
            End If
        Next


        'Glazing profiles
        Dim glazingProfiles = GetGlazingDimensions(newAecObjs, myFile)
        For Each aecOpening In newAecObjs
            If glazingProfiles.ContainsKey(aecOpening) Then
                aecOpening.CompleteTgmPset("IsGlazed", True)
            Else
                aecOpening.CompleteTgmPset("IsGlazed", False)
            End If
        Next

        'Other points and directions
        GetOpeningOrientation(newAecObjs, myFile)
        GetOpeningInsertPoint(newAecObjs, myFile)
        GetOpeningSwingDirection(newAecObjs, myFile)


        'Get resulting Dico
        CommonFunctions.RunFilter(TempWorkspace, listToLoad)
        Dim openingParametersDico As New Dictionary(Of AecOpening, (height As Double, width As Double, glazing As Double, heightComparison As Double?, widthComparison As Double?, glazingComparison As Double?, orientation As String))
        For Each aecOpening In newAecObjs
            Dim openingParam As (height As Double, width As Double, glazing As Double, heightComparison As Double?, widthComparison As Double?, glazingComparison As Double?, orientation As String)
            openingParam.height = Math.Round(aecOpening.HeightByTgm * 100.0, 1)
            openingParam.heightComparison = aecOpening.Metaobject.GetAttribute("CoherenceIfcGeoHeight", True)?.Value * 100.0
            openingParam.width = Math.Round(aecOpening.WidthByTgm * 100.0, 1)
            openingParam.widthComparison = aecOpening.Metaobject.GetAttribute("CoherenceIfcGeoWidth", True)?.Value * 100.0
            openingParam.glazing = aecOpening.GlazedAreaByTgm * 100.0
            openingParam.glazingComparison = aecOpening.Metaobject.GetAttribute("CoherenceIfcGeoGlaArea", True)?.Value * 100.0
            openingParam.orientation = aecOpening.Metaobject.GetAttribute("OrientationByTgm", True)?.Value
            openingParametersDico.Add(aecOpening, openingParam)
        Next

        Dim comparisonDico As New Dictionary(Of Category, (widthCompList As List(Of CommonAecFunctions.MeasureComparison), heightCompList As List(Of CommonAecFunctions.MeasureComparison), glazCompList As List(Of CommonAecFunctions.MeasureComparison)))
        Dim dimensionsArrayDico As New Dictionary(Of Category, (dimArray As Array, widthCount As Integer, heightCount As Integer))
        Dim badCountDico As New Dictionary(Of Category, (Integer, Integer))
        For Each openingCategory In categOpeningsDico.Keys

            Dim nbDIco As New Dictionary(Of (Double, Double), List(Of AecOpening))
            Dim heightWidth As (Double, Double) = (0, 0)

            'GET COMPARISON DICO
            Dim widthCompList As New List(Of CommonAecFunctions.MeasureComparison)
            Dim heightCompList As New List(Of CommonAecFunctions.MeasureComparison)
            Dim glazCompList As New List(Of CommonAecFunctions.MeasureComparison)

            For Each aecOpening In categOpeningsDico(openingCategory)
                Dim openingParameters = openingParametersDico(aecOpening)

                'If nbDIco.ContainsKey((Math.Round(openingParameters.height, 1), Math.Round(openingParameters.width, 1))) Then 'PK ON ARRONDI DE NOUVEAU ?????????????????????????
                '    nbDIco((Math.Round(openingParameters.height, 1), Math.Round(openingParameters.width, 1))).AddRange({aecOpening})
                'Else
                '    nbDIco.Add((Math.Round(openingParameters.height, 1), Math.Round(openingParameters.width, 1)), New List(Of AecOpening) From {aecOpening})
                'End If
                If nbDIco.ContainsKey((openingParameters.height, openingParameters.width)) Then
                    nbDIco((openingParameters.height, openingParameters.width)).AddRange({aecOpening})
                Else
                    nbDIco.Add((openingParameters.height, openingParameters.width), New List(Of AecOpening) From {aecOpening})
                End If

                'Width comparison
                If openingParameters.widthComparison Is Nothing Then
                    widthCompList.Add(CommonAecFunctions.MeasureComparison.NR)
                ElseIf openingParameters.widthComparison > myFile.ReportInfos.OpeningAnalysis_widthTolerance Then
                    widthCompList.Add(CommonAecFunctions.MeasureComparison.different)
                Else
                    widthCompList.Add(CommonAecFunctions.MeasureComparison.equal)
                End If
                'Height comparison
                If openingParameters.heightComparison Is Nothing Then
                    heightCompList.Add(CommonAecFunctions.MeasureComparison.NR)
                ElseIf openingParameters.heightComparison > myFile.ReportInfos.OpeningAnalysis_heightTolerance Then
                    heightCompList.Add(CommonAecFunctions.MeasureComparison.different)
                Else
                    heightCompList.Add(CommonAecFunctions.MeasureComparison.equal)
                End If
                'Glazing comparison
                If openingParameters.glazingComparison Is Nothing Then
                    glazCompList.Add(CommonAecFunctions.MeasureComparison.NR)
                ElseIf openingParameters.glazingComparison > myFile.ReportInfos.OpeningAnalysis_glazAreaTolerance Then
                    glazCompList.Add(CommonAecFunctions.MeasureComparison.different)
                Else
                    glazCompList.Add(CommonAecFunctions.MeasureComparison.equal)
                End If

            Next
            comparisonDico.Add(openingCategory, (widthCompList, heightCompList, glazCompList))

            'GET ARRAY DIMENSIONS
            Dim widthList = categOpeningsDico(openingCategory).Select(Function(o) openingParametersDico(o).width).ToList.OrderBy(Function(o) o).ToHashSet
            '-------------------> ARRONDIR AU MM
            Dim heightList = categOpeningsDico(openingCategory).Select(Function(o) openingParametersDico(o).height).ToList.OrderBy(Function(o) o).ToHashSet
            Dim i, j As Integer
            Dim myArray(widthList.Count, heightList.Count) As Double
            myArray(0, 0) = Nothing
            'first row
            For j = 1 To heightList.Count
                myArray(0, j) = heightList(j - 1)
            Next
            'first column
            For i = 1 To widthList.Count
                myArray(i, 0) = widthList(i - 1)
            Next
            'fill array
            For i = 1 To widthList.Count
                For j = 1 To heightList.Count
                    myArray(i, j) = categOpeningsDico(openingCategory).Where(Function(o) openingParametersDico(o).width = widthList(i - 1) And openingParametersDico(o).height = heightList(j - 1)).Count
                Next
            Next
            dimensionsArrayDico.Add(openingCategory, (myArray, widthList.Count, heightList.Count))

            Dim maxheight, minheight, maxWidth, minWidth As Double
            If openingCategory = Category.Door Then
                maxheight = myFile.ReportInfos.OpeningAnalysis_DoorMaxHeight
                minheight = myFile.ReportInfos.OpeningAnalysis_DoorMinHeight
                maxWidth = myFile.ReportInfos.OpeningAnalysis_DoorMaxWidth
                minWidth = myFile.ReportInfos.OpeningAnalysis_DoorMinWidth
            ElseIf openingCategory = Category.Window Then
                maxheight = myFile.ReportInfos.OpeningAnalysis_WindowMaxHeight
                minheight = myFile.ReportInfos.OpeningAnalysis_WindowMinHeight
                maxWidth = myFile.ReportInfos.OpeningAnalysis_WindowMaxWidth
                minWidth = myFile.ReportInfos.OpeningAnalysis_WindowMinWidth
            End If

            Dim nbBadDImensions As Integer = 0
            For Each heightWidth In nbDIco.Keys
                If CDbl(heightWidth.Item1) < minheight * 100 Or CDbl(heightWidth.Item1) > maxheight * 100 Or CDbl(heightWidth.Item2) < minWidth * 100 Or CDbl(heightWidth.Item2) > maxWidth * 100 Then
                    nbBadDImensions += nbDIco(heightWidth).Count
                End If
            Next
            badCountDico.Add(openingCategory, (nbBadDImensions, nbDIco.Keys.Count))
        Next

        'GET ORIENTATIONS DICO
        Dim orientationsDico As New Dictionary(Of GeographicOrientation, (openingNumber As Integer, glazedArea As Double))
        'For Each orientationSt As String In System.[Enum].GetNames(GetType(GeographicOrientation))
        '    Dim orientedOpenings = openingParametersDico.Where(Function(o) o.Value.orientation = orientationSt And o.Value.glazing > 0.0) 'Get only glazed openings
        '    Dim orientedGlazedArea = orientedOpenings.Select(Function(o) o.Value.glazing).Sum
        '    orientationsDico.Add(orientationSt, (orientedOpenings.Count, orientedGlazedArea))
        'Next
        For Each orientation As GeographicOrientation In System.Enum.GetValues(GetType(GeographicOrientation))
            Dim orientedOpenings = openingParametersDico.Where(Function(o) o.Value.orientation = orientation.ToString And o.Value.glazing > 0.0) 'Get only glazed openings
            Dim orientedGlazedArea = orientedOpenings.Select(Function(o) o.Value.glazing).Sum
            orientationsDico.Add(orientation, (orientedOpenings.Count, orientedGlazedArea))
        Next

#End Region



#Region "FILL EXCEL"
        Dim openingWsheet As Worksheet = myFile.InsertProdActionTemplateInAuditWbook("ProdActions_Template", analysisWsheetName)

        'COMPARISON TABLE
        'Parametrage
        Dim compLine = openingWsheet.Range("parametrage").BottomRowIndex
        Dim compCol = openingWsheet.Range("parametrage").LeftColumnIndex
        openingWsheet.Cells(compLine + 1, compCol + 2).Value = myFile.ReportInfos.OpeningAnalysis_widthTolerance
        openingWsheet.Cells(compLine + 2, compCol + 2).Value = myFile.ReportInfos.OpeningAnalysis_heightTolerance
        openingWsheet.Cells(compLine + 3, compCol + 2).Value = myFile.ReportInfos.OpeningAnalysis_glazAreaTolerance
        openingWsheet.Cells(compLine + 4, compCol + 2).Value = myFile.ReportInfos.OpeningAnalysis_widthAtt
        openingWsheet.Cells(compLine + 5, compCol + 2).Value = myFile.ReportInfos.OpeningAnalysis_heightAtt
        openingWsheet.Cells(compLine + 6, compCol + 2).Value = myFile.ReportInfos.OpeningAnalysis_glazAreaAtt
        'Comparison
        Dim categoriesLine = openingWsheet.Range("IfcTypes").BottomRowIndex
        Dim categoriesCol = openingWsheet.Range("IfcTypes").LeftColumnIndex
        'Dim colNumber = openingWsheet.Range("IfcTypes").LeftColumnIndex - 1
        openingWsheet.Cells(categoriesLine, categoriesCol).Resize(2, 1).Merge 'Ne foncionne pas....
        openingWsheet.Cells(categoriesLine, categoriesCol + 1).Resize(2, 1).Merge 'Ne foncionne pas....
        openingWsheet.Cells(categoriesLine, categoriesCol + 2).Resize(1, 4).Merge 'Ne foncionne pas....

        Dim compCurrentLine = categoriesLine + 2
        For Each category In comparisonDico.Keys
            ''Mise en forme
            'For i = 1 To 3
            '    Dim myRange As CellRange = openingWsheet.Cells(compCurrentLine + i, compCol).Resize(1, colNumber)
            '    myRange.Insert(InsertCellsMode.ShiftCellsDown)
            '    Dim rangeCompleted As CellRange = openingWsheet.Cells(compCurrentLine + i - 1, compCol)
            '    rangeCompleted = rangeCompleted.Resize(1, colNumber)
            '    rangeCompleted.CopyFrom(formatRange, PasteSpecial.Formats)
            '    rangeCompleted.Borders.SetOutsideBorders(System.Drawing.Color.Black, DevExpress.Spreadsheet.BorderLineStyle.Thin)
            '    rangeCompleted.Borders.InsideVerticalBorders.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin
            '    rangeCompleted.Borders.InsideHorizontalBorders.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin
            'Next

            'Width
            'openingWsheet.Cells(compCurrentLine, compCol + 1).Value = "Largeur Totale"
            openingWsheet.Cells(compCurrentLine, compCol + 2).Value = comparisonDico(category).widthCompList.Count
            openingWsheet.Cells(compCurrentLine, compCol + 3).Value = comparisonDico(category).widthCompList.Where(Function(o) o = CommonAecFunctions.MeasureComparison.NR).Count
            openingWsheet.Cells(compCurrentLine, compCol + 4).Value = comparisonDico(category).widthCompList.Where(Function(o) o = CommonAecFunctions.MeasureComparison.different).Count
            openingWsheet.Cells(compCurrentLine, compCol + 5).Value = comparisonDico(category).widthCompList.Where(Function(o) o = CommonAecFunctions.MeasureComparison.equal).Count
            compCurrentLine += 1
            'Height
            'openingWsheet.Cells(compCurrentLine, compCol + 1).Value = "Hauteur Totale"
            openingWsheet.Cells(compCurrentLine, compCol + 2).Value = comparisonDico(category).heightCompList.Count
            openingWsheet.Cells(compCurrentLine, compCol + 3).Value = comparisonDico(category).heightCompList.Where(Function(o) o = CommonAecFunctions.MeasureComparison.NR).Count
            openingWsheet.Cells(compCurrentLine, compCol + 4).Value = comparisonDico(category).heightCompList.Where(Function(o) o = CommonAecFunctions.MeasureComparison.different).Count
            openingWsheet.Cells(compCurrentLine, compCol + 5).Value = comparisonDico(category).heightCompList.Where(Function(o) o = CommonAecFunctions.MeasureComparison.equal).Count
            compCurrentLine += 1
            'Glazing
            'openingWsheet.Cells(compCurrentLine, compCol + 1).Value = "Surface Vitrée"
            openingWsheet.Cells(compCurrentLine, compCol + 2).Value = comparisonDico(category).glazCompList.Count
            openingWsheet.Cells(compCurrentLine, compCol + 3).Value = comparisonDico(category).glazCompList.Where(Function(o) o = CommonAecFunctions.MeasureComparison.NR).Count
            openingWsheet.Cells(compCurrentLine, compCol + 4).Value = comparisonDico(category).glazCompList.Where(Function(o) o = CommonAecFunctions.MeasureComparison.different).Count
            openingWsheet.Cells(compCurrentLine, compCol + 5).Value = comparisonDico(category).glazCompList.Where(Function(o) o = CommonAecFunctions.MeasureComparison.equal).Count
            compCurrentLine += 1
            ''IfcType
            'openingWsheet.Cells(compCurrentLine - 1, compCol).Value = DicoObjType($"Ifc{category}")
            'openingWsheet.Cells(compCurrentLine - 1, compCol).Font.Bold = True
            'openingWsheet.Cells(compCurrentLine - 3, compCol).Resize(3, 1).Merge
        Next


        'DIMENSIONS TABLES
        Dim dimLine = openingWsheet.Range("tableaudimensions").BottomRowIndex
        Dim dimCol = openingWsheet.Range("tableaudimensions").LeftColumnIndex

        'decide how many columns have to be added
        Dim heightLongerCount As Integer = 0
        For Each ifctype In dimensionsArrayDico.Keys
            Dim myArr = dimensionsArrayDico(ifctype)
            Dim heightCount = myArr.Item3
            If heightCount > heightLongerCount Then
                heightLongerCount = heightCount
            End If
        Next
        openingWsheet.Columns.Insert(dimCol + 1, heightLongerCount, ColumnFormatMode.FormatAsNext)
        openingWsheet.Cells(0, dimCol + 1).Resize(1, heightLongerCount + 1).ColumnWidth = 150.0

        For Each category In dimensionsArrayDico.Keys
            If dimensionsArrayDico(category).widthCount = 0 Or dimensionsArrayDico(category).heightCount = 0 Then
                Continue For 'No door/window
            End If

            'FILL TITLES
            openingWsheet.Cells(dimLine, dimCol).Value = "" 'Delete template sentence
            Dim htRange As CellRange = openingWsheet.Cells(dimLine, dimCol + 2)
            Dim largRange As CellRange = openingWsheet.Cells(dimLine + 2, dimCol)
            htRange.Value = "HAUTEUR"
            largRange.Value = "LARGEUR"

            'mise en forme
            htRange.FillColor = XlColor.FromTheme(XlThemeColor.Accent1, 0.4)
            largRange.FillColor = XlColor.FromTheme(XlThemeColor.Accent1, 0.4)
            htRange.Font.Bold = True
            largRange.Font.Bold = True
            largRange.Alignment.Vertical = SpreadsheetVerticalAlignment.Center


            Dim allRange = openingWsheet.Cells(dimLine, dimCol).Resize(2 + dimensionsArrayDico(category).widthCount, 2 + dimensionsArrayDico(category).heightCount)
            Dim titleRange = openingWsheet.Cells(dimLine, dimCol).Resize(2, 2)
            titleRange.Merge
            openingWsheet.Cells(dimLine, dimCol + 2).Resize(1, dimensionsArrayDico(category).heightCount).Merge
            openingWsheet.Cells(dimLine + 2, dimCol).Resize(dimensionsArrayDico(category).widthCount, 1).Merge
            'largRange.AutoFitRows

            allRange.Borders.SetOutsideBorders(System.Drawing.Color.Black, DevExpress.Spreadsheet.BorderLineStyle.Thin)
            allRange.Borders.InsideVerticalBorders.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin
            allRange.Borders.InsideHorizontalBorders.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin
            titleRange.Borders.RemoveBorders()
            titleRange.Borders.BottomBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin
            titleRange.Borders.RightBorder.LineStyle = DevExpress.Spreadsheet.BorderLineStyle.Thin

            'FILL DIMENSIONS
            Dim k, l As Integer
            Dim wrongWidthList As New List(Of Integer)
            Dim wrongHeightList As New List(Of Integer)
            Dim maxheight, minheight, maxWidth, minWidth As Double
            If category = Category.Door Then
                maxheight = myFile.ReportInfos.OpeningAnalysis_DoorMaxHeight
                minheight = myFile.ReportInfos.OpeningAnalysis_DoorMinHeight
                maxWidth = myFile.ReportInfos.OpeningAnalysis_DoorMaxWidth
                minWidth = myFile.ReportInfos.OpeningAnalysis_DoorMinWidth
            ElseIf category = Category.Window Then
                maxheight = myFile.ReportInfos.OpeningAnalysis_WindowMaxHeight
                minheight = myFile.ReportInfos.OpeningAnalysis_WindowMinHeight
                maxWidth = myFile.ReportInfos.OpeningAnalysis_WindowMaxWidth
                minWidth = myFile.ReportInfos.OpeningAnalysis_WindowMinWidth
            End If
            For k = 0 To dimensionsArrayDico(category).dimArray.GetLength(0) - 1
                For l = 0 To dimensionsArrayDico(category).dimArray.GetLength(1) - 1
                    openingWsheet.Cells(dimLine + 1 + k, dimCol + 1 + l).Value = CDbl(dimensionsArrayDico(category).dimArray(k, l))

                    'Coloration des cases hors critère
                    If k = 0 Then
                        'width
                        If CDbl(dimensionsArrayDico(category).dimArray(k, l)) < minheight * 100 Or CDbl(dimensionsArrayDico(category).dimArray(k, l)) > maxheight * 100 Then
                            wrongHeightList.Add(dimensionsArrayDico(category).dimArray(k, l))
                            'Color Purple
                            openingWsheet.Cells(dimLine + 1 + k, dimCol + 1 + l).FillColor = Drawing.Color.FromArgb(150, 131, 236)
                        End If
                    End If
                    If l = 0 Then
                        'height
                        If CDbl(dimensionsArrayDico(category).dimArray(k, l)) < minWidth * 100 Or CDbl(dimensionsArrayDico(category).dimArray(k, l)) > maxWidth * 100 Then
                            wrongWidthList.Add(dimensionsArrayDico(category).dimArray(k, l))
                            'Color Purple
                            openingWsheet.Cells(dimLine + 1 + k, dimCol + 1 + l).FillColor = Drawing.Color.FromArgb(150, 131, 236)
                        End If
                    End If
                Next
            Next

            'mise en forme
            openingWsheet.Cells(dimLine + 1, dimCol + 2).Resize(1, 100).NumberFormat = "0.0"
            openingWsheet.Columns(dimCol + 1).NumberFormat = "0.0"
            titleRange.Font.Bold = True
            titleRange.Value = DicoObjType($"Ifc{category}")
            titleRange.Alignment.Vertical = SpreadsheetVerticalAlignment.Center
            titleRange.Alignment.RotationAngle = 0

            Dim countRange = openingWsheet.Cells(dimLine + 2, dimCol + 2).Resize(dimensionsArrayDico(category).widthCount, dimensionsArrayDico(category).heightCount)
            Dim minPoint As ConditionalFormattingValue = openingWsheet.ConditionalFormattings.CreateValue(ConditionalFormattingValueType.MinMax)
            Dim minColor = Drawing.Color.FromArgb(243, 229, 171) 'Vanilla
            Dim maxPoint As ConditionalFormattingValue = openingWsheet.ConditionalFormattings.CreateValue(ConditionalFormattingValueType.MinMax)
            Dim maxColor = Drawing.Color.FromArgb(65, 169, 76) 'Green
            Dim colorScale As ColorScale2ConditionalFormatting = openingWsheet.ConditionalFormattings.AddColorScale2ConditionalFormatting(countRange, minPoint, minColor, maxPoint, maxColor)
            colorScale.Priority = 2
            Dim expression As ExpressionConditionalFormatting = openingWsheet.ConditionalFormattings.AddExpressionConditionalFormatting(countRange, ConditionalFormattingExpressionCondition.EqualTo, 0)
            expression.Formatting.Fill.BackgroundColor = Drawing.Color.FromArgb(255, 255, 255) 'White
            expression.Priority = 1

            dimLine += dimensionsArrayDico(category).widthCount + 3
        Next


        'ORIENTATIONS TABLE
        Dim orientRange As CellRange = openingWsheet.Range("tableauorientations")
        Dim orientLine = orientRange.BottomRowIndex
        Dim orientCol = orientRange.LeftColumnIndex
        openingWsheet.Cells(orientLine, orientCol).Value = "Orientations"
        openingWsheet.Cells(orientLine, orientCol + 1).Value = "Surfaces vitrées"
        openingWsheet.Cells(orientLine, orientCol + 2).Value = "Nombre d'ouvertures vitrées"
        Dim startCell As CellRange = openingWsheet.Cells(orientLine, orientCol)
        Dim startCell2 As CellRange = openingWsheet.Cells(orientLine, orientCol + 1)
        orientLine += 1
        Dim lastCell, lastCell2 As CellRange
        For Each orient In orientationsDico.Keys
            lastCell = openingWsheet.Cells(orientLine, orientCol)
            lastCell.Value = orient.ToString
            lastCell2 = openingWsheet.Cells(orientLine, orientCol + 1)
            lastCell2.Value = orientationsDico(orient).glazedArea
            openingWsheet.Cells(orientLine, orientCol + 2).Value = orientationsDico(orient).openingNumber
            orientLine += 1
        Next
        If orientationsDico.Keys.Count > 0 Then
            Dim chartRange = startCell.Resize(lastCell2.BottomRowIndex - orientRange.BottomRowIndex + 1, 2)
            Dim chartObj As Chart = openingWsheet.Charts.Add(ChartType.Column3DClustered, chartRange)
            chartObj.TopLeftCell = openingWsheet.Cells(orientRange.BottomRowIndex, orientRange.LeftColumnIndex) 'orientRange
            chartObj.BottomRightCell = openingWsheet.Cells(orientRange.BottomRowIndex + 10, orientRange.LeftColumnIndex + 5)
            chartObj.Outline.Width = 0.0
            chartObj.Outline.SetNoFill()
        End If



        'COMPLETE STATS
        Dim statLine As Integer = 0
        Dim statCol As Integer = 0
        Dim salmonRed = Drawing.Color.FromArgb(255, 124, 124)

        myFile.ReportInfos.revDoorsNb = badCountDico(Category.Door).Item1
        myFile.ReportInfos.revWindowsNb = badCountDico(Category.Window).Item1

        openingWsheet.Cells(statLine + 1, statCol + 1).Value = myFile.ReportInfos.doorsNb
        openingWsheet.Cells(statLine + 2, statCol + 1).Value = myFile.ReportInfos.windowsNb
        openingWsheet.Cells(statLine + 4, statCol + 1).Value = myFile.ReportInfos.OpeningsNb - myFile.ReportInfos.reservationNb
        openingWsheet.Cells(statLine + 4, statCol + 1).FillColor = salmonRed

        'Nb Typologies
        openingWsheet.Cells(statLine + 7, statCol + 1).Value = badCountDico(Category.Door).Item2
        openingWsheet.Cells(statLine + 8, statCol + 1).Value = badCountDico(Category.Window).Item2

        'Dimension criteria
        openingWsheet.Cells(statLine + 11, statCol + 1).Value = "> " + (myFile.ReportInfos.OpeningAnalysis_DoorMinWidth * 100).ToString + " cm et < " + (myFile.ReportInfos.OpeningAnalysis_DoorMaxWidth * 100).ToString + " cm"
        openingWsheet.Cells(statLine + 12, statCol + 1).Value = "> " + (myFile.ReportInfos.OpeningAnalysis_DoorMinHeight * 100).ToString + " cm et < " + (myFile.ReportInfos.OpeningAnalysis_DoorMaxHeight * 100).ToString + " cm"
        openingWsheet.Cells(statLine + 13, statCol + 1).Value = "> " + (myFile.ReportInfos.OpeningAnalysis_WindowMinWidth * 100).ToString + " cm et < " + (myFile.ReportInfos.OpeningAnalysis_WindowMaxWidth * 100).ToString + " cm"
        openingWsheet.Cells(statLine + 14, statCol + 1).Value = "> " + (myFile.ReportInfos.OpeningAnalysis_WindowMinHeight * 100).ToString + " cm et < " + (myFile.ReportInfos.OpeningAnalysis_WindowMaxHeight * 100).ToString + " cm"

        openingWsheet.Cells(statLine + 17, statCol + 1).Value = badCountDico(Category.Door).Item1
        openingWsheet.Cells(statLine + 18, statCol + 1).Value = badCountDico(Category.Window).Item1

        'Mise en page
        Dim purpleError = Drawing.Color.FromArgb(150, 131, 236)
        If badCountDico(Category.Door).Item1 > 0 Then
            openingWsheet.Cells(statLine + 17, statCol + 1).FillColor = purpleError
        End If
        If badCountDico(Category.Window).Item1 > 0 Then
            openingWsheet.Cells(statLine + 18, statCol + 1).FillColor = purpleError
        End If
        openingWsheet.Columns.Item(statCol + 1).Resize(statLine + 14, statCol + 1)

#End Region

        'FILL REPORT
        myFile.ReportInfos.CompleteCriteria(4)
        myFile.ReportInfos.CompleteCriteria(5)
        myFile.ReportInfos.CompleteCriteria(6)

        myFile.AuditWorkbook.Calculate()
        myFile.AuditWorkbook.Worksheets.ActiveWorksheet = myFile.ReportWorksheet
        myFile.SaveAuditWorkbook()

#Region "CREATE TREE"
        Dim treeName = "Analyse - Portes et Fenêtres"
        Dim myTree = myFile.ProjWs.SmartAddTree(treeName)
        Dim fileNode = myTree.SmartAddNode("File", myFile.FileTgm.Name)
        While fileNode.Nodes.Count > 0 'Reset
            fileNode.RemoveNode(fileNode.Nodes.First)
        End While

        DimensionsOutputTree(myFile, dicoDimensions, dimensionsArrayDico, treeName)
        OrientationsOutputTree(myFile, orientationsDico, treeName)
        'Dim myTree = ComparisonsOutputTree(myFile, "Analyse - Portes et Fenêtres")

#End Region




        myFile.SaveAuditWorkbook()
        myFile.ProjWs.PushAllModifiedEntities()

        OutputTree = myTree
        Return New SucceededActionResult("Succeeded", System.Drawing.Color.Green, New List(Of Element) From {Doors, Windows})
    End Function

    Public Shared Function GetOpeningOverallDimensions(openings As IEnumerable(Of AecObject), myFile As FileObject) As Dictionary(Of AecObject, Profile3D)

        Dim outputWs As Workspace = myFile.ProjWs.SmartAddWorkspace(myFile.ScanWs.Name + " - Profile", "AdditionalGeometries", True, "Scan Ifc", myFile.ScanWs) ' "San Ifc" -> To be reset at all scans

        'ANALYSIS
        Dim profileDico As New Dictionary(Of AecObject, Profile3D)
        Dim exceptionList As New HashSet(Of String) From {"Succeeded"}
        For Each aecOpening In openings
            Dim excMessage As String = Nothing
            If aecOpening.VectorU.IsZero Then
                excMessage = "No local placement"
                GoTo sendError
            End If

            'get axis system
            Dim originPt As New Treegram.GeomKernel.BasicModeler.Point(0, 0, 0)
            Dim xVec As New Treegram.GeomKernel.BasicModeler.Vector(1, 0, 0)
            Dim yVec As New Treegram.GeomKernel.BasicModeler.Vector(0, 1, 0)
            Dim zVec As New Treegram.GeomKernel.BasicModeler.Vector(0, 0, 1)

            'get local bounding box
            Dim distanceAbove = 0.02
            Dim objMinU = aecOpening.MinU
            Dim objMaxU = aecOpening.MaxU
            Dim vAverage = (aecOpening.MaxV + aecOpening.MinV) / 2
            Dim localPt1 As New Treegram.GeomKernel.BasicModeler.Point(objMinU, vAverage, aecOpening.MinZ)
            Dim localPt2 As New Treegram.GeomKernel.BasicModeler.Point(objMinU, vAverage, aecOpening.MaxZ)
            Dim localPt3 As New Treegram.GeomKernel.BasicModeler.Point(objMaxU, vAverage, aecOpening.MaxZ)
            Dim localPt4 As New Treegram.GeomKernel.BasicModeler.Point(objMaxU, vAverage, aecOpening.MinZ)

            'get x and y vectors in local axis system
            Dim localAxisAngle = aecOpening.VectorU.ToBasicModelerVector.Angle(xVec)
            Dim x2Vec = xVec.Rotate2(-localAxisAngle)
            Dim y2Vec = yVec.Rotate2(-localAxisAngle)

            'get world bounding box
            Dim pt1 = Treegram.GeomKernel.BasicModeler.ChangeCoordPointInAnotherAxisSyst(localPt1, originPt, x2Vec, y2Vec, zVec)
            Dim pt2 = Treegram.GeomKernel.BasicModeler.ChangeCoordPointInAnotherAxisSyst(localPt2, originPt, x2Vec, y2Vec, zVec)
            Dim pt3 = Treegram.GeomKernel.BasicModeler.ChangeCoordPointInAnotherAxisSyst(localPt3, originPt, x2Vec, y2Vec, zVec)
            Dim pt4 = Treegram.GeomKernel.BasicModeler.ChangeCoordPointInAnotherAxisSyst(localPt4, originPt, x2Vec, y2Vec, zVec)
            Dim mySurf As New Treegram.GeomKernel.BasicModeler.Surface(New List(Of Treegram.GeomKernel.BasicModeler.Point) From {pt1, pt2, pt3, pt4}, SurfaceType.Polygon)
            If mySurf.GetNormal Is Nothing Then
                excMessage = "Incorrect surface"
                GoTo sendError
            End If

            '---WRITE OVERALL PROFILE
            Dim profileTgm = aecOpening.Metaobject.SmartAddExtension(aecOpening.Name + " - OverallProfile", M4D.Treegram.Core.Extensions.Enums.ExtensionType.AdditionalGeometries, outputWs.Name, "Tgm")
            profileTgm.SmartAddAttribute("Area", Math.Round(mySurf.Area3D, 4)) 'cm2
            Dim overProfileAtt = profileTgm.SmartAddAttribute("OverallProfile", Nothing)
            mySurf.ToProfile3d.Write(overProfileAtt)


            profileDico.Add(aecOpening, mySurf.ToProfile3d)
            AecObject.CompleteActionStateAttribute(aecOpening, "OpeningDimensions", "Succeeded")

sendError:
            If excMessage IsNot Nothing Then
                AecObject.CompleteActionStateAttribute(aecOpening, "OpeningDimensions", excMessage)
                exceptionList.Add(excMessage)
            End If
        Next
        'CompleteActionStateTree(projWs, "OpeningDimensions", exceptionList)

        Return profileDico
    End Function

    Public Shared Function GetReservationDimensions(objects As IEnumerable(Of AecObject), myFile As FileObject) As Dictionary(Of AecObject, Profile3D)

        Dim outputWs As Workspace = myFile.ProjWs.SmartAddWorkspace(myFile.ScanWs.Name + " - Profile", "AdditionalGeometries", True, "Scan Ifc", myFile.ScanWs)

        'ANALYSIS
        Dim profileDico As New Dictionary(Of AecObject, Profile3D)
        Dim exceptionList As New HashSet(Of String) From {"Succeeded"}
        Dim creationDico As New Dictionary(Of MetaObject, Tuple(Of List(Of Profile3D), Color4))
        For Each aecObj In objects
            Dim excMessage As String = Nothing
            Dim reservation As AecObject = Nothing
            If aecObj.GetType.Equals(GetType(AecOpening)) Then
                reservation = CType(aecObj, AecOpening).Reservation
            ElseIf aecObj.GetType.Equals(GetType(AecCurtainWall)) Then
                reservation = CType(aecObj, AecCurtainWall).Reservation
            Else
                excMessage = "Cannot get reservation on this type of aec object"
                GoTo sendError
            End If
            If reservation Is Nothing Then
                aecObj.CompleteTgmPset("HasReservation", False)
                excMessage = "No opening element"
                GoTo sendError
            Else
                aecObj.CompleteTgmPset("HasReservation", True)
            End If
            Dim ifcOpenEltProfile = reservation.IfcProfile
            If ifcOpenEltProfile Is Nothing Then
                aecObj.CompleteTgmPset("HasReservation", False) 'Pas convaincu...
                excMessage = "Opening element without rectangle profile"
                GoTo sendError
            End If
            Dim ifcOpenEltSurface = ifcOpenEltProfile.ToBasicModelerSurface

            'IFC INFOS
            Dim ifcWidth = aecObj.Metaobject.GetAttribute(myFile.ReportInfos.OpeningAnalysis_widthAtt)
            Dim ifcHeight = aecObj.Metaobject.GetAttribute(myFile.ReportInfos.OpeningAnalysis_heightAtt)

            'GEO MEASURES
            Dim points = New List(Of Treegram.GeomKernel.BasicModeler.Point)
            points.AddRange(ifcOpenEltSurface.Points)
            Dim pointsZ = points.Select(Function(o) o.Z).ToList
            Dim array() As IList = {pointsZ, points}
            CommonFunctions.SortLists(Of Double)(array, Function(i1 As Double, i2 As Double)
                                                            Return i1.CompareTo(i2)
                                                        End Function)
            Dim geoWidth = Math.Round(Treegram.GeomKernel.BasicModeler.Distance(points(0), points(1)), 3)
            Dim geoHeight = Math.Round(pointsZ.Max - pointsZ.Min, 3)

            'SAVE INFOS
            Dim spaceAnaAtt = aecObj.AnalysisExtension.SmartAddAttribute("DimensionsAnalysis", Nothing, True)
            Dim ifcAtt = spaceAnaAtt.SmartAddAttribute("ByIfcInformation", Nothing)
            ifcAtt.SmartAddAttribute("IfcAnaWidth", CDbl(ifcWidth?.Value))
            ifcAtt.SmartAddAttribute("IfcAnaHeight", CDbl(ifcHeight?.Value))

            Dim geoAtt = spaceAnaAtt.SmartAddAttribute("ByGeometry", Nothing)
            geoAtt.SmartAddAttribute("GeoAnaWidth", geoWidth)
            geoAtt.SmartAddAttribute("GeoAnaHeight", geoHeight)

            Dim resultAtt = spaceAnaAtt.SmartAddAttribute("Result", Nothing, True)
            resultAtt.SmartAddAttribute("ResAnaWidth", geoWidth)
            resultAtt.SmartAddAttribute("ResAnaHeight", geoHeight)
            If ifcWidth Is Nothing Then
                resultAtt.SmartAddAttribute("CoherenceIfcGeoWidth", Nothing, True)
                'resultAtt.SmartAddAttribute("AnaWidth_NoIfcInfo", "True", True) 'On peut s'en passer...
            Else
                Dim widthDiff = Math.Round(Math.Abs(geoWidth - CDbl(ifcWidth.Value)), 3)
                resultAtt.SmartAddAttribute("CoherenceIfcGeoWidth", widthDiff, True)
                'resultAtt.SmartAddAttribute("AnaWidth_NoIfcInfo", "False", True) 'On peut s'en passer...
            End If

            If ifcHeight Is Nothing Then
                resultAtt.SmartAddAttribute("CoherenceIfcGeoHeight", Nothing, True)
                'resultAtt.SmartAddAttribute("AnaHeigth_NoIfcInfo", "True", True) 'On peut s'en passer...
            Else
                Dim heightDiff = Math.Round(Math.Abs(geoHeight - CDbl(ifcHeight.Value)), 3)
                resultAtt.SmartAddAttribute("CoherenceIfcGeoHeight", heightDiff, True)
                'resultAtt.SmartAddAttribute("AnaHeigth_NoIfcInfo", "False", True) 'On peut s'en passer...
            End If

            '---WRITE OPENING PROFILE
            Dim profileTgm = aecObj.Metaobject.SmartAddExtension(aecObj.Name + " - OpeningProfile", M4D.Treegram.Core.Extensions.Enums.ExtensionType.AdditionalGeometries, outputWs.Name, "Tgm")
            profileTgm.SmartAddAttribute("Area", ifcOpenEltSurface.Area3D)

            '---WRITE GLAZING PROFILE
            Dim openProfileAtt = profileTgm.SmartAddAttribute("OpeningProfile", Nothing)
            ifcOpenEltProfile.Write(openProfileAtt)

            If Treegram.GeomFunctions.Models.GetGeometryModels(profileTgm).Count = 0 Then
                Dim myColor = New SharpDX.Color4(213 / 255, 190 / 255, 235 / 255, 1)
                creationDico.Add(profileTgm, New Tuple(Of List(Of Profile3D), Color4)({ifcOpenEltProfile}.ToList, myColor))
            End If

            profileDico.Add(aecObj, ifcOpenEltProfile)
            AecObject.CompleteActionStateAttribute(aecObj, "OpeningDimensions", "Succeeded")

sendError:
            If excMessage IsNot Nothing Then
                AecObject.CompleteActionStateAttribute(aecObj, "OpeningDimensions", excMessage)
                exceptionList.Add(excMessage)
            End If
        Next
        'CompleteActionStateTree(projWs, "OpeningDimensions", exceptionList)

        'CREATE BEM GEOMETRY
        If creationDico.Count > 0 Then
            myFile.ProjWs.PushAllModifiedEntities()
            Treegram.GeomFunctions.Models.CreateGeometryInTreegramFromSurface(outputWs, creationDico, "OpeningProfile")
        End If

        Return profileDico
    End Function

    Public Shared Function GetGlazingDimensions(openings As IEnumerable(Of AecObject), myFile As FileObject) As Dictionary(Of AecObject, List(Of Profile3D))

        Dim outputWs As Workspace = myFile.ProjWs.SmartAddWorkspace(myFile.ScanWs.Name + " - Profile", "AdditionalGeometries", True, "Scan Ifc", myFile.ScanWs) ' "San Ifc" -> To be reset at all scans

        Treegram.GeomKernel.BasicModeler.DefaultTolerance = -8

        'ANALYSIS
        Dim newGeomCount = 0
        Dim profileDico As New Dictionary(Of AecObject, List(Of Profile3D))
        Dim exceptionList As New HashSet(Of String) From {"Succeeded"}
        Dim creationDico As New Dictionary(Of MetaObject, Tuple(Of List(Of Profile3D), Color4))
        For Each aecOpening In openings
            Dim excMessage As String = Nothing

            'IFC INFOS
            Dim ifcGlazingAreaDbl As Double = Nothing
            If myFile.ReportInfos.OpeningAnalysis_glazAreaAtt IsNot Nothing Then
                Dim ifcGlazingArea = aecOpening.Metaobject.GetAttribute(myFile.ReportInfos.OpeningAnalysis_glazAreaAtt)
                If ifcGlazingArea IsNot Nothing Then ifcGlazingAreaDbl = CDbl(ifcGlazingArea.Value.ToString.Replace(".", ","))
            End If

            'GEO MEASURES
            Dim surfList As List(Of Treegram.GeomKernel.BasicModeler.Surface) = GetGlazedArbitrarySurfaces(aecOpening)
            Dim geoGlazingArea = Math.Round(surfList.Select(Function(surf) surf.Area3D).Sum, 4) 'cm2

            'COMPLETE EXTENSION
            Dim spaceAnaAtt = aecOpening.AnalysisExtension.SmartAddAttribute("DimensionsAnalysis", Nothing, True)
            Dim ifcAtt = spaceAnaAtt.SmartAddAttribute("ByIfcInformation", Nothing)
            ifcAtt.SmartAddAttribute("IfcAnaGlaArea", ifcGlazingAreaDbl)

            Dim geoAtt = spaceAnaAtt.SmartAddAttribute("ByGeometry", Nothing)
            geoAtt.SmartAddAttribute("GeoAnaGlaArea", geoGlazingArea)
            geoAtt.SmartAddAttribute("GeoAnaGlaNumber", surfList.Count)

            Dim resultAtt = spaceAnaAtt.SmartAddAttribute("Result", Nothing, True)
            resultAtt.SmartAddAttribute("ResAnaGlaArea", geoGlazingArea)
            aecOpening.CompleteTgmPset("GlazedAreaByTgm", geoGlazingArea) 'Client purpose
            If Not ifcGlazingAreaDbl = Nothing Then
                Dim areaDiff = Math.Round(Math.Abs(geoGlazingArea - ifcGlazingAreaDbl), 4) 'cm2
                resultAtt.SmartAddAttribute("CoherenceIfcGeoGlaArea", areaDiff, True)
            Else
                resultAtt.SmartAddAttribute("CoherenceIfcGeoGlaArea", Nothing, True)
            End If

            'CREATE BEM OBJECT
            If surfList.Count > 0 Then
                Dim glazingTgm = aecOpening.Metaobject.SmartAddExtension(aecOpening.Name + " - Glazing", M4D.Treegram.Core.Extensions.Enums.ExtensionType.AdditionalGeometries, outputWs.Name, "Tgm")
                For Each att In glazingTgm.Attributes.ToList
                    glazingTgm.RemoveAttribute(att)
                Next
                glazingTgm.SmartAddAttribute("Area", geoGlazingArea)

                '---WRITE GLAZING PROFILE
                For l As Integer = 1 To surfList.Count
                    Dim glazProfileAtt = glazingTgm.SmartAddAttribute($"GlazingProfile.{l}", Nothing)
                    surfList(l - 1).ToProfile3d.Write(glazProfileAtt)
                Next

                If Treegram.GeomFunctions.Models.GetGeometryModels(glazingTgm).Count = 0 Then
                    Dim myColor = New SharpDX.Color4(49 / 255, 140 / 255, 231 / 255, 0.5)
                    creationDico.Add(glazingTgm, New Tuple(Of List(Of Profile3D), Color4)(surfList.Select(Function(o) o.ToProfile3d).ToList, myColor))
                End If

                profileDico.Add(aecOpening, surfList.Select(Function(o) o.ToProfile3d).ToList)
            End If
            AecObject.CompleteActionStateAttribute(aecOpening, "GlazingDimensions", "Succeeded")

sendError:
            If excMessage IsNot Nothing Then
                AecObject.CompleteActionStateAttribute(aecOpening, "GlazingDimensions", excMessage)
                exceptionList.Add(excMessage)
            End If
        Next
        'CompleteActionStateTree(projWs, "OpeningDimensions", exceptionList)


        'CREATE BEM GEOMETRY
        If creationDico.Count > 0 Then
            myFile.ProjWs.PushAllModifiedEntities()
            Treegram.GeomFunctions.Models.CreateGeometryInTreegramFromSurface(outputWs, creationDico, "GlazingProfile")
        End If

        Treegram.GeomKernel.BasicModeler.DefaultTolerance = -4 'Set back default value !!!

        Return profileDico
    End Function
    Private Shared Function GetGlazedArbitrarySurfaces(aecOpening As AecObject) As List(Of Treegram.GeomKernel.BasicModeler.Surface)
        Dim vecU = aecOpening.VectorU.ToBasicModelerVector
        Dim vecV = aecOpening.VectorV.ToBasicModelerVector
        Dim originPt As New Treegram.GeomKernel.BasicModeler.Point(0, 0, 0)
        Dim vecZ As New Treegram.GeomKernel.BasicModeler.Vector(0, 0, 1)

        'Get transparent models
        Dim transparModels = Treegram.GeomFunctions.Models.GetTransparentModels(aecOpening.Metaobject, True)
        Dim finalGlazedFaces As New List(Of GlazedFace)
        For Each model In transparModels
            'Get the biggest surface from the model
            Dim transparFaces = Treegram.GeomFunctions.Models.GetModelAsFacesV2(model, True)
            If transparFaces.Count = 0 Then
                Continue For
            End If
            Dim maxSurf = transparFaces.OrderBy(Function(surf) CDbl(surf.Area3D)).Last
            Dim myGlaFace As New GlazedFace With {.OriginalFace = maxSurf}

            'Get glazing 2D profile
            Dim ptList As New List(Of Treegram.GeomKernel.BasicModeler.Point)
            For Each pt In maxSurf.Points
                Dim ptInUZ = Treegram.GeomKernel.BasicModeler.ChangeCoordPointInAnotherAxisSyst(pt, originPt, vecU, vecZ, vecV.Reverse)
                ptInUZ.Z = 0.0
                ptList.Add(ptInUZ)
            Next
            myGlaFace.FaceInUZ = New Treegram.GeomKernel.BasicModeler.Surface(ptList, SurfaceType.Polygon)
            If myGlaFace.FaceInUZ.GetNormal Is Nothing Then
                Continue For
            End If

            'Compare it with glazing surfaces already treated
            Dim alreadyTreated = False
            For Each otherGlaFace In finalGlazedFaces
                If CommonGeomFunctions.GetSuperpositionArea(myGlaFace.FaceInUZ, otherGlaFace.FaceInUZ) > 0.0 Then
                    alreadyTreated = True
                    Exit For
                End If
            Next
            If Not alreadyTreated Then
                finalGlazedFaces.Add(myGlaFace)
            End If
        Next
        Return finalGlazedFaces.Select(Function(o) o.OriginalFace).ToList
    End Function
    Private Class GlazedFace
        Property OriginalFace As Treegram.GeomKernel.BasicModeler.Surface
        Property FaceInUZ As Treegram.GeomKernel.BasicModeler.Surface
    End Class
    Private Shared Function GetGlazedRectangleSurfacesFromOpeningElt(anaOpening As AecOpening) As List(Of Treegram.GeomKernel.BasicModeler.Surface)
        Dim vecU = anaOpening.VectorU.ToBasicModelerVector
        Dim vecV = anaOpening.VectorV.ToBasicModelerVector
        Dim originPt As New Treegram.GeomKernel.BasicModeler.Point(0, 0, 0)
        Dim vecZ As New Treegram.GeomKernel.BasicModeler.Vector(0, 0, 1)

        'Get transparent models
        Dim transparModels = Treegram.GeomFunctions.Models.GetTransparentModels(anaOpening.Metaobject, True)
        Dim transparSurfaces As New Dictionary(Of Treegram.GeomKernel.BasicModeler.Surface, List(Of Double))
        For Each model In transparModels
            'Get the biggest surface from the model
            Dim transparFaces = Treegram.GeomFunctions.Models.GetModelAsFacesV2(model, True)
            If transparFaces.Count = 0 Then
                Continue For
            End If

            Dim maxSurf = transparFaces.OrderBy(Function(surf) CDbl(surf.Area3D)).Last

            'Get glazing position in local placement
            Dim minU = Double.PositiveInfinity
            Dim maxU = Double.NegativeInfinity
            Dim minZ = Double.PositiveInfinity
            Dim maxZ = Double.NegativeInfinity
            For Each pt In maxSurf.Points
                Dim ptInUV = Treegram.GeomKernel.BasicModeler.ChangeCoordPointInAnotherAxisSyst(pt, originPt, vecU, vecV, vecZ)
                If ptInUV.X < minU Then
                    minU = CDbl(Math.Round(ptInUV.X, 3))
                End If
                If ptInUV.X > maxU Then
                    maxU = CDbl(Math.Round(ptInUV.X, 3))
                End If
                If pt.Z < minZ Then
                    minZ = CDbl(Math.Round(pt.Z, 3))
                End If
                If pt.Z > maxZ Then
                    maxZ = CDbl(Math.Round(pt.Z, 3))
                End If
            Next

            'Compare it with glazing surfaces already treated
            Dim alreadyTreated = False
            For Each transparSurf In transparSurfaces
                Dim insideU, insideV As Boolean
                If transparSurf.Value(1) <= minU Or transparSurf.Value(0) >= maxU Then
                    insideU = False
                Else
                    insideU = True
                End If
                If transparSurf.Value(3) <= minZ Or transparSurf.Value(2) >= maxZ Then
                    insideV = False
                Else
                    insideV = True
                End If
                If insideU And insideV Then
                    alreadyTreated = True
                    Exit For
                End If
            Next
            If Not alreadyTreated Then
                transparSurfaces.Add(maxSurf, New List(Of Double) From {minU, maxU, minZ, maxZ})
            End If
        Next
        Dim surfList = transparSurfaces.Keys.ToList
        Return surfList
    End Function


    Private Sub GetOpeningOrientation(openings As List(Of AecOpening), myFile As FileObject)

        'ANALYSIS
        Dim exceptionList As New HashSet(Of String) From {"Succeeded"}
        For Each aecOpening In openings

            'GEO MEASURES
            Dim openingDirectionVec = aecOpening.VectorV.ToBasicModelerVector '.Reverse 'cf. IfcDoorStyleOperationEnum
            Dim openingAngle = (openingDirectionVec.Angle(myFile.North.Direction.ToBasicModelerVector, True) + 360) Mod 360
            Dim orientation As GeographicOrientation
            If openingAngle < 22.5 Then
                orientation = GeographicOrientation.N
            ElseIf openingAngle < 67.5 Then
                orientation = GeographicOrientation.NW
            ElseIf openingAngle < 112.5 Then
                orientation = GeographicOrientation.W
            ElseIf openingAngle < 157.5 Then
                orientation = GeographicOrientation.SW
            ElseIf openingAngle < 202.5 Then
                orientation = GeographicOrientation.S
            ElseIf openingAngle < 247.5 Then
                orientation = GeographicOrientation.SE
            ElseIf openingAngle < 292.5 Then
                orientation = GeographicOrientation.E
            ElseIf openingAngle < 337.5 Then
                orientation = GeographicOrientation.NE
            Else
                orientation = GeographicOrientation.N
            End If

            'SAVE INFOS
            'Dim spaceAnaAtt = anaOpening.AnalysisExtension.SmartAddAttribute("DimensionsAnalysis", Nothing, True)
            'Dim geoAtt = spaceAnaAtt.SmartAddAttribute("ByGeometry", Nothing)
            'geoAtt.SmartAddAttribute("GeoOrientation", orientation.ToString)
            'TgmPSet
            aecOpening.CompleteTgmPset("OrientationByTgm", orientation.ToString)

        Next

    End Sub
    Private Sub GetOpeningInsertPoint(openings As List(Of AecOpening), myFile As FileObject)
        Dim outputWs As Workspace = myFile.ProjWs.SmartAddWorkspace(myFile.ScanWs.Name + " - AdditionalGeometries", "AdditionalGeometries", True, "Scan Ifc", myFile.ScanWs) ' "San Ifc" -> To be reset at all scans
        If outputWs.GetAttribute("OriginPoint")?.GetAttribute("RelativeTo") Is Nothing Then
            outputWs.SetRelativeTo(myFile.ScanWs)
        End If

        'Treegram.GeomFunctions.Models.LoadGeometryModels(outputWs)

        'ANALYSIS
        Dim insertPointDico As New Dictionary(Of MetaObject, Tuple(Of Point3D, SharpDX.Color4, Double))
        For Each aecOpening In openings

            'Get Insert Point
            Dim insertPt = aecOpening.LocalOrigin()

            'Save in Tgm
            If insertPt IsNot Nothing Then
                Dim insertPtTgm = aecOpening.Metaobject.SmartAddExtension(aecOpening.Name + " - OpeningInsertPoint", M4D.Treegram.Core.Extensions.Enums.ExtensionType.AdditionalGeometries, outputWs.Name, "Tgm")
                insertPt.Write(insertPtTgm)

                If Treegram.GeomFunctions.Models.GetGeometryModels(insertPtTgm).Count = 0 Then
                    insertPointDico.Add(insertPtTgm, New Tuple(Of Point3D, SharpDX.Color4, Double)(insertPt, New SharpDX.Color4(1, 0.65, 0, 1), 0.05)) 'ORANGE
                End If
            End If
        Next

        'CREATE MODELS
        If insertPointDico.Count > 0 Then
            myFile.ProjWs.PushAllModifiedEntities()
            Treegram.GeomFunctions.Models.CreateCubeInTreegramFromPoint(outputWs, insertPointDico, "OpeningInsertPoint")
        End If

    End Sub
    Private Function GetOpeningSwingDirection(openings As List(Of AecOpening), myFile As FileObject) As Dictionary(Of MetaObject, Tuple(Of Curve3D, SharpDX.Color4, Double))
        Dim outputWs As Workspace = myFile.ProjWs.SmartAddWorkspace(myFile.ScanWs.Name + " - AdditionalGeometries", "AdditionalGeometries", True, "Scan Ifc", myFile.ScanWs) ' "San Ifc" -> To be reset at all scans
        If outputWs.GetAttribute("OriginPoint")?.GetAttribute("RelativeTo") Is Nothing Then
            outputWs.SetRelativeTo(myFile.ScanWs)
        End If

        'Treegram.GeomFunctions.Models.LoadGeometryModels(outputWs)
        'Treegram.GeomFunctions.Models.LoadGeometryModels(openings, True, True)

        'ANALYSIS
        Dim swingDirectionDico As New Dictionary(Of MetaObject, Tuple(Of Curve3D, SharpDX.Color4, Double))
        For Each aecOpening In openings

            'Get swing direction attribute
            Dim operationType As String = aecOpening.Metaobject.GetAttribute("OperationType")?.Value
            If operationType Is Nothing OrElse (operationType <> "SINGLE_SWING_LEFT" And operationType <> "SINGLE_SWING_RIGHT") Then
                Continue For
            End If

            'get axis system
            Dim originPt As New Treegram.GeomKernel.BasicModeler.Point(0, 0, 0)
            Dim xVec As New Treegram.GeomKernel.BasicModeler.Vector(1, 0, 0)
            Dim yVec As New Treegram.GeomKernel.BasicModeler.Vector(0, 1, 0)
            Dim zVec As New Treegram.GeomKernel.BasicModeler.Vector(0, 0, 1)

            'get local bounding box
            Dim distanceAbove = 0.02
            Dim vAverage = (aecOpening.MaxV + aecOpening.MinV) / 2
            Dim localPt1 As New Treegram.GeomKernel.BasicModeler.Point(aecOpening.MinU, vAverage, aecOpening.MinZ + 0.1)
            Dim localPt2 As New Treegram.GeomKernel.BasicModeler.Point(aecOpening.MaxU, vAverage, aecOpening.MinZ + 0.1)

            'get swing direction representation line
            Dim vec As Treegram.GeomKernel.BasicModeler.Vector
            Dim startLocalPt, endLocalPt As Treegram.GeomKernel.BasicModeler.Point
            Dim angle = aecOpening.VectorV.ToBasicModelerVector.Angle(aecOpening.VectorU.ToBasicModelerVector)
            If operationType = "SINGLE_SWING_LEFT" Then
                startLocalPt = localPt1
                vec = New Treegram.GeomKernel.BasicModeler.Vector(localPt1, localPt2)
                vec = vec.Rotate2(angle / 2)
                endLocalPt = vec.EndPoint
            ElseIf operationType = "SINGLE_SWING_RIGHT" Then
                startLocalPt = localPt2
                vec = New Treegram.GeomKernel.BasicModeler.Vector(localPt2, localPt1)
                vec = vec.Rotate2(-angle / 2)
                endLocalPt = vec.EndPoint
            Else
                Throw New Exception("Error")
            End If

            'get x and y vectors in local axis system
            Dim localAxisAngle = aecOpening.VectorU.ToBasicModelerVector.Angle(xVec)
            Dim x2Vec = xVec.Rotate2(-localAxisAngle)
            Dim y2Vec = yVec.Rotate2(-localAxisAngle)

            'get world bounding box
            Dim pt1 = Treegram.GeomKernel.BasicModeler.ChangeCoordPointInAnotherAxisSyst(startLocalPt, originPt, x2Vec, y2Vec, zVec)
            Dim pt2 = Treegram.GeomKernel.BasicModeler.ChangeCoordPointInAnotherAxisSyst(endLocalPt, originPt, x2Vec, y2Vec, zVec)
            Dim myLine As New Treegram.GeomKernel.BasicModeler.Line(pt1, pt2)

            'Save in Tgm
            If myLine IsNot Nothing Then
                Dim swingDirTgm = aecOpening.Metaobject.SmartAddExtension(aecOpening.Name + " - SwingDirection", M4D.Treegram.Core.Extensions.Enums.ExtensionType.AdditionalGeometries, outputWs.Name, "Tgm")

                If Treegram.GeomFunctions.Models.GetGeometryModels(swingDirTgm).Count = 0 Then
                    swingDirectionDico.Add(swingDirTgm, New Tuple(Of Curve3D, SharpDX.Color4, Double)(myLine.ToCurve3d, New SharpDX.Color4(1, 0.65, 0, 1), 0.01)) 'ORANGE
                End If
            End If
        Next

        'CREATE MODELS
        If swingDirectionDico.Count > 0 Then
            myFile.ProjWs.PushAllModifiedEntities()
            Treegram.GeomFunctions.Models.CreateCylinderInTreegramFromLine(outputWs, swingDirectionDico, "SwingDirection")
        End If

        Return swingDirectionDico
    End Function


    Private Function ComparisonsOutputTree(myFile As FileObject, Optional treeName As String = "Dimensions des Portes et Fenêtres") As Tree

        Dim myAnaVerifTree = myFile.ProjWs.SmartAddTree(treeName)
        Dim fileNode = myAnaVerifTree.SmartAddNode("File", myFile.FileTgm.Name)
        myAnaVerifTree.RemoveNode(fileNode)
        fileNode = myAnaVerifTree.SmartAddNode("File", myFile.FileTgm.Name)
        Dim compNode = fileNode.SmartAddNode("DimensionsAnalysis", Nothing,,, "Comparaison Attr vs Geom")

        'Fiabilite Ifc-Geo
        Dim widthTolerance = myFile.ReportInfos.OpeningAnalysis_widthTolerance
        Dim widthFiabilityNode = compNode.SmartAddNode("DimensionsAnalysis", Nothing,,, "Cohérence Largeurs")
        widthFiabilityNode.SmartAddNode("CoherenceIfcGeoWidth", CDbl(widthTolerance) / 100, AttributeType.Number, NodeOperator.LessOrEqual, "Ecart < " + widthTolerance.ToString + "cm")
        widthFiabilityNode.SmartAddNode("CoherenceIfcGeoWidth", CDbl(widthTolerance) / 100, AttributeType.Number, NodeOperator.Greater, "Ecart > " + widthTolerance.ToString + "cm")
        widthFiabilityNode.SmartAddNode("AnaWidth_NoIfcInfo", "True",,, "Info manquante dans l'ifc")

        'Fiabilite Ifc-Geo
        Dim heightTolerance = myFile.ReportInfos.OpeningAnalysis_heightTolerance
        Dim heightFiabilityNode = compNode.SmartAddNode("DimensionsAnalysis", Nothing,,, "Cohérence Hauteurs")
        heightFiabilityNode.SmartAddNode("CoherenceIfcGeoHeight", CDbl(heightTolerance) / 100, AttributeType.Number, NodeOperator.LessOrEqual, "Ecart < " + heightTolerance.ToString + "cm")
        heightFiabilityNode.SmartAddNode("CoherenceIfcGeoHeight", CDbl(heightTolerance) / 100, AttributeType.Number, NodeOperator.Greater, "Ecart > " + heightTolerance.ToString + "cm")
        heightFiabilityNode.SmartAddNode("AnaHeight_NoIfcInfo", "True",,, "Info manquante dans l'ifc")

        'Fiabilite Ifc-Geo
        Dim glazTolerance = myFile.ReportInfos.OpeningAnalysis_glazAreaTolerance
        Dim fiabilityNode = compNode.SmartAddNode("DimensionsAnalysis", Nothing,,, "Cohérence Surfaces Vitrées")
        fiabilityNode.SmartAddNode("CoherenceIfcGeoGlaArea", CDbl(glazTolerance) / 10000, AttributeType.Number, NodeOperator.LessOrEqual, "Ecart < " + glazTolerance.ToString + "cm2")
        fiabilityNode.SmartAddNode("CoherenceIfcGeoGlaArea", CDbl(glazTolerance) / 10000, AttributeType.Number, NodeOperator.Greater, "Ecart > " + glazTolerance.ToString + "cm2")
        fiabilityNode.SmartAddNode("AnaGla_NoIfcInfo", "True",,, "Info manquante dans l'ifc")

        'Géométries Additionnelles
        Dim addGeomNode = fileNode.SmartAddNode("workspace.type", "AdditionalGeometries",,, "Géométries Additionnelles")
        addGeomNode.SmartAddNode("object.name", "Glazing",,, "Surfaces Vitrées", True, True)
        addGeomNode.SmartAddNode("object.name", "OpeningProfile",,, "Surfaces Ouvertures", True, True)
        addGeomNode.SmartAddNode("object.name", "OpeningInsertPoint",,, "Points Insertion", True, True)
        addGeomNode.SmartAddNode("object.name", "SwingDirection",,, "Sens Ouverture", True, True)

        Return myAnaVerifTree
    End Function

    Private Function OrientationsOutputTree(myFile As FileObject, orientationsDico As Dictionary(Of GeographicOrientation, (openingNumber As Integer, glazedArea As Double)), Optional treeName As String = "Dimensions des Portes et Fenêtres") As Tree

        Dim myTree = myFile.ProjWs.SmartAddTree(treeName)
        Dim fileNode = myTree.SmartAddNode("File", myFile.FileTgm.Name)
        Dim catNode = fileNode.SmartAddNode("IsOpenable", True,,, "Portes et fenêtres")
        fileNode.SmartAddNode("object.name", "GeographicNorth",,, "Nord")
        Dim orientNode = catNode.SmartAddNode("GlazedAreaByTgm", 0.0, AttributeType.Number, NodeOperator.Greater, "Orientations")

        '--Colors
        Dim northColor As New Tuple(Of System.Windows.Media.Color, Single)(System.Windows.Media.Color.FromRgb(154, 255, 255), 0.7F)
        Dim eastColor As New Tuple(Of System.Windows.Media.Color, Single)(System.Windows.Media.Color.FromRgb(205, 255, 154), 0.7F)
        Dim southColor As New Tuple(Of System.Windows.Media.Color, Single)(System.Windows.Media.Color.FromRgb(255, 154, 154), 0.7F)
        Dim westColor As New Tuple(Of System.Windows.Media.Color, Single)(System.Windows.Media.Color.FromRgb(205, 154, 255), 0.7F)
        Dim northeastColor As New Tuple(Of System.Windows.Media.Color, Single)(System.Windows.Media.Color.FromRgb(154, 255, 179), 0.7F)
        Dim southeastColor As New Tuple(Of System.Windows.Media.Color, Single)(System.Windows.Media.Color.FromRgb(255, 230, 154), 0.7F)
        Dim southwestColor As New Tuple(Of System.Windows.Media.Color, Single)(System.Windows.Media.Color.FromRgb(255, 154, 230), 0.7F)
        Dim northwestColor As New Tuple(Of System.Windows.Media.Color, Single)(System.Windows.Media.Color.FromRgb(154, 179, 255), 0.7F)


        Dim chargedWs As New List(Of Workspace) From {myFile.ScanWs}
        Dim checkedNodes As New List(Of Node) From {fileNode}
        Dim savedColors As New Dictionary(Of Node, Tuple(Of System.Windows.Media.Color, Single))
        Dim nodeColors = New List(Of NodeColorTransparency) From {
            New NodeColorTransparency(fileNode, System.Windows.Media.Color.FromRgb(191, 191, 191), 0.1F)
        }

        For Each orientationPair In orientationsDico
            If orientationPair.Value.openingNumber > 0 Then
                Dim myNode = orientNode.SmartAddNode("OrientationByTgm", orientationPair.Key.ToString)
                checkedNodes.Add(myNode)

                Select Case orientationPair.Key
                    Case GeographicOrientation.E
                        nodeColors.Add(New NodeColorTransparency(myNode, eastColor.Item1, eastColor.Item2))
                    Case GeographicOrientation.N
                        nodeColors.Add(New NodeColorTransparency(myNode, northColor.Item1, northColor.Item2))
                    Case GeographicOrientation.NE
                        nodeColors.Add(New NodeColorTransparency(myNode, northeastColor.Item1, northeastColor.Item2))
                    Case GeographicOrientation.NW
                        nodeColors.Add(New NodeColorTransparency(myNode, northwestColor.Item1, northwestColor.Item2))
                    Case GeographicOrientation.S
                        nodeColors.Add(New NodeColorTransparency(myNode, southColor.Item1, southColor.Item2))
                    Case GeographicOrientation.SE
                        nodeColors.Add(New NodeColorTransparency(myNode, southeastColor.Item1, southeastColor.Item2))
                    Case GeographicOrientation.SW
                        nodeColors.Add(New NodeColorTransparency(myNode, southwestColor.Item1, southwestColor.Item2))
                    Case GeographicOrientation.W
                        nodeColors.Add(New NodeColorTransparency(myNode, westColor.Item1, westColor.Item2))
                End Select
            End If
        Next


        myFile.ProjWs.PushAllModifiedEntities()
        myTree.WriteColors(nodeColors)
        myFile.ProjWs.PushAllModifiedEntities()
        TgmView.CreateView(myFile.FileTgm.GetProject, myFile.FileTgm.Name & " - ORIENTATIONS", chargedWs, myTree, checkedNodes, nodeColors)

        Return myTree
    End Function

    Private Function DimensionsOutputTree(myFile As FileObject, dicoDimensions As Dictionary(Of Category, List(Of String)), dimensionsArrayDico As Dictionary(Of Category, (dimArray As Array, widthCount As Integer, heightCount As Integer)), Optional treeName As String = "Dimensions des Portes et Fenêtres") As Tree

        Dim myTree = myFile.ProjWs.SmartAddTree(treeName)
        Dim fileNode = myTree.SmartAddNode("File", myFile.FileTgm.Name)
        Dim catNode = fileNode.SmartAddNode("IsOpenable", True,,, "Portes et fenêtres")
        Dim opNode As Node = catNode.SmartAddNode("OpeningContainment", ".+",, NodeOperator.Different, "Sans objet hôte")
        opNode.IsRegex = True
        Dim dimNode As Node = catNode.SmartAddNode("DimensionsByTgm", Nothing,,, "Dimensions")
        'While dimNode.Nodes.Count > 0 'Reset
        '    dimNode.RemoveNode(dimNode.Nodes.First)
        'End While

        Dim chargedWs As New List(Of Workspace) From {myFile.ScanWs}
        Dim colorNodes As New List(Of NodeColorTransparency)
        Dim checkedNodesDim As New List(Of Node)

        Dim grey20 As New Tuple(Of System.Windows.Media.Color, Single)(System.Windows.Media.Color.FromRgb(191, 191, 191), 0.2)
        Dim red100 As New Tuple(Of System.Windows.Media.Color, Single)(System.Windows.Media.Color.FromRgb(255, 124, 124), 1.0)
        Dim Violet100 As New Tuple(Of System.Windows.Media.Color, Single)(System.Windows.Media.Color.FromRgb(150, 131, 236), 1.0)

        colorNodes.Add(New NodeColorTransparency(opNode, red100.Item1, red100.Item2))
        checkedNodesDim.Add(opNode)
        For Each category In dimensionsArrayDico.Keys
            If Not dicoDimensions.ContainsKey(category) Then
                Continue For
            End If
            Dim maxHeight, minheight, maxwidth, minwidth As Double
            If category = Category.Door Then
                maxHeight = myFile.ReportInfos.OpeningAnalysis_DoorMaxHeight
                minheight = myFile.ReportInfos.OpeningAnalysis_DoorMinHeight
                maxwidth = myFile.ReportInfos.OpeningAnalysis_DoorMaxWidth
                minwidth = myFile.ReportInfos.OpeningAnalysis_DoorMinWidth
            ElseIf category = Category.Window Then
                maxHeight = myFile.ReportInfos.OpeningAnalysis_WindowMaxHeight
                minheight = myFile.ReportInfos.OpeningAnalysis_WindowMinHeight
                maxwidth = myFile.ReportInfos.OpeningAnalysis_WindowMaxWidth
                minwidth = myFile.ReportInfos.OpeningAnalysis_WindowMinWidth
            End If

            Dim heightNode As Node = dimNode.SmartAddNode("CategoryByTgm", category.ToString,,, DicoObjType($"Ifc{category}"))
            Dim dimensionList = dicoDimensions(category).AsEnumerable.ToHashSet.ToList

            'reorder by width
            Dim dicotoreorder As New Dictionary(Of String, Double)
            For Each dimension In dimensionList
                Dim myHeight = Convert.ToDouble(dimension.Split("x")(1).Replace(".", ","))
                Dim myWidth = Convert.ToDouble(dimension.Split("x")(0).Replace(".", ","))
                dicotoreorder.Add(dimension, myWidth)
            Next
            Dim sorted = From pair In dicotoreorder
                         Order By pair.Value
            Dim sortedDictionary = sorted.ToDictionary(Function(p) p.Key, Function(p) p.Value)

            'create nodes
            For Each dimension In sortedDictionary.Keys
                Dim dimensionNode = heightNode.SmartAddNode("DimensionsByTgm", dimension,,, dimension)
                Dim myHeight = Convert.ToDouble(dimension.Split("x")(1).Replace(".", ","))
                Dim myWidth = Convert.ToDouble(dimension.Split("x")(0).Replace(".", ","))
                If myHeight < minheight * 100 Or myHeight > maxHeight * 100 Or myWidth < minwidth * 100 Or myWidth > maxwidth * 100 Then
                    colorNodes.Add(New NodeColorTransparency(dimensionNode, Violet100.Item1, Violet100.Item2))
                    checkedNodesDim.Add(dimensionNode)
                Else
                    colorNodes.Add(New NodeColorTransparency(dimensionNode, grey20.Item1, grey20.Item2))
                    checkedNodesDim.Add(dimensionNode)
                End If
            Next
        Next

        myFile.ProjWs.PushAllModifiedEntities()
        myTree.WriteColors(colorNodes)
        myFile.ProjWs.PushAllModifiedEntities()
        TgmView.CreateView(myFile.FileTgm.GetProject, myFile.FileTgm.Name & " - OPENING DIMENSIONS", chargedWs, myTree, checkedNodesDim, colorNodes)

        Return myTree
    End Function
End Class
