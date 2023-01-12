Imports M4D.Production.Core
Imports M4D.Treegram.Core.Entities
Imports M4D.Treegram.Core.Kernel
Imports M4D.Treegram.Core.Extensions.Kernel
Imports M4D.Treegram.Core.Extensions.Entities
Imports Treegram.Bam.Libraries.AuditFile
Imports DevExpress.Spreadsheet
Imports DevExpress.Export.Xl
Imports Treegram.Bam.Libraries.AEC
Imports Treegram.Bam.Functions.AEC
Imports Treegram.ConstructionManagement
Imports Treegram.Bam.Functions

Public Class Create_IfcInitialAuditProdScript
    Inherits ProdScript
    Public Sub New()
        MyBase.New("IFC :: AUDIT : Initial") 'ATTENTION A RENOMMER AVEC PRECAUTION CAR LANCÉ AVEC LE SCAN !!!!!!!!!!
        AddAction(New Create_IfcInitialAudit())
    End Sub
End Class
Public Class Create_IfcInitialAudit
    Inherits ProdAction
    Public Sub New()
        Name = "IFC :: AUDIT : Initial (ProdAction)"
        PartOfScript = True
    End Sub

    Public Overrides Sub CreateGuiInputTree()
        InputTree = TempWorkspace.AddTree("--- Temporary Input Tree ---")
        Dim fileNode = InputTree.AddNode("object.type", "File")
        fileNode.Description = "Files"
        SelectYourSetsOfInput.Add("FichiersIFC", {fileNode}.ToList())
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
                'POUR UN FILTRAGE DE TYPE INFORMATIQUE : WORKSPACE, FILE, METAOBJET
                If inputmObj.GetTgmType <> "File" OrElse inputmObj.GetAttribute("Extension").Value.ToString.ToLower <> "ifc" Then Continue For 'Throw New Exception("You must drag and drop file objects")
                launchTree.SmartAddNode("object.id", inputmObj.Id.ToString,,, inputmObj.Name)
            Next
        End If

        Return launchTree
    End Function

    <ActionMethod()>
    Public Function MyMethod(FichiersIFC As MultipleElements) As ActionResult

        ' Récupération des inputs
        Dim tgmFiles = GetInputAsMetaObjects(FichiersIFC.Source).ToHashSet 'To get rid of duplicates
        If tgmFiles.Count = 0 Then Throw New Exception("Inputs missing")

        'AUDIT
        For Each fileTgm In tgmFiles

            If fileTgm.GetAttribute("Extension").Value.ToString.ToLower <> "ifc" Then
                Throw New Exception("this file is not an ifc")
            End If
            Dim myAudit As New IfcFileObject(fileTgm)
            myAudit.PrepareFileWorkspacesAnalysis() 'Pour initialiser les ws et gérer les différentes révisions !!!

            GetAuditFromFile(myAudit)
        Next

        Return New SucceededActionResult("Succeeded", System.Drawing.Color.Green, New List(Of Element) From {FichiersIFC})
    End Function

    Public Shared Function GetAuditFromFile(myAudit As FileObject) As Workbook

        myAudit.CreateAuditWorkbook()
        'myAudit.ReportInfos.percentageMax = CDbl(myAudit.GetOrInsertSettingsWSheet("AUDIT IFC").Range("ifcproxy").Value.NumericValue)
        Try 'Classer les nouveaux settings chronologiquement, le plus ancien en premier
            'myAudit.ReportInfos.fileUnitySettings = myAudit.GetOrInsertSettingsWSheet("AUDIT IFC").Range("unite").Value.ToString
            myAudit.ReportInfos.settingsTemplate = myAudit.GetOrInsertSettingsWSheet("PRESENTATION").Range("template").Value.ToString
        Catch ex As Exception
        End Try

        ''FILTER
        Workspace.GetTemp.SmartAddTree("ALL").Filter(False, New List(Of Workspace) From {myAudit.ScanWs}).RunSynchronously()

        'FILE INFOS
        CompleteAuditFileInfos(myAudit)

        'IFC STRUCTURE
        CompleteAuditStructure(myAudit)

        'TYPES
        CompleteAuditTypes(myAudit)

        'LAYERS
        CompleteAuditLayers(myAudit)

        'FILL REPORT
        myAudit.ReportInfos.CompleteCriteria(1)
        myAudit.ReportInfos.CompleteCriteria(2)
        myAudit.ReportInfos.CompleteCriteria(3)
        myAudit.ReportInfos.CompleteCriteria(51)
        myAudit.ReportInfos.CompleteCriteria(52)
        myAudit.ReportInfos.CompleteCriteria(53)
        myAudit.ReportInfos.CompleteCriteria(54)
        myAudit.ReportInfos.CompleteCriteria(57)
        myAudit.ReportInfos.CompleteCriteria(63)
        myAudit.ReportInfos.CompleteCriteria(66)
        myAudit.ReportInfos.FillActivity()
        myAudit.ReportInfos.CompleteReportInfos()

        'SAVE
        myAudit.ProjWs.PushAllModifiedEntities()
        myAudit.AuditWorkbook.Calculate()
        myAudit.AuditWorkbook.Worksheets.ActiveWorksheet = myAudit.ReportWorksheet
        myAudit.SaveAuditWorkbook()

        Return myAudit.AuditWorkbook
    End Function

    Private Shared Sub CompleteAuditFileInfos(myAudit As IfcFileObject)
#Region "GET INFOS"

        Dim unitsAtt = myAudit.FileTgm.GetAttribute("Units")
        myAudit.ReportInfos.fileSizeInfo = Math.Round(CDbl(myAudit.FileTgm.GetAttribute("Size").Value), 1)
        Dim softSt = CStr(myAudit.FileTgm.GetAttribute("Software").Value)
        'Dim sharedPosiAtt = myAudit.originPtTgm.GetAttribute("SharedPosition")
        'myAudit.ReportInfos.X = CDbl(sharedPosiAtt.GetAttribute("E/W").Value)
        'myAudit.ReportInfos.Y = CDbl(sharedPosiAtt.GetAttribute("N/S").Value)
        'Dim projElevDb = CDbl(sharedPosiAtt.GetAttribute("Elevation").Value)
        'myAudit.ReportInfos.Angle = CDbl(sharedPosiAtt.GetAttribute("Angle").Value)
        Dim fileGeoRef = myAudit.Georeferencing

        'COMPLETE BUILDINGS AND STOREYS REFERENCES DICO
        Dim buildAndStoReferencesDico As New Dictionary(Of MetaObject, Dictionary(Of MetaObject, List(Of MetaObject)))
        Dim buildings = myAudit.SiteBuildings
        Dim buildDico = myAudit.GetBuildings(buildings, myAudit._3dObjects)
        For Each buildingTgm In buildings
            Dim buildObjs = buildDico(buildingTgm)
            'Traité le cas si buildingTgm = Nothing...

            Dim storeys = myAudit.BuildingStoreys(buildingTgm)
            CommonFunctions.SortLevelsByElevation(storeys)
            Dim stoDico = myAudit.GetStoreys(storeys, buildObjs, False) 'ATTENTION AVEC LA RÉSURSIVITÉ !!!!!

            Dim storeysDico As New Dictionary(Of MetaObject, List(Of MetaObject))
            For Each storeyTgm In storeys
                Dim stoObjs = stoDico(storeyTgm)
                'Traité le cas si storeyTgm = Nothing...

                storeysDico.Add(storeyTgm, stoObjs)
            Next
            buildAndStoReferencesDico.Add(buildingTgm, storeysDico)
        Next
        Dim layersDico = myAudit.GetLayers(myAudit._3dObjects)

        'GET FILE INFOS
        Dim _infosProject = True
        If myAudit.ScanWs.GetMetaObjects(, "IfcSite").Count = 0 Or myAudit.ScanWs.GetMetaObjects(, "IfcBuilding").Count = 0 Then
            _infosProject = False
        Else
            Dim unwantedNameList As New List(Of String) From {"", "unknown", "default"}
            For Each site In myAudit.ScanWs.GetMetaObjects(, "IfcSite")
                If unwantedNameList.Contains(site.Name.Trim.ToLower) Then
                    _infosProject = False
                    Exit For
                End If
            Next
            If _infosProject Then
                For Each build In myAudit.ScanWs.GetMetaObjects(, "IfcBuilding")
                    If unwantedNameList.Contains(build.Name.Trim.ToLower) Then
                        _infosProject = False
                        Exit For
                    End If
                Next
            End If
        End If
        myAudit.ReportInfos.HasProjectInfos = _infosProject

#End Region

#Region "FILL EXCEL"
        Dim fichierWsheet As Worksheet = myAudit.InsertProdActionTemplateInAuditWbook("AuditIFC_Template", "FICHIER")
        'fichierWsheet.Columns("G").Visible = False
        'fichierWsheet.Columns("H").Visible = False

        'FICHIER
        Dim fileStartLine = fichierWsheet.Range("fichier").BottomRowIndex
        Dim fileCol = fichierWsheet.Range("fichier").LeftColumnIndex
        fichierWsheet.Cells(fileStartLine + 1, 2).Value = myAudit.FileTgm.Name
        fichierWsheet.Cells(fileStartLine + 2, 2).Value = Math.Round(myAudit.ReportInfos.fileSizeInfo, 1).ToString + " Mo"
        fichierWsheet.Cells(fileStartLine + 3, 2).Value = softSt

        'LOCALISATION        <-- Pour le moment fonctionne que pour un seul bâtiment par fichier !!!
        Dim locaStartLine = fichierWsheet.Range("localisation").BottomRowIndex
        If buildAndStoReferencesDico.Keys(0) IsNot Nothing Then
            fichierWsheet.Cells(locaStartLine + 1, 2).Value = buildAndStoReferencesDico.Keys(0).SmartGetAttribute("BuildingAddress/_\Country")?.Value?.ToString
            fichierWsheet.Cells(locaStartLine + 2, 2).Value = buildAndStoReferencesDico.Keys(0).SmartGetAttribute("BuildingAddress/_\Region")?.Value?.ToString
            fichierWsheet.Cells(locaStartLine + 3, 2).Value = buildAndStoReferencesDico.Keys(0).SmartGetAttribute("BuildingAddress/_\Town")?.Value?.ToString
            fichierWsheet.Cells(locaStartLine + 4, 2).Value = buildAndStoReferencesDico.Keys(0).SmartGetAttribute("BuildingAddress/_\Address")?.Value?.ToString
            fichierWsheet.Cells(locaStartLine + 5, 2).Value = buildAndStoReferencesDico.Keys(0).SmartGetAttribute("BuildingAddress/_\Description")?.Value?.ToString
        End If

        'DONNEE
        Dim dataStartLine = fichierWsheet.Range("donnee").BottomRowIndex
        fichierWsheet.Cells(dataStartLine + 1, 2).Value = buildAndStoReferencesDico.Keys.Count
        fichierWsheet.Cells(dataStartLine + 2, 2).Value = buildAndStoReferencesDico.Values.SelectMany(Function(i) i).ToList.Count
        fichierWsheet.Cells(dataStartLine + 3, 2).Value = layersDico.Keys.Where(Function(o) o <> "<NONE>").Count
        fichierWsheet.Cells(dataStartLine + 4, 2).Value = myAudit._3dObjects.Count
        fichierWsheet.Cells(dataStartLine + 5, 2).Value = myAudit.AddedObjs
        fichierWsheet.Cells(dataStartLine + 6, 2).Value = myAudit.ModifiedObjs
        fichierWsheet.Cells(dataStartLine + 7, 2).Value = myAudit.DeletedObjs

        'GEOREF
        Dim georefStartLine = fichierWsheet.Range("georef").BottomRowIndex
        Dim georefCol = fichierWsheet.Range("georef").LeftColumnIndex
        fichierWsheet.Cells(georefStartLine + 1, georefCol + 1).Value = fileGeoRef.Y
        fichierWsheet.Cells(georefStartLine + 2, georefCol + 1).Value = fileGeoRef.X
        fichierWsheet.Cells(georefStartLine + 3, georefCol + 1).Value = fileGeoRef.Z
        fichierWsheet.Cells(georefStartLine + 4, georefCol + 1).Value = fileGeoRef.Angle
        Try
            fichierWsheet.Hyperlinks.Add(fichierWsheet.Cells(georefStartLine + 5, georefCol), fileGeoRef.MapUrl, True, "'--> Localiser dans Google Maps <--")
        Catch ex As Exception
        End Try

        'PLATFORM
        Dim platformStartLine = fichierWsheet.Range("plateforme").BottomRowIndex
        Dim platformCol = fichierWsheet.Range("plateforme").LeftColumnIndex
        fichierWsheet.Cells(platformStartLine + 1, platformCol + 1).Value = myAudit.FileTgm.GetAttribute("Platform")?.Value?.ToString
        fichierWsheet.Cells(platformStartLine + 2, platformCol + 1).Value = myAudit.FileTgm.SmartGetAttribute("Platform/_\PlatformFileDate")?.Value?.ToString

        'UNITES
        Dim unitsStartLine = fichierWsheet.Range("unites").BottomRowIndex
        Dim unitsCol = fichierWsheet.Range("unites").LeftColumnIndex

        Dim k As Integer
        myAudit.ReportInfos.uncomparedUnits = New List(Of String)
        myAudit.ReportInfos.invalidUnits = New List(Of String)
        myAudit.ReportInfos.validUnits = New List(Of String)
        For k = 0 To unitsAtt.Attributes.Count - 1
            Dim unitAtt = unitsAtt.Attributes(k)
            Dim myUnitRow As Integer
            Dim maxNameRange As CellRange = Nothing
            Dim maxValueRange As CellRange = Nothing
            Dim valueRange, refRange, comparisonRange As CellRange

            'LENGTH
            If unitAtt.Name.ToLower = "length" Or unitAtt.Name.ToLower = "longueur" Then
                myUnitRow = unitsStartLine + 1
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
                myUnitRow = unitsStartLine + 2
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


        '---mise en forme
        'Dim unitsRange = fichierWsheet.Range(FileObject.GetSpreadsheetColumnName(unitsCol) + (unitsStartLine + 1).ToString + ":" + FileObject.GetSpreadsheetColumnName(unitsCol + 1) + (unitsStartLine + k + 1).ToString)
        'unitsRange.Borders.SetOutsideBorders(System.Drawing.Color.Black, BorderLineStyle.Thin)
        'unitsRange.Borders.InsideVerticalBorders.LineStyle = BorderLineStyle.Thin
        'unitsRange.Borders.InsideHorizontalBorders.LineStyle = BorderLineStyle.Thin

        ''fichierWsheet.Range(auditfile.getspreadsheetColumnName(unitsCol + 1) + "1").EntireColumn.AutoFit()
        'fichierWsheet.Columns(unitsCol + 1).AutoFit()
        fichierWsheet.Columns(fileCol + 1).AutoFit()
#End Region

#Region "CREATE TREE"
        Dim myTree = myAudit.AuditTreesWs.SmartAddTree("Audit IFC - FICHIER")
        Dim fileNode = myTree.SmartAddNode("File", myAudit.FileTgm.Name)
        myTree.RemoveNode(fileNode)
        fileNode = myTree.SmartAddNode("File", myAudit.FileTgm.Name)

        fileNode.SmartAddNode("object.type", "IfcSite")
        fileNode.SmartAddNode("object.type", "IfcBuilding")
        fileNode.SmartAddNode("object.type", "IfcBuildingStorey")

        'Dim siteNode = fileNode.SmartAddNode("IfcSite", myAudit.siteTgm.Name)
        'For Each buildTgm In buildAndStoReferencesDico.Keys
        '    Dim buildNode = siteNode.SmartAddNode("IfcBuilding", buildTgm.Name)
        '    Dim buildObjs = buildAndStoReferencesDico(buildTgm)
        '    For Each storeyTgm In buildObjs.Keys
        '        Dim stoNode = buildNode.SmartAddNode("IfcBuildingStorey", storeyTgm.Name)
        '        stoNode.SmartAddNode("object.type", "IfcBuildingStorey")
        '    Next
        'Next
        fileNode.SmartAddNode("object.name", myAudit.OriginPtTgm.Name)
        fileNode.SmartAddNode("object.name", myAudit.North.Name)
        fileNode.SmartAddNode("object.type", myAudit.FileTgm.GetTgmType)

#End Region
    End Sub

    Private Shared Sub CompleteAuditStructure(myAudit As IfcFileObject)
#Region "GET INFOS"
        Dim elevAttName As String = "GlobalElevation"

        'COMPLETE TYPES AND NAMES DICO
        Dim sites = myAudit.ScanWs.GetMetaObjects(, "IfcSite").ToList
        If sites.Count <> 1 Then
            myAudit.ReportInfos.structurationCorrect = False
        End If
        Dim remainingBuildings = myAudit.ScanWs.GetMetaObjects(, "IfcBuilding").ToList
        Dim remainingStoreys = myAudit.ScanWs.GetMetaObjects(, "IfcBuildingStorey").ToList

        'deal with sites children
        Dim auditDico As New Dictionary(Of String, SortedDictionary(Of String, List(Of MetaObject)))
        For Each siteTgm In sites
            Dim childBuildings = siteTgm.GetChildren("IfcBuilding").ToList
            If childBuildings.Count > 0 Then
                Dim buildDico As New SortedDictionary(Of String, List(Of MetaObject))

                For Each childBuildTgm In childBuildings
                    Dim childStoreys = childBuildTgm.GetChildren("IfcBuildingStorey").ToList
                    If childStoreys.Count > 0 Then
                        For Each childStoTgm In childStoreys
                            remainingStoreys.Remove(childStoTgm)
                        Next
                        CommonFunctions.SortLevelsByElevation(childStoreys, elevAttName)
                        buildDico.Add(childBuildTgm.Name, childStoreys)
                    Else
                        buildDico.Add(childBuildTgm.Name, Nothing)
                    End If
                    remainingBuildings.Remove(childBuildTgm)
                Next
                auditDico.Add(siteTgm.Name, buildDico)
            Else
                auditDico.Add(siteTgm.Name, Nothing)
            End If
        Next
        'deal with remaining buildings
        If remainingBuildings.Count <> 0 Then
            Dim i As Integer = 0
            Dim buildDico As New SortedDictionary(Of String, List(Of MetaObject))
            For i = remainingBuildings.Count - 1 To 0 Step -1
                Dim buildTgm = remainingBuildings(i)
                Dim storeys = buildTgm.GetChildren("IfcBuildingStorey").ToList
                If storeys.Count > 0 Then
                    For Each childStoTgm In storeys
                        remainingStoreys.Remove(childStoTgm)
                    Next
                    CommonFunctions.SortLevelsByElevation(storeys, elevAttName)
                    buildDico.Add(buildTgm.Name, storeys)
                Else
                    buildDico.Add(buildTgm.Name, Nothing)
                End If
                remainingBuildings.Remove(buildTgm)

            Next
            auditDico.Add("<NONE>", buildDico)
        End If
        'deal with remaining storeys
        If remainingStoreys.Count <> 0 Then
            Dim buildDico As New SortedDictionary(Of String, List(Of MetaObject))
            CommonFunctions.SortLevelsByElevation(remainingStoreys, elevAttName)
            buildDico.Add("<NONE>", remainingStoreys)
            auditDico.Add("<NONE>", buildDico)
        End If

        ''deal with project references
        'Dim refDico As New SortedDictionary(Of String, List(Of MetaObject))
        'If myAudit.ProjRefWs IsNot Nothing Then
        '    Dim tgmBuildings = myAudit.ProjRefWs.GetMetaObjects(, "TgmBuilding")
        '    For Each tgmBuildTgm In tgmBuildings
        '        Dim tgmStoreys = tgmBuildTgm.GetChildren("TgmBuildingStorey").Where(Function(o) o.GetAttribute("IsUsed").Value).ToList
        '        refDico.Add(tgmBuildTgm.Name, tgmStoreys)
        '    Next
        'End If

#End Region

#Region "FILL EXCEL"
        Dim structureWsheet As Worksheet = myAudit.InsertProdActionTemplateInAuditWbook("AuditIFC_Template", "STRUCTURE IFC")
        structureWsheet.Columns("D").ColumnWidth = 0.0
        structureWsheet.Columns("E").ColumnWidth = 0.0
        structureWsheet.Columns("G").ColumnWidth = 0.0
        structureWsheet.Columns("H").ColumnWidth = 0.0
        structureWsheet.Columns("J").ColumnWidth = 0.0
        structureWsheet.Columns("K").ColumnWidth = 0.0

        Dim indice_start_line = 4
        Dim buildingsNumber = auditDico.Where(Function(n) n.Value IsNot Nothing).SelectMany(Function(o) o.Value).Count
        Dim storeysNumber = auditDico.Where(Function(n) n.Value IsNot Nothing).SelectMany(Function(o) o.Value).Where(Function(n) n.Value IsNot Nothing).SelectMany(Function(o) o.Value).Count

        'FILL ARRAY
        Dim indice_next_line = 0
        myAudit.ReportInfos.goodLvls = True
        Dim totalRows = auditDico.Keys.Count + buildingsNumber + storeysNumber
        For Each siteSt In auditDico.Keys
            Dim siteStructure = auditDico(siteSt)

            'Fill SITES
            If siteSt = "" Then siteSt = "Unknown"
            structureWsheet.Cells(indice_start_line + indice_next_line, 1).Value = siteSt

            '---Mise en forme
            Dim typeRange = structureWsheet.Range("B" + (indice_start_line + indice_next_line + 1).ToString + ":L" + (indice_start_line + indice_next_line + 1).ToString)
            typeRange.FillColor = XlColor.FromTheme(XlThemeColor.Dark1, 0.7)
            typeRange.Font.Bold = True
            typeRange.Borders.SetOutsideBorders(System.Drawing.Color.Black, BorderLineStyle.Thin)
            typeRange.Borders.InsideVerticalBorders.LineStyle = BorderLineStyle.Thin
            typeRange.Borders.InsideHorizontalBorders.LineStyle = BorderLineStyle.Thin

            '---If nothing
            If siteStructure Is Nothing Then
                structureWsheet.Cells(indice_start_line + indice_next_line, 2).Value = "-> Pas d'enfants"
                structureWsheet.Range("C" + (indice_start_line + indice_next_line).ToString).Font.Color = System.Drawing.Color.Red
                indice_next_line += 1
                myAudit.ReportInfos.structurationCorrect = False
                Continue For
            End If
            indice_next_line += 1

            Dim indice_layer_start_line As Integer = indice_next_line
            For Each buildingSt In siteStructure.Keys
                Dim buildingStructure = siteStructure(buildingSt)

                'Fill BUILDINGS
                If buildingSt = "" Then buildingSt = "Unknown"
                structureWsheet.Cells(indice_start_line + indice_next_line, 2).Value = buildingSt

                '---Mise en forme
                Dim typeRange2 = structureWsheet.Range("B" + (indice_start_line + indice_next_line + 1).ToString + ":L" + (indice_start_line + indice_next_line + 1).ToString)
                typeRange2.FillColor = XlColor.FromTheme(XlThemeColor.Dark1, 0.9)
                typeRange2.Font.Bold = True
                typeRange2.Borders.SetOutsideBorders(System.Drawing.Color.Black, BorderLineStyle.Thin)
                typeRange2.Borders.InsideVerticalBorders.LineStyle = BorderLineStyle.Thin
                typeRange2.Borders.InsideHorizontalBorders.LineStyle = BorderLineStyle.Thin

                '---If nothing
                If siteSt = "<NONE>" Then
                    structureWsheet.Cells(indice_start_line + indice_next_line, 1).Value = "Pas de parent <-"
                    structureWsheet.Range("B" + (indice_start_line + indice_next_line).ToString).Font.Color = System.Drawing.Color.Red
                    myAudit.ReportInfos.structurationCorrect = False
                End If
                If buildingStructure Is Nothing Then
                    structureWsheet.Cells(indice_start_line + indice_next_line, 5).Value = "-> Pas d'enfants"
                    structureWsheet.Range("F" + (indice_start_line + indice_next_line).ToString).Font.Color = System.Drawing.Color.Red
                    indice_next_line += 1
                    myAudit.ReportInfos.structurationCorrect = False
                    Continue For
                End If
                indice_next_line += 1

                Dim lastElevation As Double? = Nothing
                Dim indice_layer_start_line2 As Integer = indice_next_line
                For Each storeyTgm In buildingStructure

                    '---Rapport check
                    Dim myElevation = storeyTgm.GetAttribute(elevAttName, True).Value
                    If lastElevation IsNot Nothing AndAlso myElevation - lastElevation < myAudit.ReportInfos.lvlEcartMin Then
                        myAudit.ReportInfos.goodLvls = False
                    End If
                    lastElevation = myElevation

                    'Fill STOREYS
                    'dataArray(indice_next_line, 4) = storeyTgm.Name
                    'dataArray(indice_next_line, 7) = BetaTgmFcts.GetAttribute(storeyTgm, elevAttName, True).value
                    'dataArray(indice_next_line, 10) = storeyTgm.GetChildren().Count
                    structureWsheet.Cells(indice_start_line + indice_next_line, 5).Value = storeyTgm.Name
                    structureWsheet.Cells(indice_start_line + indice_next_line, 8).Value = CDbl(Math.Round(myElevation, 3)) 'Rounded to mm
                    structureWsheet.Cells(indice_start_line + indice_next_line, 11).Value = CInt(storeyTgm.GetChildren().Count)

                    '---If nothing
                    If buildingSt = "<NONE>" Then
                        structureWsheet.Cells(indice_start_line + indice_next_line, 2).Value = "Pas de parent <-"
                        structureWsheet.Range("C" + (indice_start_line + indice_next_line).ToString).Font.Color = System.Drawing.Color.Red
                        myAudit.ReportInfos.structurationCorrect = False
                    End If

                    indice_next_line += 1
                Next

                '---Mise en forme
                If indice_next_line <> indice_layer_start_line2 Then
                    Dim namesRange = structureWsheet.Range("B" + (indice_start_line + indice_layer_start_line2 + 1).ToString + ":L" + (indice_start_line + indice_next_line).ToString)
                    namesRange.Borders.SetOutsideBorders(System.Drawing.Color.Black, BorderLineStyle.Thin)
                    namesRange.Borders.InsideVerticalBorders.LineStyle = BorderLineStyle.Thin
                    namesRange.GroupRows(False)
                End If
            Next

            '---Mise en forme
            If indice_next_line <> indice_layer_start_line Then
                Dim namesRange = structureWsheet.Range("B" + (indice_start_line + indice_layer_start_line + 1).ToString + ":L" + (indice_start_line + indice_next_line).ToString)
                namesRange.Borders.SetOutsideBorders(System.Drawing.Color.Black, BorderLineStyle.Thin)
                namesRange.Borders.InsideVerticalBorders.LineStyle = BorderLineStyle.Thin
                namesRange.GroupRows(False)
            End If
        Next

        '---Mise en forme
        structureWsheet.Columns("B:C").AutoFit()
        structureWsheet.Columns("F").AutoFit()
#End Region

#Region "CREATE TREE"
        Dim myTree = myAudit.AuditTreesWs.SmartAddTree("Audit IFC - STRUCTURE IFC")
        Dim fileNode = myTree.SmartAddNode("File", myAudit.FileTgm.Name)
        myTree.RemoveNode(fileNode)
        fileNode = myTree.SmartAddNode("File", myAudit.FileTgm.Name)

        For Each siteSt In auditDico.Keys
            Dim siteStructure = auditDico(siteSt)
            Dim siteNode = fileNode.SmartAddNode("IfcSite", siteSt,,, "IfcSite : " + siteSt)
            If siteStructure Is Nothing Then Continue For

            For Each buildSt In siteStructure.Keys
                Dim buildStructure = siteStructure(buildSt)
                Dim buildNode = siteNode.SmartAddNode("IfcBuilding", buildSt,,, "IfcBuilding : " + buildSt)
                If buildStructure Is Nothing Then Continue For

                For Each storeyTgm In buildStructure
                    Dim storeyNode = buildNode.SmartAddNode("IfcBuildingStorey", storeyTgm.Name)
                Next
            Next
        Next
#End Region
    End Sub

    Private Shared Sub CompleteAuditTypes(myAudit As IfcFileObject)
#Region "GET INFOS"
        'COMPLETE TYPES AND NAMES DICO
        Dim typesAndNamesDico As New SortedDictionary(Of String, SortedDictionary(Of String, List(Of MetaObject)))
        Dim typesDico = myAudit.GetTypes(myAudit._3dObjects)
        Dim namesCount = 0
        For Each typeSt In typesDico.Keys
            Dim typedObjsList = typesDico(typeSt)
            Dim namesDico = myAudit.GetNames(typedObjsList)
            typesAndNamesDico.Add(typeSt, namesDico)
            namesCount += namesDico.Keys.Count
        Next
        Dim ifcProxyObjs = 0
        If typesDico.ContainsKey("IfcBuildingElementProxy") Then
            ifcProxyObjs = typesDico("IfcBuildingElementProxy").Count
        End If
        Dim totalObjs = typesDico.SelectMany(Function(o) o.Value).Count
        myAudit.ReportInfos.percentageGenerics = Math.Round(ifcProxyObjs / totalObjs, 3)
#End Region

#Region "FILL EXCEL"
        Dim typeWsheet As Worksheet = myAudit.InsertProdActionTemplateInAuditWbook("AuditIFC_Template", "TYPES IFC")
        Dim typeStartLine = 5
        typeWsheet.Cells(typeStartLine - 1, 1).Value = typesAndNamesDico.Keys.Count
        typeWsheet.Cells(typeStartLine - 1, 3).Value = typesAndNamesDico.SelectMany(Function(o) o.Value).SelectMany(Function(o) o.Value).Count

        'FILL ARRAY
        Dim indice_next_line = 0
        Dim indice_type_start_line As Integer
        'Dim dataArray(typesAndNamesDico.Keys.Count + namesCount - 1, 2) As Object
        For Each typeSt In typesAndNamesDico.Keys
            Dim namesDico = typesAndNamesDico(typeSt)
            Dim subTotalObjs = namesDico.Select(Function(o) o.Value.Count).Sum

            'Fill TYPES
            typeWsheet.Range("B" + (typeStartLine + indice_next_line + 1).ToString).Value = typeSt
            typeWsheet.Range("D" + (typeStartLine + indice_next_line + 1).ToString).Value = subTotalObjs
            'dataArray(indice_next_line, 0) = typeSt
            'dataArray(indice_next_line, 2) = subTotalObjs

            '---Mise en forme
            Dim typeRange = typeWsheet.Range("B" + (typeStartLine + indice_next_line + 1).ToString + ":D" + (typeStartLine + indice_next_line + 1).ToString)
            typeRange.FillColor = XlColor.FromTheme(XlThemeColor.Dark1, 0.7)
            typeRange.Font.Bold = True
            typeRange.Borders.SetOutsideBorders(System.Drawing.Color.Black, BorderLineStyle.Thin)
            typeRange.Borders.InsideVerticalBorders.LineStyle = BorderLineStyle.Thin
            typeRange.Borders.InsideHorizontalBorders.LineStyle = BorderLineStyle.Thin

            indice_next_line += 1
            indice_type_start_line = indice_next_line
            For Each namesSt In namesDico.Keys
                'Fill NAMES
                typeWsheet.Range("C" + (typeStartLine + indice_next_line + 1).ToString).Value = namesSt
                typeWsheet.Range("D" + (typeStartLine + indice_next_line + 1).ToString).Value = namesDico(namesSt).Count
                'dataArray(indice_next_line, 1) = namesSt
                'dataArray(indice_next_line, 2) = namesDico(namesSt).Count
                indice_next_line += 1
                '---Mise en forme

            Next

            '---Mise en forme
            Dim namesRange = typeWsheet.Range("B" + (typeStartLine + indice_type_start_line + 1).ToString + ":D" + (typeStartLine + indice_next_line).ToString)
            namesRange.Borders.SetOutsideBorders(System.Drawing.Color.Black, BorderLineStyle.Thin)
            namesRange.Borders.InsideVerticalBorders.LineStyle = BorderLineStyle.Thin
            namesRange.GroupRows(False)
        Next

        ''FILL EXCEL
        'typeWsheet.Range("B" + typeStartLine.ToString).Resize(typesAndNamesDico.Keys.Count + namesCount, 3).Value = dataArray

#End Region

#Region "CREATE TREE"
        Dim myTree = myAudit.AuditTreesWs.SmartAddTree("Audit IFC - TYPES IFC")
        Dim fileNode = myTree.SmartAddNode("File", myAudit.FileTgm.Name)
        myTree.RemoveNode(fileNode)
        fileNode = myTree.SmartAddNode("File", myAudit.FileTgm.Name)

        For Each typeSt In typesAndNamesDico.Keys
            If typeSt = "<NONE>" Then Continue For
            Dim typedNode = fileNode.SmartAddNode("object.type", typeSt)
            'Dim typedNames = typesAndNamesDico(typeSt)
            'For Each nameSt In typedNames.Keys 'A voir si ça n'allourdit pas trop
            '    Dim namedNode = typedNode.SmartAddNode("object.name", nameSt)
            'Next
        Next
#End Region
    End Sub

    Private Shared Sub CompleteAuditLayers(myAudit As IfcFileObject)
#Region "GET INFOS"
        'COMPLETE TYPES AND NAMES DICO
        Dim layersAndTypesDico As New SortedDictionary(Of String, SortedDictionary(Of String, List(Of MetaObject)))
        Dim layersDico = myAudit.GetLayers(myAudit._3dObjects)
        Dim typesCount = 0
        For Each layerSt In layersDico.Keys
            Dim layerObjList = layersDico(layerSt)
            Dim typesDico = myAudit.GetTypes(layerObjList)
            layersAndTypesDico.Add(layerSt, typesDico)
            typesCount += typesDico.Keys.Count
        Next

#End Region

#Region "FILL EXCEL"
        Dim calqueWsheet As Worksheet = myAudit.InsertProdActionTemplateInAuditWbook("AuditIFC_Template", "CALQUES")
        Dim typeStartLine = 5
        calqueWsheet.Cells(typeStartLine - 1, 1).Value = layersAndTypesDico.Keys.Count
        calqueWsheet.Cells(typeStartLine - 1, 3).Value = layersAndTypesDico.SelectMany(Function(o) o.Value).SelectMany(Function(o) o.Value).Count

        'FILL ARRAY
        Dim indice_next_line = 0
        Dim indice_layer_start_line As Integer
        For Each layerSt In layersAndTypesDico.Keys
            Dim typesDico = layersAndTypesDico(layerSt)
            Dim totalObjs = typesDico.Select(Function(o) o.Value.Count).Sum

            'Fill TYPES
            calqueWsheet.Range("B" + (typeStartLine + indice_next_line + 1).ToString).Value = layerSt
            calqueWsheet.Range("D" + (typeStartLine + indice_next_line + 1).ToString).Value = totalObjs

            '---Mise en forme
            Dim typeRange = calqueWsheet.Range("B" + (typeStartLine + indice_next_line + 1).ToString + ":D" + (typeStartLine + indice_next_line + 1).ToString)
            typeRange.FillColor = XlColor.FromTheme(XlThemeColor.Dark1, 0.7)
            typeRange.Font.Bold = True
            typeRange.Borders.SetOutsideBorders(System.Drawing.Color.Black, BorderLineStyle.Thin)
            typeRange.Borders.InsideVerticalBorders.LineStyle = BorderLineStyle.Thin
            typeRange.Borders.InsideHorizontalBorders.LineStyle = BorderLineStyle.Thin

            indice_next_line += 1
            indice_layer_start_line = indice_next_line
            For Each typeSt In typesDico.Keys
                'Fill NAMES
                calqueWsheet.Range("C" + (typeStartLine + indice_next_line + 1).ToString).Value = typeSt
                calqueWsheet.Range("D" + (typeStartLine + indice_next_line + 1).ToString).Value = typesDico(typeSt).Count
                indice_next_line += 1
                '---Mise en forme

            Next

            '---Mise en forme
            Dim namesRange = calqueWsheet.Range("B" + (typeStartLine + indice_layer_start_line + 1).ToString + ":D" + (typeStartLine + indice_next_line).ToString)
            namesRange.Borders.SetOutsideBorders(System.Drawing.Color.Black, BorderLineStyle.Thin)
            namesRange.Borders.InsideVerticalBorders.LineStyle = BorderLineStyle.Thin
            namesRange.GroupRows(False)
        Next

#End Region

#Region "CREATE TREE"
        Dim myTree = myAudit.AuditTreesWs.SmartAddTree("Audit IFC - CALQUES")
        Dim fileNode = myTree.SmartAddNode("File", myAudit.FileTgm.Name)
        myTree.RemoveNode(fileNode)
        fileNode = myTree.SmartAddNode("File", myAudit.FileTgm.Name)

        For Each layerSt In layersAndTypesDico.Keys
            If layerSt = "<NONE>" Then Continue For
            Dim layerNode = fileNode.SmartAddNode("Layer", layerSt)
            Dim layerTypes = layersAndTypesDico(layerSt)
            For Each typeSt In layerTypes.Keys
                Dim typeNode = layerNode.SmartAddNode("object.type", typeSt)
            Next
        Next
#End Region
    End Sub
End Class
