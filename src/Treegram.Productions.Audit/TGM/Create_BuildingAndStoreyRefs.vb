Imports M4D.Production.Core
Imports M4D.Treegram.Core.Entities
Imports M4D.Treegram.Core.Kernel
Imports M4D.Treegram.Core.Extensions.Entities
Imports M4D.Treegram.Core.Extensions.Kernel
Imports M4D.Deixi.Core
Imports Treegram.Bam.Libraries.AuditFile
Imports DevExpress.Spreadsheet
Imports Treegram.GeomKernel.BasicModeler
Imports SharpDX
Imports DevExpress.Export.Xl
Imports System.IO
Imports System.Windows.Forms
Imports Treegram.Bam.Libraries.AEC
Imports Treegram.Bam.Functions
Imports Treegram.Bam.Functions.AEC
Imports Treegram.ConstructionManagement
Imports Treegram.GeomFunctions
Imports Treegram.GeomLibrary

Public Class Create_ReferencesProdScript
    Inherits ProdScript
    Public Sub New()
        MyBase.New("TGM :: Create : Buildings and Storeys")
        AddAction(New Create_FileTgmBuildings())
        AddAction(New Create_ProjectReferences())
    End Sub
End Class

Public Class Create_ProjectReferences
    Inherits ProdAction
    Public Sub New()
        Name = "TGM :: Create : Storey References"
        PartOfScript = True
    End Sub

    Public Overrides Function GuiUpdateLauncherTree(inputs As Dictionary(Of String, IEnumerable(Of PersistEntity))) As Tree

        Dim launchTree = TempWorkspace.SmartAddTree("LauncherTree")
        While (launchTree.Nodes.FirstOrDefault() IsNot Nothing)
            launchTree.RemoveNode(launchTree.Nodes.First)
        End While

        If inputs.ContainsKey("TgmBuildings") AndAlso inputs("TgmBuildings").ToList.Count > 0 Then
            For Each objTgm As MetaObject In inputs("TgmBuildings")
                launchTree.SmartAddNode("TgmBuilding", objTgm.Name)
            Next
        End If

        Return launchTree
    End Function

    <ActionMethod()>
    Public Function MyMethod(TgmBuildings As MultipleElements) As ActionResult

        'Récupération des inputs
        Dim tgmBuilds = GetInputAsMetaObjects(TgmBuildings.Source).ToHashSet 'To get rid of duplicates
        If tgmBuilds.Count = 0 Then Throw New Exception("Inputs missing")
        Dim tgmBuilding = tgmBuilds(0)

        'Define file tgm
        Dim projWs = tgmBuilding.GetProject
        If LaunchedParentNode Is Nothing Then Throw New Exception("You can't lauch this ProdAction from this node")
        Dim fileName = LaunchedParentNode.Value?.ToString
        Dim fileWs = projWs.GetWorkspaces(fileName).FirstOrDefault
        If fileWs Is Nothing Then Throw New Exception("There is no workspace with this name : " + fileName)
        Dim fileTgm = fileWs.GetMetaObjectByName(fileName).FirstOrDefault
        If fileTgm Is Nothing Then Throw New Exception("There is no metaobject with this name : " + fileName)
        Dim myAudit As New FileObject(fileTgm)

        'Analysis
        Dim tgmStoreys = CreateProjectReferences_Vsimple(myAudit, tgmBuilding, False)

        'Tree visualization
        OutputTree = ProjectReferencingFunctions.ProjectReferencesOutputTree(projWs)

        'send result
        Dim outputs = New List(Of Element) From {New MultipleElements(tgmStoreys, "TgmStoreys", ElementState.[New])}
        Dim res = New SucceededActionResult(DateTime.Now.ToString("dd/MM/yy"), System.Drawing.Color.Green, outputs)
        Return res
    End Function

    ''' <summary>
    ''' Creates Project Reference objects for given file, Prompts user to validate StoreyReferences, and draws 3d representation of storey elevation if option drawStoreys = True
    ''' </summary>
    Private Function CreateProjectReferences_Vsimple(myAudit As FileObject, buildingRefTgm As MetaObject, Optional storeyReliabilityDetail As Boolean = False) As List(Of MetaObject)
        Dim fileTgm = myAudit.FileTgm
        Dim scanWs = myAudit.ScanWs
        Dim projWs = myAudit.ProjWs
        Dim extType = myAudit.ExtensionType
        Dim elevationAttr = "GlobalElevation"

        'Verif de l'existence d'un positionnement du fichier
        Dim originPt = fileTgm.Container.GetAttribute("OriginPoint")
        If originPt Is Nothing Then Throw New Exception("No georeferencing found for this file")

        'GET PROJ DATA WS AND FILE
        Dim projectRefWs As Workspace = buildingRefTgm.Container
        Dim boxVisuWs As Workspace

        boxVisuWs = myAudit.ProjWs.GetWorkspaces("ProjectReferences - BoxLevels").FirstOrDefault 'Old name
        If boxVisuWs IsNot Nothing Then
            boxVisuWs.Name = ProjectReference.Constants.boxLevelsName
        Else
            boxVisuWs = myAudit.ProjWs.GetWorkspaces(ProjectReference.Constants.boxLevelsName).FirstOrDefault
            If boxVisuWs Is Nothing Then
                boxVisuWs = myAudit.ProjWs.SmartAddWorkspace(ProjectReference.Constants.boxLevelsName, M4D.Treegram.Core.Constants.WorkspaceName.BoxWorkspace, True)
            End If
        End If

        If boxVisuWs.GetAttribute("OriginPoint") Is Nothing Then
            Dim origSourceAtt = myAudit.ScanWs.GetAttribute("OriginPoint")
            'Dim origDestAtt = boxVisuWs.SmartAddAttribute("OriginPoint", Nothing)
            origSourceAtt.CopyAttributeTo(boxVisuWs)
            boxVisuWs.SmartAddAttribute("GeoRefCopiedFrom", myAudit.ScanWs.Id.ToString)
            projWs.Push({boxVisuWs})
        End If

        'FILTER AND LOAD - NEW WAY TO LOAD ADDITIONAL GEOM
        Dim listToLoad As New List(Of Workspace) From {scanWs, projectRefWs}
        listToLoad.AddRange(scanWs.GetExtensionWorspaces(M4D.Treegram.Core.Extensions.Enums.ExtensionType.AdditionalGeometries))
        TempWorkspace.SmartAddTree("ALL").Filter(False, listToLoad).RunSynchronously()
        Treegram.GeomFunctions.Models.Reset3dDictionnary()

        'GET OLD REFERENCES FOR THIS BUILDING
        Dim _3dObjName = buildingRefTgm.Name
        Dim inactiveStoRefList As New List(Of String)
        Dim activeStoRefList As New List(Of String)
        Dim oldRefStoreysList As New List(Of OldRefStorey)
        If projectRefWs.MetaObjects.Count > 0 Then
            Dim stoOldReferences As List(Of MetaObject) = projectRefWs.MetaObjects.Where(Function(mo) mo.GetTgmType = ProjectReference.Constants.storeyRefName AndAlso mo.GetAttribute(ProjectReference.Constants.buildingRefName, True)?.Value = buildingRefTgm.Name).ToList
            For Each obj In stoOldReferences
                Dim myOldRefSto As New OldRefStorey(obj.Name) With {.IsActive = obj.IsActive}
                If obj.GetAttribute("InfraOrSuper")?.Value IsNot Nothing Then 'OLD NAME, à sup un jour...
                    Try
                        myOldRefSto.Type = System.[Enum].Parse(GetType(StoreyTypo), obj.GetAttribute("InfraOrSuper").Value) 'For INFRA and TOITURE types
                    Catch ex As Exception
                        myOldRefSto.Type = StoreyTypo.LOGEMENT
                    End Try
                ElseIf obj.GetAttribute("Type")?.Value IsNot Nothing Then
                    myOldRefSto.Type = System.[Enum].Parse(GetType(StoreyTypo), obj.GetAttribute("Type").Value)
                End If
                oldRefStoreysList.Add(myOldRefSto)

                'If obj.IsActive Then
                '    activeStoRefList.Add(obj.Name) 'Add to active list to keep in memory last settings
                'Else
                '    inactiveStoRefList.Add(obj.Name) 'Add to inactive list to keep in memory last settings
                'End If
            Next
        End If

#Region "COPY GRIDS & GRIDS GEOMETRY"
        If extType = "rvt" Then
            Dim Quads = New List(Of (origin As MetaObject, ref As MetaObject))

            For Each fileRel As Relation In fileTgm.Relations
                Dim oChild As MetaObject = fileRel.Target
                If oChild.Name = "REF" Then
                    For Each subRel As Relation In oChild.Relations
                        Dim subChild = subRel.Target
                        If subChild.IsActive AndAlso subChild.GetAttribute("object.type") IsNot Nothing AndAlso subChild.GetAttribute("object.type").Value = "FileGrid" Then

                            Dim newChild = subChild.CopyMetaObject(projectRefWs, True)
                            Dim newList = New List(Of MetaObject) From {newChild}

                            'projectRefWs.Push(newList)
                            newChild.SmartAssignTemplate("FileGrid")
                            buildingRefTgm.SmartAddRelation(newChild, RelationType.Containance)
                            projectRefWs.Push(newList)
                            Quads.Add((subChild, newChild))
                        End If
                    Next
                End If
            Next

            Dim gridWs As Workspace = fileTgm.Container()
            Treegram.GeomFunctions.Models.LoadGeometryModels(gridWs)

            Dim allModels As New List(Of Model)
            For Each quad In Quads
                Dim mObj = quad.origin
                Dim refobj = quad.ref
                Dim oModels = Treegram.GeomFunctions.Models.GetGeometryModels(mObj)

                For Each gridModel In oModels
                    gridModel.Tag = New Tuple(Of Integer, Integer)(refobj.Id.Container.Id - 1, refobj.Id.Entity.Id)
                Next

                allModels.AddRange(oModels)

            Next

            Dim refPath As String = Treegram.GeomFunctions.PathUtils.Get3dWorkspaceDirectory(projectRefWs)
            Dim gridPath As String = refPath + "\" + fileTgm.Name + "_Ref_Grids.3drtgm"

            Dim gridFile As New System.IO.FileInfo(gridPath)
            If gridFile.Exists Then
                gridFile.Delete()
            End If

            If allModels.Count > 0 Then
                Writer.WriteModels(gridPath, allModels, False, 1, 1)
            End If
        End If



#End Region

#Region "COMPLETE STOREY REFERENCES DICO"

        Dim levelType As String
        If extType = "ifc" Then
            levelType = "IfcBuildingStorey"
        ElseIf extType = "rvt" Then
            levelType = "FileLevel"
        Else
            Throw New Exception("File extension Not supported")
        End If
        '---Get storeys
        Dim storeysList = scanWs.GetMetaObjects(, levelType).Where(Function(o) o.GetAttribute("TgmBuilding", True)?.Value = buildingRefTgm.Name).ToList
        If storeysList.Count = 0 Then
            storeysList = scanWs.GetMetaObjects(, levelType).ToList 'Only 1 Building case
            If storeysList.Count = 0 Then
                Throw New Exception("Can't find " + levelType)
            End If
        End If
        '---Get storey objects and storey activity
        CommonFunctions.SortLevelsByElevation(storeysList, elevationAttr)
        Dim defaultStoreyType = StoreyTypo.INFRA
        Dim storeysDico As New Dictionary(Of MetaObject, (elevation As Double, objsList As List(Of MetaObject), reliability As String, type As StoreyTypo))
        For Each storeyTgm In storeysList
            'Get storey type old value
            Dim oldRefSto = oldRefStoreysList.FirstOrDefault(Function(o) o.Name = storeyTgm.Name)

            'Get storey type default value
            Dim toLowStoName = storeyTgm.Name.ToLower.Trim.Replace(" ", String.Empty)
            If extType = "rvt" AndAlso storeyTgm.GetAttribute("Etage de bâtiment")?.Value IsNot Nothing AndAlso 'Specific to revit
                                       storeyTgm.GetAttribute("Etage de bâtiment")?.Value = "Non" Then
                defaultStoreyType = StoreyTypo.AUCUN
            ElseIf storeyTgm.Equals(storeysList.Last) Then
                defaultStoreyType = StoreyTypo.TOITURE
            ElseIf defaultStoreyType = StoreyTypo.INFRA AndAlso (toLowStoName.Contains("rdc") OrElse toLowStoName.Contains("n00") OrElse toLowStoName.Contains("l00") OrElse
                toLowStoName.Contains("r+0") OrElse toLowStoName.Contains("r0") OrElse toLowStoName.Contains("niv0") OrElse toLowStoName.Contains("niveau0") OrElse toLowStoName.Contains("rez")) Then
                defaultStoreyType = StoreyTypo.MIXTE
            ElseIf defaultStoreyType = StoreyTypo.MIXTE Then
                defaultStoreyType = StoreyTypo.LOGEMENT
            End If

            'Pick storey type
            Dim storeyType As StoreyTypo
            If oldRefSto IsNot Nothing Then 'Old settings first
                storeyType = oldRefSto.Type
            Else 'Default settings then
                storeyType = defaultStoreyType
            End If

            ''Get IsUsed param default value
            'Dim isUsed As Boolean
            'If oldRefSto IsNot Nothing AndAlso Not oldRefSto.IsActive Then 'Already inactive in last revision
            '    isUsed = False
            'ElseIf oldRefSto IsNot Nothing AndAlso oldRefSto.IsActive Then 'Already active in last revision
            '    isUsed = True
            'ElseIf extType = "rvt" AndAlso storeyTgm.GetAttribute("Etage de bâtiment")?.Value IsNot Nothing AndAlso 'Specific to revit
            '   storeyTgm.GetAttribute("Etage de bâtiment")?.Value = "Non" Then
            '    isUsed = False
            'Else 'Otherwise -> active by default
            '    isUsed = True
            'End If

            'Get storey reliability with doors
            Dim globalElevationDbl As Double = CDbl(storeyTgm.GetAttribute(elevationAttr, True).Value)
            Dim elevationDbl = CommonFunctions.ConvertGlobalElevationToScanWs(globalElevationDbl, scanWs)
            Dim stoDoors = storeyTgm.GetChildren().Where(Function(d) d.GetTgmType = "IfcDoor" Or d.GetAttribute("RevitCategory", True)?.Value = "Portes").ToList
            Dim minzDico = LowerElevationDico(stoDoors)
            Dim elevation_tolerance = 0.005
            Dim numWithinTolerance As Integer = 0
            Dim reliability As Double = 0
            If minzDico IsNot Nothing AndAlso minzDico.Count > 0 Then
                For Each kvp In minzDico
                    If Math.Abs(elevationDbl - kvp.Key) < elevation_tolerance Then
                        numWithinTolerance += kvp.Value
                    End If
                Next
                Dim majorityElv As Double = minzDico.Keys(0)
                reliability = Math.Round(numWithinTolerance / stoDoors.Count, 1) * 100
            End If

            storeysDico.Add(storeyTgm, (elevationDbl, storeyTgm.GetChildren().ToList, reliability.ToString + "%", storeyType))
        Next
        'Dim buildObjs = storeysDico.SelectMany(Function(o) o.Value.objsList).ToList


#End Region

        '---Correct manualy storeys dico
        storeysDico = ProjectReferenceExcelExport_DevExpress(fileTgm, buildingRefTgm, storeysDico)
        Dim usedStoreys = storeysDico.Where(Function(o) o.Value.type <> StoreyTypo.AUCUN).Select(Function(l) l.Key).ToList


#Region "SAVE NEW DATA ORGANISATION"

        '---Delete old references for this building
        If projectRefWs.MetaObjects.Count > 0 Then
            CommonFunctions.Delete3drtgmFile(projectRefWs, _3dObjName)
            Dim objToDel As List(Of MetaObject) = projectRefWs.MetaObjects.Where(Function(mo) mo.IsActive AndAlso mo.GetTgmType = ProjectReference.Constants.storeyRefName AndAlso mo.GetAttribute(ProjectReference.Constants.buildingRefName, True)?.Value = buildingRefTgm.Name).ToList
            Dim delInt = 0
            For Each obj In objToDel
                obj.ClearMetaObject(delInt, False, False, False) 'Keep relation !!!
            Next
            CommonAecFunctions.DeactivateTgmObjs(projWs, boxVisuWs, ProjectReference.Constants.boxTemplateName, _3dObjName, buildingRefTgm.Name)
        End If

        '---Save source file
        buildingRefTgm.SmartAddAttribute("ScanWs", scanWs.Name)

        '---Save georef
        Dim translateGeo As Double() = Nothing
        Dim rotateGeo As Double? = Nothing
        CommonFunctions.GetGeodata(scanWs, rotateGeo, translateGeo)
        Dim newOriginRef = buildingRefTgm.SmartAddAttribute(originPt.Name, originPt.Value)
        newOriginRef.SmartAddAttribute("X", translateGeo(0))
        newOriginRef.SmartAddAttribute("Y", translateGeo(1))
        newOriginRef.SmartAddAttribute("Z", translateGeo(2))
        newOriginRef.SmartAddAttribute("Angle", CDbl(rotateGeo))

        '---Save revit position point
        If extType = "rvt" Then
            Dim baseProject = scanWs.GetMetaObjects("BasePoint", Nothing, False).FirstOrDefault
            Dim survey = scanWs.GetMetaObjects("SurveyPoint", Nothing, False).FirstOrDefault
            If baseProject IsNot Nothing Then
                Dim newBaseRef = buildingRefTgm.SmartAddAttribute("BasePoint", Nothing)
                For Each att In baseProject.GetAttribute("SharedPosition").Attributes
                    newBaseRef.SmartAddAttribute(att.Name, att.Value)
                Next
            End If
            If survey IsNot Nothing Then
                Dim newSurveyRef = buildingRefTgm.SmartAddAttribute("SurveyPoint", Nothing)
                For Each att In survey.GetAttribute("SharedPosition").Attributes
                    newSurveyRef.SmartAddAttribute(att.Name, att.Value)
                Next
            End If
        End If

        '---Loop on storeys
        Dim storeyGeomDico As New Dictionary(Of MetaObject, Tuple(Of List(Of Profile3D), SharpDX.Color4))
        Dim boxModels As New List(Of M4D.Deixi.Core.Model)
        Dim i As Integer = 0
        Dim isBelowInfra = True
        For i = 0 To storeysDico.Keys.Count - 1
            Dim storeyTgm = storeysDico.Keys(i)
            Dim stoObjs = storeysDico(storeyTgm).objsList
            Dim globalElevationDbl As Double = storeyTgm.GetAttribute(elevationAttr, True).Value
            Dim elevationDbl = CommonFunctions.ConvertGlobalElevationToScanWs(globalElevationDbl, projectRefWs) 'Because both BoxWs and RefWs have the same repositioning

            'Create storey reference
            Dim storeyRefTgm = buildingRefTgm.SmartAddMetaObject(storeyTgm.Name, RelationType.Containance, ProjectReference.Constants.storeyRefName)
            storeyRefTgm.SmartAddAttribute("GlobalElevation", Math.Round(globalElevationDbl, 3)) 'Rounded to mm
            storeyRefTgm.SmartAddAttribute("Type", storeysDico(storeyTgm).type.ToString)
            Dim isGroundFloor = False
            If storeysDico(storeyTgm).type = StoreyTypo.AUCUN Or storeysDico(storeyTgm).type = StoreyTypo.NIVEAU_HAUT_INFRA Then
                storeyRefTgm.IsActive = False
                Continue For
            ElseIf isBelowInfra AndAlso storeysDico(storeyTgm).type <> StoreyTypo.INFRA Then
                isGroundFloor = True
                isBelowInfra = False
            End If
            storeyRefTgm.SmartAddAttribute("GroundFloor", isGroundFloor)

            'Find out next storey in used storeys
            Dim nextStoreyElevation As Double
            Dim currentIndex = usedStoreys.IndexOf(storeyTgm)
            If currentIndex = usedStoreys.Count - 1 Then
                nextStoreyElevation = elevationDbl + 3.0
            Else
                Dim nextStoTgm = usedStoreys(currentIndex + 1)
                Dim nextStoGlobalElev = CDbl(nextStoTgm.GetAttribute(elevationAttr, True).Value)
                nextStoreyElevation = CommonFunctions.ConvertGlobalElevationToScanWs(nextStoGlobalElev, projectRefWs)
            End If

            'Define storey 2d boundary box
            Dim extrapol = 2.0
            Dim minMaxSurf As Surface = CommonAecGeomFunctions.GetGroup2dBbox(elevationDbl, stoObjs, extrapol)
            'If minMaxSurf = Nothing And stoObjs.Count > 0 Then
            '    Throw New Exception("Can't define storey 2dBbox")
            If minMaxSurf = Nothing Then
                Continue For 'No 3d representation for this reference storey
            End If
            Dim planarSurf = CommonGeomFunctions.SetElevationToSurface(New List(Of Surface) From {minMaxSurf}, elevationDbl)
            Dim myColor As New SharpDX.Color4(187 / 255, 187 / 255, 187 / 255, 0.8) 'light grey
            Dim planarTuple As New Tuple(Of List(Of Profile3D), SharpDX.Color4)(planarSurf.Select(Function(o) o.ToProfile3d).ToList, myColor)
            storeyGeomDico.Add(storeyRefTgm, planarTuple)

            'Define storey 3d boundary box
            Dim boxName = buildingRefTgm.Name + "-" + storeyTgm.Name + "-ViewBox"
            Dim storeyBoxTgm = boxVisuWs.SmartAddMetaObject(boxName, ProjectReference.Constants.boxTemplateName)
            storeyBoxTgm.SmartAddAttribute(ProjectReference.Constants.boxNameAttName, boxName)
            storeyBoxTgm.SmartAddAttribute(ProjectReference.Constants.buildingRefName, buildingRefTgm.Name)
            Dim boxBottomDelta = -0.5
            Dim boxTopDelta = 0.2
            Dim botSurf = CommonGeomFunctions.SetElevationToSurface(New List(Of Surface) From {minMaxSurf}, nextStoreyElevation + boxBottomDelta)
            Dim topSurf = CommonGeomFunctions.SetElevationToSurface(New List(Of Surface) From {minMaxSurf}, elevationDbl + boxTopDelta)
            Dim geometryBox = New Treegram.GeomFunctions.Geometries.ParallelepipedGeometry(botSurf(0).Points(0).ToVector3, botSurf(0).Points(1).ToVector3, botSurf(0).Points(3).ToVector3, botSurf(0).Points(2).ToVector3,
                                                         topSurf(0).Points(0).ToVector3, topSurf(0).Points(1).ToVector3, topSurf(0).Points(3).ToVector3, topSurf(0).Points(2).ToVector3,
                                                         New Color4(0.9F, 0.9F, 0.9F, 1.0F))
            geometryBox.ReorderIndices()

            projWs.Push({storeyBoxTgm}.ToList)
            Dim geometryModel = New M4D.Deixi.Core.Model() With {.Geometry = geometryBox, .IsInstance = False, .InstanceMatrix = SharpDX.Matrix.Identity, .ReplaceColor = False,
                                                .Tag = New Tuple(Of Integer, Integer, Integer)(storeyBoxTgm.Id.Container.Id - 1, storeyBoxTgm.Id.Entity.Id, 1)}
            geometryModel.IsInstance = True
            geometryModel.InstanceMatrix = Treegram.GeomFunctions.GeoReferencing.GetCorrectionMatrix(scanWs, boxVisuWs)
            boxModels.Add(geometryModel)
        Next

        'SET GEOMETRY TO TGM OBJ
        If storeyGeomDico.Keys.Count > 0 Then
            projWs.PushAllModifiedEntities()
            Treegram.GeomFunctions.Models.CreateGeometryInTreegramFromSurface(projectRefWs, storeyGeomDico, _3dObjName, scanWs)
            Dim boxWs3dPath = Treegram.GeomFunctions.PathUtils.Get3dWorkspaceDirectory(boxVisuWs) + "\" + _3dObjName + ".3drtgm"
            Writer.WriteModels(boxWs3dPath, boxModels, True, CType(1, Byte), 1.0F)
            'Writer.WriteModels(boxWs3dPath, boxModels, False, CType(1, Byte), 1.0F) '3drtgm corrompu
        End If
#End Region

        Return usedStoreys
    End Function



    Public Shared Function ProjectReferenceExcelExport_DevExpress(fileTgm As MetaObject, buildingRefTgm As MetaObject, storeysDico As Dictionary(Of MetaObject, (elevation As Double, objsList As List(Of MetaObject), reliability As String, type As StoreyTypo)), Optional storeyReliabilityDetail As Boolean = False) As Dictionary(Of MetaObject, (elevation As Double, objsList As List(Of MetaObject), reliability As String, type As StoreyTypo))
        Dim projWs As Workspace = fileTgm.GetProject

        'Deal with paths
        Dim projectFile = New FileInfo(projWs.Path.LocalPath)
        Dim exportDirPath = Path.Combine(projectFile.Directory.FullName, "Project References")
        If Not Directory.Exists(exportDirPath) Then
            Directory.CreateDirectory(exportDirPath)
        End If
        Dim destPath = System.IO.Path.Combine(exportDirPath, buildingRefTgm.Name + "_ProjectReferences_" + Date.Now.ToString("yyMMdd") + ".xlsx")
        If File.Exists(destPath) Then
            Try
                File.Delete(destPath)
            Catch ex As Exception
                Throw New Exception("Fermer l' XL des références de projet concerné et relancer")
            End Try
        End If

        'Get XL
        Dim myWbook As New DevExpress.Spreadsheet.Workbook
        If Not myWbook.CreateNewDocument Then
            Throw New Exception("Could not create spreadsheet document")
        End If
        Dim myWsheet As DevExpress.Spreadsheet.Worksheet = myWbook.Worksheets.ActiveWorksheet
        Dim start_line = 5
        Dim start_column = 1

        'Complete Worksheet
        Dim typesTitle = System.Enum.GetNames(GetType(StoreyTypo)).First
        For s As Integer = 1 To System.Enum.GetNames(GetType(StoreyTypo)).Count - 1
            typesTitle += " / " + System.Enum.GetNames(GetType(StoreyTypo))(s).Replace("_", " ")
        Next
        myWbook.BeginUpdate()
        myWsheet.Cells(1, start_column).Value = "Project References"
        myWsheet.Cells(2, start_column).Value = "File : " + fileTgm.Name
        Dim titles As New List(Of String) From {"STOREY", "BUILDING", "STOREY ELEVATION", "NUMBER OF OBJECTS", "STOREY RELIABILITY", typesTitle}
        Dim current_line = start_line
        Dim current_column = start_column
        For Each elt In titles
            myWsheet.Cells(current_line, current_column).Value = elt
            current_column += 1
        Next
        current_line += 1
        Dim sortedStoreys = storeysDico.Keys.ToList
        sortedStoreys.Reverse()
        For Each storeyTgm In sortedStoreys
            Dim infos = storeysDico(storeyTgm)
            myWsheet.Cells(current_line, start_column + 0).Value = storeyTgm.Name
            myWsheet.Cells(current_line, start_column + 1).Value = buildingRefTgm.Name
            myWsheet.Cells(current_line, start_column + 2).Value = Math.Round(infos.elevation, 3)
            myWsheet.Cells(current_line, start_column + 3).Value = infos.objsList.Count
            myWsheet.Cells(current_line, start_column + 4).Value = infos.reliability
            myWsheet.Cells(current_line, start_column + 5).Value = infos.type.ToString
            'myWsheet.Cells(current_line, start_column + 6).Value = infos.isUsed
            current_line += 1
        Next

        '---Mise en forme
        Dim titlesRange = myWsheet.Cells(start_line, start_column).Resize(1, 6)
        titlesRange.Font.Bold = True
        titlesRange.ColumnWidth = 600
        Dim valuesRange = myWsheet.Cells(start_line + 1, start_column).Resize(storeysDico.Keys.Count, 6)
        valuesRange.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center
        Dim rangeToModify = myWsheet.Cells(start_line + 1, start_column + 5).Resize(storeysDico.Keys.Count, 1)
        rangeToModify.FillColor = XlColor.FromArgb(248, 203, 173)
        rangeToModify.Borders.SetOutsideBorders(System.Drawing.Color.Black, BorderLineStyle.Thick)
        rangeToModify.Borders.InsideVerticalBorders.LineStyle = BorderLineStyle.Thin

        'Export as Xl file and open it
        myWbook.Calculate()
        myWbook.SaveDocument(destPath, DocumentFormat.OpenXml)
        System.Diagnostics.Process.Start(destPath)

        'User set modifications
        MessageBox.Show("If necessary, modify entries in ORANGE cells and SAVE EXCEL to take it into account (before clicking OK) !")

        'Get modifications
        Dim modifiedWbook As New DevExpress.Spreadsheet.Workbook
        If Not modifiedWbook.LoadDocument(destPath) Then
            Throw New Exception("Could not find : " + destPath)
        End If
        Dim modifiedWsheet As DevExpress.Spreadsheet.Worksheet = modifiedWbook.Worksheets.ActiveWorksheet
        current_line = start_line + 1
        Dim correctedStoreysDico As New Dictionary(Of MetaObject, (elevation As Double, objsList As List(Of MetaObject), reliability As String, type As StoreyTypo))
        For Each storeyTgm In sortedStoreys
            Dim infos = storeysDico(storeyTgm)

            Dim stoTypeSt = modifiedWsheet.Cells(current_line, start_column + 5).Value.ToString.ToUpper.Replace(" ", "_")
            Dim stoType As StoreyTypo
            If System.Enum.GetNames(GetType(StoreyTypo)).Contains(stoTypeSt) Then
                stoType = System.Enum.Parse(GetType(StoreyTypo), stoTypeSt)
            Else
                stoType = StoreyTypo.AUCUN
            End If
            'Try
            '    stoType = System.Enum.Parse(GetType(StoreyTypo), stoTypeSt)
            'Catch ex As Exception
            '    stoType = StoreyTypo.AUCUN
            'End Try
            'Dim isUsedValue = modifiedWsheet.Cells(current_line, start_column + 6).Value.BooleanValue

            correctedStoreysDico.Add(storeyTgm, (infos.elevation, infos.objsList, infos.reliability, stoType))

            current_line += 1
        Next

        Return correctedStoreysDico

    End Function



    ''' <summary>
    ''' From a list of MetaObjects, returns a dictionnary of lower elevations (keys : elevations, values : number of objects with lower elevation equal to the key)
    ''' </summary>
    ''' <param name="objs"></param>
    ''' <returns></returns>
    Private Function LowerElevationDico(objs As List(Of MetaObject)) As Dictionary(Of Double, Integer)

        Treegram.GeomFunctions.Models.LoadGeometryModels(objs, True)

        Dim n_stoDoors As Integer = 0
        Dim minzList As New List(Of Double)
        Dim minzDico As New Dictionary(Of Double, Integer)

        For Each doorTgm In objs
            Dim anaDoor As New AecDoor(doorTgm)

            If anaDoor.Models Is Nothing OrElse anaDoor.Models.Count = 0 Then
                Continue For
            End If

            Dim doorMinZ As Double = Math.Round(anaDoor.MinZ, 3)

            minzList.Add(doorMinZ)
            n_stoDoors += 1

            If minzDico.ContainsKey(doorMinZ) Then
                minzDico(doorMinZ) += 1
            Else
                minzDico.Add(doorMinZ, 1)
            End If
        Next


        'Sort elevations by number of objects
        'Dim elevations = minzDico.Keys.ToList
        'Dim numObj = minzDico.Values.ToList
        'Dim array() As IList = {numObj, elevations}
        'Core.SortLists(Of Double)(array, Function(i1 As Double, i2 As Double)
        '                                     Return i1.CompareTo(i2)
        '                                 End Function)

        'Dim sortedElevationDico As New Dictionary(Of Double, Integer)
        'For Each elv In elevations
        '    sortedElevationDico.Add(elv, minzDico(elv))
        'Next

        Dim sortedElevationDico = minzDico.OrderByDescending(Function(item) item.Value)

        Return minzDico

    End Function

    Private Class OldRefStorey
        Public Sub New(Name As String)
            Me.Name = Name
        End Sub
        Property Name As String
        Property IsActive As Boolean
        Property Type As StoreyTypo
    End Class
End Class