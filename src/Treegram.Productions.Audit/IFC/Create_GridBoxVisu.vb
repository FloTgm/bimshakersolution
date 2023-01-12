Imports M4D.Production.Core
Imports M4D.Treegram.Core.Entities
Imports M4D.Treegram.Core.Kernel
Imports M4D.Treegram.Core.Extensions.Entities
Imports M4D.Deixi.Core
Imports Treegram.Bam.Libraries.AuditFile
Imports Treegram.Bam.Functions.AEC
Imports Treegram.Bam.Functions
Imports SharpDX
Imports M4D.Treegram.Core.Extensions.Kernel
Imports Treegram.GeomKernel.BasicModeler
Imports Point = Treegram.GeomKernel.BasicModeler.Point
Imports Line = Treegram.GeomKernel.BasicModeler.Line
Imports M4D.Treegram.Core.Constants
Imports Treegram.ConstructionManagement
Imports System.IO

'Public Class CreateGridBoxVisuScript
'    Inherits ProdScript
'    Public Sub New()
'        MyBase.New("IFC :: Create : View Boxes from Grids")
'        AddAction(New CreateGridBoxVisu() With {.Name = Name, .PartOfScript = True})
'    End Sub
'End Class
Public Class CreateGridBoxVisu
    Inherits ProdAction
    Public Sub New()
    End Sub

    Public Overrides Sub CreateGuiInputTree()
        'Création d'une zone de filtrage (arbre) temporaire
        InputTree = TempWorkspace.AddTree("--- Temporary Input Tree ---")

        'Création d'un filtre : murs
        Dim wallNode = InputTree.SmartAddNode("object.type", "IfcGrid")

        'Drag and Drop automatique des objets filtrés vers l'algo (Optionnel)
        SelectYourSetsOfInput.Add("Grids", {wallNode}.ToList())
    End Sub


    Public Overrides Function GuiUpdateLauncherTree(inputs As Dictionary(Of String, IEnumerable(Of PersistEntity))) As Tree
        'Séparation des lancements par fichier
        Dim launchTree As Tree = TempWorkspace.SmartAddTree("LauncherTree")
        CommonAecLauncherTrees.FileLevelsLaunchTree(launchTree, inputs)
        Return launchTree
    End Function


    <ActionMethod()>
    Public Function MyMethod(Grids As MultipleElements) As ActionResult

        'PREPARATION ENVIRONNEMENT TGM
        '--Sortie de l'algo s'il n'y a pas d'inputs ou trop
        If GetInputAsMetaObjects(Grids.Source).Count = 0 Then
            Return New IgnoredActionResult("No input")
        End If
        If GetInputAsMetaObjects(Grids.Source).Count > 1 Then
            Return New FailedActionResult("Too many inputs")
        End If

        '--Récupération des métaobjets murs sous formé de liste
        Dim gridTgm = GetInputAsMetaObjects(Grids.Source).First

        '--Récupération du fichier source
        Dim myIfcFile = CType(CommonAecFunctions.GetAuditFileFromInputs({gridTgm}.ToList), IfcFileObject)

        '--Chargement des données BAMs
        Dim wsList = New List(Of Workspace) From {myIfcFile.ScanWs}
        wsList.AddRange(myIfcFile.ScanWs.GetExtensionWorspaces(M4D.Treegram.Core.Extensions.Enums.ExtensionType.AdditionalGeometries).ToList)
        CommonFunctions.RunFilter(TempWorkspace, wsList)


        'ANALYSE MÉTIER
        'Get storeys and elevations
        Dim allStoreys = myIfcFile.SiteBuildings.SelectMany(Function(o) myIfcFile.BuildingStoreys(o)).ToList
        Dim launchedStoreys = allStoreys.Where(Function(o) o.Name = LaunchedNode.Value).ToList
        If launchedStoreys.Count = 0 Then
            Return New FailedActionResult("Cannot find launched FileLevel")
        ElseIf launchedStoreys.Count > 1 Then
            Return New FailedActionResult("More than one FileLevel with this name")
        End If
        Dim launchedStorey = launchedStoreys.First
        Dim elevationInf = CDbl(launchedStorey.GetAttribute("Elevation").Value)
        Dim concernedBuilding = launchedStorey.GetParents("IfcBuilding").First
        Dim concernedStoreys = myIfcFile.BuildingStoreys(concernedBuilding)
        CommonFunctions.SortLevelsByElevation(concernedStoreys)
        Dim launchedStoIndex = concernedStoreys.IndexOf(launchedStorey)
        Dim elevationSup As Double
        If launchedStoIndex = concernedStoreys.Count - 1 Then
            elevationSup = elevationInf + 3.0
        Else
            Dim storeySup = concernedStoreys(launchedStoIndex + 1)
            elevationSup = CDbl(storeySup.GetAttribute("Elevation").Value)
        End If

        'Get grid lines
        Dim uAxesAtt = gridTgm.GetAttribute("UAxes")
        Dim vAxesAtt = gridTgm.GetAttribute("VAxes")
        If uAxesAtt Is Nothing Or vAxesAtt Is Nothing Then
            Return New FailedActionResult("Missing attributes")
        End If
        Dim uAxes As New List(Of Line)
        For Each uAxe In uAxesAtt.Attributes
            Dim startPt = New Point(Convert.ToDouble(uAxe.SmartGetAttribute("Point 0/_\X").Value) / 1000, Convert.ToDouble(uAxe.SmartGetAttribute("Point 0/_\Y").Value) / 1000)
            Dim endPt = New Point(Convert.ToDouble(uAxe.SmartGetAttribute("Point 1/_\X").Value / 1000), Convert.ToDouble(uAxe.SmartGetAttribute("Point 1/_\Y").Value) / 1000)
            uAxes.Add(New Line(startPt, endPt))
            If uAxes.Count = 8 Then
                Exit For 'SPECIFIQUE A SETEC
            End If
        Next
        Dim vAxes As New List(Of Line)
        Dim v As Integer = 0
        For Each vAxe In vAxesAtt.Attributes
            v += 1
            If v < 9 Then
                Continue For 'SPECIFIQUE A SETEC
            End If
            Dim startPt = New Point(Convert.ToDouble(vAxe.SmartGetAttribute("Point 0/_\X").Value / 1000), Convert.ToDouble(vAxe.SmartGetAttribute("Point 0/_\Y").Value) / 1000)
            Dim endPt = New Point(Convert.ToDouble(vAxe.SmartGetAttribute("Point 1/_\X").Value / 1000), Convert.ToDouble(vAxe.SmartGetAttribute("Point 1/_\Y").Value) / 1000)
            vAxes.Add(New Line(startPt, endPt))
        Next
        If uAxes.Count < 2 Or vAxes.Count < 2 Then
            Return New FailedActionResult("Not enough axes")
        End If

        'Get grid box ws
        Dim boxVisuWs = myIfcFile.ProjWs.GetWorkspaces(ProjectReference.Constants.boxLevelsName).FirstOrDefault
        If boxVisuWs Is Nothing Then
            boxVisuWs = myIfcFile.ProjWs.SmartAddWorkspace(ProjectReference.Constants.boxLevelsName, WorkspaceName.BoxWorkspace, True)
            myIfcFile.ProjWs.Push({boxVisuWs})
        End If

        'Get first georef ws to copy georeferencing
        Dim firstGeoReferencedWorkspace = CommonFunctions.ProjectFirstGeoReferencedWs(myIfcFile.ProjWs)
        If firstGeoReferencedWorkspace IsNot Nothing AndAlso boxVisuWs.GetAttribute(AttributeName.OriginPoint) Is Nothing Then
            Dim origSourceAtt = firstGeoReferencedWorkspace.GetAttribute(AttributeName.OriginPoint)
            origSourceAtt.CopyAttributeTo(boxVisuWs)
            boxVisuWs.SmartAddAttribute("GeoRefCopiedFrom", firstGeoReferencedWorkspace.Id.ToString)
            myIfcFile.ProjWs.Push({boxVisuWs})
        End If

        'Define xVec and yVec AND SORT AXES
        '--- U
        Dim uFirstAxe = uAxes.First
        Dim uSecAxe = uAxes(1)
        Dim uSecProj = GetNormalPoint(uFirstAxe.StartPoint, uSecAxe.StartPoint, uSecAxe.EndPoint)
        Dim xVec = New Vector(uFirstAxe.StartPoint, uSecProj).ScaleOneVector
        Dim farPt = uFirstAxe.StartPoint.Translate2(xVec, -1000.0)
        Dim uDistances As New List(Of Double)
        For Each uAxe In uAxes
            Dim uProj = GetNormalPoint(uFirstAxe.StartPoint, uAxe.StartPoint, uAxe.EndPoint)
            uDistances.Add(Distance(farPt, uProj))
        Next
        Dim uArray() As IList = {uDistances, uAxes}
        CommonFunctions.SortLists(Of Double)(uArray, Function(i1 As Double, i2 As Double)
                                                         Return i1.CompareTo(i2)
                                                     End Function)

        '--- V
        Dim vFirstAxe = vAxes.First
        Dim vSecAxe = vAxes(1)
        Dim vSecProj = GetNormalPoint(vFirstAxe.StartPoint, vSecAxe.StartPoint, vSecAxe.EndPoint)
        Dim yVec = New Vector(vFirstAxe.StartPoint, vSecProj).ScaleOneVector
        Dim vFarPt = vFirstAxe.StartPoint.Translate2(yVec, -1000.0)
        Dim vDistances As New List(Of Double)
        For Each vAxe In vAxes
            Dim vProj = GetNormalPoint(vFirstAxe.StartPoint, vAxe.StartPoint, vAxe.EndPoint)
            vDistances.Add(Distance(vFarPt, vProj))
        Next
        Dim vArray() As IList = {vDistances, vAxes}
        CommonFunctions.SortLists(Of Double)(vArray, Function(i1 As Double, i2 As Double)
                                                         Return i1.CompareTo(i2)
                                                     End Function)

        'Delete old grid boxes
        Dim objType = "GridViewBox"
        Dim delInt = 0
        For Each boxTgm In boxVisuWs.MetaObjects.ToList
            If boxTgm.Name.Split("_").Last.Split("-").First = objType AndAlso boxTgm.GetAttribute("FileLevel")?.Value = launchedStorey.Name Then
                boxTgm.ClearMetaObject(delInt, False, True, False)
            End If
        Next
        Dim box3dPath = $"{Treegram.GeomFunctions.PathUtils.Get3dWorkspaceDirectory(boxVisuWs)}\{launchedStorey.Name}_{objType}.3drtgm"
        If File.Exists(box3dPath) Then
            File.Delete(box3dPath)
        End If

        'Construct grid boxes
        Dim defaultOffset = 3.0
        Dim visuBoxes As New List(Of Model)
        For i As Integer = 0 To uAxes.Count - 1
            Dim uAxe = uAxes(i)

            Dim uPreviousOffset As Double = defaultOffset
            Dim uNextOffset As Double = defaultOffset
            If i = 0 Then 'First axe
                Dim uNextAxe = uAxes(i + 1)
                Dim uNextProj = GetNormalPoint(uAxe.StartPoint, uNextAxe.StartPoint, uNextAxe.EndPoint)
                uNextOffset = Math.Round(Distance(uAxe.StartPoint, uNextProj) / 2, 3)

            ElseIf i = uAxes.Count - 1 Then 'Last axe
                Dim uPreviousAxe = uAxes(i - 1)
                Dim uPreviousProj = GetNormalPoint(uAxe.StartPoint, uPreviousAxe.StartPoint, uPreviousAxe.EndPoint)
                uPreviousOffset = Math.Round(Distance(uPreviousProj, uAxe.StartPoint) / 2, 3)

            Else
                Dim uPreviousAxe = uAxes(i - 1)
                Dim uPreviousProj = GetNormalPoint(uAxe.StartPoint, uPreviousAxe.StartPoint, uPreviousAxe.EndPoint)
                uPreviousOffset = Math.Round(Distance(uPreviousProj, uAxe.StartPoint) / 2, 3)
                Dim uNextAxe = uAxes(i + 1)
                Dim uNextProj = GetNormalPoint(uAxe.StartPoint, uNextAxe.StartPoint, uNextAxe.EndPoint)
                uNextOffset = Math.Round(Distance(uAxe.StartPoint, uNextProj) / 2, 3)
            End If


            For j As Integer = 0 To vAxes.Count - 1
                Dim vAxe = vAxes(j)

                Dim vPreviousOffset As Double = defaultOffset
                Dim vNextOffset As Double = defaultOffset
                If j = 0 Then 'First axe
                    Dim vNextAxe = vAxes(j + 1)
                    Dim vNextProj = GetNormalPoint(vAxe.StartPoint, vNextAxe.StartPoint, vNextAxe.EndPoint)
                    vNextOffset = Math.Round(Distance(vAxe.StartPoint, vNextProj) / 2, 3)

                ElseIf j = vAxes.Count - 1 Then 'Last axe
                    Dim vPreviousAxe = vAxes(j - 1)
                    Dim vPreviousProj = GetNormalPoint(vAxe.StartPoint, vPreviousAxe.StartPoint, vPreviousAxe.EndPoint)
                    vPreviousOffset = Math.Round(Distance(vPreviousProj, vAxe.StartPoint) / 2, 3)

                Else
                    Dim vPreviousAxe = vAxes(j - 1)
                    Dim vPreviousProj = GetNormalPoint(vAxe.StartPoint, vPreviousAxe.StartPoint, vPreviousAxe.EndPoint)
                    vPreviousOffset = Math.Round(Distance(vPreviousProj, vAxe.StartPoint) / 2, 3)
                    Dim vNextAxe = vAxes(j + 1)
                    Dim vNextProj = GetNormalPoint(vAxe.StartPoint, vNextAxe.StartPoint, vNextAxe.EndPoint)
                    vNextOffset = Math.Round(Distance(vAxe.StartPoint, vNextProj) / 2, 3)
                End If

                'Intersection
                Dim errCode As ErrorCode
                Dim intersPt = GetIntersection(uAxe.StartPoint, uAxe.EndPoint, vAxe.StartPoint, vAxe.EndPoint, errCode, True)
                If errCode <> ErrorCode.noError Then
                    Continue For
                End If

                'Save tgm box
                Dim boxName = $"{launchedStorey.Name}_{objType}-U{i}V{j}"
                Dim gridViewBoxTgm = boxVisuWs.SmartAddMetaObject(boxName, ProjectReference.Constants.boxTemplateName)
                gridViewBoxTgm.SmartAddAttribute(ProjectReference.Constants.boxNameAttName, boxName)
                gridViewBoxTgm.SmartAddAttribute("FileLevel", launchedStorey.Name)
                myIfcFile.ProjWs.Push({gridViewBoxTgm}.ToList)

                'Construct box
                intersPt.Z = elevationInf
                Dim minXpt = intersPt.Translate2(xVec, -uPreviousOffset)
                Dim maxXpt = intersPt.Translate2(xVec, uNextOffset)

                Dim basGauArrPt = minXpt.Translate2(yVec, -vPreviousOffset)
                Dim basDroiArrPt = maxXpt.Translate2(yVec, -vPreviousOffset)
                Dim basGauAvPt = minXpt.Translate2(yVec, vNextOffset)
                Dim basDroiAvPt = maxXpt.Translate2(yVec, vNextOffset)

                Dim zVec As New Vector(0.0, 0.0, 1.0)
                Dim hautGauArrPt = basGauArrPt.Translate2(zVec, elevationSup - elevationInf)
                Dim hautDroiArrPt = basDroiArrPt.Translate2(zVec, elevationSup - elevationInf)
                Dim hautGauAvPt = basGauAvPt.Translate2(zVec, elevationSup - elevationInf)
                Dim hautDroiAvPt = basDroiAvPt.Translate2(zVec, elevationSup - elevationInf)

                Dim geometryBox = New Treegram.GeomFunctions.Geometries.ParallelepipedGeometry(basGauArrPt, basDroiArrPt, basGauAvPt, basDroiAvPt,
                                                                        hautGauArrPt, hautDroiArrPt, hautGauAvPt, hautDroiAvPt,
                                                                        New Color4(0.9F, 0.9F, 0.9F, 1.0F))
                geometryBox.ReorderIndices()
                Dim geometryModel = New M4D.Deixi.Core.Model() With {.Geometry = geometryBox, .IsInstance = False, .InstanceMatrix = SharpDX.Matrix.Identity, .ReplaceColor = False,
                                            .Tag = New Tuple(Of Integer, Integer, Integer)(gridViewBoxTgm.Id.Container.Id - 1, gridViewBoxTgm.Id.Entity.Id, 1)}
                geometryModel.IsInstance = True
                geometryModel.InstanceMatrix =  Treegram.GeomFunctions.GeoReferencing.GetCorrectionMatrix(boxVisuWs, myIfcFile.ScanWs)

                visuBoxes.Add(geometryModel)
            Next
        Next

        'Save it
        myIfcFile.ProjWs.PushAllModifiedEntities()
        If visuBoxes.Count > 0 Then
            Writer.WriteModels(box3dPath, visuBoxes, True, CType(1, Byte), 1.0F)
        End If

        CreateGuiOutputTree()
        Return New SucceededActionResult("Launched", System.Drawing.Color.Green, New List(Of Element) From {Grids})
    End Function

    Public Overrides Sub CreateGuiOutputTree()
        'Présentation des résultats grâce à une zone de filtrage
        OutputTree = TempWorkspace.AddTree("Visualization Boxes from Grid")
        OutputTree.SmartAddNode("object.type", "IfcGrid")
        OutputTree.SmartAddNode("object.type", "GridViewBox")

    End Sub

End Class
