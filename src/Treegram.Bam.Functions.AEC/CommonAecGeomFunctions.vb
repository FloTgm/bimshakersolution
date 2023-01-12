Imports M4D.Treegram.Core.Entities
Imports M4D.Deixi.Core
Imports Treegram.GeomKernel.BasicModeler
Imports Treegram.GeomKernel.BasicModeler.Surface
Imports M4D.Treegram.Core.Extensions.Entities
Imports Treegram.Bam.Libraries.AEC
Imports Treegram.Bam.Libraries.AuditFile
Imports SharpDX
Imports Treegram.GeomLibrary
Imports Treegram.GeomFunctions
Imports M4D.Treegram.Core.Extensions.Kernel

Public Module CommonAecGeomFunctions

    Private Function ModelsHeight(models As List(Of Model)) As Double
        If models.Count = 0 Then
            Return Nothing
        End If
        Dim minZ = Double.PositiveInfinity
        Dim maxZ = Double.NegativeInfinity
        For Each model In models
            For Each vert In model.VerticesTesselation
                If vert.Z < minZ Then minZ = vert.Z
                If vert.Z > maxZ Then maxZ = vert.Z
            Next
        Next
        Return CDbl(Math.Round(maxZ - minZ, 4))
    End Function

    Public Sub CompleteConnectedSpaces(analysisSpaceList As List(Of AecSpace))
        Dim i, j As Integer
        Dim nb = 0
        For i = 0 To analysisSpaceList.Count - 1
            Dim iSpace = analysisSpaceList(i)
            For j = i + 1 To analysisSpaceList.Count - 1
                Dim jSpace = analysisSpaceList(j)
                If iSpace.TgmBuildingStoreysAtt = jSpace.TgmBuildingStoreysAtt Then 'for 'By Building' analysis

                    'Dim iSpaceProfileSurface As new Treegram.GeomKernel.BasicModeler.Surface(iSpace.Profile.Select(Function(w) w.StartPoint).ToList(), SurfaceType.Polygon)
                    'Dim jSpaceProfileSurface As new Treegram.GeomKernel.BasicModeler.Surface(jSpace.Profile.Select(Function(w) w.StartPoint).ToList(), SurfaceType.Polygon)
                    Dim iSpaceProfileSurface = CommonGeomFunctions.SurfaceCleaningMethods(iSpace.HorizontalProfile.ToBasicModelerSurface)
                    Dim jSpaceProfileSurface = CommonGeomFunctions.SurfaceCleaningMethods(jSpace.HorizontalProfile.ToBasicModelerSurface)

                    If Treegram.GeomKernel.BasicModeler.Distance(iSpaceProfileSurface, jSpaceProfileSurface) < 0.009 Then
                        If jSpace.ConnectedSpaces Is Nothing Then
                            jSpace.ConnectedSpaces = New List(Of AecSpace) From {iSpace}
                        Else
                            jSpace.ConnectedSpaces.Add(iSpace)
                        End If
                        If iSpace.ConnectedSpaces Is Nothing Then
                            iSpace.ConnectedSpaces = New List(Of AecSpace) From {jSpace}
                        Else
                            iSpace.ConnectedSpaces.Add(jSpace)
                        End If
                        nb += 1
                    End If
                End If
            Next
            If iSpace.ConnectedSpaces Is Nothing Then
                iSpace.ConnectedSpaces = New List(Of AecSpace)
            End If
        Next

    End Sub


    ''' <summary>
    ''' Assign Models according to tgm types
    ''' </summary>
    ''' <param name="myAecObj"></param>
    Public Function SmartAddLocalPlacement(ByRef myAecObj As AecObject, myFile As FileObject) As Boolean
        If Not myAecObj.VectorU.IsZero And Not myAecObj.VectorV.IsZero Then
            Return True 'Already consolidated
        End If

        If myFile.IsIfc Then
            Dim succeeded = True
            If myAecObj.Type = "IfcCurtainWall" Then
                succeeded = DefineLocalPlacementForCurtainWall(myAecObj) 'Get rid of some composants to get the right UV vectors !!
            ElseIf myAecObj.Type = "IfcStair" Then
                succeeded = DefineLocalPlacementForStair(myAecObj) 'Get rid of some composants to get the right UV vectors !!
            Else 'IfcWall, IfcDoor, IfcWindow...
                If myAecObj.Metaobject.GetAttribute("LocalPlacement") Is Nothing Then
                    succeeded = False
                Else
                    'Sometimes both vectors are inverted - Must be longer than large !
                    Dim objUvec = New Vector2(CDbl(myAecObj.Metaobject.SmartGetAttribute("LocalPlacement/_\Axis X/_\X").Value), CDbl(myAecObj.Metaobject.SmartGetAttribute("LocalPlacement/_\Axis X/_\Y").Value))
                    Dim objVvec = New Vector2(CDbl(myAecObj.Metaobject.SmartGetAttribute("LocalPlacement/_\Axis Y/_\X").Value), CDbl(myAecObj.Metaobject.SmartGetAttribute("LocalPlacement/_\Axis Y/_\Y").Value))
                    myAecObj.VectorU = objUvec
                    myAecObj.VectorV = objVvec
                    If objUvec.IsZero Or objVvec.IsZero Then
                        succeeded = False
                    Else
                        If myAecObj.MaxU - myAecObj.MinU < myAecObj.MaxV - myAecObj.MinV Then
                            myAecObj.VectorU = objVvec
                            myAecObj.VectorV = objUvec.ToBasicModelerVector.Reverse.ToVector2
                        End If
                    End If
                End If
            End If
            If Not succeeded Then
                Return False
            End If

        ElseIf myFile.IsRevit Then
            Dim succeeded = True
            If myAecObj.GetType Is GetType(AecWall) Then
                Dim axis = CType(myAecObj, AecWall).Axis

                If axis Is Nothing Then
                    'Lecture du scan Brut
                    Dim axisPoints = New List(Of Vector2)
                    If CType(myAecObj, AecWall).AdditionalGeometryExtension IsNot Nothing AndAlso CType(myAecObj, AecWall).AdditionalGeometryExtension.GetAttribute("Line", True) IsNot Nothing Then
                        Dim curveAtt = CType(myAecObj, AecWall).AdditionalGeometryExtension.GetAttribute("Line", True)
                        If curveAtt IsNot Nothing Then
                            For Each pointAtt In curveAtt.Attributes
                                axisPoints.Add(New SharpDX.Vector2(CSng(pointAtt.GetAttribute("X").Value), CSng(pointAtt.GetAttribute("Y").Value)))
                                axis = New Curve2D(axisPoints)
                            Next
                        End If
                    Else
                        succeeded = False
                    End If
                End If
                If axis.Points.Count = 2 Then
                    myAecObj.VectorU = New Vector(axis.Points.First.ToBasicModelerPoint, axis.Points.Last.ToBasicModelerPoint).ToVector2
                    myAecObj.VectorV = myAecObj.VectorU.ToBasicModelerVector.Rotate2(Math.PI / 2).ToVector2
                    Dim objUvec = myAecObj.VectorU
                    Dim objVvec = myAecObj.VectorV
                    If objUvec.IsZero Or objVvec.IsZero Then
                        succeeded = False
                    Else
                        If myAecObj.MaxU - myAecObj.MinU < myAecObj.MaxV - myAecObj.MinV Then
                            myAecObj.VectorU = objVvec
                            myAecObj.VectorV = objUvec.ToBasicModelerVector.Reverse.ToVector2
                        End If
                    End If
                End If
            Else
            End If
        Else
            Throw New Exception("This type of file is not implemented yet")
        End If

        'Save new local placement
        Dim axisSyst = New AxisSystem(myAecObj.VectorU, myAecObj.VectorV)
        axisSyst.Write(myAecObj.AdditionalGeometryExtension)
        Return True
    End Function

    ''' <summary>
    ''' Specific to curtain walls, get local placement of the first transparent model that composes it
    ''' </summary>
    ''' <param name="anaCWall"></param>
    ''' <returns></returns>
    Public Function DefineLocalPlacementForCurtainWall(ByRef anaCWall As AecObject) As Boolean
        Dim objUvec As Treegram.GeomKernel.BasicModeler.Vector = Nothing
        Dim objVvec As Treegram.GeomKernel.BasicModeler.Vector = Nothing
        'Dim newModelsList As New List(Of Model)
        Dim localPlactFound = False

        'Thanks to children transparent models, if none, take all models
        Dim curtainWallModels As List(Of Model) = Treegram.GeomFunctions.Models.GetTransparentModels(anaCWall.Metaobject, True)
        If curtainWallModels.Count = 0 Then
            curtainWallModels = Treegram.GeomFunctions.Models.GetGeometryModels(anaCWall.Metaobject, True)
        End If
        If curtainWallModels.Count = 0 Then
            Return False
        End If
        'Find out the widest face
        Dim curtainWallFaces = curtainWallModels.Select(Function(o) Treegram.GeomFunctions.Models.GetModelAsFacesV2(o, True)).SelectMany(Function(o) o).ToList
        If curtainWallFaces.Count = 0 Then Throw New Exception("Could not get transparent faces from model")
        Dim maxTransparSurf = curtainWallFaces.OrderBy(Function(surf) CDbl(surf.Area3D)).Last
        maxTransparSurf.CleanPoints(3)
        If maxTransparSurf.Points.Count < 3 Then Throw New Exception("Weird glazed surface")
        'Find out 3 points to build a plane
        Dim pt1 = maxTransparSurf.Points(0)
        Dim pt2 = maxTransparSurf.Points(1)
        Dim pt3 As Treegram.GeomKernel.BasicModeler.Point = Nothing
        Dim j As Integer
        For j = 2 To maxTransparSurf.Points.Count - 1
            Dim myPt = maxTransparSurf.Points(j)
            If Treegram.GeomKernel.BasicModeler.Distance(myPt, pt1, pt2, True) > 0.0001 Then
                pt3 = myPt
                Exit For
            End If
        Next
        'Find out U, V vectors
        If pt3 Is Nothing Then Throw New Exception("Weird glazed surface")
        Dim myPlane As New Treegram.GeomKernel.BasicModeler.Plane(pt1, pt2, pt3)
        If myPlane.Vector.Z < 0.001 Then
            objVvec = myPlane.Vector
            objUvec = objVvec.Rotate2(-Math.PI / 2)
            localPlactFound = True
        End If

        'For Each childTgm In anaCWall.Metaobject.GetChildren
        '    If childTgm.GetTgmType = "IfcDoor" Or childTgm.GetTgmType = "IfcWindow" Then 'Exclude DOORS and WINDOWS !!!!!!!!!!!!!!! A améliorer.......
        '        Continue For
        '    Else
        '        newModelsList.AddRange(Treegram.GeomFunctions.Models.GetGeometryModels(childTgm, True))
        '    End If

        '    'STEP2 - Through children
        '    'If Not localPlactFound Then
        '    '    If Treegram.GeomFunctions.Models.GetTransparentModels(childTgm).Count > 0 Then 'look for transparent models that are more reliable
        '    '        'Get vectors
        '    '        If BetaTgmFcts.GetAttribute(childTgm, "LocalPlacement") Is Nothing Then
        '    '            Continue For
        '    '        End If
        '    '        objUvec = new Treegram.GeomKernel.BasicModeler.Vector(CDbl(BetaTgmFcts.GetAttribute(childTgm, "LocalPlacement/_\Axis X/_\X").Value), CDbl(BetaTgmFcts.GetAttribute(childTgm, "LocalPlacement/_\Axis X/_\Y").Value), 0.0)
        '    '        objVvec = new Treegram.GeomKernel.BasicModeler.Vector(CDbl(BetaTgmFcts.GetAttribute(childTgm, "LocalPlacement/_\Axis Y/_\X").Value), CDbl(BetaTgmFcts.GetAttribute(childTgm, "LocalPlacement/_\Axis Y/_\Y").Value), 0.0)
        '    '        If objUvec.Length > 0.0 And objVvec.Length > 0.0 Then 'Parfois un des vecteurs est nul, peut-être dans le cas des poutres où il faudrait alors l'axe Z...
        '    '            localPlactFound = True
        '    '        End If
        '    '    End If
        '    'End If
        'Next

        ''STEP3 - Through extension <-- Not very reliable but take it as a last resort !
        'If Not localPlactFound Then
        '    Dim geomExt = anaCWall.Metaobject.SmartAddExtension( M4D.Treegram.Core.Extensions.Enums.ExtensionType.AdditionalGeometries)
        '    If BetaTgmFcts.GetAttribute(geomExt, "LocalPlacement") IsNot Nothing Then
        '        objUvec = new Treegram.GeomKernel.BasicModeler.Vector(CDbl(BetaTgmFcts.GetAttribute(geomExt, "LocalPlacement/_\Axis X/_\X").Value), CDbl(BetaTgmFcts.GetAttribute(geomExt, "LocalPlacement/_\Axis X/_\Y").Value), 0.0)
        '        objVvec = new Treegram.GeomKernel.BasicModeler.Vector(CDbl(BetaTgmFcts.GetAttribute(geomExt, "LocalPlacement/_\Axis Y/_\X").Value), CDbl(BetaTgmFcts.GetAttribute(geomExt, "LocalPlacement/_\Axis Y/_\Y").Value), 0.0)
        '        localPlactFound = True
        '    End If
        'End If

        ''Assign to analysis object
        'If newModelsList.Count <> 0 Then
        '    anaCWall.Models = newModelsList
        'Else
        '    anaCWall.Models = Treegram.GeomFunctions.Models.GetGeometryModels(anaCWall.Metaobject, True)
        'End If
        If Not localPlactFound Or anaCWall.Models.Count = 0 Then
            Return False
        End If

        'Sometimes both vectors are inverted - Must be longer than large !
        Dim tempCWall As New AecObject(anaCWall.Metaobject)
        'tempCWall.Models = anaCWall.Models
        tempCWall.VectorU = objUvec.ToVector2
        tempCWall.VectorV = objVvec.ToVector2
        If tempCWall.MaxU - tempCWall.MinU < tempCWall.MaxV - tempCWall.MinV Then
            anaCWall.VectorU = objVvec.ToVector2
            anaCWall.VectorV = objUvec.Reverse.ToVector2
        Else
            anaCWall.VectorU = objUvec.ToVector2
            anaCWall.VectorV = objVvec.ToVector2
        End If


        Return True
    End Function

    ''' <summary>
    ''' Specific to stairs, except IfcRailing in models
    ''' </summary>
    ''' <param name="anaStair"></param>
    ''' <returns></returns>
    Public Function DefineLocalPlacementForStair(ByRef anaStair As AecObject) As Boolean
        Dim objUvec As Treegram.GeomKernel.BasicModeler.Vector = Nothing
        Dim objVvec As Treegram.GeomKernel.BasicModeler.Vector = Nothing
        'Dim newModelsList As New List(Of Model)
        Dim localPlactFound = False

        'Try to get infos from IfcStair
        'newModelsList.AddRange(anaStair.Models(False))
        If anaStair.Metaobject.GetAttribute("LocalPlacement") IsNot Nothing Then
            objUvec = New Treegram.GeomKernel.BasicModeler.Vector(CDbl(anaStair.Metaobject.SmartGetAttribute("LocalPlacement/_\Axis X/_\X").Value), CDbl(anaStair.Metaobject.SmartGetAttribute("LocalPlacement/_\Axis X/_\Y").Value), 0.0)
            objVvec = New Treegram.GeomKernel.BasicModeler.Vector(CDbl(anaStair.Metaobject.SmartGetAttribute("LocalPlacement/_\Axis Y/_\X").Value), CDbl(anaStair.Metaobject.SmartGetAttribute("LocalPlacement/_\Axis Y/_\Y").Value), 0.0)
            localPlactFound = True
        End If

        'Try to get infos to its children, except IfcRailing
        For Each childTgm In anaStair.Metaobject.GetChildren
            If childTgm.GetTgmType = "IfcRailing" Then
                Continue For
                'Else
                '    newModelsList.AddRange(Treegram.GeomFunctions.Models.GetGeometryModels(childTgm, True))
            End If
            If Not localPlactFound Then
                'If Treegram.GeomFunctions.Models.GetTransparentModels(childTgm, True).Count > 0 Then' WTF ????
                'Get vectors
                If childTgm.GetAttribute("LocalPlacement") Is Nothing Then
                    Continue For
                End If
                objUvec = New Treegram.GeomKernel.BasicModeler.Vector(CDbl(childTgm.SmartGetAttribute("LocalPlacement/_\Axis X/_\X").Value), CDbl(childTgm.SmartGetAttribute("LocalPlacement/_\Axis X/_\Y").Value), 0.0)
                objVvec = New Treegram.GeomKernel.BasicModeler.Vector(CDbl(childTgm.SmartGetAttribute("LocalPlacement/_\Axis Y/_\X").Value), CDbl(childTgm.SmartGetAttribute("LocalPlacement/_\Axis Y/_\Y").Value), 0.0)
                localPlactFound = True
                'End If
            End If
        Next

        'Assign to analysis object
        'anaStair.Models = newModelsList
        If Not localPlactFound Then
            Return False
        End If
        anaStair.VectorU = objUvec.ToVector2
        anaStair.VectorV = objVvec.ToVector2

        Return True
    End Function

    Public Function GetGroup2dBbox(elevationDbl As Double, objReq As List(Of MetaObject), extrapol As Double) As Surface
        'Get min and max from bbox
        Dim minX = Double.PositiveInfinity
        Dim minY = Double.PositiveInfinity
        Dim maxX = Double.NegativeInfinity
        Dim maxY = Double.NegativeInfinity

        For Each objTgm In objReq
            Dim aecObj As New AecObject(objTgm)
            Dim bboxMinMax = aecObj.BoundingBoxMinMax
            If bboxMinMax.Item1 Is Nothing Then
                Continue For
            End If
            If CDbl(bboxMinMax.Item1.Point.X) < minX Then minX = CDbl(bboxMinMax.Item1.Point.X)
            If bboxMinMax.Item1.Point.Y < minY Then minY = bboxMinMax.Item1.Point.Y
            If bboxMinMax.Item2.Point.X > maxX Then maxX = bboxMinMax.Item2.Point.X
            If bboxMinMax.Item2.Point.Y > maxY Then maxY = bboxMinMax.Item2.Point.Y
        Next
        If minX = Double.PositiveInfinity Or minY = Double.PositiveInfinity Or maxX = Double.NegativeInfinity Or maxY = Double.NegativeInfinity Then
            Return Nothing
        End If

        'Create surface
        Dim ptList As New List(Of Treegram.GeomKernel.BasicModeler.Point) From {
                New Treegram.GeomKernel.BasicModeler.Point(minX - extrapol, minY - extrapol, elevationDbl),
                New Treegram.GeomKernel.BasicModeler.Point(minX - extrapol, maxY + extrapol, elevationDbl),
                New Treegram.GeomKernel.BasicModeler.Point(maxX + extrapol, maxY + extrapol, elevationDbl),
                New Treegram.GeomKernel.BasicModeler.Point(maxX + extrapol, minY - extrapol, elevationDbl)}
        Dim minMaxSurf As New Treegram.GeomKernel.BasicModeler.Surface(ptList, SurfaceType.Polygon)
        Return minMaxSurf
    End Function

    Public Sub DispatchContainerBoundaries(objList As List(Of AecObject), ByRef containedInList As List(Of AecObject), ByRef containerList As List(Of AecObject), Optional containedIfAlone As Boolean = True)


        Dim i As Integer = 0
        For Each myObj In objList
            Dim containanceCount = 0
            Dim isContainedCount = 0

            Dim j As Integer = 0
            For Each otherObj In objList
                If i <> j AndAlso Treegram.GeomKernel.BasicModeler.Distance(myObj.HorizontalProfile.ToBasicModelerSurface, otherObj.HorizontalProfile.ToBasicModelerSurface) = 0 Then

                    ''DEBUG
                    'Dim debug As Boolean = False
                    'If debug Then
                    '    Dim nullDistanceList As New List(of Treegram.GeomKernel.BasicModeler.Surface) From {myObj.ProfileSurface, otherObj.ProfileSurface}
                    '    DrawPolylines(nullDistanceList, "NullDistance", True, System.Drawing.Color.DarkOrange)
                    'End If

                    Dim intersArea = CommonGeomFunctions.GetSuperpositionArea(myObj.HorizontalProfile.ToBasicModelerSurface, otherObj.HorizontalProfile.ToBasicModelerSurface)
                    If intersArea / otherObj.HorizontalProfile.ToBasicModelerSurface.Area > 0.99 And intersArea / myObj.HorizontalProfile.ToBasicModelerSurface.Area < 1.0 Then 'HYPOTHESE !!!
                        containanceCount += 1
                    ElseIf intersArea / myObj.HorizontalProfile.ToBasicModelerSurface.Area > 0.99 And intersArea / otherObj.HorizontalProfile.ToBasicModelerSurface.Area < 1.0 Then 'HYPOTHESE !!!
                        isContainedCount += 1
                    End If
                End If
                j += 1
            Next
            'If containanceCount = 0 Then
            '    ContainedInList.Add(myObj)
            'Else
            '    ContainerList.Add(myObj)
            'End If

            If containanceCount > 0 Then
                containerList.Add(myObj)
            End If
            If isContainedCount > 0 Then
                containedInList.Add(myObj)
            End If
            If containanceCount = 0 And isContainedCount = 0 Then
                If containedIfAlone Then
                    containedInList.Add(myObj)
                Else
                    containerList.Add(myObj)
                End If
            End If

            i += 1
        Next

    End Sub


End Module



