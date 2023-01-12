Imports System.Runtime.CompilerServices
Imports Treegram.GeomLibrary
Imports Treegram.Bam.Libraries.AEC
Imports Treegram.GeomFunctions
Imports Treegram.GeomKernel.BasicModeler.Surface

Public Module AecOpeningExtensions
    'Extensions using BasicModeler

    ''' <summary>Get the barycenter of the door frame. Caution this property uses object models</summary>
    Public Function FrameCenterPoint(aecOpening) As Point3D

        'get heigth for each model
        Dim heightList As New List(Of Double)
        For Each objModel In aecOpening.Models
            Dim minZ = Double.PositiveInfinity
            Dim maxZ = Double.NegativeInfinity
            For Each vert In objModel.VerticesTesselation
                If vert.Z < minZ Then minZ = vert.Z
                If vert.Z > maxZ Then maxZ = vert.Z
            Next
            heightList.Add(Math.Round(maxZ - minZ, 3))
        Next

        'sort models by height
        Dim array() As IList = {heightList, aecOpening.Models}
        CommonFunctions.SortLists(Of Double)(array, Function(i1 As Double, i2 As Double)
                                                        Return i1.CompareTo(i2)
                                                    End Function)

        'get the barycenter for the heighest model, supposed to be the door frame
        Dim frameCenterPt = Treegram.GeomFunctions.Models.GetModelVerticesBarycenter(aecOpening.Models.Last)
        Return New Point3D(frameCenterPt)
    End Function


    ''' <summary>
    ''' Get IfcOpeningElement barycenter - Can be reached directly from an IfcDoor or IfcWindow - Needs GetOpeningElementsRelationsWithWalls First !!!
    ''' </summary>
    <Extension()>
    Public Function ReservationGravityCenter(aecOpen As AecOpening) As Point3D
        Dim ifcOpenMod = aecOpen.Reservation3dModel()
        If ifcOpenMod Is Nothing Then
            Return Nothing
        End If
        Dim ptList = ifcOpenMod.SelectMany(Function(o) o.Points).ToList
        Return Treegram.GeomKernel.BasicModeler.Barycenter(ptList).ToPoint3d
    End Function


    ''' <summary>
    ''' Get IfcOpeningElement 3D model - Can be reached directly from an IfcDoor or IfcWindow
    ''' </summary>
    <Extension()>
    Public Function Reservation3dModel(aecOpen As AecOpening) As List(Of Treegram.GeomKernel.BasicModeler.Surface)
        'Get wall thickness
        If aecOpen.WallOwner Is Nothing OrElse aecOpen.WallOwner.ThicknessByTgm = Nothing OrElse aecOpen.WallOwner.HorizontalProfile Is Nothing Then
            Return Nothing
        End If
        Dim wallThickness = aecOpen.WallOwner.ThicknessByTgm

        ''Get reservation object
        'Dim aecResa As AecOpening = Nothing
        'If aecOpen.IsReservation Then
        '    aecResa = aecOpen
        'ElseIf aecOpen.Metaobject.GetParents("IfcOpeningElement").FirstOrDefault IsNot Nothing Then
        '    aecResa = New AecOpening(aecOpen.Metaobject.GetParents("IfcOpeningElement").FirstOrDefault)
        'End If

        'Get profile
        Dim resaProfile As Profile3D = Nothing
        'If aecResa IsNot Nothing Then
        '    resaProfile = aecResa.IfcProfile
        'Else 'Curtain walls are concerned
        resaProfile = aecOpen.OpeningProfile
            If resaProfile Is Nothing Then
                resaProfile = aecOpen.OverallProfile
            End If
        'End If
        Dim resaSurface = resaProfile.ToBasicModelerSurface

        'To deal with weird profiles - the profile has to be VERTICAL and PERPENDICULAR to wall axis
        If resaSurface Is Nothing OrElse resaSurface.GetNormal.Z <> 0.0 Then Return Nothing
        Dim myAngle = Math.Abs(resaSurface.GetNormal.Angle(aecOpen.WallOwner.Axis.ToBasicModelerLine.Vector, True)) Mod 180
        Dim ANGLE_TOLERANCE = 5 'deg
        If myAngle < 90 - ANGLE_TOLERANCE Or myAngle > 90 + ANGLE_TOLERANCE Then ' not perpendicular to wall axis ! (ex : THIAIS)
            Dim openProfile = aecOpen.OpeningProfile
            If openProfile Is Nothing Then
                openProfile = aecOpen.OverallProfile
            End If
            resaSurface = openProfile.ToBasicModelerSurface
            If resaSurface Is Nothing OrElse resaSurface.GetNormal.Z <> 0.0 Then Return Nothing
        End If


        Dim myPtList As List(Of Treegram.GeomKernel.BasicModeler.Point) = resaSurface.Points.ToList

        ''Get Opening X and Y vector :
        Dim indexList As New List(Of Integer) From {0, 1, 2, 3}
        Dim ptElevations = myPtList.Select(Function(o) o.Z).ToList
        Dim array() As IList = {ptElevations, indexList}
        CommonFunctions.SortLists(Of Double)(array, Function(i1 As Double, i2 As Double)
                                                        Return i1.CompareTo(i2)
                                                    End Function)
        Dim openingXvec As New Treegram.GeomKernel.BasicModeler.Vector(myPtList(indexList(0)), myPtList(indexList(1)))
        Dim openingYvec = openingXvec.Rotate2(Math.PI / 2)

        'Test min and max with object projection - Could be object vertices, but it's faster like this
        Dim minY = Double.PositiveInfinity
        Dim maxY = Double.NegativeInfinity
        Dim forewardPoint As New Treegram.GeomKernel.BasicModeler.Point(myPtList(0).X, myPtList(0).Y)
        For Each pt In aecOpen.WallOwner.HorizontalProfile.Points
            Dim newCoord = Treegram.GeomKernel.BasicModeler.ChangeCoordPointInAnotherAxisSyst({pt.X, pt.Y, 0.0}, {forewardPoint.X, forewardPoint.Y, 0.0}, {openingXvec.X, openingXvec.Y, 0.0}, {openingYvec.X, openingYvec.Y, 0.0}, {0, 0, 1})
            If newCoord(1) < minY Then
                minY = newCoord(1)
            End If
            If newCoord(1) > maxY Then
                maxY = newCoord(1)
            End If
        Next

        'Translate points
        '- front
        Dim frontPtList As New List(Of Treegram.GeomKernel.BasicModeler.Point)
        frontPtList.Add(myPtList(0).Translate2(openingYvec, maxY)) 'bottom left
        frontPtList.Add(myPtList(1).Translate2(openingYvec, maxY)) 'bottom right
        frontPtList.Add(myPtList(2).Translate2(openingYvec, maxY)) 'top right
        frontPtList.Add(myPtList(3).Translate2(openingYvec, maxY)) 'top left

        '- back
        Dim backPtList As New List(Of Treegram.GeomKernel.BasicModeler.Point)
        backPtList.Add(myPtList(0).Translate2(openingYvec, minY)) 'bottom left
        backPtList.Add(myPtList(1).Translate2(openingYvec, minY)) 'bottom right
        backPtList.Add(myPtList(2).Translate2(openingYvec, minY)) 'top right
        backPtList.Add(myPtList(3).Translate2(openingYvec, minY)) 'top left

        'Create surfaces
        Dim bottomList As New List(Of Treegram.GeomKernel.BasicModeler.Point) From {frontPtList(0), frontPtList(1), backPtList(1), backPtList(0)}
        Dim topList As New List(Of Treegram.GeomKernel.BasicModeler.Point) From {frontPtList(3), frontPtList(2), backPtList(2), backPtList(3)}
        Dim leftList As New List(Of Treegram.GeomKernel.BasicModeler.Point) From {frontPtList(0), frontPtList(3), backPtList(3), backPtList(0)}
        Dim rightList As New List(Of Treegram.GeomKernel.BasicModeler.Point) From {frontPtList(1), frontPtList(2), backPtList(2), backPtList(1)}
        Dim frontList As New List(Of Treegram.GeomKernel.BasicModeler.Point) From {frontPtList(0), frontPtList(1), frontPtList(2), frontPtList(3)}
        Dim backList As New List(Of Treegram.GeomKernel.BasicModeler.Point) From {backPtList(0), backPtList(1), backPtList(2), backPtList(3)}
        Dim surfacesList As New List(Of Treegram.GeomKernel.BasicModeler.Surface)
        surfacesList.Add(New Treegram.GeomKernel.BasicModeler.Surface(bottomList, SurfaceType.Polygon))
        surfacesList.Add(New Treegram.GeomKernel.BasicModeler.Surface(topList, SurfaceType.Polygon))
        surfacesList.Add(New Treegram.GeomKernel.BasicModeler.Surface(leftList, SurfaceType.Polygon))
        surfacesList.Add(New Treegram.GeomKernel.BasicModeler.Surface(rightList, SurfaceType.Polygon))
        surfacesList.Add(New Treegram.GeomKernel.BasicModeler.Surface(frontList, SurfaceType.Polygon))
        surfacesList.Add(New Treegram.GeomKernel.BasicModeler.Surface(backList, SurfaceType.Polygon))
        '---CAREFUL--- Indications "left, right, etc." are not always corrects

        Return surfacesList
    End Function


End Module

