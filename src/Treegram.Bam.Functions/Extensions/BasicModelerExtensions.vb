Imports System.Runtime.CompilerServices
Imports SharpDX
Imports Treegram.GeomLibrary

Public Module BasicModelerExtensions

    <Extension()>
    Public Function ToVector2(point As Treegram.GeomKernel.BasicModeler.Point) As Vector2
        Return New Vector2(point.X, point.Y)
    End Function
    <Extension()>
    Public Function ToVector2(vector As Treegram.GeomKernel.BasicModeler.Vector) As Vector2
        Return New Vector2(vector.X, vector.Y)
    End Function

    <Extension()>
    Public Function ToVector3(point As Treegram.GeomKernel.BasicModeler.Point) As Vector3
        Return New Vector3(point.X, point.Y, point.Z)
    End Function
    <Extension()>
    Public Function ToPoint2d(point As Treegram.GeomKernel.BasicModeler.Point) As Point2D
        Return New Point2D(point.X, point.Y)
    End Function
    <Extension()>
    Public Function ToPoint3d(point As Treegram.GeomKernel.BasicModeler.Point) As Point3D
        Return New Point3D(point.X, point.Y, point.Z)
    End Function

    <Extension()>
    Public Function ToProfile3d(surface As Treegram.GeomKernel.BasicModeler.Surface) As Profile3D
        Dim profile As New Profile3D
        For Each basicPt In surface.Points
            profile.Points.Add(New Vector3(basicPt.X, basicPt.Y, basicPt.Z))
        Next
        Return profile
    End Function
    <Extension()>
    Public Function ToProfile2d(surface As Treegram.GeomKernel.BasicModeler.Surface) As Profile2D
        Dim profile As New Profile2D
        For Each basicPt In surface.Points
            profile.Points.Add(New Vector2(basicPt.X, basicPt.Y))
        Next
        Return profile
    End Function
    <Extension()>
    Public Function ToCurve3d(line As Treegram.GeomKernel.BasicModeler.Line) As Curve3D
        Dim curve As New Curve3D()
        curve.Points.Add(line.StartPoint.ToVector3)
        curve.Points.Add(line.EndPoint.ToVector3)
        Return curve
    End Function
    <Extension()>
    Public Function ToCurve2d(line As Treegram.GeomKernel.BasicModeler.Line) As Curve2D
        Dim curve As New Curve2D()
        curve.Points.Add(line.StartPoint.ToVector2)
        curve.Points.Add(line.EndPoint.ToVector2)
        Return curve
    End Function
End Module

