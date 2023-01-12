Imports System.Runtime.CompilerServices
Imports System.Web
Imports Treegram.GeomLibrary
Imports Treegram.GeomKernel.BasicModeler

Public Module GeomLibraryExtensions

    <Extension()>
    Public Function ToBasicModelerPoint(shPt As SharpDX.Vector3) As Treegram.GeomKernel.BasicModeler.Point
        Return New Treegram.GeomKernel.BasicModeler.Point(shPt.X, shPt.Y, shPt.Z)
    End Function
    <Extension()>
    Public Function ToBasicModelerPoint(point As Point3D) As Treegram.GeomKernel.BasicModeler.Point
        Return New Treegram.GeomKernel.BasicModeler.Point(point.Point.X, point.Point.Y, point.Point.Z)
    End Function
    <Extension()>
    Public Function ToBasicModelerPoint(shPt As SharpDX.Vector2) As Treegram.GeomKernel.BasicModeler.Point
        Return New Treegram.GeomKernel.BasicModeler.Point(shPt.X, shPt.Y)
    End Function
    <Extension()>
    Public Function ToBasicModelerPoint(point As Point2D) As Treegram.GeomKernel.BasicModeler.Point
        Return New Treegram.GeomKernel.BasicModeler.Point(point.Point.X, point.Point.Y)
    End Function
    <Extension()>
    Public Function ToBasicModelerVector(shPt As SharpDX.Vector2) As Treegram.GeomKernel.BasicModeler.Vector
        Return New Treegram.GeomKernel.BasicModeler.Vector(shPt.X, shPt.Y, 0.0)
    End Function
    <Extension()>
    Public Function ToBasicModelerLine(curve As Curve2D) As Treegram.GeomKernel.BasicModeler.Line
        Return New Treegram.GeomKernel.BasicModeler.Line(curve.Points.First.ToBasicModelerPoint, curve.Points(1).ToBasicModelerPoint)
    End Function
    <Extension()>
    Public Function ToBasicModelerLine(curve As Curve3D) As Treegram.GeomKernel.BasicModeler.Line
        Return New Treegram.GeomKernel.BasicModeler.Line(curve.Points.First.ToBasicModelerPoint, curve.Points(1).ToBasicModelerPoint)
    End Function
    <Extension()>
    Public Function ToBasicModelerSurface(profile As Profile3D) As Treegram.GeomKernel.BasicModeler.Surface
        Return New Treegram.GeomKernel.BasicModeler.Surface(profile.Points.Select(Function(o) o.ToBasicModelerPoint).ToList, Treegram.GeomKernel.BasicModeler.Surface.SurfaceType.Polygon)
    End Function
    <Extension()>
    Public Function ToBasicModelerSurface(profile As Profile2D) As Treegram.GeomKernel.BasicModeler.Surface
        Return New Treegram.GeomKernel.BasicModeler.Surface(profile.Points.Select(Function(o) o.ToBasicModelerPoint).ToList, Treegram.GeomKernel.BasicModeler.Surface.SurfaceType.Polygon)
    End Function
    <Extension()>
    Public Function ToBasicModelerBoundary(profile As Profile3D) As List(Of Treegram.GeomKernel.BasicModeler.Line)
        Dim boundary As New List(Of Line)
        For k As Integer = 0 To profile.Points.Count - 1
            If k = profile.Points.Count - 1 Then
                boundary.Add(New Line(profile.Points(k).ToBasicModelerPoint, profile.Points.First.ToBasicModelerPoint))
            Else
                boundary.Add(New Line(profile.Points(k).ToBasicModelerPoint, profile.Points(k + 1).ToBasicModelerPoint))
            End If
        Next
        Return boundary
    End Function
    <Extension()>
    Public Function ToBasicModelerBoundary(profile As Profile2D) As List(Of Treegram.GeomKernel.BasicModeler.Line)
        Dim boundary As New List(Of Line)
        For k As Integer = 0 To profile.Points.Count - 1
            If k = profile.Points.Count - 1 Then
                boundary.Add(New Line(profile.Points(k).ToBasicModelerPoint, profile.Points.First.ToBasicModelerPoint))
            Else
                boundary.Add(New Line(profile.Points(k).ToBasicModelerPoint, profile.Points(k + 1).ToBasicModelerPoint))
            End If
        Next
        Return boundary
    End Function

End Module

