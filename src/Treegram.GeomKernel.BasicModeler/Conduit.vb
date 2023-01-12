''' <summary>
''' Represents a Conduit
''' </summary>
Public Class Conduit
    ' Ne pas utiliser pour le moment 


    Public Property Nodes As List(Of Point)
    Public Property HeightDirection As Vector
    Public Property Surface As Surface

    ''' <summary>
    ''' Create a new Conduit by surface and nodes
    ''' </summary>
    ''' <param name="surface"> Surface extruded </param>
    ''' <param name="heightDirection"> ? </param>
    ''' <param name="nodes"> Nodes defining the path of Conduit </param>
    Public Sub New(surface As Surface, heightDirection As Vector, nodes As List(Of Point))
        Me.Surface = surface
        Me.Nodes = nodes
        Me.HeightDirection = heightDirection
    End Sub

    ''' <summary>
    ''' Get the Volume measured
    ''' </summary>
    ''' <returns></returns>
    Public Function GetVolumes() As List(Of Volume)
        Dim list As New List(Of Volume)
        Dim rayon, longueur, largeur As Double
        Select Case Surface.Type
            Case Surface.SurfaceType.Circle
                rayon = Distance(Surface.Points(0), Surface.Points(1))
            Case Surface.SurfaceType.Rectangle
                longueur = Distance(Surface.Points(0), Surface.Points(1))
                largeur = Distance(Surface.Points(1), Surface.Points(2))
            Case Else
                Throw New Exception("Define surface type of the conduit before checking the volumes")
        End Select
        For I = 1 To Nodes.Count - 1
            Dim pointA As Point = Nodes(I - 1)
            Dim pointB As Point = Nodes(I)
            Select Case Surface.Type
                Case Surface.SurfaceType.Circle
                    Dim vector As New Vector With {.X = pointB.X - pointA.X, .Y = pointB.Y - pointA.Y, .Z = pointB.Z - pointA.Z}
                    vector = vector.ScaleOneVector
                    Dim zAxis As Vector = New Vector With {.X = 0, .Y = 0, .Z = 1}
                    Dim xAxis As Vector = Vector.CrossProduct(vector, zAxis)
                    zAxis = Vector.CrossProduct(vector, xAxis)
                    Dim pointZ As Point = zAxis.GetPointAtDistance(rayon, pointA)
                    Dim points As New List(Of Point)
                    points.Add(pointA)
                    points.Add(pointZ)
                    points.Add(pointB)
                    Dim volume As New Volume(points, Volume.VolumeType.Cylinder)
                    list.Add(volume)
                Case Surface.SurfaceType.Rectangle
                    Dim vector As New Vector With {.X = pointB.X - pointA.X, .Y = pointB.Y - pointA.Y, .Z = pointB.Z - pointA.Z}
                    vector = vector.ScaleOneVector
                    Dim zAxis As Vector = New Vector With {.X = 0, .Y = 0, .Z = 1}
                    Dim xAxis As Vector = Vector.CrossProduct(vector, zAxis)
                    zAxis = Vector.CrossProduct(vector, xAxis)
                    Dim topPoint As Point = zAxis.GetPointAtDistance(largeur / 2, pointA)
                    Dim originPoint As Point = xAxis.GetPointAtDistance(-longueur / 2, topPoint)
                    Dim widthPoint As Point = xAxis.GetPointAtDistance(longueur, originPoint)
                    Dim heightPoint As Point = zAxis.GetPointAtDistance(-largeur, originPoint)
                    Dim lengthPoint As Point = vector.GetPointAtDistance(Distance(pointA, pointB), originPoint)
                    Dim points As New List(Of Point)
                    points.Add(originPoint)
                    points.Add(widthPoint)
                    points.Add(heightPoint)
                    points.Add(lengthPoint)
                    Dim volume As New Volume(points, Volume.VolumeType.Parallelepiped)
                    list.Add(volume)
            End Select
        Next
        Return list
    End Function

End Class
