''' <summary>
''' Represents a Vertex for polygon operations
''' </summary>
''' For surface operations we construct a data structure with two surfaces
''' The structure created is a representation of the polygons where we add the intersection points between the two perimeters
''' See the Greiner/Hormann method for clipping polygons for more informations
Public Class VertexData
    'Private Const PRECISION As Integer = 4

    ''' <summary>
    ''' Point of the Vertex
    ''' </summary>
    Public Point As Point

    ''' <summary>
    ''' Following Vertex, Precedent Vertex, Neigbour Vertex
    ''' </summary>
    Public Suiv, Prec, Neighbour As VertexData

    ''' <summary>
    ''' Intersect and entry_exit properties of Vertex
    ''' </summary>
    Public Intersect, EntryExit As Boolean?
    Public Alpha As Double?

    ''' <summary>
    ''' Creates a new VertexData by Point
    ''' </summary>
    ''' <param name="point"> Point location of Vertex </param>
    ''' <param name="prec"> Precedent Vertex </param>
    ''' <param name="suiv"> Following Vertex </param>
    ''' <param name="intersect"> Intersect property of Vertex </param>
    Public Sub New(point As Point, Optional prec As VertexData = Nothing, Optional suiv As VertexData = Nothing, Optional intersect As Boolean = False)
        Me.Point = point
        Me.Suiv = suiv
        Me.Prec = prec
        Me.Intersect = intersect
    End Sub

    Public Shared Operator =(vertex As VertexData, other As VertexData) As Boolean
        If IsNothing(vertex) Or IsNothing(other) Then
            Throw New NullReferenceException
        Else
            Return (Math.Round(vertex.Point.X, -DefaultTolerance) = Math.Round(other.Point.X, -DefaultTolerance) AndAlso Math.Round(vertex.Point.Y, -DefaultTolerance) = Math.Round(other.Point.Y, -DefaultTolerance) AndAlso Math.Round(vertex.Point.Z, -DefaultTolerance) = Math.Round(other.Point.Z, -DefaultTolerance))
        End If
    End Operator
    Public Shared Operator <>(vertex As VertexData, other As VertexData) As Boolean
        Return Not vertex = other
    End Operator

    ''' <summary>
    ''' Check if the vertex is in the polygon formed by Vertices
    ''' </summary>
    ''' <param name="vertices"> List of Vertex forming the polygon </param>
    ''' <returns></returns>
    Public Function IsVertexInsidePolygonVertices(vertices As List(Of VertexData)) As Boolean
        Dim testPoint = Me.Point
        Dim polyList As New List(Of Point)
        For Each element In vertices
            polyList.Add(element.Point)
        Next
        Dim polygon As New Surface(polyList, Surface.SurfaceType.Polygon)
        Dim wn = polygon.WindingNumber(testPoint)
        Return (wn Mod 2 = 1)
    End Function
End Class
