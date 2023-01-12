''' <summary>
''' Represents a Curve
''' </summary>
Public MustInherit Class Curve

        ''' <summary>
        ''' Length of the Curve
        ''' </summary>
        ''' <returns> Length </returns>
        MustOverride ReadOnly Property Length As Double

        ''' <summary>
        ''' MidPoint of the Curve
        ''' </summary>
        ''' <returns> MidPoint </returns>
        MustOverride ReadOnly Property MidPoint As Point

        ''' <summary>
        ''' StartPoint of the Curve
        ''' </summary>
        ''' <returns> StartPoint </returns>
        MustOverride Property StartPoint As Point

        ''' <summary>
        ''' EndPoint of the Curve
        ''' </summary>
        ''' <returns> EndPoint </returns>
        MustOverride Property EndPoint As Point

#Region "Old"
        'Public Shared Widening Operator CType(points As Point()) As Curve
        '    If points.Count <> 2 Then
        '        Throw New ArgumentOutOfRangeException("points", "Nombre de points incorret.")
        '    ElseIf points(0) Is Nothing Or points(1) Is Nothing Then
        '        Throw New NullReferenceException("Au moins un des points est null.")
        '    End If

        '    Return New Line(points(0), points(1))
        'End Operator
#End Region

    End Class
