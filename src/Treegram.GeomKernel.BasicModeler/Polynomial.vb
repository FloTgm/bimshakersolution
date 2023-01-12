Public Class Polynomial

    Private Property MCoefficients As List(Of Double)

    Public Property Coefficients As List(Of Double)
        Get
            Return MCoefficients
        End Get
        Set(value As List(Of Double))
            Dim newCoeffs As List(Of Double) = Nothing
            For Each coeff In value
                newCoeffs.Add(RoundToSignificantDigits(coeff, RealPrecision))
            Next
            MCoefficients = newCoeffs
        End Set
    End Property

    Public ReadOnly Property Degree As Integer
        Get
            Me.Simplify()
            Return Me.MCoefficients.Count - 1
        End Get
    End Property

    ''' <summary>
    ''' New Polynomial.
    ''' </summary>
    ''' <param name="coefficients"> Coefficients of Polynome in ascending order. </param>
    Public Sub New(coefficients As List(Of Double))
        If coefficients Is Nothing OrElse coefficients.Count = 0 Then
            Me.MCoefficients = {0.0}.ToList
        Else
            Me.Coefficients = coefficients
        End If
        Simplify()
    End Sub

    ''' <summary>
    ''' Null Polynomial.
    ''' </summary>
    Public Sub New()
        Me.MCoefficients = {0.0}.ToList
    End Sub

    ''' <summary>
    ''' New Polynomial from string expression.
    ''' </summary>
    ''' <param name="str"></param>
    Public Sub New(str As String)
        Dim coeffs As New List(Of Double)
        Dim lowerStr = str.ToLower
        Dim split = lowerStr.Split({"+"}, StringSplitOptions.None)
        Dim dico As New Dictionary(Of Integer, Double)

        For Each elem In split
            Dim secondSplit = elem.Split({"x"}, StringSplitOptions.None)
            Dim key As Integer
            Dim value As Double
            If secondSplit.Count > 1 Then
                Dim puissance = elem.Split({"x"}, StringSplitOptions.None)(1)
                If puissance = "" Then
                    key = 1
                Else
                    key = CInt(puissance)
                End If
            Else
                key = 0
            End If
            Dim coeff = secondSplit(0)
            If coeff = "" Then
                value = 1
            Else
                value = CDbl(coeff)
            End If
            dico.Add(key, value)
        Next
        dico = (From entry In dico Order By entry.Key Ascending).ToDictionary(Function(pair) pair.Key, Function(pair) pair.Value)

        Dim i = 0
        While i <= dico.Last.Key
            If dico.ContainsKey(i) Then
                coeffs.Add(dico(i))
            Else
                coeffs.Add(0)
            End If
            i += 1
        End While

        Me.MCoefficients = coeffs
        Simplify()
    End Sub

    ''' <summary>
    ''' Value at parameter
    ''' </summary>
    ''' <param name="t"> Parameter </param>
    ''' <returns></returns>
    Public Function Value(t As Double) As Double
        Dim res = 0.0
        If MCoefficients.Count = 0 Then
            Return 0
        End If
        For i = 0 To MCoefficients.Count - 1
            res += MCoefficients(i) * Math.Pow(t, i)
        Next
        Return RoundToSignificantDigits(res, RealPrecision)
    End Function

    ''' <summary>
    ''' Derived Polynomial.
    ''' </summary>
    ''' <returns></returns>
    Public Function Derivee() As Polynomial
        Dim newCoeffs As New List(Of Double)
        If Me.MCoefficients.Count < 2 Then
            Return New Polynomial()
        Else
            For i = 1 To Me.MCoefficients.Count - 1
                newCoeffs.Add(i * Me.MCoefficients(i))
            Next
            Return New Polynomial(newCoeffs)
        End If
    End Function

    ''' <summary>
    ''' Primitive of the Polynomial, that vanishes at zeroValue.
    ''' </summary>
    ''' <param name="zeroValue"> Value of vanishing </param>
    ''' <returns></returns>
    Public Function Primitive(Optional zeroValue As Double = 0) As Polynomial
        Dim newCoeffs As New List(Of Double) From {zeroValue}
        For i = 0 To Me.Coefficients.Count - 1
            newCoeffs.Add(1 / (i + 1) * Me.MCoefficients(i))
        Next
        Return New Polynomial(newCoeffs)
    End Function

    ''' <summary>
    ''' Integral between two values.
    ''' </summary>
    ''' <param name="startValue"></param>
    ''' <param name="endValue"></param>
    ''' <returns></returns>
    Public Function Integrale(startValue As Double, endValue As Double) As Double
        Return Me.Primitive.Value(endValue) - Me.Primitive.Value(startValue)
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="polynomial1"></param>
    ''' <param name="polynomial2"></param>
    ''' <returns></returns>
    Public Shared Operator +(polynomial1 As Polynomial, polynomial2 As Polynomial) As Polynomial
        Dim long1 = polynomial1.MCoefficients.Count
        Dim long2 = polynomial2.MCoefficients.Count
        If long1 = 0 Then
            Return polynomial2
        End If
        If long2 = 0 Then
            Return polynomial1
        End If
        Dim newCoeffs As New List(Of Double)
        For i = 0 To Math.Max(long1, long2) - 1
            If i < long1 Then
                If i < long2 Then
                    newCoeffs.Add(polynomial1.MCoefficients(i) + polynomial2.MCoefficients(i))
                Else
                    newCoeffs.Add(polynomial1.MCoefficients(i))
                End If
            Else
                newCoeffs.Add(polynomial2.MCoefficients(i))
            End If
        Next
        Return New Polynomial(newCoeffs)
    End Operator

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="polynomial1"></param>
    ''' <param name="polynomial2"></param>
    ''' <returns></returns>
    Public Shared Operator -(polynomial1 As Polynomial, polynomial2 As Polynomial) As Polynomial
        Dim long1 = polynomial1.MCoefficients.Count
        Dim long2 = polynomial2.MCoefficients.Count
        If long1 = 0 Then
            Return polynomial2
        End If
        If long2 = 0 Then
            Return polynomial1
        End If
        Dim newCoeffs As New List(Of Double)
        For i = 0 To Math.Max(long1, long2) - 1
            If i < long1 Then
                If i < long2 Then
                    newCoeffs.Add(polynomial1.MCoefficients(i) - polynomial2.MCoefficients(i))
                Else
                    newCoeffs.Add(polynomial1.MCoefficients(i))
                End If
            Else
                newCoeffs.Add(-polynomial2.MCoefficients(i))
            End If
        Next
        Return New Polynomial(newCoeffs)
    End Operator

    Public Shared Operator *(polynomial1 As Polynomial, polynomial2 As Polynomial) As Polynomial
        Dim dico As New Dictionary(Of Integer, Double)
        For i = 0 To polynomial1.Coefficients.Count - 1
            For j = 0 To polynomial2.Coefficients.Count - 1
                If dico.ContainsKey(i * j) Then
                    dico(i * j) += polynomial1.MCoefficients(i) * polynomial2.MCoefficients(j)
                Else
                    dico.Add(i * j, polynomial1.MCoefficients(i) * polynomial2.MCoefficients(j))
                End If
            Next
        Next

        dico = (From entry In dico Order By entry.Key Ascending).ToDictionary(Function(pair) pair.Key, Function(pair) pair.Value)

        Dim coeffs As New List(Of Double)
        Dim k = 0
        While k <= dico.Last.Key
            If dico.ContainsKey(k) Then
                coeffs.Add(dico(k))
            Else
                coeffs.Add(0)
            End If
            k += 1
        End While

        Return New Polynomial(coeffs)
    End Operator

    ''' <summary>
    ''' Euclidian division with quotient and remainder
    ''' </summary>
    ''' <param name="dividend"></param>
    ''' <param name="divisor"></param>
    ''' <returns> (Quotient, Remainder) </returns>
    Public Shared Operator /(dividend As Polynomial, divisor As Polynomial) As Tuple(Of Polynomial, Polynomial)
        divisor.Simplify()
        If divisor.MCoefficients.Count = 0 Then
            Throw New DivideByZeroException
        End If
        dividend.Simplify()
        If dividend.MCoefficients.Count < divisor.MCoefficients.Count Then
            Return New Tuple(Of Polynomial, Polynomial)(New Polynomial(), dividend)
        End If

        Dim quotient, reste As New List(Of Double)
        Dim degDivisor = divisor.Degree
        Dim degDividend = dividend.Degree
        Dim dicoSum As New Dictionary(Of Integer, List(Of Double))
        Dim dicoQuotRem As New Dictionary(Of Integer, Double)

        For i = 0 To degDividend
            dicoSum.Add(i, {dividend.MCoefficients(i)}.ToList)
        Next
        For j = degDividend To 0 Step -1
            Dim y = 0.0
            For Each x In dicoSum(j)
                y += x
            Next
            If j >= degDivisor Then
                dicoQuotRem.Add(j, y / divisor.MCoefficients.Last)
                For k = 0 To degDivisor - 1
                    dicoSum(j - degDivisor + k).Add(-dicoQuotRem(j) * divisor.MCoefficients(k))
                Next
            Else
                dicoQuotRem.Add(j, y)
            End If
        Next

        dicoQuotRem = (From entry In dicoQuotRem Order By entry.Key Ascending).ToDictionary(Function(pair) pair.Key, Function(pair) pair.Value)
        For n = 0 To degDividend
            If n < degDivisor Then
                reste.Add(dicoQuotRem(n))
            Else
                quotient.Add(dicoQuotRem(n))
            End If
        Next

        Return New Tuple(Of Polynomial, Polynomial)(New Polynomial(quotient), New Polynomial(reste))
    End Operator

    Public Shared Operator =(polynomial1 As Polynomial, polynomial2 As Polynomial) As Boolean
        polynomial1.Simplify()
        polynomial2.Simplify()
        If polynomial1.MCoefficients.Count <> polynomial2.MCoefficients.Count Then
            Return False
        End If
        For i = 0 To polynomial1.MCoefficients.Count - 1
            If polynomial1.MCoefficients(i) <> polynomial2.MCoefficients(i) Then
                Return False
            End If
        Next
        Return True
    End Operator

    Public Shared Operator <>(polynomial1 As Polynomial, polynomial2 As Polynomial) As Boolean
        Return Not (polynomial1 = polynomial2)
    End Operator

    ''' <summary>
    ''' Simplify coeeficients
    ''' </summary>
    Private Sub Simplify()
        Dim isNull = (Me.MCoefficients.Last = 0.0)
        While isNull
            Me.MCoefficients.RemoveAt(Me.MCoefficients.Count - 1)
            If Me.Coefficients.Count <> 0 Then
                isNull = (Me.MCoefficients.Last = 0.0)
            Else
                Me.MCoefficients = {0.0}.ToList
                isNull = False
            End If
        End While
    End Sub

    Public Overrides Function ToString() As String
        Dim str As String = ""
        For k = MCoefficients.Count - 1 To 0 Step -1
            If k = 0 Then
                str += MCoefficients(k).ToString + " + "
            ElseIf k = 1 Then
                str += MCoefficients(k).ToString + "X" + " + "
            Else

                str += MCoefficients(k).ToString + "X^" + k.ToString + " + "
            End If
        Next
        str = str.Substring(0, str.Count - 3)
        Return "Function polynomiale : " + str
    End Function
End Class
