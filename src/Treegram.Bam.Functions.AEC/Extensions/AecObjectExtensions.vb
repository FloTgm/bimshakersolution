Imports System.Runtime.CompilerServices
Imports SharpDX
Imports Treegram.GeomLibrary
Imports Treegram.Bam.Libraries.AEC
Imports M4D.Treegram.Core.Extensions.Entities
Imports M4D.Treegram.Core.Extensions.Kernel
Imports SharpDX.Direct3D11

Public Module AecObjectExtensions
    'Extensions using BasicModeler

    ''' <summary>Get 3d object models</summary>
    <Extension()>
    Public Function Models(aecObj As AecObject, Optional includeDecomposition As Boolean = True) As List(Of M4D.Deixi.Core.Model)
        Return Treegram.GeomFunctions.Models.GetGeometryModels(aecObj.Metaobject, includeDecomposition)
    End Function

    ''' <summary>Get 3d object gross models</summary>
    <Extension()>
    Public Function GrossModels(aecObj As AecObject) As List(Of M4D.Deixi.Core.Model)
        Dim grossExt = aecObj.GrossExtension
        If grossExt Is Nothing Then
            Return Nothing
        End If
        Return Treegram.GeomFunctions.Models.GetGeometryModels(grossExt, True)
    End Function

    <Extension()>
    Public Function SmartAddExtremumZ(aecObj As AecObject) As Tuple(Of Double, Double)
        'MIN Z
        Dim minZ = aecObj.MinZ
        If minZ = Nothing Then
            minZ = Double.PositiveInfinity
            For Each model In aecObj.Models
                For Each vert In model.VerticesTesselation
                    If vert.Z < minZ Then minZ = vert.Z
                Next
            Next
            If minZ <> Double.PositiveInfinity Then
                aecObj.CompleteTgmPset("MinZ", Math.Round(minZ, 5))
            Else
                Return New Tuple(Of Double, Double)(Nothing, Nothing)
            End If
        End If
        'MAX Z
        Dim maxZ = aecObj.MaxZ
        If maxZ = Nothing Then
            maxZ = Double.NegativeInfinity
            For Each model In aecObj.Models
                For Each vert In model.VerticesTesselation
                    If vert.Z > maxZ Then maxZ = vert.Z
                Next
            Next
            If maxZ <> Double.NegativeInfinity Then
                aecObj.CompleteTgmPset("MaxZ", Math.Round(maxZ, 5))
            Else
                Return New Tuple(Of Double, Double)(Nothing, Nothing)
            End If
        End If
        Return New Tuple(Of Double, Double)(minZ, maxZ)
    End Function


    'CENTER POINTS
    ''' <summary>Get the barycenter. Caution : this property uses object models</summary>
    <Extension()>
    Public Function SmartAddBarycenter(aecObj As AecObject) As Point3D

        If aecObj.Barycenter Is Nothing Then
            'OPTION 1
            'Dim barycenter = Treegram.GeomFunctions.Models.GetMetaObjectVerticesBarycenter(aecObj.Metaobject, True)

            'OPTION 2 : better for doors and windows
            Dim ptList As New List(Of Vector3)
            For Each model In aecObj.Models
                ptList.Add(Treegram.GeomFunctions.Models.GetModelVerticesBarycenter(model))
            Next
            If ptList.Count = 0 Then Return Nothing
            Dim barycenter = CommonGeomFunctions.GetBarycenterFromPoints(ptList)

            If Not barycenter.IsZero Then
                Dim baryAtt = aecObj.AdditionalGeometryExtension.SmartAddAttribute("BarycenterByTgm", Nothing)
                Dim baryTgmPt As New Point3D(barycenter)
                baryTgmPt.Write(baryAtt)
                Return baryTgmPt
            Else
                Return Nothing
            End If
        Else
            Return aecObj.Barycenter
        End If
    End Function

    ''' <summary>Get max U coord in object local axis (from LocalPlacement). Caution : this property uses object models</summary>
    <Extension()>
    Public Function MaxU(aecObj As AecObject) As Double
        Return CommonGeomFunctions.XYmodelsMinMax(aecObj.Models, aecObj.VectorU).Item2
    End Function

    ''' <summary>Get min U coord in object local axis (from LocalPlacement). Caution : this property uses object models</summary>
    <Extension()>
    Public Function MinU(aecObj As AecObject) As Double
        Return CommonGeomFunctions.XYmodelsMinMax(aecObj.Models, aecObj.VectorU).Item1
    End Function

    ''' <summary>Get max V coord in object local axis (from LocalPlacement). Caution : this property uses object models</summary>
    <Extension()>
    Public Function MaxV(aecObj As AecObject) As Double
        Return CommonGeomFunctions.XYmodelsMinMax(aecObj.Models, aecObj.VectorV).Item2
    End Function

    ''' <summary>Get min V coord in object local axis (from LocalPlacement). Caution : this property uses object models</summary>
    <Extension()>
    Public Function MinV(aecObj As AecObject) As Double
        Return CommonGeomFunctions.XYmodelsMinMax(aecObj.Models, aecObj.VectorV).Item1
    End Function

End Module

