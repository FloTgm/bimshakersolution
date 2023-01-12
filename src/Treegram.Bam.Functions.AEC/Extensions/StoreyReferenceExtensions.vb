
Imports System.Runtime.CompilerServices
Imports M4D.Treegram.Core.Entities
Imports Treegram.ConstructionManagement

Public Module Extensions
    <Extension()>
    Public Function LocalElevation(storeyRef As StoreyReference, scanWs As Workspace) As Double
        Dim globalPoint = Treegram.GeomFunctions.GeoReferencing.TransformPointInGlobalAxes(New Treegram.GeomKernel.BasicModeler.Point(0, 0, 0), scanWs)
        Return Math.Round(storeyRef.GlobalElevation - globalPoint.Z, 4)
    End Function
End Module

