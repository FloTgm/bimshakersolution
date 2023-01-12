
Imports System.Runtime.CompilerServices
Imports M4D.Treegram.Core.Entities
Imports M4D.Treegram.Core.Extensions.Entities
Imports Treegram.ConstructionManagement

Public Module StoreyReferenceExtensions
    <Extension()>
    Public Function AecObjects(storeyRef As StoreyReference, scanWs As Workspace) As List(Of AecObject)
        Dim tgmObjs = storeyRef.Metaobject.GetChildren.Where(Function(o) CType(o.Container, Workspace).Equals(scanWs)).ToList
        Return tgmObjs.Select(Function(o) New AecObject(o)).ToList
    End Function

End Module

