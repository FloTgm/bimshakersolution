Imports M4D.Treegram.Core.Entities
Imports M4D.Treegram.Core.Extensions.Entities

Public Class IfcFileObject
    Inherits FileObject
    Public Sub New(fileTgm As MetaObject)
        MyBase.New(fileTgm)
    End Sub

#Region "Specific IFC Objects"

    Public ReadOnly Property IfcTypeTgm As MetaObject
        Get
            Return FileTgm.GetChildren(, "IfcType").First
        End Get
    End Property
    Public ReadOnly Property IfcTypes As List(Of MetaObject)
        Get
            Return IfcTypeTgm.GetChildren.ToList
        End Get
    End Property

    Public ReadOnly Property SiteTgm As MetaObject
        Get
            Return _3dObjTgm.GetChildren("IfcSite").FirstOrDefault
        End Get
    End Property
    Public ReadOnly Property SiteBuildings As List(Of MetaObject)
        Get
            'Return siteTgm.GetChildren("IfcBuilding").ToList 'intéressant mais bloquant
            Return ScanWs.GetMetaObjects(, "IfcBuilding").ToList
        End Get
    End Property
    Public ReadOnly Property BuildingStoreys(buildingTgm As MetaObject) As List(Of MetaObject)
        Get
            Return buildingTgm.GetChildren("IfcBuildingStorey").ToList
        End Get
    End Property

    Private _getBuildings As New Dictionary(Of MetaObject, List(Of MetaObject))
    Public ReadOnly Property GetBuildings(buildingsList As List(Of MetaObject), objsList As List(Of MetaObject)) As Dictionary(Of MetaObject, List(Of MetaObject))
        Get
            'prepare dico
            _getBuildings = New Dictionary(Of MetaObject, List(Of MetaObject))
            For Each buildTgm In buildingsList
                _getBuildings.Add(buildTgm, New List(Of MetaObject))
            Next
            'fill dico
            For Each objTgm In objsList
                Dim buildingTgm = objTgm.GetParents("IfcBuilding",,, True).FirstOrDefault
                If buildingTgm Is Nothing Then
                    Continue For
                    'If _GetBuildings.ContainsKey(Nothing) Then
                    '    _GetBuildings(Nothing).Add(objTgm)
                    'Else
                    '    _GetBuildings.Add(Nothing, New List(Of MetaObject) From {objTgm})
                    'End If
                Else
                    _getBuildings(buildingTgm).Add(objTgm)
                End If
            Next
            Return _getBuildings
        End Get
    End Property


    Private _getStoreys As New Dictionary(Of MetaObject, List(Of MetaObject))
    Public ReadOnly Property GetStoreys(storeysList As List(Of MetaObject), objsList As List(Of MetaObject), Optional recursively As Boolean = False) As Dictionary(Of MetaObject, List(Of MetaObject))
        Get
            'prepare dico
            _getStoreys = New Dictionary(Of MetaObject, List(Of MetaObject))
            For Each storeyTgm In storeysList
                _getStoreys.Add(storeyTgm, New List(Of MetaObject))
            Next
            'fill dico
            For Each objTgm In objsList
                Dim storeyTgm = objTgm.GetParents("IfcBuildingStorey",,, recursively).FirstOrDefault
                If storeyTgm Is Nothing Then
                    Continue For
                    'If _GetStoreys.ContainsKey(Nothing) Then
                    '    _GetStoreys(Nothing).Add(objTgm)
                    'Else
                    '    _GetStoreys.Add(Nothing, New List(Of MetaObject) From {objTgm})
                    'End If
                Else
                    _getStoreys(storeyTgm).Add(objTgm)
                End If
            Next
            Return _getStoreys
        End Get
    End Property

    Private _getLayers As New SortedDictionary(Of String, List(Of MetaObject))
    Public ReadOnly Property GetLayers(objsList As List(Of MetaObject)) As SortedDictionary(Of String, List(Of MetaObject))
        Get
            'If _GetLayers Is Nothing Then '<--- SURTOUT PAS !!!
            _getLayers = New SortedDictionary(Of String, List(Of MetaObject))
            For Each objTgm In objsList
                Dim layerSt = objTgm.GetAttribute("Layer")?.Value
                If layerSt Is Nothing OrElse layerSt = "" Then
                    layerSt = "<NONE>"
                End If
                If _getLayers.ContainsKey(layerSt) Then
                    _getLayers(layerSt).Add(objTgm)
                Else
                    _getLayers.Add(layerSt, New List(Of MetaObject) From {objTgm})
                End If
            Next
            'End If
            Return _getLayers
        End Get
    End Property

    Private _getTypes As New SortedDictionary(Of String, List(Of MetaObject))
    Public ReadOnly Property GetTypes(objsList As List(Of MetaObject)) As SortedDictionary(Of String, List(Of MetaObject))
        Get
            'If _GetTypes Is Nothing Then '<--- SURTOUT PAS !!!
            _getTypes = New SortedDictionary(Of String, List(Of MetaObject))
            For Each objTgm In objsList
                Dim typeSt = objTgm.GetTgmType
                If typeSt Is Nothing OrElse typeSt = "" Then
                    typeSt = "<NONE>"
                End If
                If _getTypes.ContainsKey(typeSt) Then
                    _getTypes(typeSt).Add(objTgm)
                Else
                    _getTypes.Add(typeSt, New List(Of MetaObject) From {objTgm})
                End If
            Next
            'End If
            Return _getTypes
        End Get
    End Property

    Public ReadOnly Property GetNames(objsList As List(Of MetaObject)) As SortedDictionary(Of String, List(Of MetaObject))
        Get
            Dim namesDico As New SortedDictionary(Of String, List(Of MetaObject))
            For Each objTgm In objsList
                Dim nameSt = objTgm.Name
                If nameSt Is Nothing OrElse nameSt = "" Then
                    nameSt = "<NONE>"
                End If
                If namesDico.ContainsKey(nameSt) Then
                    namesDico(nameSt).Add(objTgm)
                Else
                    namesDico.Add(nameSt, New List(Of MetaObject) From {objTgm})
                End If
            Next
            Return namesDico
        End Get
    End Property
#End Region

End Class
