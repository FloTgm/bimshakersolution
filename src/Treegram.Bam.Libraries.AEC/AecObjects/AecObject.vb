Imports M4D.Treegram.Core.Entities
Imports M4D.Treegram.Core.Kernel
Imports M4D.Treegram.Core.Extensions.Entities
Imports M4D.Treegram.Core.Extensions.Kernel
Imports Treegram.ConstructionManagement
Imports Treegram.GeomLibrary
Imports SharpDX

Public Class AecObject

    Private _anaExtTgm, _psetExtTgm, _profileExtTgm, _horizontalProfileExtTgm, _axisExtTgm, _additionalGeometryExtension As MetaObject
    Private _height As Double?
    Dim _envAnaRes As EnvelopePosition?
    Dim _envOwner As String

    Public Sub New(Optional objTgm As MetaObject = Nothing)
        If objTgm IsNot Nothing Then
            Me.Metaobject = objTgm
        End If
    End Sub

#Region "INFOS"
    Public Property Metaobject As MetaObject
    Public ReadOnly Property ContainerWs As Workspace
        Get
            Return CType(Me.Metaobject.Container, Workspace)
        End Get
    End Property
    Public ReadOnly Property Name As String
        Get
            Return Metaobject.Name
        End Get
    End Property
    Public ReadOnly Property Type As String
        Get
            Return Me.Metaobject.GetTgmType
        End Get
    End Property
    Private _category As Category
    Public ReadOnly Property Category As Category
        Get
            If _category = Nothing Then
                Dim categAtt = Me.Metaobject.GetAttribute("CategoryByTgm", True)
                If categAtt IsNot Nothing AndAlso System.Enum.GetNames(GetType(Category)).Contains(categAtt.Value.ToString) Then
                    _category = System.[Enum].Parse(GetType(Category), categAtt.Value.ToString)
                Else
                    _category = Category.Unknown
                End If
            End If
            Return _category
        End Get
    End Property
    Public ReadOnly Property LodByTgm As Integer?
        Get
            If Me.Metaobject.GetAttribute("LodByTgm") IsNot Nothing Then
                Return CInt(Me.Metaobject.GetAttribute("LodByTgm", True).Value)
            Else
                Return 0
            End If
        End Get
    End Property

    Public ReadOnly Property AnalysisExtension As MetaObject
        Get
            If _anaExtTgm Is Nothing Then
                If Me.ContainerWs.Name.ToLower.Contains("tgmcreation") Then 'Pas super comme condition... à voir si on veut le même comportement que pour les extensions PsetTgm : càd créées en amont
                    Throw New Exception("No analysis extension for tgm objects")
                End If
                '_AnaExtTgm = Me.Metaobject.SmartAddExtension( M4D.Treegram.Core.Extensions.Enums.ExtensionType.Analysis) 'After several revisions, this method create to many objects...

                Dim extWsName = Me.ContainerWs.Name + " - " + M4D.Treegram.Core.Extensions.Enums.ExtensionType.Analysis.ToString
                Dim extWs = Me.Metaobject.GetProject.GetWorkspaces(extWsName).FirstOrDefault
                If extWs Is Nothing Then
                    Throw New Exception("Analysis workspace is missing for this file : " + Me.ContainerWs.Name)
                End If

                Dim extObjName = Me.Name + " - " + M4D.Treegram.Core.Extensions.Enums.ExtensionType.Analysis.ToString
                'Dim extObj = baseObj.Extensions.FirstOrDefault(Function(o) o.Name = extensionName) 'PAS SAFE SI WS PAS CHARGE...
                _anaExtTgm = extWs.MetaObjects.FirstOrDefault(Function(o) o.Name = extObjName AndAlso o.Extend IsNot Nothing AndAlso o.Extend.Equals(Me.Metaobject))
                If _anaExtTgm IsNot Nothing Then
                    _anaExtTgm.IsActive = True
                ElseIf extWs.MetaObjects.FirstOrDefault(Function(o) Not o.IsActive AndAlso o.Name = extObjName) IsNot Nothing Then 'It is supposed to be clean : no attribute, no relation, etc.
                    _anaExtTgm = extWs.MetaObjects.FirstOrDefault(Function(o) Not o.IsActive AndAlso o.Name = extObjName)
                    _anaExtTgm.IsActive = True
                    _anaExtTgm.Extend = Me.Metaobject
                Else
                    _anaExtTgm = extWs.AddMetaObject(extObjName)
                    _anaExtTgm.Extend = Me.Metaobject
                End If

            End If
            Return _anaExtTgm
        End Get
    End Property
    Public ReadOnly Property PsetTgmExtension As MetaObject
        Get
            If _psetExtTgm Is Nothing Then
                _psetExtTgm = Me.Metaobject.Extensions.FirstOrDefault(Function(o) o.Name.Substring(Math.Max(0, o.Name.Length - 15)) = $" - {ProjectReference.Constants.psetTgmWsType}")
            End If
            Return _psetExtTgm
        End Get
    End Property
    Public ReadOnly Property AxisExtension As MetaObject
        Get
            If _axisExtTgm Is Nothing Then
                _axisExtTgm = Me.Metaobject.Extensions.FirstOrDefault(Function(o) o.Name.Substring(Math.Max(0, o.Name.Length - 7)) = " - Axis")
            End If
            Return _axisExtTgm
        End Get
    End Property
    Public ReadOnly Property ProfileExtension As MetaObject
        Get
            If _profileExtTgm Is Nothing Then
                _profileExtTgm = Me.Metaobject.Extensions.FirstOrDefault(Function(o) o.Name.Substring(Math.Max(0, o.Name.Length - 10)) = " - Profile")
            End If
            Return _profileExtTgm
        End Get
    End Property
    Public ReadOnly Property HorizontalProfileExtension As MetaObject
        Get
            If _horizontalProfileExtTgm Is Nothing Then
                _horizontalProfileExtTgm = Me.Metaobject.Extensions.FirstOrDefault(Function(o) o.Name.Substring(Math.Max(0, o.Name.Length - 20)) = " - HorizontalProfile")
            End If
            Return _horizontalProfileExtTgm
        End Get
    End Property
    Public ReadOnly Property AdditionalGeometryExtension As MetaObject
        Get
            If _additionalGeometryExtension Is Nothing Then
                _additionalGeometryExtension = Me.Metaobject.Extensions.FirstOrDefault(Function(o) o.Name.Substring(Math.Max(0, o.Name.Length - 23)) = " - AdditionalGeometries")
            End If
            Return _additionalGeometryExtension
        End Get
    End Property
    Public ReadOnly Property GrossExtension As MetaObject
        Get
            Dim grossExt = Me.Metaobject.Extensions.Where(Function(m) m.Name.Contains("BrutGeometry") Or m.Name.Contains("GrossGeometry")).FirstOrDefault
            Return grossExt
        End Get
    End Property
    ''' <summary>
    ''' Simply get the extension if there is one, otherwise get the metaobject
    ''' </summary>
    Public ReadOnly Property GetWritableTgmObj As MetaObject
        Get
            'If Me.Metaobject.Extensions.Count > 0 Then
            If Me.ContainerWs.Name.ToLower.Contains("tgmcreation") Then
                Return Me.Metaobject
            Else
                Return Me.AnalysisExtension
            End If
        End Get
    End Property

    Public ReadOnly Property TgmBuildingStoreysAtt As String
        Get
            Return Me.Metaobject.GetAttribute("TgmBuildingStorey", True).Value.ToString
        End Get
    End Property
    Public ReadOnly Property EntranceTgmStorey(myBuildingRef As BuildingReference) As StoreyReference
        Get
            Dim entranceStoreySt = Me.Metaobject.GetAttribute("StoreyByTgm").Value
            Return myBuildingRef.SortedStoreys.FirstOrDefault(Function(o) o.Name = entranceStoreySt)
        End Get
    End Property

    Public ReadOnly Property TgmBuildingStoreys(myBuildingRef As BuildingReference) As List(Of StoreyReference)
        Get
            Dim storeys = Me.Metaobject.GetParents(Nothing, Nothing, RelationType.Containance).ToList.Where(Function(o) o.GetTgmType = "TgmBuildingStorey").ToList
            Dim stoRefList As New List(Of StoreyReference)
            For Each sto In myBuildingRef.SortedStoreys
                If storeys.FirstOrDefault(Function(o) o.Name = sto.Name) IsNot Nothing Then
                    stoRefList.Add(sto)
                End If
            Next
            Return stoRefList
        End Get
    End Property

    Public ReadOnly Property IfcBuilding As MetaObject
        Get
            Return Me.Metaobject.GetParents("IfcBuilding",,, True).FirstOrDefault
        End Get
    End Property

    Public ReadOnly Property IfcBuildingStorey As MetaObject
        Get
            Return Me.Metaobject.GetParents("IfcBuildingStorey",,, True).FirstOrDefault
        End Get
    End Property
    Public Property IsDone As Boolean

    Public ReadOnly Property IfcProfile As Profile3D
        Get
            If ProfileExtension IsNot Nothing Then
                'Get attribute profile
                Dim profileAtt = ProfileExtension.GetAttribute(IfcProfileType.RectangleProfile.ToString, True)
                If profileAtt Is Nothing Then
                    profileAtt = ProfileExtension.GetAttribute(IfcProfileType.ArbitraryProfile.ToString, True)
                    If profileAtt Is Nothing Then
                        Return Nothing
                    End If
                End If
                'Read points
                Dim ptList As New List(Of Vector3)
                For i = 0 To profileAtt.Attributes.Count - 2 'because currently Point 5 = Point 1
                    Dim myPtAtt = profileAtt.Attributes(i)
                    If myPtAtt.GetAttribute("Z") IsNot Nothing Then
                        ptList.Add(New Vector3(CDbl(myPtAtt.GetAttribute("X").Value), CDbl(myPtAtt.GetAttribute("Y").Value), CDbl(myPtAtt.GetAttribute("Z").Value)))
                    Else
                        ptList.Add(New Vector3(CDbl(myPtAtt.GetAttribute("X").Value), CDbl(myPtAtt.GetAttribute("Y").Value), 0.0))
                    End If
                Next
                Return New Profile3D(ptList)
            Else
                Return Nothing
            End If

        End Get
    End Property

#End Region

#Region "GEOMETRY"

    Private _hasGeometry
    Public ReadOnly Property HasGeometry As Boolean?
        Get
            If _hasGeometry Is Nothing Then
                _hasGeometry = Me.Metaobject.GetAttribute("HasGeometry", True)?.Value
            End If
            Return _hasGeometry
        End Get
    End Property

    Public ReadOnly Property IsSpace As Boolean
        Get
            Return Me.Metaobject.GetAttribute("IsSpace", True)?.Value
        End Get
    End Property
    Public ReadOnly Property IsSlab As Boolean
        Get
            Return Me.Metaobject.GetAttribute("IsSlab", True)?.Value
        End Get
    End Property
    Public ReadOnly Property IsOpenable As Boolean
        Get
            Return Me.Metaobject.GetAttribute("IsOpenable", True)?.Value
        End Get
    End Property
    Public ReadOnly Property IsVerticalSeparator As Boolean
        Get
            Return Me.Metaobject.GetAttribute("IsVerticalSeparator", True)?.Value
        End Get
    End Property
    'Public ReadOnly Property IsHorizontalSeparator As Boolean
    '    Get
    '        Return Me.Metaobject.GetAttribute("IsHorizontalSeparator", True)?.Value
    '    End Get
    'End Property
    Public ReadOnly Property IsGlazed As Boolean
        Get
            Return Me.Metaobject.GetAttribute("IsGlazed", True)?.Value
        End Get
    End Property
    Public ReadOnly Property IsReservation As Boolean
        Get
            Return Me.Metaobject.GetAttribute("IsReservation", True)?.Value
        End Get
    End Property

    Private _thicknessByTgm As Double?
    Public ReadOnly Property ThicknessByTgm As Double 'Pas à la bonne place mais utile pour les séparateurs verticaux
        Get
            If _thicknessByTgm Is Nothing Then
                '_ThicknessByTgm = CDbl(Me.PsetTgmExtension.GetAttribute("ThicknessByTgm", True)?.Value) 'NE FONCTIONNE PAS avec les éléments superposés
                _thicknessByTgm = CDbl(Me.Metaobject.GetAttribute("ThicknessByTgm", True)?.Value)
            End If
            Return _thicknessByTgm
        End Get
    End Property
    Public ReadOnly Property WidthByTgm As Double
        Get
            Return Me.Metaobject.GetAttribute("WidthByTgm", True).Value
        End Get
    End Property

    Private _barycenter As Point3D
    Public ReadOnly Property Barycenter As Point3D
        Get
            If _barycenter Is Nothing Then
                If Me.AdditionalGeometryExtension IsNot Nothing AndAlso Me.AdditionalGeometryExtension.GetAttribute("BarycenterByTgm") IsNot Nothing Then
                    Dim baryAtt = Me.AdditionalGeometryExtension.GetAttribute("BarycenterByTgm")
                    Dim bary3dPt As New Point3D()
                    If bary3dPt.Read(baryAtt) Then
                        _barycenter = bary3dPt
                    End If
                End If
            End If
            Return _barycenter
        End Get
    End Property
    Public ReadOnly Property Barycenter2D As Point2D
        Get
            If _barycenter IsNot Nothing Then
                Return New Point2D(_barycenter.Point.X, _barycenter.Point.Y)
            Else
                Return Nothing
            End If
        End Get
    End Property

    Private _vectorU As Vector2
    ''' <summary>Get U vector in local axis (from LocalPlacement).</summary>
    Public Property VectorU As Vector2
        Get
            If _vectorU.IsZero Then
                If Me.AdditionalGeometryExtension IsNot Nothing AndAlso Me.AdditionalGeometryExtension.GetAttribute("AxisSystemByTgm") IsNot Nothing Then
                    _vectorU = New Vector2(CDbl(Me.AdditionalGeometryExtension.SmartGetAttribute("AxisSystemByTgm/_\Axis U/_\X").Value), CDbl(Me.AdditionalGeometryExtension.SmartGetAttribute("AxisSystemByTgm/_\Axis U/_\Y").Value))
                Else
                    'Throw New Exception("No local placement")
                    ''Return Nothing
                End If
            End If
            Return _vectorU
        End Get
        Set(value As Vector2)
            _vectorU = value
        End Set
    End Property

    Private _vectorV As Vector2
    ''' <summary>Get V vector in local axis (from LocalPlacement).</summary>
    Public Property VectorV As Vector2
        Get
            If _vectorV.IsZero Then
                If Me.AdditionalGeometryExtension IsNot Nothing AndAlso Me.AdditionalGeometryExtension.GetAttribute("AxisSystemByTgm") IsNot Nothing Then
                    _vectorV = New Vector2(CDbl(Me.AdditionalGeometryExtension.SmartGetAttribute("AxisSystemByTgm/_\Axis V/_\X").Value), CDbl(Me.AdditionalGeometryExtension.SmartGetAttribute("AxisSystemByTgm/_\Axis V/_\Y").Value))
                Else
                    'Throw New Exception("No local placement")
                    ''Return Nothing
                End If
            End If
            Return _vectorV
        End Get
        Set(value As Vector2)
            _vectorV = value
        End Set
    End Property
    Private _axisSystem2d As AxisSystem
    Public ReadOnly Property AxisSystem As AxisSystem
        Get
            If Me.AdditionalGeometryExtension IsNot Nothing AndAlso _axisSystem2d Is Nothing Then
                Dim axisSyst = New AxisSystem
                If axisSyst.Read(Me.AdditionalGeometryExtension) Then
                    _axisSystem2d = axisSyst
                End If
            End If
            Return _axisSystem2d
        End Get
    End Property
    Private _horizProfile2d As Profile2D
    Public ReadOnly Property HorizontalProfile As Profile2D
        Get
            If Me.HorizontalProfileExtension IsNot Nothing AndAlso _horizProfile2d Is Nothing Then
                Dim horiz2dProfile = New Profile2D
                If horiz2dProfile.Read(Me.HorizontalProfileExtension) Then
                    _horizProfile2d = horiz2dProfile
                End If
            End If
            Return _horizProfile2d
        End Get
    End Property

    Private _horizProfile As Profile3D
    Public ReadOnly Property HorizontalProfile(z As Double) As Profile3D
        Get
            If Me.HorizontalProfileExtension IsNot Nothing AndAlso (_horizProfile Is Nothing OrElse _horizProfile.Points.First.Z <> z) Then
                Dim horiz2dProfile = New Profile2D
                If horiz2dProfile.Read(Me.HorizontalProfileExtension) Then
                    Dim point3dList As New List(Of Vector3)
                    For Each pt In horiz2dProfile.Points
                        point3dList.Add(New Vector3(pt.X, pt.Y, z))
                    Next
                    _horizProfile = New Profile3D(point3dList)
                End If
            End If
            Return _horizProfile
        End Get
    End Property


    ''' <summary>Get the boundingbox center. Caution : it is projected onto XY plan. Caution 2 : this property may use object models</summary>
    Private _bBoxCenterPt As Point3D = Nothing
    Public ReadOnly Property BoundingBoxCenter As Point3D
        Get
            If _bBoxCenterPt Is Nothing AndAlso BoundingBoxMinMax.Item1 IsNot Nothing Then
                Dim x = (BoundingBoxMinMax.Item2.Point.X + BoundingBoxMinMax.Item1.Point.X) / 2
                Dim y = (BoundingBoxMinMax.Item2.Point.Y + BoundingBoxMinMax.Item1.Point.Y) / 2
                Dim z = (BoundingBoxMinMax.Item2.Point.Z + BoundingBoxMinMax.Item1.Point.Z) / 2
                _bBoxCenterPt = New Point3D(x, y, z)
            End If
            Return _bBoxCenterPt
        End Get
    End Property


    Private _bBoxMinMax As Tuple(Of Point3D, Point3D)
    ''' <summary>Get the boundingbox min and max with attributes. Item1 is min and Item2 is max</summary>
    Public ReadOnly Property BoundingBoxMinMax As Tuple(Of Point3D, Point3D)
        Get
            If _bBoxMinMax Is Nothing Then

                'recuperation boundingbox
                Dim BoundingboxAtt = Me.Metaobject.GetAttribute("RelativeBoundingBox", True)
                If BoundingboxAtt Is Nothing Then
                    BoundingboxAtt = Me.Metaobject.GetAttribute("BoundingBox", True) 'If there is no transformation
                    If BoundingboxAtt Is Nothing Then
                        Return New Tuple(Of Point3D, Point3D)(Nothing, Nothing)
                    End If
                End If

                'Max
                Dim maxAtt = BoundingboxAtt.GetAttribute("Max", True)
                Dim maxPoint As New Point3D(maxAtt.GetAttribute("X").Value, maxAtt.GetAttribute("Y").Value, maxAtt.GetAttribute("Z").Value)

                'Min
                Dim minAtt = BoundingboxAtt.GetAttribute("Min", True)
                Dim minPoint As New Point3D(minAtt.GetAttribute("X").Value, minAtt.GetAttribute("Y").Value, minAtt.GetAttribute("Z").Value)

                _bBoxMinMax = New Tuple(Of Point3D, Point3D)(minPoint, maxPoint)
            End If
            Return _bBoxMinMax

        End Get
    End Property

    Public ReadOnly Property HeightByTgm As Double
        Get
            Return Me.Metaobject.GetAttribute("HeightByTgm", True).Value
        End Get
    End Property


    ''' <summary>Get object height. Caution : this property uses object models</summary>
    Public ReadOnly Property Height As Double
        Get
            If _height Is Nothing Then
                _height = MaxZ - MinZ
            End If
            Return _height
        End Get
    End Property


    Private _maxZ, _minZ As Double
    ''' <summary>Get coord Z of the object top. Caution : this property uses object models</summary>
    Public ReadOnly Property MaxZ As Double
        Get
            If _maxZ = Nothing Then
                If Metaobject.GetAttribute("MaxZ") IsNot Nothing Then
                    _maxZ = CDbl(Metaobject.GetAttribute("MaxZ").Value)
                Else
                    If BoundingBoxMinMax.Item2 Is Nothing Then
                        _maxZ = Nothing
                    Else
                        _maxZ = BoundingBoxMinMax.Item2.Point.Z
                    End If
                End If
            End If
            Return _maxZ
        End Get
    End Property

    ''' <summary>Get coord Z of the object bottom. Caution : this property uses object models</summary>
    Public ReadOnly Property MinZ As Double
        Get
            If _minZ = Nothing Then
                If Metaobject.GetAttribute("MinZ") IsNot Nothing Then
                    _minZ = CDbl(Metaobject.GetAttribute("MinZ").Value)
                Else
                    If BoundingBoxMinMax.Item1 Is Nothing Then
                        _minZ = Nothing
                    Else
                        _minZ = BoundingBoxMinMax.Item1.Point.Z
                    End If
                End If
            End If
            Return _minZ
        End Get
    End Property

    '' <summary>Get coord Z of the object top. Caution : this property uses object models</summary>
    'Public ReadOnly Property MaxZ As Double
    '    Get
    '        If _MaxZ = Nothing Then
    '            If Metaobject.GetAttribute("MaxZ") IsNot Nothing Then
    '                _MaxZ = CDbl(Metaobject.GetAttribute("MaxZ").Value)
    '            Else

    '                _MaxZ = Double.NegativeInfinity
    '                If Me.Models.Count = 0 Then Throw New Exception("No Model")
    '                For Each model In Me.Models
    '                    For Each vert In model.VerticesTesselation
    '                        If vert.Z > _MaxZ Then _MaxZ = vert.Z
    '                    Next
    '                Next
    '                If _MaxZ = Double.NegativeInfinity Then Throw New Exception("MaxZ = NegativeInfinity")
    '                _MaxZ = CDbl(Math.Round(_MaxZ, 3))
    '            End If
    '        End If
    '        Return _MaxZ
    '    End Get
    'End Property
    '''' <summary>Get coord Z of the object bottom. Caution : this property uses object models</summary>
    'Public ReadOnly Property MinZ As Double
    '    Get
    '        If _MinZ = Nothing Then
    '            If Metaobject.GetAttribute("MinZ") IsNot Nothing Then
    '                _MinZ = CDbl(Metaobject.GetAttribute("MinZ").Value)
    '            Else

    '                _MinZ = Double.PositiveInfinity
    '                If Me.Models.Count = 0 Then Throw New Exception("No Model")
    '                For Each model In Me.Models
    '                    For Each vert In model.VerticesTesselation
    '                        If vert.Z < _MinZ Then _MinZ = vert.Z
    '                    Next
    '                Next
    '                If _MinZ = Double.PositiveInfinity Then Throw New Exception("MinZ = PositiveInfinity")
    '                _MinZ = CDbl(Math.Round(_MinZ, 3))
    '            End If
    '        End If
    '        Return _MinZ
    '    End Get
    'End Property

    Private _location As Point3D
    ''' <summary>Get origin point from LocalPlacement.</summary>
    Public Property LocalOrigin As Point3D
        Get
            If _location Is Nothing Then
                If Me.Metaobject.GetAttribute("LocalPlacement") Is Nothing Then
                    Return Nothing
                Else
                    _location = New Point3D(CDbl(Me.Metaobject.SmartGetAttribute("LocalPlacement/_\Location/_\X").Value), CDbl(Me.Metaobject.SmartGetAttribute("LocalPlacement/_\Location/_\Y").Value), CDbl(Me.Metaobject.SmartGetAttribute("LocalPlacement/_\Location/_\Z").Value))
                End If
            End If
            Return _location
        End Get
        Set(value As Point3D)
            _location = value
        End Set
    End Property

#End Region

#Region "ENVELOPE"
    Public Property PositionRelativeToEnvelope As EnvelopePosition
        Get
            If _envAnaRes Is Nothing Then
                Dim envAtt = Me.Metaobject.GetAttribute("EnvelopePositionByTgm", True)?.Value
                If envAtt IsNot Nothing AndAlso System.Enum.GetNames(GetType(EnvelopePosition)).Contains(envAtt.ToString) Then
                    _envAnaRes = System.[Enum].Parse(GetType(EnvelopePosition), envAtt.ToString)
                Else
                    _envAnaRes = EnvelopePosition.NonDefini
                End If
            End If
            Return _envAnaRes
        End Get
        Set(value As EnvelopePosition) 'A sup ! on fait que lire dans ces classes
            _envAnaRes = value
        End Set
    End Property
    Public Property EnvelopeOwner As String
        Get
            If _envOwner Is Nothing Then
                _envOwner = Me.Metaobject.GetAttribute("EnvelopeOwnerByTgm", True)?.Value
            End If
            Return _envOwner
        End Get
        Set(value As String)
            _envOwner = value
        End Set
    End Property

#End Region

#Region "ACTION STATE"
    Public Shared Sub CompleteActionStateTree(projWs As Workspace, actionName As String, exceptionList As HashSet(Of String))
        Dim stateTree = projWs.SmartAddTree("Action States")
        Dim actionNode = stateTree.SmartAddNode(actionName, Nothing)
        For Each excMessage In exceptionList
            actionNode.SmartAddNode(actionName, excMessage)
        Next
    End Sub
    Public Shared Sub CompleteActionStateAttribute(anaObj As AecObject, actionName As String, state As String)
        Dim stateAtt = anaObj.GetWritableTgmObj.SmartAddAttribute("ActionState", Nothing)
        stateAtt.SmartAddAttribute(actionName, state)
    End Sub
#End Region

#Region "PSET TREEGRAM"
    Public Function CompleteTgmPset(attName As String, attValue As Object, Optional propagate As Boolean = False) As Attribute
        If Me.PsetTgmExtension IsNot Nothing Then
            Return Me.PsetTgmExtension.SmartAddAttribute(attName, attValue, propagate)
            'ElseIf Me.ContainerWs.GetAttribute("LinkedData") IsNot Nothing Then
            '    Return Me.Metaobject.SmartAddAttribute(attName, attValue, propagate) 'For objects from TgmCreationWs
        Else
            Return Me.Metaobject.SmartAddAttribute(attName, attValue, propagate)
            'Throw New Exception("Cannot add PsetTgm attribute to this object, probably because extension doesn't exist yet.")
        End If
    End Function

    Public Shared Function CompleteTgmPset(objTgm As MetaObject, attName As String, attValue As Object, Optional propagate As Boolean = False) As Attribute
        Dim anaObj As New AecObject(objTgm)
        Return anaObj.CompleteTgmPset(attName, attValue, propagate)
    End Function
#End Region


End Class


