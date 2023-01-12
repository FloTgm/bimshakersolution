Imports System.IO
Imports M4D.Treegram.Core.Extensions.Entities
Imports M4D.Deixi.Core
Imports M4D.Treegram.Core.Kernel
Imports M4D.Treegram.Core.Entities
Imports M4D.Treegram.Design.Kernel
Imports M4D.Treegram.Core.Extensions.Kernel
Imports M4D.Treegram.Core.Constants

Public Module CommonFunctions

    Public Sub SortLists(Of T)(ByRef lists As IList(), comparisonMethod As Func(Of T, T, Integer))
        Dim firstList = lists.First
        For index = firstList.Count - 1 To 1 Step -1
            For j_index = 0 To index - 1
                If comparisonMethod(firstList(j_index + 1), firstList(j_index)) < 0 Then
                    For Each list In lists
                        Dim temp = list(j_index)
                        list(j_index) = list(j_index + 1)
                        list(j_index + 1) = temp
                    Next
                End If
            Next
        Next
    End Sub
    Public Sub RunFilter(projWs As Workspace, Optional wsList As List(Of Workspace) = Nothing)
        projWs.SmartAddTree("ALL").Filter(False, wsList).RunSynchronously()
        'projWs.Join(wsList)
    End Sub

    Public Function GetFileMetaObject(ByVal workspace As Workspace) As MetaObject
        If workspace Is Nothing Then Throw New Exception("invalid inputs")
        Dim metaObject = workspace.MetaObjects.FirstOrDefaultWithoutLoading()
        If metaObject Is Nothing Then Return Nothing
        If metaObject.GetAttribute("object.type")?.Value?.ToString() = "File" Then Return metaObject
        Return workspace.MetaObjects.FirstOrDefault(Function(mo) mo.GetAttribute("object.type")?.Value?.ToString() = "File")
    End Function

    Public Sub Delete3drtgmFile(outputWs As Workspace, _3dFileName As String)
        Dim directory3dBin = Treegram.GeomFunctions.PathUtils.Get3dWorkspaceDirectory(outputWs)
        _3dFileName = _3dFileName.Replace(".3drtgm", "")
        Dim file3dBinPath = Path.Combine(directory3dBin, _3dFileName + ".3drtgm")
        If File.Exists(file3dBinPath) Then
            File.Delete(file3dBinPath)
        End If
    End Sub

    Public Function ConvertProjElevationInScanWs(elevation As Double, sourceWs As Workspace, destWs As Workspace) As Double
        Dim elevPoint As New SharpDX.Vector3(0, 0, elevation)
        Dim newElevPoint = Treegram.GeomFunctions.GeoReferencing.TransformPointInOtherWorkspaceAxes(elevPoint, sourceWs, destWs)
        Return newElevPoint.Z
    End Function

    Public Function ConvertGlobalElevationToScanWs(elevation As Double, destWs As Workspace) As Double
        Dim elevPoint As New SharpDX.Vector3(0, 0, elevation)
        Dim newElevPoint = Treegram.GeomFunctions.GeoReferencing.TransformPointFromGlobalAxes(elevPoint, destWs)
        Return newElevPoint.Z
    End Function

    Public Sub GetGeodata(ByVal workspace As Workspace, ByRef rotationToApply As Double?, ByRef translationToApply As Double())
        Dim basePoint = workspace.Attributes.Where(Function(a) a.Name = "OriginPoint").FirstOrDefault
        Dim angleResult = Nothing, xValue = Nothing, yValue = Nothing, zValue = Nothing

        If basePoint IsNot Nothing Then
            Dim x = basePoint.GetAttribute("X")
            Dim y = basePoint.GetAttribute("Y")
            Dim z = basePoint.GetAttribute("Z")
            Dim angle = basePoint.GetAttribute("Angle")
            Dim relativeTo = basePoint.GetAttribute("RelativeTo")

            If relativeTo IsNot Nothing Then
                Try
                    Dim relativeToId = New Identifier(TryCast(relativeTo.Value, String))
                    If relativeToId <> workspace.Id Then
                        Dim relativeToWorkspace = (TryCast(workspace.Container, Workspace)).Workspaces.[Get](relativeToId)
                        GetGeodata(relativeToWorkspace, rotationToApply, translationToApply)
                    End If
                Catch
                End Try
            End If

            If angle IsNot Nothing AndAlso angle.TryGetValueAsDouble(angleResult) Then
                If rotationToApply Is Nothing Then rotationToApply = 0
                rotationToApply += angleResult
            End If

            If x IsNot Nothing AndAlso y IsNot Nothing AndAlso z IsNot Nothing AndAlso x.TryGetValueAsDouble(xValue) AndAlso y.TryGetValueAsDouble(yValue) AndAlso z.TryGetValueAsDouble(zValue) Then
                If translationToApply Is Nothing Then
                    translationToApply = {xValue, yValue, zValue}
                Else
                    translationToApply = {translationToApply(0) + xValue, translationToApply(1) + yValue, translationToApply(2) + zValue}
                End If
            End If
        End If
    End Sub

    Private Sub FindOutChildrenAtt(att As Attribute, ByRef moAttCOunt As Integer, ByRef moFilledAttCOunt As Integer)
        For Each childAtt In att.Attributes
            If childAtt.Attributes.Count > 0 Then 'Doesn't take in account parent attribute - To discuss...
                FindOutChildrenAtt(childAtt, moAttCOunt, moFilledAttCOunt)
            Else
                moAttCOunt += 1
                Try
                    If Not String.IsNullOrEmpty(childAtt?.Value?.ToString) Then
                        If childAtt.Value.ToString <> "" Then
                            moFilledAttCOunt += 1
                        End If
                    End If
                Catch ex As Exception
                End Try
            End If
        Next
    End Sub

    Public Sub SortLevelsByElevation(storeysList As List(Of MetaObject), Optional elevAttName As String = "GlobalElevation")
        For Each Storey In storeysList
            If Storey.GetAttribute(elevAttName, True) Is Nothing Then Throw New Exception("storey elevation missing")
        Next
        'Sort levels
        Dim elevations = storeysList.Select(Function(o) CDbl(o.GetAttribute(elevAttName, True).Value)).ToList
        Dim array() As IList = {elevations, storeysList}
        SortLists(Of Double)(array, Function(i1 As Double, i2 As Double)
                                        Return i1.CompareTo(i2)
                                    End Function)
    End Sub

    Public Function GetWorkspaceContainers(objsTgm As List(Of MetaObject)) As Dictionary(Of Workspace, List(Of MetaObject))
        Dim wsDico As New Dictionary(Of Workspace, List(Of MetaObject))
        For Each objTgm In objsTgm
            Dim containerWs As Workspace = objTgm.Container
            If wsDico.ContainsKey(containerWs) Then
                wsDico(containerWs).Add(objTgm)
            Else
                wsDico.Add(containerWs, New List(Of MetaObject) From {objTgm})
            End If
        Next
        Return wsDico
    End Function

    Public Function GetTgmTypes(objReq As List(Of MetaObject)) As Dictionary(Of String, List(Of MetaObject))
        Dim typesDico As New Dictionary(Of String, List(Of MetaObject))
        For Each objTgm In objReq
            Dim ifcType = objTgm.GetTgmType
            If typesDico.ContainsKey(ifcType) Then
                typesDico(ifcType).Add(objTgm)
            Else
                typesDico.Add(ifcType, New List(Of MetaObject) From {objTgm})
            End If
        Next
        Return typesDico
    End Function

    Public Function GetRevitCategories(objReq As List(Of MetaObject)) As Dictionary(Of String, List(Of MetaObject))
        Dim categsDico As New Dictionary(Of String, List(Of MetaObject))
        For Each objTgm In objReq
            Dim rvtCateg = objTgm.GetAttribute("RevitCategory").Value
            If categsDico.ContainsKey(rvtCateg) Then
                categsDico(rvtCateg).Add(objTgm)
            Else
                categsDico.Add(rvtCateg, New List(Of MetaObject) From {objTgm})
            End If
        Next
        Return categsDico
    End Function

    Public Function GetSousProjet(objReq As List(Of MetaObject)) As Dictionary(Of String, List(Of MetaObject))
        Dim ssProjDico As New Dictionary(Of String, List(Of MetaObject))
        For Each objTgm In objReq
            Dim ssProj = objTgm?.GetAttribute("PG_IDENTITY_DATA")?.GetAttribute("Sous-projet")?.Value
            If Not IsNothing(ssProj) Then
                If ssProjDico.ContainsKey(ssProj) Then
                    ssProjDico(ssProj).Add(objTgm)
                Else
                    ssProjDico.Add(ssProj, New List(Of MetaObject) From {objTgm})
                End If
            End If
        Next
        Return ssProjDico
    End Function

    Public Function GetBuildingsFromTgm(objsTgm As List(Of MetaObject)) As Dictionary(Of MetaObject, List(Of MetaObject))
        Return GetBuildingsFromSource(objsTgm, "TgmBuilding")
    End Function
    Public Function GetBuildingsFromIfc(objsTgm As List(Of MetaObject)) As Dictionary(Of MetaObject, List(Of MetaObject))
        Return GetBuildingsFromSource(objsTgm, "IfcBuilding")
    End Function

    Private Function GetBuildingsFromSource(objsTgm As List(Of MetaObject), buildingObjType As String) As Dictionary(Of MetaObject, List(Of MetaObject))
        Dim buildingsDico As New Dictionary(Of MetaObject, List(Of MetaObject))
        For Each objTgm In objsTgm
            Dim buildingTgm = objTgm.GetParents(buildingObjType,,, True).FirstOrDefault
            If buildingTgm Is Nothing Then Continue For
            If buildingsDico.ContainsKey(buildingTgm) Then
                buildingsDico(buildingTgm).Add(objTgm)
            Else
                buildingsDico.Add(buildingTgm, New List(Of MetaObject) From {objTgm})
            End If
        Next
        buildingsDico.OrderBy(Function(k) k)
        Return buildingsDico
    End Function


    Public Function GetStoreysFromIfc(objsTgm As List(Of MetaObject), Optional recursively As Boolean = True) As Dictionary(Of MetaObject, List(Of MetaObject))
        Dim storeysDico As New Dictionary(Of MetaObject, List(Of MetaObject))
        For Each objTgm In objsTgm
            Dim storeyTgm = objTgm.GetParents("IfcBuildingStorey",,, recursively).FirstOrDefault
            If storeyTgm IsNot Nothing Then
                If storeysDico.ContainsKey(storeyTgm) Then
                    storeysDico(storeyTgm).Add(objTgm)
                Else
                    storeysDico.Add(storeyTgm, New List(Of MetaObject) From {objTgm})
                End If
            End If
        Next
        storeysDico.OrderBy(Function(k) k)
        Return storeysDico
    End Function

    ''' <summary>
    ''' Returns dictionnary of key : FileLevel, value : list of objects from initial list with relation to the FileLevel.
    ''' An object can belong to several storeys
    ''' </summary>
    Public Function GetStoreysFromRevit(objsTgm As List(Of MetaObject)) As Dictionary(Of MetaObject, List(Of MetaObject))
        Dim storeysDico As New Dictionary(Of MetaObject, List(Of MetaObject))
        For Each objTgm In objsTgm
            Dim storeys = objTgm.GetParents(Nothing, Nothing, RelationType.Containance).ToList.Where(Function(o) o.GetTgmType = "FileLevel").ToList
            For Each stoRef In storeys
                If storeysDico.ContainsKey(stoRef) Then
                    storeysDico(stoRef).Add(objTgm)
                Else
                    storeysDico.Add(stoRef, New List(Of MetaObject) From {objTgm})
                End If
            Next
        Next
        storeysDico.OrderBy(Function(k) k)
        Return storeysDico
    End Function

    ''' <summary>
    ''' For a given attribute and list of MetaObjects, returns a dictionnary with values of attributes as keys and the list of MetaObjects with given attribute value as dictionnary value
    ''' </summary>
    Public Function GetAttributeDico(attrName As String, objsTgm As List(Of MetaObject)) As Dictionary(Of String, List(Of MetaObject))
        Dim attributeDico As New Dictionary(Of String, List(Of MetaObject))

        For Each objTgm In objsTgm
            If objTgm.GetAttribute(attrName, True) IsNot Nothing Then
                Dim attrValue = objTgm.GetAttribute(attrName, True).Value

                If attributeDico.ContainsKey(attrValue) Then
                    attributeDico(attrValue).Add(objTgm)
                Else
                    attributeDico.Add(attrValue, New List(Of MetaObject) From {objTgm})
                End If
            End If
        Next
        attributeDico.OrderBy(Function(k) k)
        Return attributeDico
    End Function

    Public Function RefilterObjectsByTgmStorey(tgmObjs As List(Of MetaObject), storeyName As String) As List(Of MetaObject)
        'Dim newSpacesTgmList = SpacesTgmList.Where(Function(o) o.GetParents("TgmBuildingStorey").Contains(storeyRefTgm)).ToList
        'Return tgmObjs.Where(Function(o)?.Value.Contains(storeyName)).ToList
        Dim resultObjs As New List(Of MetaObject)
        For Each objTgm In tgmObjs
            Dim concatStoreys = objTgm.GetAttribute("StoreyByTgm", True)
            If concatStoreys Is Nothing Then
                Continue For
            End If
            Dim storeys = Strings.Split(concatStoreys.Value, "/_\")
            If storeys.Contains(storeyName) Then
                resultObjs.Add(objTgm)
            End If
        Next
        Return resultObjs
    End Function

    Public Function ProjectFirstGeoReferencedWs(projWs As Workspace) As Workspace
        Dim firstGeoReferencedWorkspace As Workspace = Nothing
        For Each ws In projWs.Workspaces.ToList
            If ws.GetAttribute(AttributeName.OriginPoint) IsNot Nothing Then
                firstGeoReferencedWorkspace = ws
                Exit For
            End If
        Next

        Return firstGeoReferencedWorkspace
    End Function
End Module













