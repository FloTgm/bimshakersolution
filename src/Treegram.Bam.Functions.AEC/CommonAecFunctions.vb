Imports M4D.Treegram.Core.Extensions.Entities
Imports M4D.Treegram.Core.Extensions.Kernel
Imports M4D.Treegram.Core.Kernel
Imports M4D.Treegram.Core.Entities
Imports Treegram.Bam.Libraries.AEC
Imports Treegram.Bam.Libraries.AuditFile
Imports Treegram.ConstructionManagement
Imports Treegram.GeomLibrary
Imports Treegram.GeomKernel.BasicModeler
Imports M4D.Treegram.Design.Kernel
Imports Treegram.GeomFunctions

Public Module CommonAecFunctions


    Public Sub DeactivateTgmObjs(projWs As Workspace, outputWs As Workspace, objTgmType As String, Optional _3dFileName As String = "", Optional buildingName As String = "", Optional storeyName As String = "", Optional apartName As String = "", Optional clearRelations As Boolean = True)
        'DEACTIVATE OBJ
        Dim delInt As Integer
        Dim objToDel As List(Of MetaObject) = outputWs.MetaObjects.Where(Function(mo) mo.IsActive AndAlso mo.GetAttribute("object.type")?.Value = objTgmType).ToList
        For Each obj In objToDel
            If buildingName <> "" Then 'If building criteria
                If obj.GetAttribute("TgmBuilding", True)?.Value <> buildingName Then 'If No building OR Different
                    Continue For
                End If
            End If
            If storeyName <> "" Then 'If storey criteria
                'If obj.GetAttribute("TgmBuildingStorey", True)?.Value Is Nothing OrElse Not obj.GetAttribute("TgmBuildingStorey", True).Value.ToString.Contains(storeyName) Then 'If no storey OR Different
                If obj.GetAttribute("TgmBuildingStorey", True)?.Value Is Nothing OrElse Not Strings.Split(obj.GetAttribute("TgmBuildingStorey", True).Value.ToString, "/_\").Any(Function(o) o = storeyName) Then ' <-- To resolve problem with "R+1 and "R+11" for example
                    Continue For
                End If
            End If

            If apartName <> "" Then 'If apart criteria
                If obj.GetAttribute("ApartmentByTgm", True)?.Value Is Nothing OrElse Not obj.GetAttribute("ApartmentByTgm", True).Value = apartName Then '.ToString.Contains(apartName) Then
                    Continue For
                End If
            End If

            obj.ClearMetaObject(delInt, False, True, clearRelations)
        Next

        'DELETE MODELS
        If _3dFileName <> "" Then
            CommonFunctions.Delete3drtgmFile(outputWs, _3dFileName)
        End If

        If delInt > 0 Then projWs.PushAllModifiedEntities()
    End Sub

    ''' <summary>
    ''' Nécessaire tant qu'on aura des éléments multi-étages dans les maquettes
    ''' </summary>
    ''' <param name="aecObjs"></param>
    ''' <param name="storeyName"></param>
    ''' <returns></returns>
    Public Function RefilterObjectsByTgmStorey(aecObjs As IEnumerable(Of AecObject), storeyName As String) As IEnumerable(Of AecObject)
        'Dim newAnaObjs = anaObjs.Where(Function(o) o.Metaobject.GetParents("TgmBuildingStorey").Contains(storeyRef.Metaobject)).ToList
        'Return anaObjs.Where(Function(o) o.Metaobject.GetAttribute("StoreyByTgm", True)?.Value.contains(storeyName)).ToList
        Dim resultObjs As New List(Of Libraries.AEC.AecObject)
        For Each aecObj In aecObjs
            Dim concatStoreys = aecObj.Metaobject.GetAttribute("StoreyByTgm", True)
            If concatStoreys Is Nothing Then
                Continue For
            End If
            Dim storeys = Strings.Split(concatStoreys.Value, "/_\")
            If storeys.Contains(storeyName) Then
                resultObjs.Add(aecObj)
            End If
        Next
        Return resultObjs
    End Function

    Public Function RefilterObjectsByTgmApartment(aecObjs As IEnumerable(Of AecObject), apartmentName As String) As IEnumerable(Of AecObject)
        Dim resultObjs As New List(Of Libraries.AEC.AecObject)
        For Each aecObj In aecObjs
            Dim concatAparts = aecObj.Metaobject.GetAttribute("ApartmentBelongingByTgm", True)
            If concatAparts Is Nothing Then
                Continue For
            End If
            Dim aparts = Strings.Split(concatAparts.Value, "/_\")
            If aparts.Contains(apartmentName) Then
                resultObjs.Add(aecObj)
            End If
        Next
        Return resultObjs
    End Function

    Public Function ExcludeNoGeomObjects(aecObjList As IEnumerable(Of AecObject)) As IEnumerable(Of AecObject)
        Dim newObjList As New List(Of AecObject)
        For Each aecObj In aecObjList
            If aecObj.Models(True).Count = 0 Then
                aecObj.CompleteTgmPset("HasGeometry", False)
                aecObj.CompleteTgmPset("LodByTgm", 0)
            Else
                aecObj.CompleteTgmPset("HasGeometry", True)
                newObjList.Add(aecObj)
            End If
        Next
        Return newObjList
    End Function

    ''' <summary>
    ''' Returns dictionnary of key : StoreyReference, value : list of objects from initial list with relation to the 
    ''' An object can belong to several storeys
    ''' </summary>
    Public Function GetStoreysFromTgm(objsTgm As List(Of MetaObject), Optional recursively As Boolean = True) As Dictionary(Of MetaObject, List(Of MetaObject))
        Dim storeysDico As New Dictionary(Of MetaObject, List(Of MetaObject))
        For Each objTgm In objsTgm
            Dim storeys = objTgm.GetParents(ProjectReference.Constants.storeyRefName,,, recursively).ToList
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

    Public Function GetEnvelopesFromTgm(objsTgm As List(Of MetaObject), channelWs As Workspace, storeyTgm As MetaObject) As Dictionary(Of String, List(Of MetaObject))
        Dim envDico As New Dictionary(Of String, List(Of MetaObject))
        For Each objTgm In objsTgm
            Dim anaObj As New Libraries.AEC.AecObject(objTgm)
            If anaObj.EnvelopeOwner Is Nothing Then
                Continue For
            End If
            For Each envSt In anaObj.EnvelopeOwner.Split(";")
                Dim envTgm = channelWs.GetMetaObjects(envSt, "EnvelopeByTgm").FirstOrDefault
                If envTgm IsNot Nothing AndAlso envTgm.GetParents(ProjectReference.Constants.storeyRefName).FirstOrDefault.Equals(storeyTgm) Then
                    If envDico.ContainsKey(envSt) Then
                        envDico(envSt).Add(objTgm)
                    Else
                        envDico.Add(envSt, New List(Of MetaObject) From {objTgm})
                    End If
                End If
            Next
        Next
        envDico.OrderBy(Function(k) k)
        Return envDico
    End Function

    Public Function GetApartmentsFromTgm(objsTgm As List(Of MetaObject)) As Dictionary(Of MetaObject, List(Of MetaObject))
        Dim apartDico As New Dictionary(Of MetaObject, List(Of MetaObject))
        For Each objTgm In objsTgm
            Dim aparts = objTgm.GetParents("ApartmentByTgm").ToList
            For Each apartTgm In aparts
                If apartDico.ContainsKey(apartTgm) Then
                    apartDico(apartTgm).Add(objTgm)
                Else
                    apartDico.Add(apartTgm, New List(Of MetaObject) From {objTgm})
                End If
            Next
        Next
        apartDico.OrderBy(Function(k) k)
        Return apartDico
    End Function

    Public Function GetAuditFileFromInputs(allInputs As List(Of MetaObject), Optional createPsetTgmExt As Boolean = True) As FileObject
        Dim projWs = allInputs.First.GetProject
        Dim objWsDico = CommonFunctions.GetWorkspaceContainers(allInputs)
        'Dim scanWs = objWsDico.Keys.FirstOrDefault(Function(w) w.GetAttribute("Source")?.Value = "Scan") 'ScanWs
        Dim scanWs = objWsDico.Keys.FirstOrDefault(Function(w) w.GetMetaObjects.First.GetTgmType = "File") 'ScanWs
        Dim linkedWs = objWsDico.Keys.FirstOrDefault(Function(w) w.GetAttribute("LinkedData")?.Value IsNot Nothing) 'ChannelWs
        If scanWs Is Nothing And linkedWs IsNot Nothing Then
            Dim wsId = linkedWs.GetAttribute("LinkedData").Value.ToString
            scanWs = projWs.Workspaces.Get(New Identifier(wsId))
        ElseIf scanWs Is Nothing Then 'For old projects, to delete one day !
            Throw New Exception("Depuis l'ajout des ""GhostData"", les scans doivent être relancés sur tous les projets ! Si cela a déjà été fait, me contacter (Flo)")
        End If
        Dim fileTgm = scanWs.GetMetaObjects(, "File").First

        Dim myAuditFile
        If fileTgm.GetAttribute("Extension").Value.ToString().ToLower = "ifc" Then
            myAuditFile = New IfcFileObject(fileTgm)
        Else
            myAuditFile = New FileObject(fileTgm)
        End If

        If createPsetTgmExt AndAlso objWsDico.ContainsKey(scanWs) Then
            If myAuditFile.PsetTgmWs IsNot Nothing Then
                myAuditFile.CreatePsetTgmExtensions(objWsDico(scanWs))
            Else
                Throw New Exception("PsetTgm workspace is not created yet for this file, you must scan your file again !")
            End If
        End If

        Return myAuditFile
    End Function

    Public Function GetAuditIfcFileFromInputs(allInputs As List(Of MetaObject), Optional createPsetTgmExt As Boolean = True) As IfcFileObject
        Dim projWs = allInputs.First.GetProject
        Dim objWsDico = CommonFunctions.GetWorkspaceContainers(allInputs)
        'Dim scanWs = objWsDico.Keys.FirstOrDefault(Function(w) w.GetAttribute("Source")?.Value = "Scan") 'ScanWs
        Dim scanWs = objWsDico.Keys.FirstOrDefault(Function(w) w.MetaObjects.Any(Function(o) o.GetTgmType = "File")) 'ScanWs
        Dim linkedWs = objWsDico.Keys.FirstOrDefault(Function(w) w.GetAttribute("LinkedData")?.Value IsNot Nothing) 'ChannelWs
        If scanWs Is Nothing And linkedWs IsNot Nothing Then
            Dim wsId = linkedWs.GetAttribute("LinkedData").Value.ToString
            scanWs = projWs.Workspaces.Get(New Identifier(wsId))
        ElseIf scanWs Is Nothing Then 'For old projects, to delete one day !
            Throw New Exception("Depuis l'ajout des ""GhostData"", les scans doivent être relancés sur tous les projets ! Si cela a déjà été fait, me contacter (Flo)")
        End If
        Dim fileTgm = scanWs.GetMetaObjects(, "File").First
        Dim myAuditFile As New IfcFileObject(fileTgm)

        If createPsetTgmExt AndAlso objWsDico.ContainsKey(scanWs) Then
            If myAuditFile.PsetTgmWs IsNot Nothing Then
                myAuditFile.CreatePsetTgmExtensions(objWsDico(scanWs))
            Else
                Throw New Exception("PsetTgm workspace is not created yet for this file, you must scan your file again !")
            End If
        End If

        Return myAuditFile
    End Function


    Public Function GetIfcSpaceClientSpecifications(myAuditFile As IfcFileObject) As Dictionary(Of String, (AttName As String, AttValue As String))
        Dim ifcSpaceNamingDico As New Dictionary(Of String, (AttName As String, AttValue As String))
        'Dim ifcSpaceNamingDico As New Dictionary(Of String, Tuple(Of String, String))
        Dim ifcSpaceWsheet = myAuditFile.GetOrInsertSettingsWSheet("PIECES ARCHI")
        Dim descriptionRange = ifcSpaceWsheet.Range("description")
        Dim startLine = descriptionRange.BottomRowIndex + 1
        Dim descriptionCol = descriptionRange.LeftColumnIndex
        Dim nameCol = descriptionCol + 1
        Dim valueCol = descriptionCol + 2
        Dim currentLine = startLine
        Do While Not String.IsNullOrWhiteSpace(ifcSpaceWsheet.Cells(currentLine, descriptionCol).Value.TextValue)
            Dim myDescription = ifcSpaceWsheet.Cells(currentLine, descriptionCol).Value.TextValue
            Dim myAttName = ifcSpaceWsheet.Cells(currentLine, nameCol).Value.TextValue
            Dim myAttValue As String
            If ifcSpaceWsheet.Cells(currentLine, valueCol).Value.IsBoolean Then
                myAttValue = ifcSpaceWsheet.Cells(currentLine, valueCol).Value.BooleanValue.ToString
            Else
                myAttValue = ifcSpaceWsheet.Cells(currentLine, valueCol).Value.TextValue
            End If
            ifcSpaceNamingDico.Add(myDescription, (myAttName, myAttValue))
            currentLine += 1
        Loop

        Return ifcSpaceNamingDico
    End Function

    Public AxisAndProfile_DistanceAboveGeom = 0.02


    ''' <summary>
    ''' Create or Read objects axis
    ''' CAUTION ! you must load CHILDREN and EXTENSIONS geometry models before calling this function !
    ''' </summary>
    ''' <param name="aecObjs"></param>
    ''' <param name="myFile"></param>
    ''' <returns></returns>
    Public Function SmartAddAxis(aecObjs As IEnumerable(Of AecObject), myFile As FileObject) As Dictionary(Of AecObject, Curve2D)

        If aecObjs.Count = 0 Then Return Nothing
        Dim outputWs As Workspace = myFile.ProjWs.SmartAddWorkspace(myFile.ScanWs.Name + " - Axis", "AdditionalGeometries", True, "Scan", myFile.ScanWs) 'Pas très logique de garder comme source Scan, à adapter

        Dim myDico As New Dictionary(Of AecObject, Curve2D)
        Dim creationDico As New Dictionary(Of MetaObject, Tuple(Of Curve3D, SharpDX.Color4))
        Dim exceptionList As New HashSet(Of String) From {"Succeeded"}

        Dim i As Integer = 0
        For Each aecObj In aecObjs
            Dim excMessage As String = Nothing
            Dim objAxisExt = aecObj.Metaobject.SmartAddExtension(aecObj.Name + " - Axis", M4D.Treegram.Core.Extensions.Enums.ExtensionType.AdditionalGeometries, outputWs.Name)

            'READ AXIS ATTRIBUTES
            Dim axis = Nothing
            If aecObj.GetType Is GetType(AecWall) Then
                axis = CType(aecObj, AecWall).Axis
            ElseIf aecObj.GetType Is GetType(AecCurtainWall) Then
                axis = CType(aecObj, AecCurtainWall).Axis
            Else
                excMessage = "Cannot get axis on this type of aec object"
                GoTo sendError
            End If

            If axis IsNot Nothing Then
                myDico.Add(aecObj, axis)
                If Treegram.GeomFunctions.Models.GetGeometryModels(objAxisExt).Count = 0 Then
                    'DRAW AXIS IN TGM
                    creationDico.Add(objAxisExt, New Tuple(Of Curve3D, SharpDX.Color4)(axis.ToCurve3d(aecObj.MaxZ + AxisAndProfile_DistanceAboveGeom), New SharpDX.Color4(0F, 1.0F, 0F, 1.0F)))
                    i += 1
                End If
            Else
                'CALCULATE AXIS
                If aecObj.VectorU.IsZero Then
                    excMessage = "TgmLocalPlacement not defined"
                    GoTo sendError
                End If
                axis = CalculateAxis(aecObj)
                If axis Is Nothing Then
                    excMessage = "Couldn't calculate an axis"
                    GoTo sendError
                End If
                myDico.Add(aecObj, axis)

                'WRITE ATTRIBUTES IN TGM
                axis.Write(objAxisExt)

                'DRAW AXIS IN TGM
                creationDico.Add(objAxisExt, New Tuple(Of Curve3D, SharpDX.Color4)(axis.ToCurve3d(aecObj.MaxZ + AxisAndProfile_DistanceAboveGeom), New SharpDX.Color4(0F, 1.0F, 0F, 1.0F)))
                i += 1
            End If

sendError:
            If excMessage Is Nothing Then
                Libraries.AEC.AecObject.CompleteActionStateAttribute(aecObj, "GetAxisAction", "Succeeded")
            Else
                Libraries.AEC.AecObject.CompleteActionStateAttribute(aecObj, "GetAxisAction", excMessage)
                exceptionList.Add(excMessage)
            End If
        Next

        'CompleteActionStateTree(projWs, "GetAxisAction", exceptionList)

        If creationDico.Count > 0 Then
            myFile.ProjWs.PushAllModifiedEntities() 'essential before calling CreateGeometryInTreegramFromLine function
            Treegram.GeomFunctions.Models.CreateGeometryInTreegramFromLine(outputWs, creationDico, "TgmAxes")
        End If
        myFile.ProjWs.PushAllModifiedEntities()

        Return myDico
    End Function

    ''' <summary>
    ''' Geometry Models must be loaded before calling function
    ''' </summary>
    ''' <returns></returns>
    Public Function CalculateAxis(aecObj As AecObject) As Curve2D

        Dim ANGLETOLERANCE As Double = 1 'deg
        Dim MINLENGTH As Double = 0.001 'm

        'Measure wall
        Dim originPt As New Treegram.GeomKernel.BasicModeler.Point(0, 0, 0)
        Dim axisZvec As New Treegram.GeomKernel.BasicModeler.Vector(0, 0, 1)
        Dim globalXpt As New Treegram.GeomKernel.BasicModeler.Point(1, 0, 0)
        Dim globalYpt As New Treegram.GeomKernel.BasicModeler.Point(0, 1, 0)
        Dim globalXinUVpt = Treegram.GeomKernel.BasicModeler.ChangeCoordPointInAnotherAxisSyst(globalXpt, originPt, aecObj.VectorU.ToBasicModelerVector, aecObj.VectorV.ToBasicModelerVector, axisZvec)
        Dim globalYinUVpt = Treegram.GeomKernel.BasicModeler.ChangeCoordPointInAnotherAxisSyst(globalYpt, originPt, aecObj.VectorU.ToBasicModelerVector, aecObj.VectorV.ToBasicModelerVector, axisZvec)
        Dim axisXinUV As New Treegram.GeomKernel.BasicModeler.Vector(originPt, globalXinUVpt)
        Dim axisYinUV As New Treegram.GeomKernel.BasicModeler.Vector(originPt, globalYinUVpt)
        Dim objMinV = aecObj.MinV

        Dim deltaV = Math.Abs(aecObj.MaxV - objMinV)

        'Create axis - placed at z coord MaxZ of wall
        Dim axisStartPt As New Treegram.GeomKernel.BasicModeler.Point(aecObj.MinU, objMinV + deltaV / 2, 0.0)
        Dim axisEndPt As New Treegram.GeomKernel.BasicModeler.Point(aecObj.MaxU, objMinV + deltaV / 2, 0.0)

        Dim axisStartPtGlobal = Treegram.GeomKernel.BasicModeler.ChangeCoordPointInAnotherAxisSyst(axisStartPt, originPt, axisXinUV, axisYinUV, axisZvec)
        Dim axisEndPtGlobal = Treegram.GeomKernel.BasicModeler.ChangeCoordPointInAnotherAxisSyst(axisEndPt, originPt, axisXinUV, axisYinUV, axisZvec)

        Dim axis As New Treegram.GeomKernel.BasicModeler.Line(axisStartPtGlobal, axisEndPtGlobal)
        Dim angle As Double = Math.Abs(axis.Vector.Angle(aecObj.VectorU.ToBasicModelerVector, True)) Mod 180

        'VERIF RESULTAT
        If angle < ANGLETOLERANCE Then  'HYPOTHESES !!!
            Return axis.ToCurve2d
        Else
            Return Nothing
        End If

    End Function

    ''' <summary>
    ''' Create or Read objects oriented profile
    ''' CAUTION, you must load CHILDREN and EXTENSIONS geometry models before calling this function !
    ''' </summary>
    ''' <param name="aecObjs"></param>
    ''' <param name="myFile"></param>
    ''' <returns></returns>
    Public Function SmartAddOrientedBBoxProfile(aecObjs As IEnumerable(Of AecObject), myFile As FileObject) As Dictionary(Of AecObject, Profile2D)

        Dim outputWs As Workspace = myFile.ProjWs.SmartAddWorkspace($"{myFile.ScanWs.Name} - Profile", "AdditionalGeometries", True, "Scan", myFile.ScanWs) 'Pas très logique de garder comme source Scan, à adapter

        Dim i As Integer = 0
        Dim myDico As New Dictionary(Of AecObject, Profile2D)
        Dim creationDico As New Dictionary(Of MetaObject, Tuple(Of List(Of Curve3D), SharpDX.Color4))
        Dim exceptionList As New HashSet(Of String) From {"Succeeded"}
        For Each aecObj In aecObjs
            Dim excMessage As String = Nothing

            ''GET LOCALPLACEMENTS
            'If Not CommonAecGeomFunctions.SmartAddLocalPlacement(anaObj, myFile) Then
            '    excMessage = "Can't get local placement"
            '    GoTo sendError
            'End If

            'READ PROFILE ATTRIBUTES
            'Dim objProfileExt = anaObj.Metaobject.SmartAddExtension($"{anaObj.Name} - {prefix}Profile",  M4D.Treegram.Core.Extensions.Enums.ExtensionType.AdditionalGeometries, outputWs.Name, "Scan") 'Pas très logique de garder comme source Scan, à adapter
            Dim objProfileExt = aecObj.Metaobject.SmartAddExtension($"{aecObj.Name} - HorizontalProfile", M4D.Treegram.Core.Extensions.Enums.ExtensionType.AdditionalGeometries, outputWs.Name)
            Dim profile = aecObj.HorizontalProfile 'Needs correct models

            If profile IsNot Nothing AndAlso profile.Points.Count > 0 Then
                If Treegram.GeomFunctions.Models.GetGeometryModels(objProfileExt).Count = 0 Then
                    '---SAVE PROFILE GEOMETRY IN TGM
                    Dim profile3d = profile.ToProfile3d(aecObj.MaxZ + AxisAndProfile_DistanceAboveGeom)
                    creationDico.Add(objProfileExt, New Tuple(Of List(Of Curve3D), SharpDX.Color4)(profile3d.ToCurve3dList, New SharpDX.Color4(1.0F, 0F, 0F, 1.0F)))
                    i += 1
                End If
            Else
                'CALCULATE PROFILE
                If aecObj.VectorU.IsZero Then
                    excMessage = "You must define LocalPlacement first"
                    GoTo sendError
                End If
                Dim boundary2d = CalculateOrientedBboxProfile(aecObj)
                If boundary2d Is Nothing OrElse boundary2d.Count = 0 Then
                    excMessage = "Couldn't calculate a profile"
                    GoTo sendError
                End If
                profile = New Profile2D(boundary2d)
                profile.Write(objProfileExt)

                'SAVE PROFILE GEOMETRY IN TGM
                Dim boundary3d = boundary2d.Select(Function(o) o.ToCurve3d(aecObj.MaxZ + AxisAndProfile_DistanceAboveGeom)).ToList
                creationDico.Add(objProfileExt, New Tuple(Of List(Of Curve3D), SharpDX.Color4)(boundary3d, New SharpDX.Color4(1.0F, 0F, 0F, 1.0F)))
                i += 1
            End If

            myDico.Add(aecObj, profile)
            AecObject.CompleteActionStateAttribute(aecObj, "GetProfileAction", "Succeeded")

sendError:
            If excMessage IsNot Nothing Then
                AecObject.CompleteActionStateAttribute(aecObj, "GetProfileAction", excMessage)
                exceptionList.Add(excMessage)
            End If
        Next
        'CompleteActionStateTree(projWs, "GetProfileAction", exceptionList)

        If creationDico.Count > 0 Then
            myFile.ProjWs.PushAllModifiedEntities() 'Required before calling CreateGeometryInTreegramFromLine function
            'Treegram.GeomFunctions.Models.CreateGeometryInTreegramFromLine(outputWs, creationDico, $"Tgm{prefix}Profiles")
            Treegram.GeomFunctions.Models.CreateGeometryInTreegramFromLine(outputWs, creationDico, "HorizontalProfile")
        End If
        If myDico.Count > 0 Then
            myFile.ProjWs.PushAllModifiedEntities()
        End If
        Return myDico
    End Function

    Public Function CalculateOrientedBboxProfile(aecObj As AecObject) As List(Of Curve2D)

        Dim lineList As List(Of Treegram.GeomKernel.BasicModeler.Line)
        'Dim ifcProfileAtt = AnalysisObject.GetIfcProfileAtt(aecobj.Metaobject)
        'If ifcProfileAtt IsNot Nothing Then
        '    lineList = AnalysisObject.GetIfcProfile(ifcProfileAtt)
        'Else
        Dim xVec As New Treegram.GeomKernel.BasicModeler.Vector(1, 0, 0)
        Dim yVec As New Treegram.GeomKernel.BasicModeler.Vector(0, 1, 0)
        Dim zVec As New Treegram.GeomKernel.BasicModeler.Vector(0, 0, 1)
        Dim originPt As New Treegram.GeomKernel.BasicModeler.Point(0, 0, 0)

        'get local bounding box
        Dim objMinU = aecObj.MinU
        Dim objMaxU = aecObj.MaxU
        Dim objMinV = aecObj.MinV
        Dim objMaxV = aecObj.MaxV
        Dim localPt1 As New Treegram.GeomKernel.BasicModeler.Point(objMinU, objMinV, 0.0)
        Dim localPt2 As New Treegram.GeomKernel.BasicModeler.Point(objMinU, objMaxV, 0.0)
        Dim localPt3 As New Treegram.GeomKernel.BasicModeler.Point(objMaxU, objMaxV, 0.0)
        Dim localPt4 As New Treegram.GeomKernel.BasicModeler.Point(objMaxU, objMinV, 0.0)

        'get x and y vectors in local axis system
        Dim localAxisAngle = aecObj.VectorU.ToBasicModelerVector.Angle(xVec)
        xVec = xVec.Rotate2(-localAxisAngle)
        yVec = yVec.Rotate2(-localAxisAngle)

        'get world bounding box
        Dim pt1 = Treegram.GeomKernel.BasicModeler.ChangeCoordPointInAnotherAxisSyst(localPt1, originPt, xVec, yVec, zVec)
        Dim pt2 = Treegram.GeomKernel.BasicModeler.ChangeCoordPointInAnotherAxisSyst(localPt2, originPt, xVec, yVec, zVec)
        Dim pt3 = Treegram.GeomKernel.BasicModeler.ChangeCoordPointInAnotherAxisSyst(localPt3, originPt, xVec, yVec, zVec)
        Dim pt4 = Treegram.GeomKernel.BasicModeler.ChangeCoordPointInAnotherAxisSyst(localPt4, originPt, xVec, yVec, zVec)
        lineList = New List(Of Line) From {New Line(pt1, pt2), New Line(pt2, pt3), New Line(pt3, pt4), New Line(pt4, pt1)}

        Return lineList.Select(Function(o) o.ToCurve2d).ToList
    End Function


    Public Function SmartAddProjectionProfile(aecObjs As IEnumerable(Of AecObject), myFile As FileObject) As Dictionary(Of AecObject, Profile2D)

        Dim outputWs As Workspace = myFile.ProjWs.SmartAddWorkspace(myFile.ScanWs.Name + " - Profile", "AdditionalGeometries", True, "Scan", myFile.ScanWs)

        Dim i As Integer = 0
        Dim myDico As New Dictionary(Of AecObject, Profile2D)
        Dim creationDico As New Dictionary(Of MetaObject, Tuple(Of List(Of Curve3D), SharpDX.Color4))
        Dim exceptionList As New HashSet(Of String) From {"Succeeded"}
        For Each aecObj In aecObjs
            Dim excMessage As String = Nothing

            Dim objProfileExt = aecObj.Metaobject.SmartAddExtension(aecObj.Name + " - HorizontalProfile", M4D.Treegram.Core.Extensions.Enums.ExtensionType.AdditionalGeometries, outputWs.Name)
            Dim profile = aecObj.HorizontalProfile '(aecObj.MaxZ + 0.02)

            If profile IsNot Nothing AndAlso profile.Points.Count > 0 Then

                If Treegram.GeomFunctions.Models.GetGeometryModels(objProfileExt).Count = 0 Then
                    '---SAVE PROFILE GEOMETRY IN TGM
                    Dim profile3d = profile.ToProfile3d(aecObj.MaxZ + AxisAndProfile_DistanceAboveGeom)
                    creationDico.Add(objProfileExt, New Tuple(Of List(Of Curve3D), SharpDX.Color4)(profile3d.ToCurve3dList, New SharpDX.Color4(1.0F, 0F, 0F, 1.0F)))
                    i += 1
                End If
            Else
                'CALCULATE PROFILE
                Dim boundaries = Projections2D.GetHorizProjFacingDownward(aecObj.Metaobject, 4) ', aecObj.MaxZ + 0.02)
                If boundaries.Count = 0 Then
                    excMessage = "Couldn't find any profile"
                    GoTo sendError
                End If
                Dim boundary2d = CommonGeomFunctions.BiggestProfile(boundaries)
                profile = New Profile2D(boundary2d)
                profile.Write(objProfileExt)

                'SAVE PROFILE GEOMETRY IN TGM
                Dim boundary3d = boundary2d.Select(Function(o) o.ToCurve3d(aecObj.MaxZ + AxisAndProfile_DistanceAboveGeom)).ToList
                creationDico.Add(objProfileExt, New Tuple(Of List(Of Curve3D), SharpDX.Color4)(boundary3d, New SharpDX.Color4(1.0F, 0F, 0F, 1.0F)))
                i += 1
            End If

            myDico.Add(aecObj, profile)
            AecObject.CompleteActionStateAttribute(aecObj, "GetProjProfileAction", "Succeeded")

sendError:
            If excMessage IsNot Nothing Then
                AecObject.CompleteActionStateAttribute(aecObj, "GetProjProfileAction", excMessage)
                exceptionList.Add(excMessage)
            End If
        Next
        'CompleteActionStateTree(projWs, "GetProfileAction", exceptionList)

        If creationDico.Count > 0 Then
            myFile.ProjWs.PushAllModifiedEntities() 'Required before calling CreateGeometryInTreegramFromLine function
            Treegram.GeomFunctions.Models.CreateGeometryInTreegramFromLine(outputWs, creationDico, "HorizontalProfile")
        End If
        If myDico.Count > 0 Then
            myFile.ProjWs.PushAllModifiedEntities()
        End If

        Return myDico
    End Function

    Enum QuantityComparison
        Same
        LittleDifference
        BigDifference
        Different
        Unknown
        Invalid
    End Enum

    Public Enum MeasureComparison
        NR
        inferior
        equal
        superior
        different
    End Enum

End Module













