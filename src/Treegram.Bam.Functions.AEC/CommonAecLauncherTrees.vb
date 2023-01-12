Imports M4D.Treegram.Core.Entities
Imports M4D.Treegram.Core.Extensions.Entities
Imports M4D.Treegram.Core.Kernel
Imports Treegram.ConstructionManagement

Public Module CommonAecLauncherTrees
    Public Sub ApartmentsLaunchTree(ByRef launchTree As Tree, inputs As Dictionary(Of String, IEnumerable(Of PersistEntity)))
        While launchTree.Nodes.Count > 0
            launchTree.RemoveNode(launchTree.Nodes.First)
        End While

        'Dim allNode = launchTree.SmartAddNode("", Nothing,,, "ALL") 'TEST car ça devrait mieux marcher pour les objets dupliqués
        If inputs.Values.SelectMany(Function(i) i).ToList.Count > 0 Then

            Dim mObjList As New List(Of MetaObject)
            For Each inputObj In inputs.Values.SelectMany(Function(i) i).ToList
                Dim inputmObj As MetaObject = inputObj
                mObjList.Add(inputmObj)
            Next

            ''PROJECT DATA
            'Dim projWs, scanWs, channelWs, referencesWs, envFunctionalModelWs, spaceFunctionalModelWs, analysisWs As Workspace
            'TgmKernel.GetProjectWorkspacesFromObject(mObjList(0), projWs, scanWs, channelWs, referencesWs, envFunctionalModelWs, spaceFunctionalModelWs, analysisWs)
            Dim myAuditFile = CommonAecFunctions.GetAuditIfcFileFromInputs(mObjList, False)

            'loop on buildings
            Dim buildingsDico = CommonFunctions.GetBuildingsFromTgm(mObjList)
            If buildingsDico IsNot Nothing AndAlso buildingsDico.Count > 0 Then
                For Each buildingTgm In buildingsDico.Keys
                    Dim buildNode = launchTree.SmartAddNode(ProjectReference.Constants.buildingRefName, buildingTgm.Name)

                    'loop on storeys
                    Dim storeysDico = CommonAecFunctions.GetStoreysFromTgm(buildingsDico(buildingTgm))
                    If storeysDico IsNot Nothing AndAlso storeysDico.Count > 0 Then
                        Dim keys = storeysDico.Keys.ToList
                        CommonFunctions.SortLevelsByElevation(keys, "GlobalElevation")
                        For Each storeyTgm In keys
                            Dim storeyNode = buildNode.SmartAddNode(ProjectReference.Constants.storeyRefName, storeyTgm.Name,,,, True, True) 'Match before and match after for duplex !

                            'loop on envelopes
                            Dim envelopDico = CommonAecFunctions.GetEnvelopesFromTgm(storeysDico(storeyTgm), myAuditFile.TgmCreationWs, storeyTgm)
                            If envelopDico IsNot Nothing AndAlso envelopDico.Count > 0 Then
                                For Each envSt In envelopDico.Keys
                                    Dim envNode = storeyNode.SmartAddNode("EnvelopeOwnerByTgm", envSt,,,, True, True) 'Match before and match after for duplex !

                                    'loop on apartments
                                    Dim apartDico = CommonAecFunctions.GetApartmentsFromTgm(envelopDico(envSt))
                                    If apartDico IsNot Nothing AndAlso apartDico.Count > 0 Then
                                        For Each apartTgm In apartDico.Keys

                                            If apartTgm.GetAttribute(ProjectReference.Constants.storeyRefName).Value.ToString.Contains(storeyTgm.Name) Then 'To avoid confusion with duplex and multi-storeys elements
                                                envNode.SmartAddNode(apartTgm.GetTgmType, apartTgm.Name)
                                            End If
                                        Next
                                    End If
                                Next
                            End If
                        Next
                    End If
                Next
            End If
        End If
    End Sub

    Public Sub ApartmentsLaunchTreeV2(ByRef launchTree As Tree, inputs As Dictionary(Of String, IEnumerable(Of PersistEntity)))
        While launchTree.Nodes.Count > 0
            launchTree.RemoveNode(launchTree.Nodes.First)
        End While

        'Dim allNode = launchTree.SmartAddNode("", Nothing,,, "ALL") 'TEST car ça devrait mieux marcher pour les objets dupliqués
        If inputs.Values.SelectMany(Function(i) i).ToList.Count > 0 Then

            Dim mObjList As New List(Of MetaObject)
            For Each inputObj In inputs.Values.SelectMany(Function(i) i).ToList
                Dim inputmObj As MetaObject = inputObj
                mObjList.Add(inputmObj)
            Next

            'loop on buildings
            Dim buildingsDico = CommonFunctions.GetBuildingsFromTgm(mObjList)
            If buildingsDico IsNot Nothing AndAlso buildingsDico.Count > 0 Then
                For Each buildingTgm In buildingsDico.Keys
                    Dim buildNode = launchTree.SmartAddNode(ProjectReference.Constants.buildingRefName, buildingTgm.Name)

                    'loop on storeys
                    Dim storeysDico = CommonAecFunctions.GetStoreysFromTgm(buildingsDico(buildingTgm))
                    If storeysDico IsNot Nothing AndAlso storeysDico.Count > 0 Then
                        Dim keys = storeysDico.Keys.ToList
                        CommonFunctions.SortLevelsByElevation(keys, "GlobalElevation")
                        For Each storeyTgm In keys
                            Dim storeyNode = buildNode.SmartAddNode(ProjectReference.Constants.storeyRefName, storeyTgm.Name,,,, True, True) 'Match before and match after for duplex !

                            'loop on apartments
                            Dim apartDico = CommonAecFunctions.GetApartmentsFromTgm(storeysDico(storeyTgm))
                            If apartDico IsNot Nothing AndAlso apartDico.Count > 0 Then
                                For Each apartTgm In apartDico.Keys

                                    If apartTgm.GetAttribute(ProjectReference.Constants.storeyRefName).Value.ToString.Contains(storeyTgm.Name) Then 'To avoid confusion with duplex and multi-storeys elements
                                        storeyNode.SmartAddNode(apartTgm.GetTgmType, apartTgm.Name)
                                    End If
                                Next
                            End If
                        Next
                    End If
                Next
            End If
        End If
    End Sub

    Public Sub ApartmentsLaunchTreeV3(ByRef launchTree As Tree, inputs As Dictionary(Of String, IEnumerable(Of PersistEntity)))
        While launchTree.Nodes.Count > 0
            launchTree.RemoveNode(launchTree.Nodes.First)
        End While

        'Dim allNode = launchTree.SmartAddNode("", Nothing,,, "ALL") 'TEST car ça devrait mieux marcher pour les objets dupliqués
        If inputs.Values.SelectMany(Function(i) i).ToList.Count > 0 Then

            Dim mObjList As New List(Of MetaObject)
            For Each inputObj In inputs.Values.SelectMany(Function(i) i).ToList
                Dim inputmObj As MetaObject = inputObj
                mObjList.Add(inputmObj)
            Next

            'loop on buildings
            Dim buildingsDico = CommonFunctions.GetBuildingsFromTgm(mObjList)
            If buildingsDico IsNot Nothing AndAlso buildingsDico.Count > 0 Then
                For Each buildingTgm In buildingsDico.Keys
                    Dim buildNode = launchTree.SmartAddNode(ProjectReference.Constants.buildingRefName, buildingTgm.Name)

                    'loop on apartments
                    Dim apartDico = CommonAecFunctions.GetApartmentsFromTgm(buildingsDico(buildingTgm))
                    If apartDico IsNot Nothing AndAlso apartDico.Count > 0 Then
                        For Each apartTgm In apartDico.Keys
                            buildNode.SmartAddNode(apartTgm.GetTgmType, apartTgm.Name)
                        Next
                    End If
                Next
            End If
        End If
    End Sub

    Public Sub EnvelopesLaunchTree(ByRef launchTree As Tree, inputs As Dictionary(Of String, IEnumerable(Of PersistEntity)))
        While launchTree.Nodes.Count > 0
            launchTree.RemoveNode(launchTree.Nodes.First)
        End While

        'Dim allNode = launchTree.SmartAddNode("", Nothing,,, "ALL") 'TEST car ça devrait mieux marcher pour les objets dupliqués
        If inputs.Values.SelectMany(Function(i) i).ToList.Count > 0 Then

            Dim mObjList As New List(Of MetaObject)
            For Each inputObj In inputs.Values.SelectMany(Function(i) i).ToList
                Dim inputmObj As MetaObject = inputObj
                mObjList.Add(inputmObj)
            Next

            'PROJECT DATA
            Dim myAuditFile = CommonAecFunctions.GetAuditIfcFileFromInputs(mObjList, False)

            'loop on buildings
            Dim buildingsDico = CommonFunctions.GetBuildingsFromTgm(mObjList)
            If buildingsDico IsNot Nothing AndAlso buildingsDico.Count > 0 Then
                For Each buildingTgm In buildingsDico.Keys
                    Dim buildNode = launchTree.SmartAddNode(ProjectReference.Constants.buildingRefName, buildingTgm.Name)

                    'loop on storeys
                    Dim storeysDico = CommonAecFunctions.GetStoreysFromTgm(buildingsDico(buildingTgm))
                    If storeysDico IsNot Nothing AndAlso storeysDico.Count > 0 Then
                        Dim keys = storeysDico.Keys.ToList
                        CommonFunctions.SortLevelsByElevation(keys, "GlobalElevation")
                        For Each storeyTgm In keys
                            Dim storeyNode = buildNode.SmartAddNode(ProjectReference.Constants.storeyRefName, storeyTgm.Name,,,, True, True) 'Match before and match after for duplex !

                            'loop on envelopes
                            Dim envelopDico = CommonAecFunctions.GetEnvelopesFromTgm(storeysDico(storeyTgm), myAuditFile.TgmCreationWs, storeyTgm)
                            If envelopDico IsNot Nothing AndAlso envelopDico.Count > 0 Then
                                For Each envSt In envelopDico.Keys
                                    storeyNode.SmartAddNode("EnvelopeOwnerByTgm", envSt,,,, True, True) 'Match before and match after for duplex !
                                Next
                            End If
                        Next
                    End If
                Next
            End If
        End If
    End Sub

    Public Sub StoreyReferencesLaunchTree(ByRef launchTree As Tree, inputs As Dictionary(Of String, IEnumerable(Of PersistEntity)), Optional storeyMatchBefore As Boolean = True, Optional storeyMatchAfter As Boolean = True)
        While launchTree.Nodes.Count > 0
            launchTree.RemoveNode(launchTree.Nodes.First)
        End While

        'Dim allNode = launchTree.SmartAddNode("", Nothing,,, "ALL") 'TEST car ça devrait mieux marcher pour les objets dupliqués
        If inputs.Values.SelectMany(Function(i) i).ToList.Count > 0 Then

            Dim mObjList As New List(Of MetaObject)
            For Each inputObj In inputs.Values.SelectMany(Function(i) i).ToList
                Dim inputmObj As MetaObject = inputObj
                mObjList.Add(inputmObj)
            Next

            Dim buildingsDico = CommonFunctions.GetBuildingsFromTgm(mObjList)

            'Dim fileNode = launchTree.AddNode("object.type", "File")
            If buildingsDico IsNot Nothing AndAlso buildingsDico.Count > 0 Then
                For Each buildingTgm In buildingsDico.Keys
                    Dim buildNode = launchTree.SmartAddNode(ProjectReference.Constants.buildingRefName, buildingTgm.Name)
                    Dim storeysDico = CommonAecFunctions.GetStoreysFromTgm(buildingsDico(buildingTgm))
                    If storeysDico IsNot Nothing AndAlso storeysDico.Count > 0 Then
                        Dim keys = storeysDico.Keys.ToList
                        CommonFunctions.SortLevelsByElevation(keys, "GlobalElevation")
                        For Each storeyTgm In keys
                            'If storeyTgm.GetAttribute("IsUsed")?.Value = False Then Continue For
                            buildNode.SmartAddNode(ProjectReference.Constants.storeyRefName, storeyTgm.Name,,,, storeyMatchBefore, storeyMatchAfter) 'Match before and match after for duplex !
                        Next
                    End If
                Next
            End If
        End If
    End Sub

    Public Sub BuildingReferencesLaunchTree(ByRef launchTree As Tree, inputs As Dictionary(Of String, IEnumerable(Of PersistEntity)))
        While launchTree.Nodes.Count > 0
            launchTree.RemoveNode(launchTree.Nodes.First)
        End While

        'Dim allNode = launchTree.SmartAddNode("", Nothing,,, "ALL") 'TEST car ça devrait mieux marcher pour les objets dupliqués
        If inputs.Values.SelectMany(Function(i) i).ToList.Count > 0 Then

            Dim mObjList As New List(Of MetaObject)
            For Each inputObj In inputs.Values.SelectMany(Function(i) i).ToList
                Dim inputmObj As MetaObject = inputObj
                mObjList.Add(inputmObj)
            Next

            Dim buildingsDico = CommonFunctions.GetBuildingsFromTgm(mObjList)

            'Dim fileNode = launchTree.AddNode("object.type", "File")
            If buildingsDico IsNot Nothing AndAlso buildingsDico.Count > 0 Then
                For Each buildingTgm In buildingsDico.Keys
                    Dim buildNode = launchTree.SmartAddNode("TgmBuilding", buildingTgm.Name)

                Next
            End If
        End If
    End Sub

    Public Sub IfcBuildingsLaunchTree(ByRef launchTree As Tree, inputs As Dictionary(Of String, IEnumerable(Of PersistEntity)))
        While launchTree.Nodes.Count > 0
            launchTree.RemoveNode(launchTree.Nodes.First)
        End While
        If inputs.Values.SelectMany(Function(i) i).ToList.Count > 0 Then

            Dim mObjList As New List(Of MetaObject)
            For Each inputObj In inputs.Values.SelectMany(Function(i) i).ToList
                Dim inputmObj As MetaObject = inputObj
                mObjList.Add(inputmObj)
            Next
            Dim buildingsDico = CommonFunctions.GetBuildingsFromIfc(mObjList)
            If buildingsDico IsNot Nothing AndAlso buildingsDico.Count > 0 Then
                For Each buildingTgm In buildingsDico.Keys
                    Dim buildNode = launchTree.SmartAddNode("IfcBuilding", buildingTgm.Name)
                Next
            End If
        End If
    End Sub

    Public Sub FileLevelsLaunchTree(ByRef launchTree As Tree, inputs As Dictionary(Of String, IEnumerable(Of PersistEntity)))
        While launchTree.Nodes.Count > 0
            launchTree.RemoveNode(launchTree.Nodes.First)
        End While

        If inputs.Values.SelectMany(Function(i) i).ToList.Count > 0 Then

            Dim mObjList As New List(Of MetaObject)
            For Each inputObj In inputs.Values.SelectMany(Function(i) i).ToList
                Dim inputmObj As MetaObject = inputObj
                mObjList.Add(inputmObj)
            Next

            Dim myFile = CommonAecFunctions.GetAuditFileFromInputs({mObjList.First}.ToList, False)
            Dim storeysDico As Dictionary(Of MetaObject, List(Of MetaObject)) = Nothing
            If myFile.IsIfc Then
                storeysDico = CommonFunctions.GetStoreysFromIfc(mObjList)
            ElseIf myFile.IsRevit Then
                storeysDico = CommonFunctions.GetStoreysFromRevit(mObjList)
            End If

            If storeysDico IsNot Nothing AndAlso storeysDico.Count > 0 Then
                Dim keys = storeysDico.Keys.ToList
                CommonFunctions.SortLevelsByElevation(keys)

                For Each storeyTgm In keys
                    'If storeyTgm.GetAttribute("IsUsed").Value = False Then Continue For
                    launchTree.SmartAddNode("FileLevel", storeyTgm.Name)
                Next
            End If
        End If
    End Sub

    Public Sub FilesLaunchTree(ByRef launchTree As Tree, inputs As Dictionary(Of String, IEnumerable(Of PersistEntity)), Optional specifiedInputSt As String = "")
        'Suppression des anciens noeuds
        While launchTree.Nodes.Count > 0
            launchTree.RemoveNode(launchTree.Nodes.First)
        End While
        'Complete with inputs
        Dim containersList As New HashSet(Of Workspace)
        If specifiedInputSt = "" Then
            For Each obj As MetaObject In inputs.Values.SelectMany(Function(i) i).ToList
                containersList.Add(CType(obj.Container, Workspace))
            Next
        Else
            If inputs.ContainsKey(specifiedInputSt) AndAlso inputs(specifiedInputSt).ToList.Count > 0 Then
                For Each obj As MetaObject In inputs(specifiedInputSt)
                    containersList.Add(CType(obj.Container, Workspace))
                Next
            End If
        End If

        For Each containerWs In containersList
            Dim fileTgm As MetaObject = containerWs.GetMetaObjects(, "File").FirstOrDefault
            If fileTgm IsNot Nothing Then
                Dim fileNode = launchTree.SmartAddNode("File", fileTgm.Name)
            End If
        Next
    End Sub

End Module




