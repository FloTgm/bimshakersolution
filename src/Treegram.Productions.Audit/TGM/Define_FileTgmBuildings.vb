Imports M4D.Production.Core
Imports M4D.Treegram.Core.Entities
Imports M4D.Treegram.Core.Kernel
Imports M4D.Treegram.Core.Extensions.Entities
Imports DevExpress.Spreadsheet
Imports Treegram.Bam.Libraries.AEC
Imports Treegram.Bam.Functions.AEC
Imports Treegram.Bam.Libraries.AuditFile
Imports Treegram.ConstructionManagement
Imports Treegram.Bam.Functions
Imports M4D.Treegram.Core.Extensions.Kernel

Public Class Define_FileTgmBuildingsScript
    Inherits ProdScript
    Public Sub New()
        MyBase.New("TGM :: Define : File TgmBuildings")
        AddAction(New Define_FileTgmBuildings)
    End Sub
End Class
Public Class Define_FileTgmBuildings
    Inherits ProdAction
    Public Sub New()
        Name = "TGM :: Define : File TgmBuildings (ProdAction)"
        PartOfScript = True
    End Sub

    Public Overrides Sub CreateGuiInputTree()
        InputTree = TempWorkspace.AddTree("--- Temporary Input Tree ---")
        Dim myFileNode = InputTree.AddNode("object.type", "File")
        myFileNode.Description = "Files"
        SelectYourSetsOfInput.Add("Files", New List(Of Node) From {myFileNode})
    End Sub

    Public Overrides Function GuiUpdateLauncherTree(inputs As Dictionary(Of String, IEnumerable(Of PersistEntity))) As Tree

        Dim launchTree = TempWorkspace.SmartAddTree("LauncherTree")
        'launchTree.SmartAddNode("File", Nothing,,, "<ALL>")

        ' Suppression des anciens noeuds
        While launchTree.Nodes.Count > 0
            launchTree.RemoveNode(launchTree.Nodes.First)
        End While

        If inputs.Values.SelectMany(Function(i) i).ToList.Count > 0 Then
            For Each inputObj In inputs.Values.SelectMany(Function(i) i).ToList
                Dim inputmObj As MetaObject = inputObj
                If inputmObj.GetTgmType <> "File" Then Continue For 'Throw New Exception("You must drag and drop file objects")
                'launchTree.SmartAddNode("object.id", inputmObj.Id.ToString,,, inputmObj.Name)
                launchTree.SmartAddNode("object.name", inputmObj.Name)
            Next
        End If
        Return launchTree
    End Function

    <ActionMethod()>
    Public Function MyMethod(Files As MultipleElements) As ActionResult
        If Files.Source.Count = 0 Then
            Return New IgnoredActionResult("No Input File")
        End If
        ' Récupération des inputs
        Dim tgmFiles = GetInputAsMetaObjects(Files.Source).ToHashSet 'To get rid of duplicates mo
        Dim myAuditFile = CommonAecFunctions.GetAuditIfcFileFromInputs(tgmFiles.ToList)
        If myAuditFile.IsIfc Then
            myAuditFile.CreatePsetTgmExtensions(myAuditFile.ScanWs.GetMetaObjects(, "IfcBuildingStorey").ToList) 'To write TgmBuilding on them
            'myAuditFile.ScanWs.MetaObjects.Where(Function(o) o.GetTgmTypes(True).Contains("FileLevel")).ToList ' MERDIQUE car on récupère le template IfcBuildingStorey dans cette requête
        Else
            myAuditFile.CreatePsetTgmExtensions(myAuditFile.ScanWs.GetMetaObjects(, "FileLevel").ToList) 'To write TgmBuilding on them
        End If

        'ANALYSIS
        Dim buildAndStoReferencesDico = DefineTgmBuildings(myAuditFile)

        'ADD BUILDINGS TO DICTIONARY - For scope use
        Dim nodesTree = myAuditFile.DictionaryWs.SmartAddTree("Saved nodes")
        For Each buildingTgm In buildAndStoReferencesDico.Keys.ToList
            nodesTree.SmartAddNode("TgmBuilding", buildingTgm.Name)
        Next

        'output tree
        OutputTree = TgmBuildingsOutputTree(TempWorkspace, myAuditFile, buildAndStoReferencesDico)

        'send result
        Dim outputs = New List(Of Element) From {New MultipleElements(buildAndStoReferencesDico.Keys.Select(Function(o) o.Metaobject).ToList, "TgmBuildings", ElementState.[New])}
        Return New SucceededActionResult("Succeeded", System.Drawing.Color.Green, outputs)
    End Function

    Private Function DefineTgmBuildings(myFile As FileObject) As Dictionary(Of BuildingReference, List(Of MetaObject))
        Dim projRef As New ProjectReference(myFile.ProjWs)
        Dim buildingRefNames = projRef.Buildings.Select(Function(o) o.Name).ToList
        Dim buildingRefSourceWorkspaces = projRef.Buildings.Select(Function(o) o.FromBeemerFile).ToList

        'COMPLETE BUILDINGS AND STOREYS DICO
        Dim reorderingTree = myFile.ProjWs.Trees.FirstOrDefault(Function(tr) tr.Name = "Buildings Reordering")
        Dim buildAndStoReferencesDico As New Dictionary(Of BuildingReference, List(Of MetaObject))
        Dim scanWsTgmBuilding = myFile.ScanWs.GetAttribute("TgmBuilding", True)?.Value

        If buildingRefSourceWorkspaces.Contains(myFile.ScanWs) Then
            '---REFERENCES HAVE BEEN CREATED FROM THIS FILE
            'Only one building per file for this case
            Dim buildingRef = projRef.Buildings.First(Function(o) o.FromBeemerFile.Equals(myFile.ScanWs))
            If myFile.IsIfc Then
                Dim ifcStoreys = myFile.ScanWs.GetMetaObjects(, "IfcBuildingStorey").ToList
                CommonFunctions.SortLevelsByElevation(ifcStoreys, "Elevation")
                buildAndStoReferencesDico.Add(buildingRef, ifcStoreys)
            ElseIf myFile.IsRevit Then
                Dim rvtLevels = myFile.ScanWs.GetMetaObjects(, "FileLevel").ToList
                CommonFunctions.SortLevelsByElevation(rvtLevels, "OriginElevation")
                buildAndStoReferencesDico.Add(buildingRef, rvtLevels)
            Else
                Throw New Exception("File extension not supported")
            End If

        ElseIf scanWsTgmBuilding IsNot Nothing AndAlso buildingRefNames.Contains(scanWsTgmBuilding.ToString) Then
            '---THANKS TO ATTACHED WORKSPACE
            'Only one building per file for this case
            Dim buildingRef = projRef.Buildings.First(Function(o) o.Name = scanWsTgmBuilding.ToString)
            If myFile.IsIfc Then
                Dim ifcStoreys = myFile.ScanWs.GetMetaObjects(, "IfcBuildingStorey").ToList
                CommonFunctions.SortLevelsByElevation(ifcStoreys, "Elevation")
                buildAndStoReferencesDico.Add(buildingRef, ifcStoreys)
            ElseIf myFile.IsRevit Then
                Dim rvtLevels = myFile.ScanWs.GetMetaObjects(, "FileLevel").ToList
                CommonFunctions.SortLevelsByElevation(rvtLevels, "OriginElevation")
                buildAndStoReferencesDico.Add(buildingRef, rvtLevels)
            Else
                Throw New Exception("File extension not supported")
            End If

        ElseIf reorderingTree IsNot Nothing AndAlso reorderingTree.Nodes.FirstOrDefault(Function(o) o.Value = myFile.FileTgm.Name) IsNot Nothing Then
            '---THANKS TO A REORDERING TREE
            reorderingTree.Filter(False, New List(Of Workspace) From {myFile.ScanWs}).RunSynchronously()
            Dim fileNode = reorderingTree.Nodes.FirstOrDefault(Function(o) o.Value = myFile.FileTgm.Name)
            For Each buildNode In fileNode.Nodes
                Dim storeysList As List(Of MetaObject) = Nothing
                If myFile.IsIfc Then
                    storeysList = buildNode.FilteredMetaObjects.Where(Function(o) o.GetTgmType = "IfcBuildingStorey").ToList
                ElseIf myFile.IsRevit Then
                    storeysList = buildNode.FilteredMetaObjects.Where(Function(o) o.GetTgmType = "FileLevel").ToList
                Else
                    Throw New Exception("File extension not supported")
                End If

                If buildingRefNames.Contains(buildNode.ToString) Then
                    Dim buildingRef = projRef.Buildings.First(Function(o) o.Name = buildNode.ToString)

                    If storeysList?.Count > 0 Then
                        CommonFunctions.SortLevelsByElevation(storeysList, "GlobalElevation")
                        If buildAndStoReferencesDico.ContainsKey(buildingRef) Then
                            storeysList.AddRange(buildAndStoReferencesDico(buildingRef)) 'Complete storeys list
                            buildAndStoReferencesDico.Remove(buildingRef) 'Remove old key
                        End If
                        buildAndStoReferencesDico.Add(buildingRef, storeysList) 'Takes description into account
                    End If
                End If

            Next
            'Dim remainingStoreysList = fileNode.BlockedMetaObjects.Where(Function(o) o.GetTgmType = "IfcBuildingStorey").ToList
            'If remainingStoreysList.Count > 0 Then
            '    buildAndStoReferencesDico.Add("DefaultTgmBuilding", remainingStoreysList)
            'End If
        Else
            '---THANKS TO FILE STRUCTURE
            If myFile.IsIfc Then
                Dim ifcBuildings = myFile.ScanWs.GetMetaObjects(, "IfcBuilding").ToList
                If ifcBuildings.Count = 1 And projRef.Buildings.Count = 1 Then ' 1 building/1 ref CASE !
                    Dim ifcStoreys = ifcBuildings.First.GetChildren("IfcBuildingStorey").ToList
                    CommonFunctions.SortLevelsByElevation(ifcStoreys, "Elevation")
                    buildAndStoReferencesDico.Add(projRef.Buildings.First, ifcStoreys)

                Else
                    For Each ifcBuildingTgm In ifcBuildings
                        Dim ifcStoreys = ifcBuildingTgm.GetChildren("IfcBuildingStorey").ToList
                        CommonFunctions.SortLevelsByElevation(ifcStoreys, "Elevation")
                        If buildingRefNames.Contains(ifcBuildingTgm.Name) Then
                            Dim buildingRef = projRef.Buildings.First(Function(o) o.Name = ifcBuildingTgm.Name)
                            buildAndStoReferencesDico.Add(buildingRef, ifcStoreys)
                        End If
                    Next
                End If

            ElseIf myFile.IsRevit Then
                Dim rvtLevels = myFile.ScanWs.GetMetaObjects(, "FileLevel").ToList
                CommonFunctions.SortLevelsByElevation(rvtLevels, "OriginElevation")

                If projRef.Buildings.Count = 1 Then ' 1 building/1 ref CASE !
                    buildAndStoReferencesDico.Add(projRef.Buildings.First, rvtLevels)
                Else
                    Dim buildingName As String = myFile.FileTgm.SmartGetAttribute("ProjectInformations/_\BuildingName")?.Value?.ToString
                    If buildingRefNames.Contains(buildingName) Then
                        Dim buildingRef = projRef.Buildings.First(Function(o) o.Name = buildingName)
                        buildAndStoReferencesDico.Add(buildingRef, rvtLevels)
                    End If
                End If
            Else
                Throw New Exception("File extension not supported")
            End If
        End If


        'SAVE TGM BUILDINGS ON STOREYS AND FILE
        If myFile.FileTgm.GetAttribute(ProjectReference.Constants.buildingRefName) IsNot Nothing Then
            myFile.FileTgm.RemoveAttribute(ProjectReference.Constants.buildingRefName)
        End If
        Dim buildingConcatSt As String = ""
        Dim i As Integer
        For i = 0 To buildAndStoReferencesDico.Keys.Count - 1
            Dim buildingRef = buildAndStoReferencesDico.Keys(i)
            'If String.IsNullOrWhiteSpace(buildingName) Then buildingName = "DefaultTgmBuilding"
            For Each storeyTgm In buildAndStoReferencesDico(buildingRef)
                'ATTENTION ON ECRIT SUR LE SCAN !!!!!!!!!!!!
                storeyTgm.SmartAddAttribute(ProjectReference.Constants.buildingRefName, buildingRef.Name, True) 'PAS LE CHOIX POUR QUE LA VUE DE CREATION DES CIRCULATIONS VERTICALES FONCTIONNE AVEC LE SCOPE
                AecObject.CompleteTgmPset(storeyTgm, ProjectReference.Constants.buildingRefName, buildingRef.Name, True) 'Celui-ci ne se propage pas aux enfants de l'IfcBuildingStorey...
            Next
            If i > 0 Then
                buildingConcatSt = buildingConcatSt + "/_\" + buildingRef.Name
            Else
                buildingConcatSt = buildingRef.Name
            End If
        Next
        AecObject.CompleteTgmPset(myFile.FileTgm, ProjectReference.Constants.buildingRefName, buildingConcatSt)

        'SAVE FILE ON TGM BUILDINGS
        For Each buildingRef In buildAndStoReferencesDico.Keys
            If buildingRef.Metaobject.GetAttribute("File")?.Value Is Nothing Then 'Essential for BimShaker storey analysis launcher trees
                buildingRef.Metaobject.SmartAddAttribute("File", myFile.FileTgm.Name + ";", True)
            ElseIf Not buildingRef.Metaobject.GetAttribute("File")?.Value.ToString.Contains(myFile.FileTgm.Name) Then
                buildingRef.Metaobject.SmartAddAttribute("File", buildingRef.Metaobject.GetAttribute("File").Value + myFile.FileTgm.Name + ";", True)
            End If
        Next

        'CLEAN FILE ON OTHER TGM BUILDINGS
        For Each buildingRefToClean In projRef.Buildings.Where(Function(o) Not buildAndStoReferencesDico.Keys.Contains(o)).ToList
            If buildingRefToClean.Metaobject.GetAttribute("File")?.Value IsNot Nothing AndAlso
               buildingRefToClean.Metaobject.GetAttribute("File").Value.ToString.Contains(myFile.FileTgm.Name) Then
                Dim fileAttValue As String = buildingRefToClean.Metaobject.GetAttribute("File").Value
                Dim newFileAttValue = fileAttValue.Replace(myFile.FileTgm.Name + ";", "")
                buildingRefToClean.Metaobject.SmartAddAttribute("File", newFileAttValue, True)
            End If
        Next


        myFile.ProjWs.PushAllModifiedEntities()
        Return buildAndStoReferencesDico
    End Function


    Private Function TgmBuildingsOutputTree(projWs As Workspace, myAuditFile As FileObject, buildAndStoReferencesDico As Dictionary(Of BuildingReference, List(Of MetaObject))) As Tree
        Dim myTree As Tree = projWs.SmartAddTree("File TgmBuildings")
        Dim outFileNode = myTree.SmartAddNode("File", myAuditFile.FileTgm.Name)
        For Each buildingKey In buildAndStoReferencesDico.Keys
            Dim buildingName = buildingKey.Name
            Dim tgmBuildNode = outFileNode.SmartAddNode("TgmBuilding", buildingName,,, "TgmBuilding : " + buildingName)
            For Each storeyTgm In buildAndStoReferencesDico(buildingKey)
                tgmBuildNode.SmartAddNode(storeyTgm.GetAttribute("object.type").Value, storeyTgm.Name)
            Next
        Next
        Return myTree
    End Function
End Class
