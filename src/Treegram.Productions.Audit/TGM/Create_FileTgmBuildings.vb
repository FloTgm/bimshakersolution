Imports M4D.Production.Core
Imports M4D.Treegram.Core.Entities
Imports M4D.Treegram.Core.Kernel
Imports M4D.Treegram.Core.Extensions.Entities
Imports M4D.Treegram.Core.Extensions.Kernel
Imports DevExpress.Spreadsheet
Imports Treegram.Bam.Libraries.AEC
Imports Treegram.Bam.Functions.AEC
Imports Treegram.Bam.Libraries.AuditFile
Imports Treegram.ConstructionManagement
Imports Treegram.Bam.Functions

'Public Class Create_FileTgmBuildingsScript
'    Inherits ProdScript
'    Public Sub New()
'        MyBase.New("TGM :: Create : File TgmBuildings (MUTE)")
'        AddAction(New Create_FileTgmBuildings)
'    End Sub
'End Class
Public Class Create_FileTgmBuildings
    Inherits ProdAction
    Public Sub New()
        Name = "TGM :: Create : File TgmBuildings (MUTE)"
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
        'analysis
        Dim buildAndStoReferencesDico = DefineTgmBuildings(myAuditFile)

        'ADD BUILDINGS TO DICTIONARY - For scope use
        Dim nodesTree = myAuditFile.DictionaryWs.SmartAddTree("Saved nodes")
        For Each buildingTgm In buildAndStoReferencesDico.Keys.ToList
            nodesTree.SmartAddNode("TgmBuilding", buildingTgm.Name)
        Next

        'output tree
        OutputTree = TgmBuildingsOutputTree(TempWorkspace, myAuditFile, buildAndStoReferencesDico)

        'send result
        Dim outputs = New List(Of Element) From {New MultipleElements(buildAndStoReferencesDico.Keys.ToList, "TgmBuildings", ElementState.[New])}
        Return New SucceededActionResult("Succeeded", System.Drawing.Color.Green, outputs)
    End Function

    Private Function DefineTgmBuildings(myAuditFile As FileObject) As Dictionary(Of MetaObject, List(Of MetaObject))

        'COMPLETE BUILDINGS AND STOREYS DICO
        Dim buildAndStoReferencesDico = GetFileTgmBuildings(myAuditFile)

        'SAVE TGM BUILDINGS ON OBJECTS 
        If myAuditFile.FileTgm.GetAttribute("TgmBuilding") IsNot Nothing Then
            myAuditFile.FileTgm.RemoveAttribute("TgmBuilding")
        End If
        Dim buildingConcatSt As String = ""
        Dim i As Integer
        For i = 0 To buildAndStoReferencesDico.Keys.Count - 1
            Dim buildingKey = buildAndStoReferencesDico.Keys(i)
            Dim buildingName = buildingKey
            If String.IsNullOrWhiteSpace(buildingName) Then buildingName = "DefaultTgmBuilding"
            For Each storeyTgm In buildAndStoReferencesDico(buildingKey)
                'ATTENTION ON ECRIT SUR LE SCAN !!!!!!!!!!!!
                storeyTgm.SmartAddAttribute("TgmBuilding", buildingName, True) 'PAS LE CHOIX POUR QUE LA VUE DE CREATION DES CIRCULATIONS VERTICALES FONCTIONNE AVEC LE SCOPE
                AecObject.CompleteTgmPset(storeyTgm, ProjectReference.Constants.buildingRefName, buildingName, True) 'Celui-ci ne se propage pas aux enfants de l'IfcBuildingStorey...
            Next
            If i > 0 Then
                buildingConcatSt = buildingConcatSt + "/_\" + buildingName
            Else
                buildingConcatSt = buildingName
            End If
        Next
        'myAuditFile.FileTgm.SmartAddAttribute("TgmBuilding", buildingConcatSt) 'Très important de ne pas propager !!!
        AecObject.CompleteTgmPset(myAuditFile.FileTgm, ProjectReference.Constants.buildingRefName, buildingConcatSt)

        'CREATE TGM BUILDINGS
        Dim newBuildAndStoReferencesDico As New Dictionary(Of MetaObject, List(Of MetaObject))
        For i = 0 To buildAndStoReferencesDico.Keys.Count - 1
            Dim buildingKey = buildAndStoReferencesDico.Keys(i)
            Dim buildingName = buildingKey
            If String.IsNullOrWhiteSpace(buildingName) Then buildingName = "DefaultTgmBuilding"

            'Get ProjRef ws
            Dim projectRefWs As Workspace
            projectRefWs = myAuditFile.ProjWs.Workspaces.Where(Function(ws) ws.Name = M4D.Treegram.Core.Constants.WorkspaceName.TreegramProjectReferenceWorkspace).FirstOrDefault
            If projectRefWs Is Nothing Then
                projectRefWs = myAuditFile.ProjWs.SmartAddWorkspace(M4D.Treegram.Core.Constants.WorkspaceName.TreegramProjectReferenceWorkspace, Nothing)
                Dim origSourceAtt = myAuditFile.ScanWs.GetAttribute("OriginPoint")
                'Dim origDestAtt = projectRefWs.SmartAddAttribute("OriginPoint", Nothing)
                origSourceAtt.CopyAttributeTo(projectRefWs)
                projectRefWs.SmartAddAttribute("ScanWs", myAuditFile.ScanWs.Id.ToString)
            End If

            'Create TgmBuilding
            Dim buildingRefTgm = projectRefWs.SmartAddMetaObject(buildingName, ProjectReference.Constants.buildingRefName)
            If buildingRefTgm.GetAttribute("File")?.Value Is Nothing Then 'Necessary for storey analysis launcher trees
                buildingRefTgm.SmartAddAttribute("File", myAuditFile.FileTgm.Name, True)
            ElseIf Not buildingRefTgm.GetAttribute("File")?.Value.ToString.Contains(myAuditFile.FileTgm.Name) Then
                buildingRefTgm.SmartAddAttribute("File", buildingRefTgm.GetAttribute("File").Value + ";" + myAuditFile.FileTgm.Name, True)
            End If
            newBuildAndStoReferencesDico.Add(buildingRefTgm, buildAndStoReferencesDico.Values(i))
        Next

        myAuditFile.ProjWs.PushAllModifiedEntities()
        Return newBuildAndStoReferencesDico
    End Function


    ''' <summary>
    ''' Find out what are TgmBuildings for a specific file, and use 'Buildings Reordering' Tree if existing
    ''' </summary>
    Private Function GetFileTgmBuildings(myAuditFile As FileObject) As Dictionary(Of String, List(Of MetaObject))
        Dim reorderingTree = myAuditFile.ProjWs.Trees.FirstOrDefault(Function(tr) tr.Name = "Buildings Reordering")

        Dim buildAndStoReferencesDico As New Dictionary(Of String, List(Of MetaObject))
        If myAuditFile.ScanWs.GetAttribute("TgmBuilding", True)?.Value IsNot Nothing Then
            '---THANKS TO ATTACHED WORKSPACE
            'Only one building per file for this case
            Dim buildingName = myAuditFile.ScanWs.GetAttribute("TgmBuilding").Value
            If myAuditFile.IsIfc Then
                Dim ifcStoreys = myAuditFile.ScanWs.GetMetaObjects(, "IfcBuildingStorey").ToList
                CommonFunctions.SortLevelsByElevation(ifcStoreys, "Elevation")
                buildAndStoReferencesDico.Add(buildingName, ifcStoreys)
            ElseIf myAuditFile.IsRevit Then
                Dim rvtLevels = myAuditFile.ScanWs.GetMetaObjects(, "FileLevel").ToList
                CommonFunctions.SortLevelsByElevation(rvtLevels, "OriginElevation")
                buildAndStoReferencesDico.Add(buildingName, rvtLevels)
            Else
                Throw New Exception("File extension not supported")
            End If

        ElseIf reorderingTree IsNot Nothing AndAlso reorderingTree.Nodes.FirstOrDefault(Function(o) o.Value = myAuditFile.FileTgm.Name) IsNot Nothing Then
            '---THANKS TO A REORDERING TREE
            reorderingTree.Filter(False, New List(Of Workspace) From {myAuditFile.ScanWs}).RunSynchronously()
            Dim fileNode = reorderingTree.Nodes.FirstOrDefault(Function(o) o.Value = myAuditFile.FileTgm.Name)
            For Each buildNode In fileNode.Nodes
                Dim storeysList As List(Of MetaObject) = Nothing
                If myAuditFile.IsIfc Then
                    storeysList = buildNode.FilteredMetaObjects.Where(Function(o) o.GetTgmType = "IfcBuildingStorey").ToList
                ElseIf myAuditFile.IsRevit Then
                    storeysList = buildNode.FilteredMetaObjects.Where(Function(o) o.GetTgmType = "FileLevel").ToList
                Else
                    Throw New Exception("File extension not supported")
                End If
                If storeysList?.Count > 0 Then
                    CommonFunctions.SortLevelsByElevation(storeysList, "GlobalElevation")
                    If buildAndStoReferencesDico.ContainsKey(buildNode.ToString) Then
                        storeysList.AddRange(buildAndStoReferencesDico(buildNode.ToString)) 'Complete storeys list
                        buildAndStoReferencesDico.Remove(buildNode.ToString) 'Remove old key
                    End If
                    buildAndStoReferencesDico.Add(buildNode.ToString, storeysList) 'Takes description into account
                End If
            Next
            Dim remainingStoreysList = fileNode.BlockedMetaObjects.Where(Function(o) o.GetTgmType = "IfcBuildingStorey").ToList
            If remainingStoreysList.Count > 0 Then
                buildAndStoReferencesDico.Add("DefaultTgmBuilding", remainingStoreysList)
            End If
        Else
            '---THANKS TO FILE STRUCTURE
            If myAuditFile.IsIfc Then
                Dim ifcBuildings = myAuditFile.ScanWs.GetMetaObjects(, "IfcBuilding").ToList
                For Each ifcBuildingTgm In ifcBuildings
                    Dim ifcStoreys = ifcBuildingTgm.GetChildren("IfcBuildingStorey").ToList
                    CommonFunctions.SortLevelsByElevation(ifcStoreys, "Elevation")
                    Dim buildingName = If(ifcBuildingTgm.Name = "", "DefaultTgmBuilding", ifcBuildingTgm.Name)
                    buildAndStoReferencesDico.Add(buildingName, ifcStoreys)
                Next
            ElseIf myAuditFile.IsRevit Then
                Dim rvtLevels = myAuditFile.ScanWs.GetMetaObjects(, "FileLevel").ToList
                CommonFunctions.SortLevelsByElevation(rvtLevels, "OriginElevation")
                Dim buildingName As String = myAuditFile.FileTgm.SmartGetAttribute("ProjectInformations/_\BuildingName")?.Value?.ToString
                If buildingName = "" Then buildingName = "DefaultTgmBuilding"
                buildAndStoReferencesDico.Add(buildingName, rvtLevels)
            Else
                Throw New Exception("File extension not supported")
            End If
        End If

        Return buildAndStoReferencesDico
    End Function

    Private Function TgmBuildingsOutputTree(projWs As Workspace, myAuditFile As FileObject, buildAndStoReferencesDico As Dictionary(Of MetaObject, List(Of MetaObject))) As Tree
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
