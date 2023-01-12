Imports M4D.Production.Core
Imports M4D.Treegram.Core.Entities
Imports M4D.Treegram.Core.Kernel
Imports M4D.Treegram.Core.Extensions.Entities
Imports Treegram.Bam.Functions

Public Class Create_DynamicTreesScript
    Inherits ProdScript
    Public Sub New()
        MyBase.New("TGM :: Create : Dynamic Trees")
        AddAction(New Create_DynamicTrees())
    End Sub
End Class
Public Class Create_DynamicTrees
    Inherits ProdAction
    Public Sub New()
        Name = "TGM :: Create : Dynamic Trees (ProdAction)"
        PartOfScript = True
    End Sub
    Public Overrides Sub CreateGuiInputTree()
        InputTree = TempWorkspace.AddTree("--- Temporary Input Tree ---")
        Dim fileNode As Node = InputTree.AddNode("object.type", "File")
        InputTree.SmartAddNode("object.type", "File").Description = "Files"

        SelectYourSetsOfInput.Add("Fichiers", {fileNode}.ToList)
    End Sub

    Public Overrides Function GuiUpdateLauncherTree(inputs As Dictionary(Of String, IEnumerable(Of PersistEntity))) As Tree

        Dim launchTree = TempWorkspace.SmartAddTree("LauncherTree")

        ' Suppression des anciens noeuds
        While launchTree.Nodes.Count > 0
            launchTree.RemoveNode(launchTree.Nodes.First)
        End While

        If inputs.Values.SelectMany(Function(i) i).ToList.Count > 0 Then
            For Each inputObj In inputs.Values.SelectMany(Function(i) i).ToList
                Dim inputmObj As MetaObject = inputObj
                If inputmObj.GetTgmType <> "File" Then Continue For 'Throw New Exception("You must drag and drop file objects")
                launchTree.SmartAddNode("object.id", inputmObj.Id.ToString,,, inputmObj.Name)
            Next
        End If

        Return launchTree
    End Function

    <ActionMethod()>
    Public Function MyMethod(Fichiers As MultipleElements) As ActionResult
        If Fichiers.Source.Count = 0 Then
            Return New IgnoredActionResult("No Input File")
        End If
        'Aller chercher tous les objets
        Dim tgmFiles = GetInputAsMetaObjects(Fichiers.Source).ToHashSet 'To get rid of duplicates

        'Creation de l'arbre
        For Each fileTgm In tgmFiles
            OutputTree = GetAttrAnalysis(fileTgm)
        Next

        Return New SucceededActionResult("Succeeded", System.Drawing.Color.Green, New List(Of Element) From {Fichiers})
    End Function

    Public Shared Function GetAttrAnalysis(fileTgm) As Tree
        'Try
        Dim scanWs As Workspace = fileTgm.Container
        Dim projWs As Workspace = scanWs.Container
        Dim typesDico As Dictionary(Of String, List(Of MetaObject))
        Dim ssProjDico As Dictionary(Of String, List(Of MetaObject))
        Dim objList As New List(Of MetaObject)

        Dim _3DObjsList As New List(Of MetaObject)
        Dim subProjList As New List(Of MetaObject)

        scanWs.Join({scanWs})
        scanWs.SmartJoin( M4D.Treegram.Core.Extensions.Enums.ExtensionType.AdditionalAttributes)

        Dim listToLoad As New List(Of Workspace) From {scanWs}
        Dim myTree = projWs.SmartAddTree("Audit - DynamicTree")
        myTree.Filter(False, listToLoad).RunSynchronously()

        _3DObjsList = getAllObjFromAttr(scanWs.MetaObjects.ToList, "object.containedin", "3D")
#Region "getInfos"
        If fileTgm.GetAttribute("Extension").Value = "ifc" Then
            typesDico = CommonFunctions.GetTgmTypes(_3DObjsList)
        Else
            typesDico = CommonFunctions.GetRevitCategories(_3DObjsList)
            ssProjDico = CommonFunctions.GetSousProjet(_3DObjsList)
        End If
#End Region


#Region "CREATE TREE"


        Dim fileNode = myTree.SmartAddNode("File", fileTgm.Name)
        myTree.RemoveNode(fileNode)
        fileNode = myTree.SmartAddNode("File", fileTgm.Name)

        'Create categories/types SubTree
        Dim typeNode As Node
        Dim projNode As Node
        If fileTgm.GetAttribute("Extension").Value = "ifc" Then
            typeNode = fileNode.SmartAddNode("object.type", "ifc",,, "Objets IFC",, True)
        Else
            typeNode = fileNode.SmartAddNode("RevitCategory", Nothing)
            projNode = fileNode.SmartAddNode("Sous-projet", Nothing)
        End If

        For Each typeSt In typesDico.Keys
            Dim typeSsNode As Node
            If fileTgm.GetAttribute("Extension").Value = "ifc" Then
                typeSsNode = typeNode.SmartAddNode("object.type", typeSt)
            Else
                typeSsNode = typeNode.SmartAddNode("RevitCategory", typeSt)
                For Each typeSst In ssProjDico.Keys
                    Dim ssNode As Node
                    ssNode = projNode.SmartAddNode("Sous-projet", typeSst)
                Next
            End If
        Next

        'Create Sous-Projet SubTree

#End Region

        projWs.PushAllModifiedEntities()
        myTree.UpdateModificationDate()
        projWs.Save(Nothing)
        projWs.Push(Nothing)

        Return myTree
        'Catch ex As Exception
        'End Try
    End Function

    Public Shared Function getAllObjFromAttr(metaObjList As List(Of MetaObject), myAttributeName As String, Optional myattributeValue As String = "") As List(Of MetaObject)
        Dim filteredList As New List(Of MetaObject)

        If metaObjList.Count = 0 Then
            Return Nothing
        End If

        For Each elt In metaObjList
            Dim myVal = elt.GetAttribute(myAttributeName)?.Value
            If elt.IsActive And Not IsNothing(myVal) Then
                If myattributeValue = "" Or myVal = myattributeValue Then
                    filteredList.Add(elt)
                End If
            End If
        Next
        Return filteredList
    End Function
End Class
