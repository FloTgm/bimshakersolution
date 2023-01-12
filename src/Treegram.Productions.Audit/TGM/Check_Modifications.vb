Imports M4D.Production.Core
Imports M4D.Treegram.Core.Entities
Imports M4D.Treegram.Core.Kernel
Imports M4D.Treegram.Core.Extensions.Kernel
Imports M4D.Treegram.Core.Extensions.Entities
Imports Tgm.Deixi.Services
Imports M4D.Treegram.Core.Extensions.Models
Public Class Check_ModificationsScript
    Inherits ProdScript
    Public Sub New()
        MyBase.New("TGM :: Check : Modifications")
        AddAction(New Check_Modifications())
    End Sub
End Class
Public Class Check_Modifications
    Inherits ProdAction
    Public Sub New()
        Name = "TGM :: Check : Modifications (Prodaction)"
        PartOfScript = True
    End Sub

    Public Overrides Sub CreateGuiInputTree()
        InputTree = TempWorkspace.AddTree("--- Temporary Input Tree ---")
        Dim fileNode As Node = InputTree.AddNode("object.type", "File")
        fileNode.Description = "Files"

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

        ' Récupération des inputs
        Dim tgmFiles = GetInputAsMetaObjects(Fichiers.Source).ToHashSet 'To get rid of duplicates

        'COPY TREE FROM DICO
        Dim scanWs As Workspace = tgmFiles(0).Container
        Dim projWs As Workspace = tgmFiles(0).GetProject

        Dim dicoWs As Workspace = projWs.GetWorkspaces("Dictionary").FirstOrDefault
        If dicoWs IsNot Nothing Then
            For Each dicoTree In dicoWs.Trees
                If dicoTree.Name = "Modifications between current and previous revision" Then
                    If projWs.Trees.FirstOrDefault(Function(o) o.Name = dicoTree.Name) Is Nothing Then
                        dicoTree.CopyTree(projWs)
                    End If
                End If
            Next
        End If

        'VIEW - Version avec une vue créé par file 
        Dim chargedWs As New List(Of Workspace) From {scanWs}
        Dim checkedNodes As New List(Of Node)
        'Dim savedColors As New Dictionary(Of Node, Tuple(Of System.Windows.Media.Color, Single))
        Dim colorNodes As New List(Of NodeColorTransparency)

        OutputTree = projWs.Trees.Where(Function(o) o.Name = "Modifications between current and previous revision").FirstOrDefault
        Dim myTree = OutputTree
        Dim nonmodifiedNode As Node = myTree.Nodes.Where(Function(o) o.Name = "Modification Status" And o.Value.ToString = "Unchanged").First
        Dim modifiedattrNode As Node = myTree.Nodes.Where(Function(o) o.Name = "Modification Status" And o.Value.ToString = "Modified").First
        Dim modifiedgeoNode As Node = modifiedattrNode.Nodes.First
        Dim newNode As Node = myTree.Nodes.Where(Function(o) o.Name = "Modification Status" And o.Value.ToString = "New").First

        checkedNodes.AddRange({nonmodifiedNode, modifiedattrNode, modifiedgeoNode, newNode})
        'savedColors.Add(nonmodifiedNode, New Tuple(Of Windows.Media.Color, Single)(Windows.Media.Color.FromRgb(191, 191, 191), 1.0F))
        'savedColors.Add(modifiedattrNode, New Tuple(Of Windows.Media.Color, Single)(Windows.Media.Color.FromRgb(255, 184, 54), 1.0F))
        'savedColors.Add(modifiedgeoNode, New Tuple(Of Windows.Media.Color, Single)(Windows.Media.Color.FromRgb(244, 132, 129), 1.0F))
        'savedColors.Add(newNode, New Tuple(Of Windows.Media.Color, Single)(Windows.Media.Color.FromRgb(182, 250, 182), 1.0F))

        colorNodes.Add(New NodeColorTransparency(nonmodifiedNode, Windows.Media.Color.FromRgb(191, 191, 191), 1.0F))
        colorNodes.Add(New NodeColorTransparency(modifiedattrNode, Windows.Media.Color.FromRgb(255, 184, 54), 1.0F))
        colorNodes.Add(New NodeColorTransparency(modifiedgeoNode, Windows.Media.Color.FromRgb(244, 132, 129), 1.0F))
        colorNodes.Add(New NodeColorTransparency(newNode, Windows.Media.Color.FromRgb(182, 250, 182), 1.0F))

        TgmView.CreateView(projWs, tgmFiles(0).Name & " - MODIFICATIONS", chargedWs, myTree, checkedNodes, colorNodes)
        'END OF VIEW CREATION

        Return New SucceededActionResult("Succeeded", System.Drawing.Color.Green, New List(Of Element) From {Fichiers})

    End Function

End Class
