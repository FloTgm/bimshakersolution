Imports M4D.Production.Core
Imports M4D.Treegram.Core.Entities
Imports M4D.Treegram.Core.Kernel
Imports M4D.Treegram.Core.Extensions.Entities
Imports Treegram.Bam.Functions.AEC

Public Class Analyse_WallsRvtProdScript
    Inherits ProdScript
    Public Sub New()
        MyBase.New("RVT :: Analyse : Walls")
        AddAction(New Analyse_WallsRvt() With {.Name = Name, .PartOfScript = True})
    End Sub
End Class
Public Class Analyse_WallsRvt
    Inherits Analyse_Walls

    Public Sub New()
    End Sub

    Public Overrides Sub CreateGuiInputTree()
        'Création d'une zone de filtrage (arbre) temporaire
        InputTree = TempWorkspace.AddTree("--- Temporary Input Tree ---")

        'Création d'un filtre : murs
        Dim wallNode = InputTree.SmartAddNode("object.type", "Murs",,, "Murs")

        'Drag and Drop automatique des objets filtrés vers l'algo (Optionnel)
        SelectYourSetsOfInput.Add("Murs", {wallNode}.ToList())
    End Sub

    Public Overrides Function GuiUpdateLauncherTree(inputs As Dictionary(Of String, IEnumerable(Of PersistEntity))) As Tree
        'Séparation des lancements par fichier
        Dim launchTree As Tree = TempWorkspace.SmartAddTree("LauncherTree")
        launchTree.DuplicateObjectsWhileFiltering = False
        CommonAecLauncherTrees.FilesLaunchTree(launchTree, inputs)
        Return launchTree
    End Function

End Class
