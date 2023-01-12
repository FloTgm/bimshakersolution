Imports M4D.Production.Core
Imports M4D.Treegram.Core.Entities
Imports M4D.Treegram.Core.Kernel
Imports M4D.Treegram.Core.Extensions.Entities
Imports Treegram.Bam.Libraries.AEC
Imports Treegram.Bam.Libraries.AuditFile
Imports Treegram.Bam.Functions.AEC
Imports Treegram.Bam.Functions
Imports Treegram.GeomLibrary
Imports SharpDX

Public Class Check_WallAxisRvtScript
    Inherits ProdScript
    Public Sub New()
        MyBase.New("RVT :: Check : Walls have axis")
        AddAction(New Check_RvtWallAxis() With {.Name = Name, .PartOfScript = True})
    End Sub
End Class
Public Class Check_RvtWallAxis
    Inherits Check_WallAxis
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

    Public Overrides Sub CreateGuiOutputTree()
        'Présentation des résultats grâce à une zone de filtrage
        OutputTree = TempWorkspace.AddTree("Wall has axis ?")
        Dim wallNode = OutputTree.SmartAddNode("object.type", "Murs",,, "Murs")
        wallNode.SmartAddNode("Axis2dByTgm", "None")
        wallNode.SmartAddNode("Axis2dByTgm", "Polyline")
        wallNode.SmartAddNode("Axis2dByTgm", "TrimmedCurve")
    End Sub

End Class
