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

Public Class Check_WallAxisScript
    Inherits ProdScript
    Public Sub New()
        MyBase.New("IFC :: Check : Walls have axis")
        AddAction(New Check_WallAxis() With {.Name = Name, .PartOfScript = True})
    End Sub
End Class
Public Class Check_WallAxis
    Inherits ProdAction
    Public Sub New()
    End Sub

    Public Overrides Sub CreateGuiInputTree()
        'Création d'une zone de filtrage (arbre) temporaire
        InputTree = TempWorkspace.AddTree("--- Temporary Input Tree ---")

        'Création d'un filtre : murs
        Dim wallNode = InputTree.SmartAddNode("object.type", "IfcWall",,, "Murs")

        'Drag and Drop automatique des objets filtrés vers l'algo (Optionnel)
        SelectYourSetsOfInput.Add("Murs", {wallNode}.ToList())
    End Sub


    Public Overrides Function GuiUpdateLauncherTree(inputs As Dictionary(Of String, IEnumerable(Of PersistEntity))) As Tree
        'Séparation des lancements par fichier
        Dim launchTree As Tree = TempWorkspace.SmartAddTree("LauncherTree")
        CommonAecLauncherTrees.FilesLaunchTree(launchTree, inputs)
        Return launchTree
    End Function


    <ActionMethod()>
    Public Function MyMethod(Murs As MultipleElements) As ActionResult

        'PREPARATION ENVIRONNEMENT TGM
        '--Sortie de l'algo s'il n'y a pas d'inputs
        If GetInputAsMetaObjects(Murs.Source).Count = 0 Then
            Return New IgnoredActionResult("No input")
        End If

        '--Récupération des métaobjets murs sous formé de liste
        Dim tgmObjsList = GetInputAsMetaObjects(Murs.Source).ToHashSet.ToList

        '--Récupération du fichier source
        Dim myFile = CommonAecFunctions.GetAuditFileFromInputs(tgmObjsList)

        '--Chargement des données BAMs
        Dim wsList = New List(Of Workspace) From {myFile.ScanWs}
        wsList.AddRange(myFile.ScanWs.GetExtensionWorspaces(M4D.Treegram.Core.Extensions.Enums.ExtensionType.AdditionalGeometries).ToList)
        CommonFunctions.RunFilter(TempWorkspace, wsList)


        'ANALYSE MÉTIER
        Dim totalWallsNb = tgmObjsList.Count
        Dim noAxisWallsNb = 0
        For Each wall In tgmObjsList
            Dim aecWall As New AecWall(wall)

            Dim axis = GetAxisFromScan(aecWall, myFile)

            If axis Is Nothing Then
                noAxisWallsNb += 1
                aecWall.CompleteTgmPset("Axis2dByTgm", "None")
            Else
                aecWall.CompleteTgmPset("Axis2dByTgm", "Polyline")
            End If
        Next
        myFile.ProjWs.PushAllModifiedEntities()


        'BIM-SHAKER : Complète le critère dans le rapport d'audit
        Dim visa As ReportInfos.visa
        Dim comment As String
        If noAxisWallsNb = 0 Then
            visa = ReportInfos.visa.OK
            comment = "Tous les murs ont un axe"
        Else
            visa = ReportInfos.visa.REV
            comment = $"{noAxisWallsNb} murs n'ont pas d'axe"
        End If
        myFile.ReportInfos.CompleteCriteria("Tgm_AxisCriteria", visa, comment)


        CreateGuiOutputTree()
        Return New SucceededActionResult("Launched", System.Drawing.Color.Green, New List(Of Element) From {Murs})
    End Function

    Private Shared Function GetAxisFromScan(aecWall As AecWall, myFile As FileObject) As Curve2D
        Dim axis2d As Curve2D = Nothing
        If myFile.IsIfc Then
            If aecWall.AxisExtension IsNot Nothing Then
                Dim curveAtt = aecWall.AxisExtension.GetAttribute("Curve2D", True)
                If curveAtt IsNot Nothing Then
                    Dim axisPoints = New List(Of Vector2)
                    For Each pointAtt In curveAtt.Attributes
                        axisPoints.Add(New SharpDX.Vector2(CSng(pointAtt.GetAttribute("X").Value), CSng(pointAtt.GetAttribute("Y").Value)))
                    Next
                    axis2d = New Curve2D(axisPoints)
                End If
            End If

        ElseIf myFile.IsRevit Then
            If aecWall.AdditionalGeometryExtension IsNot Nothing Then
                Dim curveAtt = aecWall.AdditionalGeometryExtension.GetAttribute("Line", True)
                If curveAtt IsNot Nothing Then
                    Dim axisPoints = New List(Of Vector2)
                    For Each pointAtt In curveAtt.Attributes
                        axisPoints.Add(New SharpDX.Vector2(CSng(pointAtt.GetAttribute("X").Value), CSng(pointAtt.GetAttribute("Y").Value)))
                    Next
                    axis2d = New Curve2D(axisPoints)
                End If
            End If
        Else
            Throw New Exception("This type of file is not implemented yet")
        End If
        Return axis2d
    End Function

    Public Overrides Sub CreateGuiOutputTree()
        'Présentation des résultats grâce à une zone de filtrage
        OutputTree = TempWorkspace.AddTree("Wall has axis ?")
        Dim wallNode = OutputTree.SmartAddNode("object.type", "IfcWall",,, "Murs")
        wallNode.SmartAddNode("Axis2dByTgm", "None")
        wallNode.SmartAddNode("Axis2dByTgm", "Polyline")
        wallNode.SmartAddNode("Axis2dByTgm", "TrimmedCurve")
    End Sub

End Class
