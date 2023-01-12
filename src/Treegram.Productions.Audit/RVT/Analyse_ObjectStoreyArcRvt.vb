Imports M4D.Production.Core
Imports M4D.Treegram.Core.Entities
Imports M4D.Treegram.Core.Kernel
Imports M4D.Treegram.Dashboard.Production
Imports SimpleBetaTgmFcts

Public Class Analyse_ObjectStoreyArcRvt
    Inherits AdvancedProdAction
    Public Sub New()
        Name = "RVT :: AUDIT : Storey Analysis ARC (CQ-BIM)"
    End Sub

    Public Overrides Sub CreateGuiInputTree()
        InputTree = TempWorkspace.AddTree("--- Temporary Input Tree ---")
        Dim objectNode = InputTree.SmartAddNode("object.containedIn", "3D", Nothing, Nothing, "Objets 3D")

        '---REVIT objects
        Dim revNode = objectNode.SmartAddNode("RevitCategory", Nothing, Nothing, Nothing, "Objets REVIT")
        Dim categList As New List(Of String) From {"Accessoire de canalisation", "Accessoire de gaine", "Appareils d'appel malade", "Appareils de communication", "Appareils sanitaires", "Appareils téléphoniques", "Armature à béton", "Armature surfacique", "Armature surfacique (treillis)", "Bouche d'aération", "Canalisation", "Canalisation souple", "Cheminement de fabrication MEP", "Chemins de câbles", "Conduits", "Connexions structurelles", "Coupleur d'armature structurelle", "Dalles structurelles", "Direction principale du ferraillage", "Dispositifs d'alarme incendie", "Dispositifs de données", "Dispositifs de sécurité", "Dispositifs d'éclairage", "Eléments", "Eléments de détail", "Environnement", "Equipement de génie climatique", "Equipement électrique", "Equipement spécialisé", "Escalier", "Espaces", "Espaces réservés aux canalisations", "Espaces réservés aux gaines", "Fenêtres", "Fils", "Fondations", "Gaine", "Gaine flexible", "Garde-corps", "Installations électriques", "Isolations des canalisations", "Isolations des gaines", "Lignes", "Luminaires", "Meneaux de murs-rideaux", "Meubles de rangement", "Mobilier", "Modèles génériques", "Murs", "Murs-rideaux", "Ossatures", "Ouvertures de cage", "Panneaux de murs-rideaux", "Parking", "Pièces", "Plafonds", "Plantes", "Portes", "Poteaux", "Poteaux porteurs", "Poutres à treillis", "Raccords de canalisation", "Raccords de chemins de câbles", "Raccords de conduits", "Raccords de gaine", "Raidisseurs", "Rampes d'accès", "Réseau de canalisations de fabrication MEP", "Réseau de gaines de fabrication MEP", "Réseaux de poutres", "Revêtements des gaines", "Routes", "Site", "Sols", "Sprinklers", "Surfaces", "Systèmes de mobilier", "Systèmes de murs-rideaux", "Tirants de fabrication MEP", "Toits", "Topographie", "Volumes", "Zones HVAC"}
        categList.Sort()

        Dim categNodeList As New List(Of Node)

        For Each categSt In categList
            categNodeList.Add(revNode.SmartAddNode("RevitCategory", categSt))
        Next
        '---REF ETAGES
        Dim etageNode = InputTree.SmartAddNode("object.type", "FileLevel", Nothing, Nothing, "Etages")
        Dim storeyNode = InputTree.SmartAddNode("object.type", "TgmBuildingStorey", Nothing, Nothing, "Etages References")
        storeyNode.SmartAddNode("IsUsed", False,,, "Non valides")

        SelectYourSetsOfInput.Add("Storeys", {etageNode}.ToList())
        SelectYourSetsOfInput.Add("Objects3D", categNodeList)

    End Sub

    Public Overrides Function GuiUpdateLauncherTree(inputs As Dictionary(Of String, IEnumerable(Of PersistEntity))) As Tree

        Dim launchTree = TempWorkspace.SmartAddTree("LauncherTree")

        ' Suppression des anciens noeuds
        While launchTree.Nodes.Count > 0
            launchTree.RemoveNode(launchTree.Nodes.First)
        End While

        '---Complete with inputs
        Dim containersList As New List(Of Workspace)
        If inputs.ContainsKey("Objects3D") AndAlso inputs("Objects3D").ToList.Count > 0 Then
            ''Get first File founded - To discuss...
            'Dim firstObj As MetaObject = inputs("Objects3D")(0)
            'Dim scanWs As Workspace = firstObj.Container
            'Dim fileTgm As MetaObject = scanWs.GetMetaObjects(, "File").First
            'Dim fileNode = launchTree.SmartAddNode("File", fileTgm.Name)

            'Try looping on each object
            For Each obj As MetaObject In inputs("Objects3D")
                If Not containersList.Contains(CType(obj.Container, Workspace)) Then
                    containersList.Add(CType(obj.Container, Workspace))
                End If
            Next
        End If
        For Each containerWs In containersList
            Dim fileTgm As MetaObject = containerWs.GetMetaObjects(, "File").First
            Dim fileNode = launchTree.SmartAddNode("File", fileTgm.Name)
        Next
        Return launchTree
    End Function

    <ActionMethod(Scope:=MethodScope.OnNew Or MethodScope.OnUpdate)>
    Public Function MyMethod(Objects3D As MultipleElements, Storeys As MultipleElements) As ActionResult

        'GET INPUTS
        Dim objs3dList = CType(Objects3D.Source, IEnumerable(Of PersistEntity)).ToHashSet 'To get rid of duplicates
        Dim storeysList = CType(Storeys.Source, IEnumerable(Of PersistEntity)).ToHashSet 'To get rid of duplicates
        If objs3dList.Count = 0 Or storeysList.Count = 0 Then Throw New Exception("Inputs missing")
        Dim tgmObjsList As List(Of MetaObject) = objs3dList.Cast(Of MetaObject).ToList
        Dim tgmStoreysList As List(Of MetaObject) = storeysList.Cast(Of MetaObject).ToList
        If tgmObjsList.Count = 0 Or tgmStoreysList.Count = 0 Then Throw New Exception("Inputs missing")

        'AUDIT
        OutputTree = Audit_Library.Core.GetStoreyAnalysis(tgmStoreysList, tgmObjsList)

        Return Nothing
    End Function

End Class
