Imports Treegram.Bam.Libraries.AuditFile
Imports M4D.Production.Core
Imports M4D.Treegram.Core.Entities
Imports M4D.Treegram.Core.Kernel
Imports M4D.Treegram.Core.Extensions.Kernel
Imports M4D.Treegram.Core.Extensions.Entities

'TRAVAIL EN COURS
Public Class Create_RvtAuditInitial_Prodscript
    Inherits ProdScript
    Public Sub New()
        MyBase.New("RVT :: Audit : Initial")
        AddAction(New Create_RvtAuditInitial())
    End Sub
End Class
Public Class Create_RvtAuditInitial
    Inherits ProdAction
    Public Sub New()
        Name = "RVT :: Audit : Initial (ProdAction)"
        PartOfScript = True
    End Sub

    Public Overrides Sub CreateGuiInputTree()
        InputTree = TempWorkspace.AddTree("--- Temporary Input Tree ---")
        Dim fileNode As Node = InputTree.AddNode("object.type", "File")
        InputTree.SmartAddNode("object.type", "File").Description = "Files"

        SelectYourSetsOfInput.Add("FichiersRVT", {fileNode}.ToList())
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
                'POUR UN FILTRAGE DE TYPE INFORMATIQUE : WORKSPACE, FILE, METAOBJET
                If inputmObj.GetTgmType <> "File" OrElse inputmObj.GetAttribute("Extension").Value.ToString.ToLower <> "rvt" Then
                    Continue For 'Throw New Exception("You must drag and drop file objects")
                End If
                launchTree.SmartAddNode("object.id", inputmObj.Id.ToString,,, inputmObj.Name)
            Next
        End If
        Return launchTree
    End Function

    <ActionMethod()>
    Public Function ActionMethod(FichiersRVT As MultipleElements) As ActionResult

        If FichiersRVT.Source.Count = 0 Then
            Return New IgnoredActionResult("No input")
        End If
        ' Récupération des inputs
        Dim tgmFiles = GetInputAsMetaObjects(FichiersRVT.Source).ToHashSet 'To get rid of duplicates

        'AUDIT
        For Each fileTgm In tgmFiles

            If fileTgm.GetAttribute("Extension").Value.ToString.ToLower <> "rvt" Then
                Throw New Exception("this file is not a rvt File")
            End If
            Dim myAudit As New FileObject(fileTgm)
            'myAudit.PrepareFileWorkspacesAnalysis() <------------- Deja initialisé dans le scan

            CreateRvtInitialAUdit(myAudit)
            myAudit.SaveAuditWorkbook()
        Next

        Return New SucceededActionResult("Succeeded", System.Drawing.Color.Green, New List(Of Element) From {FichiersRVT})
    End Function

    Private Function CreateRvtInitialAUdit(myAudit As FileObject) As Boolean
        Dim myReport = myAudit.AuditWorkbook
        ''FILTER
        Workspace.GetTemp.SmartAddTree("ALL").Filter(False, New List(Of Workspace) From {myAudit.ScanWs}).RunSynchronously()

        'Levels <-------- A préparer
        'myAudit.ReportInfos.CompleteCriteria(52)
        'métriques du fichier
        myAudit.ReportInfos.FillActivity()

        Dim mySIze = CDbl(myAudit.FileTgm.GetAttribute("Size").Value)
        If Not IsNothing(myAudit.FileTgm?.GetAttribute("Size")) Then myAudit.ReportInfos.fileSizeInfo = Math.Round(CDbl(myAudit.FileTgm.GetAttribute("Size").Value), 1)
        myAudit.ReportInfos.CompleteCriteria(53)
        'File name -> Pas d'infos à aller chercher
        myAudit.ReportInfos.CompleteCriteria(63)
        'Emplacement fichier -> Pas d'infos à aller chercher
        myAudit.ReportInfos.CompleteCriteria(66)

        If Not IsNothing(myAudit.FileTgm?.GetAttribute("Name")) Then
            myAudit.ReportInfos.projectNameInfo = myAudit.FileTgm.GetAttribute("Name").Value
            myAudit.ReportInfos.projectNameBool = True
        Else
            myAudit.ReportInfos.projectNameBool = False
        End If
        myAudit.ReportInfos.CompleteInfo("projectName")

        If Not IsNothing(myAudit.FileTgm?.GetAttribute("ProjectInformations")?.GetAttribute("BuildingName")) Then
            myAudit.ReportInfos.buildingNameInfo = myAudit.FileTgm.GetAttribute("ProjectInformations").GetAttribute("BuildingName").Value
            myAudit.ReportInfos.buildingNameBool = True
        Else
            myAudit.ReportInfos.projectNameBool = False
        End If
        myAudit.ReportInfos.CompleteInfo("buildingName")

        If Not IsNothing(myAudit.FileTgm?.GetAttribute("ProjectInformations")?.GetAttribute("clientName")) Then
            myAudit.ReportInfos.buildingNameInfo = myAudit.FileTgm.GetAttribute("ProjectInformations").GetAttribute("clientName").Value
            myAudit.ReportInfos.buildingNameBool = True
        Else
            myAudit.ReportInfos.projectNameBool = False
        End If
        myAudit.ReportInfos.CompleteInfo("clientName")

        Dim groupReq = From obj In myAudit.ScanWs.MetaObjects Where obj.IsActive AndAlso Not IsNothing(obj?.GetAttribute("RevitCategory")) AndAlso obj.GetAttribute("RevitCategory").Value = "Groupes de modèles" AndAlso Not IsNothing(obj?.GetAttribute("object.containedIn")) AndAlso obj.GetAttribute("object.containedIn").Value = "3D"

        Dim grouptypeList As New List(Of String)
        For Each obj In groupReq
            Dim myType = obj.GetAttribute("RevitType").Value
            If Not grouptypeList.Contains(myType.replace("(membre exclu)", "")) Then
                grouptypeList.Add(myType.replace("(membre exclu)", ""))
            End If
        Next
        myAudit.ReportInfos.nbGroup3DInfo = groupReq.Count

        If groupReq.Count > 0 Then
            myAudit.ReportInfos.moy3DInfo = Math.Round(groupReq.Count / grouptypeList.Count, 1)
        Else
            myAudit.ReportInfos.moy3DInfo = 0.0
        End If
        myAudit.ReportInfos.CompleteInfo("nbGrp3D")
        myAudit.ReportInfos.CompleteInfo("moyGrp3D")
        myAudit.ReportInfos.CompleteCriteria(57)

        Dim geneReq = From obj In myAudit.ScanWs.MetaObjects Where obj.IsActive AndAlso Not IsNothing(obj?.GetAttribute("RevitCategory")) AndAlso obj.GetAttribute("RevitCategory").Value = "Modèle générique" AndAlso Not IsNothing(obj?.GetAttribute("object.containedIn")) AndAlso obj.GetAttribute("object.containedIn").Value = "3D"


        myAudit.SaveAuditWorkbook()
        Return True
    End Function

End Class
