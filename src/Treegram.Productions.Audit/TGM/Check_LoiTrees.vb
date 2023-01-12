Imports M4D.Production.Core
Imports M4D.Treegram.Core.Entities
Imports M4D.Treegram.Core.Kernel
Imports M4D.Treegram.Core.Extensions.Kernel
Imports M4D.Treegram.Core.Extensions.Entities
Imports Treegram.Bam.Libraries.AuditFile
Imports Tgm.Deixi.Services
Imports M4D.Treegram.Core.Extensions.Models

Public Class Check_LoiTreesScript
    Inherits ProdScript
    Public Sub New()
        MyBase.New("TGM :: Check : LOI_Niveau d'Information")
        AddAction(New Check_LoiTrees())
    End Sub
End Class
Public Class Check_LoiTrees
    Inherits ProdAction
    Public Sub New()
        Name = "TGM :: Check : LOI Trees (ProdAction)"
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

        ' Récupération des inputs
        Dim tgmFiles = GetInputAsMetaObjects(Fichiers.Source).ToHashSet 'To get rid of duplicates

        'CHECK
        CheckLoiTrees(tgmFiles(0))

#Region "VIEW CREATION"
        Dim scanWs As Workspace = tgmFiles(0).Container
        Dim projWs As Workspace = tgmFiles(0).GetProject

        'Récuperation de l'arbre LOI_Protection incendie file
        If projWs.Trees.Where(Function(o) o.Name = "LOI_Protection Incendie").FirstOrDefault IsNot Nothing Then
            Dim chargedWs As New List(Of Workspace) From {scanWs}
            Dim checkedNodes As New List(Of Node)
            'Dim savedColors As New Dictionary(Of Node, Tuple(Of System.Windows.Media.Color, Single))
            Dim colorNodes As New List(Of NodeColorTransparency)

            Dim myTree = projWs.Trees.Where(Function(o) o.Name = "LOI_Protection Incendie").FirstOrDefault
            For Each catNode In myTree.Nodes.ToList
                checkedNodes.Add(catNode)
                'savedColors.Add(catNode, New Tuple(Of System.Windows.Media.Color, Single)(System.Windows.Media.Color.FromRgb(154, 204, 255), 0.5F))
                colorNodes.Add(New NodeColorTransparency(catNode, System.Windows.Media.Color.FromRgb(154, 204, 255), 0.5F))

                checkedNodes.Add(catNode.Nodes(0))
                'savedColors.Add(catNode.Nodes(0), New Tuple(Of System.Windows.Media.Color, Single)(System.Windows.Media.Color.FromRgb(229, 117, 114), 0.5F))
                colorNodes.Add(New NodeColorTransparency(catNode, System.Windows.Media.Color.FromRgb(229, 117, 114), 0.5F))
            Next

            TgmView.CreateView(projWs, tgmFiles(0).Name & " - LOI_Protection Incendie", chargedWs, myTree, checkedNodes, colorNodes)
        End If
#End Region

        'Dim outputs = New List(Of Element) From {New MultipleElements(Fichiers, "Fichiers", Nothing, ElementState.[New])}

        Return New SucceededActionResult("Succeeded", New List(Of Element) From {Fichiers})
    End Function

    Public Shared Function CheckLoiTrees(fileTgm As MetaObject) As Tree

        Dim myAudit As New IfcFileObject(fileTgm)
        Dim auditWb = myAudit.AuditWorkbook
        If auditWb Is Nothing Then Throw New Exception("Launch initial audit first")
        Dim repSheet = myAudit.ReportWorksheet

        'get TagInfo
        Dim fileWs = fileTgm.Container
        Dim myTagAtt = fileTgm?.GetAttribute("BimShakerSettings").GetAttribute("Trade")
        Dim tagList As New List(Of String)
        Dim tagExist = False
        If Not IsNothing(myTagAtt) Then
            tagExist = True
        Else

        End If

        'Get dico ws
        Dim dicoWs As Workspace = myAudit.ProjWs.GetWorkspaces("Dictionary").FirstOrDefault
        If dicoWs IsNot Nothing Then
            For Each dicoTree In dicoWs.Trees
                Dim treeName As String = dicoTree.Name.Replace(" ", "")
                If treeName.Split("-")(0).ToLower = "loi" OrElse treeName.Split("_")(0).ToLower = "loi" Then
                    If myAudit.ProjWs.Trees.FirstOrDefault(Function(o) o.Name = dicoTree.Name) Is Nothing Then
                        dicoTree.CopyTree(myAudit.ProjWs)
                    End If
                End If
            Next
        End If

        'Get loi trees
        Dim loiTrees As New List(Of Tree)
        For Each myTree In myAudit.ProjWs.Trees
            If Not myTree.IsActive Then
                Continue For
            End If

            'If tagExist Then
            '    Dim tagSplit As String() = myTagAtt.Value.split(";")
            '    For Each myTag In tagSplit
            '        If Not myTree.Name.Contains("_" + myTag) Then
            '        End If
            '    Next
            'End If

            Dim treeName As String = myTree.Name.Replace(" ", "")
            If treeName.Split("-")(0).ToLower = "loi" OrElse treeName.Split("_")(0).ToLower = "loi" Then
                If tagExist AndAlso myTagAtt.Value <> "" Then
                    Dim tagSplit As String() = myTagAtt.Value.split(";")
                    For Each myTag In tagSplit
                        If myTree.Name.Contains("_" + myTag + "_") Then
                            loiTrees.Add(myTree)
                            Exit For
                        End If
                    Next
                Else
                    loiTrees.Add(myTree)
                End If
            End If
        Next
        If loiTrees.Count = 0 Then
            Return Nothing
        End If


        'Check for each tree falling objects
        For Each myTree In loiTrees
            myTree.Filter(myAudit._3dObjects)
            myAudit.ReportInfos.LoiCriteria_LoiTreeName = myTree.Name
            Dim mainNodes = myTree.Nodes
            Dim totalCount As Integer = 0

            If mainNodes.Count > 0 Then
                For Each node In mainNodes
                    Dim blkCount As Integer = node.BlockedEntities.Count
                    Dim filteredCount As Integer = node.FilteredEntities.Count
                    totalCount = totalCount + blkCount + (filteredCount - blkCount)
                Next
            Else

            End If
            'Check falling objects
            myAudit.ReportInfos.LoiCriteria_ObjAllFallen = True
            Dim blockedObjCount As Integer = 0
            For Each node In myTree.AllNodes
                If node.Nodes.Count <> 0 AndAlso node.BlockedEntities.Count > 0 Then
                    blockedObjCount = blockedObjCount + node.BlockedEntities.Count
                    myAudit.ReportInfos.LoiCriteria_ObjAllFallen = False
                    'Exit For
                End If
            Next
            Dim LOIseulAtt = myTree?.GetAttribute("LOI_Seuil")

            myAudit.ReportInfos.LoiCriteria_NoObj = False
            If totalCount > 0 Then
                myAudit.ReportInfos.LoiCriteria_LoiPercentage = Math.Floor((((totalCount - blockedObjCount) / totalCount) * 100))
                Dim myStat = Math.Floor((((totalCount - blockedObjCount) / totalCount) * 100)).ToString + "% (" + (totalCount - blockedObjCount).ToString + "/" + totalCount.ToString + ")"
                myAudit.ReportInfos.LoiCriteria_LoiPercentageStr = myStat

                If IsNothing(LOIseulAtt) OrElse LOIseulAtt.Value.ToString = "" Then
                    myAudit.ReportInfos.LoiCriteria_Seuil = 100
                Else
                    Try
                        myAudit.ReportInfos.LoiCriteria_Seuil = CDbl(LOIseulAtt.Value)
                    Catch ex As Exception
                        myAudit.ReportInfos.LoiCriteria_Seuil = 100
                    End Try
                End If
            Else
                myAudit.ReportInfos.LoiCriteria_LoiPercentageStr = "0% (0/0)"
                myAudit.ReportInfos.LoiCriteria_LoiPercentage = 0
                myAudit.ReportInfos.LoiCriteria_ObjAllFallen = False
                myAudit.ReportInfos.LoiCriteria_NoObj = True
            End If
            'Fill report
            myAudit.ReportInfos.CompleteCriteria(64)
        Next

        myAudit.ProjWs.PushAllModifiedEntities()
        myAudit.AuditWorkbook.Calculate()
        myAudit.SaveAuditWorkbook()

        Return Nothing

    End Function
End Class
