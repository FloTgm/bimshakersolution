Imports Treegram.Bam.Libraries.AuditFile
Imports M4D.Production.Core
Imports M4D.Treegram.Core.Entities
Imports M4D.Treegram.Core.Kernel
Imports M4D.Treegram.Core.Extensions.Kernel
Imports M4D.Treegram.Core.Extensions.Entities

Public Class Check_ParameterExistsScript
    Inherits ProdScript
    Public Sub New()
        MyBase.New("RVT :: Check : Reference Parameters")
        AddAction(New Check_ParameterExists())
    End Sub
End Class
Public Class Check_ParameterExists
    Inherits ProdAction
    Public Sub New()
        Name = "RVT :: Check : Reference Parameters"
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
        Dim RevitFileTgmObjs = GetInputAsMetaObjects(FichiersRVT.Source).ToHashSet 'To get rid of duplicates

        OutputTree = CheckParamReference(RevitFileTgmObjs(0))

        Return New SucceededActionResult("Succeeded", System.Drawing.Color.Green, New List(Of Element) From {FichiersRVT})
    End Function

    Private Function CheckParamReference(RevitFileTgmObj As MetaObject) As Tree
        Dim myAudit As New IfcFileObject(RevitFileTgmObj)

        Dim settingWb = myAudit.SettingsWorkbook
        Dim refSheet = settingWb.Worksheets.Item("VALEURS DE REFERENCE")
        If IsNothing(refSheet) Then

        End If



        Return Nothing
    End Function

End Class
