Imports Treegram.Bam.Libraries.AuditFile
Imports M4D.Production.Core
Imports M4D.Treegram.Core.Entities
Imports M4D.Treegram.Core.Kernel
Imports M4D.Treegram.Core.Extensions.Kernel
Imports M4D.Treegram.Core.Extensions.Entities

Public Class Create_FamilyTreeScript
    Inherits ProdScript
    Public Sub New()
        MyBase.New("RVT :: Create : Family Tree")
        AddAction(New Create_FamilyTree())
    End Sub
End Class
Public Class Create_FamilyTree
    Inherits ProdAction
    Public Sub New()
        Name = "RVT :: Create : Family Tree (ProdAction)"
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

        OutputTree = CreateRvtFamilyTree(RevitFileTgmObjs(0))

        Return New SucceededActionResult("Succeeded", System.Drawing.Color.Green, New List(Of Element) From {FichiersRVT})
    End Function

    Private Function CreateRvtFamilyTree(RevitFileTgmObj As MetaObject) As Tree
        Dim myAudit As New IfcFileObject(RevitFileTgmObj)

#Region "GET INFOS"
        ' Metaobject > Category > Famille > Type
        Dim dictionary = New Dictionary(Of Tuple(Of String, String), Dictionary(Of Tuple(Of String, String), List(Of Tuple(Of String, String))))
        Dim noCategoryTuple = New Tuple(Of String, String)("No Category", String.Empty)
        Dim noFamilyTuple = New Tuple(Of String, String)("No Family", String.Empty)

        ' Recherche si fichier
        Dim at = RevitFileTgmObj.GetAttribute("Extension", True)
        If at Is Nothing OrElse Not String.Equals(at.Value.ToString().ToLower(), "rvt") Then
            Throw New Exception("Revit scan error")
        End If

        ' Fichier de type revit ?
        Dim revitTypeMo = RevitFileTgmObj.Relations.FirstOrDefault(Function(r) String.Equals(r.Target.Name, "RevitType"))?.Target
        If revitTypeMo Is Nothing Then
            Throw New Exception("Revit scan error")
        End If

        'Dim moTuple = New Tuple(Of String, String)(CType(RevitFileTgmObj.Templates.First(), MetaObject).Name, RevitFileTgmObj.Name)
        'dictionary.Add(moTuple, New Dictionary(Of Tuple(Of String, String), Dictionary(Of Tuple(Of String, String), List(Of Tuple(Of String, String))))())

        For Each typeRelation In revitTypeMo.Relations
            'On prend le mo contenant le nom du type
            Dim moType = typeRelation.Target

            'On cherche son template avec une relation avec un parent le nom est RevitFamily (peut ne pas exister)
            Dim moFamily = moType.Templates.Cast(Of MetaObject)().FirstOrDefault(Function(moTemplate) moTemplate.Parents.Any(Function(relation) String.Equals(relation.Parent.Name, "RevitFamily")))

            'Le mo où l'on cherche la catégorie
            Dim moFindCategory = If(moFamily, moType)

            'On cherche dans ce mo la relation avec un parent dont le nom est RevitCategory
            Dim moCategory = moFindCategory.Templates.Cast(Of MetaObject)().FirstOrDefault(Function(moTemplate) moTemplate.Parents.Any(Function(relation) String.Equals(relation.Parent.Name, "RevitCategory")))

            'Y'a plus qu'à prendre les noms
            Dim type = New Tuple(Of String, String)("RevitType", moType.Name)
            Dim family = If(moFamily IsNot Nothing, New Tuple(Of String, String)("RevitFamily", moFamily.Name), If(moType.GetAttribute("RevitFamily") IsNot Nothing, New Tuple(Of String, String)("RevitFamilySystem", moType.GetAttribute("RevitFamily").Value.ToString), noFamilyTuple))
            Dim category = If(moCategory IsNot Nothing, New Tuple(Of String, String)("RevitCategory", moCategory.Name), noCategoryTuple)

            ' Construction du dictionnaire
            Dim element = dictionary
            If Not element.ContainsKey(category) Then
                element.Add(category, New Dictionary(Of Tuple(Of String, String), List(Of Tuple(Of String, String)))())
                element(category).Add(family, New List(Of Tuple(Of String, String))())
                element(category)(family).Add(type)
            Else
                If Not element(category).ContainsKey(family) Then
                    element(category).Add(family, New List(Of Tuple(Of String, String))())
                    element(category)(family).Add(type)
                Else
                    element(category)(family).Add(type)
                End If
            End If
        Next

        'Ajout des Catégories sans types
        Dim revitCategoryMo = RevitFileTgmObj.Relations.FirstOrDefault(Function(r) String.Equals(r.Target.Name, "RevitCategory"))?.Target
        If dictionary.Count = revitCategoryMo.Relations.Count Then
            'Throw New Exception("Revit scan error")
        End If
        For Each categoryRelation In revitCategoryMo.Relations
            Dim moCategory = categoryRelation.Target
            Dim category = New Tuple(Of String, String)("RevitCategory", moCategory.Name)
            If Not dictionary.ContainsKey(New Tuple(Of String, String)("RevitCategory", moCategory.Name)) Then
                dictionary.Add(category, New Dictionary(Of Tuple(Of String, String), List(Of Tuple(Of String, String)))())
            End If
        Next

        'Sort categories
        Dim sortedDict = dictionary.OrderBy(Function(o) o.Key.Item2).ToDictionary(Function(pair) pair.Key, Function(pair) pair.Value)
#End Region

#Region "CREATE TREE"
        Dim myTree = myAudit.AuditTreesWs.SmartAddTree("Audit - CATEGORISATION")
        Dim fileNode = myTree.SmartAddNode("File", myAudit.FileTgm.Name)
        myTree.RemoveNode(fileNode)
        fileNode = myTree.SmartAddNode("File", myAudit.FileTgm.Name)
        myTree.DuplicateObjectsWhileFiltering = False

        For Each categoryDictionaryElement In sortedDict
            Dim catNode As Node
            If (categoryDictionaryElement.Key.Equals(noCategoryTuple)) Then
                'catNode = moNode.AddNode(Nothing, "Catégorie")
                'catNode.Operator = NodeOperator.Different
                'catNode.Description = "No category"
                Continue For
            Else
                catNode = fileNode.AddNode(categoryDictionaryElement.Key.Item1, categoryDictionaryElement.Key.Item2)
                catNode.Description = "Category : " + categoryDictionaryElement.Key.Item2
            End If

            For Each familyDictionaryElement In categoryDictionaryElement.Value
                Dim famNode As Node
                If (familyDictionaryElement.Key.Equals(noFamilyTuple)) Then
                    'famNode = catNode.AddNode(Nothing, "Famille")
                    'famNode.Operator = NodeOperator.Different
                    'famNode.Description = "No family"
                    Continue For
                Else
                    If familyDictionaryElement.Key.Item1.Contains("System") Then
                        famNode = catNode.AddNode("RevitFamily", familyDictionaryElement.Key.Item2)
                        famNode.Description = "FamilySystem : " + familyDictionaryElement.Key.Item2
                    Else
                        famNode = catNode.AddNode(familyDictionaryElement.Key.Item1, familyDictionaryElement.Key.Item2)
                        famNode.Description = "Family : " + familyDictionaryElement.Key.Item2
                    End If
                End If

                For Each typeDictionaryElement In familyDictionaryElement.Value
                    Dim typeNode = famNode.AddNode(typeDictionaryElement.Item1, typeDictionaryElement.Item2)
                    typeNode.Description = "Type : " + typeDictionaryElement.Item2
                Next
            Next

            If categoryDictionaryElement.Value.ContainsKey(noFamilyTuple) Then
                Dim nofamNode = catNode.AddNode("RevitType", Nothing)
                nofamNode.Description = "No family"

                For Each type In categoryDictionaryElement.Value(noFamilyTuple)
                    Dim typeNode = nofamNode.AddNode(type.Item1, type.Item2)
                    typeNode.Description = "Type : " + type.Item2
                Next
            End If
        Next

        If sortedDict.ContainsKey(noCategoryTuple) Then
            Dim noCatNode = fileNode.AddNode("RevitType", Nothing)
            noCatNode.Description = "No category"

            For Each type In sortedDict(noCategoryTuple)
                Dim typeNode = noCatNode.AddNode(type.Key.Item1, type.Key.Item2)
                typeNode.Description = "Type : " + type.Key.Item2
            Next
        End If

        ''Crée l'arbre définitif
        'If Save_The_Resulting_Tree Then
        '    Dim inputList As New List(Of MetaObject)
        '    For Each imputedObj In InputObjects
        '        Dim mObj As MetaObject = imputedObj
        '        inputList.Add(mObj)
        '    Next
        '    Dim projWs As Workspace = inputList(0).GetProject
        '    Dim resultTree As Tree = projWs.SmartAddTree("Revit Family Tree")
        '    resultTree.DuplicateObjectsWhileFiltering = False
        '    For Each StructureRevitFile In dictionary
        '        Dim FileNodeDefinitiveTree = resultTree.SmartAddNode(StructureRevitFile.Key.Item1, StructureRevitFile.Key.Item2)
        '        'Vide d'abord l'arbre
        '        While FileNodeDefinitiveTree.Nodes.Count > 0
        '            FileNodeDefinitiveTree.RemoveNode(FileNodeDefinitiveTree.Nodes(0))
        '        End While
        '        'Copie l'arbre temporaire
        '        Dim fileNodeOfOutputTree = OutputTree.SmartAddNode(StructureRevitFile.Key.Item1, StructureRevitFile.Key.Item2)
        '        For Each node In fileNodeOfOutputTree.Nodes
        '            FileNodeDefinitiveTree.CloneNode(node, True, True)
        '        Next
        '    Next
        'End If
#End Region

        Return myTree
    End Function

End Class
