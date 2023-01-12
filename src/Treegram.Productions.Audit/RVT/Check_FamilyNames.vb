Imports M4D.Production.Core
Imports M4D.Treegram.Core.Entities
Imports M4D.Treegram.Core.Kernel
Imports M4D.Treegram.Core.Extensions.Kernel
Imports M4D.Treegram.Core.Extensions.Entities
Imports Treegram.Bam.Libraries.AuditFile
Imports DevExpress.Spreadsheet
Public Class CheckFamilyNameProdScript
    Inherits ProdScript
    Public Sub New()
        MyBase.New("RVT :: Check : Family Names")
        AddAction(New CheckFamilyName())
    End Sub
End Class
Public Class CheckFamilyName
    Inherits ProdAction
    Public Sub New()
        Name = "RVT :: Check : Family Names (ProdAction)"
        PartOfScript = True
    End Sub

    ''' <summary>
    ''' Méthode appelée lors de la sélection de l'action
    ''' Elle n'est pas obligatoire
    ''' </summary>
    Public Overrides Sub OnActionSelected()

    End Sub

    ''' <summary>
    ''' Si jamais vous avez besoin de créer un arbre input, vous pouvez utiliser cette méthode.
    ''' Il n'est pas obligatoire de la rajouter
    ''' </summary>
    Public Overrides Sub CreateGuiInputTree()
        InputTree = TempWorkspace.AddTree("--- Temporary Input Tree ---")
        Dim fileNode As Node = InputTree.AddNode("object.type", "File")
        'InputTree.AddNode("object.type", "File").Description = "Files"

        Dim mandatName As String = ""
        SelectYourSetsOfInput.Add("FichiersRevit", {fileNode}.ToList())
        'SelectYourSetsOfInput.Add("MandatoryName", {mandatName}.ToList())

    End Sub

    ''' <summary>
    ''' Initialise l'Output Tree
    ''' Pas obligatoire non plus
    ''' </summary>
    Public Overrides Sub CreateGuiOutputTree()
        'OutputTree = TempWorkspace.AddTree("Arbre output temporaire")
    End Sub

    ''' <param name="inputs">La collection d'un paramètre et de ses PersistEntity(pour le moment uniquement des metaObjects</param>
    ''' <returns></returns>
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
                If inputmObj.GetTgmType <> "File" OrElse inputmObj.GetAttribute("Extension").Value.ToString.ToLower <> "rvt" Then Continue For 'Throw New Exception("You must drag and drop file objects")
                launchTree.SmartAddNode("object.id", inputmObj.Id.ToString,,, inputmObj.Name)
            Next
        End If

        Return launchTree
    End Function
    <ActionMethod()>
    Public Function MyMethod(FichiersRevit As MultipleElements) As ActionResult
        ' Récupération des inputs
        If FichiersRevit.Source.Count = 0 Then
            Return New IgnoredActionResult("No input")
        End If

        Dim tgmFiles = GetInputAsMetaObjects(FichiersRevit.Source).ToHashSet 'To get rid of duplicates

        Dim sortedDict As New Dictionary(Of Tuple(Of String, String), Dictionary(Of Tuple(Of String, String), List(Of Tuple(Of String, String))))
        'AUDIT
        For Each fileTgm In tgmFiles
            OutputTree = Complete3DNames(fileTgm, sortedDict)
        Next

        Return New SucceededActionResult("Succeeded", System.Drawing.Color.Green, New List(Of Element) From {FichiersRevit})

    End Function

    Private Function Complete3DNames(RevitFileTgmObj As MetaObject, ByRef sortedDict As Dictionary(Of Tuple(Of String, String), Dictionary(Of Tuple(Of String, String), List(Of Tuple(Of String, String))))) As Tree

        Dim myAudit As New IfcFileObject(RevitFileTgmObj)
        Dim auditWb = myAudit.AuditWorkbook
        If auditWb Is Nothing Then Throw New Exception("Launch initial audit first")

        'Noms interdits
        Dim forbiddenStrList As New List(Of String)
        forbiddenStrList.Add("famille")
        forbiddenStrList.Add("family")
        forbiddenStrList.Add("type")

        'Noms obligatoires
        Dim mandatoryBool = True
        Dim mandatoryList As New List(Of String)

#Region "GET INFOS"

        Dim auditWsheet As DevExpress.Spreadsheet.Worksheet = auditWb.Worksheets.Item("FAMILLES")
        Dim famCol As Integer = 2
        Dim typeCol As Integer = 3
        Dim lastL As Integer = 0
        Dim typeColor = auditWsheet.Cells(7, typeCol).FillColor
        Dim famColor = auditWsheet.Cells(6, typeCol).FillColor
        'Dim famColor = auditWsheet.Cells(3 + lastLine, typeCol).FillColor
        Dim categColor = auditWsheet.Cells(5, typeCol).FillColor

        While auditWsheet.Cells(3 + lastL, typeCol).Value.ToString <> ""
            lastL = lastL + 1
        End While

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

        For k = 5 To lastL + 3
            'On prend le mo contenant le nom du type
            Dim myCell = auditWsheet.Cells(k, typeCol)
            If myCell.FillColor = typeColor Then
                Dim moType = auditWsheet.Cells(k, typeCol).Value.ToString
                Dim n As Integer = 0
                'On cherche son template avec une relation avec un parent le nom est RevitFamily (peut ne pas exister)
                While auditWsheet.Cells(k - n, typeCol).FillColor <> famColor
                    n += 1
                End While
                Dim moFamily = auditWsheet.Cells(k - n, famCol).Value.ToString

                Dim m As Integer = 0
                While auditWsheet.Cells(k - m, typeCol).FillColor <> categColor
                    m += 1
                End While
                Dim moCategory = auditWsheet.Cells(k - m, 1).Value.ToString

                'Y'a plus qu'à prendre les noms
                Dim type = New Tuple(Of String, String)("RevitType", moType)
                Dim family As New Tuple(Of String, String)("RevitFamily", moFamily)
                Dim category As New Tuple(Of String, String)("RevitCategory", moCategory)

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
            End If
        Next

        'Ajout des Catégories sans types
        'Dim revitCategoryMo = RevitFileTgmObj.Relations.FirstOrDefault(Function(r) String.Equals(r.Target.Name, "RevitCategory"))?.Target
        ''If dictionary.Count = revitCategoryMo.Relations.Count Then
        ''    Throw New Exception("Revit scan error")
        ''End If
        'For Each categoryRelation In revitCategoryMo.Relations
        '    Dim moCategory = categoryRelation.Target
        '    Dim category = New Tuple(Of String, String)("RevitCategory", moCategory.Name)
        '    If Not dictionary.ContainsKey(New Tuple(Of String, String)("RevitCategory", moCategory.Name)) Then
        '        dictionary.Add(category, New Dictionary(Of Tuple(Of String, String), List(Of Tuple(Of String, String)))())
        '    End If
        'Next

        'Sort categories
        sortedDict = dictionary.OrderBy(Function(o) o.Key.Item2).ToDictionary(Function(pair) pair.Key, Function(pair) pair.Value)

#End Region


#Region "ANALYSE"

        'ANALYSE DES FAMILLES
        Dim famNameList, typeNameList, famList As New List(Of String)
        Dim nbForbidden As Integer = 0
        'Par categorie
        For Each categoryDictionaryElement In sortedDict
            'On verifie si il y a des familles
            If (categoryDictionaryElement.Key.Equals(noCategoryTuple)) Then
                Continue For
            Else
            End If

            'Par famille
            For Each familyDictionaryElement In categoryDictionaryElement.Value

                'On verifie si il y a une famille
                If (familyDictionaryElement.Key.Equals(noFamilyTuple)) Then
                    Continue For
                Else
                    'On compare avec la forbiddenList ET si famille Fenetre = fenetre
                    For Each forbiddenStr In forbiddenStrList
                        If familyDictionaryElement.Key.Item2.ToLower.Contains(forbiddenStr) Or familyDictionaryElement.Key.Item2.ToLower = categoryDictionaryElement.Key.Item2.ToLower Then
                            nbForbidden = nbForbidden + 1
                            famNameList.Add(familyDictionaryElement.Key.Item2.ToLower)
                            'Si oui, le travail sur la famille est fini
                            GoTo badFamilyName
                        End If
                    Next
                End If

                'Verification si = famillyName + " " + i.tostring (murs 1, murs 2... etc) = renommé automatiquement par RVT
                Dim myName = familyDictionaryElement.Key.Item2.ToLower.Split(" ")
                Dim m As Integer

                If myName.Count > 1 AndAlso (myName(0).ToLower = categoryDictionaryElement.Key.Item2.ToLower And Integer.TryParse(myName(1), m)) Then
                    nbForbidden = nbForbidden + 1
                    'Verif si mandatoryBool
                    'ElseIf mandatoryBool = True Then
                    '    For Each mandStr In mandatoryList
                    '        If Not (myName.Contains(mandStr)) Then
                    '            famNameList.Add(familyDictionaryElement.Key.Item2.ToLower)

                    '            nbForbidden = nbForbidden + 1
                    '            GoTo badFamilyName
                    '        End If
                    '    Next
                End If

badFamilyName:
                Dim typeList As New List(Of String)
                'On compare les noms des types de chaque famille
                For Each typeDictionaryElement In familyDictionaryElement.Value
                    For Each forbiddenStr In forbiddenStrList
                        If typeDictionaryElement.Item2.ToLower.Contains(forbiddenStr) Or typeDictionaryElement.Item2.ToLower = categoryDictionaryElement.Key.Item2.ToLower Or typeDictionaryElement.Item2.ToLower = familyDictionaryElement.Key.Item2.ToLower Then
                            typeNameList.Add(typeDictionaryElement.Item2.ToLower)
                            GoTo nextone2
                        End If
                    Next



                    Dim typeName = typeDictionaryElement.Item2.ToLower.Split(" ")
                    Dim n As Integer
                    If typeName.Count > 1 AndAlso (typeName(0).ToLower = categoryDictionaryElement.Key.Item2.ToLower And Integer.TryParse(typeName(1), n)) Then
                        nbForbidden = nbForbidden + 1
                        Exit For
                    ElseIf mandatoryBool = True Then
                        For Each mandStr In mandatoryList
                            If Not (typeName.Contains(mandStr)) Then
                                typeNameList.Add(typeDictionaryElement.Item2.ToLower)

                                GoTo nextone2
                            End If
                        Next
                    End If

                    If typeName.Count > 1 AndAlso (typeName(0).ToLower = familyDictionaryElement.Key.Item2.ToLower And Integer.TryParse(typeName(1), n)) Then
                        typeNameList.Add(typeDictionaryElement.Item2.ToLower)

                        Exit For
                    End If

                    'Lister les types pour les comparer entre eux

nextone2:
                    typeList.Add(typeDictionaryElement.Item2.ToLower)
                Next

                Dim u, v, wt As Integer
                If typeList.Count > 1 Then
                    For p = 0 To typeList.Count - 1
                        Dim fam = typeList.Item(p)
                        Dim typeSplit = fam.Split(" ")

                        If typeSplit.Count > 1 Then
                            Dim typeSuff As String = ""
                            For w = 0 To typeSplit.Count - 2
                                typeSuff = typeSuff + " " + typeSplit(w)
                            Next
                            For q = p + 1 To typeList.Count - 2
                                Dim qtypeSplit = typeList.Item(q).Split(" ")
                                Dim qtypesuff As String = ""
                                For w = 0 To qtypeSplit.Count - 2
                                    qtypesuff = qtypesuff + " " + qtypeSplit(w)
                                Next

                                If (typeSplit.Count > 1 And qtypeSplit.Count > 1) AndAlso (typeSuff = qtypesuff And (Integer.TryParse(qtypeSplit(1), v)) Or Integer.TryParse(typeSplit(1), u)) Then
                                    nbForbidden = nbForbidden + 1
                                    typeNameList.Add(fam)
                                    typeNameList.Add(typeList.Item(q))
                                    Exit For
                                End If
                            Next
                        End If
                    Next
                End If
                famList.Add(familyDictionaryElement.Key.Item2.ToLower)
            Next
            typeNameList = typeNameList.ToHashSet.ToList

            Dim r, s, t As Integer
            If famList.Count > 1 Then
                For p = 0 To famList.Count - 1
                    Dim fam = famList.Item(p)
                    Dim famSplit = fam.Split(" ")
                    Dim famSuff As String = ""

                    If famSplit.Count > 0 Then

                        For t = 0 To famSplit.Count - 2
                            famSuff = famSuff + " " + famSplit(t)
                        Next

                        For q = p + 1 To famList.Count - 2
                            Dim qfamSplit = famList.Item(q).Split(" ")

                            'Il faut bien vérifier que tous les noms soient égaux (on ne veut pas que "murs de base" et "murs soutainement 20" soient détectés par exemple)
                            If qfamSplit.Count = famSplit.Count And (famSplit.Count > 1 And qfamSplit.Count > 1) Then

                                Dim qfamsuff As String = ""
                                For t = 0 To qfamSplit.Count - 2
                                    qfamsuff = qfamsuff + " " + qfamSplit(t)
                                Next

                                If (famSuff.ToLower = qfamsuff.ToLower And (Integer.TryParse(qfamSplit(qfamSplit.Count - 1), s) Or Integer.TryParse(famSplit(famSplit.Count - 1), r))) Then
                                    nbForbidden = nbForbidden + 1
                                    famNameList.Add(fam)
                                    famNameList.Add(famList.Item(q))
                                    Exit For
                                End If

                            End If
                        Next
                    End If
                Next
            End If
            'famNameList.Clear()
        Next
        famNameList = famNameList.ToHashSet.ToList
#End Region


#Region "ANALYSE VUES"
        Dim forbiddenViewList As New List(Of String) From {"vue", "view", "dessin", "coupe", "plan", "Feuille", "Sheet", "niveau", "level", "elevation", "detail", "east", "west", "north", "south", "site", "est", "ouest", "nord", "sud"}
        Dim containsViewList As New List(Of String) From {"{", "sans nom"}
        Dim exactFOrbiddenList As New List(Of String) From {"east", "west", "north", "south", "site", "est", "ouest", "nord", "sud"}
        Dim levelList As New List(Of String)

        Dim myWs As Workspace = RevitFileTgmObj.Container
        Dim levelReq = From Mo In myWs.MetaObjects Where Mo.IsActive AndAlso Not IsNothing(Mo.GetAttribute("Object.type")) AndAlso Mo?.GetAttribute("Object.type").Value = "FileLevel"


        Dim viewWsheet As DevExpress.Spreadsheet.Worksheet = auditWb.Worksheets.Item("VUES")
        Dim vuesCol As Integer = 2
        Dim lastLine As Integer = 0

        Dim myColor As Drawing.Color = Drawing.Color.Red

        Dim j = 0
        While viewWsheet.Cells(3 + lastLine, vuesCol).Value.ToString <> ""
            lastLine = lastLine + 1
        End While

        Dim nameViewList As New List(Of String)
        Dim forbViewName As Integer = 0

        'Dim myColor As Drawing.Color = Drawing.Color.Red
        For i = 5 To lastLine + 3

            If viewWsheet.Cells(i, vuesCol).Value.ToString <> "" Then
                Dim mySplit = viewWsheet.Cells(i, vuesCol).Value.ToString.Split(" ")
                'Cas forbidden Coupe + " " + 1
                Dim w As Integer = 0
                For Each forbiddenStr In forbiddenViewList
                    If mySplit.Count > 1 AndAlso (mySplit(0).ToLower = forbiddenStr And Integer.TryParse(mySplit(1), w)) Then
                        nameViewList.Add(viewWsheet.Cells(i, vuesCol).Value.ToString)
                        forbViewName += 1
                        viewWsheet.Cells(i, vuesCol).Font.Color = myColor
                    End If
                Next

                'contiens sans nom ou une accolade
                For Each containsStr In containsViewList
                    If viewWsheet.Cells(i, vuesCol).Value.ToString.ToLower.Contains(containsStr) Then
                        nameViewList.Add(viewWsheet.Cells(i, vuesCol).Value.ToString)
                        forbViewName += 1
                        viewWsheet.Cells(i, vuesCol).Font.Color = myColor
                    End If
                Next

                'est exactement égal (vues de départ)
                For Each exactStr In exactFOrbiddenList
                    If viewWsheet.Cells(i, vuesCol).Value.ToString.ToLower = exactStr Then
                        nameViewList.Add(viewWsheet.Cells(i, vuesCol).Value.ToString)
                        forbViewName += 1
                        viewWsheet.Cells(i, vuesCol).Font.Color = myColor
                    End If
                Next

                'duplicata sauvage
                If viewWsheet.Cells(i, vuesCol).Value.ToString.ToLower.Contains("copie") Then
                    nameViewList.Add(viewWsheet.Cells(i, vuesCol).Value.ToString)
                    forbViewName += 1
                    viewWsheet.Cells(i, vuesCol).Font.Color = myColor
                End If

                'Nom = Nom etage
                For Each lvlMo In levelReq
                    If viewWsheet.Cells(i, vuesCol).Value.ToString.ToLower = lvlMo.Name Then
                        nameViewList.Add(viewWsheet.Cells(i, vuesCol).Value.ToString)
                        forbViewName += 1
                        viewWsheet.Cells(i, vuesCol).Font.Color = myColor
                    End If
                Next
            End If
            nameViewList = nameViewList.ToHashSet.ToList
        Next

#End Region

#Region "VERIF SHEETS"

        Dim sheetWsheet As DevExpress.Spreadsheet.Worksheet = auditWb.Worksheets.Item("FEUILLES")
        Dim sheetCOl As Integer = 4
        lastLine = 0

        j = 0
        While viewWsheet.Cells(3 + lastLine, sheetCOl).Value.ToString <> ""
            lastLine = lastLine + 1
        End While

        Dim nameSheetList As New List(Of String)
        Dim forbSheetName As Integer = 0
        For i = 5 To lastLine + 3

            If sheetWsheet.Cells(i, sheetCOl).Value.ToString <> "" Then
                'contiens sans nom ou une accolade
                For Each containsStr In containsViewList
                    If sheetWsheet.Cells(i, sheetCOl).Value.ToString.ToLower.Contains(containsStr) Then
                        nameSheetList.Add(sheetWsheet.Cells(i, sheetCOl).Value.ToString)
                        forbSheetName += 1
                        sheetWsheet.Cells(i, sheetCOl).Font.Color = myColor
                    End If
                Next
            End If
            nameSheetList = nameSheetList.ToHashSet.ToList
        Next

#End Region


#Region "RENSEIGNEMENT EXCEL"
        'Dim auditWsheet As DevExpress.Spreadsheet.Worksheet = auditWb.Worksheets.Item("STATISTIQUES D'ELEMENTS")
        'Dim famCol As Integer = 2
        'Dim typeCol As Integer = 3
        'lastLine = 0
        'Dim typeColor = auditWsheet.Cells(7, typeCol).FillColor
        'Dim famColor = auditWsheet.Cells(6, typeCol).FillColor
        ''Dim famColor = auditWsheet.Cells(3 + lastLine, typeCol).FillColor
        'Dim categColor = auditWsheet.Cells(5, typeCol).FillColor

        j = 0
        While auditWsheet.Cells(3 + lastLine, typeCol).Value.ToString <> ""
            lastLine = lastLine + 1
        End While
        Dim myRichtxt As New RichTextString()

        'myRichtxt.AddTextRun("", New RichTextRunFont("Calibri", 11, Drawing.Color.Red))
        For i = 7 To lastLine
            If auditWsheet.Cells(i, famCol).Value.ToString <> "" Then
                For Each fam In famNameList
                    If auditWsheet.Cells(i, famCol).Value.ToString.ToLower = fam Then
                        'aaaaaaaaaaaaaaaaaaaaa
                        auditWsheet.Cells(i, famCol).Font.Color = myColor
                    End If
                Next
            End If

            'Si nom type = nom famille
            If auditWsheet.Cells(i, famCol).Value.ToString <> "" Then
                For Each fam In typeNameList
                    If auditWsheet.Cells(i, famCol).Value.ToString.ToLower = fam Then
                        'aaaaaaaaaaaaaaaaaaaaa
                        auditWsheet.Cells(i, famCol).Font.Color = myColor
                    End If
                Next
            End If

            If auditWsheet.Cells(i, typeCol).Value.ToString <> "" Then
                For Each type In typeNameList
                    If auditWsheet.Cells(i, typeCol).Value.ToString.ToLower = type Then
                        'aaaaaaaaaaaaaaaaaaaaa

                        auditWsheet.Cells(i, typeCol).Font.Color = myColor
                    End If
                Next
            End If
        Next
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
                Continue For
            Else
                catNode = fileNode.AddNode(categoryDictionaryElement.Key.Item1, categoryDictionaryElement.Key.Item2)
                catNode.Description = "Category : " + categoryDictionaryElement.Key.Item2
            End If

            For Each familyDictionaryElement In categoryDictionaryElement.Value
                Dim famNode As Node
                If (familyDictionaryElement.Key.Equals(noFamilyTuple)) Then
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

#End Region

        'Familles
        If nbForbidden > 0 Then
            myAudit.ReportInfos.FamilyNamesCriteria_goodName = False
            myAudit.ReportInfos.FamilyNamesCriteria_numberOfBadNames = famNameList.Count + typeNameList.Count
        Else
            myAudit.ReportInfos.FamilyNamesCriteria_goodName = True
        End If

        'Vues
        If forbViewName > 0 Then
            myAudit.ReportInfos.ViewNamesCriteria_goodName = False
            myAudit.ReportInfos.ViewNamesCriteria_numberOfBadNames = nameViewList.Count
        Else
            myAudit.ReportInfos.FamilyNamesCriteria_goodName = True
        End If

        'feuilles
        If forbSheetName > 0 Then
            myAudit.ReportInfos.SheetNamesCriteria_goodName = False
            myAudit.ReportInfos.SheetNamesCriteria_numberOfBadNames = nameSheetList.Count
        Else
            myAudit.ReportInfos.SheetNamesCriteria_goodName = True
        End If

        myAudit.ProjWs.PushAllModifiedEntities()

        myAudit.ReportInfos.CompleteCriteria(29)
        myAudit.ReportInfos.CompleteCriteria(30)
        myAudit.ReportInfos.CompleteCriteria(31)

        myAudit.AuditWorkbook.Calculate()
        myAudit.AuditWorkbook.Worksheets.ActiveWorksheet = myAudit.ReportWorksheet
        myAudit.SaveAuditWorkbook()

        Return myTree
    End Function
End Class
