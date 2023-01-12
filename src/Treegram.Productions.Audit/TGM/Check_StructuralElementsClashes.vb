Imports M4D.Production.Core
Imports M4D.Treegram.Core.Entities
Imports M4D.Treegram.Core.Kernel
Imports M4D.Treegram.Core.Extensions.Kernel
Imports M4D.Treegram.Core.Extensions.Entities
Imports Treegram.Bam.Libraries.AuditFile
Imports DevExpress.Spreadsheet
Imports DevExpress.Export.Xl
Imports Treegram.Bam.Libraries.AEC
Imports Treegram.Bam.Functions.AEC
Imports Treegram.Bam.Functions

Public Class Check_StructuralElementsClashesIfcScript
    Inherits ProdScript
    Public Sub New()
        MyBase.New("IFC :: Clash : Structural Elements")
        AddAction(New Check_StructuralElementsClashesIfc())
    End Sub
End Class
Public Class Check_StructuralElementsClashesRvtScript
    Inherits ProdScript
    Public Sub New()
        MyBase.New("RVT :: Clash : Structural Elements")
        AddAction(New Check_StructuralElementsClashesRvt())
    End Sub
End Class


Public Class Check_StructuralElementsClashesIfc
    Inherits Check_StructuralElementsClashes
    Public Sub New()
        Name = "IFC :: Clash : Structural Elements (ProdAction)"
        PartOfScript = True
    End Sub

    Public Overrides Sub CreateGuiInputTree()
        InputTree = TempWorkspace.AddTree("--- Temporary Input Tree ---")
        Dim objectNode = InputTree.SmartAddNode("object.containedIn", "3D",,, "Objets 3D")

        '---IFC objects
        Dim ifcNode = objectNode.SmartAddNode("object.type", "ifc",,, "Objets IFC",, True)
        Dim strTypesList As New List(Of String) From {"IfcWall", "IfcColumn", "IfcBeam", "IfcSlab"}
        strTypesList.Sort()

        Dim myNodeList As New List(Of Node)

        For Each typeSt In strTypesList
            myNodeList.Add(ifcNode.SmartAddNode("object.type", typeSt))
        Next

        SelectYourSetsOfInput.Add("Objects3D", myNodeList)
    End Sub
End Class
Public Class Check_StructuralElementsClashesRvt
    Inherits Check_StructuralElementsClashes
    Public Sub New()
        Name = "RVT :: Clash : Structural Elements (ProdAction)"
        PartOfScript = True
    End Sub

    Public Overrides Sub CreateGuiInputTree()

        InputTree = TempWorkspace.AddTree("--- Temporary Input Tree ---")
        Dim objectNode = InputTree.SmartAddNode("object.containedIn", "3D",,, "Objets 3D")

        '---REVIT objects
        Dim revNode = objectNode.SmartAddNode("RevitCategory", Nothing,, , "Objets REVIT")
        Dim categList As New List(Of String) From {"Armature à béton", "Armature surfacique", "Armature surfacique (treillis)", "Connexions structurelles", "Coupleur d'armature structurelle", "Dalles structurelles", "Direction principale du ferraillage", "Eléments", "Eléments de détail", "Escalier", "Fondations", "Lignes", "Modèles génériques", "Murs", "Ossature", "Ouvertures de cage", "Panneaux de murs-rideaux", "Pièces", "Poteaux", "Poteaux porteurs", "Poutres à treillis", "Raidisseurs", "Rampes d'accès", "Réseaux de poutres", "Sols", "Toits", "Volume"}

        categList.Sort()
        Dim myNodeList As New List(Of Node)
        For Each categSt In categList
            myNodeList.Add(revNode.SmartAddNode("RevitCategory", categSt))
        Next
        SelectYourSetsOfInput.Add("Objects3D", myNodeList)
    End Sub
End Class

Public Class Check_StructuralElementsClashes
    Inherits ProdAction
    Public Sub New()
        Name = "TGM :: Clash : Structural Elements"
        PartOfScript = True
    End Sub

    Public Overrides Sub CreateGuiInputTree()

        InputTree = TempWorkspace.AddTree("--- Temporary Input Tree ---")
        Dim objectNode = InputTree.SmartAddNode("object.containedIn", "3D",,, "Objets 3D")

        '---IFC objects
        Dim ifcNode = objectNode.SmartAddNode("object.type", "ifc",,, "Objets IFC",, True)
        Dim strTypesList As New List(Of String) From {"IfcWall", "IfcColumn", "IfcBeam", "IfcSlab"}
        strTypesList.Sort()
        For Each typeSt In strTypesList
            ifcNode.SmartAddNode("object.type", typeSt)
        Next

        '---REVIT objects
        Dim revNode = objectNode.SmartAddNode("RevitCategory", Nothing,, , "Objets REVIT")
        Dim categList As New List(Of String) From {"Murs", "Sols", "Plafonds", "Portes", "Fenêtres", "Murs-rideaux", "Poteaux", "Poteaux Porteurs", "Poutres", "Escaliers", "Espaces", "Pièces", "Ossature", "Garde-corps", "Panneaux de murs-rideaux", "Modèles génériques", "Equipement spécialisé", "Canalisation", "Equipement de génie climatique", "Fondation", "Chemins de câbles", "Luminaires", "Gaines", "Bouche d'aération", "Installations électriques", "Equipements électriques", "Conduits", "Volées"}
        categList.Sort()
        For Each categSt In categList
            revNode.SmartAddNode("RevitCategory", categSt)
        Next
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

    <ActionMethod()>
    Public Function MyMethod(Objects3D As MultipleElements) As ActionResult
        If Objects3D.Source.Count = 0 Then
            Return New IgnoredActionResult("No 3D Objects")
        End If

        ' Récupération des inputs
        Dim tgmObjsList = GetInputAsMetaObjects(Objects3D.Source).ToHashSet 'To get rid of duplicates

        'ANALYSIS
        OutputTree = CompleteAuditStructuralClashes(tgmObjsList.ToList)

        Return New SucceededActionResult("Succeeded", System.Drawing.Color.Green, New List(Of Element) From {Objects3D})
    End Function

    Private Function CompleteAuditStructuralClashes(tgmObjsList As List(Of MetaObject)) As Tree

        Dim myAudit = CommonAecFunctions.GetAuditFileFromInputs(tgmObjsList)
        Dim auditWb = myAudit.AuditWorkbook
        If auditWb Is Nothing Then Throw New Exception("Launch initial audit first")
        Dim analysisWsheetName = "CLASHS STRUCTURELS"
        Dim structClashSettingsWsheet = myAudit.GetOrInsertSettingsWSheet(analysisWsheetName)
        If structClashSettingsWsheet Is Nothing Then Throw New Exception("Could not find """ + analysisWsheetName + """ worksheet in settings workbook")
        myAudit.ReportInfos.contactTolerance = CDbl(structClashSettingsWsheet.Range("contactTolerance").Value.NumericValue)

#Region "GET INFOS"
        Treegram.GeomFunctions.Models.Reset3dDictionnary()
        Treegram.GeomFunctions.Models.LoadGeometryModels(tgmObjsList, True)
        Dim typesDico As Dictionary(Of String, List(Of MetaObject))
        If myAudit.IsIfc Then
            typesDico = CommonFunctions.GetTgmTypes(tgmObjsList)
        Else
            typesDico = CommonFunctions.GetRevitCategories(tgmObjsList)
        End If
        'Dim typesAndDuplicatesDico As New Dictionary(Of String, List(Of (clashInfo As Treegram.GeomFunctions.Clashs.ClashInfo, obj1 As MetaObject, obj2 As MetaObject, Name As String)))
        Dim ClashesDico As New Dictionary(Of String, List(Of (clashInfo As Treegram.GeomFunctions.Clashs.ClashInfo, obj1 As MetaObject, obj2 As MetaObject, Name As String)))
        Dim objsCheckedCount = 0
        Dim clashesCount = 0
        Dim i, j As Integer
        If typesDico.Keys.Count < 2 Then
            Throw New Exception("Less than 2 ifc types")
        End If
        For i = 0 To typesDico.Keys.Count - 1
            Dim typeSt = typesDico.Keys(i)
            Dim typedObjs = typesDico(typeSt)
            objsCheckedCount += typedObjs.Count

            'Dim duplicateList As New List(Of (clashInfo As Treegram.GeomFunctions.Clashs.ClashInfo, obj1 As MetaObject, obj2 As MetaObject, Name As String))
            For Each firstObj In typedObjs
                If Treegram.GeomFunctions.Models.GetGeometryModels(firstObj, True).Count = 0 Then
                    Continue For
                End If

                For j = i + 1 To typesDico.Keys.Count - 1
                    Dim otherTypeSt = typesDico.Keys(j)
                    Dim otherTypedObjs = typesDico(otherTypeSt)
                    For Each secObj In otherTypedObjs
                        Dim myClash As Treegram.GeomFunctions.Clashs.ClashInfo
                        Try
                            myClash = Treegram.GeomFunctions.Clashs.AdvanceClash(firstObj, secObj, myAudit.ReportInfos.contactTolerance, False, True)
                        Catch ex As Exception
                            Continue For
                        End Try
                        If myClash.ClashState <> M4D.Deixi.Core.Enum.ClashState.None And myClash.ClashState <> M4D.Deixi.Core.Enum.ClashState.Contact Then
                            Dim clashName = "Clash-" + (clashesCount + 1).ToString
                            'duplicateList.Add((myClash, firstObj, secObj, clashName))
                            Dim attribute1 = AecObject.CompleteTgmPset(firstObj, "Clashes", Nothing)
                            Dim attribute2 = AecObject.CompleteTgmPset(secObj, "Clashes", Nothing)
                            Dim att1 As Attribute = attribute1.SmartAddAttribute(clashName, Nothing)
                            Dim att2 As Attribute = attribute2.SmartAddAttribute(clashName, Nothing)
                            att1.SmartAddAttribute("ClashType", myClash.ClashState.ToString)
                            att2.SmartAddAttribute("ClashType", myClash.ClashState.ToString)

                            If ClashesDico.ContainsKey(myClash.ClashState.ToString) Then
                                ClashesDico(myClash.ClashState.ToString).Add((myClash, firstObj, secObj, clashName))
                            Else
                                ClashesDico.Add(myClash.ClashState.ToString, New List(Of (clashInfo As Treegram.GeomFunctions.Clashs.ClashInfo, obj1 As MetaObject, obj2 As MetaObject, Name As String)) From {(myClash, firstObj, secObj, clashName)})
                            End If
                            clashesCount += 1
                        End If
                    Next
                Next
            Next
            'typesAndDuplicatesDico.Add(typeSt, duplicateList)
        Next
        myAudit.ReportInfos.StructuralClashCriteria_clashesNumber = clashesCount
#End Region

#Region "FILL EXCEL"
        Dim doublonWsheet As Worksheet = myAudit.InsertProdActionTemplateInAuditWbook("ProdActions_Template", "CLASH STRUCTURELS")
        Dim wsStartLine = 5
        Dim objIdAtt, objType As String
        If myAudit.ExtensionType = "ifc" Then
            doublonWsheet.Cells(wsStartLine - 3, 1).Value = "TYPES IFC"
            objIdAtt = "GlobalId"
            objType = "object.type"
        Else
            doublonWsheet.Cells(wsStartLine - 3, 1).Value = "CATEGORIES RVT"
            objIdAtt = "RevitId"
            objType = "RevitCategory"
        End If
        doublonWsheet.Cells(wsStartLine - 1, 1).Value = ClashesDico.Keys.Count
        doublonWsheet.Cells(wsStartLine - 1, 2).Value = ClashesDico.SelectMany(Function(o) o.Value).Count
        doublonWsheet.Cells(wsStartLine - 1, 3).Value = "(" + objIdAtt + ")"
        doublonWsheet.Cells(wsStartLine - 1, 5).Value = "(" + objIdAtt + ")"

        'FILL ARRAY
        Dim indice_next_line = 0
        Dim indice_layer_start_line As Integer
        Dim totalRows = ClashesDico.SelectMany(Function(o) o.Value).Count + ClashesDico.Keys.Count
        If totalRows > 0 Then
            '---Prepare array
            Dim arrayLinesNb = totalRows
            Dim arrayColsNb = 6

            For Each clashTypeSt In ClashesDico.Keys
                Dim typedClashes = ClashesDico(clashTypeSt)

                'Fill TYPES
                doublonWsheet.Range("B" + (wsStartLine + indice_next_line + 1).ToString).Value = clashTypeSt
                doublonWsheet.Range("C" + (wsStartLine + indice_next_line + 1).ToString).Value = typedClashes.Count

                '---Mise en forme
                Dim typeRange = doublonWsheet.Range("B" + (wsStartLine + indice_next_line + 1).ToString + ":" + FileObject.GetSpreadsheetColumnName(arrayColsNb) + (wsStartLine + indice_next_line + 1).ToString)
                typeRange.FillColor = XlColor.FromTheme(XlThemeColor.Dark1, 0.7)
                typeRange.Font.Bold = True
                typeRange.Borders.SetOutsideBorders(System.Drawing.Color.Black, BorderLineStyle.Thin)
                typeRange.Borders.InsideVerticalBorders.LineStyle = BorderLineStyle.Thin
                typeRange.Borders.InsideHorizontalBorders.LineStyle = BorderLineStyle.Thin

                indice_next_line += 1
                indice_layer_start_line = indice_next_line
                For Each clashInfo In typedClashes
                    'Fill DUPLICATES
                    doublonWsheet.Range("C" + (wsStartLine + indice_next_line + 1).ToString).Value = clashInfo.Name
                    doublonWsheet.Range("D" + (wsStartLine + indice_next_line + 1).ToString).Value = clashInfo.obj1.GetAttribute(objIdAtt).Value.ToString
                    doublonWsheet.Range("E" + (wsStartLine + indice_next_line + 1).ToString).Value = clashInfo.obj1.GetAttribute(objType).Value.ToString
                    doublonWsheet.Range("F" + (wsStartLine + indice_next_line + 1).ToString).Value = clashInfo.obj2.GetAttribute(objIdAtt).Value.ToString
                    doublonWsheet.Range("G" + (wsStartLine + indice_next_line + 1).ToString).Value = clashInfo.obj2.GetAttribute(objType).Value.ToString
                    indice_next_line += 1
                Next

                '---Mise en forme
                If indice_next_line <> indice_layer_start_line Then
                    Dim namesRange = doublonWsheet.Range("B" + (wsStartLine + indice_layer_start_line + 1).ToString + ":" + FileObject.GetSpreadsheetColumnName(arrayColsNb) + (wsStartLine + indice_next_line).ToString)
                    namesRange.Borders.SetOutsideBorders(System.Drawing.Color.Black, BorderLineStyle.Thin)
                    namesRange.Borders.InsideVerticalBorders.LineStyle = BorderLineStyle.Thin
                    namesRange.GroupRows(False)
                End If
            Next
        End If
#End Region

#Region "CREATE TREE"
        Dim myTree = myAudit.AuditTreesWs.SmartAddTree("Audit - CLASH STRUCTURELS")
        Dim fileNode = myTree.SmartAddNode("File", myAudit.FileTgm.Name)
        myTree.RemoveNode(fileNode)
        fileNode = myTree.SmartAddNode("File", myAudit.FileTgm.Name)

        For Each clashTypeSt In ClashesDico.Keys
            Dim typeNode = fileNode.SmartAddNode("ClashType", clashTypeSt)
            For Each clashInfo In ClashesDico(clashTypeSt).ToList
                Dim clashNode = typeNode.SmartAddNode(clashInfo.Name, Nothing)
                clashNode.SmartAddNode("object.id", clashInfo.obj1.Id.ToString,,, clashInfo.obj1.Name)
                clashNode.SmartAddNode("object.id", clashInfo.obj2.Id.ToString,,, clashInfo.obj2.Name)
            Next
        Next
#End Region

        'FILL REPORT
        myAudit.ReportInfos.CompleteCriteria(65)

        myAudit.ProjWs.PushAllModifiedEntities()
        myAudit.AuditWorkbook.Calculate()
        myAudit.AuditWorkbook.Worksheets.ActiveWorksheet = myAudit.ReportWorksheet
        myAudit.SaveAuditWorkbook()

        Return myTree

    End Function
End Class
