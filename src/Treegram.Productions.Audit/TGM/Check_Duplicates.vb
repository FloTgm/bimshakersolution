Imports DevExpress.Export.Xl
Imports DevExpress.Spreadsheet
Imports M4D.Production.Core
Imports M4D.Treegram.Core.Entities
Imports M4D.Treegram.Core.Kernel
Imports M4D.Treegram.Core.Extensions.Entities
Imports M4D.Treegram.Core.Extensions.Kernel
Imports Treegram.Bam.Libraries.AEC
Imports Treegram.Bam.Functions.AEC
Imports TGM.Deixi.Services
Imports M4D.Treegram.Core.Extensions.Models
Imports Treegram.Bam.Functions

Public Class Check_DuplicatesScript
    Inherits ProdScript
    Public Sub New()
        MyBase.New("TGM :: Check : Duplicates")
        AddAction(New Check_Duplicates())
    End Sub
End Class
Public Class Check_Duplicates
    Inherits ProdAction
    Public Sub New()
        Name = "TGM :: Check : Duplicates (ProdAction)"
        PartOfScript = True
    End Sub

    Public Overrides Sub CreateGuiInputTree()

        InputTree = TempWorkspace.AddTree("--- Temporary Input Tree ---")
        Dim objectNode = InputTree.SmartAddNode("object.containedIn", "3D",,, "Objets 3D")
        SelectYourSetsOfInput.Add("Objects3D", {objectNode}.ToList())

        ''---IFC objects
        'Dim ifcNode = objectNode.SmartAddNode("object.type", "ifc",,, "Objets IFC",, True)
        'Dim buildingNode = ifcNode.SmartAddNode("ProductType", "BUILDING")
        ''------Architecture
        'Dim archNode = ifcNode.SmartAddNode("object.type", "ifc",,, "Architecture",, True)
        'Dim arcTypesList As New List(Of String) From {"IfcFlowTerminal", "IfcBuildingElementPart", "IfcAnnotation", "IfcBuildingElementProxy", "IfcUnitaryEquipment", "IfcStair", "IfcWindow", "IfcRailing", "IfcLightFixtureType", "IfcCurtainWall", "IfcFurniture", "IfcWall", "IfcDistributionElement", "IfcSpace", "IfcCovering", "IfcDoor", "IfcColumn", "IfcRamp", "IfcAssembly", "IfcSite", "IfcSystemFurnitureElement", "IfcRoof"}
        'arcTypesList.Sort()
        'For Each typeSt In arcTypesList
        '    archNode.SmartAddNode("object.type", typeSt)
        'Next
        ''------Fluides
        'Dim fluidNode = ifcNode.SmartAddNode("object.type", "ifc",,, "Fluide",, True)
        'Dim fluidtypesList As New List(Of String) From {"IfcValveType", "IfcBuildingElementProxy", "IfcSwitchingDeviceType", "IfcFlowTerminal", "IfcElectricApplianceType", "IfcAirTerminal", "IfcPipeSegment", "IFCCableCarrierSegment", "IfcAlarmType", "IfcBuildingElementPart", "IfcAnnotation", "IfcUnitaryEquipment", "IfcSpace", "IfcDuctSegment", "Non exporté", "IfcCovering", "IfcLightFixtureType", "IfcPipeFitting", "IFCCableCarrierFittingType", "IFCCableCarrierFitting", "IfcDuctFitting", "IfcFireSuppressionTerminalType"}
        'fluidtypesList.Sort()
        'For Each typeSt In fluidtypesList
        '    fluidNode.SmartAddNode("object.type", typeSt)
        'Next
        ''------Structure
        'Dim strTypesList As New List(Of String) From {"IfcReinforcingMesh", "IfcReinforcingBar", "IfcReinforcementMesh", "IfcMechanicalFastener", "IfcBuildingElementPart", "IfcAnnotation", "IfcStair", "IfcSlab", "IfcBuildingElementProxy", "IfcWall", "IfcDistributionElement", "IfcSpace", "IfcColumn", "IfcAssembly", "IfcPlate", "IfcRamp", "IfcRoof"}


        ''---REVIT objects
        'Dim revNode = objectNode.SmartAddNode("RevitCategory",Nothing,, , "Objets REVIT")
        'Dim arcCategList As New List(Of String) From {"Appareils sanitaires", "Eléments", "Eléments de détail", "Environnement", "Equipement de génie climatique", "Equipement électrique", "Equipement spécialisé", "Escalier", "Fenêtres", "Garde-corps", "Installations électriques", "Lignes", "Luminaires", "Meneaux de murs-rideaux", "Meubles de rangement", "Mobilier", "Modèles génériques", "Murs", "Murs-rideaux", "Ossature", "Ouvertures de cage", "Panneaux de murs-rideaux", "Parking", "Pièces", "Plafonds", "Plantes", "Portes", "Poteaux", "Poteaux porteurs", "Rampes d'accès", "Réseaux de poutres", "Routes", "Site", "Surfaces", "Systèmes de mobilier", "Systèmes de murs-rideaux", "Toits", "Topographie", "Volume"}
        'Dim fluCategList As New List(Of String) From {"Accessoire de canalisation", "Accessoire de gaine", "Appareils d'appel malade", "Appareils de communication", "Appareils sanitaires", "Appareils téléphoniques", "Bouche d'aération", "Canalisation", "Canalisation souple", "Cheminement de fabrication MEP", "Chemins de câbles", "Conduits", "Dispositifs d'alarme incendie", "Dispositifs de données", "Dispositifs de sécurité", "Dispositifs d'éclairage", "Eléments", "Eléments de détail", "Equipement de génie climatique", "Equipement électrique", "Espaces", "Espaces réservés aux canalisations", "Espaces réservés aux gaines", "Fils", "Gaine", "Gaine flexible", "Installations électriques", "Isolations des canalisations", "Isolations des gaines", "Lignes", "Luminaires", "Modèles génériques", "Parking", "Raccords de canalisation", "Raccords de chemins de câbles", "Raccords de conduits", "Raccords de gaine", "Réseau de canalisations de fabrication MEP", "Réseau de gaines de fabrication MEP", "Revêtements des gaines", "Sprinklers", "Tirants de fabrication MEP", "Volume", "Zones HVAC"}
        'Dim strCategList As New List(Of String) From {"Armature à béton", "Armature surfacique", "Armature surfacique (treillis)", "Connexions structurelles", "Coupleur d'armature structurelle", "Dalles structurelles", "Direction principale du ferraillage", "Eléments", "Eléments de détail", "Escalier", "Fondations", "Lignes", "Modèles génériques", "Murs", "Ossature", "Ouvertures de cage", "Panneaux de murs-rideaux", "Pièces", "Poteaux", "Poteaux porteurs", "Poutres à treillis", "Raidisseurs", "Rampes d'accès", "Réseaux de poutres", "Sols", "Toits", "Volume"}
        'categList.Sort()
        'For Each categSt In categList
        '    revNode.SmartAddNode("RevitCategory", categSt)
        'Next
    End Sub

    Public Overrides Function GuiUpdateLauncherTree(inputs As Dictionary(Of String, IEnumerable(Of PersistEntity))) As Tree
        Dim launchTree As Tree = TempWorkspace.SmartAddTree("LauncherTree")
        launchTree.DuplicateObjectsWhileFiltering = False
        CommonAecLauncherTrees.FilesLaunchTree(launchTree, inputs)
        Return launchTree
    End Function

    <ActionMethod()>
    Public Function MyMethod(Objects3D As MultipleElements) As ActionResult

        If Objects3D.Source.Count = 0 Then
            Return New IgnoredActionResult("No 3D Objects Input")
        End If

        ' Récupération des inputs
        Dim tgmObjsList = GetInputAsMetaObjects(Objects3D.Source).ToHashSet 'To get rid of duplicates
        Dim outputList As New List(Of MetaObject)

        'ANALYSIS
        OutputTree = CompleteAuditDuplicates(tgmObjsList.ToList, outputList)

#Region "VIEW CREATION"
        Dim scanWs As Workspace = tgmObjsList(0).Container

        Dim chargedWs As New List(Of Workspace) From {scanWs}
        Dim checkedNodes As New List(Of Node)
        'Dim savedColors As New Dictionary(Of Node, Tuple(Of System.Windows.Media.Color, Single))
        Dim colorNodes As New List(Of NodeColorTransparency)

        'Colors used
        Dim grey10 As New Tuple(Of System.Windows.Media.Color, Single)(System.Windows.Media.Color.FromRgb(191, 191, 191), 0.1F)
        Dim red100 As New Tuple(Of System.Windows.Media.Color, Single)(System.Windows.Media.Color.FromRgb(255, 0, 0), 1.0F)

        'Récuperation du file
        Dim fileTgm As MetaObject = scanWs.GetMetaObjects(, "File").First

        Dim myTree = OutputTree
        Dim fileNode = myTree.SmartAddNode("File", fileTgm.Name)

        'savedColors.Add(fileNode, grey10)
        colorNodes.Add(New NodeColorTransparency(fileNode, grey10.Item1, grey10.Item2))
        checkedNodes.Add(fileNode)
        For Each catNode In fileNode.Nodes.ToList
            'savedColors.Add(catNode, grey10)
            colorNodes.Add(New NodeColorTransparency(catNode, grey10.Item1, grey10.Item2))
            checkedNodes.Add(catNode)
            For Each dupNode In catNode.Nodes.ToList
                'savedColors.Add(dupNode, red100)
                colorNodes.Add(New NodeColorTransparency(dupNode, red100.Item1, red100.Item2))
                checkedNodes.Add(dupNode)
                For Each eltNode In dupNode.Nodes.ToList
                    'savedColors.Add(eltNode, red100)
                    colorNodes.Add(New NodeColorTransparency(eltNode, red100.Item1, red100.Item2))
                    checkedNodes.Add(eltNode)
                Next
            Next
        Next

        TgmView.CreateView(fileTgm.GetProject, fileTgm.Name & " - DOUBLONS", chargedWs, myTree, checkedNodes, colorNodes)
#End Region

        Dim outputs = New List(Of Element) From {New MultipleElements(outputList, "DoublonsList", ElementState.[New])}
        Return New SucceededActionResult($"{outputList.Count} duplicate(s)", System.Drawing.Color.Green, outputs)
    End Function

    Private Function CompleteAuditDuplicates(tgmObjsList As List(Of MetaObject), ByRef outputList As List(Of MetaObject), Optional contactTolerance As Double = 0.01) As Tree

        Dim myAudit = CommonAecFunctions.GetAuditFileFromInputs(tgmObjsList) 'Initialize PsetTgm extensions
        Dim auditWb = myAudit.AuditWorkbook
        If auditWb Is Nothing And myAudit.IsIfc Then
            Create_IfcInitialAudit.GetAuditFromFile(myAudit)
        ElseIf auditWb Is Nothing Then
            Throw New Exception("Launch initial audit first")
        End If

#Region "GET INFOS"
        Treegram.GeomFunctions.Models.Reset3dDictionnary()
        Treegram.GeomFunctions.Models.LoadGeometryModels(tgmObjsList, True)
        Dim typesDico As Dictionary(Of String, List(Of MetaObject))
        If myAudit.ExtensionType = "ifc" Then
            typesDico = CommonFunctions.GetTgmTypes(tgmObjsList)
        Else
            typesDico = CommonFunctions.GetRevitCategories(tgmObjsList)
        End If
        Dim typesAndDuplicatesDico As New Dictionary(Of String, List(Of (clashInfo As Treegram.GeomFunctions.Clashs.ClashInfo, myList As List(Of MetaObject), Name As String)))
        Dim objsCheckedCount = 0
        Dim duplicatesCount = 0

        For Each typeSt In typesDico.Keys
            Dim typedObjs = typesDico(typeSt)
            Dim i As Integer
            For i = 0 To typedObjs.Count - 2
                Dim firstObj = typedObjs(i)
                Dim anaFirstObj As New AecObject(firstObj)
                If Not IsNothing(anaFirstObj.PsetTgmExtension) Then
                    firstObj.RemoveAttribute("Duplicates")
                End If
            Next
        Next

        For Each typeSt In typesDico.Keys
            Dim typedObjs = typesDico(typeSt)
            If typedObjs.Count < 2 Then Continue For
            objsCheckedCount += typedObjs.Count

            Dim duplicateList As New List(Of (clashInfo As Treegram.GeomFunctions.Clashs.ClashInfo, myList As List(Of MetaObject), Name As String))
            Dim doublonscheckList As New List(Of MetaObject)

            Dim i, j As Integer
            For i = 0 To typedObjs.Count - 2
                Dim firstObj = typedObjs(i)
                If Treegram.GeomFunctions.Models.GetGeometryModels(firstObj, True).Count = 0 Then
                    Continue For
                End If

                Dim doublonsList As New List(Of MetaObject)
                Dim M As Integer = 1
                For j = i + 1 To typedObjs.Count - 1
                    Dim secObj = typedObjs(j)

                    Dim identical = False
                    Try
                        identical = Treegram.GeomFunctions.Models.CompareModels(firstObj, secObj)
                    Catch ex As Exception 'objects without Models for example
                    End Try

                    If identical Then
                        If M = 1 Then
                            doublonsList.Add(firstObj)
                        End If
                        doublonsList.Add(secObj)
                        M = M + 1
                    End If

                Next

                Dim duplicateName As String = ""
                Dim alreadyInformed = False
                Dim n = 1
                For Each elt In doublonsList
                    'Verifier si les objets ne sont pas attribués

                    For Each myDbl In doublonscheckList
                        If elt.Id = myDbl.Id Then
                            alreadyInformed = True
                        End If
                    Next

                    If alreadyInformed = True Then
                        Exit For
                    Else
                        Dim attribute1 = AecObject.CompleteTgmPset(elt, "Duplicates", Nothing)
                        duplicateName = "Duplicate-" + (duplicatesCount + 1).ToString
                        Dim att1 As Attribute = attribute1.SmartAddAttribute(duplicateName, Nothing) 'ATTENTION : On écrit sur le scan !!!

                        Try
                            att1.SmartAddAttribute("Instance", n.ToString)
                            n += 1
                        Catch ex As Exception
                            Dim tested = 1
                        End Try

                    End If
                Next

                If alreadyInformed = False And doublonsList.Count > 1 Then
                    duplicatesCount += 1
                    duplicateList.Add((Nothing, doublonsList, duplicateName))
                    For Each dbl In doublonsList
                        doublonscheckList.Add(dbl)
                        outputList.Add(dbl)
                    Next
                End If
            Next
            If duplicateList.Count > 0 Then
                typesAndDuplicatesDico.Add(typeSt, duplicateList)
            End If
        Next
        myAudit.ReportInfos.DuplicateCriteria_duplicatesNumber = duplicatesCount

#End Region

#Region "FILL EXCEL"
        Dim doublonWsheet As Worksheet = myAudit.InsertProdActionTemplateInAuditWbook("ProdActions_Template", "DOUBLONS")
        Dim typeStartLine = 5
        Dim objIdAtt As String
        If myAudit.ExtensionType = "ifc" Then
            doublonWsheet.Cells(typeStartLine - 3, 1).Value = "TYPES IFC"
            objIdAtt = "GlobalId"
        Else
            doublonWsheet.Cells(typeStartLine - 3, 1).Value = "CATEGORIES RVT"
            objIdAtt = "RevitId"
        End If
        doublonWsheet.Cells(typeStartLine - 1, 1).Value = typesAndDuplicatesDico.Keys.Count
        doublonWsheet.Cells(typeStartLine - 1, 2).Value = duplicatesCount
        doublonWsheet.Cells(typeStartLine - 1, 3).Value = typesAndDuplicatesDico.SelectMany(Function(o) o.Value).Count
        doublonWsheet.Cells(typeStartLine - 1, 4).Value = "(" + objIdAtt + ")"
        doublonWsheet.Cells(typeStartLine - 1, 5).Value = "(" + objIdAtt + ")"

        'FILL ARRAY
        Dim indice_next_line = 0
        Dim indice_layer_start_line As Integer
        Dim totalRows = typesAndDuplicatesDico.SelectMany(Function(o) o.Value).Count + typesAndDuplicatesDico.Keys.Count
        If totalRows > 0 Then

            For Each typeSt In typesAndDuplicatesDico.Keys
                Dim duplicates = typesAndDuplicatesDico(typeSt)

                'Fill TYPES
                doublonWsheet.Range("B" + (typeStartLine + indice_next_line + 1).ToString).Value = typeSt
                doublonWsheet.Range("C" + (typeStartLine + indice_next_line + 1).ToString).Value = typesDico(typeSt).Count
                doublonWsheet.Range("D" + (typeStartLine + indice_next_line + 1).ToString).Value = duplicates.Count

                '---Mise en forme
                Dim typeRange = doublonWsheet.Range("B" + (typeStartLine + indice_next_line + 1).ToString + ":F" + (typeStartLine + indice_next_line + 1).ToString)
                typeRange.FillColor = XlColor.FromTheme(XlThemeColor.Dark1, 0.7)
                typeRange.Font.Bold = True
                typeRange.Borders.SetOutsideBorders(System.Drawing.Color.Black, BorderLineStyle.Thin)
                typeRange.Borders.InsideVerticalBorders.LineStyle = BorderLineStyle.Thin
                typeRange.Borders.InsideHorizontalBorders.LineStyle = BorderLineStyle.Thin

                indice_next_line += 1
                indice_layer_start_line = indice_next_line
                For Each duplicateInfo In duplicates
                    'Fill DUPLICATES
                    doublonWsheet.Range("E" + (typeStartLine + indice_next_line + 1).ToString).Value = duplicateInfo.myList.Item(0).GetAttribute(objIdAtt).Value.ToString

                    Dim concat As String = ""
                    For i = 1 To duplicateInfo.myList.Count - 1
                        Dim elt = duplicateInfo.myList.Item(i)
                        Dim idStr = elt.GetAttribute(objIdAtt).Value.ToString

                        If i = 1 Then
                            concat = idStr
                        Else
                            concat = concat + ";" + idStr
                        End If

                    Next
                    doublonWsheet.Range("F" + (typeStartLine + indice_next_line + 1).ToString).Value = concat

                    indice_next_line += 1
                Next

                '---Mise en forme
                If indice_next_line <> indice_layer_start_line Then
                    Dim namesRange = doublonWsheet.Range("B" + (typeStartLine + indice_layer_start_line + 1).ToString + ":F" + (typeStartLine + indice_next_line).ToString)
                    namesRange.Borders.SetOutsideBorders(System.Drawing.Color.Black, BorderLineStyle.Thin)
                    namesRange.Borders.InsideVerticalBorders.LineStyle = BorderLineStyle.Thin
                    namesRange.GroupRows(False)
                End If
            Next
        End If
#End Region

#Region "CREATE TREE"
        Dim myTree = myAudit.AuditTreesWs.SmartAddTree("Audit - DOUBLONS")
        Dim fileNode = myTree.SmartAddNode("File", myAudit.FileTgm.Name)
        For Each typeSt In typesAndDuplicatesDico.Keys
            Dim typeNode As Node
            Dim typeNode2 As Node
            If myAudit.ExtensionType = "ifc" Then
                typeNode = fileNode.SmartAddNode("object.type", typeSt)
            Else
                typeNode = fileNode.SmartAddNode("RevitCategory", typeSt)
            End If
            Dim typedDuplicates = typesAndDuplicatesDico(typeSt)


            For Each duplicate In typedDuplicates
                Dim dupliNode = typeNode.SmartAddNode(duplicate.Name, Nothing)
                For Each elt In duplicate.myList
                    dupliNode.SmartAddNode("object.id", elt.Id.ToString,,, elt.Name)
                Next
            Next
        Next
#End Region
        myAudit.ReportInfos.CompleteCriteria(55)

        'SAVE PROJECT
        myAudit.ProjWs.PushAllModifiedEntities()
        myAudit.AuditWorkbook.Calculate()
        myAudit.AuditWorkbook.Worksheets.ActiveWorksheet = myAudit.ReportWorksheet
        myAudit.SaveAuditWorkbook()
        Return myTree

    End Function
End Class

Public Class Clean_Duplicates
    Inherits ProdScript
    Public Sub New()
        MyBase.New("TGM :: Clean : Duplicates")
        AddAction(New Check_Duplicates())
        AddAction(New Deactivate_Duplicates())
    End Sub
End Class
Public Class Deactivate_Duplicates
    Inherits ProdAction
    Public Sub New()
        Name = "TGM :: Deactivate : Duplicates"
        Description = "Deactivate Duplicates"
        PartOfScript = True
    End Sub

    Public Overrides Function GuiUpdateLauncherTree(inputs As Dictionary(Of String, IEnumerable(Of PersistEntity))) As Tree

        Dim launchTree = TempWorkspace.SmartAddTree("LauncherTree")
        While (launchTree.Nodes.FirstOrDefault() IsNot Nothing)
            launchTree.RemoveNode(launchTree.Nodes.First)
        End While

        Dim tgmObjList As New List(Of MetaObject)
        If inputs.ContainsKey("DoublonsList") AndAlso inputs("DoublonsList").ToList.Count > 0 Then

            tgmObjList = GetInputAsMetaObjects(inputs("DoublonsList")).ToList
            Dim scanWs As Workspace = tgmObjList(0).Container
            Dim fileTgm As MetaObject = scanWs.GetMetaObjects(, "File").First
            Dim typesDico As Dictionary(Of String, List(Of MetaObject))

            If fileTgm.GetAttribute("Extension").Value = "ifc" Then
                typesDico = CommonFunctions.GetTgmTypes(tgmObjList)
            Else
                typesDico = CommonFunctions.GetRevitCategories(tgmObjList)
            End If


            Dim refInt As Integer = 0
            Dim minInt As Integer = 10000000
            For Each typeSt In typesDico.Keys
                Dim typedObjs = typesDico(typeSt)
                Dim typeNode As Node
                If fileTgm.GetAttribute("Extension").Value = "ifc" Then
                    typeNode = launchTree.SmartAddNode("object.type", typeSt)
                Else
                    typeNode = launchTree.SmartAddNode("RevitCategory", typeSt)
                End If
                For Each objTgm As MetaObject In typedObjs
                    Dim aecObj As New AecObject(objTgm)
                    Dim att = aecObj.PsetTgmExtension.GetAttribute("Duplicates")
                    For Each myAtt In att.Attributes
                        If myAtt.Name.Contains("Duplicate-") Then
                            Dim myNum = myAtt.Name.Split("-")(1)
                            Dim myInt = CInt(myNum)
                            If myInt > refInt Then
                                refInt = myInt
                            End If

                            'On cherche le plus petit duplicata
                            If myInt < minInt Then
                                minInt = myInt
                            End If
                        End If
                    Next
                Next
                For i = minInt To refInt
                    typeNode.SmartAddNode("Duplicate-" + i.ToString, Nothing)
                Next
            Next
        End If

        Return launchTree
    End Function


    <ActionMethod()>
    Public Function MyMethod(DoublonsList As MultipleElements) As ActionResult
        Dim dblList = GetInputAsMetaObjects(DoublonsList.Source).ToHashSet.ToList
        If dblList.Count = 0 Then Throw New Exception("Inputs missing")

        Dim refAttcount As Integer = 0
        Dim objToKeep As MetaObject
        For Each elt As MetaObject In dblList
            Dim attcount = elt.Attributes.Count

            If attcount > refAttcount Then
                refAttcount = attcount
                objToKeep = elt
            End If
        Next

        Dim n = 0
        For Each obj In dblList
            If obj.Id = objToKeep.Id Then
                Dim k = 0
            Else
                obj.IsActive = False
                n += 1
            End If
        Next

        Dim outputs = New List(Of Element) From {New MultipleElements(dblList, "DoublonsList", ElementState.[New])}
        Dim res = New SucceededActionResult(n.ToString + " Objets désactivés", System.Drawing.Color.Green, outputs)
        Return res
    End Function

End Class