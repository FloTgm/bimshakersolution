Imports Treegram.Bam.Libraries.AuditFile
Imports DevExpress.Export.Xl
Imports DevExpress.Spreadsheet
Imports M4D.Production.Core
Imports M4D.Treegram.Core.Entities
Imports M4D.Treegram.Core.Extensions.Models
Imports M4D.Treegram.Core.Kernel
Imports M4D.Treegram.Core.Extensions.Kernel
Imports M4D.Treegram.Core.Extensions.Entities
Imports Treegram.Bam.Libraries.AEC
Imports Treegram.Bam.Functions.AEC
Imports TGM.Deixi.Services
Imports Treegram.Bam.Functions

Public Class Analyse_ObjectStoreyRefScript
    Inherits ProdScript
    Public Sub New()
        MyBase.New("TGM :: Analyse : Objects Storey With References")
        AddAction(New Analyse_ObjectStoreyRef())
    End Sub
End Class
Public Class Analyse_ObjectStoreyRef
    Inherits Analyse_ObjectStorey
    Public Sub New()
        Name = "TGM :: Analyse : Objects Storey With References (ProdAction)"
        PartOfScript = True
    End Sub
    Public Overrides Sub CreateGuiInputTree()
        InputTree = MyBase.TempWorkspace.AddTree("--- Temporary Input Tree ---")
        Dim objectNode = InputTree.SmartAddNode("object.containedIn", "3D",,, "Objets 3D")
        SelectYourSetsOfInput.Add("Objects3D", {objectNode}.ToList())
        Dim storeyNode = InputTree.SmartAddNode("object.type", "TgmBuildingStorey",,, "Etages References")
        SelectYourSetsOfInput.Add("Storeys", {storeyNode}.ToList())
    End Sub
End Class

Public Class Analyse_ObjectStoreyScript
    Inherits ProdScript
    Public Sub New()
        MyBase.New("TGM :: Analyse : Objects Storey")
        AddAction(New Analyse_ObjectStorey())
    End Sub
End Class

Public Class Analyse_ObjectStorey
    Inherits ProdAction
    Public Sub New()
        Name = "TGM :: Analyse : Objects Storey (ProdAction)"
        PartOfScript = True
    End Sub

    Public Shared noStoreyMsg = "Objet sans étage"
    Public Shared noGeometryMsg = "Objet sans géométrie"
    Public Shared unexistingStoreyMsg = "Etage non existant"
    Public Shared badStoreyMsg = "Mauvais étage"
    Public Shared goodStoreyMsg = "Bon étage"

    Public Overrides Sub CreateGuiInputTree()
        InputTree = MyBase.TempWorkspace.AddTree("--- Temporary Input Tree ---")
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
        'Dim categList As New List(Of String) From {"Murs", "Sols", "Plafonds", "Portes", "Fenêtres", "Murs-rideaux", "Poteaux", "Poteaux Porteurs", "Poutres", "Escaliers", "Espaces", "Pièces", "Ossature", "Garde-corps", "Panneaux de murs-rideaux", "Modèles génériques", "Equipement spécialisé", "Canalisation", "Equipement de génie climatique", "Fondation", "Chemins de câbles", "Luminaires", "Gaines", "Bouche d'aération", "Installations électriques", "Equipements électriques", "Conduits", "Volées", "Site", "Dispositifs de sécurité", "Appareils téléphoniques", "Equipement spécialisé", "Mobilier", "Dispositifs d'alarme incendie", "Appareils de communication"}
        'categList.Sort()
        'For Each categSt In categList
        '    revNode.SmartAddNode("RevitCategory", categSt)
        'Next
        'Dim arcCategList As New List(Of String) From {"Appareils sanitaires", "Eléments", "Eléments de détail", "Environnement", "Equipement de génie climatique", "Equipement électrique", "Equipement spécialisé", "Escalier", "Fenêtres", "Garde-corps", "Installations électriques", "Lignes", "Luminaires", "Meneaux de murs-rideaux", "Meubles de rangement", "Mobilier", "Modèles génériques", "Murs", "Murs-rideaux", "Ossature", "Ouvertures de cage", "Panneaux de murs-rideaux", "Parking", "Pièces", "Plafonds", "Plantes", "Portes", "Poteaux", "Poteaux porteurs", "Rampes d'accès", "Réseaux de poutres", "Routes", "Site", "Sols", "Surfaces", "Systèmes de mobilier", "Systèmes de murs-rideaux", "Toits", "Topographie", "Volume"}
        'Dim fluCategList As New List(Of String) From {"Accessoire de canalisation", "Accessoire de gaine", "Appareils d'appel malade", "Appareils de communication", "Appareils sanitaires", "Appareils téléphoniques", "Bouche d'aération", "Canalisation", "Canalisation souple", "Cheminement de fabrication MEP", "Chemins de câbles", "Conduits", "Dispositifs d'alarme incendie", "Dispositifs de données", "Dispositifs de sécurité", "Dispositifs d'éclairage", "Eléments", "Eléments de détail", "Equipement de génie climatique", "Equipement électrique", "Espaces", "Espaces réservés aux canalisations", "Espaces réservés aux gaines", "Fils", "Gaine", "Gaine flexible", "Installations électriques", "Isolations des canalisations", "Isolations des gaines", "Lignes", "Luminaires", "Modèles génériques", "Parking", "Raccords de canalisation", "Raccords de chemins de câbles", "Raccords de conduits", "Raccords de gaine", "Réseau de canalisations de fabrication MEP", "Réseau de gaines de fabrication MEP", "Revêtements des gaines", "Sprinklers", "Tirants de fabrication MEP", "Volume", "Zones HVAC"}
        'Dim strCategList As New List(Of String) From {"Armature à béton", "Armature surfacique", "Armature surfacique (treillis)", "Connexions structurelles", "Coupleur d'armature structurelle", "Dalles structurelles", "Direction principale du ferraillage", "Eléments", "Eléments de détail", "Escalier", "Fondations", "Lignes", "Modèles génériques", "Murs", "Ossature", "Ouvertures de cage", "Panneaux de murs-rideaux", "Pièces", "Poteaux", "Poteaux porteurs", "Poutres à treillis", "Raidisseurs", "Rampes d'accès", "Réseaux de poutres", "Sols", "Toits", "Volume"}

        '---REF ETAGES
        Dim etageNode = InputTree.SmartAddNode("object.type", "FileLevel",,, "Etages")
        'Dim storeyNode = InputTree.SmartAddNode("object.type", "TgmBuildingStorey",,, "Etages References")
        SelectYourSetsOfInput.Add("Storeys", {etageNode}.ToList())

    End Sub

    Public Overrides Function GuiUpdateLauncherTree(inputs As Dictionary(Of String, IEnumerable(Of PersistEntity))) As Tree
        Dim launchTree As Tree = TempWorkspace.SmartAddTree("LauncherTree")
        launchTree.DuplicateObjectsWhileFiltering = True
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
            Dim fileNode = launchTree.SmartAddNode("File", fileTgm.Name,,,, True, True) '<--- Match before and match after to work with TgmBuildingStoreys
        Next
        Return launchTree
    End Function

    <ActionMethod()>
    Public Function MyMethod(Objects3D As MultipleElements, Storeys As MultipleElements) As ActionResult

        If Objects3D.Source.Count = 0 Then
            Return New IgnoredActionResult("No 3D Objects Input")
        ElseIf Storeys.Source.Count = 0 Then
            Return New IgnoredActionResult("No levels Input")
        End If
        'GET INPUTS
        Dim objs3dList = CType(Objects3D.Source, IEnumerable(Of PersistEntity)).ToHashSet 'To get rid of duplicates
        Dim storeysList = CType(Storeys.Source, IEnumerable(Of PersistEntity)).ToHashSet 'To get rid of duplicates
        Dim tgmObjsList As List(Of MetaObject) = objs3dList.Cast(Of MetaObject).ToList
        Dim tgmStoreysList As List(Of MetaObject) = storeysList.Cast(Of MetaObject).ToList

        If tgmObjsList.Count = 0 Then
            Return New IgnoredActionResult("No 3D Objects Input")
        ElseIf tgmStoreysList.Count = 0 Then
            Return New IgnoredActionResult("No levels Input")
        End If

        'AUDIT
        Dim outTree = GetStoreyAnalysis(tgmStoreysList, tgmObjsList)

#Region "VIEW CREATION"
        Dim scanWs As Workspace = tgmObjsList(0).Container

        'Récuperation du file
        Dim fileTgm As MetaObject = scanWs.GetMetaObjects(, "File").First
        Dim myAudit As New FileObject(fileTgm)

        Dim chargedWs As New List(Of Workspace) From {scanWs} 'Workspaces à charger

        'Recherche du WS de Reference s'il existe, et rajout à la liste des workspaces à charger
        Dim wsRef As Boolean = False
        If myAudit.ProjRefWs IsNot Nothing Then
            wsRef = True
            chargedWs.Add(myAudit.ProjRefWs)
        End If

        Dim checkedNodes As New List(Of Node) 'Noeuds à cocher
        Dim savedColors As New Dictionary(Of Node, Tuple(Of System.Windows.Media.Color, Single)) 'Dictionnaire couleurs et transparence des noeuds
        Dim colorNodes As New List(Of NodeColorTransparency)

        Dim fileNode = outTree.SmartAddNode("File", fileTgm.Name)
        Dim buildNodeList = fileNode.Nodes.ToList

        For Each buildNode In buildNodeList
            Dim allStoNode = buildNode.Nodes.Where(Function(o) o.Value = "").First

            Dim goodStoreyNode = allStoNode.Nodes.Where(Function(o) o.Value = goodStoreyMsg).First
            Dim badStoreyNode = allStoNode.Nodes.Where(Function(o) o.Value = badStoreyMsg).First
            Dim unexistingStoreyNode = allStoNode.Nodes.Where(Function(o) o.Value = unexistingStoreyMsg).First
            Dim noStoreyNode = allStoNode.Nodes.Where(Function(o) o.Value = noStoreyMsg).First
            Dim noGeometryNode = allStoNode.Nodes.Where(Function(o) o.Value = noGeometryMsg).First

            checkedNodes.AddRange({allStoNode, goodStoreyNode, badStoreyNode, unexistingStoreyNode, noStoreyNode, noGeometryNode})
            'savedColors.Add(allStoNode, New Tuple(Of System.Windows.Media.Color, Single)(System.Windows.Media.Color.FromRgb(191, 191, 191), 0.1F)) '10% grey
            'savedColors.Add(goodStoreyNode, New Tuple(Of System.Windows.Media.Color, Single)(System.Windows.Media.Color.FromRgb(144, 238, 144), 0.1F)) '10% light green
            'savedColors.Add(badStoreyNode, New Tuple(Of System.Windows.Media.Color, Single)(System.Windows.Media.Color.FromRgb(255, 0, 0), 1.0F)) '100% red
            'savedColors.Add(unexistingStoreyNode, New Tuple(Of System.Windows.Media.Color, Single)(System.Windows.Media.Color.FromRgb(75, 172, 198), 1.0F)) '100% celestial blue
            'savedColors.Add(noStoreyNode, New Tuple(Of System.Windows.Media.Color, Single)(System.Windows.Media.Color.FromRgb(255, 165, 0), 1.0F)) '100% orange
            'savedColors.Add(noGeometryNode, New Tuple(Of System.Windows.Media.Color, Single)(System.Windows.Media.Color.FromRgb(128, 100, 162), 1.0F)) '100% purple
            colorNodes.Add(New NodeColorTransparency(allStoNode, System.Windows.Media.Color.FromRgb(191, 191, 191), 0.1F))
            colorNodes.Add(New NodeColorTransparency(goodStoreyNode, System.Windows.Media.Color.FromRgb(144, 238, 144), 0.1F))
            colorNodes.Add(New NodeColorTransparency(badStoreyNode, System.Windows.Media.Color.FromRgb(255, 0, 0), 1F))
            colorNodes.Add(New NodeColorTransparency(unexistingStoreyNode, System.Windows.Media.Color.FromRgb(75, 172, 198), 1F))
            colorNodes.Add(New NodeColorTransparency(noStoreyNode, System.Windows.Media.Color.FromRgb(255, 165, 0), 1F))
            colorNodes.Add(New NodeColorTransparency(noGeometryNode, System.Windows.Media.Color.FromRgb(128, 100, 162), 1F))
        Next

        TgmView.CreateView(myAudit.ProjWs, fileTgm.Name & " - ANALYSE ETAGE", chargedWs, outTree, checkedNodes, colorNodes)
#End Region

        OutputTree = outTree

        Return New SucceededActionResult("Succeeded", System.Drawing.Color.Green, New List(Of Element) From {Objects3D})
    End Function

    Private Function GetStoreyAnalysis(tgmStoreysList As List(Of MetaObject), tgmObjsList As List(Of MetaObject)) As Tree

        Dim myAudit = CommonAecFunctions.GetAuditIfcFileFromInputs(tgmObjsList) 'Initialize PsetTgm extensions
        Dim auditWb = myAudit.AuditWorkbook
        If auditWb Is Nothing Then Throw New Exception("Launch initial audit first")
        Dim analysisWsheetName = "ANALYSE ETAGES"
        Dim stoAnaSettingsWsheet = myAudit.GetOrInsertSettingsWSheet(analysisWsheetName)
        If stoAnaSettingsWsheet Is Nothing Then Throw New Exception("Could not find """ + analysisWsheetName + """ worksheet in settings workbook")
        myAudit.ReportInfos.deltaInf = CDbl(stoAnaSettingsWsheet.Range("deltainf").Value.NumericValue)
        myAudit.ReportInfos.deltaSup = CDbl(stoAnaSettingsWsheet.Range("deltasup").Value.NumericValue)

        'FILTER AND LOAD
        TempWorkspace.SmartAddTree("ALL").Filter(False, New List(Of Workspace) From {myAudit.ScanWs}).RunSynchronously()
        Treegram.GeomFunctions.Models.Reset3dDictionnary()
        Treegram.GeomFunctions.Models.LoadGeometryModels(myAudit.ScanWs)
        'BetaTgmFcts.SmartJoin(myAudit.ScanWs,  M4D.Treegram.Core.Extensions.Enums.ExtensionType.AdditionalGeometries) 'Plus utile...

        Dim withRef As Boolean = False
        Dim tgmStoCount = tgmStoreysList.Where(Function(o) o.GetTgmType = "TgmBuildingStorey").Count
        If tgmStoCount = tgmStoreysList.Count Then
            If myAudit.FileTgm.GetAttribute("TgmBuilding") Is Nothing Then
                Throw New Exception("You must define TgmBuildings first !")
            End If
            withRef = True
        ElseIf tgmStoCount > 0 Then
            Throw New Exception("Les ""Storeys"" doivent être de même type !")
        End If

#Region "ANALYSE"
        '---sort tgm objects by TgmBuilding
        Dim elevAttName = "GlobalElevation"
        Dim buildDico As New Dictionary(Of String, List(Of MetaObject))
        Dim buildings As List(Of MetaObject)
        Dim buildingType, storeyType As String
        If withRef Then 'WITH REFERENCES
            buildingType = "TgmBuilding"
            storeyType = "TgmBuildingStorey"
            buildings = myAudit.ProjRefWs.GetMetaObjects(, buildingType).ToList
            Dim storeysDico
            If myAudit.IsIfc Then
                storeysDico = CommonFunctions.GetStoreysFromIfc(tgmObjsList)
            ElseIf myAudit.IsRevit Then
                storeysDico = CommonFunctions.GetStoreysFromRevit(tgmObjsList)
            Else
                Throw New Exception("Extension not implemented yet")
            End If
            For Each sto In storeysDico
                Dim storeyTgm = sto.Key
                Dim tgmBuildName = storeyTgm.GetAttribute(buildingType, True)?.Value
                If tgmBuildName Is Nothing Then
                    Throw New Exception("You must define TgmBuildings first !")
                End If
                If buildDico.ContainsKey(tgmBuildName) Then
                    buildDico(tgmBuildName).AddRange(sto.Value)
                Else
                    If buildings.Select(Function(o) o.Name.ToLower.Trim).Contains(tgmBuildName.tolower.trim) Then
                        buildDico.Add(tgmBuildName, sto.Value)
                    Else
                        Throw New Exception("This building """ + tgmBuildName + """ doesn't have storey references")
                    End If
                End If
            Next
        Else
            If myAudit.IsIfc Then 'IFC WITHOUT REFERENCES
                buildingType = "IfcBuilding"
                storeyType = "IfcBuildingStorey"
                buildings = myAudit.ScanWs.GetMetaObjects(, buildingType).ToList
                Dim buildTgmDico = CommonFunctions.GetBuildingsFromIfc(tgmObjsList)
                For Each build In buildTgmDico
                    buildDico.Add(build.Key.Name, build.Value)
                Next
            ElseIf myAudit.IsRevit Then 'REVIT WITHOUT REFERENCES
                buildingType = "None"
                storeyType = "FileLevel"
                buildings = Nothing
                Dim buildingName = myAudit.FileTgm.SmartGetAttribute("ProjectInformations/_\BuildingName")?.Value
                If buildingName Is Nothing Then buildingName = "DefaultTgmBuilding"
                Dim buildTgmDico = CommonFunctions.GetStoreysFromRevit(tgmObjsList)
                buildDico.Add(buildingName, tgmObjsList)
            Else
                Throw New Exception("Extension not implemented yet")
            End If
        End If


        '---fill building dictionary
        Dim storeyAnalysisDico As New Dictionary(Of String, (_3dObjs As List(Of MetaObject), storeys As List(Of String), elevations As List(Of Double)))
        If withRef Or myAudit.IsIfc Then
            For Each buildTgm As MetaObject In buildings
                If Not buildDico.ContainsKey(buildTgm.Name) Then
                    Continue For
                End If
                Dim storeys = buildTgm.GetChildren(storeyType).Where(Function(o) tgmStoreysList.Contains(o)).ToList
                CommonFunctions.SortLevelsByElevation(storeys, elevAttName)
                Dim stoNames = storeys.Select(Function(o) o.Name).ToList
                Dim stoElevations = storeys.Select(Function(o) CommonFunctions.ConvertGlobalElevationToScanWs(CDbl(o.GetAttribute(elevAttName, True).Value), myAudit.ScanWs)).ToList
                storeyAnalysisDico.Add(buildTgm.Name, (buildDico(buildTgm.Name), stoNames, stoElevations))
            Next
        ElseIf myAudit.IsRevit Then
            Dim storeys = myAudit.ScanWs.GetMetaObjects(, storeyType).Where(Function(o) tgmStoreysList.Contains(o)).ToList
            CommonFunctions.SortLevelsByElevation(storeys, elevAttName)
            Dim stoNames = storeys.Select(Function(o) o.Name).ToList
            Dim stoElevations = storeys.Select(Function(o) CommonFunctions.ConvertGlobalElevationToScanWs(CDbl(o.GetAttribute(elevAttName, True).Value), myAudit.ScanWs)).ToList
            storeyAnalysisDico.Add(buildDico.Keys(0), (buildDico.Values(0), stoNames, stoElevations))
        End If


        '---prepare storey analysis dictionnary (to get storey in the right order)
        Dim storeyAnalysisResultDico As New Dictionary(Of String, Dictionary(Of String, List(Of (aecObj As AecObject, comment As String))))
        For Each buildingName In storeyAnalysisDico.Keys
            Dim _3dObjs = storeyAnalysisDico(buildingName)._3dObjs
            Dim storeysDico As Dictionary(Of MetaObject, List(Of MetaObject))
            If myAudit.IsIfc Then
                storeysDico = CommonFunctions.GetStoreysFromIfc(buildDico(buildingName)) 'Ici on prend tous les étages présent dans le bâtiment
            ElseIf myAudit.IsRevit Then
                storeysDico = CommonFunctions.GetStoreysFromRevit(buildDico(buildingName))
            End If
            Dim storeys = storeysDico.Keys.ToList
            CommonFunctions.SortLevelsByElevation(storeys, elevAttName)
            Dim buildingResultDico As New Dictionary(Of String, List(Of (aecObj As AecObject, comment As String)))
            buildingResultDico.Add("<NONE>", New List(Of (aecObj As AecObject, comment As String)))
            For Each ifcSto In storeys
                buildingResultDico.Add(ifcSto.Name, New List(Of (aecObj As AecObject, comment As String)))
            Next
            storeyAnalysisResultDico.Add(buildingName, buildingResultDico)
        Next

        '---start analysis
        For Each buildingName In storeyAnalysisDico.Keys
            Dim _3dObjs = storeyAnalysisDico(buildingName)._3dObjs
            'Dim buildingResultDico As New Dictionary(Of String, List(Of (objTgm As MetaObject, comment As String)))
            For Each objTgm In _3dObjs
                AnalyseObjStorey(myAudit, storeyAnalysisDico, myAudit.ReportInfos.deltaInf, myAudit.ReportInfos.deltaSup, buildingName, storeyAnalysisResultDico, New AecObject(objTgm))
            Next
            'storeyAnalysisResultDico.Add(buildingName, buildingResultDico)
        Next
        '---rapport
        myAudit.ReportInfos.badStoreyNumber = storeyAnalysisResultDico.SelectMany(Function(o) o.Value).SelectMany(Function(o) o.Value).Where(Function(o) o.comment = badStoreyMsg).ToList.Count

#End Region

#Region "FILL EXCEL"
        Dim storeyAnaWsheet As Worksheet = myAudit.InsertProdActionTemplateInAuditWbook("ProdActions_Template", analysisWsheetName)
        Dim typeStartLine = 5
        'storeyAnaWsheet.Cells(typeStartLine - 1, 2).value = storeyAnalysisResultDico.Keys.Count
        Dim storeysNumber = storeyAnalysisResultDico.SelectMany(Function(o) o.Value).Count
        'storeyAnaWsheet.Cells(typeStartLine - 1, 3).value = storeysNumber
        Dim objsAnalysed = storeyAnalysisResultDico.SelectMany(Function(o) o.Value).SelectMany(Function(o) o.Value)
        Dim wrongObjs = objsAnalysed.Where(Function(o) o.comment <> goodStoreyMsg).ToList
        'storeyAnaWsheet.Cells(typeStartLine - 1, 4).value = "'" + wrongObjs.Count.ToString + "/" + objsAnalysed.Count.ToString
        Dim objIdAtt, categoryAtt As String
        If myAudit.ExtensionType = "ifc" Then
            storeyAnaWsheet.Cells(typeStartLine - 2, 6).Value = "Types IFC"
            objIdAtt = "GlobalId"
            categoryAtt = "object.type"
        Else
            storeyAnaWsheet.Cells(typeStartLine - 2, 6).Value = "Catégories RVT"
            objIdAtt = "RevitId"
            categoryAtt = "RevitCategory"
        End If

        'PREPARE ARRAY
        Dim indice_next_line = 0
        Dim totalRows = storeyAnalysisResultDico.Keys.Count + storeysNumber + objsAnalysed.Count
        Dim arrayColsNb = 7

        For Each buildingSt In storeyAnalysisResultDico.Keys
            Dim buildingAnalysis = storeyAnalysisResultDico(buildingSt)

            'Fill BUILDINGS
            If String.IsNullOrWhiteSpace(buildingSt) Then buildingSt = "Unknown"
            storeyAnaWsheet.Range("B" + (typeStartLine + indice_next_line + 1).ToString).Value = buildingSt
            Dim objsAnalysed2 = buildingAnalysis.SelectMany(Function(o) o.Value)
            Dim wrongObjs2 = objsAnalysed2.Where(Function(o) o.comment <> goodStoreyMsg).ToList
            storeyAnaWsheet.Range("D" + (typeStartLine + indice_next_line + 1).ToString).Value = "'" + wrongObjs2.Count.ToString + "/" + objsAnalysed2.Count.ToString

            '---Mise en forme
            Dim typeRange = storeyAnaWsheet.Range("B" + (typeStartLine + indice_next_line + 1).ToString + ":" + FileObject.GetSpreadsheetColumnName(arrayColsNb) + (typeStartLine + indice_next_line + 1).ToString)
            typeRange.FillColor = XlColor.FromTheme(XlThemeColor.Dark1, 0.7)
            typeRange.Font.Bold = True
            typeRange.Borders.SetOutsideBorders(System.Drawing.Color.Black, BorderLineStyle.Thin)
            typeRange.Borders.InsideVerticalBorders.LineStyle = BorderLineStyle.Thin
            typeRange.Borders.InsideHorizontalBorders.LineStyle = BorderLineStyle.Thin

            indice_next_line += 1
            Dim indice_layer_start_line As Integer = indice_next_line
            For Each storeySt In buildingAnalysis.Keys
                Dim storeyAnalysis = buildingAnalysis(storeySt)
                If storeySt = "<NONE>" AndAlso storeyAnalysis.Count = 0 Then
                    Continue For
                End If

                'Fill STOREYS
                storeyAnaWsheet.Range("C" + (typeStartLine + indice_next_line + 1).ToString).Value = storeySt
                Dim wrongObjs3 = storeyAnalysis.Where(Function(o) o.comment <> goodStoreyMsg).ToList
                storeyAnaWsheet.Range("D" + (typeStartLine + indice_next_line + 1).ToString).Value = "'" + wrongObjs3.Count.ToString + "/" + storeyAnalysis.Count.ToString


                '---Mise en forme
                Dim typeRange2 = storeyAnaWsheet.Range("B" + (typeStartLine + indice_next_line + 1).ToString + ":" + FileObject.GetSpreadsheetColumnName(arrayColsNb) + (typeStartLine + indice_next_line + 1).ToString)
                typeRange2.FillColor = XlColor.FromTheme(XlThemeColor.Dark1, 0.9)
                typeRange2.Font.Bold = True
                typeRange2.Borders.SetOutsideBorders(System.Drawing.Color.Black, BorderLineStyle.Thin)
                typeRange2.Borders.InsideVerticalBorders.LineStyle = BorderLineStyle.Thin
                typeRange2.Borders.InsideHorizontalBorders.LineStyle = BorderLineStyle.Thin


                indice_next_line += 1
                Dim indice_layer_start_line2 As Integer = indice_next_line
                For Each objAnalysis In storeyAnalysis
                    Dim aecObj = objAnalysis.aecObj
                    If objAnalysis.comment = goodStoreyMsg Then Continue For

                    'Fill OBJECTS
                    'If myAudit.IsRevit Then
                    '    dataArray(indice_next_line, 5) = objTgm.GetAttribute("RevitCategory")?.Value
                    'ElseIf myAudit.IsIfc Then
                    '    dataArray(indice_next_line, 5) = objTgm.GetTgmType
                    'End If

                    storeyAnaWsheet.Range("E" + (typeStartLine + indice_next_line + 1).ToString).Value = aecObj.Name
                    storeyAnaWsheet.Range("F" + (typeStartLine + indice_next_line + 1).ToString).Value = aecObj.Metaobject.GetAttribute(objIdAtt).Value.ToString
                    storeyAnaWsheet.Range("G" + (typeStartLine + indice_next_line + 1).ToString).Value = aecObj.Metaobject.GetAttribute(categoryAtt)?.Value.ToString
                    storeyAnaWsheet.Range("H" + (typeStartLine + indice_next_line + 1).ToString).Value = objAnalysis.comment
                    indice_next_line += 1
                Next

                '---Mise en forme
                If indice_next_line <> indice_layer_start_line2 Then
                    Dim namesRange = storeyAnaWsheet.Range("B" + (typeStartLine + indice_layer_start_line2 + 1).ToString + ":" + FileObject.GetSpreadsheetColumnName(arrayColsNb) + (typeStartLine + indice_next_line).ToString)
                    namesRange.Borders.SetOutsideBorders(System.Drawing.Color.Black, BorderLineStyle.Thin)
                    namesRange.Borders.InsideVerticalBorders.LineStyle = BorderLineStyle.Thin
                    namesRange.GroupRows(False)
                End If
            Next

            '---Mise en forme
            If indice_next_line <> indice_layer_start_line Then
                Dim namesRange = storeyAnaWsheet.Range("B" + (typeStartLine + indice_layer_start_line + 1).ToString + ":" + FileObject.GetSpreadsheetColumnName(arrayColsNb) + (typeStartLine + indice_next_line).ToString)
                namesRange.Borders.SetOutsideBorders(System.Drawing.Color.Black, BorderLineStyle.Thin)
                namesRange.Borders.InsideVerticalBorders.LineStyle = BorderLineStyle.Thin
                namesRange.GroupRows(False)
            End If
        Next

        '---Mise en forme
        storeyAnaWsheet.Columns("D").Font.Bold = False
        storeyAnaWsheet.Columns("E").AutoFit()
        storeyAnaWsheet.Columns("G").AutoFit()
        storeyAnaWsheet.Columns("H").AutoFit()
#End Region

#Region "CREATE TREE"
        Dim myTree = myAudit.AuditTreesWs.SmartAddTree("Audit - ANALYSE ETAGES")
        Dim fileNode = myTree.SmartAddNode("File", myAudit.FileTgm.Name)
        myTree.RemoveNode(fileNode)
        fileNode = myTree.SmartAddNode("File", myAudit.FileTgm.Name)

        For Each buildingName In storeyAnalysisResultDico.Keys
            Dim buildNode As Node
            If Not withRef And myAudit.IsRevit Then
                buildNode = fileNode.SmartAddNode("", Nothing,,, buildingName)
            Else
                buildNode = fileNode.SmartAddNode(buildingType, buildingName)
            End If

            '--All node
            Dim sourceStoreyType As String
            If myAudit.IsIfc Then
                sourceStoreyType = "IfcBuildingStorey"
            Else
                sourceStoreyType = "FileLevel"
            End If
            Dim allNode = buildNode.SmartAddNode(sourceStoreyType, Nothing,,, "Tous les étages")
            allNode.SmartAddNode("AuditStoreyAnalysisComment", goodStoreyMsg)
            allNode.SmartAddNode("AuditStoreyAnalysisComment", badStoreyMsg)
            allNode.SmartAddNode("AuditStoreyAnalysisComment", unexistingStoreyMsg)
            allNode.SmartAddNode("AuditStoreyAnalysisComment", noStoreyMsg)
            allNode.SmartAddNode("AuditStoreyAnalysisComment", noGeometryMsg)

            '--Storeys node
            Dim storeys = storeyAnalysisResultDico(buildingName).Keys.Where(Function(o) o <> "NONE").ToList
            For Each storeyName In storeys
                Dim storeyNode = buildNode.SmartAddNode(sourceStoreyType, storeyName)
                storeyNode.SmartAddNode("AuditStoreyAnalysisComment", goodStoreyMsg)
                storeyNode.SmartAddNode("AuditStoreyAnalysisComment", badStoreyMsg)
                storeyNode.SmartAddNode("AuditStoreyAnalysisComment", unexistingStoreyMsg)
                storeyNode.SmartAddNode("AuditStoreyAnalysisComment", noGeometryMsg)
            Next
        Next
        Dim stoNode = myTree.SmartAddNode("object.type", "TgmBuildingStorey")
        'stoNode.SmartAddNode("IsUsed", True)
#End Region

        'FILL REPORT
        myAudit.ReportInfos.CompleteCriteria(56)

        myAudit.ProjWs.PushAllModifiedEntities()
        myAudit.AuditWorkbook.Calculate()
        myAudit.AuditWorkbook.Worksheets.ActiveWorksheet = myAudit.ReportWorksheet
        myAudit.SaveAuditWorkbook()

        Return myTree

    End Function

    Private Sub AnalyseObjStorey(myAudit As IfcFileObject, storeyAnalysisDico As Dictionary(Of String, (_3dObjs As List(Of MetaObject), storeys As List(Of String), elevations As List(Of Double))), deltaInf As Double, deltaSup As Double, buildingName As String, buildingResultDico As Dictionary(Of String, Dictionary(Of String, List(Of (aecObj As AecObject, comment As String)))), aecObj As AecObject)

        '--Get audit parent storey
        Dim auditParentStorey As MetaObject
        If myAudit.IsRevit Then
            auditParentStorey = aecObj.Metaobject.GetParents("FileLevel",,, True).FirstOrDefault
        ElseIf myAudit.IsIfc Then
            auditParentStorey = aecObj.Metaobject.GetParents("IfcBuildingStorey",,, True).FirstOrDefault
        Else
            Throw New Exception("Extension not implemented yet")
        End If

        '--Get selected storeys
        Dim SelectedStoreyNames = storeyAnalysisDico(buildingName).storeys
        Dim storeyElevations = storeyAnalysisDico(buildingName).elevations
        If SelectedStoreyNames.Count = 0 Then GoTo Comparison

        '--Obj min max Z
        Dim minZ = Double.PositiveInfinity
        Dim maxZ = Double.NegativeInfinity
        If aecObj.Metaobject.GetAttribute("MaxZ") IsNot Nothing AndAlso aecObj.Metaobject.GetAttribute("MinZ") IsNot Nothing Then
            'If MaxZ and MinZ are already on scan
            minZ = CDbl(aecObj.Metaobject.GetAttribute("MinZ").Value)
            maxZ = CDbl(aecObj.Metaobject.GetAttribute("MaxZ").Value)
        Else
            'Else Analyse models
            For Each spaceModel In aecObj.Models(True)
                For Each vert In spaceModel.VerticesTesselation
                    If vert.Z < minZ Then minZ = vert.Z
                    If vert.Z > maxZ Then maxZ = vert.Z
                Next
            Next
            'Save in Tgm
            aecObj.CompleteTgmPset("MinZ", Math.Round(minZ, 5))
            aecObj.CompleteTgmPset("MaxZ", Math.Round(maxZ, 5))
        End If

        '--Obj storeys
        Dim ResultStoreyNames As New List(Of String) 'One wall can belong to several storeys, in case of full height wall
        If aecObj.Models(True).Count > 0 Then
            'First level case - everything underneath belong to this level !
            If minZ < storeyElevations(1) + deltaSup Then
                ResultStoreyNames.Add(SelectedStoreyNames.First)
            End If
            'Mid level cases
            Dim i As Integer
            For i = 1 To SelectedStoreyNames.Count - 2
                Dim elevInf = storeyElevations(i)
                Dim elevSup = storeyElevations(i + 1)
                If (elevInf + deltaInf <= Math.Round(minZ, 4) And Math.Round(minZ, 4) <= elevSup + deltaSup) Or
                          (elevInf + deltaInf <= Math.Round(maxZ, 4) And Math.Round(maxZ, 4) <= elevSup + deltaSup) Or
                              (Math.Round(minZ, 4) < elevInf + deltaInf And Math.Round(maxZ, 4) > elevSup + deltaSup) Then
                    ResultStoreyNames.Add(SelectedStoreyNames(i))
                End If
            Next
            'Last level case - everything above belong to this level !
            If maxZ > storeyElevations.Last + deltaInf Then
                ResultStoreyNames.Add(SelectedStoreyNames.Last)
            End If
        End If

Comparison:
        '--Comparison
        Dim myObjTuple As (aecObj As AecObject, comment As String)
        Dim auditStoreyName As String
        If auditParentStorey Is Nothing Then
            auditStoreyName = "<NONE>"
            myObjTuple = (aecObj, Analyse_ObjectStorey.noStoreyMsg)
        ElseIf aecObj.Models(True).Count = 0 OrElse aecObj.Models(True).Select(Function(o) o.VerticesTesselation).Count = 0 Then
            auditStoreyName = auditParentStorey.Name
            myObjTuple = (aecObj, noGeometryMsg)
        ElseIf Not SelectedStoreyNames.Select(Function(o) o.ToLower.Trim).Contains(auditParentStorey.Name.ToLower.Trim) Then
            auditStoreyName = auditParentStorey.Name
            myObjTuple = (aecObj, unexistingStoreyMsg)
        ElseIf ResultStoreyNames.Count = 0 Then 'ça devrait pas se produire... à sup à terme !
            auditStoreyName = "<IMPOSSIBLE>"
            myObjTuple = (aecObj, Analyse_ObjectStorey.noStoreyMsg)
        ElseIf Not ResultStoreyNames.Select(Function(o) o.ToLower.Trim).Contains(auditParentStorey.Name.ToLower.Trim) Then
            auditStoreyName = ResultStoreyNames.First
            For j As Integer = 1 To ResultStoreyNames.Count - 1
                auditStoreyName += "/_\" + ResultStoreyNames(j)
            Next
            myObjTuple = (aecObj, badStoreyMsg)
        Else
            auditStoreyName = auditParentStorey.Name
            myObjTuple = (aecObj, goodStoreyMsg)
        End If

        '---Fill dico for excel purpose
        If auditParentStorey Is Nothing Then
            buildingResultDico(buildingName)("<NONE>").Add(myObjTuple)
        Else
            buildingResultDico(buildingName)(auditParentStorey.Name).Add(myObjTuple)
        End If

        '--Write result in tgm
        aecObj.CompleteTgmPset("AuditStoreyAnalysisResult", auditStoreyName)
        aecObj.CompleteTgmPset("AuditStoreyAnalysisComment", myObjTuple.comment)

    End Sub

End Class
