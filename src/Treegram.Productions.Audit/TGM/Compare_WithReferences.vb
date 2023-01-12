Imports System.Windows.Forms
Imports Treegram.Bam.Libraries.AuditFile
Imports M4D.Production.Core
Imports M4D.Treegram.Core.Entities
Imports M4D.Treegram.Core.Kernel
Imports M4D.Treegram.Core.Extensions.Kernel
Imports M4D.Treegram.Core.Extensions.Entities
Imports DevExpress.Spreadsheet
Imports Treegram.Bam.Libraries.AEC
Imports Treegram.Bam.Functions.AEC
Imports Treegram.ConstructionManagement
Imports Treegram.Bam.Functions

Public Class Compare_WithReferencesScript
    Inherits ProdScript
    Public Sub New()
        MyBase.New("TGM :: Compare : With References")
        AddAction(New Compare_WithReferences)
    End Sub
End Class
Public Class Compare_WithReferences
    Inherits ProdAction
    Public Sub New()
        Name = "TGM :: Compare : With References (ProdAction)"
        PartOfScript = True
    End Sub

    Public Overrides Sub CreateGuiInputTree()
        InputTree = TempWorkspace.AddTree("--- Temporary Input Tree ---")
        Dim myFileNode = InputTree.AddNode("object.type", "File")
        myFileNode.Description = "Files"
        SelectYourSetsOfInput.Add("Files", New List(Of Node) From {myFileNode})
    End Sub

    Public Overrides Function GuiUpdateLauncherTree(inputs As Dictionary(Of String, IEnumerable(Of PersistEntity))) As Tree

        Dim launchTree = TempWorkspace.SmartAddTree("LauncherTree")

        ' Suppression des anciens noeuds
        While launchTree.Nodes.Count > 0
            launchTree.RemoveNode(launchTree.Nodes.First)
        End While

        If inputs.Values.SelectMany(Function(i) i).ToList.Count > 0 Then
            Dim count = 0
            For Each inputObj In inputs.Values.SelectMany(Function(i) i).ToList
                Dim inputmObj As MetaObject = inputObj
                Dim ext = inputmObj.GetAttribute("Extension").Value.ToString.ToLower
                If inputmObj.GetTgmType <> "File" OrElse (ext <> "rvt" And ext <> "ifc") Then Continue For 'Throw New Exception("You must drag and drop file objects")
                launchTree.SmartAddNode("object.id", inputmObj.Id.ToString,,, inputmObj.Name)
                count += 1
            Next
            If count = 0 Then MessageBox.Show("The inputs type is not corresponding to the prodAction (ifc or revit in this case).")
        End If
        Return launchTree
    End Function

    <ActionMethod()>
    Public Function MyMethod(Files As MultipleElements) As ActionResult
        If Files.Source.Count = 0 Then
            Return New IgnoredActionResult("No Input File")
        End If
        ' Récupération des inputs
        Dim tgmFiles = GetInputAsMetaObjects(Files.Source).ToHashSet 'To get rid of duplicates


        'analysis
        CompareWithReferences(tgmFiles(0))

        Return New SucceededActionResult("Succeeded", System.Drawing.Color.Green, New List(Of Element) From {Files})
    End Function

    Private Sub CompareWithReferences(fileTgm As MetaObject)

        Dim myAudit As New IfcFileObject(fileTgm)
        Dim auditWb = myAudit.AuditWorkbook
        If auditWb Is Nothing Then Throw New Exception("Launch initial audit first")

        'Filter
        TempWorkspace.SmartAddTree( "ALL").Filter(False, New List(Of Workspace) From {myAudit.ScanWs}).RunSynchronously()

        'Get ref ws
        If myAudit.ProjRefWs Is Nothing Then Throw New Exception("No references information in Project")

#Region "Get buildings and storeys references"
        Dim fileTgmBuildingName = fileTgm.GetAttribute("TgmBuilding", True)?.Value?.ToString 'Recursively for revit case, to not define tgm buildings before initial audit
        If fileTgmBuildingName = "" Then Throw New Exception("Define TgmBuildings for this file first")

        Dim batObjRefs As New List(Of MetaObject)
        Dim refDico As New Dictionary(Of String, List(Of StoreyRefToCompare))
        For Each batName In Strings.Split(fileTgmBuildingName, "/_\")
            Dim batObjRef = myAudit.ProjRefWs.GetMetaObjectByName(batName).Where(Function(m) m.IsActive).FirstOrDefault
            If batObjRef Is Nothing OrElse batObjRef.GetChildren(ProjectReference.Constants.storeyRefName).ToList.Count = 0 Then Throw New Exception("TgmBuilding """ + batName + """ isn't referenced")
            batObjRefs.Add(batObjRef)

            Dim globalAltiList As New List(Of StoreyRefToCompare)
            Dim refBuildStoreys = batObjRef.GetChildren(ProjectReference.Constants.storeyRefName)
            For Each refBuildSto In refBuildStoreys
                Dim globalElevation = refBuildSto.GetAttribute("GlobalElevation").Value
                Dim localElevation = CommonFunctions.ConvertGlobalElevationToScanWs(globalElevation, myAudit.ScanWs)
                Dim stoRefToComp As New StoreyRefToCompare With {
                    .storeyName = refBuildSto.Name,
                    .globalStoreyElevation = Math.Round(globalElevation, 3),
                    .localStoreyElevation = Math.Round(localElevation, 3),
                    .found = False}
                globalAltiList.Add(stoRefToComp)
            Next
            refDico.Add(batObjRef.Name, globalAltiList)
        Next
#End Region

#Region "Get TgmBuildings origins"
        Dim batOrigins As New List(Of Attribute)
        For Each batObjRef In batObjRefs
            Dim batOrigin = batObjRef.Attributes.Where(Function(a) a.Name = "OriginPoint").FirstOrDefault
            If batOrigin Is Nothing Then Throw New Exception("TgmBuilding """ + batObjRef.Name + """ doesn't have Goereferencing information")
            batOrigins.Add(batOrigin)
        Next
        Dim xVals, yVals, zVals, angleVals As New HashSet(Of Double)
        For k = 0 To batOrigins.Count - 1
            Dim batOrigin = batOrigins(k)
            Dim xBat = batOrigin.GetAttribute("X")
            Dim yBat = batOrigin.GetAttribute("Y")
            Dim zBat = batOrigin.GetAttribute("Z")
            Dim angleBat = batOrigin.GetAttribute("Angle")
            If xBat Is Nothing OrElse yBat Is Nothing OrElse zBat Is Nothing OrElse angleBat Is Nothing OrElse
                xBat.Value Is Nothing OrElse yBat.Value Is Nothing OrElse zBat.Value Is Nothing OrElse angleBat.Value Is Nothing Then
                Throw New Exception("TgmBuilding """ + batObjRefs(k).Name + """ doesn't have Goereferencing information")
                Exit Sub
            Else
                'xVals.Add(Math.Round(xBat.Value, 2)) 'Rounded to cm
                'yVals.Add(Math.Round(yBat.Value, 2)) 'Rounded to cm
                'zVals.Add(Math.Round(zBat.Value, 2)) 'Rounded to cm
                'angleVals.Add(Math.Round(angleBat.Value, 2)) 'Rounded to cm
                xVals.Add(xBat.Value)
                yVals.Add(yBat.Value)
                zVals.Add(zBat.Value)
                angleVals.Add(angleBat.Value)
            End If
        Next
        If xVals.Count <> 1 Or yVals.Count <> 1 Or zVals.Count <> 1 Or angleVals.Count <> 1 Then
            'myAudit.ReportInfos.Xref = Double.NaN
            'myAudit.ReportInfos.Yref = Double.NaN
            'myAudit.ReportInfos.Zref = Double.NaN
            'myAudit.ReportInfos.AngleRef = Double.NaN
            myAudit.ReportInfos.refGeoRef = New GeoReferencing() With {.X = Double.NaN, .Y = Double.NaN, .Z = Double.NaN, .Angle = Double.NaN}
        Else
            'myAudit.ReportInfos.Xref = xVals.First
            'myAudit.ReportInfos.Yref = yVals.First
            'myAudit.ReportInfos.Zref = zVals.First
            'myAudit.ReportInfos.AngleRef = angleVals.First
            myAudit.ReportInfos.refGeoRef = New GeoReferencing() With {.X = xVals.First, .Y = yVals.First, .Z = zVals.First, .Angle = angleVals.First}

        End If
#End Region

#Region "Get grids references"
        Dim allGrids As New List(Of MetaObject)
        For Each oRef As MetaObject In myAudit.ProjRefWs.MetaObjects
            If oRef.GetAttribute("object.type") IsNot Nothing AndAlso oRef.GetAttribute("object.type").Value = "FileGrid" Then
                allGrids.Add(oRef)
            End If
        Next
#End Region

        'Get audit file georef <-- Aller lire les attributs plutôt !!!
        Dim auditFileGeoRef = myAudit.Georeferencing
        'Dim translateGeo As Double() = Nothing
        'Dim rotateGeo As Double? = Nothing
        'CommonFunctions.GetGeodata(myAudit.ScanWs, rotateGeo, translateGeo)
        'If translateGeo IsNot Nothing AndAlso rotateGeo IsNot Nothing Then
        '    myAudit.ReportInfos.X = Math.Round(translateGeo(0), 2) 'Rounded to cm
        '    myAudit.ReportInfos.Y = Math.Round(translateGeo(1), 2) 'Rounded to cm
        '    myAudit.ReportInfos.Z = Math.Round(translateGeo(2), 2) 'Rounded to cm
        '    myAudit.ReportInfos.Angle = Math.Round(CDbl(rotateGeo), 2) 'Rounded to cm
        'Else
        '    Throw New Exception("File miss Georeferencing information")
        'End If

#Region "FILL EXCEL"
        'Compare audit excel with references
        myAudit.ReportInfos.lvlNOK = 0
        myAudit.ReportInfos.lvlREV = 0

#Region "IFC"

        If myAudit.ExtensionType = "ifc" Then

            'Complete ifc buildings and levels
            Dim structureWsheet = FileObject.GetWorksheet(auditWb, "STRUCTURE IFC")
            If structureWsheet Is Nothing Then Throw New Exception("worksheet ""STRUCTURE IFC"" missing in audit!")
            structureWsheet.Columns("D").Visible = True
            structureWsheet.Columns("E").Visible = True
            structureWsheet.Columns("G").Visible = True
            structureWsheet.Columns("H").Visible = True
            structureWsheet.Columns("J").Visible = True
            structureWsheet.Columns("K").Visible = True
            Dim lastRow = structureWsheet.GetUsedRange.BottomRowIndex 'Last storey row
            Dim myRow As Integer
            Dim currentTgmBuildingName As String = Nothing
            For myRow = 5 To lastRow + 1

                If Not String.IsNullOrWhiteSpace(structureWsheet.Range("C" + myRow.ToString).Value.TextValue) Then 'if it is a building row, compare building names
                    Dim ifcBuilding = structureWsheet.Range("C" + myRow.ToString).Value.TextValue
                    Dim correspondingTgmBuilding = refDico.Keys.FirstOrDefault(Function(o) o.ToLower = ifcBuilding.ToLower)
                    If correspondingTgmBuilding Is Nothing Then 'If not the same name, but...
                        If Strings.Split(fileTgmBuildingName, "/_\").Count = 1 Then 'Only one TgmBuilding corresponds to this file
                            currentTgmBuildingName = Strings.Split(fileTgmBuildingName, "/_\")(0)
                        ElseIf myAudit.SiteBuildings.Count = 1 Then 'Buildings reordering case
                            currentTgmBuildingName = fileTgmBuildingName
                        Else
                            currentTgmBuildingName = Nothing
                        End If
                        If currentTgmBuildingName IsNot Nothing Then structureWsheet.Range("D" + myRow.ToString).Value = currentTgmBuildingName
                        structureWsheet.Range("E" + myRow.ToString).Value = "NOK"
                    Else
                        currentTgmBuildingName = correspondingTgmBuilding
                        structureWsheet.Range("D" + myRow.ToString).Value = currentTgmBuildingName
                        structureWsheet.Range("E" + myRow.ToString).Value = "OK"
                    End If

                ElseIf Not String.IsNullOrWhiteSpace(structureWsheet.Range("F" + myRow.ToString).Value.TextValue) Then 'if it is a storey row, compare storey names and elevations
                    If currentTgmBuildingName = Nothing Then
                        Continue For
                    End If
                    Dim currentStorey As String = structureWsheet.Range("F" + myRow.ToString).Value.TextValue

                    'Get corresponding TgmBuildingStorey
                    Dim correspondingTgmStorey As StoreyRefToCompare = Nothing
                    If Strings.Split(currentTgmBuildingName, "/_\").Count > 1 Then
                        For Each buildName In Strings.Split(currentTgmBuildingName, "/_\") 'Buildings reordering case
                            correspondingTgmStorey = refDico(buildName).FirstOrDefault(Function(o) o.storeyName.ToLower = currentStorey.ToLower)
                            If correspondingTgmStorey IsNot Nothing Then
                                structureWsheet.Range("D" + myRow.ToString).Value = buildName
                                Exit For
                            End If
                        Next
                    Else
                        correspondingTgmStorey = refDico(currentTgmBuildingName).FirstOrDefault(Function(o) o.storeyName.ToLower = currentStorey.ToLower)
                    End If


                    If correspondingTgmStorey Is Nothing Then
                        structureWsheet.Range("H" + myRow.ToString).Value = "REV"
                        myAudit.ReportInfos.lvlREV += 1
                    Else
                        correspondingTgmStorey.found = True
                        structureWsheet.Range("G" + myRow.ToString).Value = correspondingTgmStorey.storeyName
                        structureWsheet.Range("H" + myRow.ToString).Value = "OK"

                        'Altimetry comparison
                        Dim elevation As Double = structureWsheet.Range("I" + myRow.ToString).Value.NumericValue
                        structureWsheet.Range("J" + myRow.ToString).Value = correspondingTgmStorey.globalStoreyElevation
                        If elevation <> correspondingTgmStorey.globalStoreyElevation Then
                            structureWsheet.Range("K" + myRow.ToString).Value = "NOK"
                            myAudit.ReportInfos.lvlNOK += 1
                        Else
                            structureWsheet.Range("K" + myRow.ToString).Value = "OK"
                        End If
                    End If
                End If
            Next

            'Add missing storey references
            Dim missingRefSto = refDico.SelectMany(Function(o) o.Value).Where(Function(l) l.found = False).ToList
            If missingRefSto.Count > 0 Then
                For Each refSto In missingRefSto
                    structureWsheet.Range("G" + myRow.ToString).Value = refSto.storeyName
                    structureWsheet.Range("H" + myRow.ToString).Value = "NOK"
                    myAudit.ReportInfos.lvlNOK += 1
                    myRow += 1
                Next
            End If

            '---Mise en forme
            structureWsheet.Columns("D").AutoFit()
            'myAudit.ReportInfos.CompleteCriteria(52, True) 'Levels

            'Compare georef
            If Not Double.IsNaN(myAudit.ReportInfos.refGeoRef.X) Then
                Dim fileWsheet = FileObject.GetWorksheet(auditWb, "FICHIER")
                If fileWsheet Is Nothing Then Throw New Exception("worksheet ""FICHIER"" missing in audit!")
                'fileWsheet.Columns("G").Visible = True
                'fileWsheet.Columns("H").Visible = True
                Dim georefRange = fileWsheet.Range("georef")
                myAudit.ReportInfos.goodGeoRef = True

                fileWsheet.Cells(georefRange.BottomRowIndex + 1, georefRange.LeftColumnIndex + 2).Value = myAudit.ReportInfos.refGeoRef.Y
                fileWsheet.Cells(georefRange.BottomRowIndex + 1, georefRange.LeftColumnIndex + 2).Alignment.Horizontal = SpreadsheetHorizontalAlignment.Right
                'fileWsheet.Cells(georefRange.BottomRowIndex + 1, georefRange.Column + 3).formulaR1C1 = "=IF(RC[-1]="""","""",IF(RC[-2]=RC[-1],""OK"",""NOK""))"
                If Math.Abs(myAudit.ReportInfos.refGeoRef.Y - auditFileGeoRef.Y) < 0.01 Then 'BETTER WAY TO COMPARE COORDINATES !!!
                    fileWsheet.Cells(georefRange.BottomRowIndex + 1, georefRange.LeftColumnIndex + 3).Value = "OK"
                Else
                    fileWsheet.Cells(georefRange.BottomRowIndex + 1, georefRange.LeftColumnIndex + 3).Value = "NOK"
                    myAudit.ReportInfos.goodGeoRef = False
                End If

                fileWsheet.Cells(georefRange.BottomRowIndex + 2, georefRange.LeftColumnIndex + 2).Value = myAudit.ReportInfos.refGeoRef.X
                fileWsheet.Cells(georefRange.BottomRowIndex + 2, georefRange.LeftColumnIndex + 2).Alignment.Horizontal = SpreadsheetHorizontalAlignment.Right
                If Math.Abs(myAudit.ReportInfos.refGeoRef.X - auditFileGeoRef.X) < 0.01 Then
                    fileWsheet.Cells(georefRange.BottomRowIndex + 2, georefRange.LeftColumnIndex + 3).Value = "OK"
                Else
                    fileWsheet.Cells(georefRange.BottomRowIndex + 2, georefRange.LeftColumnIndex + 3).Value = "NOK"
                    myAudit.ReportInfos.goodGeoRef = False
                End If

                fileWsheet.Cells(georefRange.BottomRowIndex + 3, georefRange.LeftColumnIndex + 2).Value = myAudit.ReportInfos.refGeoRef.Z
                fileWsheet.Cells(georefRange.BottomRowIndex + 3, georefRange.LeftColumnIndex + 2).Alignment.Horizontal = SpreadsheetHorizontalAlignment.Right
                If Math.Abs(myAudit.ReportInfos.refGeoRef.Z - auditFileGeoRef.Z) < 0.01 Then
                    fileWsheet.Cells(georefRange.BottomRowIndex + 3, georefRange.LeftColumnIndex + 3).Value = "OK"
                Else
                    fileWsheet.Cells(georefRange.BottomRowIndex + 3, georefRange.LeftColumnIndex + 3).Value = "NOK"
                    myAudit.ReportInfos.goodGeoRef = False
                End If

                fileWsheet.Cells(georefRange.BottomRowIndex + 4, georefRange.LeftColumnIndex + 2).Value = myAudit.ReportInfos.refGeoRef.Angle
                fileWsheet.Cells(georefRange.BottomRowIndex + 4, georefRange.LeftColumnIndex + 2).Alignment.Horizontal = SpreadsheetHorizontalAlignment.Right
                If Math.Abs(myAudit.ReportInfos.refGeoRef.Angle - auditFileGeoRef.Angle) < 0.01 Then
                    fileWsheet.Cells(georefRange.BottomRowIndex + 4, georefRange.LeftColumnIndex + 3).Value = "OK"
                Else
                    fileWsheet.Cells(georefRange.BottomRowIndex + 4, georefRange.LeftColumnIndex + 3).Value = "NOK"
                    myAudit.ReportInfos.goodGeoRef = False
                End If
            End If
            myAudit.ReportInfos.CompleteCriteria(1, True)

#End Region

#Region "RVT"

        ElseIf myAudit.ExtensionType = "rvt" Then
            '///////////////Compare GeoRef
            If Math.Abs(myAudit.ReportInfos.refGeoRef.Y - auditFileGeoRef.Y) > 0.01 OrElse
                Math.Abs(myAudit.ReportInfos.refGeoRef.X - auditFileGeoRef.X) > 0.01 OrElse
                Math.Abs(myAudit.ReportInfos.refGeoRef.Z - auditFileGeoRef.Z) > 0.01 Then

                myAudit.ReportInfos.goodGeoRef = False
            Else
                myAudit.ReportInfos.goodGeoRef = True
            End If

            Dim fileWsheet = FileObject.GetWorksheet(auditWb, "METRIQUES")
            If fileWsheet Is Nothing Then Throw New Exception("worksheet ""METRIQUES"" missing in audit!")

            'remplissage de l'excel : georef
            Dim refCol = 3
            Dim refLine = 11
            'fileWsheet.Cells(refLine, refCol).Value = auditFileGeoRef.Y
            fileWsheet.Cells(refLine, refCol).Value = myAudit.ReportInfos.refGeoRef.Y
            If Math.Abs(myAudit.ReportInfos.refGeoRef.Y - auditFileGeoRef.Y) < 0.01 Then
                fileWsheet.Cells(refLine, refCol + 1).Value = "OK"
            Else
                fileWsheet.Cells(refLine, refCol + 1).Value = "NOK"
            End If
            'fileWsheet.Cells(refLine + 1, refCol).Value = auditFileGeoRef.X
            fileWsheet.Cells(refLine + 1, refCol).Value = myAudit.ReportInfos.refGeoRef.X
            If Math.Abs(myAudit.ReportInfos.refGeoRef.X - auditFileGeoRef.X) < 0.01 Then
                fileWsheet.Cells(refLine + 1, refCol + 1).Value = "OK"
            Else
                fileWsheet.Cells(refLine + 1, refCol + 1).Value = "NOK"
            End If
            'fileWsheet.Cells(refLine + 2, refCol).Value = auditFileGeoRef.Z
            fileWsheet.Cells(refLine + 2, refCol).Value = myAudit.ReportInfos.refGeoRef.Z
            If Math.Abs(myAudit.ReportInfos.refGeoRef.Z - auditFileGeoRef.Z) < 0.01 Then
                fileWsheet.Cells(refLine + 2, refCol + 1).Value = "OK"
            Else
                fileWsheet.Cells(refLine + 2, refCol + 1).Value = "NOK"
            End If
            fileWsheet.Cells(refLine + 3, refCol).Value = myAudit.ReportInfos.refGeoRef.Angle * 180.0 / Math.PI 'convert to degrees
            If Math.Abs(auditFileGeoRef.Angle - myAudit.ReportInfos.refGeoRef.Angle) < 1.0 Then
                fileWsheet.Cells(refLine + 3, refCol + 1).Value = "OK"
            Else
                fileWsheet.Cells(refLine + 3, refCol + 1).Value = "NOK"
            End If
            fileWsheet.Columns(refCol).AutoFit()
            fileWsheet.Columns(refCol + 1).AutoFit()


            'Check BaseProject and Survey
            Dim baseProject = myAudit.ScanWs.GetMetaObjects("BasePoint", Nothing, False).FirstOrDefault
            Dim survey = myAudit.ScanWs.GetMetaObjects("SurveyPoint", Nothing, False).FirstOrDefault
            Dim refBases = batObjRefs.Select(Function(o) o.GetAttribute("BasePoint")).ToHashSet
            Dim refSurveys = batObjRefs.Select(Function(o) o.GetAttribute("SurveyPoint")).ToHashSet



            If refBases IsNot Nothing AndAlso refBases.Count = 1 AndAlso refBases.First IsNot Nothing Then
                'remplissage de l'excel : basePoint
                refCol = 3
                refLine = 6
                Dim refBase = refBases.First
                If baseProject Is Nothing Then
                    myAudit.ReportInfos.goodBasePt = False
                Else
                    myAudit.ReportInfos.goodBasePt = True
                    Dim myLine = 0
                    For Each att In baseProject.GetAttribute("SharedPosition").Attributes
                        If myLine = 0 Then 'correction cohérence excel/beemer
                            att = baseProject.GetAttribute("SharedPosition").Attributes.Item(1)
                        ElseIf myLine = 1 Then
                            att = baseProject.GetAttribute("SharedPosition").Attributes.Item(0)
                        End If
                        fileWsheet.Cells(refLine + myLine, refCol).Value = CDbl(refBase.GetAttribute(att.Name).Value)
                        If Math.Abs(refBase.GetAttribute(att.Name).Value - att.Value) < 0.01 Then 'BETTER WAY TO COMPARE COORDINATES !!!
                            fileWsheet.Cells(refLine + myLine, refCol + 1).Value = "OK"
                        Else
                            fileWsheet.Cells(refLine + myLine, refCol + 1).Value = "NOK"
                            myAudit.ReportInfos.goodBasePt = False
                        End If
                        'If refBase.GetAttribute(att.Name).Value <> att.Value Then
                        '    myAudit.ReportInfos.goodBasePt = False
                        '    Exit For
                        'End If
                        myLine += 1
                    Next
                End If
            End If
            If refSurveys IsNot Nothing AndAlso refSurveys.Count = 1 AndAlso refSurveys.First IsNot Nothing Then
                refCol = 3
                refLine = 3

                Dim refSurvey = refSurveys.First
                If survey Is Nothing Then
                    myAudit.ReportInfos.goodBasePt = False
                Else
                    myAudit.ReportInfos.goodSurveyPt = True
                    Dim myLine = 0
                    For Each att In survey.GetAttribute("SharedPosition").Attributes
                        If myLine = 0 Then 'correction cohérence excel/beemer
                            att = survey.GetAttribute("SharedPosition").Attributes.Item(1)
                        ElseIf myLine = 1 Then
                            att = survey.GetAttribute("SharedPosition").Attributes.Item(0)
                        End If
                        fileWsheet.Cells(refLine + myLine, refCol).Value = CDbl(refSurvey.GetAttribute(att.Name).Value)
                        If Math.Abs(refSurvey.GetAttribute(att.Name).Value - att.Value) < 0.01 Then 'BETTER WAY TO COMPARE COORDINATES !!!
                            fileWsheet.Cells(refLine + myLine, refCol + 1).Value = "OK"
                        Else
                            fileWsheet.Cells(refLine + myLine, refCol + 1).Value = "NOK"
                            myAudit.ReportInfos.goodSurveyPt = False
                        End If
                        myLine += 1
                        'If refSurvey.GetAttribute(att.Name).Value <> att.Value Then
                        '    myAudit.ReportInfos.goodSurveyPt = False
                        '    Exit For
                        'End If
                    Next
                End If
            End If

            myAudit.ReportInfos.CompleteCriteria(21, True)
            myAudit.ReportInfos.CompleteCriteria(22, True)
            myAudit.ReportInfos.CompleteCriteria(23, True)


            '///////////////Compare Grids
            myAudit.ReportInfos.goodGrids = True 'by default
            myAudit.ReportInfos.gridNOK = 0

            Dim strWsheet = FileObject.GetWorksheet(auditWb, "STRUCTURATION")
            If strWsheet Is Nothing Then Throw New Exception("worksheet ""STRUCTURATION"" missing in audit!")

            Dim gridRange = strWsheet.Range("QUAD")
            Dim gridColumn = gridRange.LeftColumnIndex

            Dim lastRow = strWsheet.GetUsedRange().RowCount

            Dim dicoA = New Dictionary(Of String, Integer)
            Dim listB = New List(Of String)

            For i = 5 To lastRow - 1
                'Column grids audit
                Dim auditGrid As String = strWsheet.Cells(i, gridColumn).Value.ToString
                If auditGrid.Contains("NOT FOUND") Then
                    strWsheet.Cells(i, gridColumn).Value = ""
                End If

                'Column grids ref
                Dim refGrid As String = strWsheet.Cells(i, gridColumn + 1).Value.ToString
                If Not String.IsNullOrWhiteSpace(refGrid) Then
                    strWsheet.Cells(i, gridColumn + 1).Value = ""
                    strWsheet.Cells(i, gridColumn + 2).Value = ""
                End If
            Next

            For i = 5 To lastRow - 1
                Dim auditGrid As String = strWsheet.Cells(i, gridColumn).Value.ToString
                Dim splitAuditName() As String = Split(auditGrid, "(")
                Dim nameAuditGrid As String = splitAuditName.ElementAt(0)

                If Not String.IsNullOrWhiteSpace(nameAuditGrid) Then
                    dicoA.Add(nameAuditGrid, i)

                    For Each oGrid As MetaObject In allGrids
                        If oGrid.GetAttribute("AuditAlias") Is Nothing Then Throw New Exception("You need to create grids references.")

                        Dim oValue As String = oGrid.GetAttribute("AuditAlias").Value
                        Dim splitoValue() As String = Split(oValue, "(")
                        Dim nameoGrid As String = splitoValue.ElementAt(0)

                        If Not String.IsNullOrWhiteSpace(nameoGrid) Then
                            listB.Add(nameoGrid)
                        End If

                        If nameAuditGrid = nameoGrid Then

                            strWsheet.Cells(i, gridColumn + 1).Value = oValue 'écrire la valeur

                            If auditGrid = oValue Then 'same value ?
                                strWsheet.Cells(i, gridColumn + 2).Value = "OK"
                            Else
                                strWsheet.Cells(i, gridColumn + 2).Value = "NOK"

                                myAudit.ReportInfos.goodGrids = False
                                myAudit.ReportInfos.gridNOK += 1
                            End If

                            GoTo nextGrid
                        End If
                    Next
                End If
nextGrid:
            Next

            Dim listA As New List(Of String)
            For Each oDico In dicoA
                listA.Add(oDico.Key)
            Next

            Dim items1() As String = listA.ToArray
            Dim items2() As String = listB.ToArray
            Dim list3 As New List(Of String)
            Dim list4 As New List(Of String)
            list3.AddRange(items2.Except(items1).ToArray)
            list4.AddRange(items1.Except(items2).ToArray)

            Dim j As Integer = 0
            If list3.Count = 0 Then
                'all good
            Else
                For Each except In list3

                    For Each oGrid As MetaObject In allGrids
                        Dim oValue As String = oGrid.GetAttribute("AuditAlias").Value
                        Dim splitoValue() As String = Split(oValue, "(")
                        Dim nameoGrid = splitoValue.ElementAt(0)

                        If nameoGrid = except Then
                            Dim writingRow = listA.Count + 5 + j
                            strWsheet.Cells(writingRow, gridColumn).Value = "NOT FOUND"
                            strWsheet.Cells(writingRow, gridColumn + 1).Value = oValue
                            strWsheet.Cells(writingRow, gridColumn + 2).Value = "NOK"

                            myAudit.ReportInfos.goodGrids = False
                            myAudit.ReportInfos.gridNOK += 1

                            j = j + 1
                            GoTo nextExceptGrid
                        End If

                    Next
nextExceptGrid:
                Next
            End If

            If list4.Count = 0 Then
                'all good
            Else
                For Each except In list4
                    For Each itemA In dicoA
                        If itemA.Key = except Then
                            strWsheet.Cells(itemA.Value, gridColumn + 1).Value = "NOT FOUND"
                            strWsheet.Cells(itemA.Value, gridColumn + 2).Value = "NOK"

                            myAudit.ReportInfos.goodGrids = False
                            myAudit.ReportInfos.gridNOK += 1
                        End If
                    Next
                Next
            End If

            strWsheet.Columns(gridColumn + 1).AutoFit()
            strWsheet.Columns(gridColumn + 2).AutoFit()

            myAudit.ReportInfos.CompleteCriteria(24, True)


            '///////////////compare Levels
            Dim listOfRefLvl As New List(Of StoreyRefToCompare)
            myAudit.ReportInfos.goodLvls = True

            'get reflevels from tgm
            If fileTgmBuildingName.Contains(";") Then
                Dim splitTgmBuilding() As String = Split(fileTgmBuildingName, ";")

                For Each fileTgmB In splitTgmBuilding 'pour chaque tgmbuilding
                    For Each refItem In refDico 'pour chaque bât ref
                        If refItem.Key = fileTgmB Then 'if same, add levels
                            listOfRefLvl.AddRange(refItem.Value)
                        End If
                    Next
                Next

            Else
                For Each refItem In refDico
                    If refItem.Key = fileTgmBuildingName Then
                        listOfRefLvl.AddRange(refItem.Value)
                    End If
                Next
            End If

            'dans l'audit
            Dim lvlRange = strWsheet.Range("LEVELS")
            Dim lvlColumn = lvlRange.LeftColumnIndex

            Dim dicoLvl As New Dictionary(Of String, Integer)
            Dim listRef = New List(Of String)

            'RESET 
            For i = 5 To lastRow - 1
                'Column lvls audit
                Dim auditLvl As String = strWsheet.Cells(i, lvlColumn).Value.ToString
                If auditLvl.Contains("NOT FOUND") Then
                    strWsheet.Cells(i, lvlColumn).Value = ""
                End If

                'Column lvls ref
                Dim refLvl As String = strWsheet.Cells(i, lvlColumn + 1).Value.ToString
                If Not String.IsNullOrWhiteSpace(refLvl) Then
                    strWsheet.Cells(i, lvlColumn + 1).Value = ""
                    strWsheet.Cells(i, lvlColumn + 2).Value = ""
                End If
            Next

            For i = 5 To lastRow - 1
                Dim auditLvl As String = strWsheet.Cells(i, lvlColumn).Value.ToString
                Dim splitAuditName() As String = Split(auditLvl, "(")
                Dim nameAuditLvl As String = splitAuditName.ElementAt(0).Replace(" ", "")
                Dim valueAuditLvl As String = splitAuditName.ElementAt(splitAuditName.Count - 1)

                Dim valueLvlDB As Double

                If Not String.IsNullOrWhiteSpace(nameAuditLvl) Then
                    Dim revisedValue = valueAuditLvl.Replace("(", "")
                    Dim rerevisedValue = revisedValue.Replace(")", "")
                    valueLvlDB = Convert.ToDouble(rerevisedValue)

                    dicoLvl.Add(nameAuditLvl, i)

                    For Each refLevel In listOfRefLvl

                        Dim refName As String = refLevel.storeyName.Replace(" ", "")
                        Dim refValue As String = refLevel.localStoreyElevation
                        listRef.Add(refName)

                        If nameAuditLvl = refName Then

                            strWsheet.Cells(i, lvlColumn + 1).Value = refName + " (" + refValue.ToString + ")"

                            If valueLvlDB = refValue Then
                                strWsheet.Cells(i, lvlColumn + 2).Value = "OK"
                            Else
                                strWsheet.Cells(i, lvlColumn + 2).Value = "NOK"

                                myAudit.ReportInfos.goodLvls = False
                                myAudit.ReportInfos.lvlNOK += 1
                            End If
                            GoTo nextLvl

                        End If
                    Next
                End If
nextLvl:
            Next

            Dim listAudit As New List(Of String)
            For Each oLvl In dicoLvl
                listAudit.Add(oLvl.Key)
            Next

            Dim itemsAudit() As String = listAudit.ToArray
            Dim itemsRef() As String = listRef.ToArray
            Dim list5 As New List(Of String)
            Dim list6 As New List(Of String)
            list5.AddRange(itemsRef.Except(itemsAudit).ToArray)
            list6.AddRange(itemsAudit.Except(itemsRef).ToArray)


            Dim k As Integer = 0
            If list5.Count = 0 Then
                'all good
            Else
                For Each except In list5

                    For Each refLvl In listOfRefLvl
                        If refLvl.storeyName = except Then
                            Dim writingRow = listAudit.Count + 5 + k
                            strWsheet.Cells(writingRow, lvlColumn).Value = "NOT FOUND"
                            strWsheet.Cells(writingRow, lvlColumn + 1).Value = refLvl.storeyName + " (" + refLvl.localStoreyElevation.ToString + ")"
                            strWsheet.Cells(writingRow, lvlColumn + 2).Value = "NOK"

                            myAudit.ReportInfos.goodLvls = False
                            myAudit.ReportInfos.lvlNOK += 1

                            k = k + 1
                            GoTo nextExceptLvl
                        End If

                    Next
nextExceptLvl:
                Next
            End If

            If list6.Count = 0 Then
                'all good
            Else
                For Each except In list6
                    For Each oLvl In dicoLvl
                        If oLvl.Key = except Then
                            strWsheet.Cells(oLvl.Value, lvlColumn + 1).Value = "NOT FOUND"
                            strWsheet.Cells(oLvl.Value, lvlColumn + 2).Value = "NOK"

                            myAudit.ReportInfos.goodLvls = False
                            myAudit.ReportInfos.lvlNOK += 1
                        End If
                    Next
                Next
            End If

            strWsheet.Columns(lvlColumn + 1).AutoFit()
            strWsheet.Columns(lvlColumn + 2).AutoFit()
        End If
#End Region

#End Region

        'Common to ifc and rvt
        myAudit.ReportInfos.CompleteCriteria(51, True) 'North direction
        myAudit.ReportInfos.CompleteCriteria(52, True) 'Levels

        ''Hide reference rows - A cause de l'audit revit qui ne le fait pas.... à sup à terme !
        'Dim lastRow2 As Integer = myAudit.ReportWorksheet.GetUsedRange.BottomRowIndex
        'For i = 1 To lastRow2
        '    Dim refCell = myAudit.ReportWorksheet.Range("D" + i.ToString)
        '    If refCell.Value IsNot Nothing AndAlso refCell.Value = "Yes" Then
        '        myAudit.ReportWorksheet.Rows(refCell.BottomRowIndex).Visible = False
        '    End If
        'Next

        myAudit.ProjWs.PushAllModifiedEntities()
        'myAudit.ReportInfos.ShowRefReportRows()
        myAudit.AuditWorkbook.Calculate()
        myAudit.AuditWorkbook.Worksheets.ActiveWorksheet = myAudit.ReportWorksheet
        myAudit.SaveAuditWorkbook()

    End Sub

    Private Class StoreyRefToCompare
        Public Property storeyName As String
        Public Property globalStoreyElevation As Double
        Public Property localStoreyElevation As Double
        Public Property found As Boolean
    End Class



End Class
