Imports M4D.Treegram.Core.Entities
Imports M4D.Treegram.Core.Extensions.Entities
Imports M4D.Treegram.Core.Extensions.Kernel
Imports System.IO
Imports System.Reflection
Imports DevExpress.Spreadsheet
Imports Treegram.ConstructionManagement

Public Class FileObject
    Public Sub New(fileTgm As MetaObject)
        Me.FileTgm = fileTgm
        SettingsWorkbook = Nothing
        ReportInfos = New ReportInfos(Me)

        'Get language...

    End Sub

    Public ReadOnly Property ReportInfos As ReportInfos

#Region "Common Objects"
    Public Property FileTgm As MetaObject
    Public ReadOnly Property ExtensionType As String
        Get
            Return FileTgm.GetAttribute("Extension").Value.ToString().ToLower
        End Get
    End Property
    Public ReadOnly Property ScanWs As Workspace
        Get
            Return CType(FileTgm.Container, Workspace)
        End Get
    End Property
    Public ReadOnly Property ProjWs As Workspace
        Get
            Return FileTgm.GetProject
        End Get
    End Property
    Public ReadOnly Property AuditTreesWs As Workspace
        Get
            Return ProjWs.SmartAddWorkspace("Audit Trees")
        End Get
    End Property
    Public ReadOnly Property DictionaryWs As Workspace
        Get
            Return ProjWs.SmartAddWorkspace("Dictionary")
        End Get
    End Property

    Private _projRefWs As Workspace
    Public ReadOnly Property ProjRefWs As Workspace
        Get
            If _projRefWs Is Nothing Then
                _projRefWs = ProjWs.Workspaces.Where(Function(ws) ws.Name = M4D.Treegram.Core.Constants.WorkspaceName.TreegramProjectReferenceWorkspace).FirstOrDefault
            End If
            Return _projRefWs
        End Get
    End Property
    Private _achatWs As Workspace
    Public ReadOnly Property AchatWs As Workspace
        Get
            If _achatWs Is Nothing Then
                _achatWs = ProjWs.Workspaces.Where(Function(ws) ws.Name = ProjectReference.Constants.projectAchatWsName).FirstOrDefault
            End If
            Return _achatWs
        End Get
    End Property
    Public ReadOnly Property LinkedWsList As List(Of Workspace)
        Get
            Return DataWsList("LinkedData")
        End Get
    End Property
    Public ReadOnly Property GhostWsList As List(Of Workspace)
        Get
            Return DataWsList("GhostData")
        End Get
    End Property
    Public ReadOnly Property AdvancedWsList As List(Of Workspace)
        Get
            Return DataWsList("AdvancedData")
        End Get
    End Property
    Private Function DataWsList(dataName As String) As List(Of Workspace)
        Dim linkedList As New List(Of Workspace)
        For Each ws In ProjWs.Workspaces
            If ws.GetAttribute(dataName)?.Value IsNot Nothing AndAlso ws.GetAttribute(dataName).Value.ToString = ScanWs.Id.ToString Then
                linkedList.Add(ws)
            End If
        Next
        Return linkedList
    End Function

    Public Sub CreatePsetTgmExtensions(objsToExtend As List(Of MetaObject))
        'Create PsetTgm extensions for those input objects
        For Each objTgm In objsToExtend
            If Not CType(objTgm.Container, Workspace).Equals(ScanWs) Then
                Continue For
            End If
            Dim extName = $"{objTgm} - {ProjectReference.Constants.psetTgmWsType}"
            Dim extObj = objTgm.Extensions.FirstOrDefault(Function(o) o.Name = extName)
            If extObj IsNot Nothing Then
                'S'il existe actif ou non
                extObj.IsActive = True
            ElseIf PsetTgmWs.MetaObjects.FirstOrDefault(Function(o) o.Name = extName AndAlso o.Extend IsNot Nothing AndAlso o.Extend.Equals(objTgm)) IsNot Nothing Then
                'Si les jonctions ne sont pas faites (Un peu plus long comme requête) <-- Peut être inutile !!!!!
                extObj = PsetTgmWs.MetaObjects.FirstOrDefault(Function(o) o.Name = extName AndAlso o.Extend IsNot Nothing AndAlso o.Extend.Equals(objTgm))
                extObj.IsActive = True
            ElseIf PsetTgmWs.MetaObjects.FirstOrDefault(Function(o) Not o.IsActive AndAlso o.Name = extName) IsNot Nothing Then
                'Si l'objet existe mais désactivé et sans aucun lien
                extObj = PsetTgmWs.MetaObjects.FirstOrDefault(Function(o) Not o.IsActive AndAlso o.Name = extName) 'It is supposed to be clean : no attribute, no relation, etc.
                extObj.IsActive = True
                extObj.Extend = objTgm
            Else
                'Sinon, création de l'extension
                extObj = PsetTgmWs.AddMetaObject(extName)
                extObj.Extend = objTgm
            End If
        Next
        ScanWs.Join({PsetTgmWs}.ToList) ' VERY IMPORTANT
        'ProjWs.SmartAddTree("ALL").Filter(False, {ScanWs}.ToList).RunSynchronously() ' VERY IMPORTANT

    End Sub

    Public ReadOnly Property PsetTgmWs As Workspace
        Get
            Return GhostWsList.FirstOrDefault
        End Get
    End Property
    Public ReadOnly Property TgmCreationWs As Workspace
        Get
            Return LinkedWsList.FirstOrDefault(Function(o) o.Name.Contains($" - {ProjectReference.Constants.tgmCreationWsType}"))
        End Get
    End Property
    Public ReadOnly Property ShellFunctionalModelWs As Workspace
        Get
            Return AdvancedWsList.FirstOrDefault(Function(o) o.Name.Contains(" - ShellFunctionalModel"))
        End Get
    End Property
    Public ReadOnly Property SpaceFunctionalModelWs As Workspace
        Get
            Return AdvancedWsList.FirstOrDefault(Function(o) o.Name.Contains(" - SpaceFunctionalModel"))
        End Get
    End Property
    Public ReadOnly Property AnalysisWs As Workspace
        Get
            Return AdvancedWsList.FirstOrDefault(Function(o) o.Name.Contains(" - Analysis"))
        End Get
    End Property
    Public ReadOnly Property _3dObjTgm As MetaObject
        Get
            Return FileTgm.GetChildren(, "3D").First
        End Get
    End Property
    Public ReadOnly Property _3dObjects As List(Of MetaObject)
        Get
            Return _3dObjTgm.GetChildren.ToList
        End Get
    End Property
    Public ReadOnly Property AddedObjs As Integer
        Get
            Return _3dObjTgm.GetAttribute("New", True).Value
        End Get
    End Property
    Public ReadOnly Property ModifiedObjs As Integer
        Get
            Return _3dObjTgm.GetAttribute("Modified", True).Value
        End Get
    End Property
    Public ReadOnly Property DeletedObjs As Integer
        Get
            Return _3dObjTgm.GetAttribute("Deleted", True).Value
        End Get
    End Property

    Public ReadOnly Property RefTgm As MetaObject
        Get
            Return FileTgm.GetChildren(, "REF").First
        End Get
    End Property
    Public ReadOnly Property North As GeographicNorth
        Get
            Dim northTgm = RefTgm.GetChildren(, "GeographicNorth").First
            Return New GeographicNorth(northTgm)
        End Get
    End Property
    Public ReadOnly Property Georeferencing As GeoReferencing 'LAMBERT 93
        Get
            Dim sharedPosiAtt = OriginPtTgm.GetAttribute("SharedPosition")
            Return New GeoReferencing() With {.X = CDbl(sharedPosiAtt.GetAttribute("E/W").Value),
                                              .Y = CDbl(sharedPosiAtt.GetAttribute("N/S").Value),
                                              .Z = CDbl(sharedPosiAtt.GetAttribute("Elevation").Value),
                                              .Angle = CDbl(sharedPosiAtt.GetAttribute("Angle").Value)}
        End Get
    End Property

    Public ReadOnly Property OriginPtTgm As MetaObject
        Get
            For Each childRel In RefTgm.Relations
                Dim childTgm = childRel.Target
                If childTgm.Name = "OriginPoint" Then
                    Return childTgm
                End If
            Next
            Return Nothing
        End Get
    End Property

    Public ReadOnly Property IsRevit As Boolean
        Get
            If ExtensionType = "rvt" Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property
    Public ReadOnly Property IsIfc As Boolean
        Get
            If ExtensionType = "ifc" Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property
#End Region


#Region "Xl Objects"

    'Private settingsTemplateWorkbookName = "Settings_Template.xlsx"
    Private _reportWorkbookName As String = "CQM_Template.xlsx"
    Private ReadOnly _reportWorksheetName As String = "RAPPORT"
    Private ReadOnly _warningWorkbookName As String = "Classification Avertissements Revit.xlsx"

    Private _auditWb As Workbook = Nothing
    Public Property AuditWorkbook As Workbook
        Get
            If _auditWb Is Nothing Then
                'Get file
                Dim projectFile = New FileInfo(Me.ProjWs.Path.LocalPath)
                Dim dirPath = Path.Combine(projectFile.Directory.FullName, "Audits", FileTgm.Name, Now.Year.ToString + "_W" + CalendarUtils.GetActualWeekOfyearAsString)
                Dim filePath = System.IO.Path.Combine(dirPath, "Audit_" + Now.Year.ToString + "_W" + CalendarUtils.GetActualWeekOfyearAsString + "_" + FileTgm.Name + ".xlsx")
                'Get workbook
                If File.Exists(filePath) Then
                    _auditWb = New Workbook
                    If Not _auditWb.LoadDocument(filePath) Then
                        Throw New Exception("This file cannot be opened :  " + filePath)
                    End If
                    _auditWb.BeginUpdate() 'A tester si vraiment utile !?
                Else
                    'NE PAS EFFACER - PROCESS POUR ALLER CHERCHER LAUDIT DE LA SEMAINE PRECEDENTE
                    'Dim n = 1
                    'Dim myWeek As Integer = 0
                    'Dim myYear As Integer = 0
                    'While _AuditWb Is Nothing And n <= 54

                    '    If CInt(TgmToolBox.GetActualWeekOfyearAsString) - n <= 0 Then
                    '        myWeek = CInt(TgmToolBox.GetActualWeekOfyearAsString) + 54
                    '        myYear = CInt(Now.Year.ToString) - 1
                    '    Else
                    '        myWeek = CInt(TgmToolBox.GetActualWeekOfyearAsString) - n
                    '        myYear = CInt(Now.Year.ToString)
                    '    End If
                    '    Dim lastdirPath = Path.Combine(projectFile.Directory.FullName, "Audits", FileTgm.Name, myYear.ToString + "_W" + myWeek.ToString)
                    '    Dim lastfilePath = System.IO.Path.Combine(lastdirPath, "Audit_" + myYear.ToString + "_W" + myWeek.ToString + "_" + FileTgm.Name + ".xlsx")

                    '    If File.Exists(lastfilePath) Then
                    '        _AuditWb = New Workbook
                    '        If Not _AuditWb.LoadDocument(lastfilePath) Then Continue While
                    '        _AuditWb.BeginUpdate() 'A tester si vraiment utile !?
                    '    End If
                    '    n += 1
                    'End While
                End If
            End If
            Return _auditWb
        End Get
        Set(value As Workbook)
            _auditWb = value
        End Set
    End Property
    Public Property WarningWorkbook As Workbook

    Private _settingsWb As Workbook = Nothing
    Private _settingsWbPath As String = Nothing
    Public Property SettingsWorkbook As Workbook
        Get
            If _settingsWb Is Nothing Then
                'Get project paths
                Dim projectFile = New FileInfo(Me.ProjWs.Path.LocalPath)
                Dim templatePath1 = Path.Combine(projectFile.Directory.FullName, "Audits", FileTgm.Name, "Settings", "Settings-" + FileTgm.Name + ".xlsx")
                'Dim templatePath2 = Path.Combine(projectFile.Directory.FullName, "Audits", "Settings-" + FileTgm.Name + ".xlsx") 'obsolète...
                'Load workbook
                Dim dirAuditPath = Path.Combine(projectFile.Directory.FullName, "Audits", FileTgm.Name)
                Dim dirAudit As New DirectoryInfo(dirAuditPath)

                If Not dirAudit.Exists Then
                    dirAudit.Create()
                End If

                If File.Exists(templatePath1) Then
                    _settingsWb = New Workbook
                    _settingsWbPath = templatePath1
                    If Not _settingsWb.LoadDocument(templatePath1) Then
                        Throw New Exception("This file cannot be opened : " + templatePath1)
                    End If
                    'ElseIf File.Exists(templatePath2) Then
                    '    _SettingsWb = New Workbook
                    '    _SettingsWbPath = templatePath2
                    '    If Not _SettingsWb.LoadDocument(templatePath2) Then
                    '        Throw New Exception("This file cannot be opened : " + templatePath2)
                    '    End If
                Else
                    Throw New Exception("Settings file doesn't exist")
                    ''If user didn't reference his file via Beemer Prod, create the default settings
                    'Dim defaultSettingsWbName As String
                    'If IsRevit Then
                    '    defaultSettingsWbName = "Revit_Architecture_Audit_WithoutRef_Français"
                    'Else
                    '    defaultSettingsWbName = "Ifc_Architecture_Audit_WithoutRef_Français"
                    'End If
                    'Dim exeAssemblyPath = Path.GetDirectoryName(Assembly.GetCallingAssembly().Location)

                    'Dim defaultSettingsPath As String
                    'If File.Exists(exeAssemblyPath + "\ressources\ProdAction_Selection\" + defaultSettingsWbName + ".xlsx") Then
                    '    'Debug mode
                    '    defaultSettingsPath = Path.Combine(exeAssemblyPath, "ressources", "ProdAction_Selection", defaultSettingsWbName + ".xlsx")
                    'Else
                    '    'Compiled mode
                    '    defaultSettingsPath = Path.Combine(exeAssemblyPath, "actions", "ressources", "ProdAction_Selection", defaultSettingsWbName + ".xlsx")
                    'End If

                    '_SettingsWb = New Workbook
                    'If Not _SettingsWb.LoadDocument(defaultSettingsPath) Then
                    '    Throw New Exception("This file cannot be opened : " + defaultSettingsPath)
                    'End If
                    '_SettingsWb.SaveDocument(templatePath1, DocumentFormat.OpenXml)
                    '_SettingsWbPath = templatePath1
                End If
            End If
            Return _settingsWb
        End Get
        Set(value As Workbook)
            _settingsWb = value
        End Set
    End Property
    'Commentaire Push
    Public ReadOnly Property SeuilWsheet As Worksheet
        Get
            If Me.IsIfc Then
                Return Me.GetOrInsertSettingsWSheet("VALEURS DE REFERENCE IFC")
            Else
                Return Me.GetOrInsertSettingsWSheet("VALEURS DE REFERENCE")
            End If
        End Get
    End Property
    Public ReadOnly Property ReportWorksheet As Worksheet
        Get
            Return GetWorksheet(Me.AuditWorkbook, _reportWorksheetName)
        End Get
    End Property


    ''' <summary>
    ''' Get report worksheet to create a new audit workbook
    ''' </summary>
    ''' <returns></returns>
    Public Function CreateAuditWorkbook() As Workbook


        _auditWb = New Workbook
        If SettingsWorkbook Is Nothing Then Throw New Exception("Settings file missing")
        Dim settingsReportWsheet = GetWorksheet(SettingsWorkbook, _reportWorksheetName)
        If settingsReportWsheet IsNot Nothing Then 'Get template in project files
            Dim newReportWsheet As Worksheet = _auditWb.Worksheets.FirstOrDefault
            If newReportWsheet Is Nothing Then
                _auditWb.Worksheets.Add(_reportWorksheetName) 'Pas utile...
            Else
                newReportWsheet.Name = _reportWorksheetName
            End If
            newReportWsheet.CopyFrom(settingsReportWsheet)
            ReportInfos.HideReportColumnsAndRows(newReportWsheet) 'Pour les anciens templates qui ne cachaient pas les colonnes L et M

        Else 'Get template in Tgm files
            'Get path
            If Me.Language = LanguageEnum.EN Then
                _reportWorkbookName = _reportWorkbookName.Replace(".xlsx", "-ENG.xlsx")
            End If
            Dim exeAssemblyPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
            Dim templatePath1 = exeAssemblyPath + "\actions\ressources\Report_Template\" + _reportWorkbookName 'Compiled mode
            Dim templatePath2 = exeAssemblyPath + "\ressources\Report_Template\" + _reportWorkbookName 'Debug mode
            'Get file
            If File.Exists(templatePath1) Then
                If Not _auditWb.LoadDocument(templatePath1) Then
                    Throw New Exception("This file cannot be opened : " + templatePath1)
                End If
            ElseIf File.Exists(templatePath2) Then
                If Not _auditWb.LoadDocument(templatePath2) Then
                    Throw New Exception("This file cannot be opened : " + templatePath2)
                End If
            End If
            'Get worksheet
            Dim reportWsheet = GetWorksheet(_auditWb, _reportWorksheetName)
            ReportInfos.HideReportColumnsAndRows(reportWsheet)



            'ADD IT TO FILE SETTINGS
            settingsReportWsheet = SettingsWorkbook.Worksheets.Insert(0, _reportWorksheetName)
            settingsReportWsheet.CopyFrom(reportWsheet)
            Try
                File.Delete(_settingsWbPath) 'delete to replace
            Catch ex As Exception
                Throw New Exception("Fermer l'XL settings concerné et relancer")
            End Try
            SettingsWorkbook.SaveDocument(_settingsWbPath, DocumentFormat.OpenXml)
        End If


        _auditWb.BeginUpdate() 'A tester si vraiment utile !?
        Return _auditWb
    End Function

    Public Sub PrepareFileWorkspacesAnalysis()

        'GET OR CREATE WORKSPACES FOR AUDIT
        Dim psetWs As Workspace = ProjWs.SmartAddWorkspace(ScanWs.Name + " - " + ProjectReference.Constants.psetTgmWsType,, False, "Tgm", ScanWs)
        psetWs.SmartAddAttribute("GhostData", ScanWs.Id.ToString)

        Dim analysisWs = ProjWs.SmartAddWorkspace(ScanWs.Name + " - " + M4D.Treegram.Core.Extensions.Enums.ExtensionType.Analysis.ToString, M4D.Treegram.Core.Extensions.Enums.ExtensionType.Analysis.ToString, True, "Tgm")
        analysisWs.SmartAddAttribute("AdvancedData", ScanWs.Id.ToString)

        'CLEAN FILE WORKSPACES
        Dim wsToClean As New List(Of Workspace)
        wsToClean.Add(SpaceFunctionalModelWs)
        wsToClean.Add(ShellFunctionalModelWs)
        wsToClean.AddRange(LinkedWsList) 'TgmCreation, ValoAcquereur, etc.
        'wsToClean.Add(myAuditFile.TgmCreationWs)
        wsToClean.Add(PsetTgmWs)
        wsToClean.Add(analysisWs)
        wsToClean.RemoveAll(Function(o) o Is Nothing)
        Dim delInt = 0
        For Each ws In wsToClean
            For l As Integer = 0 To ws.MetaObjects.Count - 1
                ws.MetaObjects(l).ClearMetaObject(delInt, False, True, True)
            Next
        Next

        'CLEAN REFERENCES WS (Datcha Case)
        If ProjRefWs IsNot Nothing Then

            'FILTER
            Dim listToLoad = {ScanWs, ProjRefWs}.ToList
            ProjWs.SmartAddTree("ALL").Filter(False, listToLoad).RunSynchronously()

            Dim projRef As New ProjectReference(ProjWs)
            For Each stoRef In projRef.Buildings.SelectMany(Function(o) o.Storeys).ToList
                'Remove child relations targetting scanWs
                Dim children = stoRef.Metaobject.GetChildren.ToList
                Dim j As Integer
                For j = children.Count - 1 To 0 Step -1
                    If ScanWs.Equals(CType(children(j).Container, Workspace)) Then
                        stoRef.Metaobject.RemoveRelation(children(j))
                    End If
                Next
            Next
        End If

    End Sub

    Public Sub SaveAuditWorkbook()
        Dim projectFile = New FileInfo(Me.ProjWs.Path.LocalPath)

        '---Export
        Dim exportDirPath = Path.Combine(projectFile.Directory.FullName, "Audits", FileTgm.Name, Now.Year.ToString + "_W" + CalendarUtils.GetActualWeekOfyearAsString)
        If Not Directory.Exists(exportDirPath) Then
            Directory.CreateDirectory(exportDirPath)
        End If
        Dim destPath = System.IO.Path.Combine(exportDirPath, "Audit_" + Now.Year.ToString + "_W" + CalendarUtils.GetActualWeekOfyearAsString + "_" + FileTgm.Name + ".xlsx")
        If File.Exists(destPath) Then
            Try
                File.Delete(destPath) 'delete if same week
            Catch ex As Exception
                Me.AuditWorkbook.EndUpdate()
                Throw New Exception("Fermer l'audit XL concerné et relancer. S'il n'est pas ouvert, il faut alors enlever la lecture seule sur l'XL.")
            End Try
        End If

        '---Save as XL file
        Me.AuditWorkbook.SaveDocument(destPath, DocumentFormat.OpenXml)
        Me.AuditWorkbook.EndUpdate()

    End Sub

    Public Function InsertProdActionTemplateInAuditWbook(wbName As String, wsName As String) As Worksheet
        '---Get Audit Workbook
        If Me.AuditWorkbook Is Nothing Then Throw New Exception("Audit wBook must be already loaded")
        '---Get Template Worksheet
        Dim templateWs As Worksheet = GetProdActionTemplateWsheet(wbName, wsName)
        '---Delete Ws if existing
        Dim existWs = GetWorksheet(Me.AuditWorkbook, wsName)
        If existWs IsNot Nothing Then
            Me.AuditWorkbook.Worksheets.Remove(existWs)
        End If
        '---Copy Ws
        Dim insertedWs As Worksheet = Me.AuditWorkbook.Worksheets.Add(wsName)
        insertedWs.CopyFrom(templateWs)

        Return insertedWs
    End Function

    Public Function GetProdActionTemplateWsheet(wbName As String, wsName As String) As Worksheet
        '---Get Template Workbook
        Dim projectFile = New FileInfo(Me.ProjWs.Path.LocalPath)
        Dim exeAssemblyPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
        Dim finalTpltPath As String
        If File.Exists(exeAssemblyPath + "\actions\ressources\ProdAction_Template(BAM)\" + wbName + ".xlsx") Then
            finalTpltPath = exeAssemblyPath + "\actions\ressources\ProdAction_Template(BAM)\" + wbName + ".xlsx" 'Compiled mode
        Else
            finalTpltPath = exeAssemblyPath + "\ressources\ProdAction_Template(BAM)\" + wbName + ".xlsx" 'Debug mode
        End If
        Dim templateWb As New Workbook
        If Not templateWb.LoadDocument(finalTpltPath) Then
            Throw New Exception("This file cannot be opened : " + finalTpltPath)
        End If
        '---Get Template Worksheet
        Dim templateWs = GetWorksheet(templateWb, wsName)
        If templateWs Is Nothing Then Throw New Exception("Template worksheet """ + wsName + """, in workbook """ + wbName + """, missing !")
        Return templateWs
    End Function

    Public Function GetOrInsertSettingsWSheet(wsheetName As String) As Worksheet
        If Me.SettingsWorkbook Is Nothing Then Throw New Exception("Settings file missing")
        Dim myWsheet = GetWorksheet(Me.SettingsWorkbook, wsheetName)

        'Insert worksheet in settings if missing
        If myWsheet Is Nothing Then
            'Get default settings workbook
            Dim finalTpltPath
            Dim exeAssemblyPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
            If File.Exists(exeAssemblyPath + "\actions\ressources\ProdAction_Settings\Settings_" + wsheetName + ".xlsx") Then
                finalTpltPath = exeAssemblyPath + "\actions\ressources\ProdAction_Settings\Settings_" + wsheetName + ".xlsx" 'Compiled mode
            Else
                finalTpltPath = exeAssemblyPath + "\ressources\ProdAction_Settings\Settings_" + wsheetName + ".xlsx" 'Debug mode
            End If
            Dim defaultSettingsWorkbook As New Workbook
            If Not defaultSettingsWorkbook.LoadDocument(finalTpltPath) Then
                Throw New Exception("This file cannot be opened : " + finalTpltPath)
            End If
            'Get default settings worksheet
            Dim toCopyWsheet = GetWorksheet(defaultSettingsWorkbook, wsheetName)
            If toCopyWsheet Is Nothing Then Throw New Exception("Template worksheet """ + wsheetName + """, in workbook ""Settings_" + wsheetName + """, missing !")
            'Insert in project settings wsheet
            myWsheet = SettingsWorkbook.Worksheets.Add(wsheetName)
            myWsheet.CopyFrom(toCopyWsheet)
            'Save modifications in project settings
            Try
                File.Delete(_settingsWbPath) 'delete to replace
            Catch ex As Exception
                SettingsWorkbook.EndUpdate()
                Throw New Exception("Fermer l'XL settings concerné et relancer")
            End Try
            SettingsWorkbook.SaveDocument(_settingsWbPath, DocumentFormat.OpenXml)
            SettingsWorkbook.EndUpdate()

            'Set presentation as active worksheet
            Dim prezWs = GetWorksheet(SettingsWorkbook, "PRESENTATION")
            If prezWs IsNot Nothing Then
                SettingsWorkbook.Worksheets.ActiveWorksheet = prezWs
            End If
        End If

        Return myWsheet
    End Function


    Public Shared Function GetWorksheet(wb As Workbook, name As String) As Worksheet
        If wb IsNot Nothing Then
            For Each ws In wb.Worksheets
                If ws.Name = name Then
                    Return ws
                End If
            Next
        End If
        Return Nothing
    End Function

    Public Shared Function GetSpreadsheetColumnName(columnNumber As Integer) As String
        Dim dividend As Integer = columnNumber + 1
        Dim columnName As String = String.Empty
        Dim modulo As Integer

        While dividend > 0
            modulo = (dividend - 1) Mod 26
            columnName = Convert.ToChar(65 + modulo).ToString() & columnName
            dividend = CInt((dividend - modulo) / 26)
        End While

        Return columnName
    End Function
#End Region

#Region "Languages"
    Private _language As LanguageEnum?
    Public ReadOnly Property Language As LanguageEnum
        Get
            If _language Is Nothing Then
                Try
                    Dim langSt = FileTgm.GetAttribute("Language", True).Value.ToString()
                    _language = [Enum].Parse(GetType(LanguageEnum), langSt)
                Catch ex As Exception
                    _language = LanguageEnum.FR
                End Try
                If _language = LanguageEnum.FR Then
                    System.Threading.Thread.CurrentThread.CurrentUICulture = New Globalization.CultureInfo("fr")
                    System.Threading.Thread.CurrentThread.CurrentCulture = New Globalization.CultureInfo("fr")
                ElseIf _language = LanguageEnum.EN Then
                    System.Threading.Thread.CurrentThread.CurrentUICulture = Nothing
                    System.Threading.Thread.CurrentThread.CurrentCulture = Nothing
                End If
            End If
            Return _language
        End Get
    End Property
#End Region

End Class

Public Enum LanguageEnum
    FR
    EN
End Enum
