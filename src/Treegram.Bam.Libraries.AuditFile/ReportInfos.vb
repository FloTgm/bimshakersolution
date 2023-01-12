Imports M4D.Treegram.Core.Extensions.Entities
Imports M4D.Treegram.Core.Kernel
Imports DevExpress.Spreadsheet
Imports DevExpress.Export.Xl
Imports Treegram.ConstructionManagement

Public Class ReportInfos
    Private ReadOnly _myAuditFile As FileObject
    Public Sub New(myAudit As FileObject)
        _myAuditFile = myAudit
    End Sub

    Public Sub CompleteCriteria(criteriaName As String, visa As visa, comment As String, Optional excelMoreInfoTab As String = "")
        Dim commentRange, visaRange, nameRange, moreInfoRange As CellRange
        If GetCriteriaRanges(commentRange, visaRange, nameRange, moreInfoRange, criteriaName) Then
            visaRange.Value = CorrectLanguageVisa(visa)
            commentRange.Value = comment
            moreInfoRange.Value = excelMoreInfoTab

            _myAuditFile.SaveAuditWorkbook()
        End If
    End Sub

    Public Sub CompleteCriteria(criteriaNumber As Integer, Optional withRef As Boolean = False)
        Dim commentRange, visaRange, nameRange, moreInfoRange, plusRange As CellRange
        Dim maxNameRange, maxValueRange As CellRange
        Select Case criteriaNumber
#Region "IFC CRITERIAS"
            Case 1 'GeoPositioning
                If GetCriteriaRanges(commentRange, visaRange, nameRange, moreInfoRange, criteriaNumber) Then
                    If Not withRef Then
                        If Xmin <= _myAuditFile.Georeferencing.X And _myAuditFile.Georeferencing.X <= Xmax And
                              Ymin <= _myAuditFile.Georeferencing.Y And _myAuditFile.Georeferencing.Y <= Ymax Then
                            visaRange.Value = CorrectLanguageVisa(visa.OK)
                            'commentRange.Value = Criterias.Cri1_Vali1
                            commentRange.Value = CorrectLanguageDescription(description.Cri1_Vali1)
                        Else
                            visaRange.Value = CorrectLanguageVisa(visa.NOK)
                            commentRange.Value = CorrectLanguageDescription(description.Cri1_Ref1)
                        End If
                    Else
                        If goodGeoRef Then
                            visaRange.Value = CorrectLanguageVisa(visa.OK)
                            commentRange.Value = CorrectLanguageDescription(description.Cri1_Vali2)
                        Else
                            If Double.IsNaN(refGeoRef.X) OrElse Double.IsNaN(refGeoRef.Y) OrElse Double.IsNaN(refGeoRef.Z) Then
                                visaRange.Value = CorrectLanguageVisa(visa.REV)
                                commentRange.Value = CorrectLanguageDescription(description.Cri1_Rev2)
                            Else
                                visaRange.Value = CorrectLanguageVisa(visa.NOK)
                                commentRange.Value = CorrectLanguageDescription(description.Cri1_Ref2)
                            End If
                        End If
                    End If
                    moreInfoRange.Value = "[FICHIER]"
                End If

            Case 2 'Project Information
                If GetCriteriaRanges(commentRange, visaRange, nameRange, moreInfoRange, criteriaNumber) Then
                    If HasProjectInfos Then
                        visaRange.Value = CorrectLanguageVisa(visa.INFO)
                        commentRange.Value = "IfcSite : """ + _myAuditFile.ScanWs.GetMetaObjects(, "IfcSite")(0).Name + """ ; IfcBuilding : """ + _myAuditFile.ScanWs.GetMetaObjects(, "IfcBuilding")(0).Name + """"
                    Else
                        visaRange.Value = CorrectLanguageVisa(visa.REV)
                        commentRange.Value = CorrectLanguageDescription(description.Cri2_Rev)
                    End If
                    moreInfoRange.Value = "[STRUCTURE IFC]"
                End If

            Case 3 'IFC Architecture
                If GetCriteriaRanges(commentRange, visaRange, nameRange, moreInfoRange, criteriaNumber) Then
                    If structurationCorrect Then
                        visaRange.Value = CorrectLanguageVisa(visa.OK)
                        commentRange.Value = "L'architecture IFC est conforme à la norme" 'A traduire
                    Else
                        visaRange.Value = CorrectLanguageVisa(visa.NOK)
                        commentRange.Value = "L'architecture IFC n'est pas conforme à la norme" 'A traduire
                    End If
                    moreInfoRange.Value = "[STRUCTURE IFC]"
                End If
            Case 4 'Ouvertures : liaison avec Objet support
                If GetCriteriaRanges(commentRange, visaRange, nameRange, moreInfoRange, criteriaNumber) Then
                    If reservationNb = OpeningsNb Then
                        visaRange.Value = CorrectLanguageVisa(visa.OK)
                    Else
                        visaRange.Value = CorrectLanguageVisa(visa.REV)
                    End If
                    commentRange.Value = Math.Floor(reservationNb / OpeningsNb * 100).ToString + "% (" + reservationNb.ToString + "/" + OpeningsNb.ToString + ") "
                    moreInfoRange.Value = "[OUVERTURES]"
                End If
            Case 5 'Portes : dimensions acceptables
                If GetCriteriaRanges(commentRange, visaRange, nameRange, moreInfoRange, criteriaNumber) Then
                    If revDoorsNb = 0 Then
                        visaRange.Value = CorrectLanguageVisa(visa.OK)
                    Else
                        visaRange.Value = CorrectLanguageVisa(visa.REV)
                    End If
                    commentRange.Value = Math.Floor((doorsNb - revDoorsNb) / doorsNb * 100).ToString + "% (" + (doorsNb - revDoorsNb).ToString + "/" + doorsNb.ToString + ") "
                    moreInfoRange.Value = "[OUVERTURES]"
                End If
            Case 6 'Portes : dimensions acceptables
                If GetCriteriaRanges(commentRange, visaRange, nameRange, moreInfoRange, criteriaNumber) Then
                    If revWindowsNb = 0 Then
                        visaRange.Value = CorrectLanguageVisa(visa.OK)
                    Else
                        visaRange.Value = CorrectLanguageVisa(visa.REV)
                    End If
                    commentRange.Value = Math.Floor((windowsNb - revWindowsNb) / windowsNb * 100).ToString + "% (" + (windowsNb - revWindowsNb).ToString + "/" + windowsNb.ToString + ") "
                    moreInfoRange.Value = "[OUVERTURES]"
                End If
#End Region

#Region "RVT CRITERIAS"
            Case 21 'Georeferencing
                If GetCriteriaRanges(commentRange, visaRange, nameRange, moreInfoRange, criteriaNumber) Then
                    If withRef Then
                        If goodGeoRef Then
                            visaRange.Value = CorrectLanguageVisa(visa.OK)
                            commentRange.Value = CorrectLanguageDescription(description.Cri21_Vali)
                        Else
                            visaRange.Value = CorrectLanguageVisa(visa.REV)
                            commentRange.Value = CorrectLanguageDescription(description.Cri21_Rev)
                        End If
                    Else
                        'OrgineInterne
                        If geoRefCorrect Then
                            visaRange.Value = "Validé"
                            commentRange.Value = "Cohérent avec le système Lambert93CC"
                        Else
                            visaRange.Value = "A revoir"
                            commentRange.Value = "Non cohérent avec le système Lambert93CC"
                        End If
                    End If
                End If
            Case 22 'Base point
                'ShowRow(criteriaNumber)
                If GetCriteriaRanges(commentRange, visaRange, nameRange, moreInfoRange, criteriaNumber) Then
                    If goodBasePt Is Nothing Then
                        commentRange.Value = CorrectLanguageDescription(description.Cri22_NoRef)
                        visaRange.Value = ""
                        visaRange.FillColor = XlColor.FromTheme(XlThemeColor.Accent4, 0.6) 'Sensé être bleu... (XlThemeColor.xlThemeColorAccent5)
                        Return
                    End If
                    If goodBasePt Then
                        visaRange.Value = CorrectLanguageVisa(visa.OK)
                        commentRange.Value = CorrectLanguageDescription(description.Cri22_Vali)
                    Else
                        visaRange.Value = CorrectLanguageVisa(visa.REV)
                        commentRange.Value = CorrectLanguageDescription(description.Cri22_Rev)
                    End If
                End If

            Case 23 'Survey Point
                'ShowRow(criteriaNumber)
                If GetCriteriaRanges(commentRange, visaRange, nameRange, moreInfoRange, criteriaNumber) Then
                    If goodSurveyPt Is Nothing Then
                        commentRange.Value = CorrectLanguageDescription(description.Cri23_NoRef)
                        visaRange.Value = ""
                        visaRange.FillColor = XlColor.FromTheme(XlThemeColor.Accent4, 0.6) 'Sensé être bleu... (XlThemeColor.xlThemeColorAccent5)
                        Return
                    End If
                    If goodSurveyPt Then
                        visaRange.Value = CorrectLanguageVisa(visa.OK)
                        commentRange.Value = CorrectLanguageDescription(description.Cri23_Vali)
                    Else
                        visaRange.Value = CorrectLanguageVisa(visa.REV)
                        commentRange.Value = CorrectLanguageDescription(description.Cri23_Rev)
                    End If
                End If

            Case 24 'Grids
                'ShowRow(criteriaNumber)
                If GetCriteriaRanges(commentRange, visaRange, nameRange, moreInfoRange, criteriaNumber) Then
                    If withRef Then
                        If goodGrids Is Nothing Then
                            commentRange.Value = CorrectLanguageDescription(description.Cri24_NoRef)
                            visaRange.Value = ""
                            visaRange.FillColor = XlColor.FromTheme(XlThemeColor.Accent4, 0.6) 'Censé être bleu... (XlThemeColor.xlThemeColorAccent5)
                            Return
                        End If
                        If goodGrids Then
                            visaRange.Value = CorrectLanguageVisa(visa.OK)
                            commentRange.Value = CorrectLanguageDescription(description.Cri24_Vali)
                        Else
                            visaRange.Value = CorrectLanguageVisa(visa.REV)
                            commentRange.Value = String.Format(CorrectLanguageDescription(description.Cri24_Rev), gridNOK.ToString)
                        End If
                    Else
                        If GetCriteriaRanges(commentRange, visaRange, nameRange, plusRange, criteriaNumber) Then
                            visaRange.Value = "Pour information"
                            commentRange.Value = "Nombre de quadrillages : " + nbGrids.ToString()
                        End If
                    End If
                End If

            Case 26
                'Absences imports DWG
                If GetCriteriaRanges(commentRange, visaRange, nameRange, plusRange, criteriaNumber) Then
                    If GetMaxRanges(maxNameRange, maxValueRange, "maxImpDwg") Then
                        Dim maxDwg As Integer
                        If maxValueRange.Value Is Nothing OrElse Not maxValueRange.Value.IsNumeric Then
                            visaRange.Value = "Pour information"
                            commentRange.Value = "Nombre d'imports dwg : " + nbImportDwgInfo.ToString()
                        Else
                            maxDwg = maxValueRange.Value.NumericValue
                            If nbImportDwgInfo <= maxDwg Then
                                visaRange.Value = "Validé"
                                commentRange.Value = "Nombre d'imports dwg : " + nbImportDwgInfo.ToString() + " <= " + maxDwg.ToString

                            Else
                                visaRange.Value = "Refusé"
                                commentRange.Value = "Nombre d'imports dwg : " + nbImportDwgInfo.ToString() + "> " + maxDwg.ToString

                            End If
                        End If
                        CompleteStatisticColumns(visaRange, nbImportDwgInfo, maxDwg)
                    End If
                End If
            Case 27
                'Avertissements
                If GetCriteriaRanges(commentRange, visaRange, nameRange, plusRange, criteriaNumber) Then
                    If GetMaxRanges(maxNameRange, maxValueRange, "maxCritère27") Then
                        If maxValueRange.Value.NumericValue = 0 Then
                            visaRange.Value = CorrectLanguageVisa(visa.INFO)
                            commentRange.Value = "Nombre d'avertissements : " + nbWarningsInfo.ToString()
                        Else
                            If nbWarningsInfo <= maxValueRange.Value.NumericValue Then
                                visaRange.Value = CorrectLanguageVisa(visa.OK)
                                commentRange.Value = "Nombre d'avertissements : " + nbWarningsInfo.ToString() + " <= " + maxValueRange.Value.NumericValue.ToString
                            Else
                                visaRange.Value = CorrectLanguageVisa(visa.NOK)
                                commentRange.Value = "Nombre d'avertissements : " + nbWarningsInfo.ToString() + " > " + maxValueRange.Value.NumericValue.ToString
                            End If
                        End If
                        CompleteStatisticColumns(visaRange, nbWarningsInfo, maxValueRange.Value.NumericValue)
                    End If
                End If
            Case 28 'Critical warnings
                'ShowRow(criteriaNumber)
                If GetCriteriaRanges(commentRange, visaRange, nameRange, moreInfoRange, criteriaNumber) Then

                    If nbBadWarnings = 0 Then
                        visaRange.Value = CorrectLanguageVisa(visa.OK)
                        commentRange.Value = CorrectLanguageDescription(description.Cri28_Vali)
                    Else
                        visaRange.Value = CorrectLanguageVisa(visa.REV)
                        commentRange.Value = String.Format(CorrectLanguageDescription(description.Cri28_Rev), nbBadWarnings.ToString)
                    End If
                    moreInfoRange.Value = "[AVERTISSEMENTS]"
                    CompleteStatisticColumns(visaRange, nbBadWarnings)

                End If
            Case 29 'Families Name
                'ShowRow(criteriaNumber)
                If GetCriteriaRanges(commentRange, visaRange, nameRange, moreInfoRange, criteriaNumber) Then

                    If FamilyNamesCriteria_goodName Then
                        visaRange.Value = CorrectLanguageVisa(visa.OK)
                        commentRange.Value = CorrectLanguageDescription(description.Cri29_Vali)
                    Else
                        visaRange.Value = CorrectLanguageVisa(visa.REV)
                        commentRange.Value = String.Format(CorrectLanguageDescription(description.Cri29_Rev), FamilyNamesCriteria_numberOfBadNames.ToString)
                    End If
                End If
            Case 30 'Views Name
                'ShowRow(criteriaNumber)
                If GetCriteriaRanges(commentRange, visaRange, nameRange, moreInfoRange, criteriaNumber) Then

                    If ViewNamesCriteria_goodName Then
                        visaRange.Value = CorrectLanguageVisa(visa.OK)
                        commentRange.Value = CorrectLanguageDescription(description.Cri30_Vali)
                    Else
                        visaRange.Value = CorrectLanguageVisa(visa.REV)
                        commentRange.Value = String.Format(CorrectLanguageDescription(description.Cri30_Rev), ViewNamesCriteria_numberOfBadNames.ToString)
                    End If
                End If
            Case 31 'Sheets Name
                'ShowRow(criteriaNumber)
                If GetCriteriaRanges(commentRange, visaRange, nameRange, moreInfoRange, criteriaNumber) Then

                    If SheetNamesCriteria_goodName Then
                        visaRange.Value = CorrectLanguageVisa(visa.OK)
                        commentRange.Value = CorrectLanguageDescription(description.Cri31_Vali)
                    Else
                        visaRange.Value = CorrectLanguageVisa(visa.REV)
                        commentRange.Value = String.Format(CorrectLanguageDescription(description.Cri31_Rev), SheetNamesCriteria_numberOfBadNames.ToString)
                    End If
                End If
#End Region

#Region "COMMON CRITERIAS"
            Case 51 'Angle True North / Project North
                If withRef Then
                    'ShowRow(criteriaNumber)
                    If GetCriteriaRanges(commentRange, visaRange, nameRange, moreInfoRange, criteriaNumber) Then
                        If refGeoRef.Angle = _myAuditFile.Georeferencing.Angle Then
                            visaRange.Value = CorrectLanguageVisa(visa.OK)
                            commentRange.Value = CorrectLanguageDescription(description.Cri51_Vali)
                        Else
                            If Double.IsNaN(refGeoRef.Angle) Then
                                visaRange.Value = CorrectLanguageVisa(visa.REV)
                                commentRange.Value = CorrectLanguageDescription(description.Cri51_Rev)
                            Else
                                visaRange.Value = CorrectLanguageVisa(visa.NOK)
                                commentRange.Value = CorrectLanguageDescription(description.Cri51_Ref)
                            End If
                        End If
                        moreInfoRange.Value = "[FICHIER]"
                    End If
                End If
            Case 52 'Levels
                'ShowRow(criteriaNumber)
                If GetCriteriaRanges(commentRange, visaRange, nameRange, moreInfoRange, criteriaNumber) Then

                    If Not withRef Then
                        If goodLvls Then
                            visaRange.Value = CorrectLanguageVisa(visa.OK)
                            'commentRange.Value = String.Format(CorrectLanguageDescription(description.Cri52_Vali1), lvlEcartMin.ToString) 'Pas de commentaire d'après la doc de Manuel
                        Else
                            visaRange.Value = CorrectLanguageVisa(visa.REV)
                            commentRange.Value = String.Format(CorrectLanguageDescription(description.Cri52_Rev1), lvlEcartMin.ToString)
                        End If

                    Else
                        If lvlNOK > 0 Then
                            visaRange.Value = CorrectLanguageVisa(visa.NOK)
                            commentRange.Value = String.Format(CorrectLanguageDescription(description.Cri52_Ref2), lvlNOK.ToString)
                        ElseIf lvlREV > 0 Then
                            visaRange.Value = CorrectLanguageVisa(visa.REV)
                            commentRange.Value = String.Format(CorrectLanguageDescription(description.Cri52_Rev2), lvlREV.ToString)
                        Else
                            visaRange.Value = CorrectLanguageVisa(visa.OK)
                            commentRange.Value = CorrectLanguageDescription(description.Cri52_Vali2)
                        End If
                    End If
                    If _myAuditFile.ExtensionType = "rvt" Then
                        moreInfoRange.Value = "[STRUCTURATION]"
                    Else
                        moreInfoRange.Value = "[STRUCTURE IFC]"
                    End If
                End If

            Case 53 'File Size - No need for language distinction for description
                If GetCriteriaRanges(commentRange, visaRange, nameRange, moreInfoRange, criteriaNumber) Then
                    Dim commentValue As String = ""
                    'Onglet à renommer
                    If _myAuditFile.IsRevit Then
                        commentValue = "[METRIQUES]"
                    Else
                        commentValue = "[FICHIER]"
                    End If

                    If GetMaxRanges(maxNameRange, maxValueRange, "maxCritère53") Then
                        Dim FileSizeLimit = maxValueRange.Value.NumericValue
                        If fileSizeInfo < FileSizeLimit Then
                            visaRange.Value = CorrectLanguageVisa(visa.OK)
                            commentRange.Value = fileSizeInfo.ToString + "Mo (< " + FileSizeLimit.ToString + "Mo)"
                        ElseIf FileSizeLimit = 0 Then
                            visaRange.Value = CorrectLanguageVisa(visa.INFO)
                            commentRange.Value = fileSizeInfo.ToString
                        Else
                            visaRange.Value = CorrectLanguageVisa(visa.NOK)
                            commentRange.Value = fileSizeInfo.ToString + "Mo (> " + FileSizeLimit.ToString + "Mo)"
                        End If
                        CompleteStatisticColumns(visaRange, fileSizeInfo, FileSizeLimit)
                    End If

                    moreInfoRange.Value = commentValue
                End If
            Case 54 'File Units
                'Unités
                If GetCriteriaRanges(commentRange, visaRange, nameRange, plusRange, criteriaNumber) Then
                    'If _myAuditFile.IsIfc Then
                    'Onglet à renommer
                    If _myAuditFile.IsRevit Then
                        plusRange.Value = "[METRIQUES]"
                    Else
                        plusRange.Value = "[FICHIER]"
                    End If
                    If validUnits.Count = 0 And invalidUnits.Count = 0 Then
                        visaRange.Value = CorrectLanguageVisa(visa.INFO)
                        If uncomparedUnits.Count = 0 Then
                            commentRange.Value = "Aucune unité renseignée"
                        Else
                            Dim newString As String = uncomparedUnits.First
                            For j = 1 To uncomparedUnits.Count - 1
                                newString += ", " + uncomparedUnits.ElementAt(j)
                            Next
                            commentRange.Value = "Unité(s) renseignée(s) : " + newString
                        End If
                    ElseIf invalidUnits.Count > 0 Then
                        visaRange.Value = CorrectLanguageVisa(visa.NOK)
                        Dim newString As String = invalidUnits.First
                        For j = 1 To invalidUnits.Count - 1
                            newString += ", " + invalidUnits.ElementAt(j)
                        Next
                        commentRange.Value = "Unité(s) non conforme(s) : " + newString

                    Else
                        visaRange.Value = CorrectLanguageVisa(visa.OK)
                        commentRange.Value = "Toutes les unités sont conformes aux unités de référence"
                    End If

                End If

            Case 55 'Duplicates
                If GetCriteriaRanges(commentRange, visaRange, nameRange, moreInfoRange, criteriaNumber) Then
                    If DuplicateCriteria_duplicatesNumber = 0 Then
                        visaRange.Value = CorrectLanguageVisa(visa.OK)
                        commentRange.Value = CorrectLanguageDescription(description.Cri55_Vali)
                    Else
                        visaRange.Value = CorrectLanguageVisa(visa.NOK)
                        commentRange.Value = String.Format(CorrectLanguageDescription(description.Cri55_Rev), DuplicateCriteria_duplicatesNumber.ToString)
                    End If
                    moreInfoRange.Value = "[DOUBLONS]"
                    CompleteStatisticColumns(visaRange, DuplicateCriteria_duplicatesNumber)
                End If

            Case 56 'Objects attached to correct floor Level
                If GetCriteriaRanges(commentRange, visaRange, nameRange, moreInfoRange, criteriaNumber) Then

                    Dim settingsSt = "(deltaInf = " + deltaInf.ToString + "m, deltaSup = " + deltaSup.ToString + "m)"
                    If badStoreyNumber = 0 Then
                        visaRange.Value = CorrectLanguageVisa(visa.OK)
                        commentRange.Value = String.Format(CorrectLanguageDescription(description.Cri56_Vali), settingsSt)
                    Else
                        visaRange.Value = CorrectLanguageVisa(visa.NOK)
                        commentRange.Value = String.Format(CorrectLanguageDescription(description.Cri56_Rev), badStoreyNumber.ToString, settingsSt)
                    End If
                    moreInfoRange.Value = "[ANALYSE ETAGES]"
                    CompleteStatisticColumns(visaRange, badStoreyNumber)

                End If

            Case 57 'Objects with no type - No need for language distinction for description
                If GetCriteriaRanges(commentRange, visaRange, nameRange, moreInfoRange, criteriaNumber) Then
                    If GetMaxRanges(maxNameRange, maxValueRange, "maxTypeObj") Then
                        Dim maxTypeObj As Double = maxValueRange.Value.NumericValue
                        If maxTypeObj = 0 Then
                            visaRange.Value = CorrectLanguageVisa(visa.INFO)
                            commentRange.Value = "Nombre d'objets non typés : " + (percentageGenerics * 100).ToString
                        Else
                            If percentageGenerics <= maxTypeObj / 100 Then
                                CorrectLanguageVisa(visa.OK)
                                commentRange.Value = "Nombre d'objets non typés : " + (percentageGenerics * 100).ToString + "% <= " + (maxTypeObj).ToString + "%"
                            Else
                                visaRange.Value = CorrectLanguageVisa(visa.NOK)
                                commentRange.Value = "Nombre d'objets non typés : " + (percentageGenerics * 100).ToString + "% > " + (maxTypeObj).ToString + "%"
                            End If
                        End If
                        CompleteStatisticColumns(visaRange, percentageGenerics, maxTypeObj / 100)
                    End If
                    If _myAuditFile.IsIfc Then
                        moreInfoRange.Value = "[TYPES IFC]"
                    End If
                End If
            Case 61 'Columns
                If GetCriteriaRanges(commentRange, visaRange, nameRange, moreInfoRange, criteriaNumber) Then
                    If ColumnGeoCriteria_badColumnGeoCount = 0 Then
                        visaRange.Value = CorrectLanguageVisa(visa.OK)
                        commentRange.Value = CorrectLanguageDescription(description.Cri61_Vali)
                    Else
                        visaRange.Value = CorrectLanguageVisa(visa.REV)
                        commentRange.Value = String.Format(CorrectLanguageDescription(description.Cri61_Rev), ColumnGeoCriteria_badColumnGeoCount.ToString)
                    End If
                    moreInfoRange.Value = "[METIER POTEAUX]"
                    CompleteStatisticColumns(visaRange, ColumnGeoCriteria_badColumnGeoCount)

                End If
            Case 63 'File name
                If GetCriteriaRanges(commentRange, visaRange, nameRange, moreInfoRange, criteriaNumber) Then
                    visaRange.Value = CorrectLanguageVisa(visa.INFO)
                    commentRange.Value = _myAuditFile.FileTgm.Name
                End If
            Case 64 'LOI TREES - No need for language distinction for description
                If GetCriteriaRanges(commentRange, visaRange, nameRange, moreInfoRange, criteriaNumber) Then
                    commentRange.RowHeight = 0 'Hide
                    _myAuditFile.ReportWorksheet.Rows(commentRange.BottomRowIndex).Insert()
                    Dim actualCommentRange = _myAuditFile.ReportWorksheet.Cells(commentRange.BottomRowIndex, commentRange.LeftColumnIndex)
                    Dim actualVisaRange = _myAuditFile.ReportWorksheet.Cells(visaRange.BottomRowIndex, visaRange.LeftColumnIndex)
                    Dim actualNameRange = _myAuditFile.ReportWorksheet.Cells(nameRange.BottomRowIndex, nameRange.LeftColumnIndex)

                    If LoiCriteria_NoObj Then
                        actualVisaRange.Value = CorrectLanguageVisa(visa.INFO)
                        actualCommentRange.Value = LoiCriteria_LoiPercentageStr
                    Else
                        If LoiCriteria_LoiPercentage >= LoiCriteria_Seuil Then
                            actualVisaRange.Value = CorrectLanguageVisa(visa.OK)
                            actualCommentRange.Value = LoiCriteria_LoiPercentageStr + "    (>= " + LoiCriteria_Seuil.ToString + " %)"
                        Else
                            actualVisaRange.Value = CorrectLanguageVisa(visa.REV)
                            actualCommentRange.Value = LoiCriteria_LoiPercentageStr + "    (< " + LoiCriteria_Seuil.ToString + " %)"
                        End If
                    End If

                    actualNameRange.Value = LoiCriteria_LoiTreeName
                    'actualCommentRange.Value = LoiCriteria_LoiPercentageStr

                    'mise en forme
                    Dim newRange = actualNameRange.Resize(1, 4)
                    'Dim newRange = _myAuditFile.ReportWorksheet.Range(AuditFile.GetSpreadsheetColumnName(nameRange.LeftColumnIndex) + (nameRange.BottomRowIndex + 1).ToString + ":" + AuditFile.GetSpreadsheetColumnName(moreInfoRange.LeftColumnIndex) + (nameRange.BottomRowIndex + 1).ToString)
                    newRange.Borders.SetOutsideBorders(System.Drawing.Color.Black, BorderLineStyle.Thin)
                    newRange.Borders.InsideVerticalBorders.LineStyle = BorderLineStyle.Thin

                    actualVisaRange.Resize(1, 3).Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center
                End If
            Case 65 'Structural Clash
                If GetCriteriaRanges(commentRange, visaRange, nameRange, moreInfoRange, criteriaNumber) Then
                    If StructuralClashCriteria_clashesNumber = 0 Then
                        visaRange.Value = CorrectLanguageVisa(visa.OK)
                        commentRange.Value = CorrectLanguageDescription(description.Cri65_Vali)
                    Else
                        visaRange.Value = CorrectLanguageVisa(visa.REV)
                        commentRange.Value = String.Format(CorrectLanguageDescription(description.Cri65_Rev), StructuralClashCriteria_clashesNumber.ToString)
                    End If
                    moreInfoRange.Value = "[CLASHS STRUCTURELS]"
                    CompleteStatisticColumns(visaRange, StructuralClashCriteria_clashesNumber)

                End If

            Case 66 'Emplacement fichier
                If GetCriteriaRanges(commentRange, visaRange, nameRange, moreInfoRange, criteriaNumber) Then
                    Dim filePath() As String = _myAuditFile.FileTgm.GetAttribute("Path").Value.split("\")
                    Dim i As Integer
                    Dim folderPath = IO.Path.GetDirectoryName(_myAuditFile.FileTgm.GetAttribute("Path").Value.ToString()) + "\..."
                    'Dim folderPath = filePath(0)
                    'For i = 1 To filePath.Count - 2
                    '    folderPath += "\" + filePath(i)
                    'Next
                    'folderPath += "\..."
                    visaRange.Value = CorrectLanguageVisa(visa.INFO)
                    commentRange.Value = folderPath
                End If
            Case 67
                If GetCriteriaRanges(commentRange, visaRange, nameRange, moreInfoRange, criteriaNumber) Then
                    If appropriateTrade = True Then
                        visaRange.Value = CorrectLanguageVisa(visa.OK)
                        commentRange.Value = CorrectLanguageDescription(description.Cri67_Vali)
                    Else
                        visaRange.Value = CorrectLanguageVisa(visa.REV)
                        commentRange.Value = CorrectLanguageDescription(description.Cri67_Rev)
                    End If
                    moreInfoRange.Value = "[APPARTENANCE METIER]"
                End If
            Case 68 'Murs : présence d'un axe
                If GetCriteriaRanges(commentRange, visaRange, nameRange, moreInfoRange, criteriaNumber) Then
                    If nbWalls = nbWallswithAxis Then
                        visaRange.Value = CorrectLanguageVisa(visa.OK)
                    Else
                        visaRange.Value = CorrectLanguageVisa(visa.REV)
                    End If
                    'Pas de traduction nécessaire
                    commentRange.Value = Math.Floor(nbWallswithAxis / nbWalls * 100).ToString + "% (" + nbWallswithAxis.ToString + "/" + nbWalls.ToString + ")"
                    moreInfoRange.Value = "[MURS]"
                End If
            Case 69 'Murs : épaisseur acceptable
                If GetCriteriaRanges(commentRange, visaRange, nameRange, moreInfoRange, criteriaNumber) Then
                    If nbWallswithAxis = nbWallsWithWidth Then
                        visaRange.Value = CorrectLanguageVisa(visa.OK)
                    Else
                        visaRange.Value = CorrectLanguageVisa(visa.REV)
                    End If
                    commentRange.Value = Math.Floor(nbWallsWithWidth / nbWallswithAxis * 100).ToString + "% (" + nbWallsWithWidth.ToString + "/" + nbWallswithAxis.ToString + ") " + CorrectLanguageDescription(description.Cri69_Vali)
                    moreInfoRange.Value = "[MURS]"
                End If

#End Region
            Case Else
                Throw New Exception("Criteria """ + criteriaNumber.ToString + """ not treated yet")
        End Select
    End Sub

    Public Sub CompleteInfo(infoRange As String)
        Dim commentRange, visaRange, nameRange, plusRange As CellRange
        Dim maxNameRange, maxValueRange As CellRange
        Select Case infoRange

            '>>/// METRIQUES DU FICHIER
            Case "nbClean"
                If GetCriteriaRanges(commentRange, visaRange, nameRange, infoRange) Then
                    If GetMaxRanges(maxNameRange, maxValueRange, "maxNbClean") Then
                        Dim maxClean As Double = maxValueRange.Value.NumericValue

                        If maxValueRange.Value.NumericValue = 0 Then
                            visaRange.Value = "Pour information"
                            commentRange.Value = "Nombre d'éléments à nettoyer : " + elementsToClean.ToString
                        Else
                            If elementsToClean <= maxClean Then
                                visaRange.Value = "Validé"
                                commentRange.Value = "Nombre d'éléments à nettoyer : " + elementsToClean.ToString + " <= " + maxClean.ToString + "."
                            Else
                                visaRange.Value = "Refusé"
                                commentRange.Value = "Nombre d'éléments à nettoyer : " + elementsToClean.ToString + " > " + maxClean.ToString + "."
                            End If
                        End If
                        CompleteStatisticColumns(visaRange, elementsToClean, maxClean)
                    End If
                End If

            Case "definedSites"
                If GetCriteriaRanges(commentRange, visaRange, nameRange, infoRange) Then
                    If GetMaxRanges(maxNameRange, maxValueRange, "maxDefinedSites") Then
                        Dim maxSites = maxValueRange.Value.NumericValue

                        If maxSites = 0 Then
                            visaRange.Value = "Pour information"
                            commentRange.Value = "Nombre de sites : " + nbSitesInfo.ToString
                        Else
                            If nbSitesInfo <= maxSites Then
                                visaRange.Value = "Validé"
                                commentRange.Value = "Nombre de sites : " + nbSitesInfo.ToString + " <= " + maxSites.ToString
                            Else
                                visaRange.Value = "Refusé"
                                commentRange.Value = "Nombre de sites : " + nbSitesInfo.ToString + " <> " + maxSites.ToString
                            End If
                        End If
                        CompleteStatisticColumns(visaRange, nbSitesInfo, maxSites)
                    End If
                End If


            '>>/// DONNEES GENERALES DE PROJET
            Case "projectName"
                If GetCriteriaRanges(commentRange, visaRange, nameRange, infoRange) Then
                    If GetMaxRanges(maxNameRange, maxValueRange, "refProjectName") Then
                        Dim refProject = maxValueRange.Value.TextValue
                        If projectNameBool AndAlso Not String.IsNullOrEmpty(refProject) Then
                            If projectNameInfo = refProject Then
                                visaRange.Value = "Validé"
                                commentRange.Value = "Conforme à la référence : " + Chr(34) + projectNameInfo + Chr(34)
                            Else
                                If projectNameInfo <> refProject Then
                                    visaRange.Value = "A revoir"
                                    commentRange.Value = Chr(34) + projectNameInfo + Chr(34) + " n'est pas le nom du projet de référence " + Chr(34) + refProject + Chr(34) + "."
                                End If
                            End If
                        ElseIf projectNameBool AndAlso String.IsNullOrEmpty(refProject) Then
                            visaRange.Value = "Pour information"
                            commentRange.Value = "Nom du projet : " + Chr(34) + projectNameInfo + Chr(34)

                        ElseIf projectNameBool = False Then
                            visaRange.Value = "Refusé"
                            commentRange.Value = "Le champ du nom du projet est vide."
                        End If
                        CompleteStatisticColumns(visaRange, projectNameInfo, refProject)
                    End If
                End If

            Case "buildingName"
                If GetCriteriaRanges(commentRange, visaRange, nameRange, infoRange) Then
                    If GetMaxRanges(maxNameRange, maxValueRange, "refBuildingName") Then
                        Dim refBuilding = maxValueRange.Value.TextValue
                        If buildingNameBool AndAlso Not String.IsNullOrEmpty(refBuilding) Then
                            If buildingNameInfo = refBuilding Then
                                visaRange.Value = "Validé"
                                commentRange.Value = "Conforme à la référence : " + Chr(34) + buildingNameInfo + Chr(34)
                            Else
                                If buildingNameInfo <> refBuilding Then
                                    visaRange.Value = "A revoir"
                                    commentRange.Value = Chr(34) + buildingNameInfo + Chr(34) + " n'est pas le nom du bâtiment de référence " + Chr(34) + refBuilding + Chr(34) + "."
                                End If
                            End If

                        ElseIf buildingNameBool AndAlso String.IsNullOrEmpty(refBuilding) Then
                            visaRange.Value = "Pour information"
                            commentRange.Value = "Nom du bâtiment : " + Chr(34) + buildingNameInfo + Chr(34)

                        ElseIf buildingNameBool = False Then
                            visaRange.Value = "Refusé"
                            commentRange.Value = "Le champ du nom du bâtiment est vide."
                        End If
                        CompleteStatisticColumns(visaRange, buildingNameInfo, refBuilding)
                    End If
                End If


            Case "clientName"
                If GetCriteriaRanges(commentRange, visaRange, nameRange, infoRange) Then
                    If GetMaxRanges(maxNameRange, maxValueRange, "refClientName") Then
                        Dim refClient = maxValueRange.Value.TextValue
                        If clientNameBool AndAlso Not String.IsNullOrEmpty(refClient) Then
                            If clientNameInfo = refClient Then
                                visaRange.Value = "Validé"
                                commentRange.Value = "Conforme à la référence : " + Chr(34) + clientNameInfo + Chr(34)
                            Else
                                If clientNameInfo <> refClient Then
                                    visaRange.Value = "A revoir"
                                    commentRange.Value = Chr(34) + clientNameInfo + Chr(34) + "n'est pas le nom du client de référence " + Chr(34) + refClient + Chr(34) + "."
                                End If
                            End If

                        ElseIf clientNameBool AndAlso String.IsNullOrEmpty(refClient) Then
                            visaRange.Value = "Pour information"
                            commentRange.Value = "Nom du client : " + Chr(34) + clientNameInfo + Chr(34)

                        ElseIf clientNameBool = False Then
                            visaRange.Value = "Refusé"
                            commentRange.Value = "Le champ du nom du client est vide."
                        End If
                        CompleteStatisticColumns(visaRange, clientNameInfo, refClient)
                    End If
                End If

            Case "nbPhases"
                If GetCriteriaRanges(commentRange, visaRange, nameRange, infoRange) Then
                    If GetMaxRanges(maxNameRange, maxValueRange, "maxNbPhases") Then
                        Dim maxPhases = maxValueRange.Value.NumericValue
                        If maxPhases = 0 Then
                            visaRange.Value = "Pour information"
                            commentRange.Value = "Nombre de phases : " + nbPhasesInfo.ToString
                        Else
                            If nbPhasesInfo <= maxPhases Then
                                visaRange.Value = "Validé"
                                commentRange.Value = "Nombre de phases : " + nbPhasesInfo.ToString + " <= " + maxPhases.ToString
                            Else
                                visaRange.Value = "Refusé"
                                commentRange.Value = "Nombre de phases : " + nbPhasesInfo.ToString + " > " + maxPhases.ToString

                            End If
                        End If
                        CompleteStatisticColumns(visaRange, nbPhasesInfo, maxPhases)
                    End If
                End If

            Case "nbOptions"
                If GetCriteriaRanges(commentRange, visaRange, nameRange, infoRange) Then
                    If GetMaxRanges(maxNameRange, maxValueRange, "maxNbOptions") Then
                        Dim maxOptions = maxValueRange.Value.NumericValue
                        If maxOptions = 0 Then
                            visaRange.Value = "Pour information"
                            commentRange.Value = "Nombre de variantes : " + nbOptionsInfo.ToString
                        Else
                            If nbOptionsInfo <= maxOptions Then
                                visaRange.Value = "Validé"
                                commentRange.Value = "Nombre de variantes : " + nbOptionsInfo.ToString + " <= " + maxOptions.ToString
                            Else
                                visaRange.Value = "Refusé"
                                commentRange.Value = "Nombre de variantes : " + nbOptionsInfo.ToString + " > " + maxOptions.ToString
                            End If
                        End If
                        CompleteStatisticColumns(visaRange, nbOptionsInfo, maxOptions)
                    End If
                End If



            '>>/// STRUCTURATION DU FICHIER
            Case "nbSubProjects"
                If GetCriteriaRanges(commentRange, visaRange, nameRange, infoRange) Then
                    If GetMaxRanges(maxNameRange, maxValueRange, "maxNbSubProjects") Then
                        Dim maxSubProjects = maxValueRange.Value.NumericValue

                        If maxSubProjects = 0 Then
                            visaRange.Value = "Pour information"
                            commentRange.Value = "Nombre de sous-projets : " + nbSubProjectsInfo.ToString
                        Else
                            If nbSubProjectsInfo <= maxSubProjects Then
                                visaRange.Value = "Validé"
                                commentRange.Value = "Nombre de sous-projets : " + nbSubProjectsInfo.ToString + " <= " + maxSubProjects.ToString
                            Else
                                visaRange.Value = "Refusé"
                                commentRange.Value = "Nombre de sous-projets : " + nbSubProjectsInfo.ToString + " > " + maxSubProjects.ToString
                            End If
                        End If
                        CompleteStatisticColumns(visaRange, nbSubProjectsInfo, maxSubProjects)
                    End If
                End If

            Case "nameSubProjects"
                If GetCriteriaRanges(commentRange, visaRange, nameRange, infoRange) Then
                    If GetMaxRanges(maxNameRange, maxValueRange, "refSubProjects") Then
                        Dim subProjectColumn = maxNameRange.LeftColumnIndex
                        Dim subProjectRowMax = _myAuditFile.SeuilWsheet.GetUsedRange().RowCount
                        Dim refSubProjects As New List(Of String)

                        For i = 5 To subProjectRowMax
                            Dim valueRow = _myAuditFile.SeuilWsheet.Cells(i, subProjectColumn).Value.TextValue
                            If Not String.IsNullOrEmpty(valueRow) Then
                                refSubProjects.Add(valueRow)
                            End If
                        Next

                        If refSubProjects.Count > 0 AndAlso nameSubProjectsInfo.Count > 0 Then
                            Dim items1() As String = refSubProjects.ToArray
                            Dim items2() As String = nameSubProjectsInfo.ToArray
                            Dim listRefWithoutActual As New List(Of String)
                            listRefWithoutActual.AddRange(items1.Except(items2).ToArray)

                            If listRefWithoutActual.Count > 0 Then
                                visaRange.Value = CorrectLanguageVisa(visa.NOK)
                                commentRange.Value = "Au moins un sous-projet de référence n'est pas présent dans la maquette."
                            Else
                                visaRange.Value = CorrectLanguageVisa(visa.OK)
                                commentRange.Value = "Tous les sous-projets de référence sont présents dans la maquette."
                            End If
                        ElseIf refSubProjects.Count = 0 Then
                            visaRange.Value = CorrectLanguageVisa(visa.MANUAL)
                        End If
                    End If
                End If

            Case "projectParameters"
                If GetCriteriaRanges(commentRange, visaRange, nameRange, infoRange) Then
                    If GetMaxRanges(maxNameRange, maxValueRange, "maxProjectParameters") Then
                        Dim maxPPara = maxValueRange.Value.NumericValue

                        If maxPPara = 0 Then
                            visaRange.Value = CorrectLanguageVisa(visa.INFO)
                            commentRange.Value = "Nombre de paramètres de projet : " + projectParametersInfo.ToString
                        Else
                            If projectParametersInfo <= maxPPara Then
                                visaRange.Value = CorrectLanguageVisa(visa.OK)
                                commentRange.Value = "Nombre de paramètres de projet : " + projectParametersInfo.ToString + " <= " + maxPPara.ToString
                            Else
                                visaRange.Value = CorrectLanguageVisa(visa.NOK)
                                commentRange.Value = "Nombre de paramètres de projet : " + projectParametersInfo.ToString + " > " + maxPPara.ToString
                            End If
                        End If
                        CompleteStatisticColumns(visaRange, projectParametersInfo, maxPPara)
                    End If
                End If

            Case "sharedParameters"
                If GetCriteriaRanges(commentRange, visaRange, nameRange, infoRange) Then
                    If GetMaxRanges(maxNameRange, maxValueRange, "refSharedParameters") Then
                        Dim sharedPPColumn = maxNameRange.LeftColumnIndex
                        Dim sharedPPRowMax = _myAuditFile.SeuilWsheet.GetUsedRange().RowCount
                        Dim refSharedPP As New List(Of String)

                        For i = 5 To sharedPPRowMax
                            Dim valueRow = _myAuditFile.SeuilWsheet.Cells(i, sharedPPColumn).Value.TextValue
                            If Not String.IsNullOrEmpty(valueRow) Then
                                refSharedPP.Add(valueRow)
                            End If
                        Next

                        If refSharedPP.Count > 0 Then
                            Dim items1() As String = refSharedPP.ToArray
                            Dim items2() As String = nameSharedParametersInfo.ToArray
                            Dim listRefWithoutActual As New List(Of String)
                            listRefWithoutActual.AddRange(items1.Except(items2).ToArray)

                            If listRefWithoutActual.Count > 0 Then
                                visaRange.Value = "Refusé"
                                commentRange.Value = "La liste de paramètres partagés (" + nbSharedParametersInfo.ToString + ") est non conforme."
                            Else
                                visaRange.Value = "Validé"
                                commentRange.Value = "La liste de paramètres partagés (" + nbSharedParametersInfo.ToString + ") est conforme."
                            End If

                        ElseIf refSharedPP.Count = 0 Then
                            If sharedParametersInfo = True Then
                                visaRange.Value = "Pour information"
                                commentRange.Value = "Présence de " + nbSharedParametersInfo.ToString + " paramètres partagés."
                            Else
                                visaRange.Value = "Pour information"
                                commentRange.Value = "Absence de paramètres partagés."
                            End If

                        End If
                    End If
                End If

            Case "startingViewStatus"
                If GetCriteriaRanges(commentRange, visaRange, nameRange, infoRange) Then
                    If GetMaxRanges(maxNameRange, maxValueRange, "refStartingView") Then
                        Dim refStartingView = maxValueRange.Value.TextValue
                        If String.IsNullOrEmpty(refStartingView) Then
                            If existingStartingView Then
                                commentRange.Value = "Présence d'une vue de démarrage."
                            Else
                                commentRange.Value = "Absence d'une vue de démarrage."
                            End If
                            visaRange.Value = "Pour information"
                        Else
                            If refStartingView = "<Derniers ouverts>" Then 'à tester
                                visaRange.Value = "Refusé"
                                commentRange.Value = "Aucune vue de démarrage définie."
                            Else
                                visaRange.Value = "Validé"
                                commentRange.Value = "Une vue de démarrage est définie."
                            End If

                        End If
                    End If
                End If

            Case "startingViewName"
                If GetCriteriaRanges(commentRange, visaRange, nameRange, infoRange) Then
                    If GetMaxRanges(maxNameRange, maxValueRange, "refStartingView") Then
                        Dim refStartingView = maxValueRange.Value.TextValue
                        If String.IsNullOrEmpty(refStartingView) Then
                            If existingStartingView Then
                                commentRange.Value = "Vue de démarrage : " + nameStartingView
                            Else
                                commentRange.Value = "Absence d'une vue de démarrage."
                            End If
                            visaRange.Value = "Pour information"
                        Else
                            If refStartingView = nameStartingView Then
                                visaRange.Value = "Validé"
                                commentRange.Value = nameStartingView + " est le nom de démarrage de référence (" + refStartingView + ")"
                            Else
                                visaRange.Value = "Refusé"
                                commentRange.Value = nameStartingView + " n'est pas le nom de démarrage de référence (" + refStartingView + ")"
                            End If

                        End If
                    End If
                End If

            Case "nbViews"
                If GetCriteriaRanges(commentRange, visaRange, nameRange, infoRange) Then
                    If GetMaxRanges(maxNameRange, maxValueRange, "maxNbViews") Then
                        Dim maxViews = maxValueRange.Value.NumericValue

                        If maxViews = 0 Then
                            visaRange.Value = "Pour information"
                            commentRange.Value = "Nombre de vues : " + nbViewsInfo.ToString
                        Else
                            If nbViewsInfo <= maxViews Then
                                visaRange.Value = "Validé"
                                commentRange.Value = "Nombre de vues : " + nbViewsInfo.ToString + " <= " + maxViews.ToString
                            Else
                                visaRange.Value = "Refusé"
                                commentRange.Value = "Nombre de vues : " + nbViewsInfo.ToString + " > " + maxViews.ToString
                            End If
                        End If
                        CompleteStatisticColumns(visaRange, nbViewsInfo, maxViews)
                    End If
                End If

            Case "nbNomen"
                If GetCriteriaRanges(commentRange, visaRange, nameRange, infoRange) Then
                    If GetMaxRanges(maxNameRange, maxValueRange, "maxNbNomen") Then
                        Dim maxNomen = maxValueRange.Value.NumericValue

                        If maxNomen = 0 Then
                            visaRange.Value = "Pour information"
                            commentRange.Value = "Nombre de nomenclatures : " + nbNomenInfo.ToString
                        Else
                            If nbNomenInfo <= maxNomen Then
                                visaRange.Value = "Validé"
                                commentRange.Value = "Nombre de nomenclatures : " + nbNomenInfo.ToString + " <= " + maxNomen.ToString

                            Else
                                visaRange.Value = "Refusé"
                                commentRange.Value = "Nombre de nomenclatures : " + nbNomenInfo.ToString + " > " + maxNomen.ToString
                            End If
                        End If
                        CompleteStatisticColumns(visaRange, nbNomenInfo, maxNomen)
                    End If
                End If

            Case "nbLegends"
                If GetCriteriaRanges(commentRange, visaRange, nameRange, infoRange) Then
                    If GetMaxRanges(maxNameRange, maxValueRange, "maxNbLegends") Then
                        Dim maxLegends = maxValueRange.Value.NumericValue

                        If maxLegends = 0 Then
                            visaRange.Value = "Pour information"
                            commentRange.Value = "Nombre de légendes : " + nbLegendsInfo.ToString
                        Else
                            If nbLegendsInfo <= maxLegends Then
                                visaRange.Value = "Validé"
                                commentRange.Value = "Nombre de légendes : " + nbLegendsInfo.ToString + " <= " + maxLegends.ToString

                            Else
                                visaRange.Value = "Refusé"
                                commentRange.Value = "Nombre de légendes : " + nbLegendsInfo.ToString + " > " + maxLegends.ToString
                            End If
                        End If
                        CompleteStatisticColumns(visaRange, nbLegendsInfo, maxLegends)
                    End If
                End If

            Case "nbSheets"
                If GetCriteriaRanges(commentRange, visaRange, nameRange, infoRange) Then
                    If GetMaxRanges(maxNameRange, maxValueRange, "maxNbSheets") Then
                        Dim maxSheets = maxValueRange.Value.NumericValue
                        If maxSheets = 0 Then
                            visaRange.Value = "Pour information"
                            commentRange.Value = "Nombre de feuilles : " + nbSheetsInfo.ToString
                        Else
                            If nbSheetsInfo <= maxSheets Then
                                visaRange.Value = "Validé"
                                commentRange.Value = "Nombre de feuilles : " + nbSheetsInfo.ToString + " <= " + maxSheets.ToString

                            Else
                                visaRange.Value = "Refusé"
                                commentRange.Value = "Nombre de feuilles : " + nbSheetsInfo.ToString + " > " + maxSheets.ToString
                            End If
                        End If
                        CompleteStatisticColumns(visaRange, nbSheetsInfo, maxSheets)
                    End If
                End If

            Case "nbLinks"
                If GetCriteriaRanges(commentRange, visaRange, nameRange, infoRange) Then
                    If GetMaxRanges(maxNameRange, maxValueRange, "maxNbLinks") Then
                        Dim maxLinks = maxValueRange.Value.NumericValue

                        If maxLinks = 0 Then
                            visaRange.Value = "Pour information"
                            commentRange.Value = "Nombre de liens : " + nbLinksInfo.ToString
                        Else
                            If nbLinksInfo <= maxLinks Then
                                visaRange.Value = "Validé"
                                commentRange.Value = "Nombre de liens : " + nbLinksInfo.ToString + " <= " + maxLinks.ToString
                            Else
                                visaRange.Value = "Refusé"
                                commentRange.Value = "Nombre de liens : " + nbLinksInfo.ToString + " > " + maxLinks.ToString
                            End If
                        End If
                        CompleteStatisticColumns(visaRange, nbLinksInfo, maxLinks)
                    End If
                End If

            Case "nameOrgaView"
                If GetCriteriaRanges(commentRange, visaRange, nameRange, infoRange) Then
                    commentRange.Value = "Nom de l'arborescence de vues : " + nameOrgaViewInfo
                    visaRange.Value = "Pour information"
                End If

            Case "nameOrgaSheet"
                If GetCriteriaRanges(commentRange, visaRange, nameRange, infoRange) Then
                    commentRange.Value = "Nom de l'arborescence de feuilles : " + nameOrgaSheetInfo
                    visaRange.Value = "Pour information"
                End If

            Case "nbMaterials"
                If GetCriteriaRanges(commentRange, visaRange, nameRange, infoRange) Then
                    If GetMaxRanges(maxNameRange, maxValueRange, "maxNbMaterials") Then
                        Dim maxMaterials = maxValueRange.Value.NumericValue

                        If maxMaterials = 0 Then
                            visaRange.Value = "Pour information"
                            commentRange.Value = "Nombre de matériaux : " + nbMaterialsInfo.ToString
                        Else
                            If nbMaterialsInfo <= maxMaterials Then
                                visaRange.Value = "Validé"
                                commentRange.Value = "Nombre de matériaux : " + nbMaterialsInfo.ToString + " <= " + maxMaterials.ToString
                            Else
                                visaRange.Value = "Refusé"
                                commentRange.Value = "Nombre de matériaux : " + nbMaterialsInfo.ToString + " > " + maxMaterials.ToString
                            End If
                        End If
                        CompleteStatisticColumns(visaRange, nbMaterialsInfo, maxMaterials)
                    End If
                End If



            '>>/// MODELISATIONS

            Case "nbObj2D"
                If GetCriteriaRanges(commentRange, visaRange, nameRange, infoRange) Then
                    If GetMaxRanges(maxNameRange, maxValueRange, "maxNbObj2D") Then
                        Dim maxObj2D = maxValueRange.Value.NumericValue

                        If maxObj2D = 0 Then
                            visaRange.Value = "Pour information"
                            commentRange.Value = "Nombre d'objets 2D : " + nbObjs2DInfo.ToString
                        Else
                            If nbObjs2DInfo <= maxObj2D Then
                                visaRange.Value = "Validé"
                                commentRange.Value = "Nombre d'objets 2D : " + nbObjs2DInfo.ToString + " <= " + maxObj2D.ToString
                            Else
                                visaRange.Value = "Refusé"
                                commentRange.Value = "Nombre d'objets 2D : " + nbObjs2DInfo.ToString + " > " + maxObj2D.ToString
                            End If
                        End If
                        CompleteStatisticColumns(visaRange, nbObjs2DInfo, maxObj2D)
                    End If
                End If

            Case "nbGrp3D"
                If GetCriteriaRanges(commentRange, visaRange, nameRange, infoRange) Then
                    If GetMaxRanges(maxNameRange, maxValueRange, "maxGrp3D") Then
                        Dim maxGrp3D = maxValueRange.Value.NumericValue

                        If maxGrp3D = 0 Then
                            visaRange.Value = "Pour information"
                            commentRange.Value = "Nombre de groupes 3D : " + nbGroup3DInfo.ToString
                        Else
                            If nbGroup3DInfo <= maxGrp3D Then
                                visaRange.Value = "Validé"
                                commentRange.Value = "Nombre de groupes 3D : " + nbGroup3DInfo.ToString + " <= " + maxGrp3D.ToString
                            Else
                                visaRange.Value = "Refusé"
                                commentRange.Value = "Nombre de groupes 3D : " + nbGroup3DInfo.ToString + " > " + maxGrp3D.ToString
                            End If
                        End If
                        CompleteStatisticColumns(visaRange, nbGroup3DInfo, maxGrp3D)
                    End If
                End If

            Case "nbGrp2D"
                If GetCriteriaRanges(commentRange, visaRange, nameRange, infoRange) Then
                    If GetMaxRanges(maxNameRange, maxValueRange, "maxGrp2D") Then
                        Dim maxGrp2D = maxValueRange.Value.NumericValue

                        If maxGrp2D = 0 Then
                            visaRange.Value = "Pour information"
                            commentRange.Value = "Nombre de groupes 2D : " + nbGroup2DInfo.ToString
                        Else

                            If nbGroup2DInfo <= maxGrp2D Then
                                visaRange.Value = "Validé"
                                commentRange.Value = "Nombre de groupes 2D : " + nbGroup2DInfo.ToString + " <= " + maxGrp2D.ToString
                            Else
                                visaRange.Value = "Refusé"
                                commentRange.Value = "Nombre de groupes 2D : " + nbGroup2DInfo.ToString + " > " + maxGrp2D.ToString
                            End If
                        End If
                        CompleteStatisticColumns(visaRange, nbGroup2DInfo, maxGrp2D)
                    End If
                End If

            Case "moyGrp3D"
                If GetCriteriaRanges(commentRange, visaRange, nameRange, infoRange) Then
                    If GetMaxRanges(maxNameRange, maxValueRange, "maxMoy3D") Then
                        Dim maxMoy3D = maxValueRange.Value.NumericValue

                        If maxMoy3D = 0 Then
                            visaRange.Value = "Pour information"
                            commentRange.Value = "Moyenne d'instanciation des groupes 3D : " + moy3DInfo.ToString
                        Else
                            If moy3DInfo < maxMoy3D Then
                                visaRange.Value = "A revoir"
                                commentRange.Value = "Moyenne d'instanciation des groupes 3D : " + moy3DInfo.ToString + " < " + maxMoy3D.ToString
                            Else
                                visaRange.Value = "Validé"
                                commentRange.Value = "Moyenne d'instanciation des groupes 3D : " + moy3DInfo.ToString + " => " + maxMoy3D.ToString
                            End If
                        End If
                        CompleteStatisticColumns(visaRange, moy3DInfo, maxMoy3D)
                    End If
                End If

            Case "moyGrp2D"
                If GetCriteriaRanges(commentRange, visaRange, nameRange, infoRange) Then
                    If GetMaxRanges(maxNameRange, maxValueRange, "maxMoy2D") Then
                        Dim maxMoy2D = maxValueRange.Value.NumericValue

                        If maxMoy2D = 0 Then
                            visaRange.Value = "Pour information"
                            commentRange.Value = "Moyenne d'instanciation des groupes 2D : " + moy2DInfo.ToString

                        Else
                            If moy2DInfo < maxMoy2D Then
                                visaRange.Value = "A revoir"
                                commentRange.Value = "Moyenne d'instanciation des groupes 2D : " + moy2DInfo.ToString + " < " + maxMoy2D.ToString
                            Else
                                visaRange.Value = "Validé"
                                commentRange.Value = "Moyenne d'instanciation des groupes 2D : " + moy2DInfo.ToString + " => " + maxMoy2D.ToString
                            End If
                        End If
                        CompleteStatisticColumns(visaRange, moy2DInfo, maxMoy2D)
                    End If
                End If

            Case "nbModelLines"
                If GetCriteriaRanges(commentRange, visaRange, nameRange, infoRange) Then
                    If GetMaxRanges(maxNameRange, maxValueRange, "maxModelLines") Then
                        Dim maxModelLines = maxValueRange.Value.NumericValue

                        If maxModelLines = 0 Then
                            visaRange.Value = "Pour information"
                            commentRange.Value = "Nombre de lignes de modèles : " + nbModelLinesInfo.ToString
                        Else
                            If nbModelLinesInfo <= maxModelLines Then
                                visaRange.Value = "Validé"
                                commentRange.Value = "Nombre de lignes de modèles : " + nbModelLinesInfo.ToString + " <= " + maxModelLines.ToString
                            Else
                                visaRange.Value = "Refusé"
                                commentRange.Value = "Nombre de lignes de modèles : " + nbModelLinesInfo.ToString + " > " + maxModelLines.ToString
                            End If
                        End If
                        CompleteStatisticColumns(visaRange, nbModelLinesInfo, maxModelLines)
                    End If
                End If

            Case "nbRefPlanes"
                If GetCriteriaRanges(commentRange, visaRange, nameRange, infoRange) Then
                    If GetMaxRanges(maxNameRange, maxValueRange, "maxRefPlanes") Then
                        Dim maxRefPlanes = maxValueRange.Value.NumericValue

                        If maxRefPlanes = 0 Then
                            visaRange.Value = "Pour information"
                            commentRange.Value = "Nombre de plans de référence : " + nbRefPlanesInfo.ToString
                        Else
                            If nbRefPlanesInfo <= maxRefPlanes Then
                                visaRange.Value = "Validé"
                                commentRange.Value = "Nombre de plans de référence : " + nbRefPlanesInfo.ToString + " <= " + maxRefPlanes.ToString
                            Else
                                visaRange.Value = "Refusé"
                                commentRange.Value = "Nombre de plans de référence : " + nbRefPlanesInfo.ToString + " > " + maxRefPlanes.ToString
                            End If
                        End If
                        CompleteStatisticColumns(visaRange, nbRefPlanesInfo, maxRefPlanes)
                    End If
                End If

            Case "makeup2D"
                If GetCriteriaRanges(commentRange, visaRange, nameRange, infoRange) Then
                    If GetMaxRanges(maxNameRange, maxValueRange, "maxMakeUp2D") Then
                        Dim maxMakeUp2D = maxValueRange.Value.NumericValue

                        If maxMakeUp2D = 0 Then
                            visaRange.Value = "Pour information"
                            commentRange.Value = "Nombre d'éléments de maquillage 2D : " + makeup2DInfo.ToString
                        Else
                            If makeup2DInfo <= maxMakeUp2D Then
                                visaRange.Value = "Validé"
                                commentRange.Value = "Nombre d'éléments de maquillage 2D : " + makeup2DInfo.ToString + " <= " + maxMakeUp2D.ToString
                            Else
                                visaRange.Value = "Refusé"
                                commentRange.Value = "Nombre d'éléments de maquillage 2D : " + makeup2DInfo.ToString + " > " + maxMakeUp2D.ToString
                            End If
                        End If
                        CompleteStatisticColumns(visaRange, makeup2DInfo, maxMakeUp2D)
                    End If
                End If

            Case "hiddenObj"
                If GetCriteriaRanges(commentRange, visaRange, nameRange, infoRange) Then
                    If GetMaxRanges(maxNameRange, maxValueRange, "maxHiddenObj") Then
                        Dim maxHiddenObj = maxValueRange.Value.NumericValue

                        If maxHiddenObj = 0 Then
                            visaRange.Value = "Pour information"
                            commentRange.Value = "Nombre d'éléments masqués non robustement : " + hiddenObjInfo.ToString
                        Else
                            If hiddenObjInfo <= maxHiddenObj Then
                                visaRange.Value = "Validé"
                                commentRange.Value = "Nombre d'éléments masqués non robustement : " + hiddenObjInfo.ToString + " <= " + maxHiddenObj.ToString
                            Else
                                visaRange.Value = "Refusé"
                                commentRange.Value = "Nombre d'éléments masqués non robustement : " + hiddenObjInfo.ToString + " > " + maxHiddenObj.ToString
                            End If
                        End If
                        CompleteStatisticColumns(visaRange, hiddenObjInfo, maxHiddenObj)
                    End If
                End If
        End Select
    End Sub



    Private Function GetCriteriaRanges(ByRef commentRange As CellRange, ByRef visaRange As CellRange, ByRef nameRange As CellRange, ByRef moreInfoRange As CellRange, criteriaNumber As Integer) As Boolean
        Return GetCriteriaRanges(commentRange, visaRange, nameRange, moreInfoRange, "critère" + criteriaNumber.ToString)
    End Function
    Private Function GetCriteriaRanges(ByRef commentRange As CellRange, ByRef visaRange As CellRange, ByRef nameRange As CellRange, ByRef moreInfoRange As CellRange, criteriaString As String) As Boolean
        Try
            nameRange = _myAuditFile.ReportWorksheet.Range(criteriaString)
        Catch ex As Exception
            Return False
        End Try
        If nameRange Is Nothing Then Return False
        visaRange = _myAuditFile.ReportWorksheet.Cells(nameRange.BottomRowIndex, nameRange.LeftColumnIndex + 1)
        commentRange = _myAuditFile.ReportWorksheet.Cells(nameRange.BottomRowIndex, nameRange.LeftColumnIndex + 2)
        moreInfoRange = _myAuditFile.ReportWorksheet.Cells(nameRange.BottomRowIndex, nameRange.LeftColumnIndex + 3)
        Return True
    End Function
    Public Function GetCriteriaRanges(ByRef commentRange As CellRange, ByRef visaRange As CellRange, ByRef nameRange As CellRange, criteriaString As String) As Boolean
        Try
            nameRange = _myAuditFile.ReportWorksheet.Range(criteriaString)
        Catch ex As Exception
            Return False
        End Try
        If nameRange Is Nothing Then Return False
        visaRange = _myAuditFile.ReportWorksheet.Cells(nameRange.TopRowIndex, nameRange.LeftColumnIndex + 1)
        commentRange = _myAuditFile.ReportWorksheet.Cells(nameRange.TopRowIndex, nameRange.LeftColumnIndex + 2)
        Return True
    End Function
    Public Function GetMaxRanges(ByRef maxNameRange As CellRange, ByRef maxValueRange As CellRange, maxString As String) As Boolean
        Try
            maxNameRange = _myAuditFile.SeuilWsheet.Range(maxString)
        Catch ex As Exception
            Return False
        End Try
        If maxNameRange Is Nothing Then
            Return False
        Else
            maxValueRange = _myAuditFile.SeuilWsheet.Cells(maxNameRange.TopRowIndex, maxNameRange.LeftColumnIndex + 1)
            Return True
        End If
    End Function
    Private Sub CompleteStatisticColumns(visaRange As CellRange, critValue As Object, Optional critReference As Object = Nothing)
        If critValue Is Nothing Then
            _myAuditFile.ReportWorksheet.Cells(visaRange.BottomRowIndex, visaRange.LeftColumnIndex + 5).Value = "Null"
        ElseIf critValue.GetType = GetType(Integer) Then
            _myAuditFile.ReportWorksheet.Cells(visaRange.BottomRowIndex, visaRange.LeftColumnIndex + 5).Value = CType(critValue, Integer)
        ElseIf critValue.GetType = GetType(Double) Then
            _myAuditFile.ReportWorksheet.Cells(visaRange.BottomRowIndex, visaRange.LeftColumnIndex + 5).Value = CType(critValue, Double)
        Else
            _myAuditFile.ReportWorksheet.Cells(visaRange.BottomRowIndex, visaRange.LeftColumnIndex + 5).Value = CType(critValue, String)
        End If
        If critReference Is Nothing Then
            _myAuditFile.ReportWorksheet.Cells(visaRange.BottomRowIndex, visaRange.LeftColumnIndex + 6).Value = "Null"
        ElseIf critReference.GetType = GetType(Integer) Then
            _myAuditFile.ReportWorksheet.Cells(visaRange.BottomRowIndex, visaRange.LeftColumnIndex + 6).Value = CType(critReference, Integer)
        ElseIf critReference.GetType = GetType(Double) Then
            _myAuditFile.ReportWorksheet.Cells(visaRange.BottomRowIndex, visaRange.LeftColumnIndex + 6).Value = CType(critReference, Double)
        Else
            _myAuditFile.ReportWorksheet.Cells(visaRange.BottomRowIndex, visaRange.LeftColumnIndex + 6).Value = CType(critReference, String)
        End If
    End Sub

    <Obsolete("Plus pertinent avec la customsation des rapports")>
    Private Sub ShowRow(criteriaNumber As Integer)
        Try
            Dim nameRange = _myAuditFile.ReportWorksheet.Range("critère" + criteriaNumber.ToString)
            If nameRange IsNot Nothing Then
                _myAuditFile.ReportWorksheet.Rows(nameRange.BottomRowIndex).Visible = True
            End If
        Catch ex As Exception
        End Try
    End Sub

    Public Sub HideReportColumnsAndRows(reportWsheet As Worksheet)
        Dim i As Integer
        Dim lastRow As Integer = reportWsheet.GetUsedRange.BottomRowIndex

        'Hide rvt rows
        Dim fileTypeToHide As String
        If _myAuditFile.IsRevit Then
            fileTypeToHide = "Ifc"
        Else
            fileTypeToHide = "Rvt"
        End If
        For i = 1 To lastRow
            Dim myCell = reportWsheet.Range("B" + i.ToString)
            If myCell.Value.TextValue IsNot Nothing AndAlso myCell.Value.TextValue = fileTypeToHide Then
                myCell.RowHeight = 0.0
            End If
        Next
        ''Hide reference rows <---- A AMELIORER : ne pas cacher les lignes si c'est pour les réafficher ensuite
        'For i = 1 To lastRow
        '    Dim myCell = reportWsheet.Range("D" + i.ToString)
        '    If myCell.Value.TextValue IsNot Nothing AndAlso myCell.Value.TextValue.ToLower = "Yes" Then
        '        myCell.RowHeight = 0.0
        '    End If
        'Next
        'Hide columns
        reportWsheet.Range("B1").ColumnWidth = 0.0
        reportWsheet.Range("C1").ColumnWidth = 0.0
        reportWsheet.Range("D1").ColumnWidth = 0.0
        reportWsheet.Range("E1").ColumnWidth = 0.0
        reportWsheet.Range("K1").ColumnWidth = 0.0
        reportWsheet.Range("L1").ColumnWidth = 0.0
        reportWsheet.Range("M1").ColumnWidth = 0.0
    End Sub

    Public Sub FillActivity()
        Dim commentRange, visaRange, nameRange As CellRange
        Dim maxNameRange, maxValueRange As CellRange

        If GetCriteriaRanges(commentRange, visaRange, nameRange, "ajoutés") Then
            Dim addedRange = _myAuditFile.ReportWorksheet.Range("ajoutés")
            commentRange.Value = _myAuditFile.AddedObjs
            visaRange.Value = CorrectLanguageVisa(visa.INFO)
        End If

        If GetCriteriaRanges(commentRange, visaRange, nameRange, "modifiés") Then
            Dim modifiedRange = _myAuditFile.ReportWorksheet.Range("modifiés")
            commentRange.Value = _myAuditFile.ModifiedObjs
            visaRange.Value = CorrectLanguageVisa(visa.INFO)
        End If

        If GetCriteriaRanges(commentRange, visaRange, nameRange, "supprimés") Then
            Dim deletedRange = _myAuditFile.ReportWorksheet.Range("supprimés")
            commentRange.Value = _myAuditFile.DeletedObjs
            visaRange.Value = CorrectLanguageVisa(visa.INFO)
        End If

        If _myAuditFile.IsRevit Then
            Dim fileref As M4D.Treegram.Core.Entities.MetaObject = _myAuditFile.FileTgm
            Dim myWs As M4D.Treegram.Core.Entities.Workspace = fileref.Container
            Dim _3dObj = fileref.Relations.Select(Function(r) r.Target).FirstOrDefault(Function(t) t.IsActive AndAlso t.Name = "3D")

            Dim nbTemplates = myWs.MetaObjects.Where(Function(m) m.IsActive AndAlso m.IsTemplate).Count
            Dim levelsObj = fileref.Relations.Select(Function(r) r.Target).FirstOrDefault(Function(t) t.IsActive AndAlso t.Name = "Levels")
            If GetCriteriaRanges(commentRange, visaRange, nameRange, "objs3D") Then
                If GetMaxRanges(maxNameRange, maxValueRange, "maxObjs3D") Then
                    Dim maxObj3D = maxValueRange.Value.NumericValue

                    If maxObj3D = 0 Then
                        visaRange.Value = "Pour information"
                        commentRange.Value = "Nombre d'objets 3D : " + CInt(_3dObj.GetAttribute("NewTotal", True)?.Value).ToString
                    Else
                        If CInt(_3dObj.GetAttribute("NewTotal", True)?.Value) <= maxObj3D Then
                            visaRange.Value = "Validé"
                            commentRange.Value = "Nombre d'objets 3D : " + CInt(_3dObj.GetAttribute("NewTotal", True)?.Value).ToString + " <= " + maxObj3D.ToString

                        Else
                            visaRange.Value = "Refusé"
                            commentRange.Value = "Nombre d'objets 3D : " + CInt(_3dObj.GetAttribute("NewTotal", True)?.Value).ToString + " > " + maxObj3D.ToString
                        End If
                    End If
                End If
            End If

            If GetCriteriaRanges(commentRange, visaRange, nameRange, "ratioObj3D") Then
                If GetMaxRanges(maxNameRange, maxValueRange, "maxRatioObj3D") Then
                    Dim maxRatioObj3D = maxValueRange.Value.NumericValue * 100
                    Dim ratio3D = Math.Round(CInt(_3dObj.GetAttribute("NewTotal", True)?.Value) / (CInt(_3dObj.GetAttribute("NewTotal", True)?.Value) + nbObjs2DInfo) * 100, 1)

                    If maxRatioObj3D = 0 Then
                        visaRange.Value = "Pour information"
                        commentRange.Value = "Ratio 3D : " + ratio3D.ToString + "% "
                    Else
                        If ratio3D <= maxRatioObj3D Then
                            visaRange.Value = "Refusé"
                            commentRange.Value = "Ratio 3D : " + ratio3D.ToString + "% <= " + maxRatioObj3D.ToString + "% - [3D/(3D+2D)]"

                        Else
                            visaRange.Value = "Validé"
                            commentRange.Value = "Ratio 3D : " + ratio3D.ToString + "% > " + maxRatioObj3D.ToString + "% - [3D/(3D+2D)]"
                        End If
                    End If


                End If

            End If
            Dim nb3dObjs = _3dObj.Relations.Count
            Dim generics = _3dObj.Relations.Select(Function(rel) rel.Target).Where(Function(obj) ("Modèles génériques").Equals(obj.GetAttribute("RevitCategory")?.Value))
            percentageGenerics = Math.Round(generics.Count / nb3dObjs, 3)

        Else
            If GetCriteriaRanges(commentRange, visaRange, nameRange, "objs3D") Then
                Dim objsRange = _myAuditFile.ReportWorksheet.Range("objs3D")
                commentRange.Value = _myAuditFile._3dObjects.Count
                visaRange.Value = CorrectLanguageVisa(visa.INFO)
            End If
        End If

    End Sub

    Public Sub CompleteReportInfos()

        'Date of visa
        Dim dateRange As CellRange = Nothing
        Try
            dateRange = _myAuditFile.ReportWorksheet.Range("datevisa")
        Catch ex As Exception
            'Throw New Exception("Le titre de l'audit n'est pas défini")
        End Try
        If dateRange IsNot Nothing Then
            Select Case _myAuditFile.Language
                Case LanguageEnum.EN
                    If dateRange.Value.TextValue = "" Then
                        _myAuditFile.ReportWorksheet.Cells(dateRange.BottomRowIndex, dateRange.LeftColumnIndex - 1).Value = "Date of visa :"
                        dateRange.Value = Date.Now.ToString("d")
                    Else 'OLD REPORT CASE
                        dateRange.Value = "Date of visa :    " + Date.Now.ToString("d")
                    End If
                Case LanguageEnum.FR
                    If dateRange.Value.TextValue = "" Then
                        _myAuditFile.ReportWorksheet.Cells(dateRange.BottomRowIndex, dateRange.LeftColumnIndex - 1).Value = "Date de visa :"
                        dateRange.Value = Date.Now.ToString("d")
                    Else 'OLD REPORT CASE
                        dateRange.Value = "Date de Visa :    " + Date.Now.ToString("d")
                    End If
            End Select
        End If

        'Template
        Dim templateRef As String = String.Empty
        Dim templateRange As CellRange = Nothing
        Try
            templateRef = _myAuditFile.GetOrInsertSettingsWSheet("PRESENTATION").Range("template").Value.ToString
            templateRange = _myAuditFile.ReportWorksheet.Range("template")
        Catch ex As Exception
            'Throw New Exception("Le titre de l'audit n'est pas défini")
        End Try
        If String.IsNullOrEmpty(templateRef) And templateRange IsNot Nothing Then
            If templateRange.Value.TextValue = "" Then
                templateRange.Value = "Inconnu"
            Else 'OLD REPORT CASE
                templateRange.Value = "Template : Inconnu"
            End If
        ElseIf templateRange IsNot Nothing Then
            If templateRange.Value.TextValue = "" Then
                templateRange.Value = templateRef
            Else 'OLD REPORT CASE
                templateRange.Value = "Template : " + templateRef
            End If
        End If

        'Titre
        Dim titreRange As CellRange = Nothing
        Try
            titreRange = _myAuditFile.ReportWorksheet.Range("titre")
        Catch ex As Exception
            'Throw New Exception("Le titre de l'audit n'est pas défini")
        End Try
        If titreRange IsNot Nothing Then
            If _myAuditFile.IsRevit Then
                titreRange.Value = "BIM Shaker (REVIT)"
            Else
                titreRange.Value = "BIM Shaker (IFC)"
            End If
        End If
    End Sub


    Private Function CorrectLanguageVisa(_visa As visa) As String
        Select Case _myAuditFile.Language
            Case LanguageEnum.FR
                Select Case _visa
                    Case visa.OK
                        Return "Validé"
                    Case visa.NOK
                        Return "Refusé"
                    Case visa.REV
                        Return "A revoir"
                    Case visa.INFO
                        Return "Pour information"
                    Case visa.TODO
                        Return "Analyse non lancée"
                    Case visa.MANUAL
                        Return "Analyse manuelle"
                End Select
            Case LanguageEnum.EN
                Select Case _visa
                    Case visa.OK
                        Return "Accepted"
                    Case visa.NOK
                        Return "Rejected"
                    Case visa.REV
                        Return "To be revised"
                    Case visa.INFO
                        Return "For information"
                    Case visa.TODO
                        Return "Analysis not launched"
                    Case visa.MANUAL
                        Return "Manual analysis"
                End Select
        End Select
    End Function

    Public Enum visa
        OK
        NOK
        REV
        INFO
        TODO
        MANUAL
    End Enum

    Private Function CorrectLanguageDescription(st As description) As String
        If Not dico.ContainsKey(st) Then
            Throw New ArgumentException(st + " is not part of the dictionary")
        End If
        Select Case _myAuditFile.Language
            Case LanguageEnum.FR
                Return dico(st).FR
            Case LanguageEnum.EN
                Return dico(st).EN
        End Select
    End Function

    Public Enum description
        Cri1_Vali1
        Cri1_Ref1
        Cri1_Vali2
        Cri1_Rev2
        Cri1_Ref2
        Cri2_Rev
        Cri21_Vali
        Cri21_Rev
        Cri22_NoRef
        Cri22_Vali
        Cri22_Rev
        Cri23_NoRef
        Cri23_Vali
        Cri23_Rev
        Cri24_NoRef
        Cri24_Vali
        Cri24_Rev
        Cri25_Vali
        Cri25_Rev
        Cri28_Vali
        Cri28_Rev
        Cri29_Vali
        Cri29_Rev
        Cri30_Vali
        Cri30_Rev
        Cri31_Vali
        Cri31_Rev
        Cri51_Vali
        Cri51_Rev
        Cri51_Ref
        Cri52_Vali1
        Cri52_Vali2
        Cri52_Rev1
        Cri52_Rev2
        Cri52_Ref2
        Cri54_NoInfo
        Cri54_Vali
        Cri54_Ref
        Cri55_Vali
        Cri55_Rev
        Cri56_Vali
        Cri56_Rev
        Cri61_Vali
        Cri61_Rev
        Cri65_Vali
        Cri65_Rev
        Cri67_Vali
        Cri67_Rev
        Cri68_Vali
        Cri68_Rev
        Cri69_Vali
        Cri69_Rev
        Cri70_Vali
        Cri70_Rev
        Cri71_Vali
        Cri71_Rev
        Cri72_Vali
        Cri72_Rev
    End Enum

    Public dico As New Dictionary(Of description, (FR As String, EN As String)) From {{description.Cri1_Vali1, ("Cohérent au système Lambert93CC", "Consistent with Lambert93CC system")}, 'CRITERIA 1
                                                                                      {description.Cri1_Ref1, ("Non cohérent au système Lambert93CC", "Inconsistent with Lambert93CC system")},
                                                                                      {description.Cri1_Vali2, ("GéoPositionnement conforme à la Référence", "Geopositioning compatible with Reference")},
                                                                                      {description.Cri1_Rev2, ("Multi références de Géopositionnement", "Multiple geopositioning references")},
                                                                                      {description.Cri1_Ref2, ("GéoPositionnement non conforme à la Référence", "Geopositioning inconsistent with Reference")},
                                                                                      {description.Cri2_Rev, ("IfcSite et/ou IfcBuilding inexistant(s) ou sans nom(s)", "IfcSite and/or IfcBuilding do not exist or have no name")}, 'CRITERIA 2
                                                                                      {description.Cri21_Vali, ("GéoPositionnement conforme à la Référence", "Geopositioning compatible with Reference")}, 'CRITERIA 21
                                                                                      {description.Cri21_Rev, ("GéoPositionnement non conforme à la Référence", "Geopositioning inconsistent with Reference")},
                                                                                      {description.Cri22_NoRef, ("Pas de référence pour comparer", "No reference for comparison")}, 'CRITERIA 22
                                                                                      {description.Cri22_Vali, ("Point de Base conforme à la Référence", "Project Base Point compatible with Reference")},
                                                                                      {description.Cri22_Rev, ("Point de Base non conforme à la Référence", "Project Base Point inconsistent with Reference")},
                                                                                      {description.Cri23_NoRef, ("Pas de référence pour comparer", "No reference for comparison")}, 'CRITERIA 23
                                                                                      {description.Cri23_Vali, ("Point de Topographie conforme à la Référence", "Survey Point compatible with Reference")},
                                                                                      {description.Cri23_Rev, ("Point de Topographie non conforme à la Référence", "Survey Point inconsistent with Reference")},
                                                                                      {description.Cri24_NoRef, ("Pas de référence pour comparer", "No reference for comparison")}, 'CRITERIA 24
                                                                                      {description.Cri24_Vali, ("Quadrillages conformes à la Référence", "Grids compatible with Reference")},
                                                                                      {description.Cri24_Rev, ("{0} quadrillages non conformes", "{0} inconsistent grids")},
                                                                                      {description.Cri25_Vali, ("Les familles sont bien nommées", "All families have been correctly named")}, 'CRITERIA 25
                                                                                      {description.Cri25_Rev, ("{0} familles sont mal nommées", "{0} families have not been named correctly")},
                                                                                      {description.Cri28_Vali, ("Pas d'avertissement préoccupant", "No critical warnings")}, 'CRITERIA 28
                                                                                      {description.Cri28_Rev, ("{0} avertissements préoccupants", "{0} critical warnings")},
                                                                                      {description.Cri29_Vali, ("Elements 3D bien nommés", "3D Objects names compliant with codification")}, 'CRITERIA 29
                                                                                      {description.Cri29_Rev, ("{0} Elements 3D à renommer", "{0} 3D Elements need renaming")},
                                                                                      {description.Cri30_Vali, ("Vues bien nommées", "Views names compliant with codification")}, 'CRITERIA 30
                                                                                      {description.Cri30_Rev, ("{0} Vues à renommer", "{0} views need renaming")},
                                                                                      {description.Cri31_Vali, ("Feuilles bien nommées", "Sheets names compliant with codification")}, 'CRITERIA 31
                                                                                      {description.Cri31_Rev, ("{0} Feuilles à renommer", "{0} sheets need renaming")},
                                                                                      {description.Cri51_Vali, ("Rotation conforme à la Référence", "Angle compatible with Reference")}, 'CRITERIA 51
                                                                                      {description.Cri51_Rev, ("Multi références de Rotation", "Multiple angle references")},
                                                                                      {description.Cri51_Ref, ("Rotation non conforme à la Référence", "Angle inconsistent with Reference")},
                                                                                      {description.Cri52_Vali1, ("Tous les niveaux sont correctement espacés (>{0}m)", "The distance between floor levels is sufficient (>{0}m)")}, 'CRITERIA 52
                                                                                      {description.Cri52_Vali2, ("Tous les niveaux de référence sont présents", "Floor Levels consistent with Reference")},
                                                                                      {description.Cri52_Rev1, ("Les espaces entre niveaux sont parfois inférieurs à {0}m", "The distance between some floor levels is too small (<{0}m)")},
                                                                                      {description.Cri52_Rev2, ("Tous les niveaux de référence sont présents. {0} niveau(x) supplémentaire(s) à vérifier", "Check {0} level(s) that aren't referenced")},
                                                                                      {description.Cri52_Ref2, ("Il manque {0} niveau(x) de référence", "There are {0} referenced levels missing")},
                                                                                      {description.Cri54_NoInfo, ("{0} (Info manquante dans les settings)", "{0} (Information missing in Settings)")}, 'CRITERIA 54
                                                                                      {description.Cri54_Vali, ("L'unité de longueur (""{0}"") est conforme aux unités de référence", "à traduire...")},
                                                                                      {description.Cri54_Ref, ("L'unité de longueur (""{0}"") n'est pas conforme aux unités de référence (""{1}"")", "à traduire...")},
                                                                                      {description.Cri55_Vali, ("aucun doublon", "no duplicates")}, 'CRITERIA 55
                                                                                      {description.Cri55_Rev, ("{0} doublon(s)", "{0} duplicates")},
                                                                                      {description.Cri56_Vali, ("Tous les objets sont au bon étage {0}", "All 3D objects are attached to the correct floor level {0}")}, 'CRITERIA 56
                                                                                      {description.Cri56_Rev, ("{0} objet(s) sont au mauvais étage {1}", "{0} object(s) are attached to the wrong level {1}")},
                                                                                      {description.Cri61_Vali, ("Tous les poteaux ont une géométrie cohérente", "All columns' geometry is coherent")}, 'CRITERIA 61
                                                                                      {description.Cri61_Rev, ("{0} poteau(x) a/ont une géométrie incohérente", "The geometry of {0} column(s) is incoherent")},
                                                                                      {description.Cri65_Vali, ("aucun clash", "no clashes")}, 'CRITERIA 65
                                                                                      {description.Cri65_Rev, ("{0} clash(s)", "{0} clashes")},
                                                                                      {description.Cri67_Vali, ("Catégories des objets 3D  conformes à la discipline de la maquette", "Categories of 3D objects coherent with model trade")}, 'CRITERIA 67
                                                                                      {description.Cri67_Rev, ("Catégories des objets 3D non conformes à la discipline de la maquette", "Categories of 3D objects incoherent with model trade")},
                                                                                      {description.Cri69_Vali, (" des murs avec axe", " of wall with axis")}, 'CRITERIA 69
                                                                                      {description.Cri69_Rev, (" des murs avec axe", " of wall with axis")}}

#Region "VARIABLES"
    Public Property HasProjectInfos As Boolean
    Public Property postalAdressExist As Boolean = True
    Public Property structurationCorrect As Boolean = True
    Public Property settingsTemplate As String
    Public Property DuplicateCriteria_duplicatesNumber As Integer

    'LoiCriteria
    Public Property LoiCriteria_ObjAllFallen As Boolean
    Public Property LoiCriteria_NoObj As Boolean
    Public Property LoiCriteria_LoiPercentageStr As String
    Public Property LoiCriteria_Seuil As Double
    Public Property LoiCriteria_LoiTreeName As String
    Public Property LoiCriteria_LoiPercentage As Double

    'GeoRefCriteria
    Public Property Xmin As Double = 1250000 'D'après la doc de Jean
    Public Property Xmax As Double = 2200000
    Public Property Ymin As Double = 1080000
    Public Property Ymax As Double = 9320000

    Public Property refGeoRef As GeoReferencing
    Public Property goodGeoRef As Boolean?
    Public Property goodBasePt As Boolean?
    Public Property goodSurveyPt As Boolean?

    'GridRefCriteria
    Public Property goodGrids As Boolean?
    Public Property gridNOK As Integer

    'LvlRefCriteria
    Public Property goodLvls As Boolean?
    Public Property lvlNOK As Integer
    Public Property lvlREV As Integer
    Public Property lvlEcartMin As Double = 2.0

    'TradeCriteria
    Public Property appropriateTrade As Boolean?

    'WarningCriteria
    Public Property nbBadWarnings As Integer

    'ObjStoreyCriteria
    Public Property badStoreyNumber As Integer
    Public Property deltaInf As Double
    Public Property deltaSup As Double
    'wallCriteria
    Public Property nbWalls As Integer
    Public Property nbWallswithAxis As Integer
    Public Property nbWallsWithWidth As Double

    'FamilyNames
    Public Property FamilyNamesCriteria_goodName As Boolean
    Public Property FamilyNamesCriteria_numberOfBadNames As Integer
    Public Property ViewNamesCriteria_goodName As Boolean
    Public Property ViewNamesCriteria_numberOfBadNames As Integer
    Public Property SheetNamesCriteria_goodName As Boolean
    Public Property SheetNamesCriteria_numberOfBadNames As Integer

    'ColumnGeoCriteria
    Public Property ColumnGeoCriteria_badColumnGeoCount As Integer
    Public Property ColumnGeoCriteria_columnWidthLimit As Double
    Public Property ColumnGeoCriteria_columnRatioLimit As Double

    'OpeningAnalysis
    Public Property OpeningAnalysis_widthAtt As String
    Public Property OpeningAnalysis_heightAtt As String
    Public Property OpeningAnalysis_glazAreaAtt As String
    Public Property OpeningAnalysis_widthMeasure As Double
    Public Property OpeningAnalysis_heightMeasure As Double
    Public Property OpeningAnalysis_glazAreaMeasure As Double
    Public Property OpeningAnalysis_widthTolerance As Double
    Public Property OpeningAnalysis_heightTolerance As Double
    Public Property OpeningAnalysis_DoorMaxHeight As Double
    Public Property OpeningAnalysis_glazAreaTolerance As Double

    Public Property OpeningAnalysis_DoorMinHeight As Double
    Public Property OpeningAnalysis_DoorMaxWidth As Double
    Public Property OpeningAnalysis_DoorMinWidth As Double
    Public Property OpeningAnalysis_WindowMaxHeight As Double
    Public Property OpeningAnalysis_WindowMinHeight As Double
    Public Property OpeningAnalysis_WindowMaxWidth As Double
    Public Property OpeningAnalysis_WindowMinWidth As Double

    Public Property OpeningsNb As Integer
    Public Property reservationNb As Integer

    Public Property doorsNb As Integer
    Public Property windowsNb As Integer
    Public Property revDoorsNb As Integer
    Public Property revWindowsNb As Integer

    'StructuralClashCriteria
    Public Property StructuralClashCriteria_clashesNumber As Double
    Public Property contactTolerance As Double

    'métriques du fichier
    Public Property filePathInfo As String
    Public Property fileName As String
    Public Property fileSizeInfo As Double

    'units
    Public Property validUnits As List(Of String)
    Public Property invalidUnits As List(Of String)
    Public Property uncomparedUnits As List(Of String)

    <Obsolete("A remplacer par fileSizeInfo")>
    Public Property FileSizeCriteria_fileSize As Double
    <Obsolete("A remplacer par validUnits")>
    Public Property lengthInfo As String
    Public Property angleInfo As String
    Public Property surfInfo As String
    Public Property volInfo As String
    Public Property nbWarningsInfo As Integer
    Public Property elementsToClean As Double
    Public Property nbImportDwgInfo As Integer

    'géoréférencement
    Public Property geoRefCorrect As Boolean = False
    Public Property nbSitesInfo As Integer


    'données générales de projet
    Public Property projectNameBool As Boolean
    Public Property buildingNameBool As Boolean
    Public Property clientNameBool As Boolean
    Public Property projectNameInfo As String
    Public Property buildingNameInfo As String
    Public Property clientNameInfo As String
    Public Property nbPhasesInfo As Integer
    Public Property nbOptionsInfo As Integer

    'Public Property storeyAltiCorrect As Boolean = True
    Public Property nbGrids As Integer

    'structuration du fichier
    Public Property nbSubProjectsInfo As Double
    Public Property nameSubProjectsInfo As List(Of String)

    Public Property projectParametersInfo As Double
    Public Property sharedParametersInfo As Boolean
    Public Property nbSharedParametersInfo As Double
    Public Property nameSharedParametersInfo As List(Of String)

    Public Property existingStartingView As Boolean
    Public Property nameStartingView As String

    Public Property nbViewsInfo As Double
    Public Property nbNomenInfo As Double
    Public Property nbLegendsInfo As Double
    Public Property nbSheetsInfo As Double
    Public Property nbLinksInfo As Double

    Public Property nameOrgaViewInfo As String
    Public Property nameOrgaSheetInfo As String

    Public Property nbMaterialsInfo As Double

    'modélisations
    Public Property nbObjs2DInfo As Double
    Public Property nbGroup3DInfo As Double
    Public Property nbGroup2DInfo As Double
    Public Property moy3DInfo As Double
    Public Property moy2DInfo As Double

    Public Property percentageGenerics As Double

    Public Property nbModelLinesInfo As Integer
    Public Property nbRefPlanesInfo As Integer

    Public Property makeup2DInfo As Integer
    Public Property hiddenObjInfo As Integer
#End Region

End Class

