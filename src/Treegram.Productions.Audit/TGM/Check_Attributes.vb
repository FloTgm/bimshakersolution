Imports Treegram.Bam.Libraries.AEC
Imports M4D.Production.Core
Imports M4D.Treegram.Core.Entities
Imports M4D.Treegram.Core.Kernel
Imports M4D.Treegram.Core.Extensions.Kernel
Imports M4D.Treegram.Core.Extensions.Entities
Imports Treegram.ConstructionManagement
Imports DevExpress.Pdf.Native.BouncyCastle.X509
Imports DevExpress.Xpo.Helpers

Public Class Check_Attributes
    Inherits ProdAction
    Public Sub New()
        Name = "TGM :: Check if Objects Attributes (Stand by...)"
        Description = "Description de l'action"
        PartOfScript = True
    End Sub
    Public Overrides Sub CreateGuiInputTree()
        InputTree = TempWorkspace.AddTree("--- Temporary Input Tree ---")
        InputTree.AddNode("object.type", "File").Description = "Files"
    End Sub

    Public Overrides Function GuiUpdateLauncherTree(inputs As Dictionary(Of String, IEnumerable(Of PersistEntity))) As Tree

        Dim launchTree = TempWorkspace.SmartAddTree("LauncherTree")

        ' Suppression des anciens noeuds
        While launchTree.Nodes.Count > 0
            launchTree.RemoveNode(launchTree.Nodes.First)
        End While
        Dim fileNode = launchTree.SmartAddNode("File", "",,, "Files")

        If inputs.Values.SelectMany(Function(i) i).ToList.Count > 0 Then
            For Each inputObj In inputs.Values.SelectMany(Function(i) i).ToList
                Dim inputmObj As MetaObject = inputObj
                'POUR UN FILTRAGE DE TYPE INFORMATIQUE : WORKSPACE, FILE, METAOBJET
                If inputmObj.GetTgmType <> "File" OrElse inputmObj.GetAttribute("Extension").Value.ToString.ToLower <> "ifc" Then Continue For 'Throw New Exception("You must drag and drop file objects")
                launchTree.SmartAddNode("object.id", inputmObj.Id.ToString,,, inputmObj.Name)
            Next
        End If

        Return launchTree
    End Function

    <ActionMethod()>
    Public Function MyMethod(myFichier As MultipleElements) As ActionResult

        If myFichier.Source.Count = 0 Then
            Return New IgnoredActionResult("No Input File")
        End If

        ' Récupération des inputs
        Dim tgmFiles = GetInputAsMetaObjects(myFichier.Source).ToList


        'AUDIT
        For Each fileTgm In tgmFiles

            'Nb de métaObjects
            Dim myWs As Workspace = fileTgm.Container
            Dim metaOCount = myWs.MetaObjects.Count

            myWs.SmartJoin( M4D.Treegram.Core.Extensions.Enums.ExtensionType.AdditionalAttributes)

            Dim attCount As Integer = 0
            Dim filledAttCount As Integer = 0
            For Each mo In myWs.MetaObjects

                If mo.GetAttribute("object.containedIn")?.Value = "3D" Then

                    'Au niveau metaO
                    Dim moAttCount As Integer = 0
                    Dim moFilledAttCount As Integer = 0
                    For Each att In mo.Attributes
                        'Skip Tgm attributes
                        If att.Name = ProjectReference.Constants.psetTgmWsType Then
                            Continue For
                        End If
                        'Doesn't take in account parent attribute - To discuss...
                        If att.Attributes.Count > 0 Then
                            FindOutChildrenAtt(att, moAttCount, moFilledAttCount)
                        Else
                            moAttCount += 1
                            Try
                                If Not String.IsNullOrEmpty(att?.Value?.ToString) Then
                                    If att.Value.ToString <> "" Then
                                        moFilledAttCount += 1
                                    End If
                                End If
                            Catch ex As Exception
                            End Try
                        End If
                    Next

                    'Au niveau de l'extension AdditionalAttributes
                    Dim addAttExt = mo.Extensions.Where(Function(o) CType(o.Container, Workspace).Name.Contains("AdditionalAttributes")).FirstOrDefault
                    If addAttExt IsNot Nothing Then
                        For Each att In mo.Attributes
                            'Doesn't take in account parent attribute - To discuss...
                            If att.Attributes.Count > 0 Then
                                FindOutChildrenAtt(att, moAttCount, moFilledAttCount)
                            Else
                                moAttCount += 1
                                Try
                                    If Not String.IsNullOrEmpty(att?.Value?.ToString) Then
                                        If att.Value.ToString <> "" Then
                                            moFilledAttCount += 1
                                        End If
                                    End If
                                Catch ex As Exception
                                End Try
                            End If
                        Next
                    End If


                    'Nb total d'attributs au niveau du projet
                    attCount += moAttCount
                    'Nb de filledAttributes
                    filledAttCount += moFilledAttCount

                    Dim anaAtt = AecObject.CompleteTgmPset(mo, "AttributesAnalysis", Nothing)
                    Dim myAtt = anaAtt.SmartAddAttribute("AttributesCount", moAttCount.ToString, False)
                    Dim myAtt1 = anaAtt.SmartAddAttribute("FilledAttributesCount", moFilledAttCount.ToString, False)
                    Dim myAtt2 = anaAtt.SmartAddAttribute("FilledPercent", Math.Round(moFilledAttCount / moAttCount, 2).ToString, False)

                    'Pour la tranche de filtrage (en pourcentage vrai)
                    Dim myAtt3 = anaAtt.SmartAddAttribute("FilledTranche", (Math.Floor(moFilledAttCount / moAttCount * 10) * 10).ToString, False)

                    'typage des attributs pour filtrage
                    'myAtt.Type = AttributeType.Number
                    'myAtt1.Type = AttributeType.Number
                    'myAtt2.Type = AttributeType.Number
                    'myAtt3.Type = AttributeType.Number
                End If
            Next

            'Nb moyen d'attributs
            Dim anaAtt2 = AecObject.CompleteTgmPset(fileTgm, "AttributesAnalysis", Nothing)
            Dim myAtt4 = anaAtt2.SmartAddAttribute("AttributesCount", attCount.ToString, False)
            Dim myAtt5 = anaAtt2.SmartAddAttribute("FilledAttributesCount", filledAttCount.ToString, False)
            Dim myAtt6 = anaAtt2.SmartAddAttribute("FilledPercent", (Math.Round(filledAttCount / attCount, 2) * 100).ToString, False)

            'myAtt4.Type = AttributeType.Number
            'myAtt5.Type = AttributeType.Number
            'myAtt6.Type = AttributeType.Number

            OutputTree = setOutputTree(fileTgm)
        Next

        Return Nothing
    End Function

    Private Sub FindOutChildrenAtt(att As Attribute, ByRef moAttCOunt As Integer, ByRef moFilledAttCOunt As Integer)
        For Each childAtt In att.Attributes
            If childAtt.Attributes.Count > 0 Then 'Doesn't take in account parent attribute - To discuss...
                FindOutChildrenAtt(childAtt, moAttCOunt, moFilledAttCOunt)
            Else
                moAttCOunt += 1
                Try
                    If Not String.IsNullOrEmpty(childAtt?.Value?.ToString) Then
                        If childAtt.Value.ToString <> "" Then
                            moFilledAttCOunt += 1
                        End If
                    End If
                Catch ex As Exception
                End Try
            End If
        Next
    End Sub

    Public Function setOutputTree(fileTgm As MetaObject) As Tree
        Dim scanWs As Workspace = fileTgm.Container
        Dim projWs As Workspace = scanWs.Container
        Dim mytree As Tree = projWs.SmartAddTree("Audit - ATTRIBUTS")
        'mytree.Filter(False, {scanWs}).RunSynchronously()

        Dim startNode = mytree.SmartAddNode("", "",)
        Dim fileNode = startNode.SmartAddNode("File", "",,, "Files")

        'Noeuds fichiers
        Dim filesNode = fileNode.SmartAddNode("File", fileTgm.Name)
        'Dim attNode = filesNode.SmartAddNode("AttributesCount", fileTgm.GetAttribute("AttributesCount").Value, AttributeType.Number)
        'Dim filledAttNode = filesNode.SmartAddNode("FilledAttributesCount", fileTgm.GetAttribute("FilledAttributesCount").Value, AttributeType.Number)
        'Dim FilledPercent = filesNode.SmartAddNode("FilledPercent", fileTgm.GetAttribute("FilledPercent").Value, AttributeType.Number)

        Dim attNode = filesNode.SmartAddNode("AttributesCount", fileTgm.GetAttribute("AttributesCount", True).Value,,, "Nombre d'attributs total : " + fileTgm.GetAttribute("AttributesCount", True).Value.ToString)
        Dim filledAttNode = filesNode.SmartAddNode("FilledAttributesCount", fileTgm.GetAttribute("FilledAttributesCount", True).Value,,, "Nombre d'attributs renseignés : " + fileTgm.GetAttribute("FilledAttributesCount", True).Value.ToString)
        Dim FilledPercent = filesNode.SmartAddNode("FilledPercent", fileTgm.GetAttribute("FilledPercent", True).Value,,, fileTgm.GetAttribute("FilledPercent", True).Value.ToString + "% de remplissage")

        'Noeuds MetaObjects
        Dim moNode As Node = startNode.SmartAddNode("", "",,, "MetaObjects")

        'For i = 0 To 10
        '    'A voir si necessaire
        '    moNode.SmartAddNode("FilledAttributes", i * 10, AttributeType.Number)
        'Next
        For i = 0 To 9
            moNode.SmartAddNode("FilledTranche", i * 10,,, "entre " + (i * 10).ToString + "% et " + ((i + 1) * 10).ToString + "%")
        Next

        projWs.PushAllModifiedEntities()

        Return mytree
    End Function

End Class
