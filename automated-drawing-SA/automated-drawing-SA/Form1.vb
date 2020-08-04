Imports Inventor

Public Class Form1
    Dim internalDiameter, externalDiameter, height, fascia As Double ' nominal / input reference
    Dim internalDiameter1, internalDiameter2 As Double
    Dim externalDiameter1, externalDiameter2 As Double
    Dim dimensionA, dimensionB, dimensionC, dimensionD, dimensionE, dimensionG As Double
    Dim fasciaCollaudo1, fasciaCollaudo2 As Double

    Dim rodSplit1, rodSplit2, pisEndless As Boolean
    'Dim splitType As String

    Dim drawingNumber As String

    Dim partDoc As Inventor.PartDocument
    Dim param As Inventor.Parameter
    Dim invApp As Inventor.Application

    '##### Endless/Split Boolean'
    Public Sub radiobutton_rodSplit1_Click(sender As Object, e As EventArgs) Handles radiobutton_rodSplit1.Click
        If radiobutton_rodSplit1.Checked = True Then
            rodSplit1 = True
        End If
    End Sub
    Public Sub radiobutton_rodSplit2_Click(sender As Object, e As EventArgs) Handles radiobutton_rodSplit2.Click
        If radiobutton_rodSplit2.Checked = True Then
            rodSplit2 = True
        End If
    End Sub

    Public Sub radiobutton_pisEndless_Click(sender As Object, e As EventArgs) Handles radiobutton_pisEndless.Click
        If radiobutton_pisEndless.Checked = True Then
            pisEndless = True
        End If
    End Sub

    Public Sub button_ok_Click(sender As Object, e As EventArgs) Handles button_ok.Click
        '##### Get the values from textbox and store as variable (Double)'
        internalDiameter = textbox_internalDiameter.Text
        externalDiameter = textbox_externalDiameter.Text
        height = textbox_height.Text
        drawingNumber = textbox_drawingNumber.Text

        Dim fascia As Double = (externalDiameter - internalDiameter) * 0.5

        ' ##### Verifying the values
        If internalDiameter > externalDiameter Then
            MessageBox.Show("Internal Diameter must be smaller than external diameter!", "Wrong Dimension", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            Me.Close()
        End If

        If internalDiameter + 7 >= externalDiameter Then
            MessageBox.Show("Cross section is too small!", "Wrong Dimension", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            Me.Close()
        End If

        If Height > 50 Then
            MessageBox.Show("Isn't the height too high?" & vbCrLf & "Lower the height and insert a backup ring.", "Wrong Dimension", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Me.Close()
        End If
        If height < fascia + 5 Then
            MessageBox.Show("The height is too low. Make it higher than " & fascia + 5 & " mm.", "Wrong dimension", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            Me.Close()
        End If

        If internalDiameter > 2100 Or externalDiameter > 2100 Then
            MessageBox.Show("Isn't the diameter too big? Contact ufficio tecnico.", "Troppo grande", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Me.Close()
        End If

        '##### Get the Inventor Application object
        Dim invApp As Inventor.Application
        invApp = GetObject(, "Inventor.Application")

        '##### Open the part.'
        invApp.Documents.Open("\\dataserver2019\Tecnici\CARCO\EngineeringTEAM\AUTOMATIC_CREATOR\automated-drawing-SA\self-aligning-ring-1.ipt")

        '##### Get the active document. This assums it's a part document.
        partDoc = invApp.ActiveDocument

        '##### Get the Parameters collection.
        Dim params As Inventor.Parameters
        params = partDoc.ComponentDefinition.Parameters

        'Self-aligning ring 1
        '##### Get the parameter named "dimensionA_parameter"'
        Dim oDimensionAParam As Inventor.Parameter
        oDimensionAParam = params.Item("dimensionA_parameter")

        '##### Get the parameter named "fasciaCollaudo1_parameter"'
        Dim oFasciaCollaudo1Param As Inventor.Parameter
        oFasciaCollaudo1Param = params.Item("fasciaCollaudo1_parameter")

        '##### Get the parameter named "dimensionE_parameter"'
        Dim oDimensionEParam As Inventor.Parameter
        oDimensionEParam = params.Item("dimensionE_parameter")

        '##### Get the parameter named "externalDiameter1_parameter"'
        Dim oExternalDiameter1Param As Inventor.Parameter
        oExternalDiameter1Param = params.Item("externalDiameter1_parameter")

        'Self-aligning ring 2
        ' To be completed...........


        '##### Calculation parameters
        Dim dimensionAPre, fasciaCollaudo1Pre As Double

        dimensionAPre = fascia * 0.03
        fasciaCollaudo1Pre = fascia - dimensionB - dimensionG

        dimensionA = Math.Round([dimensionAPre], 1)
        dimensionB = fascia * 0.044
        dimensionG = fascia * 0.01
        dimensionC = fascia - dimensionG
        dimensionD = fascia * 0.86
        dimensionE = height
        fasciaCollaudo1 = Math.Round([fasciaCollaudo1Pre], 1)
        externalDiameter1 = externalDiameter - (fascia * 0.088)



        '##### Assign extra 1mm on diameter in case of 1-split / 2-splits.'
        If rodSplit1 = True Or rodSplit2 = True Then
            externalDiameter1 = externalDiameter1 + 1
        End If

        ' ##### Tolerance setting
        ' Fascia
        If fasciaCollaudo1 < 15 Then
            Call oFasciaCollaudo1Param.Tolerance.SetToSymmetric("0.1 mm")
        Else
            Call oFasciaCollaudo1Param.Tolerance.SetToSymmetric("0.15 mm")
        End If

        'Height ( = dimension E)
        If dimensionE < 15 Then
            Call oDimensionEParam.Tolerance.SetToSymmetric("0.1 mm")
        Else
            Call oDimensionEParam.Tolerance.SetToSymmetric("0.15 mm")
        End If

        ' TOLERANCE External Diameter
        ' DOUBLE SPLIT
        Dim rodSplit As Boolean
        If rodSplit1 = True Or rodSplit2 Then
            rodSplit = True
        End If

        If rodSplit = True And externalDiameter < 1100 Then
            Call oExternalDiameter1Param.Tolerance.SetToDeviation("1.0 mm", "-0.0 mm")
        End If
        If rodSplit = True And externalDiameter >= 1100 Then
            Dim splitToleranceExternalDiameter As Double = externalDiameter * 0.0001
            Call oExternalDiameter1Param.Tolerance.SetToDeviation(Math.Round([splitToleranceExternalDiameter], 2), "-0,0 mm")  ' standard unit is cm, thus apply extra 0.1
        End If

        'Final tolerance on external diameter (in case of split)
        Dim finalTolerancePositive As Double
        Dim finalToleranceNegative As Double

        If externalDiameter1 < 1100 Then
            finalTolerancePositive = externalDiameter1 * 0.0005
        Else
            finalTolerancePositive = 1
        End If

        finalToleranceNegative = 0

        '##### Change the equation of the parameter.'
        oDimensionAParam.Expression = dimensionA
        oFasciaCollaudo1Param.Expression = fasciaCollaudo1
        oDimensionEParam.Expression = dimensionE
        oExternalDiameter1Param.Expression = externalDiameter1

        '##### ///Controlling iProperties part'
        '##### Get the "Design Tracking Properties" property set.'
        Dim designTrackPropSet As Inventor.PropertySet
        designTrackPropSet = partDoc.PropertySets.Item("Design Tracking Properties")

        '##### Assign "Drawing N°".'
        '##### Get the "Description" property from the property set.'
        Dim descProp As Inventor.Property
        descProp = designTrackPropSet.Item("Description")
        '##### Set the value of the property using the current value of the textbox.'
        descProp.Value = textbox_object.Text

        ' ##### Assign "Material"
        ' ##### Get the "Material" property from the property set.
        Dim materialType As Inventor.Property
        materialType = designTrackPropSet.Item("Material")
        ' ##### Set the value of the property using the value from input form
        materialType.Value = comboBox_materialType.Text

        '##### Assign "Description (Endless/Double splits)".'
        '##### Get the "Project" property from the property set.'
        Dim splitProp As Inventor.Property
        splitProp = designTrackPropSet.Item("Project")
        '##### Set the value of the property using the current value of the textbox.'
        If rodSplit1 = True Then
            splitProp.Value = "Self-aligning female ring-1 (1 split)"
        End If
        If rodSplit2 = True Then
            splitProp.Value = "Self-aligning female ring-1 (2 splits)"
        End If
        If pisEndless = True Then
            splitProp.Value = "Self-aligning female ring-1 (Endless)"
        End If

        '##### Assign "Housing dimension".'
        '##### Get the "Inventor Summary Information" property set.'
        Dim inventorSummaryInfoPropSet As Inventor.PropertySet
        inventorSummaryInfoPropSet = partDoc.PropertySets.Item("Inventor Summary Information")
        '##### Get the "Subject" property from the property set.'
        Dim housingProp As Inventor.Property
        housingProp = inventorSummaryInfoPropSet.Item("Subject")
        '##### Set the value of the property using the current value of the textbox.'
        housingProp.Value = internalDiameter.ToString() & "/" & textbox_externalDiameter.Text.ToString() & " * " & textbox_housingHeight.Text.ToString()

        ' Drawing for a third party
        'If checkBox_thirdParty.Checked = True Then
        'End If


        '##### Update the document.'
        invApp.ActiveDocument.Update()

        ' ##### Add revision N° in each exported file name (Except Rev.0)
        Dim revision As String
        If comboBox_revision.Text = "0" Or comboBox_revision.Text = "" Then
            revision = ""
        Else
            revision = "_rev." & comboBox_revision.Text
        End If

        '##### Save the part-document with the assigned name (drawingNumber).'
        invApp.ActiveDocument.SaveAs("\\dataserver2019\Tecnici\CARCO\DISEGNI\TORNITURA+MODIFICHE\" & drawingNumber & revision & ".ipt", False)

        '##### Replace the reference .ipt file on the drawing.'
        Dim oDoc As Inventor.DrawingDocument
        oDoc = invApp.Documents.Open("\\dataserver2019\Tecnici\CARCO\EngineeringTEAM\AUTOMATIC_CREATOR\automated-drawing-SA\self-aligning-ring-1.idw")
        oDoc.File.ReferencedFileDescriptors(1).ReplaceReference("\\dataserver2019\Tecnici\CARCO\DISEGNI\TORNITURA+MODIFICHE\" & drawingNumber & revision & ".ipt")

        '##### Scale the drawing views according to the external diameter.'
        ' ##### View A'
        Dim oViewA As DrawingView
        oViewA = oDoc.ActiveSheet.DrawingViews.Item(1)
        If textbox_externalDiameter.Text < 100 Then
            oViewA.[Scale] = 0.8
        ElseIf textbox_externalDiameter.Text >= 100 And textbox_externalDiameter.Text < 150 Then
            oViewA.[Scale] = 0.7
        ElseIf textbox_externalDiameter.Text >= 150 And textbox_externalDiameter.Text < 200 Then
            oViewA.[Scale] = 0.65
        ElseIf textbox_externalDiameter.Text >= 200 And textbox_externalDiameter.Text < 250 Then
            oViewA.[Scale] = 0.6
        ElseIf textbox_externalDiameter.Text >= 250 And textbox_externalDiameter.Text < 300 Then
            oViewA.[Scale] = 0.55
        ElseIf textbox_externalDiameter.Text >= 300 And textbox_externalDiameter.Text < 350 Then
            oViewA.[Scale] = 0.45
        ElseIf textbox_externalDiameter.Text >= 350 And textbox_externalDiameter.Text < 400 Then
            oViewA.[Scale] = 0.4
        ElseIf textbox_externalDiameter.Text >= 400 And textbox_externalDiameter.Text < 450 Then
            oViewA.[Scale] = 0.35
        ElseIf textbox_externalDiameter.Text >= 450 And textbox_externalDiameter.Text < 500 Then
            oViewA.[Scale] = 0.3
        ElseIf textbox_externalDiameter.Text >= 500 And textbox_externalDiameter.Text < 550 Then
            oViewA.[Scale] = 0.25
        ElseIf textbox_externalDiameter.Text >= 550 And textbox_externalDiameter.Text < 600 Then
            oViewA.[Scale] = 0.4
        ElseIf textbox_externalDiameter.Text >= 600 And textbox_externalDiameter.Text < 650 Then
            oViewA.[Scale] = 0.35
        ElseIf textbox_externalDiameter.Text >= 650 And textbox_externalDiameter.Text < 700 Then
            oViewA.[Scale] = 0.3
        ElseIf textbox_externalDiameter.Text >= 700 And textbox_externalDiameter.Text < 750 Then
            oViewA.[Scale] = 0.25
        ElseIf textbox_externalDiameter.Text >= 750 And textbox_externalDiameter.Text < 800 Then
            oViewA.[Scale] = 0.2
        ElseIf textbox_externalDiameter.Text >= 800 And textbox_externalDiameter.Text < 850 Then
            oViewA.[Scale] = 0.15
        ElseIf textbox_externalDiameter.Text >= 850 And textbox_externalDiameter.Text < 900 Then
            oViewA.[Scale] = 0.1
        ElseIf textbox_externalDiameter.Text >= 900 And textbox_externalDiameter.Text < 950 Then
            oViewA.[Scale] = 0.45
        ElseIf textbox_externalDiameter.Text >= 950 And textbox_externalDiameter.Text < 1000 Then
            oViewA.[Scale] = 0.4
        ElseIf textbox_externalDiameter.Text >= 1000 And textbox_externalDiameter.Text < 1050 Then
            oViewA.[Scale] = 0.35
        ElseIf textbox_externalDiameter.Text >= 1050 And textbox_externalDiameter.Text < 1100 Then
            oViewA.[Scale] = 0.3
        ElseIf textbox_externalDiameter.Text >= 1100 And textbox_externalDiameter.Text < 1150 Then
            oViewA.[Scale] = 0.25
        ElseIf textbox_externalDiameter.Text >= 1150 And textbox_externalDiameter.Text < 1200 Then
            oViewA.[Scale] = 0.2
        ElseIf textbox_externalDiameter.Text >= 1200 And textbox_externalDiameter.Text < 1250 Then
            oViewA.[Scale] = 0.15
        ElseIf textbox_externalDiameter.Text >= 1250 And textbox_externalDiameter.Text < 1300 Then
            oViewA.[Scale] = 0.1
        Else
            oViewA.[Scale] = 0.05
        End If

        ' ##### Detail view "B".'
        Dim oViewB As DetailDrawingView
        For Each oSheet As Sheet In oDoc.Sheets
            For Each oView As DrawingView In oSheet.DrawingViews
                If oView.ViewType = DrawingViewTypeEnum.kDetailDrawingViewType Then
                    oViewB = oView
                End If
            Next
        Next

        'Set the scale of Detail View B depending on the size
        'Scale the detail drawing view according to the height.
        If textbox_height.Text < 5 Then
            oViewB.[Scale] = 3
        ElseIf textbox_height.Text >= 5 And textbox_height.Text < 20 Then
            oViewB.[Scale] = 2.5
        ElseIf textbox_height.Text >= 20 And textbox_height.Text < 35 Then
            oViewB.[Scale] = 2
        Else
            oViewB.[Scale] = 1.5
        End If

        ' ##### 3D view "View3D".'
        Dim oView3D As DrawingView
        For Each oSheet As Sheet In oDoc.Sheets
            For Each oView As DrawingView In oSheet.DrawingViews
                If oView.ViewType = DrawingViewTypeEnum.kProjectedDrawingViewType Then
                    oView3D = oView
                End If
            Next
        Next

        'Set the scale of 3D view depending on the size
        'Scale the detail drawing view according to the height.
        If textbox_externalDiameter.Text < 100 Then
            oView3D.[Scale] = 0.4
        ElseIf textbox_externalDiameter.Text >= 100 And textbox_externalDiameter.Text < 200 Then
            oView3D.[Scale] = 0.2
        ElseIf textbox_externalDiameter.Text >= 200 And textbox_externalDiameter.Text < 400 Then
            oView3D.[Scale] = 0.15
        ElseIf textbox_externalDiameter.Text >= 400 And textbox_externalDiameter.Text < 700 Then
            oView3D.[Scale] = 0.1
        ElseIf textbox_externalDiameter.Text >= 700 And textbox_externalDiameter.Text < 1000 Then
            oView3D.[Scale] = 0.05
        ElseIf textbox_externalDiameter.Text >= 1000 And textbox_externalDiameter.Text < 1200 Then
            oView3D.[Scale] = 0.04
        Else
            oView3D.[Scale] = 0.03
        End If

        ' ##### Update the revision table
        Dim oRevisionTable As RevisionTable = oDoc.ActiveSheet.RevisionTables.Item(1)
        Dim oRow As RevisionTableRow = oRevisionTable.RevisionTableRows.Item(1)
        Dim oCell1 As RevisionTableCell = oRow.Item(1)
        Dim oCell2 As RevisionTableCell = oRow.Item(2)
        Dim oCell3 As RevisionTableCell = oRow.Item(3)
        Dim oCell4 As RevisionTableCell = oRow.Item(4)
        Dim oCell5 As RevisionTableCell = oRow.Item(5)

        Dim oggi As Date = Date.Today   ' Date

        If comboBox_revision.Text = "0" Or comboBox_revision.Text = "" Then
            oCell1.Text = "0"
        Else
            oCell1.Text = comboBox_revision.Text    ' Revision N°
        End If

        If oCell1.Text = "0" Or oCell1.Text = "" Then                    ' Description (If Revision N° is "0", assign "Drawing Issue"
            oCell2.Text = "Drawing Issue"
        Else
            oCell2.Text = textbox_description.Text
        End If
        oCell3.Text = "Automated"               ' Drawn
        oCell4.Text = textbox_signature.Text    ' Approved
        oCell5.Text = oggi                      ' Date (dd/mm/yyyy)

        '##### Save the drawing-document with the assigned name (drawingNumber).'
        invApp.ActiveDocument.SaveAs("\\dataserver2019\Tecnici\CARCO\DISEGNI\TORNITURA+MODIFICHE\" & drawingNumber & revision & ".idw", False)

        '##### Update the document.'
        invApp.ActiveDocument.Update()

        ' ##### Export to PDF.'
        ' Get the active docuement.
        oDoc = invApp.ActiveDocument

        ' Save a copy as a PDF file.
        Call oDoc.SaveAs("\\dataserver2019\Tecnici\CARCO\DISEGNI\TORNITURA+MODIFICHE\" & drawingNumber & revision & ".pdf", True)

        ' Save a copy as a jpeg file.
        'Call oDoc.SaveAsBitmap("\\dataserver2019\Tecnici\CARCO\DISEGNI\TORNITURA+MODIFICHE\" + drawingNumber + ".jpg", 2303, 3258)
        'Call oDoc.SaveAs("\\dataserver2019\Tecnici\CARCO\DISEGNI\TORNITURA+MODIFICHE\" + drawingNumber + ".jpg", True)

        'SaveAsJPG("C:\Users\minso\Documents\Drawings", 3000)

        'Finishing message
        'IF SPLIT = True, ADD COLLAUDO DIMENSION INFO HERE
        If pisEndless = True Then
            MessageBox.Show("Automated drawing is generated. Please double check!", "Taaaaaac! :D", MessageBoxButtons.OK, MessageBoxIcon.None)
        End If

        If rodSplit = True Then
            MessageBox.Show("Automated drawing is generated. Please double check!" & vbCrLf & vbCrLf & "This ring is split. Final external diameter value is " & externalDiameter1 - 1 & " mm" & vbCrLf & "(Tol. +" & finalTolerancePositive & "/" & finalToleranceNegative & " mm)", "Taaaaaac! :D", MessageBoxButtons.OK, MessageBoxIcon.None)
        End If
        Me.Close()
    End Sub

End Class
