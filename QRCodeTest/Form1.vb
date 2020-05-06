Imports Microsoft.Office.Interop

Public Class Form1
    Dim oApp As Word.Application
    Dim oDoc As Word.Document
    Dim oDataFileName As String
    Dim numLabelPerPage As New ArrayList()
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim oAutoText As Word.AutoTextEntry
        oApp = CreateObject("Word.Application")
        oDoc = oApp.Documents.Add

        oDoc.Fields.Add(Range:=oApp.Selection.Range,
                        Type:=Word.WdFieldType.wdFieldMergeBarcode,
                        Text:="qrcode QR \s 55 \q 3",
                        PreserveFormatting:=False)

        oAutoText = oApp.NormalTemplate.AutoTextEntries.Add("MyLabelLayout", oDoc.Content)
        oDoc.Content.Delete()
        CreateMailMergeDataFile()

        With oDoc.MailMerge
            .MainDocumentType = Word.WdMailMergeMainDocType.wdMailingLabels
            .OpenDataSource(Name:=oDataFileName)

            oApp.MailingLabel.CreateNewDocument(Name:="8160 Address Labels", Address:="",
                                                AutoText:="MyLabelLayout", LaserTray:=Word.WdPaperTray.wdPrinterManualFeed)
            .Destination = Word.WdMailMergeDestination.wdSendToNewDocument

            .Execute()

        End With

        oAutoText.Delete()
        oDoc.Close(False)
        oApp.Visible = True

        My.Computer.FileSystem.DeleteFile(oDataFileName)
        oApp.NormalTemplate.Saved = True

    End Sub

    Public Sub CreateMailMergeDataFile()
        Dim iNumLabels As Integer
        Dim oDataDoc As Word.Document
        Dim qrCode As Guid
        iNumLabels = pagesUpDown.Value * numLabelPerPage(labelComboBox.SelectedIndex)

        oDoc.MailMerge.CreateDataSource(Name:=oDataFileName, HeaderRecord:="qrcode")
        oDataDoc = oApp.Documents.Open(oDataFileName)

        For iCount = 1 To iNumLabels
            With oDataDoc.Tables.Item(1)
                qrCode = Guid.NewGuid()
                If iCount <> iNumLabels Then
                    .Rows.Add()
                End If
                .Cell(iCount + 1, 1).Range.InsertAfter(qrCode.ToString())
            End With
        Next

        oDataDoc.Save()
        oDataDoc.Close(False)
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        oDataFileName = String.Format("{0}\labelDataSource.docx", Environment.CurrentDirectory)
        labelComboBox.Items.Add("8160 Address Labels")
        numLabelPerPage.Add(30)

        labelComboBox.SelectedIndex = 0
        pagesUpDown.Value = 1
    End Sub
End Class
