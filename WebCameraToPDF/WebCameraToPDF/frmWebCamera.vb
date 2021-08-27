#Region "ABOUT"
' / -----------------------------------------------------------------------------
' / Developer : Mr.Surapon Yodsanga (Thongkorn Tubtimkrob)
' / eMail : thongkorn@hotmail.com
' / URL: http://www.g2gnet.com (Khon Kaen - Thailand)
' / Facebook: https://www.facebook.com/g2gnet (Thailand only)
' / Facebook: https://www.facebook.com/commonindy (Worldwide)
' / More Info: http://www.g2gnet.com/webboard
' /
' / Purpose: Touchless SDK for Web camera and save PDF into disk.
' / Microsoft Visual Basic .NET (2010)
' /
' / This is open source code under @CopyLeft by Thongkorn/Common Tubtimkrob.
' / You can modify and/or distribute without to inform the developer.
' /
' / See more detail and download SDK ... http://touchless.codeplex.com/
' / Download Syncfusion Community License Free. ... https://www.syncfusion.com/products/communitylicense
' / --------------------------------------------------------------------------
#End Region

Imports TouchlessLib
Imports System.IO
Imports Syncfusion.Pdf
Imports Syncfusion.Pdf.Graphics
Imports Syncfusion.Pdf.Parsing
Imports System.Globalization

Public Class frmWebCamera
    Dim WebCamMgr As TouchlessLib.TouchlessMgr
    '// Create a folder in VB if it doesn't exist.
    Dim TempDirecory = MyPath(Application.StartupPath) & "TempPDF\"
    Dim FinalDirecory = MyPath(Application.StartupPath) & "PDF\"

    Private Sub frmWebCamera_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            Timer1.Enabled = False
            WebCamMgr.CurrentCamera.Dispose()
            WebCamMgr.Cameras.Item(cmbCamera.SelectedIndex).Dispose()
            WebCamMgr.Dispose()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub frmWebCamera_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        WebCamMgr = New TouchlessLib.TouchlessMgr
        Label2.Text = ""
        For i As Integer = 0 To WebCamMgr.Cameras.Count - 1
            cmbCamera.Items.Add(WebCamMgr.Cameras(i).ToString)
        Next
        If cmbCamera.Items.Count >= 0 Then
            cmbCamera.SelectedIndex = 0
            Timer1.Enabled = True
            btnSave.Enabled = False
            If (Not System.IO.Directory.Exists(TempDirecory)) Then System.IO.Directory.CreateDirectory(TempDirecory)
            '// Initialized and load images into DataGridView.
            Call InitDataGridView()
            Call LoadPDF2DatagridView()
        Else
            MessageBox.Show("No Web Camera, This application needs a webcam.", "Report Status", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.Close()
        End If
    End Sub

    ' / --------------------------------------------------------------------------
    Private Sub LoadPDF2DatagridView()
        Dim directory As New IO.DirectoryInfo(TempDirecory)
        If directory.Exists Then
            Dim pdfFiles() As IO.FileInfo = directory.GetFiles("*.pdf")
            Dim row As String()
            Dim iArr() As String
            For Each pdfFile As IO.FileInfo In pdfFiles
                iArr = Split(pdfFile.FullName, "\")
                row = New String() {iArr(UBound(iArr))}
                dgvData.Rows.Add(row)
            Next
        End If
        Label2.Text = "Total PDF : " & Me.dgvData.Rows.Count
    End Sub

    ' / --------------------------------------------------------------------------
    Private Sub btnSave_Click(sender As System.Object, e As System.EventArgs) Handles btnSave.Click
        If picPreview Is Nothing Then Exit Sub
        Try
            '/ Create file name
            Dim sPicName As String = TempDirecory & Format(Now, "ddMMyy-hhmmss") & ".png"
            Dim b As Bitmap = picPreview.Image
            b.Save(sPicName, System.Drawing.Imaging.ImageFormat.Png)
            '// Copy filename and change to PDF extension.
            Dim sPDF As String = Mid(sPicName, 1, Len(sPicName) - 3) & "pdf"
            '/ Create a new PDF document.
            Dim doc As New PdfDocument()
            '/ Add a page to the document
            Dim page As PdfPage = doc.Pages.Add()
            '/ Getting page size to draw the image which fits the page.
            Dim pageSize As SizeF = page.GetClientSize()
            '/ Create PDF graphics for the page.
            Dim graphics As PdfGraphics = page.Graphics
            '/ Load the image from the disk.
            Dim image As New PdfBitmap(sPicName)
            '/ Draw the image
            graphics.DrawImage(image, New RectangleF(0, 0, pageSize.Width, pageSize.Height))
            '/ Save the document
            doc.Save(sPDF)
            '/ Close the document
            doc.Close(True)
            '// Show File name in the DataGridView.
            Dim iArr() As String
            iArr = Split(sPDF, "\")
            Dim row As String()
            row = New String() {iArr(UBound(iArr))}
            dgvData.Rows.Add(row)
            '//
            Label2.Text = "Total PDF : " & Me.dgvData.Rows.Count
            Me.dgvData.Focus()
            SendKeys.Send("^{HOME}")
            SendKeys.Send("^{DOWN}")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    ' / --------------------------------------------------------------------------
    Private Sub dgvData_DoubleClick(sender As Object, e As System.EventArgs) Handles dgvData.DoubleClick
        If Me.dgvData.Rows.Count = 0 Then Return
        FilePDF = TempDirecory & dgvData.CurrentRow.Cells(0).Value.ToString
        frmShowPDF.ShowDialog()
        '// Open with default PDF Reader.
        'Process.Start(TempDirecory & dgvData.CurrentRow.Cells(0).Value.ToString)
    End Sub

    ' / --------------------------------------------------------------------------
    Private Sub InitDataGridView()
        Me.dgvData.Columns.Clear()
        '// Initialize DataGridView Control
        With dgvData
            .AllowUserToAddRows = False
            .AllowUserToDeleteRows = False
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            .AutoResizeColumns()
            .AllowUserToResizeColumns = True
            .AllowUserToResizeRows = True
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
        End With
        '// Declare columns type.
        '// Add 1th column (Index = 0), Show PDF file name.
        Dim FName As New DataGridViewTextBoxColumn()
        dgvData.Columns.Add(FName)
        With FName
            .HeaderText = "File Name"
            .ReadOnly = True
            .Visible = True
        End With
        '//
        Me.dgvData.Focus()
    End Sub

    Private Sub cmbCamera_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbCamera.SelectedIndexChanged
        WebCamMgr.CurrentCamera = WebCamMgr.Cameras.ElementAt(cmbCamera.SelectedIndex)
        '//
        'WebCamMgr.CurrentCamera.CaptureHeight = 480
        'WebCamMgr.CurrentCamera.CaptureWidth = 640
    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        picFeed.Image = WebCamMgr.CurrentCamera.GetCurrentImage()
    End Sub

    Private Sub btnCapture_Click(sender As System.Object, e As System.EventArgs) Handles btnCapture.Click
        picPreview.Image = WebCamMgr.CurrentCamera.GetCurrentImage()
        btnSave.Enabled = True
    End Sub

    '// Delete PDF & PNG File.
    Private Sub btnDelete_Click(sender As System.Object, e As System.EventArgs) Handles btnDelete.Click
        If Me.dgvData.Rows.Count = 0 Then Return
        Try
            Dim sPic As String = dgvData.CurrentRow.Cells(0).Value.ToString
            sPic = Mid(sPic, 1, Len(sPic) - 3) & "png"
            '// Delete files in folder.
            FileSystem.Kill(TempDirecory & dgvData.CurrentRow.Cells(0).Value.ToString)
            FileSystem.Kill(TempDirecory & sPic)
            '// Delete current row from dgvData
            dgvData.Rows.Remove(dgvData.CurrentRow)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    '// Merge PDF Pages to File.
    Private Sub btnMergePDF_Click(sender As System.Object, e As System.EventArgs) Handles btnMergePDF.Click
        If dgvData.RowCount = 0 Then Return
        Try
            Dim FinalDoc As New PdfDocument()
            '/ Creates a string array of source files to be merged.
            Dim Source As String() = New String(dgvData.Rows.Count - 1) {}
            For i As Integer = 0 To dgvData.Rows.Count - 1
                Source(i) = TempDirecory & dgvData.Rows(i).Cells(0).Value.ToString
            Next
            '/ Merges PDFDocument.
            PdfDocument.Merge(FinalDoc, Source)
            Dim dlgSaveFile As New SaveFileDialog()
            With dlgSaveFile
                .Filter = "PDF|*.pdf"
                .Title = "Merge PDF File"
                .DefaultExt = "pdf"
                .InitialDirectory = FinalDirecory
                .RestoreDirectory = True
            End With
            If dlgSaveFile.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                '// FinalFilePDF at modeFunction.vb
                FilePDF = dlgSaveFile.FileName
                '/ Saves the final document
                FinalDoc.Save(FilePDF)
                '/ Closes the document
                FinalDoc.Close(True)
                '/ Display new PDF with FinalFilePDF
                frmShowPDF.ShowDialog()
                '// Open with default PDF Reader.
                'Process.Start(FinalFilePDF)
            End If
            '//
            FinalDoc.Dispose()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    '// เลื่อนแถวขึ้น
    Private Sub btnUp_Click(sender As System.Object, e As System.EventArgs) Handles btnUp.Click
        With Me.dgvData
            '// หาค่า Index แถวที่เลือก
            Dim RowIndex As Integer = .SelectedCells(0).OwningRow.Index
            '// หาก Index = 0 แสดงว่าเป็นแถวบนสุด ให้จบออกจากโปรแกรมย่อย
            If RowIndex = 0 Then Return
            '//
            Dim Col As Integer = .SelectedCells(0).OwningColumn.Index
            Dim Rows As DataGridViewRowCollection = .Rows
            '// เก็บค่าแถวที่เลือก
            Dim Row As DataGridViewRow = Rows(RowIndex)
            '// ลบแถวที่เลือกออก
            Rows.Remove(Row)
            '// ไปเพิ่มแถวใหม่ ก่อนแถวที่เลือก 1 แถว (ก็เลยเสมือนมันเคลื่อนย้ายแถวได้)
            Rows.Insert(RowIndex - 1, Row)
            '// เคลียร์การเลือกแถว
            .ClearSelection()
            '// โฟกัสรายการแถวที่เลื่อนขึ้นไปแทรก
            .Rows(RowIndex - 1).Cells(Col).Selected = True
        End With

    End Sub

    '// เลื่อนแถวลง
    Private Sub btnDown_Click(sender As System.Object, e As System.EventArgs) Handles btnDown.Click
        With Me.dgvData
            Dim RowIndex As Integer = .SelectedCells(0).OwningRow.Index
            '// เช็คแถวสุดท้าย หากใช่ให้รีเทิร์นกลับ คือไม่ต้องทำอะไร
            If RowIndex = .Rows.Count - 1 Then Return
            '//
            Dim Col As Integer = .SelectedCells(0).OwningColumn.Index
            Dim Rows As DataGridViewRowCollection = .Rows
            '// เก็บค่าแถวที่เลือก
            Dim Row As DataGridViewRow = Rows(RowIndex)
            '// ลบแถวที่เลือกออก
            Rows.Remove(Row)
            '// ไปเพิ่มแถวใหม่ หลังแถวที่เลือก 1 แถว
            Rows.Insert(RowIndex + 1, Row)
            '// เคลียร์การเลือกแถว
            .ClearSelection()
            '// โฟกัสรายการแถวที่เลื่อนลงไป
            .Rows(RowIndex + 1).Cells(Col).Selected = True
        End With
    End Sub

    '// Row number of DataGridView.
    Private Sub dgvData_RowPostPaint(sender As Object, e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles dgvData.RowPostPaint
        Using b As SolidBrush = New SolidBrush(Me.dgvData.RowHeadersDefaultCellStyle.ForeColor)
            e.Graphics.DrawString(Convert.ToString(e.RowIndex + 1, CultureInfo.CurrentUICulture), e.InheritedRowStyle.Font, b, e.RowBounds.Location.X + 20, e.RowBounds.Location.Y + 4)
        End Using
    End Sub
End Class
