#Region "ABOUT"
' / -----------------------------------------------------------------------------------------------
' / Developer : Mr.Surapon Yodsanga (Thongkorn Tubtimkrob)
' / eMail : thongkorn@hotmail.com
' / URL: http://www.g2gnet.com (Khon Kaen - Thailand)
' / Facebook: https://www.facebook.com/g2gnet (For Thailand only)
' / Facebook: https://www.facebook.com/commonindy (Worldwide)
' / More Info: http://www.g2gnet.com/webboard
' /
' / Purpose: Show PDF Files with Syncfusion Community.
' / Microsoft Visual Basic .Net 2010 + Service Pack 1
' / Syncfusion Community License.
' /
' / This is open source code under @Copyleft by Thongkorn Tubtimkrob.
' / You can modify and/or distribute without to inform the developer.
' / -----------------------------------------------------------------------------------------------
#End Region

Imports Syncfusion.Pdf.Parsing
Imports Syncfusion.Pdf
Imports System.IO

Public Class frmShowPDF

    Private Sub frmShowPDF_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
        GC.SuppressFinalize(Me)
    End Sub

    Private Sub frmShowPDF_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        '// Open PDF File.
        Me.PdfViewerControl1.Load(FilePDF, "")
    End Sub
End Class
