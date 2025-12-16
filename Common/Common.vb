Imports ZXing
Imports ZXing.Common
Imports System.Drawing
Imports ZXing.QrCode
Imports ZXing.QrCode.Internal

Public Class Common
    Public address As String = "example@gmail.com"
    Public pass As String = ""
    Public MAIL_HOST As String = ""

    Public Function GenerateQRCode(ByVal text As String) As Bitmap
        Dim writer As New BarcodeWriter()
        writer.Format = BarcodeFormat.QR_CODE
        writer.Options = New EncodingOptions With {.Height = 200,.Width = 200,.Margin = 1}

        ' Trả ảnh Bitmap ra ngoài
        Return writer.Write(text)
    End Function

    Public Function GenerateQRCode1(ByVal text As String) As Bitmap
        Dim writer As New BarcodeWriter()
        writer.Format = BarcodeFormat.QR_CODE

        ' Sử dụng QrCodeEncodingOptions thay vì EncodingOptions
        Dim options As New QrCodeEncodingOptions()

        options.Width = 200
        options.Height = 200
        options.Margin = 1

        ' Thiết lập mức sửa lỗi (L, M, Q, H)
        options.ErrorCorrection = ErrorCorrectionLevel.H   ' <- thay H bằng L/M/Q/H tùy ý

        writer.Options = options

        ' Trả ảnh Bitmap ra ngoài
        Return writer.Write(text)
    End Function

    Public Function ImageToByteArray(ByVal img As Bitmap) As Byte()
        Using ms As New System.IO.MemoryStream()
            img.Save(ms, System.Drawing.Imaging.ImageFormat.Png)
            Return ms.ToArray()
        End Using
    End Function


End Class
