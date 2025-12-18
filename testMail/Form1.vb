Imports System.Net
Imports System.Net.Mail
Imports System.Data.OracleClient
Imports CrystalDecisions.Shared
Imports System.IO
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports QRCoder
Imports CrystalDecisions.CrystalReports.Engine


Public Class Form1
    Dim c As New Common.Common()
    Dim connStr As String = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=localhost)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ORCL19)));User Id=tongdung;Password=10Vini@10;"

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            Using conn As New OracleConnection(connStr)
                conn.Open()

                Console.WriteLine("✅ Kết nối thành công tới: " & conn.DataSource)

                Dim dbNameCmd As New OracleCommand("SELECT sys_context('USERENV','DB_NAME') FROM dual", conn)
                Dim dbName As String = dbNameCmd.ExecuteScalar().ToString()
                Console.WriteLine("📦 Đang kết nối đến DB: " & dbName)

                Dim whoCmd As New OracleCommand("SELECT USER FROM dual", conn)
                Console.WriteLine("👤 User hiện tại: " & whoCmd.ExecuteScalar().ToString())

                Dim pdbCmd As New OracleCommand("SELECT sys_context('USERENV','CON_NAME') FROM dual", conn)
                Console.WriteLine("🏷️ Container: " & pdbCmd.ExecuteScalar().ToString())

                Dim tablesCmd As New OracleCommand("SELECT COUNT(*) FROM all_tables WHERE table_name = 'USERS'", conn)
                Console.WriteLine("📦 Số bảng USERS thấy được: " & tablesCmd.ExecuteScalar().ToString())


                Dim sql As String = "SELECT EMAIL FROM TONGDUNG.""USERS"""


                Using cmd As New OracleCommand(sql, conn)
                    Using reader As OracleDataReader = cmd.ExecuteReader()
                        Dim found As Boolean = False
                        While reader.Read()
                            found = True
                            Dim email As String = reader("email").ToString()
                            Console.WriteLine("📧 Email: " & email)

                            If email <> "" Then
                                SendEmail(email)
                            End If
                        End While

                        If Not found Then
                            Console.WriteLine("⚠️ Không tìm thấy dòng nào trong bảng USERS.")
                        End If
                    End Using
                End Using
            End Using

            Console.WriteLine("✅ Hoàn tất gửi email!")

        Catch ex As Exception
            Console.WriteLine("❌ Lỗi: " & ex.Message)
        End Try
    End Sub


    ' ====== 2️⃣ Hàm gửi email ======
    Private Sub SendEmail(ByVal toAddress As String)
        Try
            Dim fromAddress As String = c.address
            Dim subject As String = "Thông báo tự động"
            Dim body As String = "<b>Xin chào!</b><br>Bạn nhận được email test từ hệ thống VB.NET 3.5."

            Dim mail As New MailMessage(fromAddress, toAddress, subject, body)
            mail.IsBodyHtml = True

            Dim smtp As New SmtpClient(c.MAIL_HOST, 587)
            smtp.EnableSsl = True
            smtp.Credentials = New NetworkCredential(c.address, c.pass)

            smtp.Send(mail)

            Console.WriteLine("Đã gửi tới: " & toAddress)
        Catch ex As Exception
            Console.WriteLine("Lỗi khi gửi tới " & toAddress & ": " & ex.Message)
        End Try
    End Sub



    Private Sub qr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles qr.Click
        Try
            Using conn As New OracleConnection(connStr)
                conn.Open()
                Dim sql As String = "SELECT * FROM TONGDUNG.""FRAME"""

                Using cmd As New OracleCommand(sql, conn)
                    Using reader As OracleDataReader = cmd.ExecuteReader()
                        Dim found As Boolean = False

                        While reader.Read()
                            found = True
                            Console.WriteLine("📧 Email: " & reader("QRCode").ToString())
                            Dim qrBitmap As Bitmap = c.GenerateQRCode(reader("QRCode").ToString())
                            Dim qrBitmap1 As Bitmap = c.GenerateQRCode1(reader("QRCode").ToString())
                            ' Lưu tạm QR code
                            Dim tempFile As String = Path.Combine(Path.GetTempPath(), "qr_temp.png")
                            qrBitmap.Save(tempFile, Imaging.ImageFormat.Png)
                            Dim tempFile1 As String = Path.Combine(Path.GetTempPath(), "qr_temp1.png")
                            qrBitmap1.Save(tempFile1, Imaging.ImageFormat.Png)

                            ' DataTable có cột QRCodePath kiểu String
                            Dim dt As New DataTable()
                            dt.Columns.Add("Name", GetType(String))
                            dt.Columns.Add("QRCode", GetType(String))

                            Dim row As DataRow = dt.NewRow()
                            row("Name") = reader("Name").ToString()
                            row("QRCode") = reader("QRCode").ToString()
                            dt.Rows.Add(row)

                            ' DataTable cho Subreport
                            Dim dtqrt As New DataTable("DataTable1")
                            dtqrt.Columns.Add("QRCode", GetType(String))
                            dtqrt.Columns.Add("Image", GetType(Byte))

                            Dim qrRow As DataRow = dtqrt.NewRow()
                            qrRow("QRCode") = reader("QRCode").ToString()
                            dtqrt.Rows.Add(qrRow)
                            Dim dtQr As DataTable = AddQRCode(dtqrt)

                            ' 2. Xuất PDF từ Crystal Reports (.rpt) chứa barcode cũ
                            Dim rpt As New rptQRCode() ' report hiện tại
                            rpt.SetDataSource(dt)
                            rpt.Subreports(0).SetDataSource(dtQr)
                            Dim fileName As String
                            SaveFileDialog1.Filter = "PDF Files (*.pdf*)|*.pdf"
                            If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                                fileName = SaveFileDialog1.FileName
                            Else
                                Exit Sub
                            End If
                            rpt.ExportToDisk(ExportFormatType.PortableDocFormat, fileName)

                            ' 3. Mở PDF bằng iTextSharp, thêm QR code
                            Dim readerFile As New PdfReader(fileName)
                            Dim ms As New MemoryStream()
                            Dim stamper As New PdfStamper(readerFile, ms)
                            Dim content As PdfContentByte = stamper.GetOverContent(1) ' trang đầu

                            ' Load QR code từ file PNG
                            Dim qrImage As iTextSharp.text.Image = iTextSharp.text.Image.GetInstance(tempFile)
                            qrImage.SetAbsolutePosition(400, 600) ' điều chỉnh vị trí
                            qrImage.ScaleAbsolute(150, 150)      ' điều chỉnh kích thước
                            content.AddImage(qrImage)

                            Dim qrImage1 As iTextSharp.text.Image = iTextSharp.text.Image.GetInstance(tempFile1)
                            qrImage1.SetAbsolutePosition(400, 400) ' điều chỉnh vị trí
                            qrImage1.ScaleAbsolute(150, 150)      ' điều chỉnh kích thước
                            content.AddImage(qrImage1)

                            stamper.Close()
                            readerFile.Close()

                            ' Ghi PDF cuối cùng
                            File.WriteAllBytes(fileName, ms.ToArray())
                            MessageBox.Show("Xuất PDF thành công với barcode và QRCode")
                        End While

                        If Not found Then
                            Console.WriteLine("⚠️ Không tìm thấy dòng nào trong bảng FRAME.")
                        End If
                    End Using
                End Using
            End Using

        Catch ex As Exception
            Console.WriteLine("⚠️ ex " & ex.Message)
        End Try
    End Sub

    Private Function AddQRCode(ByVal dsData As DataTable) As DataTable
        Dim dtqr As New DataTable("DataTable1")
        dtqr.Columns.Add("key_link", GetType(String))   ' dùng để link Subreport
        dtqr.Columns.Add("Image", GetType(Byte()))

        For Each row As DataRow In dsData.Rows
            ' Tạo QRCode từ giá trị FRM1
            Dim qrGenerator As New QRCodeGenerator()
            Dim qrCodeData As QRCodeData = qrGenerator.CreateQrCode(row("QRCode").ToString(), QRCodeGenerator.ECCLevel.Q)
            Dim qrCode As New QRCoder.QRCode(qrCodeData)
            Using qrBitmap As Bitmap = qrCode.GetGraphic(20)
                ' Chuyển Bitmap → Byte[]
                Dim ms As New IO.MemoryStream()
                qrBitmap.Save(ms, System.Drawing.Imaging.ImageFormat.Png)
                Dim objByteData() As Byte = ms.ToArray()
                ms.Dispose()

                ' Thêm vào DataTable
                Dim drow As DataRow = dtqr.NewRow()
                drow("key_link") = row("QRCode")        ' dùng để link Subreport
                drow("Image") = objByteData
                dtqr.Rows.Add(drow)
            End Using
        Next

        Return dtqr
    End Function

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            Using conn As New OracleConnection(connStr)
                conn.Open()
                Dim sql As String = "SELECT * FROM TONGDUNG.""FRAME"""

                Using cmd As New OracleCommand(sql, conn)
                    Using reader As OracleDataReader = cmd.ExecuteReader()
                        Dim found As Boolean = False

                        While reader.Read()
                            found = True

                            ' DataTable có cột QRCodePath kiểu String
                            Dim dt As New DataTable()
                            dt.Columns.Add("Name", GetType(String))
                            dt.Columns.Add("QRCode", GetType(String))

                            Dim row As DataRow = dt.NewRow()
                            row("Name") = reader("Name").ToString()
                            row("QRCode") = reader("QRCode").ToString()
                            dt.Rows.Add(row)

                            ' DataTable cho Subreport
                            Dim dtqrt As New DataTable("DataTable1")
                            dtqrt.Columns.Add("QRCode", GetType(String))
                            dtqrt.Columns.Add("Image", GetType(Byte))

                            Dim qrRow As DataRow = dtqrt.NewRow()
                            qrRow("QRCode") = reader("QRCode").ToString()
                            dtqrt.Rows.Add(qrRow)
                            Dim dtQr As DataTable = AddQRCode(dtqrt)

                            ' 2. Xuất PDF từ Crystal Reports (.rpt) chứa barcode cũ
                            Dim rpt As New rptQRCode() ' report hiện tại
                            rpt.SetDataSource(dt)
                            Dim subRpt As ReportDocument = rpt.OpenSubreport("SubQR.rpt")
                            subRpt.SetDataSource(dtQr)

                            Dim fileName As String
                            SaveFileDialog1.Filter = "PDF Files (*.pdf*)|*.pdf"
                            If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                                fileName = SaveFileDialog1.FileName
                            Else
                                Exit Sub
                            End If
                            rpt.ExportToDisk(ExportFormatType.PortableDocFormat, fileName)
                            MessageBox.Show("Xuất PDF thành công với barcode và QRCode")
                        End While

                        If Not found Then
                            Console.WriteLine("⚠️ Không tìm thấy dòng nào trong bảng FRAME.")
                        End If
                    End Using
                End Using
            End Using

        Catch ex As Exception
            Console.WriteLine("⚠️ ex " & ex.Message)
        End Try
    End Sub
End Class
