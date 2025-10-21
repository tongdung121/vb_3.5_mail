Imports System.Net
Imports System.Net.Mail
Imports System.Data.OracleClient

Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            Dim connStr As String = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=localhost)(PORT=1522))(CONNECT_DATA=(SERVICE_NAME=ORCLPDB)));User Id=dba_19;Password=123456;"

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


                Dim sql As String = "SELECT EMAIL FROM DBA_19.""USERS"""


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
            Dim fromAddress As String = "tongvanthienvu@gmail.com"
            Dim subject As String = "Thông báo tự động"
            Dim body As String = "<b>Xin chào!</b><br>Bạn nhận được email test từ hệ thống VB.NET 3.5."

            Dim mail As New MailMessage(fromAddress, toAddress, subject, body)
            mail.IsBodyHtml = True

            Dim smtp As New SmtpClient("smtp.gmail.com", 587)
            smtp.EnableSsl = True
            smtp.Credentials = New NetworkCredential("tongvanthienvu@gmail.com", "taryauulxecajvsx")

            smtp.Send(mail)

            Console.WriteLine("Đã gửi tới: " & toAddress)
        Catch ex As Exception
            Console.WriteLine("Lỗi khi gửi tới " & toAddress & ": " & ex.Message)
        End Try
    End Sub



End Class
