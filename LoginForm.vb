Imports System.Data.SqlClient
Imports System.Windows.Forms

Public Class LoginForm
    Inherits Form

    Private txtUsername As TextBox
    Private txtPassword As TextBox
    Private btnLogin As Button
    Private lblTitle As Label
    Private lblUsername As Label
    Private lblPassword As Label

    Private connectionString As String = "Server=localhost\SQLEXPRESS;Database=QLSinhVien;Integrated Security=True"

    Public Property UserType As String
    Public Property UserCode As String

    Public Sub New()
        InitializeControls()
    End Sub

    Private Sub InitializeControls()
        ' Thiết lập form
        Me.Text = "Đăng nhập hệ thống"
        Me.Size = New Size(400, 300)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False

        ' Tạo controls
        lblTitle = New Label()
        lblTitle.Text = "HỆ THỐNG QUẢN LÝ GIÁO DỤC"
        lblTitle.Font = New Font("Arial", 14, FontStyle.Bold)
        lblTitle.AutoSize = True
        lblTitle.Location = New Point(80, 30)

        lblUsername = New Label()
        lblUsername.Text = "Tên đăng nhập:"
        lblUsername.AutoSize = True
        lblUsername.Location = New Point(50, 100)

        txtUsername = New TextBox()
        txtUsername.Location = New Point(160, 100)
        txtUsername.Size = New Size(180, 20)

        lblPassword = New Label()
        lblPassword.Text = "Mật khẩu:"
        lblPassword.AutoSize = True
        lblPassword.Location = New Point(50, 140)

        txtPassword = New TextBox()
        txtPassword.Location = New Point(160, 140)
        txtPassword.Size = New Size(180, 20)
        txtPassword.PasswordChar = "*"c

        btnLogin = New Button()
        btnLogin.Text = "Đăng nhập"
        btnLogin.Location = New Point(160, 190)
        btnLogin.Size = New Size(100, 30)
        AddHandler btnLogin.Click, AddressOf btnLogin_Click

        Me.Controls.Add(lblTitle)
        Me.Controls.Add(lblUsername)
        Me.Controls.Add(txtUsername)
        Me.Controls.Add(lblPassword)
        Me.Controls.Add(txtPassword)
        Me.Controls.Add(btnLogin)
    End Sub

    Private Sub btnLogin_Click(sender As Object, e As EventArgs)
        ' Kiểm tra trống
        If String.IsNullOrEmpty(txtUsername.Text) OrElse String.IsNullOrEmpty(txtPassword.Text) Then
            MessageBox.Show("Vui lòng nhập tên đăng nhập và mật khẩu!", "Cảnh báo",
              MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try
            Using connection As New SqlConnection(connectionString)
                connection.Open()
                Dim query As String = "SELECT LoaiTaiKhoan, MaNguoiDung FROM TaiKhoan 
                         WHERE TenDangNhap = @username AND MatKhau = @password 
                         AND TrangThai = 'KichHoat'"

                Using command As New SqlCommand(query, connection)
                    command.Parameters.AddWithValue("@username", txtUsername.Text)
                    command.Parameters.AddWithValue("@password", txtPassword.Text)

                    Using reader As SqlDataReader = command.ExecuteReader()
                        If reader.Read() Then
                            ' Lưu thông tin vào property
                            UserType = reader("LoaiTaiKhoan").ToString()
                            UserCode = reader("MaNguoiDung").ToString()

                            MessageBox.Show("Đăng nhập thành công! ", "Thành công",
                              MessageBoxButtons.OK, MessageBoxIcon.Information)

                            Me.Hide() ' Ẩn form login

                            Dim mainForm As New MainForm(UserType, UserCode)
                            AddHandler mainForm.FormClosed, AddressOf MainForm_FormClosed
                            mainForm.Show()
                        Else
                            MessageBox.Show("Tên đăng nhập hoặc mật khẩu không đúng!", "Lỗi",
                              MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Lỗi kết nối cơ sở dữ liệu: " & ex.Message, "Lỗi",
              MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub MainForm_FormClosed(sender As Object, e As FormClosedEventArgs)
        Me.Close() ' Đóng LoginForm khi MainForm đóng
    End Sub
End Class