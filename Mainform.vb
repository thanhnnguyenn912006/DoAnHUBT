Imports System.Windows.Forms
Imports System.Data.SqlClient

Public Class MainForm
    Inherits Form

    Private menuStrip As MenuStrip
    Private statusStrip As StatusStrip
    Private userType As String
    Private userCode As String
    Private connectionString As String = "Server=localhost\SQLEXPRESS;Database=QLSinhVien;Integrated Security=True"

    Public Sub New(userType As String, userCode As String)
        Me.userType = userType
        Me.userCode = userCode
        InitializeComponent()
        InitializeMenu()
    End Sub

    Private Sub InitializeComponent()
        ' Thiết lập form chính
        Me.Text = "Hệ thống quản lý giáo dục - " & Me.userType
        Me.Size = New Size(1000, 600)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.WindowState = FormWindowState.Maximized
        Me.IsMdiContainer = True
    End Sub

    Private Sub InitializeMenu()
        ' Tạo menu chính
        menuStrip = New MenuStrip()

        ' Menu Hệ thống
        Dim systemMenu As New ToolStripMenuItem("Hệ thống")
        Dim logoutItem As New ToolStripMenuItem("Đăng xuất")
        Dim exitItem As New ToolStripMenuItem("Thoát")
        AddHandler logoutItem.Click, AddressOf logoutItem_Click
        AddHandler exitItem.Click, AddressOf exitItem_Click
        systemMenu.DropDownItems.Add(logoutItem)
        systemMenu.DropDownItems.Add(exitItem)

        ' Menu Quản lý (chỉ hiển thị cho Admin)
        If userType = "Admin" Then
            Dim manageMenu As New ToolStripMenuItem("Quản lý")
            Dim studentManageItem As New ToolStripMenuItem("Quản lý sinh viên")
            Dim teacherManageItem As New ToolStripMenuItem("Quản lý giảng viên")
            Dim classManageItem As New ToolStripMenuItem("Quản lý lớp học")
            Dim scoreManageItem As New ToolStripMenuItem("Quản lý điểm")

            AddHandler studentManageItem.Click, AddressOf studentManageItem_Click
            AddHandler teacherManageItem.Click, AddressOf teacherManageItem_Click
            AddHandler classManageItem.Click, AddressOf classManageItem_Click
            AddHandler scoreManageItem.Click, AddressOf scoreManageItem_Click

            manageMenu.DropDownItems.Add(studentManageItem)
            manageMenu.DropDownItems.Add(teacherManageItem)
            manageMenu.DropDownItems.Add(classManageItem)
            manageMenu.DropDownItems.Add(scoreManageItem)

            menuStrip.Items.Add(manageMenu)
        End If


        ' Menu Học tập (hiển thị cho SinhVien)
        If userType = "SinhVien" Then
            Dim studyMenu As New ToolStripMenuItem("Học tập")
            Dim registerItem As New ToolStripMenuItem("Đăng ký học phần")
            Dim profileItem As New ToolStripMenuItem("Cập nhật thông tin")
            Dim viewScoreItem As New ToolStripMenuItem("Xem điểm tích lũy")

            AddHandler registerItem.Click, AddressOf registerItem_Click
            AddHandler profileItem.Click, AddressOf profileStudentItem_Click
            AddHandler viewScoreItem.Click, AddressOf viewScoreItem_Click

            studyMenu.DropDownItems.Add(registerItem)
            studyMenu.DropDownItems.Add(profileItem)
            studyMenu.DropDownItems.Add(viewScoreItem)
            menuStrip.Items.Add(studyMenu)
        End If

        ' Menu Giảng dạy (chỉ hiển thị cho Giảng viên)
        If userType = "GiaoVien" Then
            Dim teachingMenu As New ToolStripMenuItem("Giảng dạy")
            Dim classListItem As New ToolStripMenuItem("Xem lớp giảng dạy")
            Dim profileItem As New ToolStripMenuItem("Cập nhật thông tin")
            Dim inputScoreItem As New ToolStripMenuItem("Nhập điểm")

            AddHandler classListItem.Click, AddressOf classListTeachItem_Click
            AddHandler profileItem.Click, AddressOf profileTeacherItem_Click
            AddHandler inputScoreItem.Click, AddressOf inputScoreItem_Click

            teachingMenu.DropDownItems.Add(classListItem)
            teachingMenu.DropDownItems.Add(profileItem)
            teachingMenu.DropDownItems.Add(inputScoreItem)
            menuStrip.Items.Add(teachingMenu)
        End If

        ' Menu Trợ giúp
        Dim helpMenu As New ToolStripMenuItem("Trợ giúp")
        Dim aboutItem As New ToolStripMenuItem("Giới thiệu")
        AddHandler aboutItem.Click, AddressOf aboutItem_Click
        helpMenu.DropDownItems.Add(aboutItem)

        ' Thêm các menu vào menuStrip
        menuStrip.Items.Add(systemMenu)
        menuStrip.Items.Add(helpMenu)

        ' Tạo thanh trạng thái
        statusStrip = New StatusStrip()
        Dim statusLabel As New ToolStripStatusLabel()
        statusLabel.Text = "Xin chào: " & GetUserName() & " | Vai trò: " & userType
        statusStrip.Items.Add(statusLabel)

        ' Thêm các controls vào form
        Me.Controls.Add(menuStrip)
        Me.Controls.Add(statusStrip)
        Me.MainMenuStrip = menuStrip
    End Sub

    Private Function GetUserName() As String
        Try
            Using connection As New SqlConnection(connectionString)
                connection.Open()
                Dim query As String = ""

                If userType = "SinhVien" Then
                    query = "SELECT HoSV + ' ' + TenSV AS HoTen FROM SinhVien WHERE MaSV = @code"
                ElseIf userType = "GiaoVien" Then
                    query = "SELECT HoTenGV AS HoTen FROM GiangVien WHERE MaGV = @code"
                Else
                    Return "Quản trị viên"
                End If

                Using command As New SqlCommand(query, connection)
                    command.Parameters.AddWithValue("@code", userCode)
                    Dim result = command.ExecuteScalar()
                    If result IsNot Nothing Then
                        Return result.ToString()
                    End If
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Lỗi khi lấy thông tin người dùng: " & ex.Message)
        End Try

        Return "Không xác định"
    End Function

    Private Sub logoutItem_Click(sender As Object, e As EventArgs)
        Dim result As DialogResult = MessageBox.Show("Bạn có chắc muốn đăng xuất?", "Xác nhận",
                                                   MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If result = DialogResult.Yes Then
            Dim loginForm As New LoginForm()
            loginForm.Show()
            Me.Close()
        End If
    End Sub

    Private Sub exitItem_Click(sender As Object, e As EventArgs)
        Dim result As DialogResult = MessageBox.Show("Bạn có chắc muốn thoát ứng dụng?", "Xác nhận",
                                                   MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If result = DialogResult.Yes Then
            Application.Exit()
        End If
    End Sub

    Private Sub studentManageItem_Click(sender As Object, e As EventArgs)
        For Each childForm As Form In Me.MdiChildren
            If TypeOf childForm Is QLSinhVienForm Then
                childForm.BringToFront()
                Return
            End If
        Next

        Dim studentForm As New QLSinhVienForm()
        studentForm.MdiParent = Me
        studentForm.Show()
    End Sub

    Private Sub teacherManageItem_Click(sender As Object, e As EventArgs)
        For Each childForm As Form In Me.MdiChildren
            If TypeOf childForm Is QLGiangVienForm Then
                childForm.BringToFront()
                Return
            End If
        Next

        Dim teacherForm As New QLGiangVienForm()
        teacherForm.MdiParent = Me
        teacherForm.Show()
    End Sub



    Private Sub registerItem_Click(sender As Object, e As EventArgs)
        For Each childForm As Form In Me.MdiChildren
            If TypeOf childForm Is DangKyHPForm Then
                childForm.BringToFront()
                Return
            End If
        Next

        Dim registerForm As New DangKyHPForm(userCode)
        registerForm.MdiParent = Me
        registerForm.Show()
    End Sub

    Private Sub aboutItem_Click(sender As Object, e As EventArgs)
        Dim message As String = "HỆ THỐNG QUẢN LÝ GIÁO DỤC - EDU MANAGER v1.0" & Environment.NewLine & Environment.NewLine &
                           "Phát triển bởi: NNT Solutions" & Environment.NewLine &
                           "Phiên bản: 1.0 Release" & Environment.NewLine &
                           "Nền tảng: Windows Forms - SQL Server" & Environment.NewLine & Environment.NewLine &
                           "Tính năng chính:" & Environment.NewLine &
                           "✓ Quản lý sinh viên và lớp học" & Environment.NewLine &
                           "✓ Phân công giảng dạy" & Environment.NewLine &
                           "✓ Báo cáo và thống kê" & Environment.NewLine &
                           "✓ Phân quyền người dùng" & Environment.NewLine & Environment.NewLine &
                           "Hỗ trợ kỹ thuật: support@nnt.com"

        MessageBox.Show(message, "Giới thiệu hệ thống", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub classManageItem_Click(sender As Object, e As EventArgs)
        For Each childForm As Form In Me.MdiChildren
            If TypeOf childForm Is frmQuanLyLop Then
                childForm.BringToFront()
                Return
            End If
        Next

        Dim classForm As New frmQuanLyLop(userType, userCode)
        classForm.MdiParent = Me
        classForm.Show()
    End Sub

    Private Sub classListTeachItem_Click(sender As Object, e As EventArgs)
        For Each childForm As Form In Me.MdiChildren
            If TypeOf childForm Is frmQuanLyLop Then
                childForm.BringToFront()
                Return
            End If
        Next

        Dim classListForm As New frmQuanLyLop(userType, userCode)
        classListForm.MdiParent = Me
        classListForm.Show()
    End Sub

    Private Sub profileStudentItem_Click(sender As Object, e As EventArgs)
        OpenFormPersonalMode(New QLSinhVienForm(userCode))
    End Sub

    Private Sub profileTeacherItem_Click(sender As Object, e As EventArgs)
        OpenFormPersonalMode(New QLGiangVienForm(userCode))
    End Sub

    Private Sub OpenFormPersonalMode(form As Form)
        For Each childForm As Form In Me.MdiChildren
            If childForm.GetType() = form.GetType() Then
                childForm.BringToFront()
                Return
            End If
        Next

        form.MdiParent = Me
        form.Show()
    End Sub

    ' Phương thức cho Admin: Quản lý điểm (Toàn quyền)
    Private Sub scoreManageItem_Click(sender As Object, e As EventArgs)
        OpenDiemForm("Admin", Nothing) ' Truyền role "Admin" và không lọc theo user
    End Sub

    ' Phương thức cho Giảng viên: Nhập điểm (Chỉ những lớp/môn họ dạy)
    Private Sub inputScoreItem_Click(sender As Object, e As EventArgs)
        OpenDiemForm("GiaoVien", userCode) ' Truyền role "GiaoVien" và mã GV
    End Sub

    ' Phương thức cho Sinh viên: Xem điểm (Chỉ điểm của chính mình)
    Private Sub viewScoreItem_Click(sender As Object, e As EventArgs)
        OpenDiemForm("SinhVien", userCode) ' Truyền role "SinhVien" và mã SV
    End Sub

    ' Phương thức chung để mở form QLDiemfrm
    Private Sub OpenDiemForm(role As String, userCode As String)
        For Each childForm As Form In Me.MdiChildren
            If TypeOf childForm Is QLDiemForm Then
                childForm.BringToFront()
                Return
            End If
        Next

        Dim diemForm As New QLDiemForm(role, userCode)
        diemForm.MdiParent = Me
        diemForm.Show()
    End Sub

End Class