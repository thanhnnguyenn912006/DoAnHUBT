Imports System.Windows.Forms
Imports System.Data
Imports System.Data.SqlClient

Public Class QLSinhVienForm
    Inherits Form

    Private dataGridView As DataGridView
    Private btnAdd As Button
    Private btnEdit As Button
    Private btnDelete As Button
    Private btnRefresh As Button
    Private btnSearch As Button
    Private txtSearch As TextBox
    Private connectionString As String = "Server=localhost\SQLEXPRESS;Database=QLSinhVien;Integrated Security=True"
    Private isPersonalMode As Boolean = False
    Private studentCode As String = ""

    Public Sub New(studentCode As String)
        Me.studentCode = studentCode
        Me.isPersonalMode = True
        InitializeComponent()
        SetupPersonalMode()
        LoadPersonalData()
    End Sub

    Public Sub New()
        InitializeComponent()
        LoadData()
    End Sub

    Private Sub InitializeComponent()
        Me.Text = "Quản lý sinh viên"
        Me.Size = New Size(900, 500)
        Me.StartPosition = FormStartPosition.CenterScreen

        Dim panelSearch As New Panel()
        panelSearch.Dock = DockStyle.Top
        panelSearch.Height = 40

        Dim lblSearch As New Label()
        lblSearch.Text = "Tìm kiếm:"
        lblSearch.Location = New Point(10, 10)
        lblSearch.AutoSize = True

        txtSearch = New TextBox()
        txtSearch.Location = New Point(80, 10)
        txtSearch.Size = New Size(200, 20)

        btnSearch = New Button()
        btnSearch.Text = "Tìm"
        btnSearch.Location = New Point(290, 10)
        btnSearch.Size = New Size(60, 23)
        AddHandler btnSearch.Click, AddressOf btnSearch_Click

        panelSearch.Controls.Add(lblSearch)
        panelSearch.Controls.Add(txtSearch)
        panelSearch.Controls.Add(btnSearch)

        dataGridView = New DataGridView()
        dataGridView.Dock = DockStyle.Fill
        dataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dataGridView.ReadOnly = True
        dataGridView.AllowUserToAddRows = False
        dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        Dim panelButtons As New Panel()
        panelButtons.Dock = DockStyle.Bottom
        panelButtons.Height = 50

        btnAdd = New Button()
        btnAdd.Text = "Thêm"
        btnAdd.Location = New Point(20, 10)
        btnAdd.Size = New Size(80, 30)
        AddHandler btnAdd.Click, AddressOf btnAdd_Click

        btnEdit = New Button()
        btnEdit.Text = "Sửa"
        btnEdit.Location = New Point(120, 10)
        btnEdit.Size = New Size(80, 30)
        AddHandler btnEdit.Click, AddressOf btnEdit_Click

        btnDelete = New Button()
        btnDelete.Text = "Xóa"
        btnDelete.Location = New Point(220, 10)
        btnDelete.Size = New Size(80, 30)
        AddHandler btnDelete.Click, AddressOf btnDelete_Click

        btnRefresh = New Button()
        btnRefresh.Text = "Làm mới"
        btnRefresh.Location = New Point(320, 10)
        btnRefresh.Size = New Size(80, 30)
        AddHandler btnRefresh.Click, AddressOf btnRefresh_Click

        panelButtons.Controls.Add(btnAdd)
        panelButtons.Controls.Add(btnEdit)
        panelButtons.Controls.Add(btnDelete)
        panelButtons.Controls.Add(btnRefresh)

        Me.Controls.Add(dataGridView)
        Me.Controls.Add(panelSearch)
        Me.Controls.Add(panelButtons)
    End Sub

    Private Sub SetupPersonalMode()
        If isPersonalMode Then
            Me.Text = "Thông tin cá nhân - Sinh viên"

            btnAdd.Visible = False
            btnDelete.Visible = False
            btnRefresh.Visible = False

            btnEdit.Location = New Point(150, 10)
            btnEdit.Text = "Cập nhật thông tin"

            Me.Controls(1).Visible = False
        End If
    End Sub

    Private Sub LoadData()
        Try
            Using connection As New SqlConnection(connectionString)
                connection.Open()
                Dim query As String = "SELECT MaSV, HoSV, TenSV, GioiTinh, NgaySinh, MaLop, Email, DienThoai, DiaChi, TrangThai FROM SinhVien"

                Using adapter As New SqlDataAdapter(query, connection)
                    Dim table As New DataTable()
                    adapter.Fill(table)
                    table.Columns.Add("GioiTinhDisplay", GetType(String), "IIF(GioiTinh = 1, 'Nam', 'Nữ')")
                    dataGridView.DataSource = table
                    dataGridView.Columns("GioiTinh").Visible = False
                    dataGridView.Columns("GioiTinhDisplay").HeaderText = "Giới tính"
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Lỗi khi tải dữ liệu: " & ex.Message)
        End Try
    End Sub

    Private Sub LoadPersonalData()
        Try
            Using connection As New SqlConnection(connectionString)
                connection.Open()
                Dim query As String = "SELECT MaSV, HoSV, TenSV, GioiTinh, NgaySinh, MaLop, Email, DienThoai, DiaChi, TrangThai FROM SinhVien WHERE MaSV = @maSV"

                Using adapter As New SqlDataAdapter(query, connection)
                    adapter.SelectCommand.Parameters.AddWithValue("@maSV", studentCode)
                    Dim table As New DataTable()
                    adapter.Fill(table)
                    table.Columns.Add("GioiTinhDisplay", GetType(String), "IIF(GioiTinh = 1, 'Nam', 'Nữ')")
                    dataGridView.DataSource = table
                    dataGridView.Columns("GioiTinh").Visible = False
                    dataGridView.Columns("GioiTinhDisplay").HeaderText = "Giới tính"
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Lỗi khi tải thông tin cá nhân: " & ex.Message)
        End Try
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs)
        Try
            Using connection As New SqlConnection(connectionString)
                connection.Open()
                Dim query As String = "SELECT MaSV, HoSV, TenSV, GioiTinh, NgaySinh, MaLop, Email, DienThoai, DiaChi, TrangThai FROM SinhVien " &
                                      "WHERE MaSV LIKE @search OR HoSV LIKE @search OR TenSV LIKE @search OR MaLop LIKE @search"

                Using adapter As New SqlDataAdapter(query, connection)
                    adapter.SelectCommand.Parameters.AddWithValue("@search", "%" & txtSearch.Text & "%")
                    Dim table As New DataTable()
                    adapter.Fill(table)
                    table.Columns.Add("GioiTinhDisplay", GetType(String), "IIF(GioiTinh = 1, 'Nam', 'Nữ')")
                    dataGridView.DataSource = table
                    dataGridView.Columns("GioiTinh").Visible = False
                    dataGridView.Columns("GioiTinhDisplay").HeaderText = "Giới tính"
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Lỗi khi tìm kiếm: " & ex.Message)
        End Try
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs)
        Dim addForm As New SinhVienDetailForm()
        If addForm.ShowDialog() = DialogResult.OK Then
            LoadData()
        End If
    End Sub

    Private Sub btnEdit_Click(sender As Object, e As EventArgs)
        If isPersonalMode Then
            Dim editForm As New SinhVienDetailForm(studentCode, isPersonalMode)
            If editForm.ShowDialog() = DialogResult.OK Then
                LoadPersonalData()
            End If
        Else
            If dataGridView.SelectedRows.Count > 0 Then
                Dim maSV As String = dataGridView.SelectedRows(0).Cells("MaSV").Value.ToString()
                Dim editForm As New SinhVienDetailForm(maSV, isPersonalMode)
                If editForm.ShowDialog() = DialogResult.OK Then
                    LoadData()
                End If
            Else
                MessageBox.Show("Vui lòng chọn một sinh viên để sửa")
            End If
        End If
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs)
        If dataGridView.SelectedRows.Count > 0 Then
            Dim maSV As String = dataGridView.SelectedRows(0).Cells("MaSV").Value.ToString()
            Dim result As DialogResult = MessageBox.Show("Bạn có chắc chắn muốn xóa sinh viên " & maSV & "?", "Xác nhận", MessageBoxButtons.YesNo)

            If result = DialogResult.Yes Then
                Try
                    Using connection As New SqlConnection(connectionString)
                        connection.Open()
                        Dim query As String = "DELETE FROM SinhVien WHERE MaSV = @maSV"

                        Using command As New SqlCommand(query, connection)
                            command.Parameters.AddWithValue("@maSV", maSV)
                            command.ExecuteNonQuery()
                            MessageBox.Show("Đã xóa sinh viên thành công")
                            LoadData()
                        End Using
                    End Using
                Catch ex As Exception
                    MessageBox.Show("Lỗi khi xóa sinh viên: " & ex.Message)
                End Try
            End If
        Else
            MessageBox.Show("Vui lòng chọn một sinh viên để xóa")
        End If
    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs)
        If isPersonalMode Then
            LoadPersonalData()
        Else
            LoadData()
        End If
    End Sub
End Class

Public Class SinhVienDetailForm
    Inherits Form

    Private txtMaSV As TextBox
    Private txtHoSV As TextBox
    Private txtTenSV As TextBox
    Private cbGioiTinh As ComboBox
    Private dtpNgaySinh As DateTimePicker
    Private txtMaLop As TextBox
    Private txtEmail As TextBox
    Private txtDienThoai As TextBox
    Private txtDiaChi As TextBox
    Private cbTrangThai As ComboBox
    Private btnSave As Button
    Private btnCancel As Button
    Private isEditMode As Boolean = False
    Private isPersonalMode As Boolean = False
    Private connectionString As String = "Server=localhost\SQLEXPRESS;Database=QLSinhVien;Integrated Security=True"

    Public Sub New()
        InitializeComponent()
        Me.Text = "Thêm sinh viên mới"
    End Sub

    Public Sub New(maSV As String, Optional personalMode As Boolean = False)
        InitializeComponent()
        Me.Text = "Sửa thông tin sinh viên"
        isEditMode = True
        isPersonalMode = personalMode
        txtMaSV.Text = maSV
        txtMaSV.Enabled = False
        LoadStudentData(maSV)
    End Sub

    Private Sub InitializeComponent()
        Me.Size = New Size(400, 400)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False

        Dim lblMaSV As New Label() With {.Text = "Mã SV:", .Location = New Point(20, 20), .AutoSize = True}
        txtMaSV = New TextBox() With {.Location = New Point(120, 20), .Size = New Size(200, 20)}

        Dim lblHoSV As New Label() With {.Text = "Họ SV:", .Location = New Point(20, 50), .AutoSize = True}
        txtHoSV = New TextBox() With {.Location = New Point(120, 50), .Size = New Size(200, 20)}

        Dim lblTenSV As New Label() With {.Text = "Tên SV:", .Location = New Point(20, 80), .AutoSize = True}
        txtTenSV = New TextBox() With {.Location = New Point(120, 80), .Size = New Size(200, 20)}

        Dim lblGioiTinh As New Label() With {.Text = "Giới tính:", .Location = New Point(20, 110), .AutoSize = True}
        cbGioiTinh = New ComboBox() With {.Location = New Point(120, 110), .Size = New Size(200, 20)}
        cbGioiTinh.Items.AddRange({"Nam", "Nữ"})
        cbGioiTinh.SelectedIndex = 0

        Dim lblNgaySinh As New Label() With {.Text = "Ngày sinh:", .Location = New Point(20, 140), .AutoSize = True}
        dtpNgaySinh = New DateTimePicker() With {.Location = New Point(120, 140), .Size = New Size(200, 20)}

        Dim lblMaLop As New Label() With {.Text = "Mã lớp:", .Location = New Point(20, 170), .AutoSize = True}
        txtMaLop = New TextBox() With {.Location = New Point(120, 170), .Size = New Size(200, 20)}

        Dim lblEmail As New Label() With {.Text = "Email:", .Location = New Point(20, 200), .AutoSize = True}
        txtEmail = New TextBox() With {.Location = New Point(120, 200), .Size = New Size(200, 20)}

        Dim lblDienThoai As New Label() With {.Text = "Điện thoại:", .Location = New Point(20, 230), .AutoSize = True}
        txtDienThoai = New TextBox() With {.Location = New Point(120, 230), .Size = New Size(200, 20)}

        Dim lblDiaChi As New Label() With {.Text = "Địa chỉ:", .Location = New Point(20, 260), .AutoSize = True}
        txtDiaChi = New TextBox() With {.Location = New Point(120, 260), .Size = New Size(200, 20)}

        Dim lblTrangThai As New Label() With {.Text = "Trạng thái:", .Location = New Point(20, 290), .AutoSize = True}
        cbTrangThai = New ComboBox() With {.Location = New Point(120, 290), .Size = New Size(200, 20)}
        cbTrangThai.Items.AddRange({"Đang học", "Tạm ngừng", "Đã tốt nghiệp", "Đã thôi học"})
        cbTrangThai.SelectedIndex = 0

        btnSave = New Button() With {.Text = "Lưu", .Location = New Point(120, 320), .Size = New Size(80, 30)}
        AddHandler btnSave.Click, AddressOf btnSave_Click

        btnCancel = New Button() With {.Text = "Hủy", .Location = New Point(220, 320), .Size = New Size(80, 30)}
        AddHandler btnCancel.Click, AddressOf btnCancel_Click

        Me.Controls.AddRange({
            lblMaSV, txtMaSV,
            lblHoSV, txtHoSV,
            lblTenSV, txtTenSV,
            lblGioiTinh, cbGioiTinh,
            lblNgaySinh, dtpNgaySinh,
            lblMaLop, txtMaLop,
            lblEmail, txtEmail,
            lblDienThoai, txtDienThoai,
            lblDiaChi, txtDiaChi,
            lblTrangThai, cbTrangThai,
            btnSave, btnCancel
        })
    End Sub

    Private Sub LoadStudentData(maSV As String)
        Try
            Using connection As New SqlConnection(connectionString)
                connection.Open()
                Dim query As String = "SELECT HoSV, TenSV, GioiTinh, NgaySinh, MaLop, Email, DienThoai, DiaChi, TrangThai FROM SinhVien WHERE MaSV = @maSV"

                Using command As New SqlCommand(query, connection)
                    command.Parameters.AddWithValue("@maSV", maSV)

                    Using reader As SqlDataReader = command.ExecuteReader()
                        If reader.Read() Then
                            txtHoSV.Text = reader("HoSV").ToString()
                            txtTenSV.Text = reader("TenSV").ToString()
                            Dim gioiTinhValue As Boolean = Convert.ToBoolean(reader("GioiTinh"))
                            cbGioiTinh.SelectedIndex = If(gioiTinhValue, 0, 1)

                            dtpNgaySinh.Value = Convert.ToDateTime(reader("NgaySinh"))
                            txtMaLop.Text = reader("MaLop").ToString()
                            txtEmail.Text = reader("Email").ToString()
                            txtDienThoai.Text = reader("DienThoai").ToString()
                            txtDiaChi.Text = reader("DiaChi").ToString()

                            Dim trangThai As String = reader("TrangThai").ToString()
                            For i As Integer = 0 To cbTrangThai.Items.Count - 1
                                If cbTrangThai.Items(i).ToString() = trangThai Then
                                    cbTrangThai.SelectedIndex = i
                                    Exit For
                                End If
                            Next
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Lỗi khi tải dữ liệu sinh viên: " & ex.Message)
        End Try
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs)
        If String.IsNullOrEmpty(txtMaSV.Text) OrElse String.IsNullOrEmpty(txtHoSV.Text) OrElse String.IsNullOrEmpty(txtTenSV.Text) Then
            MessageBox.Show("Vui lòng nhập đầy đủ thông tin bắt buộc")
            Return
        End If

        Try
            Using conn As New SqlConnection(connectionString)
                conn.Open()
                Dim query As String
                Dim gioiTinhBit As Boolean = (cbGioiTinh.SelectedIndex = 0)

                If isEditMode Then
                    If isPersonalMode Then
                        query = "UPDATE SinhVien SET HoSV=@HoSV, TenSV=@TenSV, GioiTinh=@GioiTinh, " &
                            "NgaySinh=@NgaySinh, Email=@Email, DienThoai=@DienThoai, DiaChi=@DiaChi WHERE MaSV=@MaSV"
                    Else
                        query = "UPDATE SinhVien SET HoSV=@HoSV, TenSV=@TenSV, GioiTinh=@GioiTinh, " &
                            "NgaySinh=@NgaySinh, Email=@Email, DienThoai=@DienThoai, DiaChi=@DiaChi, MaLop=@MaLop, " &
                            "TrangThai=@TrangThai WHERE MaSV=@MaSV"
                    End If
                Else
                    query = "INSERT INTO SinhVien (MaSV, HoSV, TenSV, GioiTinh, NgaySinh, MaLop, Email, DienThoai, DiaChi, TrangThai) " &
                        "VALUES (@MaSV, @HoSV, @TenSV, @GioiTinh, @NgaySinh, @MaLop, @Email, @DienThoai, @DiaChi, @TrangThai)"
                End If

                Using cmd As New SqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@MaSV", txtMaSV.Text)
                    cmd.Parameters.AddWithValue("@HoSV", txtHoSV.Text)
                    cmd.Parameters.AddWithValue("@TenSV", txtTenSV.Text)
                    cmd.Parameters.AddWithValue("@GioiTinh", gioiTinhBit)
                    cmd.Parameters.AddWithValue("@NgaySinh", dtpNgaySinh.Value)
                    cmd.Parameters.AddWithValue("@Email", If(String.IsNullOrEmpty(txtEmail.Text), DBNull.Value, txtEmail.Text))
                    cmd.Parameters.AddWithValue("@DienThoai", If(String.IsNullOrEmpty(txtDienThoai.Text), DBNull.Value, txtDienThoai.Text))
                    cmd.Parameters.AddWithValue("@DiaChi", If(String.IsNullOrEmpty(txtDiaChi.Text), DBNull.Value, txtDiaChi.Text))

                    If Not isPersonalMode OrElse Not isEditMode Then
                        cmd.Parameters.AddWithValue("@MaLop", If(String.IsNullOrEmpty(txtMaLop.Text), DBNull.Value, txtMaLop.Text))
                        cmd.Parameters.AddWithValue("@TrangThai", cbTrangThai.Text)
                    End If

                    cmd.ExecuteNonQuery()
                    MessageBox.Show("Lưu thành công!")
                    Me.DialogResult = DialogResult.OK
                    Me.Close()
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Lỗi khi lưu dữ liệu: " & ex.Message)
        End Try
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs)
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
End Class