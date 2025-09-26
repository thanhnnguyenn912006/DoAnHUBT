Imports System.Windows.Forms
Imports System.Data
Imports System.Data.SqlClient

Public Class QLGiangVienForm
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
    Private currentMaGV As String = ""

    Public Sub New()
        InitializeComponent()
        LoadData()
    End Sub

    Public Sub New(maGV As String)
        Me.currentMaGV = maGV
        Me.isPersonalMode = True
        InitializeComponent()
        SetupPersonalMode()
        LoadPersonalData()
    End Sub

    Private Sub InitializeComponent()
        Me.Text = If(isPersonalMode, "Thông tin cá nhân", "Quản lý giảng viên")
        Me.Size = New Size(900, 500)
        Me.StartPosition = FormStartPosition.CenterScreen

        If Not isPersonalMode Then
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

            Me.Controls.Add(panelSearch)
        End If

        dataGridView = New DataGridView()
        dataGridView.Dock = If(isPersonalMode, DockStyle.Fill, DockStyle.None)
        If Not isPersonalMode Then
            dataGridView.Location = New Point(0, 40)
            dataGridView.Size = New Size(900, 410)
        End If
        dataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dataGridView.ReadOnly = True
        dataGridView.AllowUserToAddRows = False
        dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        Dim panelButtons As New Panel()
        panelButtons.Dock = DockStyle.Bottom
        panelButtons.Height = 50

        If Not isPersonalMode Then
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
        Else
            btnEdit = New Button()
            btnEdit.Text = "Cập nhật thông tin"
            btnEdit.Location = New Point(20, 10)
            btnEdit.Size = New Size(150, 30)
            AddHandler btnEdit.Click, AddressOf btnEditPersonal_Click

            panelButtons.Controls.Add(btnEdit)
        End If

        Me.Controls.Add(dataGridView)
        Me.Controls.Add(panelButtons)
    End Sub

    Private Sub SetupPersonalMode()
        dataGridView.AutoGenerateColumns = True
    End Sub

    Private Sub LoadData()
        Try
            Using connection As New SqlConnection(connectionString)
                connection.Open()
                Dim query As String = "SELECT MaGV, HoTenGV, GioiTinh, Email, DienThoai, MaKhoa, TrangThai FROM GiangVien"

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
                Dim query As String = "SELECT MaGV, HoTenGV, GioiTinh, Email, DienThoai, MaKhoa, TrangThai FROM GiangVien WHERE MaGV = @maGV"

                Using adapter As New SqlDataAdapter(query, connection)
                    adapter.SelectCommand.Parameters.AddWithValue("@maGV", currentMaGV)
                    Dim table As New DataTable()
                    adapter.Fill(table)

                    table.Columns.Add("GioiTinhDisplay", GetType(String), "IIF(GioiTinh = 1, 'Nam', 'Nữ')")

                    dataGridView.DataSource = table
                    dataGridView.Columns("GioiTinh").Visible = False
                    dataGridView.Columns("GioiTinhDisplay").HeaderText = "Giới tính"
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Lỗi khi tải dữ liệu cá nhân: " & ex.Message)
        End Try
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs)
        Try
            Using connection As New SqlConnection(connectionString)
                connection.Open()
                Dim query As String = "SELECT MaGV, HoTenGV, GioiTinh, Email, DienThoai, MaKhoa, TrangThai FROM GiangVien " &
                                      "WHERE MaGV LIKE @search OR HoTenGV LIKE @search OR MaKhoa LIKE @search"

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
        Dim addForm As New GiangVienDetailForm()
        If addForm.ShowDialog() = DialogResult.OK Then
            LoadData()
        End If
    End Sub

    Private Sub btnEdit_Click(sender As Object, e As EventArgs)
        If dataGridView.SelectedRows.Count > 0 Then
            Dim maGV As String = dataGridView.SelectedRows(0).Cells("MaGV").Value.ToString()
            Dim editForm As New GiangVienDetailForm(maGV)
            If editForm.ShowDialog() = DialogResult.OK Then
                LoadData()
            End If
        Else
            MessageBox.Show("Vui lòng chọn một giảng viên để sửa")
        End If
    End Sub

    Private Sub btnEditPersonal_Click(sender As Object, e As EventArgs)
        Dim editForm As New GiangVienDetailForm(currentMaGV, True)
        If editForm.ShowDialog() = DialogResult.OK Then
            LoadPersonalData()
        End If
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs)
        If dataGridView.SelectedRows.Count > 0 Then
            Dim maGV As String = dataGridView.SelectedRows(0).Cells("MaGV").Value.ToString()
            Dim result As DialogResult = MessageBox.Show("Bạn có chắc chắn muốn xóa giảng viên " & maGV & "?", "Xác nhận", MessageBoxButtons.YesNo)

            If result = DialogResult.Yes Then
                Try
                    Using connection As New SqlConnection(connectionString)
                        connection.Open()
                        Dim query As String = "DELETE FROM GiangVien WHERE MaGV = @maGV"

                        Using command As New SqlCommand(query, connection)
                            command.Parameters.AddWithValue("@maGV", maGV)
                            command.ExecuteNonQuery()
                            MessageBox.Show("Đã xóa giảng viên thành công")
                            LoadData()
                        End Using
                    End Using
                Catch ex As Exception
                    MessageBox.Show("Lỗi khi xóa giảng viên: " & ex.Message)
                End Try
            End If
        Else
            MessageBox.Show("Vui lòng chọn một giảng viên để xóa")
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

Public Class GiangVienDetailForm
    Inherits Form

    Private txtMaGV As TextBox
    Private txtHoTenGV As TextBox
    Private cbGioiTinh As ComboBox
    Private txtEmail As TextBox
    Private txtDienThoai As TextBox
    Private txtMaKhoa As TextBox
    Private cbTrangThai As ComboBox
    Private btnSave As Button
    Private btnCancel As Button

    Private connectionString As String = "Server=localhost\SQLEXPRESS;Database=QLSinhVien;Integrated Security=True"
    Private maGVEdit As String = Nothing
    Private isPersonalMode As Boolean = False

    Public Sub New(Optional maGV As String = Nothing, Optional isPersonalMode As Boolean = False)
        Me.maGVEdit = maGV
        Me.isPersonalMode = isPersonalMode
        InitializeComponent()

        If maGVEdit IsNot Nothing Then
            LoadGiangVien(maGVEdit)
        End If
    End Sub

    Private Sub InitializeComponent()
        Me.Text = If(maGVEdit Is Nothing, "Thêm Giảng Viên",
                    If(isPersonalMode, "Cập nhật thông tin cá nhân", "Sửa Giảng Viên"))
        Me.Size = New Size(400, 400)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        Dim lblMaGV As New Label() With {.Text = "Mã GV:", .Location = New Point(20, 20), .Width = 100}
        txtMaGV = New TextBox() With {.Location = New Point(120, 20), .Width = 200}
        If isPersonalMode Then txtMaGV.Enabled = False

        Dim lblHoTen As New Label() With {.Text = "Họ tên:", .Location = New Point(20, 60), .Width = 100}
        txtHoTenGV = New TextBox() With {.Location = New Point(120, 60), .Width = 200}

        Dim lblGioiTinh As New Label() With {.Text = "Giới tính:", .Location = New Point(20, 100), .Width = 100}
        cbGioiTinh = New ComboBox() With {.Location = New Point(120, 100), .Width = 200}
        cbGioiTinh.Items.AddRange(New Object() {"Nam", "Nữ"})

        Dim lblEmail As New Label() With {.Text = "Email:", .Location = New Point(20, 140), .Width = 100}
        txtEmail = New TextBox() With {.Location = New Point(120, 140), .Width = 200}

        Dim lblDienThoai As New Label() With {.Text = "Điện thoại:", .Location = New Point(20, 180), .Width = 100}
        txtDienThoai = New TextBox() With {.Location = New Point(120, 180), .Width = 200}

        Dim lblMaKhoa As New Label() With {.Text = "Mã khoa:", .Location = New Point(20, 220), .Width = 100}
        txtMaKhoa = New TextBox() With {.Location = New Point(120, 220), .Width = 200}
        If isPersonalMode Then txtMaKhoa.Enabled = False

        Dim lblTrangThai As New Label() With {.Text = "Trạng thái:", .Location = New Point(20, 260), .Width = 100}
        cbTrangThai = New ComboBox() With {.Location = New Point(120, 260), .Width = 200}
        cbTrangThai.Items.AddRange(New Object() {"Đang giảng dạy", "Nghỉ phép", "Đã nghỉ"})
        If isPersonalMode Then cbTrangThai.Enabled = False

        btnSave = New Button() With {.Text = "Lưu", .Location = New Point(120, 310), .Size = New Size(80, 30)}
        btnCancel = New Button() With {.Text = "Hủy", .Location = New Point(220, 310), .Size = New Size(80, 30)}

        AddHandler btnSave.Click, AddressOf btnSave_Click
        AddHandler btnCancel.Click, Sub() Me.DialogResult = DialogResult.Cancel

        Me.Controls.AddRange({lblMaGV, txtMaGV, lblHoTen, txtHoTenGV, lblGioiTinh, cbGioiTinh,
                              lblEmail, txtEmail, lblDienThoai, txtDienThoai, lblMaKhoa, txtMaKhoa,
                              lblTrangThai, cbTrangThai, btnSave, btnCancel})

        If maGVEdit IsNot Nothing Then txtMaGV.Enabled = False
    End Sub

    Private Sub LoadGiangVien(maGV As String)
        Try
            Using conn As New SqlConnection(connectionString)
                conn.Open()
                Dim query As String = "SELECT * FROM GiangVien WHERE MaGV = @maGV"
                Using cmd As New SqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@maGV", maGV)
                    Using reader = cmd.ExecuteReader()
                        If reader.Read() Then
                            txtMaGV.Text = reader("MaGV").ToString()
                            txtHoTenGV.Text = reader("HoTenGV").ToString()
                            Dim gioiTinhValue As Boolean = Convert.ToBoolean(reader("GioiTinh"))
                            cbGioiTinh.Text = If(gioiTinhValue, "Nam", "Nữ")

                            txtEmail.Text = reader("Email").ToString()
                            txtDienThoai.Text = reader("DienThoai").ToString()
                            txtMaKhoa.Text = reader("MaKhoa").ToString()
                            cbTrangThai.Text = reader("TrangThai").ToString()
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Lỗi khi tải dữ liệu giảng viên: " & ex.Message)
        End Try
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs)
        If String.IsNullOrEmpty(txtMaGV.Text) OrElse String.IsNullOrEmpty(txtHoTenGV.Text) Then
            MessageBox.Show("Vui lòng nhập đầy đủ thông tin bắt buộc")
            Return
        End If

        If String.IsNullOrEmpty(cbGioiTinh.Text) OrElse (cbGioiTinh.Text <> "Nam" AndAlso cbGioiTinh.Text <> "Nữ") Then
            MessageBox.Show("Vui lòng chọn giới tính hợp lệ (Nam hoặc Nữ)")
            Return
        End If

        Dim gioiTinhBit As Boolean = (cbGioiTinh.Text = "Nam")

        Try
            Using conn As New SqlConnection(connectionString)
                conn.Open()
                Dim query As String

                If maGVEdit Is Nothing Then
                    query = "INSERT INTO GiangVien (MaGV, HoTenGV, GioiTinh, Email, DienThoai, MaKhoa, TrangThai) " &
                            "VALUES (@MaGV, @HoTenGV, @GioiTinh, @Email, @DienThoai, @MaKhoa, @TrangThai)"
                Else
                    If isPersonalMode Then
                        query = "UPDATE GiangVien SET HoTenGV=@HoTenGV, GioiTinh=@GioiTinh, " &
                                "Email=@Email, DienThoai=@DienThoai WHERE MaGV=@MaGV"
                    Else
                        query = "UPDATE GiangVien SET HoTenGV=@HoTenGV, GioiTinh=@GioiTinh, " &
                                "Email=@Email, DienThoai=@DienThoai, MaKhoa=@MaKhoa, " &
                                "TrangThai=@TrangThai WHERE MaGV=@MaGV"
                    End If
                End If

                Using cmd As New SqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@MaGV", txtMaGV.Text)
                    cmd.Parameters.AddWithValue("@HoTenGV", txtHoTenGV.Text)
                    cmd.Parameters.AddWithValue("@GioiTinh", gioiTinhBit)
                    cmd.Parameters.AddWithValue("@Email", If(String.IsNullOrEmpty(txtEmail.Text), DBNull.Value, txtEmail.Text))
                    cmd.Parameters.AddWithValue("@DienThoai", If(String.IsNullOrEmpty(txtDienThoai.Text), DBNull.Value, txtDienThoai.Text))

                    If Not isPersonalMode OrElse maGVEdit Is Nothing Then
                        cmd.Parameters.AddWithValue("@MaKhoa", txtMaKhoa.Text)
                        cmd.Parameters.AddWithValue("@TrangThai", cbTrangThai.Text)
                    End If

                    cmd.ExecuteNonQuery()
                    MessageBox.Show("Lưu thành công!")
                    Me.DialogResult = DialogResult.OK
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Lỗi khi lưu dữ liệu: " & ex.Message)
        End Try
    End Sub
End Class