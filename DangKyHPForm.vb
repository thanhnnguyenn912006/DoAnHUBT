Imports System.Data.SqlClient
Imports System.Windows.Forms

Public Class DangKyHPForm
    Inherits Form
    Private connectionString As String = "Server=localhost\SQLEXPRESS;Database=QLSinhVien;Integrated Security=True"

    ' Khai báo các controls
    Private WithEvents cmbHocKy As ComboBox
    Private WithEvents cmbNamHoc As ComboBox
    Private WithEvents cmbMonHoc As ComboBox
    Private WithEvents btnRefresh As Button
    Private dgvLopHocPhan As DataGridView
    Private WithEvents btnDangKy As Button
    Private WithEvents btnHuyDangKy As Button
    Private userCode As String

    ' Constructor nhận tham số userCode
    Public Sub New(userCode As String)
        Me.userCode = userCode
        InitializeControls()
    End Sub

    Private Sub InitializeControls()
        ' Thiết lập form
        Me.Text = "Đăng Ký Học Phần - Mã SV: " & userCode
        Me.Size = New Size(800, 600)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.MinimumSize = New Size(800, 600)

        ' Tạo và thiết lập các controls
        SetupControls()
        LoadHocKyNamHoc()
        LoadMonHoc()
        HienThiThongTinSinhVien()
    End Sub

    Private Sub SetupControls()
        ' Panel chứa các filter
        Dim filterPanel As New Panel()
        filterPanel.Dock = DockStyle.Top
        filterPanel.Height = 100
        filterPanel.BackColor = Color.LightSteelBlue

        Dim lblHocKy As New Label()
        lblHocKy.Text = "Học kỳ:"
        lblHocKy.Location = New Point(20, 20)
        lblHocKy.Width = 50
        lblHocKy.Font = New Font("Microsoft Sans Serif", 9, FontStyle.Bold)

        cmbHocKy = New ComboBox()
        cmbHocKy.Location = New Point(80, 17)
        cmbHocKy.Width = 80
        cmbHocKy.DropDownStyle = ComboBoxStyle.DropDownList
        cmbHocKy.Font = New Font("Microsoft Sans Serif", 9)

        Dim lblNamHoc As New Label()
        lblNamHoc.Text = "Năm học:"
        lblNamHoc.Location = New Point(180, 20)
        lblNamHoc.Width = 60
        lblNamHoc.Font = New Font("Microsoft Sans Serif", 9, FontStyle.Bold)

        cmbNamHoc = New ComboBox()
        cmbNamHoc.Location = New Point(250, 17)
        cmbNamHoc.Width = 100
        cmbNamHoc.DropDownStyle = ComboBoxStyle.DropDownList
        cmbNamHoc.Font = New Font("Microsoft Sans Serif", 9)

        Dim lblMonHoc As New Label()
        lblMonHoc.Text = "Môn học:"
        lblMonHoc.Location = New Point(370, 20)
        lblMonHoc.Width = 60
        lblMonHoc.Font = New Font("Microsoft Sans Serif", 9, FontStyle.Bold)

        cmbMonHoc = New ComboBox()
        cmbMonHoc.Location = New Point(440, 17)
        cmbMonHoc.Width = 200
        cmbMonHoc.DropDownStyle = ComboBoxStyle.DropDownList
        cmbMonHoc.Font = New Font("Microsoft Sans Serif", 9)
        AddHandler cmbMonHoc.SelectedIndexChanged, AddressOf cmbMonHoc_SelectedIndexChanged

        btnRefresh = New Button()
        btnRefresh.Text = "Làm mới"
        btnRefresh.Location = New Point(660, 15)
        btnRefresh.Width = 80
        btnRefresh.Font = New Font("Microsoft Sans Serif", 9)
        AddHandler btnRefresh.Click, AddressOf btnRefresh_Click

        filterPanel.Controls.Add(lblHocKy)
        filterPanel.Controls.Add(cmbHocKy)
        filterPanel.Controls.Add(lblNamHoc)
        filterPanel.Controls.Add(cmbNamHoc)
        filterPanel.Controls.Add(lblMonHoc)
        filterPanel.Controls.Add(cmbMonHoc)
        filterPanel.Controls.Add(btnRefresh)

        dgvLopHocPhan = New DataGridView()
        dgvLopHocPhan.Dock = DockStyle.Fill
        dgvLopHocPhan.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvLopHocPhan.ReadOnly = True
        dgvLopHocPhan.AllowUserToAddRows = False
        dgvLopHocPhan.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgvLopHocPhan.Font = New Font("Microsoft Sans Serif", 9)

        Dim buttonPanel As New Panel()
        buttonPanel.Dock = DockStyle.Bottom
        buttonPanel.Height = 60
        buttonPanel.BackColor = Color.LightSteelBlue

        btnDangKy = New Button()
        btnDangKy.Text = "Đăng ký"
        btnDangKy.Size = New Size(100, 35)
        btnDangKy.Location = New Point(300, 12)
        btnDangKy.Font = New Font("Microsoft Sans Serif", 9, FontStyle.Bold)
        btnDangKy.BackColor = Color.LightGreen
        AddHandler btnDangKy.Click, AddressOf btnDangKy_Click

        btnHuyDangKy = New Button()
        btnHuyDangKy.Text = "Hủy đăng ký"
        btnHuyDangKy.Size = New Size(100, 35)
        btnHuyDangKy.Location = New Point(420, 12)
        btnHuyDangKy.Font = New Font("Microsoft Sans Serif", 9, FontStyle.Bold)
        btnHuyDangKy.BackColor = Color.LightCoral
        AddHandler btnHuyDangKy.Click, AddressOf btnHuyDangKy_Click

        buttonPanel.Controls.Add(btnDangKy)
        buttonPanel.Controls.Add(btnHuyDangKy)

        Me.Controls.Add(dgvLopHocPhan)
        Me.Controls.Add(filterPanel)
        Me.Controls.Add(buttonPanel)
    End Sub

    Private Sub LoadHocKyNamHoc()
        cmbHocKy.Items.Add("1")
        cmbHocKy.Items.Add("2")
        cmbHocKy.Items.Add("3")
        cmbHocKy.Items.Add("4")
        cmbHocKy.Items.Add("5")
        cmbHocKy.Items.Add("6")
        cmbHocKy.Items.Add("7")
        cmbHocKy.Items.Add("8")
        cmbHocKy.SelectedIndex = 0

        cmbNamHoc.Items.Add("2024-2025")
        cmbNamHoc.Items.Add("2023-2024")
        cmbNamHoc.SelectedIndex = 0
    End Sub

    Private Sub LoadMonHoc()
        Try
            Using conn As New SqlConnection(connectionString)
                conn.Open()
                Dim query As String = "SELECT MaMon, TenMon FROM Mon ORDER BY TenMon"
                Using cmd As New SqlCommand(query, conn)
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        cmbMonHoc.Items.Clear()
                        While reader.Read()
                            cmbMonHoc.Items.Add(New ComboboxItem(reader("TenMon").ToString(), reader("MaMon").ToString()))
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Lỗi tải môn học: " & ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cmbMonHoc_SelectedIndexChanged(sender As Object, e As EventArgs)
        LoadLopHocPhan()
    End Sub

    Private Sub LoadLopHocPhan()
        Try
            If cmbMonHoc.SelectedItem Is Nothing Then
                dgvLopHocPhan.DataSource = Nothing
                Return
            End If

            Dim selectedMon As String = DirectCast(cmbMonHoc.SelectedItem, ComboboxItem).Value.ToString()
            Dim hocKy As String = cmbHocKy.SelectedItem.ToString()
            Dim namHoc As String = cmbNamHoc.SelectedItem.ToString()

            Using conn As New SqlConnection(connectionString)
                conn.Open()
                Dim query As String = "SELECT lhp.MaLHP, lhp.TenLHP, gv.HoTenGV, lhp.SoLuongDK, lhp.SiSoToiDa 
                                      FROM LopHocPhan lhp 
                                      INNER JOIN GiangVien gv ON lhp.MaGV = gv.MaGV 
                                      WHERE lhp.MaMon = @MaMon AND lhp.HocKy = @HocKy AND lhp.NamHoc = @NamHoc"

                Using cmd As New SqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@MaMon", selectedMon)
                    cmd.Parameters.AddWithValue("@HocKy", hocKy)
                    cmd.Parameters.AddWithValue("@NamHoc", namHoc)

                    Using adapter As New SqlDataAdapter(cmd)
                        Dim dt As New DataTable()
                        adapter.Fill(dt)

                        If dt.Rows.Count = 0 Then
                            MessageBox.Show("Không có lớp học phần nào cho môn học này!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If

                        dgvLopHocPhan.DataSource = dt
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Lỗi tải lớp học phần: " & ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs)
        LoadLopHocPhan()
    End Sub

    Private Function KiemTraSinhVienTonTai(maSV As String) As Boolean
        Try
            Using conn As New SqlConnection(connectionString)
                conn.Open()
                Dim query As String = "SELECT COUNT(*) FROM SinhVien WHERE MaSV = @MaSV"

                Using cmd As New SqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@MaSV", maSV)
                    Dim count As Integer = CInt(cmd.ExecuteScalar())
                    Return count > 0
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Lỗi kiểm tra sinh viên: " & ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    Private Function KiemTraDaDangKy(maSV As String, maLHP As String) As Boolean
        Try
            Using conn As New SqlConnection(connectionString)
                conn.Open()
                Dim query As String = "SELECT COUNT(*) FROM DangKyHocPhan WHERE MaSV = @MaSV AND MaLHP = @MaLHP"

                Using cmd As New SqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@MaSV", maSV)
                    cmd.Parameters.AddWithValue("@MaLHP", maLHP)
                    Dim count As Integer = CInt(cmd.ExecuteScalar())
                    Return count > 0
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Lỗi kiểm tra đăng ký: " & ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function
    Private Sub btnDangKy_Click(sender As Object, e As EventArgs)
        If dgvLopHocPhan.SelectedRows.Count = 0 Then
            MessageBox.Show("Vui lòng chọn lớp học phần để đăng ký!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        If Not KiemTraSinhVienTonTai(userCode) Then
            MessageBox.Show("Mã sinh viên không tồn tại trong hệ thống! Vui lòng liên hệ quản trị viên.", "Lỗi",
                      MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        Dim maLHP As String = dgvLopHocPhan.SelectedRows(0).Cells("MaLHP").Value.ToString()
        Dim tenLHP As String = dgvLopHocPhan.SelectedRows(0).Cells("TenLHP").Value.ToString()

        Try
            Using conn As New SqlConnection(connectionString)
                conn.Open()

                If KiemTraDaDangKy(userCode, maLHP) Then
                    MessageBox.Show("Bạn đã đăng ký lớp học phần này rồi!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                Dim query As String = "INSERT INTO DangKyHocPhan (MaSV, MaLHP, TrangThai) 
                                  VALUES (@MaSV, @MaLHP, N'Chờ duyệt')"

                Using cmd As New SqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@MaSV", userCode)
                    cmd.Parameters.AddWithValue("@MaLHP", maLHP)

                    Dim result As Integer = cmd.ExecuteNonQuery()
                    If result > 0 Then
                        MessageBox.Show($"Đăng ký thành công lớp: {tenLHP}! Vui lòng chờ duyệt.", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        LoadLopHocPhan()
                    End If
                End Using
            End Using
        Catch ex As SqlException
            If ex.Number = 2627 Then
                MessageBox.Show("Bạn đã đăng ký lớp học phần này rồi!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            ElseIf ex.Number = 547 Then
                MessageBox.Show("Lỗi: Mã sinh viên không hợp lệ hoặc không tồn tại!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                MessageBox.Show("Lỗi đăng ký: " & ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            MessageBox.Show("Lỗi đăng ký: " & ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub HienThiThongTinSinhVien()
        Try
            Using conn As New SqlConnection(connectionString)
                conn.Open()
                Dim query As String = "SELECT HoSV + ' ' + TenSV AS HoTen, Lop FROM SinhVien WHERE MaSV = @MaSV"

                Using cmd As New SqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@MaSV", userCode)
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        If reader.Read() Then
                            Me.Text = "Đăng Ký Học Phần - " & reader("HoTen").ToString() & " - Lớp: " & reader("Lop").ToString()
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception

        End Try
    End Sub
    Private Sub btnHuyDangKy_Click(sender As Object, e As EventArgs)
        If dgvLopHocPhan.SelectedRows.Count = 0 Then
            MessageBox.Show("Vui lòng chọn lớp học phần để hủy!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim maLHP As String = dgvLopHocPhan.SelectedRows(0).Cells("MaLHP").Value.ToString()
        Dim tenLHP As String = dgvLopHocPhan.SelectedRows(0).Cells("TenLHP").Value.ToString()

        Dim result As DialogResult = MessageBox.Show($"Bạn có chắc muốn hủy đăng ký lớp: {tenLHP}?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If result = DialogResult.Yes Then
            Try
                Using conn As New SqlConnection(connectionString)
                    conn.Open()
                    Dim query As String = "DELETE FROM DangKyHocPhan WHERE MaSV = @MaSV AND MaLHP = @MaLHP"

                    Using cmd As New SqlCommand(query, conn)
                        cmd.Parameters.AddWithValue("@MaSV", userCode)
                        cmd.Parameters.AddWithValue("@MaLHP", maLHP)

                        Dim rowsAffected As Integer = cmd.ExecuteNonQuery()
                        If rowsAffected > 0 Then
                            MessageBox.Show("Hủy đăng ký thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            LoadLopHocPhan()
                        Else
                            MessageBox.Show("Không tìm thấy đăng ký để hủy!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If
                    End Using
                End Using
            Catch ex As Exception
                MessageBox.Show("Lỗi hủy đăng ký: " & ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub
End Class

Public Class ComboboxItem
    Public Property Text As String
    Public Property Value As String

    Public Sub New(text As String, value As String)
        Me.Text = text
        Me.Value = value
    End Sub

    Public Overrides Function ToString() As String
        Return Text
    End Function
End Class