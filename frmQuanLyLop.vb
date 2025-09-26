Imports System.Data.SqlClient
Imports System.Windows.Forms

Public Class frmQuanLyLop
    Inherits Form

    Private userType As String
    Private userCode As String
    Private connectionString As String = "Server=localhost\SQLEXPRESS;Database=QLSinhVien;Integrated Security=True"
    Private dgvLopHoc As DataGridView

    Public Sub New(userType As String, Optional userCode As String = "")
        Me.userType = userType
        Me.userCode = userCode
        InitializeForm()
        LoadData()
    End Sub

    Private Sub InitializeForm()
        Me.Text = "Quản lý lớp học - " & userType
        Me.Size = New Size(900, 500)
        Me.StartPosition = FormStartPosition.CenterScreen

        ' Tạo DataGridView - ĐÃ SỬA
        dgvLopHoc = New DataGridView()
        dgvLopHoc.Name = "dgvLopHoc"
        dgvLopHoc.Dock = DockStyle.Fill
        dgvLopHoc.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgvLopHoc.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvLopHoc.ReadOnly = True
        Me.Controls.Add(dgvLopHoc)

        If userType = "Admin" Then
            Dim panel As New Panel()
            panel.Dock = DockStyle.Bottom
            panel.Height = 50

            Dim btnThem As New Button()
            btnThem.Text = "Thêm lớp"
            btnThem.Location = New Point(10, 10)
            btnThem.Size = New Size(80, 30)
            AddHandler btnThem.Click, AddressOf btnThem_Click

            Dim btnSua As New Button()
            btnSua.Text = "Sửa lớp"
            btnSua.Location = New Point(100, 10)
            btnSua.Size = New Size(80, 30)
            AddHandler btnSua.Click, AddressOf btnSua_Click

            Dim btnXoa As New Button()
            btnXoa.Text = "Xóa lớp"
            btnXoa.Location = New Point(190, 10)
            btnXoa.Size = New Size(80, 30)
            AddHandler btnXoa.Click, AddressOf btnXoa_Click

            panel.Controls.Add(btnThem)
            panel.Controls.Add(btnSua)
            panel.Controls.Add(btnXoa)
            Me.Controls.Add(panel)
        End If
    End Sub

    Private Sub LoadData()
        Try
            Using connection As New SqlConnection(connectionString)
                connection.Open()
                Dim query As String = ""
                Dim command As SqlCommand

                If userType = "Admin" Then
                    ' Admin xem tất cả lớp
                    query = "SELECT L.MaLop, L.TenLop, N.TenNganh, K.TenKhoa, L.Khoahoc, L.SiSo " &
                            "FROM Lop L " &
                            "INNER JOIN Nganh N ON L.MaNganh = N.MaNganh " &
                            "INNER JOIN Khoa K ON L.Makhoa = K.Makhoa " &
                            "ORDER BY L.MaLop"
                    command = New SqlCommand(query, connection)
                ElseIf userType = "GiaoVien" Then
                    ' Giảng viên chỉ xem lớp của mình
                    query = "SELECT L.MaLop, L.TenLop, N.TenNganh, K.TenKhoa, L.Khoahoc, L.SiSo " &
                            "FROM Lop L " &
                            "INNER JOIN Nganh N ON L.MaNganh = N.MaNganh " &
                            "INNER JOIN Khoa K ON L.Makhoa = K.Makhoa " &
                            "INNER JOIN PhanCongGiangDay PC ON L.MaLop = PC.MaLop " &
                            "WHERE PC.MaGV = @MaGV " &
                            "ORDER BY L.MaLop"
                    command = New SqlCommand(query, connection)
                    command.Parameters.AddWithValue("@MaGV", userCode)
                Else
                    MessageBox.Show("Bạn không có quyền xem thông tin này")
                    Return
                End If

                Dim adapter As New SqlDataAdapter(command)
                Dim table As New DataTable()
                adapter.Fill(table)


                dgvLopHoc.DataSource = table


                If table.Columns.Count > 0 Then
                    dgvLopHoc.Columns("MaLop").HeaderText = "Mã Lớp"
                    dgvLopHoc.Columns("TenLop").HeaderText = "Tên Lớp"
                    dgvLopHoc.Columns("TenNganh").HeaderText = "Ngành"
                    dgvLopHoc.Columns("TenKhoa").HeaderText = "Khoa"
                    dgvLopHoc.Columns("Khoahoc").HeaderText = "Khóa Học"
                    dgvLopHoc.Columns("SiSo").HeaderText = "Sĩ Số"
                End If


                If table.Rows.Count = 0 Then
                    MessageBox.Show("Không có dữ liệu lớp học để hiển thị")
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show("Lỗi khi tải dữ liệu: " & ex.Message)
        End Try
    End Sub

    Private Sub btnThem_Click(sender As Object, e As EventArgs)
        Dim frmThem As New frmThemLop()
        If frmThem.ShowDialog() = DialogResult.OK Then
            LoadData()
        End If
    End Sub

    Private Sub btnSua_Click(sender As Object, e As EventArgs)
        If dgvLopHoc.SelectedRows.Count > 0 Then
            Dim maLop As String = dgvLopHoc.SelectedRows(0).Cells("MaLop").Value.ToString()
            Dim frmSua As New frmSuaLop(maLop)
            If frmSua.ShowDialog() = DialogResult.OK Then
                LoadData()
            End If
        Else
            MessageBox.Show("Vui lòng chọn một lớp để sửa")
        End If
    End Sub

    Private Sub btnXoa_Click(sender As Object, e As EventArgs)
        If dgvLopHoc.SelectedRows.Count > 0 Then
            Dim maLop As String = dgvLopHoc.SelectedRows(0).Cells("MaLop").Value.ToString()
            Dim tenLop As String = dgvLopHoc.SelectedRows(0).Cells("TenLop").Value.ToString()

            Dim result As DialogResult = MessageBox.Show($"Bạn có chắc muốn xóa lớp '{tenLop}'?",
                                                        "Xác nhận xóa",
                                                        MessageBoxButtons.YesNo,
                                                        MessageBoxIcon.Question)

            If result = DialogResult.Yes Then
                Try
                    Using connection As New SqlConnection(connectionString)
                        connection.Open()


                        Dim queryDeletePC As String = "DELETE FROM PhanCongGiangDay WHERE MaLop = @MaLop"
                        Using commandDeletePC As New SqlCommand(queryDeletePC, connection)
                            commandDeletePC.Parameters.AddWithValue("@MaLop", maLop)
                            commandDeletePC.ExecuteNonQuery()
                        End Using


                        Dim queryDeleteLop As String = "DELETE FROM Lop WHERE MaLop = @MaLop"
                        Using commandDeleteLop As New SqlCommand(queryDeleteLop, connection)
                            commandDeleteLop.Parameters.AddWithValue("@MaLop", maLop)
                            commandDeleteLop.ExecuteNonQuery()
                            MessageBox.Show("Xóa lớp thành công")
                            LoadData()
                        End Using
                    End Using
                Catch ex As Exception
                    MessageBox.Show("Lỗi khi xóa lớp: " & ex.Message)
                End Try
            End If
        Else
            MessageBox.Show("Vui lòng chọn một lớp để xóa")
        End If
    End Sub
End Class