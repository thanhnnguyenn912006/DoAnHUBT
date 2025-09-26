Imports System.Windows.Forms
Imports System.Data.SqlClient
Imports System.Drawing

Public Class QLDiemForm
    Private currentUserRole As String
    Private currentUserCode As String

    Private WithEvents DataGridView1 As New DataGridView()
    Private WithEvents btnThem As New Button()
    Private WithEvents btnXoa As New Button()
    Private WithEvents btnSua As New Button()
    Private WithEvents btnLuu As New Button()

    Public Sub New(role As String, userCode As String)
        InitializeCustomComponent()
        currentUserRole = role
        currentUserCode = userCode

        Me.Text = If(role = "SinhVien", "Xem điểm tích lũy",
                    If(role = "GiaoVien", "Nhập điểm", "Quản lý điểm"))
    End Sub

    Private Sub InitializeCustomComponent()
        Me.Size = New Size(1000, 600)
        Me.StartPosition = FormStartPosition.CenterScreen

        ' Tạo panel chứa các nút
        Dim panel As New Panel()
        panel.Dock = DockStyle.Top
        panel.Height = 50
        panel.BackColor = Color.LightGray

        ' Thiết lập các nút
        btnThem.Text = "Thêm"
        btnThem.Size = New Size(80, 30)
        btnThem.Location = New Point(20, 10)
        btnThem.BackColor = Color.LightGreen

        btnXoa.Text = "Xóa"
        btnXoa.Size = New Size(80, 30)
        btnXoa.Location = New Point(110, 10)
        btnXoa.BackColor = Color.LightCoral

        btnSua.Text = "Sửa"
        btnSua.Size = New Size(80, 30)
        btnSua.Location = New Point(200, 10)
        btnSua.BackColor = Color.LightBlue

        btnLuu.Text = "Lưu"
        btnLuu.Size = New Size(80, 30)
        btnLuu.Location = New Point(290, 10)
        btnLuu.BackColor = Color.LightYellow


        panel.Controls.Add(btnThem)
        panel.Controls.Add(btnXoa)
        panel.Controls.Add(btnSua)
        panel.Controls.Add(btnLuu)

        DataGridView1.Dock = DockStyle.Fill
        DataGridView1.AllowUserToAddRows = False
        DataGridView1.AllowUserToDeleteRows = False
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect


        Me.Controls.Add(DataGridView1)
        Me.Controls.Add(panel)


        AddHandler Me.Load, AddressOf QLDiemForm_Load
    End Sub


    Private Sub QLDiemForm_Load(sender As Object, e As EventArgs)
        Select Case currentUserRole
            Case "Admin"
                SetupForAdmin()
            Case "GiaoVien"
                SetupForGiaoVien()
            Case "SinhVien"
                SetupForSinhVien()
        End Select
        LoadData()
    End Sub

    Private Sub SetupForAdmin()
        btnThem.Visible = True
        btnXoa.Visible = True
        btnSua.Visible = True
        btnLuu.Visible = True
        DataGridView1.ReadOnly = False
    End Sub

    Private Sub SetupForGiaoVien()
        btnThem.Visible = False
        btnXoa.Visible = False
        btnSua.Visible = True
        btnLuu.Visible = True

        DataGridView1.ReadOnly = False
        For Each col As DataGridViewColumn In DataGridView1.Columns
            If col.Name <> "Diem" Then
                col.ReadOnly = True
            End If
        Next
    End Sub

    Private Sub SetupForSinhVien()
        btnThem.Visible = False
        btnXoa.Visible = False
        btnSua.Visible = False
        btnLuu.Visible = False
        DataGridView1.ReadOnly = True
    End Sub

    Private Sub LoadData()
        Try
            Dim connectionString As String = "Server=localhost\SQLEXPRESS;Database=QLSinhVien;Integrated Security=True"
            Using connection As New SqlConnection(connectionString)
                connection.Open()

                Dim query As String = "SELECT d.MaSV, sv.HoSV, sv.TenSV, d.MaMon, m.TenMon, d.LanThi, d.Diem, d.NgayThi, d.KetQua, d.MaLHP " &
                                      "FROM Diem d " &
                                      "INNER JOIN SinhVien sv ON d.MaSV = sv.MaSV " &
                                      "INNER JOIN Mon m ON d.MaMon = m.MaMon " &
                                      "WHERE 1=1"

                If currentUserRole = "SinhVien" Then
                    query &= " AND d.MaSV = @UserCode"
                ElseIf currentUserRole = "GiaoVien" Then

                    query &= " AND d.MaMon IN (SELECT MaMon FROM LopHocPhan WHERE MaGV = @UserCode)"
                End If

                query &= " ORDER BY d.MaSV, m.TenMon, d.LanThi"

                Using command As New SqlCommand(query, connection)
                    If currentUserRole <> "Admin" Then
                        command.Parameters.AddWithValue("@UserCode", currentUserCode)
                    End If

                    Using adapter As New SqlDataAdapter(command)
                        Dim table As New DataTable()
                        adapter.Fill(table)

                        DataGridView1.DataSource = table

                        If DataGridView1.Columns.Count > 0 Then
                            DataGridView1.Columns("MaSV").HeaderText = "Mã SV"
                            DataGridView1.Columns("HoSV").HeaderText = "Họ"
                            DataGridView1.Columns("TenSV").HeaderText = "Tên"
                            DataGridView1.Columns("MaMon").HeaderText = "Mã Môn"
                            DataGridView1.Columns("TenMon").HeaderText = "Tên Môn Học"
                            DataGridView1.Columns("LanThi").HeaderText = "Lần Thi"
                            DataGridView1.Columns("Diem").HeaderText = "Điểm"
                            DataGridView1.Columns("NgayThi").HeaderText = "Ngày Thi"
                            DataGridView1.Columns("KetQua").HeaderText = "Kết Quả"
                            DataGridView1.Columns("MaLHP").HeaderText = "Mã LHP"

                            DataGridView1.Columns("NgayThi").DefaultCellStyle.Format = "dd/MM/yyyy"
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Lỗi khi tải dữ liệu: " & ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnThem_Click(sender As Object, e As EventArgs) Handles btnThem.Click
        Try
            Dim table As DataTable = CType(DataGridView1.DataSource, DataTable)
            Dim newRow As DataRow = table.NewRow()
            newRow("LanThi") = 1
            newRow("NgayThi") = DateTime.Today
            table.Rows.Add(newRow)
            DataGridView1.CurrentCell = DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0)
        Catch ex As Exception
            MessageBox.Show("Lỗi khi thêm dòng mới: " & ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnXoa_Click(sender As Object, e As EventArgs) Handles btnXoa.Click
        If DataGridView1.CurrentRow IsNot Nothing Then
            Try
                Dim result As DialogResult = MessageBox.Show("Bạn có chắc chắn muốn xóa bản ghi này?", "Xác nhận xóa", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

                If result = DialogResult.Yes Then
                    Dim maSV As String = DataGridView1.CurrentRow.Cells("MaSV").Value.ToString()
                    Dim maMon As String = DataGridView1.CurrentRow.Cells("MaMon").Value.ToString()
                    Dim lanThi As Integer = Convert.ToInt32(DataGridView1.CurrentRow.Cells("LanThi").Value)

                    Dim connectionString As String = "Server=localhost\SQLEXPRESS;Database=QLSinhVien;Integrated Security=True"
                    Using connection As New SqlConnection(connectionString)
                        connection.Open()
                        Dim query As String = "DELETE FROM Diem WHERE MaSV = @MaSV AND MaMon = @MaMon AND LanThi = @LanThi"
                        Using command As New SqlCommand(query, connection)
                            command.Parameters.AddWithValue("@MaSV", maSV)
                            command.Parameters.AddWithValue("@MaMon", maMon)
                            command.Parameters.AddWithValue("@LanThi", lanThi)
                            command.ExecuteNonQuery()
                        End Using
                    End Using

                    DataGridView1.Rows.Remove(DataGridView1.CurrentRow)

                    MessageBox.Show("Xóa thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            Catch ex As Exception
                MessageBox.Show("Lỗi khi xóa: " & ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        Else
            MessageBox.Show("Vui lòng chọn một dòng để xóa", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    Private Sub btnSua_Click(sender As Object, e As EventArgs) Handles btnSua.Click
        If DataGridView1.CurrentRow IsNot Nothing Then
            DataGridView1.BeginEdit(True)
        Else
            MessageBox.Show("Vui lòng chọn một dòng để sửa", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    Private Sub btnLuu_Click(sender As Object, e As EventArgs) Handles btnLuu.Click
        Try
            Dim connectionString As String = "Server=localhost\SQLEXPRESS;Database=QLSinhVien;Integrated Security=True"
            Using connection As New SqlConnection(connectionString)
                connection.Open()

                For Each row As DataGridViewRow In DataGridView1.Rows
                    If Not row.IsNewRow Then
                        Dim maSV As String = row.Cells("MaSV").Value.ToString()
                        Dim maMon As String = row.Cells("MaMon").Value.ToString()
                        Dim lanThi As Integer = Convert.ToInt32(row.Cells("LanThi").Value)
                        Dim diem As Decimal = If(IsDBNull(row.Cells("Diem").Value), 0, Convert.ToDecimal(row.Cells("Diem").Value))
                        Dim ngayThi As Date = If(IsDBNull(row.Cells("NgayThi").Value), DateTime.Today, Convert.ToDateTime(row.Cells("NgayThi").Value))
                        Dim ketQua As String = If(IsDBNull(row.Cells("KetQua").Value), "", row.Cells("KetQua").Value.ToString())
                        Dim maLHP As String = If(IsDBNull(row.Cells("MaLHP").Value), "", row.Cells("MaLHP").Value.ToString())

                        If String.IsNullOrEmpty(ketQua) Then
                            ketQua = If(diem >= 5, "Đạt", "Không đạt")
                        End If

                        Dim checkQuery As String = "SELECT COUNT(*) FROM Diem WHERE MaSV = @MaSV AND MaMon = @MaMon AND LanThi = @LanThi"
                        Using checkCommand As New SqlCommand(checkQuery, connection)
                            checkCommand.Parameters.AddWithValue("@MaSV", maSV)
                            checkCommand.Parameters.AddWithValue("@MaMon", maMon)
                            checkCommand.Parameters.AddWithValue("@LanThi", lanThi)
                            Dim count As Integer = Convert.ToInt32(checkCommand.ExecuteScalar())

                            If count > 0 Then
                                Dim updateQuery As String = "UPDATE Diem SET Diem = @Diem, NgayThi = @NgayThi, KetQua = @KetQua, MaLHP = @MaLHP WHERE MaSV = @MaSV AND MaMon = @MaMon AND LanThi = @LanThi"
                                Using updateCommand As New SqlCommand(updateQuery, connection)
                                    updateCommand.Parameters.AddWithValue("@Diem", diem)
                                    updateCommand.Parameters.AddWithValue("@NgayThi", ngayThi)
                                    updateCommand.Parameters.AddWithValue("@KetQua", ketQua)
                                    updateCommand.Parameters.AddWithValue("@MaLHP", If(String.IsNullOrEmpty(maLHP), DBNull.Value, maLHP))
                                    updateCommand.Parameters.AddWithValue("@MaSV", maSV)
                                    updateCommand.Parameters.AddWithValue("@MaMon", maMon)
                                    updateCommand.Parameters.AddWithValue("@LanThi", lanThi)
                                    updateCommand.ExecuteNonQuery()
                                End Using
                            Else
                                Dim insertQuery As String = "INSERT INTO Diem (MaSV, MaMon, LanThi, Diem, NgayThi, KetQua, MaLHP) VALUES (@MaSV, @MaMon, @LanThi, @Diem, @NgayThi, @KetQua, @MaLHP)"
                                Using insertCommand As New SqlCommand(insertQuery, connection)
                                    insertCommand.Parameters.AddWithValue("@MaSV", maSV)
                                    insertCommand.Parameters.AddWithValue("@MaMon", maMon)
                                    insertCommand.Parameters.AddWithValue("@LanThi", lanThi)
                                    insertCommand.Parameters.AddWithValue("@Diem", diem)
                                    insertCommand.Parameters.AddWithValue("@NgayThi", ngayThi)
                                    insertCommand.Parameters.AddWithValue("@KetQua", ketQua)
                                    insertCommand.Parameters.AddWithValue("@MaLHP", If(String.IsNullOrEmpty(maLHP), DBNull.Value, maLHP))
                                    insertCommand.ExecuteNonQuery()
                                End Using
                            End If
                        End Using
                    End If
                Next

                MessageBox.Show("Lưu dữ liệu thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information)
                LoadData()
            End Using
        Catch ex As Exception
            MessageBox.Show("Lỗi khi lưu dữ liệu: " & ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        If e.RowIndex >= 0 AndAlso e.ColumnIndex = DataGridView1.Columns("Diem").Index Then
            Dim diem As Decimal = If(IsDBNull(DataGridView1.Rows(e.RowIndex).Cells("Diem").Value), 0, Convert.ToDecimal(DataGridView1.Rows(e.RowIndex).Cells("Diem").Value))
            Dim ketQua As String = If(diem >= 5, "Đạt", "Không đạt")

            DataGridView1.Rows(e.RowIndex).Cells("KetQua").Value = ketQua
        End If
    End Sub
End Class