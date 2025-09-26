Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Drawing

Public Class frmSuaLop
    Inherits Form

    Private connectionString As String = "Server=localhost\SQLEXPRESS;Database=QLSinhVien;Integrated Security=True"
    Private maLop As String


    Private txtMaLop, txtTenLop, txtMaNganh, txtMaKhoa, txtKhoaHoc, txtSiSo As TextBox

    Public Sub New(maLop As String)
        Me.maLop = maLop
        InitializeForm()
        LoadClassData()
    End Sub

    Private Sub InitializeForm()
        Me.Text = "Sửa thông tin lớp"
        Me.Size = New Size(400, 400)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.BackColor = Color.White

        CreateControls()
    End Sub

    Private Sub CreateControls()
        Dim lblMaLop As New Label()
        lblMaLop.Text = "Mã lớp:"
        lblMaLop.Location = New Point(20, 20)
        lblMaLop.Size = New Size(100, 20)

        txtMaLop = New TextBox()
        txtMaLop.Name = "txtMaLop"
        txtMaLop.Location = New Point(120, 20)
        txtMaLop.Size = New Size(200, 20)
        txtMaLop.ReadOnly = True

        ' Tên lớp
        Dim lblTenLop As New Label()
        lblTenLop.Text = "Tên lớp:"
        lblTenLop.Location = New Point(20, 60)
        lblTenLop.Size = New Size(100, 20)

        txtTenLop = New TextBox()
        txtTenLop.Name = "txtTenLop"
        txtTenLop.Location = New Point(120, 60)
        txtTenLop.Size = New Size(200, 20)

        ' Mã ngành
        Dim lblMaNganh As New Label()
        lblMaNganh.Text = "Mã ngành:"
        lblMaNganh.Location = New Point(20, 100)
        lblMaNganh.Size = New Size(100, 20)

        txtMaNganh = New TextBox()
        txtMaNganh.Name = "txtMaNganh"
        txtMaNganh.Location = New Point(120, 100)
        txtMaNganh.Size = New Size(200, 20)

        ' Mã khoa
        Dim lblMaKhoa As New Label()
        lblMaKhoa.Text = "Mã khoa:"
        lblMaKhoa.Location = New Point(20, 140)
        lblMaKhoa.Size = New Size(100, 20)

        txtMaKhoa = New TextBox()
        txtMaKhoa.Name = "txtMaKhoa"
        txtMaKhoa.Location = New Point(120, 140)
        txtMaKhoa.Size = New Size(200, 20)

        ' Khóa học
        Dim lblKhoaHoc As New Label()
        lblKhoaHoc.Text = "Khóa học:"
        lblKhoaHoc.Location = New Point(20, 180)
        lblKhoaHoc.Size = New Size(100, 20)

        txtKhoaHoc = New TextBox()
        txtKhoaHoc.Name = "txtKhoaHoc"
        txtKhoaHoc.Location = New Point(120, 180)
        txtKhoaHoc.Size = New Size(200, 20)

        ' Sĩ số
        Dim lblSiSo As New Label()
        lblSiSo.Text = "Sĩ số:"
        lblSiSo.Location = New Point(20, 220)
        lblSiSo.Size = New Size(100, 20)

        txtSiSo = New TextBox()
        txtSiSo.Name = "txtSiSo"
        txtSiSo.Location = New Point(120, 220)
        txtSiSo.Size = New Size(200, 20)

        ' Nút Lưu
        Dim btnLuu As New Button()
        btnLuu.Text = "Lưu"
        btnLuu.Location = New Point(120, 270)
        btnLuu.Size = New Size(80, 30)
        AddHandler btnLuu.Click, AddressOf btnLuu_Click

        ' Nút Hủy
        Dim btnHuy As New Button()
        btnHuy.Text = "Hủy"
        btnHuy.Location = New Point(220, 270)
        btnHuy.Size = New Size(80, 30)
        AddHandler btnHuy.Click, AddressOf btnHuy_Click

        Me.Controls.Add(lblMaLop)
        Me.Controls.Add(txtMaLop)
        Me.Controls.Add(lblTenLop)
        Me.Controls.Add(txtTenLop)
        Me.Controls.Add(lblMaNganh)
        Me.Controls.Add(txtMaNganh)
        Me.Controls.Add(lblMaKhoa)
        Me.Controls.Add(txtMaKhoa)
        Me.Controls.Add(lblKhoaHoc)
        Me.Controls.Add(txtKhoaHoc)
        Me.Controls.Add(lblSiSo)
        Me.Controls.Add(txtSiSo)
        Me.Controls.Add(btnLuu)
        Me.Controls.Add(btnHuy)
    End Sub

    Private Sub LoadClassData()
        Try
            Using connection As New SqlConnection(connectionString)
                connection.Open()
                Dim query As String = "SELECT MaLop, TenLop, MaNganh, MaKhoa, KhoaHoc, SiSo FROM Lop WHERE MaLop = @MaLop"

                Using command As New SqlCommand(query, connection)
                    command.Parameters.AddWithValue("@MaLop", maLop)

                    Using reader As SqlDataReader = command.ExecuteReader()
                        If reader.Read() Then
                            txtMaLop.Text = reader("MaLop").ToString()
                            txtTenLop.Text = reader("TenLop").ToString()
                            txtMaNganh.Text = reader("MaNganh").ToString()
                            txtMaKhoa.Text = reader("MaKhoa").ToString()
                            txtKhoaHoc.Text = reader("KhoaHoc").ToString()
                            txtSiSo.Text = reader("SiSo").ToString()
                        Else
                            MessageBox.Show("Không tìm thấy lớp với mã: " & maLop, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Me.Close()
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Lỗi khi tải dữ liệu lớp: " & ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Close()
        End Try
    End Sub

    Private Sub btnLuu_Click(sender As Object, e As EventArgs)
        ' Kiểm tra dữ liệu
        If String.IsNullOrEmpty(txtTenLop.Text) Then
            MessageBox.Show("Vui lòng điền tên lớp!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtTenLop.Focus()
            Return
        End If

        If Not Integer.TryParse(txtSiSo.Text, Nothing) OrElse Integer.Parse(txtSiSo.Text) <= 0 Then
            MessageBox.Show("Sĩ số phải là số nguyên dương!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtSiSo.Focus()
            Return
        End If

        ' Lưu dữ liệu
        Try
            Using connection As New SqlConnection(connectionString)
                connection.Open()
                Dim query As String = "UPDATE Lop SET TenLop = @TenLop, MaNganh = @MaNganh, 
                                     MaKhoa = @MaKhoa, KhoaHoc = @KhoaHoc, SiSo = @SiSo 
                                     WHERE MaLop = @MaLop"

                Using command As New SqlCommand(query, connection)
                    command.Parameters.AddWithValue("@MaLop", maLop)
                    command.Parameters.AddWithValue("@TenLop", txtTenLop.Text)
                    command.Parameters.AddWithValue("@MaNganh", If(String.IsNullOrEmpty(txtMaNganh.Text), DBNull.Value, txtMaNganh.Text))
                    command.Parameters.AddWithValue("@MaKhoa", If(String.IsNullOrEmpty(txtMaKhoa.Text), DBNull.Value, txtMaKhoa.Text))
                    command.Parameters.AddWithValue("@KhoaHoc", If(String.IsNullOrEmpty(txtKhoaHoc.Text), DBNull.Value, txtKhoaHoc.Text))
                    command.Parameters.AddWithValue("@SiSo", Integer.Parse(txtSiSo.Text))

                    Dim result As Integer = command.ExecuteNonQuery()
                    If result > 0 Then
                        MessageBox.Show("Cập nhật lớp thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Me.DialogResult = DialogResult.OK
                    Else
                        MessageBox.Show("Không thể cập nhật lớp!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                End Using
            End Using
            Me.Close()

        Catch ex As Exception
            MessageBox.Show("Lỗi khi cập nhật lớp: " & ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnHuy_Click(sender As Object, e As EventArgs)
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
End Class