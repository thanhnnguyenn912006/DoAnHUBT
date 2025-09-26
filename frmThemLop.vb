Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Drawing

Public Class frmThemLop
    Inherits Form

    Private connectionString As String = "Server=localhost\SQLEXPRESS;Database=QLSinhVien;Integrated Security=True"

    Private txtMaLop As TextBox
    Private txtTenLop As TextBox
    Private txtMaNganh As TextBox
    Private txtMaKhoa As TextBox
    Private txtKhoaHoc As TextBox
    Private txtSiSo As TextBox
    Private btnLuu As Button
    Private btnHuy As Button

    Public Sub New()

        InitializeCustomComponents()
    End Sub

    Private Sub InitializeCustomComponents()
        Me.Text = "Thêm lớp mới"
        Me.Size = New Size(450, 400)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.BackColor = Color.White

        ' Tạo Panel chứa nội dung
        Dim panel As New Panel()
        panel.Dock = DockStyle.Fill
        panel.Padding = New Padding(20)
        Me.Controls.Add(panel)

        ' Tạo TableLayoutPanel để sắp xếp controls
        Dim tableLayout As New TableLayoutPanel()
        tableLayout.Dock = DockStyle.Fill
        tableLayout.ColumnCount = 2
        tableLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 30))
        tableLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 70))
        tableLayout.RowCount = 7
        tableLayout.Padding = New Padding(5)
        panel.Controls.Add(tableLayout)

        ' Tạo và thêm controls
        AddControlsToTableLayout(tableLayout)

        ' Tạo panel cho buttons
        Dim buttonPanel As New Panel()
        buttonPanel.Dock = DockStyle.Bottom
        buttonPanel.Height = 50
        buttonPanel.Padding = New Padding(10)
        panel.Controls.Add(buttonPanel)

        ' Tạo buttons
        CreateButtons(buttonPanel)
    End Sub

    Private Sub AddControlsToTableLayout(tableLayout As TableLayoutPanel)
        ' Mã lớp
        Dim lblMaLop As New Label()
        lblMaLop.Text = "Mã lớp:"
        lblMaLop.TextAlign = ContentAlignment.MiddleRight
        tableLayout.Controls.Add(lblMaLop, 0, 0)

        txtMaLop = New TextBox()
        txtMaLop.Margin = New Padding(3, 5, 3, 5)
        tableLayout.Controls.Add(txtMaLop, 1, 0)

        ' Tên lớp
        Dim lblTenLop As New Label()
        lblTenLop.Text = "Tên lớp:"
        lblTenLop.TextAlign = ContentAlignment.MiddleRight
        tableLayout.Controls.Add(lblTenLop, 0, 1)

        txtTenLop = New TextBox()
        txtTenLop.Margin = New Padding(3, 5, 3, 5)
        tableLayout.Controls.Add(txtTenLop, 1, 1)

        ' Mã ngành
        Dim lblMaNganh As New Label()
        lblMaNganh.Text = "Mã ngành:"
        lblMaNganh.TextAlign = ContentAlignment.MiddleRight
        tableLayout.Controls.Add(lblMaNganh, 0, 2)

        txtMaNganh = New TextBox()
        txtMaNganh.Margin = New Padding(3, 5, 3, 5)
        tableLayout.Controls.Add(txtMaNganh, 1, 2)

        ' Mã khoa
        Dim lblMaKhoa As New Label()
        lblMaKhoa.Text = "Mã khoa:"
        lblMaKhoa.TextAlign = ContentAlignment.MiddleRight
        tableLayout.Controls.Add(lblMaKhoa, 0, 3)

        txtMaKhoa = New TextBox()
        txtMaKhoa.Margin = New Padding(3, 5, 3, 5)
        tableLayout.Controls.Add(txtMaKhoa, 1, 3)

        ' Khóa học
        Dim lblKhoaHoc As New Label()
        lblKhoaHoc.Text = "Khóa học:"
        lblKhoaHoc.TextAlign = ContentAlignment.MiddleRight
        tableLayout.Controls.Add(lblKhoaHoc, 0, 4)

        txtKhoaHoc = New TextBox()
        txtKhoaHoc.Margin = New Padding(3, 5, 3, 5)
        tableLayout.Controls.Add(txtKhoaHoc, 1, 4)

        ' Sĩ số
        Dim lblSiSo As New Label()
        lblSiSo.Text = "Sĩ số:"
        lblSiSo.TextAlign = ContentAlignment.MiddleRight
        tableLayout.Controls.Add(lblSiSo, 0, 5)

        txtSiSo = New TextBox()
        txtSiSo.Margin = New Padding(3, 5, 3, 5)
        tableLayout.Controls.Add(txtSiSo, 1, 5)
    End Sub

    Private Sub CreateButtons(buttonPanel As Panel)
        ' Button Lưu
        btnLuu = New Button()
        btnLuu.Text = "Lưu"
        btnLuu.Size = New Size(80, 30)
        btnLuu.Location = New Point(buttonPanel.Width - 180, 10)
        btnLuu.BackColor = Color.LightBlue
        AddHandler btnLuu.Click, AddressOf btnLuu_Click
        buttonPanel.Controls.Add(btnLuu)

        ' Button Hủy
        btnHuy = New Button()
        btnHuy.Text = "Hủy"
        btnHuy.Size = New Size(80, 30)
        btnHuy.Location = New Point(buttonPanel.Width - 90, 10)
        btnHuy.BackColor = Color.LightCoral
        AddHandler btnHuy.Click, AddressOf btnHuy_Click
        buttonPanel.Controls.Add(btnHuy)
    End Sub

    Private Sub btnLuu_Click(sender As Object, e As EventArgs)
        If ValidateData() Then
            If InsertLop() Then
                MessageBox.Show("Thêm lớp thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.DialogResult = DialogResult.OK
                Me.Close()
            Else
                MessageBox.Show("Lỗi khi thêm lớp!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
    End Sub

    Private Function ValidateData() As Boolean
        If String.IsNullOrEmpty(txtMaLop.Text) Then
            MessageBox.Show("Vui lòng nhập mã lớp!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtMaLop.Focus()
            Return False
        End If

        If String.IsNullOrEmpty(txtTenLop.Text) Then
            MessageBox.Show("Vui lòng nhập tên lớp!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtTenLop.Focus()
            Return False
        End If

        If Not Integer.TryParse(txtSiSo.Text, Nothing) OrElse Integer.Parse(txtSiSo.Text) <= 0 Then
            MessageBox.Show("Sĩ số phải là số nguyên dương!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtSiSo.Focus()
            Return False
        End If

        Return True
    End Function

    Private Function InsertLop() As Boolean
        Try
            Using connection As New SqlConnection(connectionString)
                connection.Open()
                Dim query As String = "INSERT INTO Lop (MaLop, TenLop, MaNganh, MaKhoa, KhoaHoc, SiSo) 
                                     VALUES (@MaLop, @TenLop, @MaNganh, @MaKhoa, @KhoaHoc, @SiSo)"

                Using command As New SqlCommand(query, connection)
                    command.Parameters.AddWithValue("@MaLop", txtMaLop.Text)
                    command.Parameters.AddWithValue("@TenLop", txtTenLop.Text)
                    command.Parameters.AddWithValue("@MaNganh", If(String.IsNullOrEmpty(txtMaNganh.Text), DBNull.Value, txtMaNganh.Text))
                    command.Parameters.AddWithValue("@MaKhoa", If(String.IsNullOrEmpty(txtMaKhoa.Text), DBNull.Value, txtMaKhoa.Text))
                    command.Parameters.AddWithValue("@KhoaHoc", If(String.IsNullOrEmpty(txtKhoaHoc.Text), DBNull.Value, txtKhoaHoc.Text))
                    command.Parameters.AddWithValue("@SiSo", Integer.Parse(txtSiSo.Text))

                    Return command.ExecuteNonQuery() > 0
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Lỗi: " & ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    Private Sub btnHuy_Click(sender As Object, e As EventArgs)
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
End Class