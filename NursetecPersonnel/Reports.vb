Imports System.Data.SqlClient
Imports System.Runtime.InteropServices

Public Class Reports
    'Class wide declarations
    Dim saved As Integer = 0
    Dim atOfficeID As SqlParameter = New SqlParameter("@OfficeID", SqlDbType.Int)
    Dim formatRange
    Dim border As Microsoft.Office.Interop.Excel.Borders
    Dim xlApp As Microsoft.Office.Interop.Excel.Application
    Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Dim xlWorkSheet As Microsoft.Office.Interop.Excel.Worksheet
    Dim misValue As Object = System.Reflection.Missing.Value
    Dim suggestions As List(Of AutoCompleteStringCollection) = New List(Of AutoCompleteStringCollection)
    Dim conn As SqlConnection = New SqlConnection("Integrated Security=SSPI;Persist Security Info=False;User ID=dba;Initial Catalog=Nursetec;Data Source=NURLAPTOP26\SQLEXPRESS01")
    Dim cmd As SqlCommand = New SqlCommand
    Dim adap As SqlDataAdapter = New SqlDataAdapter
    Dim adap2 As SqlDataAdapter = New SqlDataAdapter
    Dim ds As DataSet = New DataSet
    Dim ds3 As DataSet = New DataSet
    Dim empID As Integer = 0
    Dim atEmpID As SqlParameter = New SqlParameter("@Employee_ID", SqlDbType.Int)
    Dim atStartDate As SqlParameter = New SqlParameter("@startDate", SqlDbType.Date)
    Dim atEndDate As SqlParameter = New SqlParameter("@endDate", SqlDbType.Date)
    Dim ds2 As DataSet = New DataSet
    Dim dt As DataTable = New DataTable
    Dim dt2 As DataTable = New DataTable
    Dim dv As DataView = New DataView
    Dim str As String = ""
    Dim friday = 0
    Dim PublicHoliday = 0
    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        dgvPresent.DataSource = Nothing
        saved = 0
        lblProg.Text = "Fetching data..."
        Try
            If cmbBreakReport.SelectedItem Is Nothing Then
                MessageBox.Show("Please select the type of report")
            Else
                If dtpStart.Value <= dtpEnd.Value Then
                    Me.Cursor = Cursors.WaitCursor
                    conn.Open()
                    ds2.Clear()
                    cmd.Parameters.Clear()
                    cmd.Connection = conn
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.CommandTimeout = 15
                    cmd.CommandText = "sp_FetchAllReports"
                    If txtEmpID.Text = "" Then
                        empID = Nothing
                    Else
                        empID = txtEmpID.Text
                    End If
                    atEmpID.Value = empID
                    cmd.Parameters.Add(atEmpID)
                    cmd.Parameters.Add(atStartDate)
                    cmd.Parameters.Add(atEndDate)
                    cmd.Parameters.Add(atOfficeID)
                    adap.SelectCommand = cmd
                    adap.Fill(ds2)
                    'Populating the DGV based on report type
                    If cmbBreakReport.SelectedItem.ToString = "Full Report" Then
                        dgvPresent.Visible = False
                        dgvDataBreaks.Visible = True
                        dt2 = ds2.Tables(3)
                        dgvDataBreaks.DataSource = dt2
                        'Conditionaly formatting rows in DGV
                        For i = 0 To dgvDataBreaks.RowCount - 2
                            For j = 0 To dgvDataBreaks.ColumnCount - 1
                                Dim index As Integer = dgvDataBreaks(j, i).ColumnIndex
                                If index = 1 Then
                                    Dim d As Date = FormatDateTime(dgvDataBreaks(j, i).Value, DateFormat.LongDate)
                                    FormatDateTime(dgvDataBreaks(j, i).Value, DateFormat.LongDate).ToString()
                                    Dim cmdPub As SqlCommand = New SqlCommand
                                    cmdPub.Connection = conn
                                    cmdPub.CommandType = CommandType.Text
                                    Dim date2 = CDate(d.AddDays(1)).ToString("yyyy-MM-dd")
                                    cmdPub.CommandText = "SELECT * FROM Public_Holidays WHERE DayOfHoliday='" & date2 & "'"
                                    If d.DayOfWeek = DayOfWeek.Friday Then
                                        friday = 1
                                    ElseIf cmdPub.ExecuteScalar > 0 Then
                                        friday = 1
                                    End If
                                Else
                                    dgvDataBreaks(j, i).Value.ToString()
                                End If
                                'Pretoria CBD office hours
                                If cmbOffice.SelectedValue = 6 Then
                                    If index = 2 Then
                                        Dim t = dgvDataBreaks(j, i).Value.ToString()
                                        Convert.ToDateTime(t)
                                        If t > #7:02 AM# Then
                                            dgvDataBreaks(j, i).Style.BackColor = Color.Yellow ''Late arrival
                                        End If
                                    End If
                                    If index = 3 Then
                                        Dim t2 = dgvDataBreaks(j, i).Value.ToString()
                                        Convert.ToDateTime(t2)
                                        If friday = 0 And t2 < #3:58 PM# Then                  ''Early Leaving
                                            dgvDataBreaks(j, i).Style.BackColor = Color.Yellow
                                        ElseIf t2 < #2:58 PM# Then
                                            dgvDataBreaks(j, i).Style.BackColor = Color.Yellow
                                        End If
                                    End If
                                Else
                                    'All other office's hours
                                    If index = 2 Then
                                        Dim t = dgvDataBreaks(j, i).Value.ToString()
                                        Convert.ToDateTime(t)
                                        If t > #7:32 AM# Then                                  ''Late Arrival   
                                            dgvDataBreaks(j, i).Style.BackColor = Color.Yellow
                                        End If
                                    End If
                                    If index = 3 Then
                                        Dim t2 = dgvDataBreaks(j, i).Value.ToString()
                                        Convert.ToDateTime(t2)
                                        If friday = 0 And t2 < #4:28 PM# Then                  ''Early Leaving
                                            dgvDataBreaks(j, i).Style.BackColor = Color.Yellow
                                        ElseIf t2 < #3:28 PM# Then
                                            dgvDataBreaks(j, i).Style.BackColor = Color.Yellow
                                        End If
                                    End If
                                End If
                                If index = 5 Then
                                    If friday = 1 Then
                                        If dgvDataBreaks(j, i).Value < 7 Then
                                            Dim row As DataGridViewRow = dgvDataBreaks.Rows(dgvDataBreaks(j, i).RowIndex)
                                            row.DefaultCellStyle.BackColor = Color.LightSalmon ''Less working hours on friday
                                        End If
                                    Else
                                        If dgvDataBreaks(j, i).Value < 8 Then
                                            Dim row As DataGridViewRow = dgvDataBreaks.Rows(dgvDataBreaks(j, i).RowIndex)
                                            row.DefaultCellStyle.BackColor = Color.LightSalmon  ''Less working hours every-other day
                                        End If
                                    End If
                                End If
                                If index = 4 Then
                                    If dgvDataBreaks(j, i).Value > 13 Then
                                        Dim row As DataGridViewRow = dgvDataBreaks.Rows(dgvDataBreaks(j, i).RowIndex)
                                        row.DefaultCellStyle.BackColor = Color.LightGoldenrodYellow ''missed clockout
                                    End If
                                End If
                            Next
                            friday = 0
                        Next
                    ElseIf cmbBreakReport.SelectedItem.ToString = "Absent" Then
                        conn.Close()
                        conn.Open()
                        ds2.Clear()
                        cmd.Parameters.Clear()
                        cmd.Connection = conn
                        cmd.CommandType = CommandType.StoredProcedure
                        cmd.CommandTimeout = 15
                        cmd.CommandText = "sp_FetchAllReports"
                        If txtEmpID.Text = "" Then
                            empID = Nothing
                        Else
                            empID = txtEmpID.Text
                        End If
                        atEmpID.Value = empID
                        cmd.Parameters.Add(atEmpID)
                        cmd.Parameters.Add(atStartDate)
                        cmd.Parameters.Add(atEndDate)
                        cmd.Parameters.Add(atOfficeID)
                        adap.SelectCommand = cmd
                        adap.Fill(ds2)
                        Dim absent As DataTable = ds2.Tables(4)
                        Dim Emp_id As List(Of Integer) = New List(Of Integer)
                        Dim i = 0
                        Dim days = 0
                        Dim stringdate As List(Of String)
                        Dim Emp_Nickname = ""
                        Dim Emp_Surname = ""
                        Dim da As SqlDataAdapter = New SqlDataAdapter
                        Dim Dset As DataSet = New DataSet
                        stringdate = DaysofWeek(atStartDate.Value, atEndDate.Value)
                        Emp_id = absent.Rows.Cast(Of DataRow).Select(Function(dr) Integer.Parse(dr(0))).ToList
                        dgvDataBreaks.Visible = False
                        dgvPresent.Visible = True
                        Dim header = "Employee"
                        Dim dt = New DataTable
                        dt.Columns.Add(header)
                        For x = 0 To stringdate.Count - 1
                            Dim col = stringdate(x)
                            dt.Columns.Add(col)
                        Next
                        For y = 0 To Emp_id.Count - 1
                            Dset.Clear()
                            cmd.CommandType = CommandType.Text
                            cmd.CommandText = "Select CONCAT(Surname,', ',NickName) As NickName FROM Employee WHERE Employee_ID ='" & Emp_id(y) & "'"
                            da.SelectCommand = cmd
                            da.Fill(Dset)
                            Emp_Nickname = Dset.Tables(0).Rows(0)("NickName")
                            Dim rowID = Emp_Nickname
                            dt.Rows.Add(rowID)
                        Next
                        For x = 0 To stringdate.Count - 1
                            For z = 0 To Emp_id.Count - 1
                                Dim row As DataRow = dt(z)
                                cmd.CommandType = CommandType.Text
                                cmd.CommandText = "SELECT * FROM Present WHERE Employee_ID='" & Emp_id(z) & "' AND ClockDate='" & stringdate(x) & "'"
                                If cmd.ExecuteScalar > 0 Then
                                    Row.Item(x + 1) = "X"
                                Else
                                    row.Item(x + 1) = " "
                                End If
                            Next
                        Next
                        dgvPresent.DataSource = dt
                        For c = 0 To dgvPresent.Columns.Count - 1
                            For v = 0 To dgvPresent.Rows.Count - 1
                                If dgvPresent(c, v).Value = " " Then 'Formatting rows that show absenteeism
                                    dgvPresent(c, v).Style.BackColor = Color.Red

                                End If
                            Next
                        Next
                        conn.Close()
                        dgvPresent.Columns("Employee").Frozen = True
                    End If
                    conn.Close()
                    dgvDataBreaks.AutoResizeColumns()
                    Me.Cursor = Cursors.Default
                Else
                    MessageBox.Show("End date cannot be earlier than start date")
                End If
            End If
            dv = New DataView(dt2)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            conn.Close()
            Me.Cursor = Cursors.Default
        End Try
        lblProg.Text = "Ready"
    End Sub

    Private Sub Breaks_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dgvDataBreaks.Visible = False
        dgvPresent.Visible = False
        'Populating the combo box
        rdbtnExcel.Checked = True
        lblProg.Text = "Ready"
        cmbEmployee.AutoCompleteMode = AutoCompleteMode.Append
        cmbEmployee.AutoCompleteSource = AutoCompleteSource.ListItems
        conn.Open()
        dgvDataBreaks.AutoGenerateColumns = False
        cmd.Connection = conn
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "SELECT Employee_ID, NickName FROM Employee"
        adap.SelectCommand = cmd
        adap.Fill(ds)
        cmbEmployee.DataSource = ds.Tables(0)
        cmbEmployee.DisplayMember = "NickName"
        cmbEmployee.ValueMember = "Employee_ID"
        cmbEmployee.SelectedItem = Nothing
        cmd.CommandText = "SELECT NickName FROM Employee"
        adap2.SelectCommand = cmd
        adap2.Fill(ds3)
        cmbBreakReport.Items.Add("Full Report")
        cmbBreakReport.Items.Add("Absent")
        txtEmpID.Enabled = True
        Me.CenterToScreen()
        Me.CenterToParent()
        atEndDate.Value = CType(Nothing, DateTime?)
        atStartDate.Value = CType(Nothing, DateTime?)
        conn.Close()
        atEmpID.Value = Nothing
        cmbEmployee.AutoCompleteMode = AutoCompleteMode.Suggest
        cmbEmployee.AutoCompleteSource = AutoCompleteSource.ListItems
        Dim ds4 As DataSet = New DataSet
        Dim adap3 As SqlDataAdapter = New SqlDataAdapter
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "SELECT * FROM Admin_Office"
        adap3.SelectCommand = cmd
        adap3.Fill(ds4)
        cmbOffice.DataSource = ds4.Tables(0)
        cmbOffice.DisplayMember = "Office_Description"
        cmbOffice.ValueMember = "Office_ID"
        cmbOffice.SelectedItem = Nothing
    End Sub



    Private Sub dtpStart_ValueChanged(sender As Object, e As EventArgs) Handles dtpStart.ValueChanged
        atStartDate.Value = CDate(dtpStart.Value).ToString("yyyy-MM-dd")
    End Sub

    Private Sub dtpEnd_ValueChanged(sender As Object, e As EventArgs) Handles dtpEnd.ValueChanged
        atEndDate.Value = CDate(dtpEnd.Value).ToString("yyyy-MM-dd")
    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        'Reset of input controls and DGV
        saved = 0
        cmbEmployee.DataSource = ds.Tables(0)
        cmbBreakReport.SelectedItem = Nothing
        txtEmpID.Clear()
        dtpStart.Value = Date.Now
        dtpEnd.Value = Date.Now
        ds2.Clear()
        dgvDataBreaks.DataSource = ds
        cmbEmployee.SelectedItem = Nothing
        txtEmpID.Enabled = True
        cmbOffice.SelectedItem = Nothing
        atOfficeID.Value = Nothing
    End Sub
    Private Sub cmbOffice_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbOffice.SelectedIndexChanged
        If cmbOffice.SelectedValue Is Nothing Then
        Else
            If cmbOffice.SelectedValue.ToString = "System.Data.DataRowView" Then
            Else
                atOfficeID.Value = cmbOffice.SelectedValue.ToString
                Dim dsOff As DataSet = New DataSet
                Dim cmd2 As SqlCommand = New SqlCommand
                cmd2.Connection = conn
                cmd2.CommandType = CommandType.Text
                cmd2.CommandText = "SELECT Employee_ID, NickName FROM Employee WHERE OfficeID = '" & cmbOffice.SelectedValue.ToString & "'"
                adap.SelectCommand = cmd2
                adap.Fill(dsOff)
                cmbEmployee.DataSource = dsOff.Tables(0)
                cmbEmployee.DisplayMember = "NickName"
                cmbEmployee.SelectedItem = Nothing
            End If
        End If
    End Sub
    Private Sub cmbEmployee_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbEmployee.SelectedIndexChanged
        If cmbEmployee.SelectedValue Is Nothing Then
            txtEmpID.Text = Nothing
            atEmpID.Value = Nothing
        Else
            txtEmpID.Enabled = False
            txtEmpID.Text = cmbEmployee.SelectedValue.ToString
            atEmpID.Value = cmbEmployee.SelectedValue.ToString
        End If
    End Sub

    Private Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        saved = 1

        If cmbBreakReport.SelectedItem Is Nothing Then
            MessageBox.Show("Please select the type of report")
        Else
            If dtpStart.Value <= dtpEnd.Value Then
                If cmbBreakReport.SelectedItem.ToString = "Absent" Then 'Exporting Absent report
                    Try
                        tspgbar.Visible = True
                        Dim items As Integer = dgvPresent.Rows.Count - 1
                        Dim prog As Integer = 0
                        Dim i As Integer
                        Dim j As Integer
                        Dim index = 0
                        tspgbar.Maximum = items
                        xlApp = New Microsoft.Office.Interop.Excel.Application
                        xlWorkBook = xlApp.Workbooks.Add(misValue)
                        xlWorkSheet = xlWorkBook.Sheets("sheet1")
                        'Row and column counters
                        Dim r, c As Integer
                        r = 2
                        'Naming the report and adding a Heading
                        xlWorkSheet.Range("A1:" & alpha(dgvPresent.Columns.Count) & "1").Merge()
                        conn.Open()
                        Dim currOff As String
                        cmd.CommandType = CommandType.Text
                        cmd.CommandText = "SELECT Office_Description FROM ADMIN_OFFICE WHERE Office_ID ='" & cmbOffice.SelectedValue & "'"
                        currOff = cmd.ExecuteScalar
                        conn.Close()
                        ''Excel version of the DATA GRID VIEW Formatting
                        With xlWorkSheet.Range("A1")
                            If currOff IsNot "" Then
                                .Value = "Report for " & currOff & " " & FormatDateTime(atStartDate.Value, DateFormat.LongDate) & "-" & FormatDateTime(atEndDate.Value, DateFormat.LongDate)

                            Else
                                .Value = "Report for " & FormatDateTime(atStartDate.Value, DateFormat.LongDate) & "-" & FormatDateTime(atEndDate.Value, DateFormat.LongDate)
                            End If
                            .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                            .Font.Size = 12
                            .Font.Bold = True
                            .RowHeight = 30
                        End With

                        With xlWorkSheet.Application.ActiveWindow
                            .SplitColumn = 1
                            .SplitRow = 0
                        End With
                        xlWorkSheet.Application.ActiveWindow.FreezePanes = True
                        'Exporting headers
                        For k As Integer = 1 To dgvPresent.Columns.Count
                            xlWorkSheet.Cells(2, k) = dgvPresent.Columns(k - 1).HeaderText
                            xlWorkSheet.Cells(2, k).Interior.Color = RGB(242, 242, 242)
                            xlWorkSheet.Cells(2, k).EntireRow.Font.Bold = True
                        Next
                        Dim rows = dgvPresent.Rows.Count + 1
                        For i = 0 To dgvPresent.RowCount - 2
                            For j = 0 To dgvPresent.ColumnCount - 1
                                If dgvPresent(j, i).Value.ToString = " " Then 'Formatting rows that show absenteeism
                                    xlWorkSheet.Cells(i + 3, j + 1) = dgvPresent(j, i).Value.ToString
                                    xlWorkSheet.Cells(i + 3, j + 1).Interior.Color = Color.Red
                                Else
                                    xlWorkSheet.Cells(i + 3, j + 1) = dgvPresent(j, i).Value.ToString
                                End If
                            Next
                            prog += 1
                            tspgbar.Value = prog
                            lblProg.Text = "Completed " & prog & " of " & items & " items"
                        Next
                        'Formating the sheet
                        'Border
                        formatRange = xlWorkSheet.Cells.Range("A1", alpha(dgvPresent.ColumnCount) & dgvPresent.Rows.Count + 1)
                        border = formatRange.Borders
                        xlWorkSheet.Rows.EntireColumn.AutoFit()
                        formatRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                        formatRange.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                        'Print page setup
                        With xlWorkSheet.PageSetup
                            .Zoom = False
                            .FitToPagesWide = 1
                            .FitToPagesTall = False
                            .Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape
                            .TopMargin = xlWorkSheet.Application.CentimetersToPoints(1.91)
                            .BottomMargin = xlWorkSheet.Application.CentimetersToPoints(1.91)
                            .LeftMargin = xlWorkSheet.Application.CentimetersToPoints(0.64)
                            .RightMargin = xlWorkSheet.Application.CentimetersToPoints(0.64)
                            .FooterMargin = xlWorkSheet.Application.CentimetersToPoints(0.76)
                            .HeaderMargin = xlWorkSheet.Application.CentimetersToPoints(0.76)
                        End With
                        xlApp.PrintCommunication = False
                        Dim path As String
                        If currOff IsNot "" Then
                            path = "Absentee Report for " & currOff & " " & FormatDateTime(atStartDate.Value, DateFormat.LongDate) & "-" & FormatDateTime(atEndDate.Value, DateFormat.LongDate)
                        Else
                            path = "Absentee Report for All Offices" & FormatDateTime(atStartDate.Value, DateFormat.LongDate) & "-" & FormatDateTime(atEndDate.Value, DateFormat.LongDate)
                        End If
                        Dim direct As String = "C:\ReportDB\"
                        Dim file As String = (direct + path).Replace(" ", "_")
                        'Saving the file
                        If rdbtnExcel.Checked = True Then
                            lblProg.Text = "Exporting to Excel"
                            xlWorkSheet.SaveAs(file, FileFormat:=Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLStrictWorkbook)
                            Dim result = MessageBox.Show("Would you like to view the file?", "File Saved!", MessageBoxButtons.YesNo)
                            If result = DialogResult.Yes Then
                                lblProg.Text = "Opening " + file
                                Process.Start("Excel.exe", file)
                            Else
                                MsgBox("You can find the file at: " + file)
                            End If
                        ElseIf rdbtnPDF.Checked = True Then
                            lblProg.Text = "Exporting as PDF"
                            xlWorkSheet.ExportAsFixedFormat(Type:=Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF,
                            Filename:=file,
                            Quality:=Microsoft.Office.Interop.Excel.XlFixedFormatQuality.xlQualityStandard)
                            Dim result = MessageBox.Show("Would you like to view the file?", "File Saved!", MessageBoxButtons.YesNo)
                            If result = DialogResult.Yes Then
                                lblProg.Text = "Opening " + file
                                Process.Start(file + ".pdf")
                            Else
                                MsgBox("You can find the file at: " + file)
                            End If
                        ElseIf rdbtnPrinter.Checked = True Then
                            lblProg.Text = "Printing report"
                            xlWorkSheet.PrintOutEx()
                            xlApp.ScreenUpdating = True
                        End If
                        'Closing excel And releasing objects
                        xlWorkBook.Close(False)
                        xlApp.Quit()
                        ReleaseObject(formatRange)
                        ReleaseObject(xlWorkSheet)
                        ReleaseObject(xlWorkBook)
                        ReleaseObject(xlApp)
                        'Reset of labels and progress bars
                        tspgbar.Value = 0
                        items = 0
                        prog = 0
                        lblProg.Text = "Ready"
                        tspgbar.Visible = False
                        GC.Collect()
                        GC.WaitForPendingFinalizers()
                    Catch exc As COMException
                        'Catching export errors
                        MessageBox.Show(exc.Message)
                        tspgbar.Value = 0
                        lblProg.Text = ""
                        tspgbar.Visible = False
                    End Try
                Else
                    If ds2.Tables.Count = 0 Then
                        MessageBox.Show("Please select data before exporting")
                    Else
                        Try
                            'Export declarations
                            tspgbar.Visible = True
                            Dim items As Integer = dgvDataBreaks.Rows.Count - 1
                            Dim prog As Integer = 0
                            tspgbar.Maximum = items
                            Dim i As Integer
                            Dim j As Integer
                            xlApp = New Microsoft.Office.Interop.Excel.Application
                            xlWorkBook = xlApp.Workbooks.Add(misValue)
                            xlWorkSheet = xlWorkBook.Sheets("sheet1")
                            'Row and column counters
                            Dim r, c As Integer
                            r = 2
                            'Naming the report and adding a Heading
                            xlWorkSheet.Range("A1:I1").Merge()
                            conn.Open()
                            Dim currOff As String
                            cmd.CommandType = CommandType.Text
                            cmd.CommandText = "SELECT Office_Description FROM ADMIN_OFFICE WHERE Office_ID ='" & cmbOffice.SelectedValue & "'"
                            currOff = cmd.ExecuteScalar
                            conn.Close()
                            ''Excel version of the DATA GRID VIEW Formatting
                            With xlWorkSheet.Range("A1")
                                If currOff IsNot "" Then
                                    .Value = "Report for " & currOff & " " & FormatDateTime(atStartDate.Value, DateFormat.LongDate) & "-" & FormatDateTime(atEndDate.Value, DateFormat.LongDate)

                                Else
                                    .Value = "Report for " & FormatDateTime(atStartDate.Value, DateFormat.LongDate) & "-" & FormatDateTime(atEndDate.Value, DateFormat.LongDate)
                                End If
                                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                                .Font.Size = 12
                                .Font.Bold = True
                                .RowHeight = 30
                            End With
                            'Exporting headers
                            For k As Integer = 1 To dgvDataBreaks.Columns.Count
                                xlWorkSheet.Cells(2, k) = dgvDataBreaks.Columns(k - 1).HeaderText
                                xlWorkSheet.Cells(2, k).Interior.Color = RGB(242, 242, 242)
                                xlWorkSheet.Cells(2, k).EntireRow.Font.Bold = True
                            Next
                            Dim rows = dgvDataBreaks.Rows.Count + 1
                            formatRange = xlWorkSheet.Range("F" & 3, "F" & rows)
                            With formatRange
                                .Interior.Color = RGB(242, 242, 242)
                                .Font.Bold = True
                                .NumberFormat = "General"
                            End With
                            xlWorkSheet.Range("E" & 3, "E" & rows).NumberFormat = "0,0" 'Number formatting
                            'Exporting data and formatting conditionally.
                            For i = 0 To dgvDataBreaks.RowCount - 2
                                For j = 0 To dgvDataBreaks.ColumnCount - 1
                                    Dim index As Integer = dgvDataBreaks(j, i).ColumnIndex
                                    If index = 1 Then
                                        conn.Close()
                                        conn.Open()
                                        Dim d As DateTime = FormatDateTime(dgvDataBreaks(j, i).Value, DateFormat.LongDate)
                                        xlWorkSheet.Cells(i + 3, j + 1) = FormatDateTime(dgvDataBreaks(j, i).Value, DateFormat.LongDate).ToString
                                        Dim cmdPub As SqlCommand = New SqlCommand
                                        cmdPub.Connection = conn
                                        cmdPub.CommandType = CommandType.Text
                                        Dim date2 = CDate(d.AddDays(1)).ToString("yyyy-MM-dd")
                                        cmdPub.CommandText = "SELECT * FROM Public_Holidays WHERE DayOfHoliday='" & date2 & "'"
                                        If d.DayOfWeek = DayOfWeek.Friday Then
                                            friday = 1
                                        ElseIf cmdPub.ExecuteScalar > 0 Then
                                            friday = 1
                                        Else
                                            friday = 0
                                        End If
                                    Else
                                        xlWorkSheet.Cells(i + 3, j + 1) = dgvDataBreaks(j, i).Value.ToString
                                    End If
                                    conn.Close()
                                    ''Pretoria CBD office hours
                                    If cmbOffice.SelectedValue = 6 Then
                                        If index = 2 Then
                                            Dim t = dgvDataBreaks(j, i).Value.ToString()
                                            Convert.ToDateTime(t)
                                            If t > #7:02  AM# Then
                                                xlWorkSheet.Cells(i + 3, j + 1).Interior.Color = Color.Yellow ''Late arrival
                                            End If
                                        End If
                                        If index = 3 Then
                                            Dim t2 = dgvDataBreaks(j, i).Value.ToString()
                                            Convert.ToDateTime(t2)
                                            If friday = 0 And t2 < #3:58 PM# Then                  ''Early Leaving
                                                xlWorkSheet.Cells(i + 3, j + 1).Interior.Color = Color.Yellow
                                            ElseIf t2 < #2:58 PM# Then
                                                xlWorkSheet.Cells(i + 3, j + 1).Interior.Color = Color.Yellow
                                            End If
                                        End If
                                        If index = 4 Then
                                            If dgvDataBreaks(j, i).Value > 13 Then
                                                xlWorkSheet.Cells.Range("A" & i + 3, "I" & i + 3).Interior.Color = Color.LightGoldenrodYellow
                                            End If
                                        End If
                                    Else
                                        ''All other office's hours
                                        If index = 2 Then
                                            Dim t = dgvDataBreaks(j, i).Value.ToString()
                                            Convert.ToDateTime(t)
                                            If t > #7:32 AM# Then
                                                xlWorkSheet.Cells(i + 3, j + 1).Interior.Color = Color.Yellow
                                            End If
                                        End If
                                        If index = 3 Then
                                            Dim t2 = dgvDataBreaks(j, i).Value.ToString()
                                            Convert.ToDateTime(t2)
                                            If friday = 0 And t2 < #4:28 PM# Then
                                                xlWorkSheet.Cells(i + 3, j + 1).Interior.Color = Color.Yellow
                                            ElseIf t2 < #3:28 PM# Then
                                                xlWorkSheet.Cells(i + 3, j + 1).Interior.Color = Color.Yellow
                                            End If
                                        End If
                                    End If
                                    If index = 5 Then
                                        If friday = 1 Then
                                            If dgvDataBreaks(j, i).Value < 7 Then
                                                xlWorkSheet.Cells.Range("A" & i + 3, "I" & i + 3).Interior.Color = Color.LightSalmon
                                            End If
                                        Else
                                            If dgvDataBreaks(j, i).Value < 8 Then
                                                xlWorkSheet.Cells.Range("A" & i + 3, "I" & i + 3).Interior.Color = Color.LightSalmon
                                            End If
                                        End If
                                    End If
                                    If index = 4 Then
                                        If dgvDataBreaks(j, i).Value > 13 Then
                                            xlWorkSheet.Cells.Range("A" & i + 3, "I" & i + 3).Interior.Color = Color.LightGoldenrodYellow
                                        End If
                                    End If
                                    c += 1
                                    r += 1
                                Next
                                'Incrimenting progress bar
                                prog = prog + 1
                                tspgbar.Value = prog
                                lblProg.Text = "Completed " & prog & " of " & items & " items"
                                friday = 0
                            Next
                            'Formating the sheet
                            'Border
                            formatRange = xlWorkSheet.Range("A2", "I" & rows)
                            border = formatRange.Borders
                            xlWorkSheet.Rows.EntireColumn.AutoFit()
                            formatRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                            formatRange.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                            'Date Column width
                            With xlWorkSheet.Range("B2", "B" & rows)
                                .ColumnWidth = 20
                            End With
                            'Time width and removing seconds
                            xlWorkSheet.Range("C2" & rows).ColumnWidth = 15
                            With xlWorkSheet.Range("D2", "D" & rows)
                                .ColumnWidth = 10
                                .NumberFormat = "hh:mm"
                            End With
                            'Time width and removing seconds
                            With xlWorkSheet.Range("C2", "C" & rows)
                                .ColumnWidth = 10
                                .NumberFormat = "hh:mm"
                            End With
                            'Wrapping text for the Notes                                          
                            xlWorkSheet.Range("A1").Style.WrapText = True
                            xlWorkSheet.Range("I2", "I" & rows).Style.WrapText = True
                            'Print page setup
                            With xlWorkSheet.PageSetup
                                .Zoom = False
                                .FitToPagesWide = 1
                                .FitToPagesTall = False
                                .Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape
                                .TopMargin = xlWorkSheet.Application.CentimetersToPoints(1.91)
                                .BottomMargin = xlWorkSheet.Application.CentimetersToPoints(1.91)
                                .LeftMargin = xlWorkSheet.Application.CentimetersToPoints(0.64)
                                .RightMargin = xlWorkSheet.Application.CentimetersToPoints(0.64)
                                .FooterMargin = xlWorkSheet.Application.CentimetersToPoints(0.76)
                                .HeaderMargin = xlWorkSheet.Application.CentimetersToPoints(0.76)
                            End With
                            xlApp.PrintCommunication = False
                            Dim path As String
                            If currOff IsNot "" Then
                                path = "Report for " & currOff & " " & FormatDateTime(atStartDate.Value, DateFormat.LongDate) & "-" & FormatDateTime(atEndDate.Value, DateFormat.LongDate)
                            Else
                                path = "Report for " & FormatDateTime(atStartDate.Value, DateFormat.LongDate) & "-" & FormatDateTime(atEndDate.Value, DateFormat.LongDate)
                            End If
                            Dim direct As String = "C:\ReportDB\"
                            Dim file As String = (direct + path).Replace(" ", "_")
                            'Saving the file
                            If rdbtnExcel.Checked = True Then
                                lblProg.Text = "Exporting to Excel"
                                xlWorkSheet.SaveAs(file, FileFormat:=Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLStrictWorkbook)
                                Dim result = MessageBox.Show("Would you like to view the file?", "File Saved!", MessageBoxButtons.YesNo)
                                If result = DialogResult.Yes Then
                                    lblProg.Text = "Opening " + file
                                    Process.Start("Excel.exe", file)
                                Else
                                    MsgBox("You can find the file at: " + file)
                                End If
                            ElseIf rdbtnPDF.Checked = True Then
                                lblProg.Text = "Exporting as PDF"
                                xlWorkSheet.ExportAsFixedFormat(Type:=Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF,
                                    Filename:=file,
                                    Quality:=Microsoft.Office.Interop.Excel.XlFixedFormatQuality.xlQualityStandard)
                                Dim result = MessageBox.Show("Would you like to view the file?", "File Saved!", MessageBoxButtons.YesNo)
                                If result = DialogResult.Yes Then
                                    lblProg.Text = "Opening " + file
                                    Process.Start(file + ".pdf")
                                Else
                                    MsgBox("You can find the file at: " + file)
                                End If
                            ElseIf rdbtnPrinter.Checked = True Then
                                lblProg.Text = "Printing report"
                                xlWorkSheet.PrintOutEx()
                                xlApp.ScreenUpdating = True
                            End If
                            'Closing excel And releasing objects
                            xlWorkBook.Close(False)
                            xlApp.Quit()
                            'xlApp.Dispose()
                            ReleaseObject(formatRange)
                            ReleaseObject(xlWorkSheet)
                            ReleaseObject(xlWorkBook)
                            ReleaseObject(xlApp)
                            'Reset of labels and progress bars
                            tspgbar.Value = 0
                            items = 0
                            prog = 0
                            lblProg.Text = "Ready"
                            tspgbar.Visible = False
                            GC.Collect()
                            GC.WaitForPendingFinalizers()

                        Catch exc As COMException
                            'Catching export errors
                            MessageBox.Show(exc.Message)
                            tspgbar.Value = 0
                            lblProg.Text = ""
                            tspgbar.Visible = False
                        End Try
                    End If
                End If
            Else
                MessageBox.Show("End date cannot be earlier than start date")
            End If
        End If
    End Sub
    'Releasing objects and collecting garbage
    Private Sub ReleaseObject(ByVal obj As Object)
        Try
            Dim intRel As Integer = 0
            Do
                intRel = System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            Loop While intRel > 0
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub Breaks_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        'Exit button
        If saved = 1 Then
            Dim m As Main = Main
            m.lblMainTask.Text = "Open Task: None"
        Else
            Dim ans = MessageBox.Show("Exit without printing?", "Exit", MessageBoxButtons.YesNo)
            If ans = DialogResult.Yes Then
                Dim m As Main = Main
                m.lblMainTask.Text = "Open Task: None"
            Else
                e.Cancel = True
            End If
        End If
    End Sub
    Public Shared Function Weekdays(ByRef startDate As Date, ByRef endDate As Date) As Integer
        Dim numWeekdays As Integer
        Dim totalDays As Integer
        Dim WeekendDays As Integer
        numWeekdays = 0
        WeekendDays = 0

        totalDays = DateDiff(DateInterval.Day, startDate, endDate) + 1

        For i As Integer = 1 To totalDays

            If DatePart(DateInterval.Weekday, startDate) = 1 Then
                WeekendDays = WeekendDays + 1
            End If
            If DatePart(DateInterval.Weekday, startDate) = 7 Then
                WeekendDays = WeekendDays + 1
            End If
            startDate = DateAdd("d", 1, startDate)
        Next
        numWeekdays = totalDays - WeekendDays

        Return numWeekdays
    End Function
    Public Shared Function DaysofWeek(ByRef startDate As Date, ByRef endDate As Date) As List(Of String)
        Reports.conn.Close()
        Reports.conn.Open()
        Dim days As New List(Of String)
        Dim currDay As New Date
        Dim start = CDate(startDate).ToString("yyyy-MM-dd")
        Dim ends = CDate(endDate).ToString("yyyy-MM-dd")
        currDay = CDate(start).ToString("yyyy-MM-dd")
        Dim cmdPub As SqlCommand = New SqlCommand
        cmdPub.Connection = Reports.conn
        cmdPub.CommandType = CommandType.Text
        While currDay <= ends
            cmdPub.CommandText = "SELECT * FROM Public_Holidays WHERE DayOfHoliday='" & CDate(currDay).ToString("yyyy-MM-dd") & "'"
            If currDay.DayOfWeek = DayOfWeek.Monday Or currDay.DayOfWeek = DayOfWeek.Tuesday Or currDay.DayOfWeek = DayOfWeek.Wednesday Or currDay.DayOfWeek = DayOfWeek.Thursday Or currDay.DayOfWeek = DayOfWeek.Friday Then
                If cmdPub.ExecuteScalar > 0 Then
                    currDay = currDay.AddDays(1)
                Else
                    days.Add(CDate(currDay).ToString("yyyy-MM-dd"))
                    currDay = currDay.AddDays(1)
                End If
            Else
                currDay = currDay.AddDays(1)
            End If
        End While
        Return days
        Reports.conn.Close()
    End Function

    'Converting number of columns to Alphabetical for cases like column AA-AZ (27-51)
    ' I have no idea why I called it alpha....
    Public Function alpha(number) As String 'Returns excel colum name 
        Dim c As Char
        If number > 0 AndAlso number < 27 Then
            c = Convert.ToChar(number + 64)
            Return c
        ElseIf number >= 27 Then
            number = number - 26
            c = Convert.ToChar(number + 64)
            Dim c1 = "A" + c
            Return c1
        End If
    End Function
    Private Sub rdbtnExcel_CheckedChanged(sender As Object, e As EventArgs) Handles rdbtnExcel.CheckedChanged

    End Sub

    Private Sub dgvDataBreaks_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvDataBreaks.CellContentClick

    End Sub

    Private Sub txtEmpID_TextChanged(sender As Object, e As EventArgs) Handles txtEmpID.TextChanged

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub cmbBreakReport_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbBreakReport.SelectedIndexChanged

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

    End Sub

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click

    End Sub

    Private Sub StatusStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles StatusStrip1.ItemClicked

    End Sub

    Private Sub tspgbar_Click(sender As Object, e As EventArgs) Handles tspgbar.Click

    End Sub

    Private Sub lblProg_Click(sender As Object, e As EventArgs) Handles lblProg.Click

    End Sub

    Private Sub rdbtnPDF_CheckedChanged(sender As Object, e As EventArgs) Handles rdbtnPDF.CheckedChanged

    End Sub

    Private Sub rdbtnPrinter_CheckedChanged(sender As Object, e As EventArgs) Handles rdbtnPrinter.CheckedChanged

    End Sub

    Private Sub Label6_Click(sender As Object, e As EventArgs) Handles Label6.Click

    End Sub
End Class