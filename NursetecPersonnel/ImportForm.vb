Imports System.Data.SqlClient

Public Class ImportForm
    Dim rangePickedStart As Boolean = False
    Dim rangePickedEnd As Boolean = False
    Dim atStartDate As SqlParameter = New SqlParameter("@startDate", SqlDbType.Date)
    Dim atEndDate As SqlParameter = New SqlParameter("@endDate", SqlDbType.Date)
    Dim conn As SqlConnection = New SqlConnection("Integrated Security=SSPI;Persist Security Info=False;User ID=dba;Initial Catalog=Nursetec;Data Source=NURLAPTOP26\SQLEXPRESS01")
    Dim cmd As SqlCommand = New SqlCommand
    Private Sub btnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click

        conn.Open()
        cmd.Connection = conn
        cmd.CommandType = CommandType.StoredProcedure
        cmd.CommandTimeout = 15
        If rangePickedStart = True And rangePickedEnd = True Then
            cmd.CommandText = "sp_DropRange"
            cmd.Parameters.Add(atStartDate)
            cmd.Parameters.Add(atEndDate)
            cmd.ExecuteNonQuery()
            Dim import As New Import
            import.importData()
        Else
            MessageBox.Show("Please select the range to import")
        End If
        conn.Close()

    End Sub

    Private Sub ImportForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.CenterToScreen()
        Me.CenterToParent()
    End Sub

    Private Sub dtpStart_ValueChanged(sender As Object, e As EventArgs) Handles dtpStart.ValueChanged
        atStartDate.Value = CDate(dtpStart.Value).ToString("yyyy-MM-dd")
        rangePickedStart = True
    End Sub

    Private Sub dtpEnd_ValueChanged(sender As Object, e As EventArgs) Handles dtpEnd.ValueChanged
        atEndDate.Value = CDate(dtpEnd.Value).ToString("yyyy-MM-dd")
        rangePickedEnd = True
    End Sub
End Class