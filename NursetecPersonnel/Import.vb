Imports System.Data.SqlClient
Imports System.IO
Imports System.Text.RegularExpressions
Public Class Import
    'Declarations
    Dim paramClock_Time As SqlParameter
    Dim paramEmpId As SqlParameter
    Dim paramOffice As SqlParameter
    Dim paramClockDate As SqlParameter
    Dim paramClockDateTime As SqlParameter
    Dim paramEmpName As SqlParameter
    Dim paramDetails As SqlParameter
    Dim paramClockTime As SqlParameter
    Dim paramActivityID As SqlParameter
    Dim paramNotes As SqlParameter
    Dim cmd As SqlCommand = New SqlCommand
    Dim ignored = 0
    Dim conn As SqlConnection = New SqlConnection("Integrated Security=SSPI;Persist Security Info=False;User ID=dba;Initial Catalog=Nursetec;Data Source=NURLAPTOP26\SQLEXPRESS01")
    Public Sub importData()
        Dim count As Integer = 0
        conn.Open()
        cmd.Parameters.Clear()
        cmd.CommandType = CommandType.StoredProcedure
        cmd.CommandTimeout = 15
        cmd.CommandText = "sp_Import"
        cmd.Connection = conn
        Using ofd As OpenFileDialog = New OpenFileDialog()
            If ofd.ShowDialog() = DialogResult.OK Then
                Dim lines As List(Of String) = File.ReadAllLines(ofd.FileName).ToList()
                For i As Integer = 1 To lines.Count - 1
                    Dim data As String() = lines(i).Split(",")
                    paramClockDate = New SqlParameter("@Clock_Date", SqlDbType.Date)
                    paramClockDate.Value = data(0).Replace("""", "").Trim()

                    paramEmpId = New SqlParameter("@Employee_ID", SqlDbType.Int)
                    paramEmpId.Value = data(1).Replace("""", "").Trim()

                    paramOffice = New SqlParameter("@Office", SqlDbType.VarChar)
                    paramOffice.Value = data(3).Replace("""", "").Trim()

                    paramClockTime = New SqlParameter("@Clock_Time", SqlDbType.Time)
                    paramClockTime.Value = data(5).Replace("""", "").Trim()

                    paramActivityID = New SqlParameter("@ActivityID", SqlDbType.Int)

                    If data(6).Replace("""", "").Trim() = "Punch In" Then
                        paramActivityID.Value = 1
                    Else
                        paramActivityID.Value = 2
                    End If

                    paramNotes = New SqlParameter("@Notes", SqlDbType.VarChar)
                    paramNotes.Value = data(11).Replace("""", "").Trim()
                    If paramEmpId.Value IsNot "" Then
                        cmd.Parameters.Add(paramEmpId)
                        cmd.Parameters.Add(paramOffice)
                        cmd.Parameters.Add(paramClockDate)
                        cmd.Parameters.Add(paramClockTime)
                        cmd.Parameters.Add(paramActivityID)
                        cmd.Parameters.Add(paramNotes)
                        cmd.ExecuteNonQuery()
                        cmd.Parameters.Clear()
                        count += 1
                    Else
                        'Ignores record
                        ignored += 1
                    End If
                Next
            End If
            MessageBox.Show("Successfully imported " & count & " records" & vbNewLine & ignored & " records were ignored")
            cmd.CommandType = CommandType.Text
            cmd.CommandText = "SELECT * FROM Clocking"
            Dim adap As SqlDataAdapter = New SqlDataAdapter
            Dim ds As DataSet = New DataSet
            adap.SelectCommand = cmd
            adap.Fill(ds)
            conn.Close()

        End Using

    End Sub
    Public Sub importBio()
        Dim count As Integer = 0
        conn.Open()
        cmd.Parameters.Clear()
        cmd.CommandType = CommandType.StoredProcedure
        cmd.CommandTimeout = 15
        cmd.CommandText = "sp_ImportBio"
        cmd.Connection = conn
        Dim da As SqlDataAdapter = New SqlDataAdapter
        Dim DataSet As DataSet = New DataSet
        Using ofd As OpenFileDialog = New OpenFileDialog()
            If ofd.ShowDialog() = DialogResult.OK Then
                Dim lines As List(Of String) = File.ReadAllLines(ofd.FileName).ToList()
                For i As Integer = 1 To lines.Count - 1
                    Dim data As String() = lines(i).Split(",")
                    paramClockDate = New SqlParameter("@Clock_Date", SqlDbType.Date)
                    paramClockDate.Value = Convert.ToDateTime(data(3).Replace("'", "").Trim()).ToShortDateString

                    paramEmpId = New SqlParameter("@Employee_ID", SqlDbType.Int)
                    paramEmpId.Value = data(0).Replace("'", "").Trim()

                    paramOffice = New SqlParameter("@Office", SqlDbType.VarChar)
                    cmd.CommandType = CommandType.Text
                    cmd.CommandText = "SELECT OfficeID FROM Employee WHERE Employee_ID ='" & data(0).Replace("'", "").Trim() & "'"
                    da.SelectCommand = cmd
                    cmd.ExecuteNonQuery()
                    da.Fill(DataSet)
                    paramOffice.Value = (DataSet.Tables(0).Rows(0)("OfficeID"))

                    paramClockTime = New SqlParameter("@Clock_Time", SqlDbType.Time)
                    paramClockTime.Value = Convert.ToDateTime(data(3).Replace("""", "").Trim()).ToLongTimeString


                    paramActivityID = New SqlParameter("@ActivityID", SqlDbType.Int)

                    If data(4).Replace("""", "").Trim() = "Check-in" Or data(4).Replace("""", "").Trim() = "Break-In" Or data(4).Replace("""", "").Trim() = "Overtime-In" Then
                        paramActivityID.Value = 1
                    ElseIf data(4).Replace("""", "").Trim() = "Check-out" Or data(4).Replace("""", "").Trim() = "Break-Out" Or data(4).Replace("""", "").Trim() = "Overtime-Out" Then
                        paramActivityID.Value = 2
                    End If

                    If paramEmpId.Value IsNot "" Then
                        cmd.Parameters.Add(paramEmpId)
                        cmd.Parameters.Add(paramOffice)
                        cmd.Parameters.Add(paramClockDate)
                        cmd.Parameters.Add(paramClockTime)
                        cmd.Parameters.Add(paramActivityID)
                        cmd.CommandType = CommandType.StoredProcedure
                        cmd.CommandText = "sp_ImportBio"
                        cmd.ExecuteNonQuery()
                        cmd.Parameters.Clear()
                        count += 1
                    Else
                        'Ignores record
                        ignored += 1
                    End If
                Next
            End If
            MessageBox.Show("Successfully imported " & count & " records" & vbNewLine & ignored & " records were ignored")
            cmd.CommandType = CommandType.Text
            cmd.CommandText = "SELECT * FROM Clocking"
            Dim adap As SqlDataAdapter = New SqlDataAdapter
            Dim ds As DataSet = New DataSet
            adap.SelectCommand = cmd
            adap.Fill(ds)
            conn.Close()

        End Using

    End Sub
    Public Sub importhello()
        Dim Sdate As String
        Dim DateT As String()
        Dim strDay As String
        Dim DateFinal
        Dim count As Integer = 0
        conn.Open()
        cmd.Parameters.Clear()
        cmd.CommandType = CommandType.StoredProcedure
        cmd.CommandTimeout = 15
        cmd.CommandText = "sp_importHello"
        cmd.Connection = conn
        Dim lastID = 0
        Dim da As SqlDataAdapter = New SqlDataAdapter
        Dim dataset As New DataSet
        Using ofd2 As OpenFileDialog = New OpenFileDialog()
            If ofd2.ShowDialog() = DialogResult.OK Then
                Dim lines As List(Of String) = File.ReadAllLines(ofd2.FileName).ToList()
                For i As Integer = 1 To lines.Count - 1
                    Dim data As String() = lines(i).Split(",")
                    Dim prevData As String() = lines(i - 1).Split(",")
                    If data(2).Contains("km/h") Then
                        cmd.CommandType = CommandType.Text
                        cmd.CommandText = "SELECT TOP 1 * FROM ClockingHello
                                            ORDER BY ClockID DESC"
                        'If cmd.CommandTimeout Then
                        'Else
                        '    da.SelectCommand = cmd
                        '    da.Fill(dataset)
                        '    lastID = Integer.Parse(dataset.Tables(0).Rows(0)("ClockID"))
                        'End If
                        cmd.CommandText = "UPDATE ClockingHello
                                             SET Distance ='" & data(0).Replace("""", "").Trim & "', TravelTime = '" & data(1).Replace("""", "").Trim & "'
                                             WHERE ClockID = (SELECT TOP 1 ClockID FROM ClockingHello ORDER BY ClockID Desc)"
                        cmd.ExecuteNonQuery()
                    Else
                        paramClockDateTime = New SqlParameter("@ClockDate", SqlDbType.Date)
                        paramClockDateTime.Value = data(0)
                        Sdate = data(0)
                        DateT = Sdate.Split(" ")
                        strDay = DateT(0).Replace("""", "").Trim
                        Dim month
                        If DateT(1).Replace("""", "").Trim = "Jan" Then
                            month = "January"
                        ElseIf DateT(1).Replace("""", "").Trim = "Feb" Then
                            month = "Febuary"
                        ElseIf DateT(1).Replace("""", "").Trim = "Mar" Then
                            month = "March"
                        ElseIf DateT(1).Replace("""", "").Trim = "Apr" Then
                            month = "April"
                        ElseIf DateT(1).Replace("""", "").Trim = "May" Then
                            month = "May"
                        ElseIf DateT(1).Replace("""", "").Trim = "Jun" Then
                            month = "June"
                        ElseIf DateT(1).Replace("""", "").Trim = "Jul" Then
                            month = "July"
                        ElseIf DateT(1).Replace("""", "").Trim = "Aug" Then
                            month = "August"
                        ElseIf DateT(1).Replace("""", "").Trim = "Sep" Then
                            month = "September"
                        ElseIf DateT(1).Replace("""", "").Trim = "Oct" Then
                            month = "October"
                        ElseIf DateT(1).Replace("""", "").Trim = "Nov" Then
                            month = "November"
                        ElseIf DateT(1).Replace("""", "").Trim = "Dec" Then
                            month = "December"
                        End If
                        Dim strTime = DateT(2).Replace("""", "").Trim
                        Dim strFullDate As String = strDay & "/" & month + "/" & Now.Year
                        DateFinal = strFullDate

                        paramClock_Time = New SqlParameter("@ClockTime", SqlDbType.Time)
                        paramClock_Time.Value = strTime

                        paramClockDateTime.Value = DateFinal

                        paramEmpName = New SqlParameter("@Employee_Name", SqlDbType.NVarChar)
                        paramEmpName.Value = data(1).Replace("""", "").Trim

                        paramDetails = New SqlParameter("@Details", SqlDbType.NVarChar)
                        paramDetails.Value = data(2).Replace("""", "").Trim

                        If paramEmpName.Value IsNot "" Then
                            cmd.CommandType = CommandType.StoredProcedure
                            cmd.CommandText = "sp_importHello"
                            cmd.Parameters.Add(paramClockDateTime)
                            cmd.Parameters.Add(paramClock_Time)
                            cmd.Parameters.Add(paramEmpName)
                            cmd.Parameters.Add(paramDetails)
                            cmd.ExecuteNonQuery()
                            cmd.Parameters.Clear()
                            count += 1
                        Else
                            'Ignores record
                            ignored += 1
                        End If
                    End If
                    data(2) = Nothing
                Next
            End If
            MessageBox.Show("successfully imported" & vbNewLine & ignored & " records were ignored")
            cmd.CommandType = CommandType.Text
            cmd.CommandText = "SELECT * FROM ClockingHello"
            Dim adap As SqlDataAdapter = New SqlDataAdapter
            Dim ds As DataSet = New DataSet
            adap.SelectCommand = cmd
            adap.Fill(ds)
            conn.Close()
        End Using
    End Sub
End Class
