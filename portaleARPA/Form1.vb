Imports System.Threading
Imports System.Configuration
Imports System.Data.SqlClient

Public Class Form1

    Dim connectionString As String
    Dim culture As System.Globalization.CultureInfo
    Dim reportType As Int32
    Dim section As Int32
    Dim ret As Int32
    Dim ret2 As Int32


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        connectionString = ConfigurationManager.ConnectionStrings("GLOBAL_CONN_STR").ConnectionString
        DateTimePicker1.Value = Date.Now.AddYears(-1)
        culture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US")
        culture.NumberFormat.NumberGroupSeparator = ""

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        reportType = ComboBox2.SelectedIndex
        section = 8
        Dim startDate As Date = DateTimePicker1.Value
        Dim endDate As Date = DateTimePicker2.Value

        GetData(startDate, endDate, section, reportType)

    End Sub

    Private Async Sub GetData(startTime As Date, endTime As Date, section As Int32, type As Int32)

        Dim command As System.Data.SqlClient.SqlCommand
        Dim command2 As System.Data.SqlClient.SqlCommand
        Dim reader As System.Data.SqlClient.SqlDataReader
        Dim connection As New SqlConnection(connectionString)
        Dim connection2 As New SqlConnection(connectionString)

        Try
            ' Tenta di aprire la connessione
            connection.Open()
            connection2.Open()
            MessageBox.Show("Connessione al database riuscita!", "Successo", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            ' Gestione degli errori
            MessageBox.Show("Errore durante la connessione al database: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

        Dim testCMD As Data.SqlClient.SqlCommand = New Data.SqlClient.SqlCommand("sp_AQMSNT_FILL_ARPA_MASSICI_CAMINI_NODELETE", connection)
        testCMD.CommandTimeout = 18000
        testCMD.CommandType = Data.CommandType.StoredProcedure
        testCMD.Parameters.Add("@idsez", Data.SqlDbType.Int, 11)
        testCMD.Parameters("@idsez").Direction = Data.ParameterDirection.Input
        testCMD.Parameters("@idsez").Value = section
        testCMD.Parameters.Add("@data", Data.SqlDbType.DateTime, 11)
        testCMD.Parameters("@data").Direction = Data.ParameterDirection.Input
        testCMD.Parameters("@data").Value = startTime
        testCMD.Parameters.Add("@TIPO_ESTRAZIONE", Data.SqlDbType.Int, 11)
        testCMD.Parameters("@TIPO_ESTRAZIONE").Direction = Data.ParameterDirection.Input
        testCMD.Parameters("@TIPO_ESTRAZIONE").Value = reportType
        testCMD.Parameters.Add("@retval", Data.SqlDbType.Int)
        testCMD.Parameters("@retval").Direction = Data.ParameterDirection.Output

        testCMD.ExecuteScalar()
        ret = testCMD.Parameters("@retval").Value

        testCMD.Parameters("@idsez").Value = 1
        testCMD.ExecuteScalar()

        ret2 = testCMD.Parameters("@retval").Value
        Dim log_statement As String = "SELECT * FROM [ARPA_WEB_MASSICI_CAMINI] WHERE IDX_REPORT = " & ret.ToString() & " AND TIPO_DATO IS NULL ORDER BY INS_ORDER"
        command = New System.Data.SqlClient.SqlCommand(log_statement, connection)
        reader = command.ExecuteReader()
        log_statement = "SELECT * FROM [ARPA_WEB_MASSICI_CAMINI] WHERE IDX_REPORT = " & ret2.ToString() & " AND TIPO_DATO IS NULL ORDER BY INS_ORDER"
        command2 = New System.Data.SqlClient.SqlCommand(log_statement, connection2)
        Dim reader2 As System.Data.SqlClient.SqlDataReader
        reader2 = command2.ExecuteReader()

        If (reader.HasRows) Then
            While reader.Read()
                reader2.Read()
                Console.WriteLine(String.Format("Il valore NOX è: {0}", reader("E1Q_NOX")))
                Console.WriteLine(String.Format("Il valore SO2 è: {0}", reader("E1Q_SO2")))
                Console.WriteLine(String.Format("Il valore SO2 è: {0}", reader("E1Q_POLVERI")))
            End While
        End If


        Button1.Enabled = False



        'ProgressBar1.Location = New Point(465, 501)
        'ProgressBar1.Visible = True


        ' For i As Integer = 1 To 4
        'Await Task.Delay(5000) ' Simula un lavoro lungo in modo asincrono
        'ProgressBar1.Value += 25
        'Next


        'ProgressBar1.Visible = False

        Button2.Visible = True
        Button3.Visible = True

    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged

        Dim startDate = DateTimePicker1.Value
        Dim endDate = DateTimePicker2.Value

        If endDate >= startDate Then
            Button1.Enabled = True
        Else
            Button1.Enabled = False
        End If

    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged

        Dim startDate = DateTimePicker1.Value
        Dim endDate = DateTimePicker2.Value

        If endDate >= startDate Then
            Button1.Enabled = True
        Else
            Button1.Enabled = False
        End If

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

        Dim combobox As ComboBox = CType(sender, ComboBox)

        If combobox.SelectedIndex = 1 Then
            DateTimePicker1.CustomFormat = "MMMM yyyy"
            DateTimePicker2.CustomFormat = "MMMM yyyy"

        ElseIf combobox.SelectedIndex = 0 Then
            DateTimePicker1.CustomFormat = "yyyy"
            DateTimePicker2.CustomFormat = "yyyy"

        End If

    End Sub
End Class
