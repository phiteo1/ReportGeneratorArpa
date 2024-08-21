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

    Private Async Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        reportType = ComboBox2.SelectedIndex
        section = 8
        Dim startDate As Date = DateTimePicker1.Value
        Dim endDate As Date = DateTimePicker2.Value
        ProgressBar1.Location = New Point(465, 501)
        ProgressBar1.Visible = True
        ProgressBar1.Maximum = 100
        Dim progress As New Progress(Of Integer)(Sub(v)
                                                     ' Questo lambda viene eseguito nel contesto del thread della GUI
                                                     ' Quindi può aggiornare in sicurezza i controlli del modulo
                                                     ProgressBar1.Value = v
                                                 End Sub)

        ' Esegui l'operazione in un altro thread
        Await Task.Run(Sub() GetData(progress, startDate, endDate, section, reportType))
        ProgressBar1.Visible = False
        Button1.Enabled = False
        Button2.Visible = True
        Button3.Visible = True
    End Sub

    Private Async Sub GetData(progress As IProgress(Of Integer), startTime As Date, endTime As Date, section As Int32, type As Int32)

        Dim command As System.Data.SqlClient.SqlCommand
        Dim command2 As System.Data.SqlClient.SqlCommand
        Dim reader As System.Data.SqlClient.SqlDataReader
        Dim connection As New SqlConnection(connectionString)
        Dim connection2 As New SqlConnection(connectionString)
        Dim queryNumber As Integer = 0
        Dim queriesCount As Integer = 4
        Dim progressStep As Integer = 100 \ queriesCount

        Try
            ' Tenta di aprire la connessione
            connection.Open()
            connection2.Open()
        Catch ex As Exception
            ' Gestione degli errori
            MessageBox.Show("Errore durante la connessione al database: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
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
        queryNumber += 1
        progress.Report(queryNumber * progressStep)

        testCMD.Parameters("@idsez").Value = 1
        testCMD.ExecuteScalar()

        queryNumber += 1
        progress.Report(queryNumber * progressStep)

        ret2 = testCMD.Parameters("@retval").Value
        Dim log_statement As String = "SELECT * FROM [ARPA_WEB_MASSICI_CAMINI] WHERE IDX_REPORT = " & ret.ToString() & " AND TIPO_DATO IS NULL ORDER BY INS_ORDER"
        command = New System.Data.SqlClient.SqlCommand(log_statement, connection)

        reader = command.ExecuteReader()
        log_statement = "SELECT * FROM [ARPA_WEB_MASSICI_CAMINI] WHERE IDX_REPORT = " & ret2.ToString() & " AND TIPO_DATO IS NULL ORDER BY INS_ORDER"
        command2 = New System.Data.SqlClient.SqlCommand(log_statement, connection2)
        Dim reader2 As System.Data.SqlClient.SqlDataReader
        reader2 = command2.ExecuteReader()
        queryNumber += 1
        progress.Report(queryNumber * progressStep)

        If (reader.HasRows) Then
            queryNumber += 1
            progress.Report(queryNumber * progressStep)
        Else
            MessageBox.Show("Nessun dato trovato per l'anno selezionato! ")
        End If



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
