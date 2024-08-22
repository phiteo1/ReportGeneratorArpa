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
    Dim dgv As DataGridView
    Dim datanh3 As String
    Dim hiddenColumns As New List(Of String)()


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        connectionString = ConfigurationManager.ConnectionStrings("GLOBAL_CONN_STR").ConnectionString
        DateTimePicker1.Value = Date.Now.AddYears(-1)
        culture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US")
        culture.NumberFormat.NumberGroupSeparator = ""
        SetDataGridView()


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
        Dim dataTable As DataTable
        dataTable = Await Task.Run(Function() GetData(progress, startDate, endDate, section, reportType))
        ShowDataGridView(dataTable)
        'ProgressBar1.Visible = False
        'Button1.Enabled = False
        'Button2.Visible = True
        'Button3.Visible = True
    End Sub

    Private Function GetData(progress As IProgress(Of Integer), startTime As Date, endTime As Date, section As Int32, type As Int32) As Data.DataTable

        Dim dt As New Data.DataTable()
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
            Return dt
        End Try

        dt.Columns.Add(New Data.DataColumn("IDX_REPORT", GetType(Double)))
        dt.Columns.Add(New Data.DataColumn("INS_ORDER", GetType(Integer)))
        dt.Columns.Add(New Data.DataColumn("ORA", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E1Q_NOX", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E1Q_SO2", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E1Q_POLVERI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E1Q_CO", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E1Q_COV", GetType(String)))

        dt.Columns.Add(New Data.DataColumn("E2Q_NOX", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E2Q_SO2", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E2Q_POLVERI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E2Q_CO", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E2Q_COV", GetType(String)))

        dt.Columns.Add(New Data.DataColumn("E3Q_NOX", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E3Q_SO2", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E3Q_POLVERI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E3Q_CO", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E3Q_COV", GetType(String)))

        dt.Columns.Add(New Data.DataColumn("E4Q_NOX", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E4Q_SO2", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E4Q_POLVERI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E4Q_CO", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E4Q_COV", GetType(String)))

        dt.Columns.Add(New Data.DataColumn("E7Q_NOX", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E7Q_SO2", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E7Q_POLVERI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E7Q_CO", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E7Q_COV", GetType(String)))

        dt.Columns.Add(New Data.DataColumn("E8Q_NOX", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E8Q_SO2", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E8Q_POLVERI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E8Q_CO", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E8Q_COV", GetType(String)))

        dt.Columns.Add(New Data.DataColumn("E9Q_NOX", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E9Q_SO2", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E9Q_POLVERI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E9Q_CO", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E9Q_COV", GetType(String)))


        dt.Columns.Add(New Data.DataColumn("E9Q_NH3", GetType(String)))

        dt.Columns.Add(New Data.DataColumn("E10Q_NOX", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E10Q_SO2", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E10Q_POLVERI", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E10Q_CO", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("E10Q_COV", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("NOX_SOMMA", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("SO2_SOMMA", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("POLVERI_SOMMA", GetType(String)))


        dt.Columns.Add(New Data.DataColumn("CO_SOMMA", GetType(String)))


        dt.Columns.Add(New Data.DataColumn("COV_SOMMA", GetType(String)))
        dt.Columns.Add(New Data.DataColumn("NH3_SOMMA", GetType(String)))

        dt.Columns.Add(New Data.DataColumn("NOX57_SOMMA", GetType(String)))

        Select Case reportType
            Case 0
                datanh3 = "01/01/2020"
        End Select

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
        queryNumber += 3
        progress.Report(queryNumber * progressStep)

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
        Dim dr As Data.DataRow = dt.NewRow()
        If (reader.HasRows) Then
            While reader.Read()
                reader2.Read()
                dr("IDX_REPORT") = reader("IDX_REPORT")
                dr("INS_ORDER") = String.Format("{0:n0}", reader("INS_ORDER"))
                dr("ORA") = reader("ORA") 'String.Format("{0:n2}", reader("NOX"))

                dr("E1Q_NOX") = reader("E1Q_NOX")
                dr("E1Q_SO2") = reader("E1Q_SO2")
                dr("E1Q_POLVERI") = reader("E1Q_POLVERI")
                dr("E1Q_CO") = reader("E1Q_CO")
                dr("E1Q_COV") = reader("E1Q_COV")

                dr("E2Q_NOX") = reader("E2Q_NOX")
                dr("E2Q_SO2") = reader("E2Q_SO2")
                dr("E2Q_POLVERI") = reader("E2Q_POLVERI")
                dr("E2Q_CO") = reader("E2Q_CO")
                dr("E2Q_COV") = reader("E2Q_COV")

                Try
                    dr("E3Q_NOX") = String.Format(culture, "{0:n2}", Double.Parse(reader2("E1Q_NOX"), culture.NumberFormat))
                Catch e As FormatException         ''il dato non è un double
                    dr("E3Q_NOX") = reader2("E1Q_NOX")
                Catch e As Exception When TypeOf e Is InvalidOperationException OrElse TypeOf e Is InvalidCastException ''non c'è il dato per E3
                    dr("E3Q_NOX") = "--"
                End Try

                Try
                    dr("E3Q_SO2") = String.Format(culture, "{0:n2}", Double.Parse(reader2("E1Q_SO2"), culture.NumberFormat))
                Catch e As FormatException
                    dr("E3Q_SO2") = reader2("E1Q_SO2")
                Catch e As Exception When TypeOf e Is InvalidOperationException OrElse TypeOf e Is InvalidCastException     ''non c'è il dato per E3
                    dr("E3Q_SO2") = "--"
                End Try

                Try
                    dr("E3Q_POLVERI") = String.Format(culture, "{0:n2}", Double.Parse(reader2("E1Q_POLVERI"), culture.NumberFormat))
                Catch e As FormatException
                    dr("E3Q_POLVERI") = reader2("E1Q_POLVERI")
                Catch e As Exception When TypeOf e Is InvalidOperationException OrElse TypeOf e Is InvalidCastException     ''non c'è il dato per E3
                    dr("E3Q_POLVERI") = "--"
                End Try

                Try
                    dr("E3Q_CO") = String.Format(culture, "{0:n2}", Double.Parse(reader2("E1Q_CO"), culture.NumberFormat))
                Catch e As FormatException         ''il dato non è un double
                    dr("E3Q_CO") = reader2("E1Q_CO")
                Catch e As Exception When TypeOf e Is InvalidOperationException OrElse TypeOf e Is InvalidCastException ''non c'è il dato per E3
                    dr("E3Q_CO") = "--"
                End Try


                Try
                    dr("E3Q_COV") = String.Format(culture, "{0:n2}", Double.Parse(reader2("E1Q_COV"), culture.NumberFormat))
                Catch e As FormatException         ''il dato non è un double
                    dr("E3Q_COV") = reader2("E1Q_COV")
                Catch e As Exception When TypeOf e Is InvalidOperationException OrElse TypeOf e Is InvalidCastException ''non c'è il dato per E3
                    dr("E3Q_COV") = "--"
                End Try


                dr("E4Q_NOX") = reader("E4Q_NOX")
                dr("E4Q_SO2") = reader("E4Q_SO2")
                dr("E4Q_POLVERI") = reader("E4Q_POLVERI")
                dr("E4Q_CO") = reader("E4Q_CO")
                dr("E4Q_COV") = reader("E4Q_COV")

                dr("E7Q_NOX") = reader("E7Q_NOX")
                dr("E7Q_SO2") = reader("E7Q_SO2")
                dr("E7Q_POLVERI") = reader("E7Q_POLVERI")
                dr("E7Q_CO") = reader("E7Q_CO")
                dr("E7Q_COV") = reader("E7Q_COV")

                dr("E8Q_NOX") = reader("E8Q_NOX")
                dr("E8Q_SO2") = reader("E8Q_SO2")
                dr("E8Q_POLVERI") = reader("E8Q_POLVERI")
                dr("E8Q_CO") = reader("E8Q_CO")
                dr("E8Q_COV") = reader("E8Q_COV")

                dr("E9Q_NOX") = reader("E9Q_NOX")
                dr("E9Q_SO2") = reader("E9Q_SO2")
                dr("E9Q_POLVERI") = reader("E9Q_POLVERI")
                dr("E9Q_CO") = reader("E9Q_CO")
                dr("E9Q_COV") = reader("E9Q_COV")
                If (reader("E9Q_NH3") IsNot DBNull.Value) Then
                    dr("E9Q_NH3") = reader("E9Q_NH3")
                Else
                    dr("E9Q_NH3") = "0"
                End If

                dr("E10Q_NOX") = reader("E10Q_NOX")
                dr("E10Q_SO2") = reader("E10Q_SO2")
                dr("E10Q_POLVERI") = reader("E10Q_POLVERI")
                dr("E10Q_CO") = reader("E10Q_CO")
                dr("E10Q_COV") = reader("E10Q_COV")
                dr("NOX_SOMMA") = reader("NOX_SOMMA")
                dr("SO2_SOMMA") = reader("SO2_SOMMA")
                dr("POLVERI_SOMMA") = reader("POLVERI_SOMMA")
                dr("CO_SOMMA") = reader("CO_SOMMA")


                dr("COV_SOMMA") = reader("COV_SOMMA")
                dr("NH3_SOMMA") = reader("NH3_SOMMA")
                dr("NOX57_SOMMA") = reader("NOX57_SOMMA")
                dt.Rows.Add(dr)
                dr = dt.NewRow()


                If (startTime < Date.Parse(datanh3)) Then
                    hiddenColumns.Add("E9Q_NH3")
                    hiddenColumns.Add("NH3_SOMMA")
                    hiddenColumns.Add("NOX57_SOMMA")
                End If


                For Each column As DataGridViewColumn In dgv.Columns
                    ' Verifica se il nome della colonna è nella lista delle colonne nascoste
                    If hiddenColumns.Contains(column.DataPropertyName) Then
                        column.Visible = False
                    End If
                Next

                If hiddenColumns.Count = 0 Then
                    For Each column As DataGridViewColumn In dgv.Columns
                        ' Verifica se il nome della colonna corrisponde ai nomi specificati
                        If column.DataPropertyName = "Ed9Q_NH3" Or column.DataPropertyName = "NH3_SOMMA" Or column.DataPropertyName = "NOX57_SOMMA" Then
                            column.Visible = True
                        End If
                    Next
                End If
            End While
            queryNumber += 1
            progress.Report(queryNumber * progressStep)
        End If

        Return dt
    End Function

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

    Private Sub SetDataGridView()

        dgv = New DataGridView()
        dgv.Dock = DockStyle.Fill
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        dgv.AllowUserToAddRows = False
        dgv.AllowUserToDeleteRows = False
        dgv.AllowUserToResizeRows = False
        dgv.RowHeadersVisible = False
        dgv.Width = 1800
        dgv.AutoGenerateColumns = True



        For Each col As DataGridViewColumn In dgv.Columns
            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            col.DefaultCellStyle.Font = New Font(dgv.Font, FontStyle.Bold)
        Next

    End Sub

    Private Sub ShowDataGridView(dataTable As DataTable)
        ' Nascondi tutti i controlli del modulo eccetto la DataGridView
        For Each ctrl As Control In Controls
            If Not ctrl.Equals(dgv) Then
                ctrl.Visible = False
            End If
        Next

        ' Imposta la DataGridView come visibile e imposta i dati
        dgv.DataSource = dataTable


        Controls.Add(dgv)
        dgv.Visible = True
        ' Ridimensiona la DataGridView per occupare tutto lo spazio disponibile
        dgv.Dock = DockStyle.Fill

    End Sub

End Class
