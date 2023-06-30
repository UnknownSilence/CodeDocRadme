Imports System.Windows.Forms
Imports System.Data.SqlClient

Public Class FindReplaceDialogSQL
    Dim DataSource As DataSet
    Dim selectedids As ArrayList
    Dim sortColumn As Integer = -1
    Dim bCancelEdit As Boolean
    Dim CurrentItem As ListViewItem
    Dim CurrentSB As ListViewItem.ListViewSubItem
    Public Sub New(ByVal ds As DataSet, ByVal selectedidcounts As ArrayList)
        InitializeComponent()
        Me.DataSource = ds
        Me.selectedids = selectedidcounts
        AddHandler ResultsView.ColumnClick, AddressOf Me.resultsview_ColumnClick
        If selectedids.Count <> 0 Then
            FilterComboBox.SelectedItem = "Selected Items"
        Else
            FilterComboBox.SelectedItem = "Entire Database"
            FilterComboBox.Enabled = False
        End If
    End Sub

    'Button Events
    Private Sub UpdateButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UpdateButton.Click
        For Each updateitem As ListViewItem In ResultsView.SelectedItems
            Dim updatetext As String = updateitem.SubItems(3).Text
            updateitem.SubItems(2).Text = updatetext
            updateitem.SubItems(3).Tag = "Yes"
            updateitem.BackColor = Drawing.Color.LightGreen
            Dim tagcoordinates As String = updateitem.SubItems(2).Tag
            Dim coordinates As Array = tagcoordinates.Split(",")
            Dim i As Integer = coordinates(0)
            Dim q As Integer = coordinates(1)
            Dim z As Integer = updateitem.SubItems(1).Tag
            DataSource.Tables(z).Rows(i).Item(q) = updatetext
        Next
        StatusText.Text = ResultsView.SelectedItems.Count & " Items Updated."
    End Sub
    Private Sub UpdateAllButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UpdateAllButton.Click
        For Each updateitem As ListViewItem In ResultsView.Items
            Dim updatetext As String = updateitem.SubItems(3).Text
            updateitem.SubItems(2).Text = updatetext
            updateitem.SubItems(3).Tag = "Yes"
            updateitem.BackColor = Drawing.Color.LightGreen
            Dim tagcoordinates As String = updateitem.SubItems(2).Tag
            Dim coordinates As Array = tagcoordinates.Split(",")
            Dim i As Integer = coordinates(0)
            Dim q As Integer = coordinates(1)
            Dim z As Integer = updateitem.SubItems(1).Tag
            DataSource.Tables(z).Rows(i).Item(q) = updatetext
        Next
        StatusText.Text = ResultsView.Items.Count & " Items Updated."
    End Sub
    Private Sub FindButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindButton.Click
        Dim unitconverter As New UnitConversion
        Dim convertflags = {0, 0, 0, 0, 0, 0, 0, 0}
        If KPARadio.Checked Then
            convertflags(0) = 1
        End If
        If MPARadio.Checked Then
            convertflags(1) = 1
        End If
        If BARRadio.Checked Then
            convertflags(2) = 1
        End If
        If PSIRadio.Checked Then
            convertflags(3) = 1
        End If
        If INH2ORadio.Checked Then
            convertflags(4) = 1
        End If
        If HPRadio.Checked Then
            convertflags(5) = 1
        End If
        If BTURadio.Checked Then
            convertflags(6) = 1
        End If
        If KGCM2.Checked Then
            convertflags(7) = 1
        End If

        StatusText.Text = "Working..."
        ResultsView.Clear()
        ResultsView.Sorting = Windows.Forms.SortOrder.None
        Dim searchtext As String = FindTextBox.Text
        ResultsView.Columns.Add("Drawing", 180, Windows.Forms.HorizontalAlignment.Left)
        ResultsView.Columns.Add("Data type", 100, Windows.Forms.HorizontalAlignment.Left)
        ResultsView.Columns.Add("Current Data", 100, Windows.Forms.HorizontalAlignment.Left)
        ResultsView.Columns.Add("Preview / Edit", 140, Windows.Forms.HorizontalAlignment.Left)
        ResultsView.Columns.Add("Tag", 100, Windows.Forms.HorizontalAlignment.Left)

        'if a specific table is selected, remove all other tables from the for... to statement
        Dim tablecount As Integer = DataSource.Tables.Count - 1
        Dim tablestartcount As Integer = 0
        Dim selectedtabletag As String = TableComboBox.Tag
        If Not selectedtabletag = "All" Then
            tablecount = selectedtabletag
            tablestartcount = selectedtabletag
        End If

        Dim selectedcolumn As String = ColumnComboBox.SelectedItem
        For z = tablestartcount To tablecount
            Dim workingtable As DataTable = DataSource.Tables(z)
            For i = 0 To workingtable.Rows.Count - 1

                'if a specific column is selected, remove all other columns from the for... to statement
                Dim columnscount As Integer = workingtable.Columns.Count - 1
                Dim columnsstartcount As Integer = 23
                Dim selectedcolumnstag As String = ColumnComboBox.Tag
                If Not selectedcolumnstag = "All" Then
                    columnsstartcount = workingtable.Columns(selectedcolumn).Ordinal
                    columnscount = workingtable.Columns(selectedcolumn).Ordinal
                End If

                For q = columnsstartcount To columnscount
                    Dim columnname As String = workingtable.Columns(q).ColumnName
                    Dim searchitem As String = workingtable.Rows(i).Item(q).ToString
                    '

                    'Need to review this section to be able to add data to specific columns based on tag info
                    If FilterComboBox.SelectedItem = "Selected Items" Then
                        If selectedids.Contains(workingtable.Rows(i).Item("ID_COUNT_")) Then
                            If Not searchitem = Nothing Then
                                If Not FindTextBox.Text = Nothing Then
                                    'Find box has text, only display matching results
                                    If searchitem.Contains(searchtext) Then
                                        Dim drawingname As String = workingtable.Rows(i).Item("DWG_NAME_")
                                        Dim tagid As String = workingtable.Rows(i).Item("TAG_")
                                        Dim listviewresult As ListViewItem = New ListViewItem(drawingname)
                                        listviewresult.Tag = workingtable.Rows(i).Item("ID_COUNT_")
                                        ResultsView.Items.Add(listviewresult)
                                        listviewresult.SubItems.Add(columnname)
                                        listviewresult.SubItems.Add(searchitem)
                                        listviewresult.SubItems.Add(searchitem)
                                        listviewresult.SubItems.Add(tagid)

                                        listviewresult.SubItems(1).Tag = z
                                        listviewresult.SubItems(2).Tag = i & "," & q
                                        listviewresult.SubItems(3).Tag = "No"
                                        Dim updatetext As String = UpdateTextBox.Text
                                        Select Case ActionComboBox.SelectedItem
                                            Case "Replace With:"
                                                Dim replacedtext As String = searchitem.Replace(searchtext, updatetext)
                                                listviewresult.SubItems(3).Text = replacedtext
                                            Case "Increment By:"
                                                If Not updatetext = Nothing Then
                                                    Dim incrementedsearch As String = IncrementNumber(searchtext, updatetext)
                                                    Dim incrementedtext As String = searchitem.Replace(searchtext, incrementedsearch)
                                                    listviewresult.SubItems(3).Text = incrementedtext
                                                Else
                                                    listviewresult.SubItems(3).Text = searchitem
                                                End If
                                            Case "Append Prefix:"
                                                If Not updatetext = Nothing Then
                                                    Dim prefixtext As String = updatetext & searchitem
                                                    listviewresult.SubItems(3).Text = prefixtext
                                                Else
                                                    listviewresult.SubItems(3).Text = searchitem
                                                End If
                                            Case "Append Suffix:"
                                                If Not updatetext = Nothing Then
                                                    Dim suffixtext As String = searchitem & updatetext
                                                    listviewresult.SubItems(3).Text = suffixtext
                                                Else
                                                    listviewresult.SubItems(3).Text = searchitem
                                                End If
                                            Case "Convert Metric/Imperial:"
                                                Dim convertedtext As String = unitconverter.ReplaceConvertedUnits(searchitem, convertflags)
                                                listviewresult.SubItems(3).Text = convertedtext
                                            Case "Append Metric/Imperial:"
                                                Dim convertedtext As String = unitconverter.ReplaceConvertedUnits(searchitem, convertflags)
                                                Dim appendconverttext As String = searchitem & " " & "(" & convertedtext & ")"
                                                listviewresult.SubItems(3).Text = appendconverttext
                                        End Select
                                    End If
                                Else
                                    If selectedtabletag = "All" Then
                                        MsgBox("Please select a specific table and column if you do not have a search key.")
                                        Exit Sub
                                    Else
                                        If selectedcolumnstag = "All" Then
                                            MsgBox("Please select a specific table and column if you do not have a search key.")
                                            Exit Sub
                                        End If
                                    End If
                                    'Find box has no text, display all results
                                    Dim tagid As String = workingtable.Rows(i).Item("TAG_")
                                    Dim drawingname As String = workingtable.Rows(i).Item("DWG_NAME_")
                                    Dim listviewresult As ListViewItem = New ListViewItem(drawingname)
                                    listviewresult.Tag = workingtable.Rows(i).Item("ID_COUNT_")
                                    ResultsView.Items.Add(listviewresult)
                                    listviewresult.SubItems.Add(columnname)
                                    listviewresult.SubItems.Add(searchitem)
                                    listviewresult.SubItems.Add(searchitem)
                                    listviewresult.SubItems.Add(tagid)
                                    listviewresult.SubItems(1).Tag = z
                                    listviewresult.SubItems(2).Tag = i & "," & q
                                    listviewresult.SubItems(3).Tag = "No"
                                    Dim updatetext As String = UpdateTextBox.Text
                                    Select Case ActionComboBox.SelectedItem
                                        Case "Replace With:"
                                            If Not updatetext = Nothing Then
                                                listviewresult.SubItems(3).Text = updatetext
                                            Else
                                                listviewresult.SubItems(3).Text = searchitem
                                            End If
                                        Case "Increment By:"
                                            If Not updatetext = Nothing Then
                                                If Not searchitem = Nothing Then
                                                    Dim incrementedsearch As String = IncrementNumber(searchitem, updatetext)
                                                    Dim incrementedtext As String = searchitem.Replace(searchitem, incrementedsearch)
                                                    listviewresult.SubItems(3).Text = incrementedtext
                                                End If
                                            Else
                                                listviewresult.SubItems(3).Text = searchitem
                                            End If
                                        Case "Append Prefix:"
                                            If Not updatetext = Nothing Then
                                                Dim prefixtext As String = updatetext & searchitem
                                                listviewresult.SubItems(3).Text = prefixtext
                                            Else
                                                listviewresult.SubItems(3).Text = searchitem
                                            End If
                                        Case "Append Suffix:"
                                            If Not updatetext = Nothing Then
                                                Dim suffixtext As String = searchitem & updatetext
                                                listviewresult.SubItems(3).Text = suffixtext
                                            Else
                                                listviewresult.SubItems(3).Text = searchitem
                                            End If
                                        Case "Convert Metric/Imperial:"
                                            Dim convertedtext As String = unitconverter.ReplaceConvertedUnits(searchitem, convertflags)
                                            listviewresult.SubItems(3).Text = convertedtext
                                        Case "Append Metric/Imperial:"
                                            Dim convertedtext As String = unitconverter.ReplaceConvertedUnits(searchitem, convertflags)
                                            Dim appendconverttext As String = searchitem & " " & "(" & convertedtext & ")"
                                            listviewresult.SubItems(3).Text = appendconverttext
                                    End Select
                                End If
                            End If
                        End If
                    ElseIf FilterComboBox.SelectedItem = "Entire Database" Then
                        If Not searchitem = Nothing Then
                            If Not FindTextBox.Text = Nothing Then
                                'Find box has text, only display matching results
                                If searchitem.Contains(searchtext) Then
                                    Dim drawingname As String = workingtable.Rows(i).Item("DWG_NAME_")
                                    Dim tagid As String = workingtable.Rows(i).Item("TAG_")
                                    Dim listviewresult As ListViewItem = New ListViewItem(drawingname)
                                    listviewresult.Tag = workingtable.Rows(i).Item("ID_COUNT_")
                                    ResultsView.Items.Add(listviewresult)
                                    listviewresult.SubItems.Add(columnname)
                                    listviewresult.SubItems.Add(searchitem)
                                    listviewresult.SubItems.Add(searchitem)
                                    listviewresult.SubItems.Add(tagid)

                                    listviewresult.SubItems(1).Tag = z
                                    listviewresult.SubItems(2).Tag = i & "," & q
                                    listviewresult.SubItems(3).Tag = "No"
                                    Dim updatetext As String = UpdateTextBox.Text
                                    Select Case ActionComboBox.SelectedItem
                                        Case "Replace With:"
                                            Dim replacedtext As String = searchitem.Replace(searchtext, updatetext)
                                            listviewresult.SubItems(3).Text = replacedtext
                                        Case "Increment By:"
                                            If Not updatetext = Nothing Then
                                                Dim incrementedsearch As String = IncrementNumber(searchtext, updatetext)
                                                Dim incrementedtext As String = searchitem.Replace(searchtext, incrementedsearch)
                                                listviewresult.SubItems(3).Text = incrementedtext
                                            Else
                                                listviewresult.SubItems(3).Text = searchitem
                                            End If
                                        Case "Append Prefix:"
                                            If Not updatetext = Nothing Then
                                                Dim prefixtext As String = updatetext & searchitem
                                                listviewresult.SubItems(3).Text = prefixtext
                                            Else
                                                listviewresult.SubItems(3).Text = searchitem
                                            End If
                                        Case "Append Suffix:"
                                            If Not updatetext = Nothing Then
                                                Dim suffixtext As String = searchitem & updatetext
                                                listviewresult.SubItems(3).Text = suffixtext
                                            Else
                                                listviewresult.SubItems(3).Text = searchitem
                                            End If
                                        Case "Convert Metric/Imperial:"
                                            Dim convertedtext As String = unitconverter.ReplaceConvertedUnits(searchitem, convertflags)
                                            listviewresult.SubItems(3).Text = convertedtext
                                        Case "Append Metric/Imperial:"
                                            Dim convertedtext As String = unitconverter.ReplaceConvertedUnits(searchitem, convertflags)
                                            Dim appendconverttext As String = searchitem & " " & "(" & convertedtext & ")"
                                            listviewresult.SubItems(3).Text = appendconverttext
                                    End Select
                                End If
                            Else
                                If selectedtabletag = "All" Then
                                    MsgBox("Please select a specific table and column if you do not have a search key.")
                                    Exit Sub
                                Else
                                    If selectedcolumnstag = "All" Then
                                        MsgBox("Please select a specific table and column if you do not have a search key.")
                                        Exit Sub
                                    End If
                                End If
                                'Find box has no text, display all results
                                Dim tagid As String = workingtable.Rows(i).Item("TAG_")
                                Dim drawingname As String = workingtable.Rows(i).Item("DWG_NAME_")
                                Dim listviewresult As ListViewItem = New ListViewItem(drawingname)
                                listviewresult.Tag = workingtable.Rows(i).Item("ID_COUNT_")
                                ResultsView.Items.Add(listviewresult)
                                listviewresult.SubItems.Add(columnname)
                                listviewresult.SubItems.Add(searchitem)
                                listviewresult.SubItems.Add(searchitem)
                                listviewresult.SubItems.Add(tagid)
                                listviewresult.SubItems(1).Tag = z
                                listviewresult.SubItems(2).Tag = i & "," & q
                                listviewresult.SubItems(3).Tag = "No"
                                Dim updatetext As String = UpdateTextBox.Text
                                Select Case ActionComboBox.SelectedItem
                                    Case "Replace With:"
                                        If Not updatetext = Nothing Then
                                            listviewresult.SubItems(3).Text = updatetext
                                        Else
                                            listviewresult.SubItems(3).Text = searchitem
                                        End If
                                    Case "Increment By:"
                                        If Not updatetext = Nothing Then
                                            If Not searchitem = Nothing Then
                                                Dim incrementedsearch As String = IncrementNumber(searchitem, updatetext)
                                                Dim incrementedtext As String = searchitem.Replace(searchitem, incrementedsearch)
                                                listviewresult.SubItems(3).Text = incrementedtext
                                            End If
                                        Else
                                            listviewresult.SubItems(3).Text = searchitem
                                        End If
                                    Case "Append Prefix:"
                                        If Not updatetext = Nothing Then
                                            Dim prefixtext As String = updatetext & searchitem
                                            listviewresult.SubItems(3).Text = prefixtext
                                        Else
                                            listviewresult.SubItems(3).Text = searchitem
                                        End If
                                    Case "Append Suffix:"
                                        If Not updatetext = Nothing Then
                                            Dim suffixtext As String = searchitem & updatetext
                                            listviewresult.SubItems(3).Text = suffixtext
                                        Else
                                            listviewresult.SubItems(3).Text = searchitem
                                        End If
                                    Case "Convert Metric/Imperial:"
                                        Dim convertedtext As String = unitconverter.ReplaceConvertedUnits(searchitem, convertflags)
                                        listviewresult.SubItems(3).Text = convertedtext
                                    Case "Append Metric/Imperial:"
                                        Dim convertedtext As String = unitconverter.ReplaceConvertedUnits(searchitem, convertflags)
                                        Dim appendconverttext As String = searchitem & " " & "(" & convertedtext & ")"
                                        listviewresult.SubItems(3).Text = appendconverttext
                                End Select
                            End If
                        End If
                    End If
                    '
                    
                Next
            Next
            workingtable.Dispose()
        Next
        StatusText.Text = ResultsView.Items.Count & " Results Found"
    End Sub
    Private Sub DoneButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DoneButton.Click
        Me.Close()
    End Sub
    Private Sub ApplyButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ApplyButton.Click
        Select Case MsgBox("This will apply the changes to the SQL Database. Are you sure you want to continue?", MsgBoxStyle.YesNo)
            Case MsgBoxResult.Yes
                Dim count As Integer = 0
                Dim sqlcommand As String = ""
                Dim databasename As String = DataSource.Tables(0).TableName
                For Each listeditem As ListViewItem In ResultsView.Items
                    If listeditem.SubItems(3).Tag = "Yes" Then
                        count += 1
                        'update dataset
                        Dim tagcoordinates As String = listeditem.SubItems(2).Tag
                        Dim coordinates As Array = tagcoordinates.Split(",")
                        Dim i As Integer = coordinates(0)
                        Dim q As Integer = coordinates(1)
                        Dim newtext As String = listeditem.SubItems(2).Text
                        Dim tablenum As Integer = listeditem.SubItems(1).Tag
                        DataSource.Tables(tablenum).Rows(i).Item(q) = newtext

                        'create sql command string
                        Dim tablename As String = ""
                        Select Case tablenum
                            Case 0
                                tablename = "PID_Components_Instruments"
                            Case 1
                                tablename = "PID_Components_Mechanical"
                            Case 2
                                tablename = "PID_Components_Miscellaneous"
                            Case 3
                                tablename = "PID_Components_Nozzles"
                            Case 4
                                tablename = "PID_Components_Process_Lines"
                            Case 5
                                tablename = "PID_Components_Reducers"
                            Case 6
                                tablename = "PID_Components_Valves"
                            Case 7
                                tablename = "PID_Components_Vessels"
                            Case 8
                                tablename = "PID_User1_TIE_IN"
                        End Select
                        Dim idcount As String = listeditem.Tag
                        Dim columnname As String = listeditem.SubItems(1).Text
                        sqlcommand = sqlcommand & "UPDATE " & tablename & vbCrLf & "SET " & columnname & " = '" & newtext & "'" & vbCrLf & "Where ID_COUNT_ = '" & idcount & "'" & vbCrLf
                    End If
                Next

                If count <> 0 Then
                    'send command to sql
                    Dim rcfg As New ReadConfig
                    Dim dsrc As String = rcfg.ReadLocalConfig(1).ToString.ToUpper '" & dsrc & "
                    Dim connectionString As String = "Data Source=" & dsrc & ";Initial Catalog=" & databasename & ";Integrated Security=True"
                    Dim connection As New SqlConnection(connectionString)
                    Dim cmd As New SqlCommand
                    cmd.Connection = connection
                    Try
                        connection.Open()
                        cmd.CommandText = sqlcommand
                        cmd.ExecuteNonQuery()
                    Catch ex As Exception
                        MsgBox("Error while inserting record on table... " & vbCrLf & ex.Message)
                    Finally
                        connection.Close()
                    End Try
                    MsgBox("Database has been updated.")
                    ResultsView.Clear()
                Else
                    MsgBox("No changes have been made. Nothing has been done.")
                End If

            Case MsgBoxResult.No
                MsgBox("Database has NOT been updated.")
        End Select

    End Sub

    'Text Filtering
    Private Sub UpdateTextBox_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles UpdateTextBox.KeyPress
        Select Case ActionComboBox.SelectedItem
            Case "Replace With:"
            Case "Increment By:"
                Select Case Asc(e.KeyChar)
                    Case 127
                    Case 48 To 57
                    Case 45
                    Case 8
                    Case Else
                        e.Handled = True
                End Select
            Case "Append Prefix:"
            Case "Append Suffix:"
            Case "Convert Metric/Imperial:"
                e.Handled = True
            Case "Append Metric/Imperial:"
                e.Handled = True
        End Select
    End Sub

    'Combo Boxes
    Private Sub TableComboBox_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TableComboBox.SelectedIndexChanged
        ColumnComboBox.Items.Clear()
        ColumnComboBox.Tag = "All"

        Select Case TableComboBox.SelectedItem
            Case "All Tables"
                TableComboBox.Tag = "All"
                ColumnComboBox.SelectedItem = "All Columns"
                ColumnComboBox.Enabled = False
                ColumnComboBox.Tag = "All"
            Case "Instruments"
                TableComboBox.Tag = "0"
                ColumnComboBox.Enabled = True
                Dim columnstable As DataTable = DataSource.Tables(0)
                ColumnComboBox.Items.Add("All Columns")
                For i = 23 To columnstable.Columns.Count - 1
                    ColumnComboBox.Items.Add(columnstable.Columns(i).ColumnName)
                Next
                columnstable.Dispose()
            Case "Mechanical"
                TableComboBox.Tag = "1"
                ColumnComboBox.Enabled = True
                Dim columnstable As DataTable = DataSource.Tables(1)
                ColumnComboBox.Items.Add("All Columns")
                For i = 23 To columnstable.Columns.Count - 1
                    ColumnComboBox.Items.Add(columnstable.Columns(i).ColumnName)
                Next
                columnstable.Dispose()
            Case "Miscellaneous"
                TableComboBox.Tag = "2"
                ColumnComboBox.Enabled = True
                Dim columnstable As DataTable = DataSource.Tables(2)
                ColumnComboBox.Items.Add("All Columns")
                For i = 23 To columnstable.Columns.Count - 1
                    ColumnComboBox.Items.Add(columnstable.Columns(i).ColumnName)
                Next
                columnstable.Dispose()
            Case "Nozzles"
                TableComboBox.Tag = "3"
                ColumnComboBox.Enabled = True
                Dim columnstable As DataTable = DataSource.Tables(3)
                ColumnComboBox.Items.Add("All Columns")
                For i = 23 To columnstable.Columns.Count - 1
                    ColumnComboBox.Items.Add(columnstable.Columns(i).ColumnName)
                Next
                columnstable.Dispose()
            Case "Process Lines"
                TableComboBox.Tag = "4"
                ColumnComboBox.Enabled = True
                Dim columnstable As DataTable = DataSource.Tables(4)
                ColumnComboBox.Items.Add("All Columns")
                For i = 23 To columnstable.Columns.Count - 1
                    ColumnComboBox.Items.Add(columnstable.Columns(i).ColumnName)
                Next
                columnstable.Dispose()
            Case "Reducers"
                TableComboBox.Tag = "5"
                ColumnComboBox.Enabled = True
                Dim columnstable As DataTable = DataSource.Tables(5)
                ColumnComboBox.Items.Add("All Columns")
                For i = 23 To columnstable.Columns.Count - 1
                    ColumnComboBox.Items.Add(columnstable.Columns(i).ColumnName)
                Next
                columnstable.Dispose()
            Case "Valves"
                TableComboBox.Tag = "6"
                ColumnComboBox.Enabled = True
                Dim columnstable As DataTable = DataSource.Tables(6)
                ColumnComboBox.Items.Add("All Columns")
                For i = 23 To columnstable.Columns.Count - 1
                    ColumnComboBox.Items.Add(columnstable.Columns(i).ColumnName)
                Next
                columnstable.Dispose()
            Case "Vessels"
                TableComboBox.Tag = "7"
                ColumnComboBox.Enabled = True
                Dim columnstable As DataTable = DataSource.Tables(7)
                ColumnComboBox.Items.Add("All Columns")
                For i = 23 To columnstable.Columns.Count - 1
                    ColumnComboBox.Items.Add(columnstable.Columns(i).ColumnName)
                Next
                columnstable.Dispose()
            Case "Tie Ins"
                TableComboBox.Tag = "8"
                ColumnComboBox.Enabled = True
                Dim columnstable As DataTable = DataSource.Tables(8)
                ColumnComboBox.Items.Add("All Columns")
                For i = 23 To columnstable.Columns.Count - 1
                    ColumnComboBox.Items.Add(columnstable.Columns(i).ColumnName)
                Next
                columnstable.Dispose()
        End Select
    End Sub
    Private Sub ColumnComboBox_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ColumnComboBox.SelectedIndexChanged
        If ColumnComboBox.SelectedItem = "All Columns" Then
            ColumnComboBox.Tag = "All"
        Else
            ColumnComboBox.Tag = Nothing
        End If
    End Sub
    Private Sub ActionComboBox_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ActionComboBox.SelectedIndexChanged
        'UpdateTextBox.Clear()
        Select Case ActionComboBox.SelectedItem
            Case "Replace With:"
                UpdateTextBox.Enabled = True
                MetricGroupBox.Visible = False
                ImperialGroupBox1.Visible = False
                ImperialGroupBox2.Visible = False
                Dim ReplaceGroupSize As New System.Drawing.Size(530, 50)
                Dim ResultsViewLocation As New System.Drawing.Size(14, 127)
                Dim SubtractSize As New System.Drawing.Size(44, 205)
                Dim ResultsViewSize As New System.Drawing.Size
                ResultsViewSize = System.Drawing.Size.Subtract(Me.Size, SubtractSize)
                ReplaceGroupBox.Size = ReplaceGroupSize
                ResultsView.Location = ResultsViewLocation
                ResultsView.Size = ResultsViewSize
            Case "Increment By:"
                UpdateTextBox.Enabled = True
                MetricGroupBox.Visible = False
                ImperialGroupBox1.Visible = False
                ImperialGroupBox2.Visible = False
                Dim ReplaceGroupSize As New System.Drawing.Size(530, 50)
                Dim ResultsViewLocation As New System.Drawing.Size(14, 127)
                Dim SubtractSize As New System.Drawing.Size(44, 205)
                Dim ResultsViewSize As New System.Drawing.Size
                ResultsViewSize = System.Drawing.Size.Subtract(Me.Size, SubtractSize)
                ReplaceGroupBox.Size = ReplaceGroupSize
                ResultsView.Location = ResultsViewLocation
                ResultsView.Size = ResultsViewSize
            Case "Append Prefix:"
                UpdateTextBox.Enabled = True
                MetricGroupBox.Visible = False
                ImperialGroupBox1.Visible = False
                ImperialGroupBox2.Visible = False
                Dim ReplaceGroupSize As New System.Drawing.Size(530, 50)
                Dim ResultsViewLocation As New System.Drawing.Size(14, 127)
                Dim SubtractSize As New System.Drawing.Size(44, 205)
                Dim ResultsViewSize As New System.Drawing.Size
                ResultsViewSize = System.Drawing.Size.Subtract(Me.Size, SubtractSize)
                ReplaceGroupBox.Size = ReplaceGroupSize
                ResultsView.Location = ResultsViewLocation
                ResultsView.Size = ResultsViewSize
            Case "Append Suffix:"
                UpdateTextBox.Enabled = True
                MetricGroupBox.Visible = False
                ImperialGroupBox1.Visible = False
                ImperialGroupBox2.Visible = False
                Dim ReplaceGroupSize As New System.Drawing.Size(530, 50)
                Dim ResultsViewLocation As New System.Drawing.Size(14, 127)
                Dim SubtractSize As New System.Drawing.Size(44, 205)
                Dim ResultsViewSize As New System.Drawing.Size
                ResultsViewSize = System.Drawing.Size.Subtract(Me.Size, SubtractSize)
                ReplaceGroupBox.Size = ReplaceGroupSize
                ResultsView.Location = ResultsViewLocation
                ResultsView.Size = ResultsViewSize
            Case "Convert Metric/Imperial:"
                UpdateTextBox.Enabled = False
                MetricGroupBox.Visible = True
                ImperialGroupBox1.Visible = True
                ImperialGroupBox2.Visible = True
                Dim ReplaceGroupSize As New System.Drawing.Size(530, 103)
                Dim ResultsViewLocation As New System.Drawing.Size(14, 180)
                Dim SubtractSize As New System.Drawing.Size(44, 258)
                Dim ResultsViewSize As New System.Drawing.Size
                ResultsViewSize = System.Drawing.Size.Subtract(Me.Size, SubtractSize)
                ReplaceGroupBox.Size = ReplaceGroupSize
                ResultsView.Location = ResultsViewLocation
                ResultsView.Size = ResultsViewSize
            Case "Append Metric/Imperial:"
                UpdateTextBox.Enabled = False
                MetricGroupBox.Visible = True
                ImperialGroupBox1.Visible = True
                ImperialGroupBox2.Visible = True
                Dim ReplaceGroupSize As New System.Drawing.Size(530, 103)
                Dim ResultsViewLocation As New System.Drawing.Size(14, 180)
                Dim SubtractSize As New System.Drawing.Size(44, 258)
                Dim ResultsViewSize As New System.Drawing.Size
                ResultsViewSize = System.Drawing.Size.Subtract(Me.Size, SubtractSize)
                ReplaceGroupBox.Size = ReplaceGroupSize
                ResultsView.Location = ResultsViewLocation
                ResultsView.Size = ResultsViewSize
        End Select
    End Sub
    

    'Editing the Listview
    Private Sub resultsview_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs)
        If e.Column <> sortColumn Then
            sortColumn = e.Column
            ResultsView.Sorting = System.Windows.Forms.SortOrder.Ascending
        Else

            If ResultsView.Sorting = System.Windows.Forms.SortOrder.Ascending Then
                ResultsView.Sorting = System.Windows.Forms.SortOrder.Descending
            Else
                ResultsView.Sorting = System.Windows.Forms.SortOrder.Ascending
            End If
        End If

        ResultsView.Sort()
        ResultsView.ListViewItemSorter = New ListViewItemComparer(e.Column, ResultsView.Sorting)
    End Sub
    Private Sub ResultsView_MouseDoubleClick(ByVal sender As Object, _
        ByVal e As System.Windows.Forms.MouseEventArgs) Handles ResultsView.MouseDoubleClick
        ' This subroutine checks where the double-clicking was performed and
        ' initiates in-line editing if user double-clicked on the right subitem

        CurrentItem = ResultsView.GetItemAt(e.X, e.Y)
        If CurrentItem Is Nothing Then Exit Sub
        CurrentSB = CurrentItem.GetSubItemAt(e.X, e.Y)
        Dim iSubIndex As Integer = CurrentItem.SubItems.IndexOf(CurrentSB)
        Select Case iSubIndex
            Case 3
                ' This column is allowed to be edited. So continue the code
            Case Else

                Exit Sub
        End Select
        Dim lLeft = CurrentSB.Bounds.Left + 2
        Dim lWidth As Integer = CurrentSB.Bounds.Width
        With TextBox1
            .SetBounds(lLeft + ResultsView.Left, CurrentSB.Bounds.Top + _
                       ResultsView.Top, lWidth, CurrentSB.Bounds.Height)
            .Text = CurrentSB.Text
            .Show()
            .Focus()
        End With
    End Sub
    Private Sub TextBox1_KeyPress(ByVal sender As Object, _
            ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        ' This subroutine closes the text box
        Select Case e.KeyChar
            Case Microsoft.VisualBasic.ChrW(Keys.Return)
                bCancelEdit = False
                e.Handled = True
                TextBox1.Hide()
            Case Microsoft.VisualBasic.ChrW(Keys.Escape)
                bCancelEdit = True
                e.Handled = True
                TextBox1.Hide()
        End Select
    End Sub
    Private Sub TextBox1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.LostFocus
        TextBox1.Hide()
        If bCancelEdit = False Then
            CurrentSB.Text = TextBox1.Text
        Else
            bCancelEdit = False
        End If
        ResultsView.Focus()
    End Sub

    'Radio Buttons
    

    'Functions
    Function IncrementNumber(ByVal text As String, ByVal incrementvalue As Integer) As String
        Dim letterindexarray As Array = text.ToCharArray
        Dim lettersinstring As String = ""
        Dim numbersinstring As String = ""
        Dim formatstring As String = ""
        Dim integertoincrement As New Integer
        Dim incrementednumber As String = ""
        For i = 0 To letterindexarray.Length - 1
            If Char.IsDigit(letterindexarray(i)) Then
                numbersinstring = numbersinstring & letterindexarray(i)
            Else
                lettersinstring = lettersinstring & letterindexarray(i)
            End If
        Next
        If Not numbersinstring = "" Then
            If Int32.TryParse(numbersinstring, integertoincrement) Then
                incrementednumber = integertoincrement + incrementvalue
            End If
            Dim formattedincrement As String = incrementednumber.ToString
            If Not numbersinstring.Length <= incrementednumber.Length Then
                Dim addedzeros As String = ""
                Dim zerostoadd As Integer = numbersinstring.Length - incrementednumber.Length
                For i = 0 To zerostoadd - 1
                    addedzeros = addedzeros & "0"
                Next
                formattedincrement = addedzeros & incrementednumber
            End If
            Dim replacedtext As String = text.Replace(numbersinstring, formattedincrement)
            Return replacedtext
        Else
            Return text
        End If
    End Function
End Class

'Sorting the Listview
Class ListViewItemComparer
    Implements IComparer
    Private col As Integer
    Private order As System.Windows.Forms.SortOrder
    Public Sub New()
        col = 0
        order = System.Windows.Forms.SortOrder.Ascending
    End Sub

    Public Sub New(ByVal column As Integer, ByVal order As System.Windows.Forms.SortOrder)
        col = column
        Me.order = order
    End Sub
    Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare
        Dim returnVal As Integer = -1
        returnVal = [String].Compare(CType(x, ListViewItem).SubItems(col).Text, CType(y, ListViewItem).SubItems(col).Text)
        If order = System.Windows.Forms.SortOrder.Descending Then
            returnVal *= -1
        End If
        Return returnVal
        'Dim lx As ListViewItem = CType(x, ListViewItem)
        'Dim ly As ListViewItem = CType(y, ListViewItem)
        'Select Case (col)
        '    Case 0
        '        ' first column
        '        Dim c As Integer = String.Compare(lx.SubItems(0).Text, ly.SubItems(0).Text)
        '        ' take care of the other columns if needed
        '        ' modify c to respond to ascending
        '        ' or descending order, as needed
        '        Return c
        'End Select

    End Function
End Class



