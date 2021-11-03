Imports System.Reflection
Imports System.Security.AccessControl
Imports System.Threading
Imports System.Linq
Imports System.Text
Imports System.Runtime.InteropServices
Imports System.IO
Imports System.Runtime.CompilerServices
Imports MySql.Data.MySqlClient

Public Class Form1
    Dim con As New MySqlConnection("server=localhost; user=root; password=;database=athletes;")
    Dim thread1 As System.Threading.Thread

    Public Class custom
        Private Shared mLabels() As String
        Private Shared mLabelIndex As Integer
        Public Shared Sub PatchMsgBox(ByVal labels() As String)
            ''--- Updates message box buttons
            mLabels = labels
            Application.OpenForms(0).BeginInvoke(New FindWindowDelegate(AddressOf FindMsgBox), GetCurrentThreadId())
        End Sub

        Private Shared Sub FindMsgBox(ByVal tid As Integer)
            ''--- Enumerate the windows owned by the UI thread
            EnumThreadWindows(tid, AddressOf EnumWindow, IntPtr.Zero)
        End Sub

        Private Shared Function EnumWindow(ByVal hWnd As IntPtr, ByVal lp As IntPtr) As Boolean
            ''--- Is this the message box?
            Dim sb As New StringBuilder(256)
            GetClassName(hWnd, sb, sb.Capacity)
            If sb.ToString() <> "#32770" Then Return True
            ''--- Got it, now find the buttons
            mLabelIndex = 0
            EnumChildWindows(hWnd, AddressOf FindButtons, IntPtr.Zero)
            Return False
        End Function

        Private Shared Function FindButtons(ByVal hWnd As IntPtr, ByVal lp As IntPtr) As Boolean
            Dim sb As New StringBuilder(256)
            GetClassName(hWnd, sb, sb.Capacity)
            If sb.ToString() = "Button" And mLabelIndex <= UBound(mLabels) Then
                ''--- Got one, update text
                SetWindowText(hWnd, mLabels(mLabelIndex))
                mLabelIndex += 1
            End If
            Return True
        End Function

        ''--- P/Invoke declarations
        Private Delegate Sub FindWindowDelegate(ByVal tid As Integer)
        Private Delegate Function EnumWindowDelegate(ByVal hWnd As IntPtr, ByVal lp As IntPtr) As Boolean
        Private Declare Auto Function EnumThreadWindows Lib "user32.dll" (ByVal tid As Integer, ByVal callback As EnumWindowDelegate, ByVal lp As IntPtr) As Boolean
        Private Declare Auto Function EnumChildWindows Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal callback As EnumWindowDelegate, ByVal lp As IntPtr) As Boolean
        Private Declare Auto Function GetClassName Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal name As StringBuilder, ByVal maxlen As Integer) As Integer
        Private Declare Auto Function GetCurrentThreadId Lib "kernel32.dll" () As Integer
        Private Declare Auto Function SetWindowText Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal text As String) As Boolean
    End Class
    Class athlete
        Public entryId As String
        Public location As String
        Public fName As String
        Public lName As String
        Public amountJumps As Integer
    End Class


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim athletes(29) As athlete 'creates a array of objects
        Dim fileDirectory As String 'string of athlete.csv
        Dim choice As Integer '1 indicates file 2 indicats database
        Dim max As Integer

        custom.PatchMsgBox(New String() {"File", "Database"})
        'MsgBox("Please Select your preferred method of collecting data", MsgBoxStyle.YesNo)

        If (MsgBox("Please Select your preferred method of collecting data", MsgBoxStyle.YesNo) = MsgBoxResult.Yes) Then
            getfile(athletes)
            genBibVal(athletes)
        Else
            If (runDb() = True) Then 'only executes if connection is successful
                Try
                    fetchData(athletes)
                Catch ex As Exception
                    MsgBox("Unsuccessful")
                End Try
                genBibVal(athletes)
            End If
        End If

        For counter = 0 To athletes.Length - 1
            If athletes(counter).amountJumps = max Then
                MsgBox(athletes(counter).fName + " " + athletes(counter).lName)
            End If
        Next

        max = findMax(athletes) ' finds max amount of jumping jacks done

    End Sub

    Function getfile(ByRef athletes() As athlete) As String
        Dim fd As OpenFileDialog = New OpenFileDialog()
        Dim no_of_records As Integer = 0

        'sets some basic info for the dialog to use
        fd.Title = "Select csv file"
        fd.InitialDirectory = "C:\"
        fd.Filter = "DB Files|*.csv"

        'fills array of objects
        If fd.ShowDialog() = DialogResult.OK Then
            Using csvparser As New FileIO.TextFieldParser(fd.FileName)

                csvparser.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited
                csvparser.Delimiters = New String() {","}
                Dim temp As String()

                While Not csvparser.EndOfData
                    athletes(no_of_records) = New athlete
                    temp = csvparser.ReadFields()

                    athletes(no_of_records).entryId = temp(0)
                    athletes(no_of_records).location = temp(1)
                    athletes(no_of_records).fName = temp(2)
                    athletes(no_of_records).lName = temp(3)
                    athletes(no_of_records).amountJumps = temp(4)

                    no_of_records += 1
                End While
            End Using
            'outputs data filled
            For loop_counter = 0 To no_of_records - 1
                lstoutput.Items.Add(athletes(loop_counter).entryId)
                lstoutput.Items.Add(athletes(loop_counter).location)
                lstoutput.Items.Add(athletes(loop_counter).fName)
                lstoutput.Items.Add(athletes(loop_counter).lName)
                lstoutput.Items.Add(athletes(loop_counter).amountJumps)
            Next
            FileClose()
        End If
        Return fd.FileName
    End Function

    Sub genBibVal(ByRef athletes() As athlete)
        Dim fileAcc As System.IO.StreamWriter
        Dim fd As SaveFileDialog = New SaveFileDialog()
        'sets basic file dialog to follow
        fd.Title = "Save Bib Values"
        fd.InitialDirectory = "C:\"
        fd.Filter = "DB Files|*.csv"

        Dim file_access As System.IO.StreamWriter
        Dim FrstFore(29) As String
        Dim AsciiVal(29) As Integer

        'gens values to go on the bib and saves to a csv
        If fd.ShowDialog() = DialogResult.OK Then
            file_access = My.Computer.FileSystem.OpenTextFileWriter(fd.FileName, True)
            For counter = 0 To 29
                FrstFore(counter) = Mid(athletes(counter).fName, 1, 1)
                AsciiVal(counter) = Asc(athletes(counter).location)
            Next
            For counter = 0 To 29
                file_access.WriteLine(athletes(counter).entryId & "," & FrstFore(counter) & athletes(counter).lName & AsciiVal(counter))
            Next
            file_access.Close()
        End If
    End Sub

    Function findMax(ByRef athletes() As athlete) As Integer
        Dim max = athletes(0).amountJumps
        For counter = 1 To 29
            If (athletes(counter).amountJumps > max) Then
                max = athletes(counter).amountJumps
            End If
        Next
        Return max
    End Function

    Function runDb() As Boolean
        'this function is to check the connection to our db
        Try
            con.Open()
            If con.State = ConnectionState.Open Then
                MsgBox("Database Connected!")
                Return True
            Else
                MsgBox("Error")
                Return False
            End If
        Catch ex As Exception
            MsgBox("Catastrophic failure check log file in documents")
            Dim filePath As String
            filePath = System.IO.Path.Combine(My.Computer.FileSystem.SpecialDirectories.MyDocuments, "error.log")
            My.Computer.FileSystem.WriteAllText(filePath, ex.ToString(), True)
        End Try

        con.Close()
    End Function

    Sub fetchData(ByRef athletes() As athlete)
        Dim entryCmd As MySqlCommand
        Dim locationCmd As MySqlCommand
        Dim fNameCmd As MySqlCommand
        Dim lNameCmd As MySqlCommand
        Dim cmd As MySqlCommand
        Dim numJumpsCmd As MySqlCommand
        Dim counter As Integer = 1
        Dim dbLen As Integer

        con.Open()
        cmd = New MySqlCommand("SELECT COUNT(*) FROM athletes", con)
        dbLen = cmd.ExecuteScalar()

        entryCmd = New MySqlCommand("Select entryId from athletes WHERE ID = (@val1)", con)
        entryCmd.Parameters.AddWithValue("@val1", counter)

        locationCmd = New MySqlCommand("Select location from athletes WHERE ID = (@val1)", con)
        locationCmd.Parameters.AddWithValue("@val1", counter)

        fNameCmd = New MySqlCommand("Select fName from athletes WHERE ID = (@val1)", con)
        fNameCmd.Parameters.AddWithValue("@val1", counter)

        lNameCmd = New MySqlCommand("Select lName from athletes WHERE ID = (@val1)", con)
        lNameCmd.Parameters.AddWithValue("@val1", counter)

        numJumpsCmd = New MySqlCommand("Select amount jumps from athletes WHERE ID = (@val1)", con)
        numJumpsCmd.Parameters.AddWithValue("@val1", counter)

        For counter = 1 To dbLen
            athletes(counter) = New athlete
            athletes(counter).entryId = entryCmd.ExecuteScalar()
            athletes(counter).location = locationCmd.ExecuteScalar()
            athletes(counter).fName = fNameCmd.ExecuteScalar()
            athletes(counter).lName = lNameCmd.ExecuteScalar()
            athletes(counter).amountJumps = numJumpsCmd.ExecuteScalar()
        Next
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.CheckForIllegalCrossThreadCalls = False
    End Sub
End Class
