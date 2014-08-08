Imports System.IO
Imports System.Threading
Imports System.ComponentModel
Imports System.Text
Imports System.Security.Cryptography



Public Class Main_Screen

    Private busyworking As Boolean = False

    Private lastinputline As String = ""
    Private inputlines As Long = 0
    Private highestPercentageReached As Integer = 0
    Private inputlinesprecount As Long = 0
    Private pretestdone As Boolean = False
    Private primary_PercentComplete As Integer = 0
    Private percentComplete As Integer

    Private SelectedIndex As Integer = 0

    Private backupdirectory As String = ""
    Private savedirectory As String = ""

    Private AlertMessage As String = ""




    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            If ex.Message.IndexOf("Thread was being aborted") < 0 Then
                Dim Display_Message1 As New Display_Message()
                Display_Message1.Message_Textbox.Text = "The Application encountered the following problem: " & vbCrLf & identifier_msg & ": " & ex.Message.ToString

                Display_Message1.Timer1.Interval = 1000
                Display_Message1.ShowDialog()
                Dim dir As System.IO.DirectoryInfo = New System.IO.DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
                If dir.Exists = False Then
                    dir.Create()
                End If
                dir = Nothing
                Dim filewriter As System.IO.StreamWriter = New System.IO.StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
                filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & " - " & identifier_msg & ": " & ex.ToString)
                filewriter.WriteLine("")
                filewriter.Flush()
                filewriter.Close()
                filewriter = Nothing
            End If
            ex = Nothing
            identifier_msg = Nothing
        Catch exc As Exception
            MsgBox("An error occurred in the application's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub




   



    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim result As DialogResult
        result = FolderBrowserDialog1.ShowDialog
        If result = Windows.Forms.DialogResult.OK Then
            TextBox1.Text = FolderBrowserDialog1.SelectedPath
        End If
    End Sub


    


    Private Sub cancelAsyncButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cancelAsyncButton.Click

        ' Cancel the asynchronous operation.
        Me.BackgroundWorker1.CancelAsync()

        ' Disable the Cancel button.
        cancelAsyncButton.Enabled = False
        sender = Nothing
        e = Nothing
    End Sub 'cancelAsyncButton_Click

    Private Sub PreCount_Function(ByVal worker As BackgroundWorker)
        Try
            inputlinesprecount = 0
            inputlines = 0
            Dim dinfo As DirectoryInfo
            dinfo = New DirectoryInfo(TextBox1.Text)
            'Dim backupfolder As String = (Application.StartupPath & "\").Replace("\\", "\") & "WP7D Backup " & Format(Now, "yyyyMMddHHmmss")
            'backupdirectory = backupfolder
            'If My.Computer.FileSystem.DirectoryExists(backupfolder) = False Then
            '    My.Computer.FileSystem.CreateDirectory(backupfolder)
            'End If

            For Each finfo As FileInfo In dinfo.GetFiles
                'If My.Computer.FileSystem.FileExists((finfo.FullName & "\Build.txt").Replace("\\", "\")) Then
                '    Dim mfinfo As FileInfo = New FileInfo((finfo.FullName & "\Build.txt").Replace("\\", "\"))
                '    mfinfo.CopyTo((backupfolder & "\" & finfo.Name & " - Build.txt").Replace("\\", "\"))
                '    lastinputline = "Backed up: " & mfinfo.Name
                'Else
                '    AlertMessage = AlertMessage & "Missing Build.txt File: " & finfo.Name & vbCrLf
                'End If
                inputlinesprecount = inputlinesprecount + 1
                inputlines = inputlines + 1
                worker.ReportProgress(0)
                finfo = Nothing
            Next

            'If inputlinesprecount < 1 Then
            '    My.Computer.FileSystem.DeleteDirectory(backupfolder, FileIO.DeleteDirectoryOption.DeleteAllContents)
            'End If

        Catch ex As Exception
            Error_Handler(ex, "PreCount_Function")
        End Try
    End Sub

    Private Sub startAsyncButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles startAsyncButton.Click
        Try
            If busyworking = False Then
                If My.Computer.FileSystem.DirectoryExists(TextBox1.Text) Then
                    If My.Computer.FileSystem.FileExists(TextBox2.Text) Then




                        busyworking = True


                        inputlines = 0
                        lastinputline = ""
                        highestPercentageReached = 0
                        inputlinesprecount = 0

                        backupdirectory = ""
                        savedirectory = ""
                        pretestdone = False

                        TextBox1.Enabled = False
                        Button1.Enabled = False
                        TextBox2.Enabled = False
                        Button2.Enabled = False
                        startAsyncButton.Enabled = False
                        cancelAsyncButton.Enabled = True
                        ' Start the asynchronous operation.
                        AlertMessage = ""

                        BackgroundWorker1.RunWorkerAsync(TextBox1.Text)
                    Else
                        MsgBox("Please ensure that you select an existing CSSR database to process", MsgBoxStyle.Information, "Invalid Database Selected")
                    End If
                Else
                    MsgBox("Please ensure that you select an existing directory to process", MsgBoxStyle.Information, "Invalid Directory Selected")
                End If
                End If
        Catch ex As Exception
            Error_Handler(ex, "StartWorker")
        End Try
    End Sub

    ' This event handler is where the actual work is done.
    Private Sub backgroundWorker1_DoWork(ByVal sender As Object, ByVal e As DoWorkEventArgs) Handles BackgroundWorker1.DoWork

        ' Get the BackgroundWorker object that raised this event.
        Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)

        ' Assign the result of the computation
        ' to the Result property of the DoWorkEventArgs
        ' object. This is will be available to the 
        ' RunWorkerCompleted eventhandler.
        e.Result = MainWorkerFunction(worker, e)
        sender = Nothing
        e = Nothing
        worker.Dispose()
        worker = Nothing
    End Sub 'backgroundWorker1_DoWork

    ' This event handler deals with the results of the
    ' background operation.
    Private Sub backgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        busyworking = False


        ' First, handle the case where an exception was thrown.
        If Not (e.Error Is Nothing) Then
            Error_Handler(e.Error, "backgroundWorker1_RunWorkerCompleted")
        ElseIf e.Cancelled Then
            ' Next, handle the case where the user canceled the 
            ' operation.
            ' Note that due to a race condition in 
            ' the DoWork event handler, the Cancelled
            ' flag may not have been set, even though
            ' CancelAsync was called.
            Me.ToolStripStatusLabel1.Text = "Operation Cancelled" & "   (" & inputlines & " of " & inputlinesprecount & ")"
            Me.ProgressBar1.Value = 0

        Else
            ' Finally, handle the case where the operation succeeded.
            Me.ToolStripStatusLabel1.Text = "Operation Completed" & "   (" & inputlines & " of " & inputlinesprecount & ")"
            Me.ProgressBar1.Value = 100
            If AlertMessage.Length > 0 Then
                'MsgBox("The following alerts were raised during the operation. If you wish to save these alerts, press Ctrl+C and paste it into NotePad." & vbCrLf & vbCrLf & "********************" & vbCrLf & vbCrLf & AlertMessage, MsgBoxStyle.Information, "Raised Alerts")
                MsgBox("The following files appear to have been ignored: " & vbCrLf & vbCrLf & AlertMessage, MsgBoxStyle.Information, "Files Ignored")
            End If
        End If

        TextBox1.Enabled = True
        Button1.Enabled = True
        TextBox2.Enabled = True
        Button2.Enabled = True
        startAsyncButton.Enabled = True
        cancelAsyncButton.Enabled = False

        sender = Nothing
        e = Nothing


    End Sub 'backgroundWorker1_RunWorkerCompleted

    Private Sub backgroundWorker1_ProgressChanged(ByVal sender As Object, ByVal e As ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged


        Me.ProgressBar1.Value = e.ProgressPercentage
        'If lastinputline.StartsWith("Operation Completed") Then
        'Me.ToolStripStatusLabel1.Text = lastinputline
        'Else
        Me.ToolStripStatusLabel1.Text = lastinputline & "   (" & inputlines & " of " & inputlinesprecount & ")"
        'End If


        sender = Nothing
        e = Nothing
    End Sub

    Function MainWorkerFunction(ByVal worker As BackgroundWorker, ByVal e As DoWorkEventArgs) As Boolean
        Dim result As Boolean = False
        Try
            If Me.pretestdone = False Then
                primary_PercentComplete = 0
                worker.ReportProgress(0)
                PreCount_Function(worker)
                Me.pretestdone = True
            End If

            If worker.CancellationPending Then
                e.Cancel = True
                Return False
            End If

            primary_PercentComplete = 0
            worker.ReportProgress(0)

            inputlines = 0
            lastinputline = ""


            Dim dinfo As DirectoryInfo
            dinfo = New DirectoryInfo(TextBox1.Text)

            If dinfo.Exists Then
                For Each subdir As FileInfo In dinfo.GetFiles
                    Try
                        lastinputline = "Processing: " & subdir.Name
                        ' Report progress as a percentage of the total task.
                        percentComplete = 0
                        If inputlinesprecount > 0 Then
                            percentComplete = CSng(inputlines) / CSng(inputlinesprecount) * 100
                        Else
                            percentComplete = 100
                        End If
                        primary_PercentComplete = percentComplete
                        If percentComplete > 100 Then
                            percentComplete = 100
                        End If
                        If percentComplete = 100 Then
                            lastinputline = "Operation Completed"
                        End If
                        If percentComplete > highestPercentageReached Then
                            highestPercentageReached = percentComplete
                            worker.ReportProgress(percentComplete)
                        End If
                        If worker.CancellationPending Then
                            e.Cancel = True
                            Exit For
                            Return False
                        End If


                        Dim reader As StreamReader = My.Computer.FileSystem.OpenTextFileReader(subdir.FullName)
                        Dim lineread, origlineread As String
                        Dim WP_ID As Boolean = False
                        Dim WP_ID_value As String = ""
                        Dim WP_Title As Boolean = False
                        Dim WP_Author As Boolean = False
                        Dim WP_PublicationDate As Boolean = False
                        Dim WP_Price As Boolean = False
                        Dim WP_Abstract As Boolean = False
                        Dim WP_Filename As Boolean = False

                        Dim ReadNewLine As Boolean = True

                        While reader.Peek <> -1
                            If ReadNewLine = True Then
                                lineread = reader.ReadLine.Replace("&nbsp;", " ").Replace("'", "&acute;").Replace("""", "&acute;&acute;")
                            End If
                            ReadNewLine = True

                            origlineread = lineread
                            If lineread.ToLower.IndexOf("<h1><b>") <> -1 Then
                                If lineread.ToLower.IndexOf("working paper") <> -1 Then
                                    If WP_ID = False Then
                                        lineread = StripHTMLTags(lineread)
                                        Dim tokens As String() = lineread.Split(" ")
                                        For Each token As String In tokens
                                            If token.IndexOf("/") = 2 Then
                                                token = StripHTMLTags(token.Trim.ToUpper)
                                                If token.Length < 6 Then
                                                    token = token.Replace("/", "/0")
                                                End If

                                                If (WorkerExecuteSpecialScalar("Select count(WP_ID) as Counted from WorkingPapers where WP_ID = '" & token & "'", 1)) = "0" Then
                                                    WorkerExecuteNonQuery("Insert into WorkingPapers (WP_ID) values ('" & token & "')", 1)
                                                End If
                                                WP_ID_value = token
                                                WP_ID = True

                                            End If
                                        Next
                                    End If
                                End If
                            End If
                            lineread = origlineread
                            If WP_ID = True Then
                                If WP_Title = False Then
                                    If lineread.ToLower.IndexOf("<a href=&acute;&acute;papers/") <> -1 And lineread.ToLower.IndexOf("download paper") = -1 Then
                                        lineread = StripHTMLTags(lineread)
                                        'MsgBox("Update WorkingPapers set WP_Title = '" & lineread & "' where WP_ID = '" & WP_ID_value & "'")
                                        WorkerExecuteNonQuery("Update WorkingPapers set WP_Title = """ & lineread.Replace("Title:", "").Replace("""", "'").Replace("'", "`").Trim & """ where WP_ID = """ & WP_ID_value & """", 1)
                                        WP_Title = True
                                    End If
                                End If
                                lineread = origlineread
                                If WP_Filename = False Then
                                    If lineread.ToLower.IndexOf("<a href=&acute;&acute;papers/") <> -1 Then
                                        lineread = lineread.Remove(0, lineread.IndexOf("<a href=&acute;&acute;papers/"))
                                        lineread = lineread.Remove(0, 29)
                                        lineread = lineread.Substring(0, lineread.IndexOf("&acute;&acute;"))

                                        WorkerExecuteNonQuery("Update WorkingPapers set WP_Filename = """ & lineread.Trim & """ where WP_ID = """ & WP_ID_value & """", 1)
                                        WP_Filename = True
                                    End If
                                End If
                                lineread = origlineread
                                If WP_Author = False Then
                                    lineread = StripHTMLTags(lineread)
                                    If lineread.ToLower.IndexOf("author") <> -1 Or lineread.ToLower.IndexOf("author:") <> -1 Or lineread.ToLower.IndexOf("author(s):") <> -1 Or lineread.ToLower.IndexOf("authors:") <> -1 Then
                                        lineread = lineread.Replace("Author(s):", "").Trim
                                        lineread = lineread.Replace("Authors:", "").Trim
                                        lineread = lineread.Replace("Author:", "").Trim
                                        lineread = lineread.Replace("Author", "").Trim
                                        lineread = lineread.Replace("(s)", "").Trim
                                        If lineread.Length < 1 Then
                                            origlineread = reader.ReadLine.Replace("&nbsp;", " ").Replace("'", "&acute;").Replace("""", "&acute;&acute;")
                                            ReadNewLine = False
                                            lineread = origlineread
                                            lineread = StripHTMLTags(lineread).Trim
                                        End If
                                        lineread = lineread.Replace("Author(s):", "").Trim
                                        lineread = lineread.Replace("Authors:", "").Trim
                                        lineread = lineread.Replace("Author:", "").Trim
                                        lineread = lineread.Replace("Author", "").Trim
                                        lineread = lineread.Replace("(s)", "").Trim
                                        If lineread.IndexOf("Date of Publication") <> -1 Then
                                            lineread = ""
                                        End If

                                        WorkerExecuteNonQuery("Update WorkingPapers set WP_Author = """ & lineread.Trim & """ where WP_ID = """ & WP_ID_value & """", 1)
                                        WP_Author = True
                                    End If
                                End If
                                lineread = origlineread
                                If WP_Price = False Then
                                    lineread = StripHTMLTags(lineread)
                                    If lineread.ToLower.IndexOf("price") <> -1 Or lineread.ToLower.IndexOf("price:") <> -1 Then
                                        lineread = lineread.Replace("Price:", "").Trim
                                        lineread = lineread.Replace("Price", "").Trim
                                        If lineread.Length < 1 Then
                                            origlineread = reader.ReadLine.Replace("&nbsp;", " ").Replace("'", "&acute;").Replace("""", "&acute;&acute;")
                                            ReadNewLine = False
                                            lineread = origlineread
                                            lineread = StripHTMLTags(lineread).Trim
                                        End If
                                        lineread = lineread.Replace("Price:", "").Trim
                                        lineread = lineread.Replace("Price", "").Trim
                                        If lineread.IndexOf("Abstract") <> -1 Then
                                            lineread = ""
                                        End If
                                        WorkerExecuteNonQuery("Update WorkingPapers set WP_Price = """ & lineread.Trim & """ where WP_ID = """ & WP_ID_value & """", 1)
                                        WP_Price = True
                                    End If
                                End If

                                lineread = origlineread
                                If WP_PublicationDate = False Then
                                    lineread = StripHTMLTags(lineread)
                                    If lineread.ToLower.IndexOf("date of publication") <> -1 Or lineread.ToLower.IndexOf("date of publication:") <> -1 Then
                                        lineread = lineread.Replace("Date of Publication:", "").Trim
                                        lineread = lineread.Replace("Date of Publication", "").Trim
                                        If lineread.Length < 1 Then
                                            origlineread = reader.ReadLine.Replace("&nbsp;", " ").Replace("'", "&acute;").Replace("""", "&acute;&acute;")
                                            ReadNewLine = False
                                            lineread = origlineread
                                            lineread = StripHTMLTags(lineread).Trim
                                        End If

                                        If lineread.IndexOf("Price") <> -1 Then
                                            lineread = ""
                                        End If
                                        lineread = lineread.Replace("Date of Publication:", "").Trim
                                        lineread = lineread.Replace("Date of Publication", "").Trim

                                        WorkerExecuteNonQuery("Update WorkingPapers set WP_PublicationDate = """ & lineread.Trim & """ where WP_ID = """ & WP_ID_value & """", 1)
                                        WP_PublicationDate = True
                                    End If
                                End If

                                lineread = origlineread
                                If WP_Abstract = False Then
                                    lineread = StripHTMLTags(lineread)
                                    If lineread.ToLower.IndexOf("abstract") <> -1 Or lineread.ToLower.IndexOf("abstract:") <> -1 Or lineread.ToLower.IndexOf("extract") <> -1 Or lineread.ToLower.IndexOf("extract:") <> -1 Then
                                        Dim abstract As String = ""
                                        lineread = lineread.Replace("Abstract:", "").Trim
                                        lineread = lineread.Replace("Abstract", "").Trim
                                        lineread = lineread.Replace("Extract:", "").Trim
                                        lineread = lineread.Replace("Extract", "").Trim
                                        lineread = lineread & reader.ReadLine.Replace("&nbsp;", " ").Replace("'", "&acute;").Replace("""", "&acute;&acute;") & " "
                                        lineread = lineread & reader.ReadLine.Replace("&nbsp;", " ").Replace("'", "&acute;").Replace("""", "&acute;&acute;") & " "
                                        lineread = StripHTMLTags(lineread).Trim
                                        While lineread.Length > 0 Or lineread.IndexOf("Download Paper") <> -1
                                            lineread = lineread.Replace("Abstract:", "").Trim
                                            lineread = lineread.Replace("Abstract", "").Trim
                                            abstract = abstract & lineread & " "
                                            origlineread = reader.ReadLine.Replace("&nbsp;", " ").Replace("'", "&acute;").Replace("""", "&acute;&acute;")
                                            lineread = origlineread
                                            lineread = StripHTMLTags(lineread).Trim
                                        End While


                                        WorkerExecuteNonQuery("Update WorkingPapers set WP_Abstract = """ & abstract & """ where WP_ID = """ & WP_ID_value & """", 1)
                                        WP_Abstract = True
                                    End If
                                End If
                            End If
                        End While
                        reader.Close()
                        reader = Nothing

                       
                        If WP_ID = False Then
                            AlertMessage = AlertMessage & subdir.FullName & vbCrLf
                        End If


                        inputlines = inputlines + 1
                        lastinputline = "Processed: " & subdir.Name
                        ' Report progress as a percentage of the total task.
                        percentComplete = 0
                        If inputlinesprecount > 0 Then
                            percentComplete = CSng(inputlines) / CSng(inputlinesprecount) * 100
                        Else
                            percentComplete = 100
                        End If
                        primary_PercentComplete = percentComplete
                        If percentComplete > 100 Then
                            percentComplete = 100
                        End If
                        If percentComplete = 100 Then
                            lastinputline = "Operation Completed"
                        End If
                        If percentComplete > highestPercentageReached Then
                            highestPercentageReached = percentComplete
                            worker.ReportProgress(percentComplete)
                        End If
                        subdir = Nothing
                        If worker.CancellationPending Then
                            e.Cancel = True
                            Exit For
                            dinfo = Nothing
                            Return False
                        End If
                    Catch ex As Exception
                        Error_Handler(ex, "Parsing HTML File")
                    End Try

                Next
            End If
            dinfo = Nothing




        Catch ex As Exception
            Error_Handler(ex, "MainWorkerFunction")
        End Try
        worker.Dispose()
        worker = Nothing
        e = Nothing
        Return result

    End Function

    Private Sub Form1_Close(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        Try
            Me.ToolStripStatusLabel1.Text = "Application Closing"
            SaveSettings()
        Catch ex As Exception
            Error_Handler(ex, "Application Close")
        End Try
    End Sub

    Private Sub LoadSettings()
        Try
            Dim configfile As String = (Application.StartupPath & "\config.sav").Replace("\\", "\")
            If My.Computer.FileSystem.FileExists(configfile) Then
                Dim reader As StreamReader = New StreamReader(configfile)
                Dim lineread As String
                Dim variablevalue As String
                While reader.Peek <> -1
                    lineread = reader.ReadLine
                    If lineread.IndexOf("=") <> -1 Then

                        variablevalue = lineread.Remove(0, lineread.IndexOf("=") + 1)

                        If lineread.StartsWith("SelectFolder=") Then
                            Dim dinfo As DirectoryInfo = New DirectoryInfo(variablevalue)
                            If dinfo.Exists Then
                                FolderBrowserDialog1.SelectedPath = variablevalue
                                TextBox1.Text = variablevalue
                            End If
                            dinfo = Nothing
                        End If
                        If lineread.StartsWith("SelectDatabase=") Then
                            Dim dinfo As FileInfo = New FileInfo(variablevalue)
                            If dinfo.Exists Then
                                OpenFileDialog1.FileName = variablevalue
                                TextBox2.Text = variablevalue
                            End If
                            dinfo = Nothing
                        End If

                        'If lineread.StartsWith("SetVariable=") Then
                        '    ComboBox1.SelectedIndex = variablevalue
                        'End If

                        'If lineread.StartsWith("PixelValue=") Then
                        '    NumericUpDown2.Value = variablevalue
                        'End If

                    End If
                End While
                reader.Close()
                reader = Nothing
            End If
        Catch ex As Exception
            Error_Handler(ex, "Load Settings")
        End Try
    End Sub

    Private Sub SaveSettings()
        Try
            Dim configfile As String = (Application.StartupPath & "\config.sav").Replace("\\", "\")

            Dim writer As StreamWriter = New StreamWriter(configfile, False)

            If TextBox1.Text.Length > 0 Then
                Dim dinfo As DirectoryInfo = New DirectoryInfo(TextBox1.Text)
                If dinfo.Exists Then
                    writer.WriteLine("SelectFolder=" & TextBox1.Text)
                End If
                dinfo = Nothing
            End If
            If TextBox2.Text.Length > 0 Then
                Dim dinfo As FileInfo = New FileInfo(TextBox2.Text)
                If dinfo.Exists Then
                    writer.WriteLine("SelectDatabase=" & TextBox2.Text)
                End If
                dinfo = Nothing
            End If
            'If ComboBox1.SelectedIndex <> -1 Then
            '    writer.WriteLine("SetVariable=" & ComboBox1.SelectedIndex)
            'End If

            'writer.WriteLine("PixelValue=" & NumericUpDown2.Value)

            writer.Flush()
            writer.Close()
            writer = Nothing

        Catch ex As Exception
            Error_Handler(ex, "Save Settings")
        End Try
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.Text = My.Application.Info.ProductName & " " & Format(My.Application.Info.Version.Major, "0000") & Format(My.Application.Info.Version.Minor, "00") & Format(My.Application.Info.Version.Build, "00") & "." & Format(My.Application.Info.Version.Revision, "00") & ""
            LoadSettings()
            Me.ToolStripStatusLabel1.Text = "Application Loaded"
        Catch ex As Exception
            Error_Handler(ex, "Application Load")
        End Try

    End Sub



    

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        Try
            Me.ToolStripStatusLabel1.Text = "About displayed"
            AboutBox1.ShowDialog()
        Catch ex As Exception
            Error_Handler(ex, "Display About Screen")
        End Try
    End Sub

    Private Sub HelpToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HelpToolStripMenuItem.Click
        Try
            Me.ToolStripStatusLabel1.Text = "Help displayed"
            HelpBox1.ShowDialog()
        Catch ex As Exception
            Error_Handler(ex, "Display Help Screen")
        End Try
    End Sub

    Private Sub TextBox1_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles TextBox1.DragDrop
        Try
            If e.Data.GetDataPresent(DataFormats.FileDrop) Then
                Dim MyFiles() As String
                Dim i As Integer

                ' Assign the files to an array.
                MyFiles = e.Data.GetData(DataFormats.FileDrop)
                ' Loop through the array and add the files to the list.
                'For i = 0 To MyFiles.Length - 1
                If MyFiles.Length > 0 Then
                    Dim finfo As DirectoryInfo = New DirectoryInfo(MyFiles(0))
                    If finfo.Exists = True Then
                        TextBox1.Text = (MyFiles(0))
                        FolderBrowserDialog1.SelectedPath = (MyFiles(0))
                    End If
                End If
                'Next
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub TextBox1_DragEnter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles TextBox1.DragEnter
        Try
            If e.Data.GetDataPresent(DataFormats.FileDrop) Then
                e.Effect = DragDropEffects.All
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub TextBox2_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles TextBox2.DragDrop
        Try
            If e.Data.GetDataPresent(DataFormats.FileDrop) Then
                Dim MyFiles() As String
                Dim i As Integer

                ' Assign the files to an array.
                MyFiles = e.Data.GetData(DataFormats.FileDrop)
                ' Loop through the array and add the files to the list.
                'For i = 0 To MyFiles.Length - 1
                If MyFiles.Length > 0 Then
                    Dim finfo As FileInfo = New FileInfo(MyFiles(0))
                    If finfo.Exists = True Then
                        TextBox2.Text = (MyFiles(0))
                        OpenFileDialog1.FileName = (MyFiles(0))
                    End If
                End If
                'Next
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub TextBox2_DragEnter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles TextBox2.DragEnter
        Try
            If e.Data.GetDataPresent(DataFormats.FileDrop) Then
                e.Effect = DragDropEffects.All
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub


    Protected Friend Function Get_Connection(ByVal dbselect As Integer) As OleDb.OleDbConnection
        'Standard(Security)
        '"Provider=sqloledb;Data Source=Aron1;Initial Catalog=pubs;User Id=sa;Password=asdasd;" 

        'Trusted(Connection)
        '"Provider=sqloledb;Data Source=Aron1;Initial Catalog=pubs;Integrated Security=SSPI;" 
        '(use serverName\instanceName as Data Source to use an specifik SQLServer instance, only SQLServer2000)

        'Prompt for username and password:
        'oConn.Provider = "sqloledb"
        'oConn.Properties("Prompt") = adPromptAlways
        'oConn.Open("Data Source=Aron1;Initial Catalog=pubs;")

        'Connect via an IP address:
        '"Provider=sqloledb;Data Source=190.190.200.100,1433;Network Library=DBMSSOCN;Initial Catalog=pubs;User ID=sa;Password=asdasd;" 
        '(DBMSSOCN=TCP/IP instead of Named Pipes, at the end of the Data Source is the port to use (1433 is the default))

        Dim connection_string As String
        'If dbserver.IndexOf(".") = -1 Then
        'connection_string = "Provider=sqloledb;Data Source=" & dbserver & ";Initial Catalog=" & dbtable & ";User Id=" & dbuser & ";Password=" & dbpassword & ";"
        'Else
        'connection_string = "Provider=sqloledb;Data Source=" & dbserver & ",1433;Network Library=DBMSSOCN;Initial Catalog=" & dbtable & ";User Id=" & dbuser & ";Password=" & dbpassword & ";"
        'End If
        'Dim connection_string As String = "User ID=" & dbuser & ";password=" & dbpassword & ";Data Source=" & dbserver & ";Tag with column collation when possible=False;Initial Catalog=" & dbtable & ";Use Procedure for Prepare=1;Auto Translate=True;Persist Security Info=False;Provider=""SQLOLEDB.1"";Use Encryption for Data=False;Packet Size=4096"
        Select Case dbselect
            Case 1
                connection_string = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=""" & TextBox2.Text & """"
            Case 2
                connection_string = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=""" & TextBox2.Text & """"
        End Select

        Dim conn As OleDb.OleDbConnection = New OleDb.OleDbConnection(connection_string)
        Return conn
    End Function



    Public Function WorkerExecuteNonQuery(ByVal SQLstatement As String, ByVal dbselect As Integer) As String
        Dim result As String
        Try

            Dim conn As OleDb.OleDbConnection
            Try
                conn = Get_Connection(dbselect)
                conn.Open()

                Dim sql As OleDb.OleDbCommand
                sql = New OleDb.OleDbCommand
                sql.CommandText = SQLstatement
                sql.Connection = conn
                result = sql.ExecuteNonQuery().ToString & " SQL Statement Succeeded"
                sql.Dispose()

            Catch ex As Exception
                Error_Handler(ex)
                result = "0 SQL Statement Failed"
            Finally
                Try
                    conn.Close()
                Catch ex1 As Exception
                    Error_Handler(ex1)
                End Try
                conn.Dispose()
            End Try
        Catch ex As Exception
            Error_Handler(ex)
            result = "0 SQL Statement Failed"


        End Try
        Return result
    End Function



    Public Function WorkerExecuteSpecialScalar(ByVal SQLstatement As String, ByVal dbselect As Integer) As String
        Dim result As String
        Try

            Dim conn As OleDb.OleDbConnection
            Try
                conn = Get_Connection(dbselect)

                conn.Open()



                Dim sql As OleDb.OleDbCommand = New OleDb.OleDbCommand
                sql.CommandText = SQLstatement
                sql.Connection = conn

                result = sql.ExecuteScalar().ToString
                sql.Dispose()

            Catch ex As Exception
                Error_Handler(ex)
                result = "0 SQL Statement Failed"
            Finally
                Try
                    conn.Close()
                Catch ex1 As Exception
                    Error_Handler(ex1)
                End Try
                conn.Dispose()
            End Try
        Catch ex As Exception
            Error_Handler(ex)
            result = "0 SQL Statement Failed"


        End Try
        Return result
    End Function

    Private Function StripHTMLTags(ByVal input As String) As String
        Dim output As String = ""
        Try
            Dim firstopen, firstclose As Integer
            output = input
            output = output.Replace("&nbsp;", " ")
            While output.IndexOf("<") <> -1 Or output.IndexOf(">") <> -1
                firstopen = -1
                firstclose = -1
                firstopen = output.IndexOf("<")
                firstclose = output.IndexOf(">")
                If firstopen <> -1 And firstclose <> -1 Then
                    If firstopen < firstclose Then
                        output = output.Remove(firstopen, firstclose - firstopen + 1)
                    Else
                        output = output.Remove(0, firstclose + 1)
                    End If
                End If
                If firstopen = -1 And firstclose = -1 Then
                    output = output
                End If
                If firstopen <> -1 And firstclose = -1 Then
                    output = output.Remove(firstopen, output.Length - firstopen)
                End If
                If firstopen = -1 And firstclose <> -1 Then
                    output = output.Remove(0, firstclose + 1)
                End If
            End While
        Catch ex As Exception
            Error_Handler(ex, "Strip HTML Tags (Error on '" & output & "')")
        End Try
        Return output
    End Function

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim result As DialogResult
        result = OpenFileDialog1.ShowDialog
        If result = Windows.Forms.DialogResult.OK Then
            TextBox2.Text = OpenFileDialog1.FileName
        End If
    End Sub
End Class
