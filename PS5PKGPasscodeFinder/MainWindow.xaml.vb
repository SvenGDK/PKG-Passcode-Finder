Imports System.Collections.Concurrent
Imports System.Text
Imports System.Threading
Imports System.Windows.Forms

Class MainWindow

    Private Tries As Long = 0
    Private PasscodeFound As Boolean = False
    Private Shared ReadOnly NewSyncLock As New Object()
    Private Shared ReadOnly NewRandom As New Random()
    Private TestedPasscodes As New ConcurrentDictionary(Of String, Boolean)()

    Sub Main()

    End Sub

    Private Sub StartButton_Click(sender As Object, e As RoutedEventArgs) Handles StartButton.Click
        Dim StartCount As Integer = Environment.ProcessorCount * 2
        Dim PasscodeCharacters As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
        Dim PasscodeLength As Integer = 31

        Task.Run(Sub()
                     Parallel.For(0, startCount, Sub(i)
                                                     FindPasscode(PasscodeCharacters, PasscodeLength)
                                                 End Sub)
                 End Sub)
    End Sub

    Private Sub FindPasscode(Characters As String, PasscodeLength As Integer)
        While True
            If PasscodeFound = False Then
                Dim passcode As String = GenerateRandomPasscode(Characters, PasscodeLength)
                If TestedPasscodes.TryAdd(passcode, True) Then
                    Dispatcher.Invoke(Sub()
                                          Console.WriteLine("Passcodes tested: " + Tries.ToString() + vbCrLf)
                                          Console.WriteLine("Trying passcode: " + passcode)
                                          StartApplication(passcode)
                                          Interlocked.Increment(Tries)
                                      End Sub)
                End If
            Else
                Exit While
            End If
        End While

        MsgBox("Passcode found!", MsgBoxStyle.Information)
    End Sub

    Private Function GenerateRandomPasscode(Characters As String, PasscodeLength As Integer) As String
        Dim NewPasscode As New StringBuilder(PasscodeLength)
        SyncLock NewSyncLock
            For i As Integer = 1 To PasscodeLength
                Dim index As Integer = NewRandom.Next(Characters.Length)
                NewPasscode.Append(Characters(index))
            Next
        End SyncLock
        Return NewPasscode.ToString()
    End Function

    Private Sub StartApplication(passcode As String)
        Dim PubToolsPath As String = My.Computer.FileSystem.CurrentDirectory + "\prospero-pub-cmd.exe"
        Dim NewProcess As New Process()
        NewProcess.StartInfo.FileName = PubToolsPath
        NewProcess.StartInfo.Arguments = "img_extract --passcode " + passcode + " """ + SelectedPKGFIleTextBox.Text + """ """ + ExtractToTextBox.Text + """"
        NewProcess.StartInfo.UseShellExecute = False
        NewProcess.StartInfo.RedirectStandardOutput = True
        NewProcess.StartInfo.CreateNoWindow = True
        NewProcess.Start()

        Dim NewProcessOutput As String = NewProcess.StandardOutput.ReadToEnd()
        NewProcess.WaitForExit()

        If NewProcessOutput.Contains("[Error]") = False Then
            PasscodeFound = True
        End If
    End Sub

    Private Sub BrowsePKGButton_Click(sender As Object, e As RoutedEventArgs) Handles BrowsePKGButton.Click
        Dim OFD As New Forms.OpenFileDialog()
        If OFD.ShowDialog() = Forms.DialogResult.OK Then
            SelectedPKGFIleTextBox.Text = OFD.FileName
        End If
    End Sub

    Private Sub BrowseOutputPathButton_Click(sender As Object, e As RoutedEventArgs) Handles BrowseOutputPathButton.Click
        Dim FBD As New FolderBrowserDialog()
        If FBD.ShowDialog() = Forms.DialogResult.OK Then
            ExtractToTextBox.Text = FBD.SelectedPath
        End If
    End Sub

End Class
