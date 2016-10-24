Imports System.Data.SqlClient
Imports System.IO
Imports System.Net
Imports System.Reflection
Imports Ionic.Zip
Imports Microsoft.SqlServer.Management.Smo

Public Class Form1

    Private Sub DownloadFile(ByVal downloadPath As String, ByVal host As String, ByVal versionFile As String, ByVal userName As String, ByVal ftppassword As String)
        Dim myFtpWebRequest As FtpWebRequest
        myFtpWebRequest = FtpWebRequest.Create(versionFile)
        myFtpWebRequest.Credentials = New NetworkCredential(userName, ftppassword)
        myFtpWebRequest.Method = WebRequestMethods.Ftp.DownloadFile
        myFtpWebRequest.UseBinary = True

        Dim myFtpWebResponse As FtpWebResponse
        myFtpWebResponse = myFtpWebRequest.GetResponse()

        Dim myStreamWriter As StreamWriter
        myStreamWriter = New StreamWriter(downloadPath & "version.txt")
        myStreamWriter.Write(New StreamReader(myFtpWebResponse.GetResponseStream()).ReadToEnd)
        myStreamWriter.Close()
        Console.WriteLine("Version.txt" & myFtpWebResponse.StatusDescription)
        myFtpWebResponse.Close()

    End Sub

    Private Sub DownloadOvc(ByVal ovcDownloadPath As String, ByVal host As String, ByVal ovcFile As String, ByVal userName As String, ByVal ftppassword As String)
        Dim myFtpWebRequest As FtpWebRequest
        myFtpWebRequest = FtpWebRequest.Create(ovcFile)
        myFtpWebRequest.Credentials = New NetworkCredential(userName, ftppassword)
        myFtpWebRequest.Method = WebRequestMethods.Ftp.DownloadFile
        myFtpWebRequest.UseBinary = True

        Dim myFtpWebResponse As FtpWebResponse
        myFtpWebResponse = myFtpWebRequest.GetResponse()

        Dim myStreamWriter As StreamWriter
        myStreamWriter = New StreamWriter(ovcDownloadPath & "ovc.zip")
        myStreamWriter.Write(New StreamReader(myFtpWebResponse.GetResponseStream()).ReadToEnd)
        myStreamWriter.Close()
        Console.WriteLine("ovc.zip" & myFtpWebResponse.StatusDescription)
        myFtpWebResponse.Close()

    End Sub

    Public Function IsConnectionAvailable() As Boolean
        Dim objUrl As New Uri("http://www.youtube.com")
        Dim objWebReq As WebRequest
        objWebReq = WebRequest.Create(objUrl)
        Dim objresp As WebResponse

        Try
            objresp = objWebReq.GetResponse
            objresp.Close()
            objresp = Nothing
            Return True

        Catch ex As Exception
            objresp = Nothing
            objWebReq = Nothing
            Return False
        End Try
    End Function

    Public Shared Sub ExtractToDirectory(sourceArchiveFileName As String, destinationDirectoryName As String)
        Console.WriteLine("Extracting file {0} to {1}", sourceArchiveFileName, destinationDirectoryName)
        Using zip1 As ZipFile = ZipFile.Read(sourceArchiveFileName)

            'AddHandler zip1.ExtractProgress, AddressOf MyExtractProgress
            Dim e As ZipEntry
            ' here, we extract every entry  
            For Each e In zip1
                e.Extract(destinationDirectoryName, ExtractExistingFileAction.OverwriteSilently)
            Next

        End Using
    End Sub

    Public Shared Sub CopyAndReplace(sourceDir, destinationDir)

        Dim requiredFiles As String() = {"*.dll", "*.sql", "*.exe", "*.txt"}
        Dim requiredFolder As String = "TemplatesExcelReports"
        'define the variable which dictates wether we copy the file or not
        Dim copy As Boolean

        'File.Copy(requiredFolder, destinationDir)
        Debug.Print("===========Copying folder=============")
        My.Computer.FileSystem.CopyDirectory(sourceDir & "OVC\" & requiredFolder, destinationDir & "\" & requiredFolder, True)
        Debug.Print("=========== Folder copied=============")
        MsgBox("Folder copying successiful")
        ' For Each filename As String In Directory.GetFiles(sourceDir, "*.sql", SearchOption.TopDirectoryOnly)

        'Dim dirFiles As String() = Directory.GetFiles(sourceDir & "OVC\", "*.*", SearchOption.TopDirectoryOnly)
        ' Dim len As Int32 = dirFiles.Length


        'For j = 0 To dirFiles.Length - 1
        'For i = 0 To requiredFiles.Length - 1

        'If dirFiles(j) = requiredFiles(i) Then
        'File.Copy(dirFiles(j), destinationDir, True)
        ' MsgBox("Success")
        ' Else
        ' MsgBox("File not found for copying")
        ' End If
        'Next

        'Next
        'Dim searchPattern As String() = {"*.dll", "*.sql", "*.exe", "*.txt"}

        If (My.Computer.FileSystem.DirectoryExists(destinationDir)) Then
            ' For Each foundFile As String In My.Computer.FileSystem.GetFiles(sourceDir & "OVC\", _
            'FileIO.SearchOption.SearchTopLevelOnly, "*.*")
            For Each foundFile As String In Directory.GetFiles(sourceDir & "OVC\", "*.*", SearchOption.TopDirectoryOnly)
                Select Case LCase(Path.GetExtension(foundFile))
                    'Select Case LCase(foundFile)
                    Case ".gif"
                        ' My.Computer.FileSystem.CopyFile(foundFile, destinationDir & "\" & Path.GetFileName(foundFile), True)
                    Case ".ico"
                    Case ".pdp"
                    Case ".rar"
                    Case ".rar"
                    Case ".xlsm"
                    Case ".xml"
                    Case ".sql"
                        My.Computer.FileSystem.CopyFile(foundFile, destinationDir & "\" & Path.GetFileName(foundFile), True)
                    Case ".exe"
                        My.Computer.FileSystem.CopyFile(foundFile, destinationDir & "\" & Path.GetFileName(foundFile), True)
                    Case ".txt"
                        My.Computer.FileSystem.CopyFile(foundFile, destinationDir & "\" & Path.GetFileName(foundFile), True)
                        'My.Computer.FileSystem.CopyFile(foundFile, destinationDir & "\" & foundFile, True)

                    Case ".dll"
                        My.Computer.FileSystem.CopyFile(foundFile, destinationDir & "\" & Path.GetFileName(foundFile), True)
                    Case ".zip"
                        ' My.Computer.FileSystem.CopyFile(foundFile, destinationDir & "\" & Path.GetFileName(foundFile), True)

                        'File.Copy(foundFile, destinationDir, True)
                    Case Else
                        'My.Computer.FileSystem.CopyFile(foundFile, destinationDir, True)
                        'My.Computer.FileSystem.CopyFile(foundFile, destinationDir & Path.GetFileName(foundFile), True)
                        'MsgBox("Files copied successifully")
                        ' My.Computer.FileSystem.CopyFile(foundFile, destinationDir & "\" & Path.GetFileName(foundFile), True)
                End Select
            Next
        Else
            MsgBox("Error")
        End If
    End Sub

    'Function to Get SQL files to be executed

    Private Function GetSql() As String()
        Try

            ' Gets the current assembly.
            Dim Asm As Assembly = [Assembly].GetExecutingAssembly()

            Dim cmd As New SqlCommand

            'Dim proc As New Process
            ' Dim sPath As String
            Dim tableText As String

            Dim targetDirectory As String = "C:\FtpFiles\OVC\UpdatedOVC\"
            Dim fileEntries As String() = Directory.GetFiles(targetDirectory, "*.sql")
            Dim j As Int32 = fileEntries.Length - 1
            Dim tArray(j) As String
            ' Process the list of .sql files found in the directory. '
            Dim fileName As String
            Dim i As Int32 = 0
            For Each fileName In fileEntries
                Console.WriteLine(fileName)
                Dim strm As Stream = File.Open(fileName, FileMode.Open)
                Dim reader As StreamReader = New StreamReader(strm)
                tableText = reader.ReadToEnd()
                'MsgBox(tableText)
                tArray(i) = tableText

                i = i + 1

            Next

            Return tArray

        Catch ex As Exception
            MsgBox("In GetSQL: " & ex.Message)
            Throw ex
        End Try
    End Function

    Private Sub ExecuteSql(ByVal databaseName As String, ByVal sql() As String)

        'Connect to the local, default instance of SQL Server.
        Dim srv As Server
        srv = New Server
        'Define a Database object variable by supplying the server and the database name arguments in the constructor.
        Dim db As Database = New Database()
        db = New Database()

        'Reference the database
        db = srv.Databases(databaseName)
        'Dim scrip As String = "C:\Users\Ezekiel\Documents\script.sql"


        Try
            For Each script As String In sql
                ' Console.Write(script)
                db.ExecuteNonQuery(script)

            Next

        Catch ex As Exception
           ' MessageBox.Show(ex.Message)
            ' returnValue = returnValue + Environment.NewLine + "Exception: " + ex.Message
            Dim ex1 As Exception
            ex1 = ex.InnerException
            Do While ex1 IsNot (Nothing)
                MessageBox.Show(ex1.Message)
                'Me.TextBox1.Text += ex1.Message
                Console.Write(ex1.Message)
                Console.Write(vbCrLf)
                ' returnValue = returnValue + Environment.NewLine + "Inner Exception: " + ex1.Message
                ex1 = ex1.InnerException
            Loop

        Finally
            ' Closing the connection should be done in a Finally block

        End Try
    End Sub

    Sub Main()

        Dim host As String = "ftp://icfkenya.cloudapp.net/PrivateLink/NormalUsers/"
        Dim versionFile As String = host & "version.txt"
        Dim ovcFile As String = host & "ovc.zip"
        Dim file As String = "version.txt"
        Dim ftpusername As String = "User1"
        Dim ftppassword As String = "user1pass"
        Dim downloadPath As String = "C:\FtpFiles/olmis"
        Dim ovcDownloadPath As String = "C:\FtpFiles/OVC"

        If IsConnectionAvailable() = True Then

            'Const N As Byte = 5

            ' For i As Byte = 0 To N - 1
            'Next i
            DownloadFile(downloadPath, host, versionFile, ftpusername, ftppassword)

            MsgBox("Version File Download Successiful!")
            'DeleteFtpFile(versionFile)

            'Read text file to get version - FTP location version
            Dim ftpversion As Int32
            Try
                ' Open the file using a stream reader.
                Using sr As New StreamReader("C:\FtpFiles/olmis/version.txt")

                    ' Read the stream to a string and write the string to the console.
                    ftpversion = sr.ReadToEnd()
                    Console.WriteLine(ftpversion)
                End Using
            Catch e As Exception
                Console.WriteLine("The ftp location version File  could not be read:")
                Console.WriteLine(e.Message)
            End Try
            'End reading file

            'Read text file to get version - Running location version
            Dim localversion As Int32
            Try
                ' Open the file using a stream reader.
                Using sr As New StreamReader("C:\FtpFiles/runningVersion/OVCVersion.txt")

                    ' Read the stream to a string and write the string to the console.
                    localversion = sr.ReadToEnd()
                    Console.WriteLine(localversion)
                End Using
            Catch e As Exception
                Console.WriteLine("The running version File  could not be read:")
                Console.WriteLine(e.Message)
            End Try
            'End reading file
            If ftpversion > localversion Then
                'Call method to download OVC zipped file
                'Debug.Print("===========Downloading OVC File=============")
                ' DownloadOvc(ovcDownloadPath, host, ovcFile, ftpusername, ftppassword)
                MsgBox("OVC Download Successiful!")

                'Unzip downloaded files
                'Dim sourceArchiveFileName As String = "C:\FtpFiles\OVC\ovc.zip"
                'Dim destinationDirectoryName As String = "C:\FtpFiles/OVC/"
                'Call method to unzip
                'ExtractToDirectory(sourceArchiveFileName, destinationDirectoryName)

                'Copy and replace files

                Dim sourceDir As String = "C:\FtpFiles\OVC\"
                Dim destinationDir As String = "C:\FtpFiles\OVC\UpdatedOVC"

                Debug.Print("===========Replacing Files=============")
                ' If System.IO.Directory.Exists(sourceDir) = True Then
                ' CopyAndReplace(sourceDir, destinationDir)
                Debug.Print("=========== End Replacing Files=============")
                'End If

                'Get SQL files

                'Function call...
                Dim sql() As String = GetSql()

                ' For Each script As String In sql
                'MsgBox(script)
                ' Next

                'Run the SQL files

                ExecuteSql("APHIAMAINDB", sql)



            End If
        Else
            MsgBox("Check your internet connection")

        End If
    End Sub


    Private Sub Button1_Click(sender As System.Object, e As EventArgs) Handles Button1.Click
        Main()
    End Sub
End Class
