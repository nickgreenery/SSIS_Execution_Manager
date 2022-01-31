Imports System
Imports System.IO
Imports System.IO.DirectoryInfo
Imports System.Xml
Imports System.Environment
Imports System.Net.Mail
Imports System.Net
Imports System.Text
Imports System.Collections.Generic
Imports System.Data.SqlClient


Module ExecutionManager

    Sub Main(ByVal args As String())
        Try

            Dim doc As New XmlDocument
            Dim nodeList As XmlNodeList
            Dim nodes As XmlNode
            Dim dir, NamePattern, SsisPackage, DtsConfig, Excelbit, StepDesc, emailServer, ToAddress, EmailBody, WorkingDirName, dtexecDirectory64, dtexecDirectory86 As String
            Dim FileProcessed As String = "0"

            EmailBody = "Below data feeds were loaded:" + Environment.NewLine + Environment.NewLine

            'Test to see if help or ? was typed in and provide instructions how to use the program
            If args(0) = "/help" Or args(0) = "/?" Then
                Console.WriteLine("Need to specify cofiguration file." & Environment.NewLine & "Example: ProgramName.exe ""c:\folder1\XML_File.xml""")
            Else

                'Read in XML document
                doc.Load(args(0))
                'populate node list
                nodeList = doc.SelectNodes("/Files/File")

                For Each nodes In nodeList

                    Try

                  
                    'map xml values to variables
                    StepDesc = nodes("Description").InnerText
                    dir = nodes("FilePath").InnerText
                    NamePattern = nodes("NamePattern").InnerText
                    SsisPackage = nodes("SsisPackage").InnerText
                    DtsConfig = nodes("ConfigFile").InnerText
                    Excelbit = nodes("Excel").InnerText
                    emailServer = doc.SelectSingleNode("/Files/EmailServer").InnerText.ToString
                    ToAddress = doc.SelectSingleNode("/Files/ToAddress").InnerText.ToString
                    WorkingDirName = doc.SelectSingleNode("/Files/WorkingDirName").InnerText.ToString
                    dtexecDirectory64 = doc.SelectSingleNode("/Files/dtexecDirectory64").InnerText.ToString
                    dtexecDirectory86 = doc.SelectSingleNode("/Files/dtexecDirectory86").InnerText.ToString
                    'OutputFile = doc.SelectSingleNode("/newfilenotification/output").InnerText.ToString
                    'If path does not end with a "\" then add it
                    If Right(dir, 1) <> "\" Then
                        dir = dir & "\"
                    End If

                    If Right(WorkingDirName, 1) <> "\" Then
                        WorkingDirName = WorkingDirName & "\"
                    End If

                    'Console.WriteLine(dir & NamePattern)

                    Dim di As DirectoryInfo = New DirectoryInfo(dir)
                    Dim fils As FileInfo() = di.GetFiles(NamePattern)

                    If fils.Count >= 1 Then


                        Dim Movefiles() As String


                      

                        Movefiles = System.IO.Directory.GetFiles(dir, NamePattern)

                        Dim MoveFile As String
                            Console.WriteLine("--------------------------------------------------------------")
                            Console.WriteLine("--------------------------------------------------------------")
                            Console.WriteLine("*****" & StepDesc & "*****")
                            Console.WriteLine("File Processing Started: " & DateTime.Now)
                            

                        Console.WriteLine("Move files to Working Directory...")
                        For Each MoveFile In Movefiles
                            'System.IO.File.Move(MoveFile, "c:\destination\" & System.IO.Path.GetFileName(File))

                            Console.WriteLine(dir & "Working\" & System.IO.Path.GetFileName(MoveFile))
                            System.IO.File.Move(MoveFile, dir & WorkingDirName & System.IO.Path.GetFileName(MoveFile))
                            System.Threading.Thread.Sleep(5000)



                        Next

                        Console.WriteLine("Executing package: [" & SsisPackage & "] For " & StepDesc)
                        PackageRunner(SsisPackage, DtsConfig, Excelbit, dir, StepDesc, dtexecDirectory64, dtexecDirectory86)
                        EmailBody = EmailBody & StepDesc & "Processed from " & dir & Environment.NewLine & vbCrLf

                        FileProcessed = 1
                    End If


                    Catch ex As Exception

                        StepDesc = nodes("Description").InnerText

                        Console.WriteLine("--------------------------------------------------------------")
                        Console.WriteLine("--------------------------------------------------------------")
                        Console.WriteLine("*****" & StepDesc & "*****")

                        'this "Error" line is coded to match the defaul way DTEXEC captures errors so our logging capturing captures this
                        Console.WriteLine("Error: " & DateTime.Now)

                        'this "Description" line is coded to match the defaul way DTEXEC captures errors so our logging capturing captures this
                        Console.WriteLine("   Description:" & ex.ToString)

                        'this "Source" line is coded to match the defaul way DTEXEC captures errors so our logging capturing captures this
                        Console.WriteLine("   Source:" & ex.ToString)


                    End Try

                Next


            End If
            If FileProcessed = 1 Then
                SendEmail(EmailBody, ToAddress, emailServer)
                Console.WriteLine("File Processing Complete: " & DateTime.Now)
            End If

        Catch ex As Exception
            Console.WriteLine("Error: Type /help or /? for assistance.")
            Console.WriteLine(ex.ToString)
        End Try
    End Sub

    Sub PackageRunner(ByVal Package As String, ByVal config As String, ByRef Excelbit As String, ByVal dir As String, ByVal procdesc As String, dtexecDirectory64 As String, dtexecDirectory86 As String)
        'Console.WriteLine("""C:\Program Files\Microsoft SQL Server\100\DTS\Binn\dtexec.exe"" /f """ + Package + """ /Conf  """ + config + """  ")

        Dim execString As String = ""

        If Excelbit = "False" Then ' 64 bit dtexex
            If config = "" Then
                execString = """" + dtexecDirectory64 + """ /FILE """ + Package + """ /CHECKPOINTING OFF /REPORTING E"
            Else
                execString = """" + dtexecDirectory64 + """ /FILE """ + Package + """  /CONFIGFILE """ + config + """ /CHECKPOINTING OFF /REPORTING E"
            End If


        End If

        If Excelbit = "True" Then ' 32 bit dtexec

            If config = "" Then
                execString = """" + dtexecDirectory86 + """ /FILE """ + Package + """ /CHECKPOINTING OFF /REPORTING E"
            Else
                execString = """" + dtexecDirectory86 + """ /FILE """ + Package + """  /CONFIGFILE """ + config + """ /CHECKPOINTING OFF /REPORTING E"
            End If

        End If

        Console.WriteLine(execString)

        Try
            Shell(execString, AppWinStyle.NormalNoFocus, True)
            'LogToTable(dir, Package, procdesc, config, Excelbit, execString, 1, "")
        Catch ex As Exception
            'LogToTable(dir, Package, procdesc, config, Excelbit, execString, 0, ex.ToString)
        End Try



    End Sub

    Sub LogToTable(ByVal dir As String, ByVal SsisPackage As String, ByVal ProcDesc As String, ByVal Config As String, ByVal Excelbit As String, ByVal cmd As String, ByVal success As String, ByVal errormsg As String)
        Dim myConnection = New SqlConnection("server=SQLRPT;uid=SsisExecutionmanager;pwd=Bx86*2wAS;database=Logging")

        If Excelbit = "True" Then
            Excelbit = 1
        Else
            Excelbit = 0
        End If



        Try
            myConnection.Open()
            Dim myCommand = New SqlCommand("INSERT INTO SsisExecutionManagerLog(ProcessingDirectory, PackageExecuted, ProcessDescription, ConfigurationFile, IsExcel, CommandExecuted, Success, ErrorMessage) VALUES('" & dir & "','" & SsisPackage & "','" & ProcDesc & "','" & Config & "','" & Excelbit & "', '" & cmd & "', '" & success & "','" & errormsg & "')")
            myCommand.Connection = myConnection
            myCommand.ExecuteNonQuery()
            'Console.WriteLine("INSERT INTO SsisExecutionManagerLog(ProcessingDirectory, PackageExecuted, ProcessDescription, ConfigurationFile, IsExcel, CommandExecuted, Success, ErrorMessage) VALUES('" & dir & "','" & SsisPackage & "','" & ProcDesc & "','" & Config & "','" & Excelbit & "', '" & cmd & "', '" & success & "','" & errormsg & "')")

        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
        myConnection.Close()
    End Sub

    Sub SendEmail(ByVal body As String, ByVal ToAddress As String, ByVal EmailServer As String)
        Try
            Dim SmtpServer As New SmtpClient()
            Dim mail As New MailMessage()
            'SmtpServer.Credentials = New Net.NetworkCredential("username@gmail.com", "password")
            'SmtpServer.Port = 587
            SmtpServer.Host = EmailServer
            mail = New MailMessage()
            mail.From = New MailAddress("SsisExecutionManager@email.com")
            mail.To.Add(ToAddress)
            mail.Subject = "Processed Data Feeds"
            mail.Body = body
            SmtpServer.Send(mail)
            Console.WriteLine("mail send")
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Sub


End Module
