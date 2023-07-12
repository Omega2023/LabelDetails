Imports System
Imports System.Collections
Imports System.Data
Imports System.Data.Common
Imports System.Data.DataTable
Imports System.Data.Odbc
Imports System.Data.SqlClient
Imports System.Drawing
Imports System.EventArgs
Imports System.Globalization
Imports System.IO
Imports System.Management
Imports System.Net
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Security.Principal
Imports System.Text
Imports System.Threading 'For Sleep
Imports System.Windows.Forms
Imports System.Windows.Forms.Application
Imports System.Xml

Public Class LabelDetails
    Inherits System.Windows.Forms.Form

#If DEBUG Then
    'debug use C:\ ...
    Dim xmlfn As String = "C:\Apps\OmegaLabels\LabelDetails.xml"
#Else
    Dim xmlfn As String = "\\OmegaFS2\Apps\OmegaLabels\LabelDetails.xml"
#End If

    'Global items
    Dim AppLogDir As String = "C:\Log\"
    Dim AppLog As String = "LabelDetails.txt"
    Dim MyApplication As String = "LabelDetails"
    Dim AppDir As String = "C:\Apps\OmegaLabels\"
    Dim LabelDetailsData As String = "LabelDetailsData.txt"
    Dim LabelDetailsKits As String = "LabelDetailsKitPage.txt"

    'LabelDetails.xml file initializations
    Dim AppVersion As String = "1.0.0.0"

#If DEBUG Then
    'debug use C:\ ...
    Dim AppXmlDir As String = "C:\My Web Sites\LABELS\" 'App Html File Folder
#Else
    Dim AppXmlDir As String = "C:\My Web Sites\LABELS\" 'App Html File Folder
#End If

    Private Sub LabelDetails_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim status As Boolean = False
        Static loaded As Boolean = False

        If Not loaded Then
            LogEvent("Delete LabelDetails Log file: " & AppLog)
            Dim fnLog As String = Path.Combine(AppLogDir, AppLog)
            If Directory.Exists(AppLogDir) Then
                If File.Exists(fnLog) Then
                    File.Delete(fnLog)
                End If
            Else
                Directory.CreateDirectory(AppLogDir)
                File.Create(fnLog)
            End If
            LogEvent("Delete LabelDetails Data file: " & LabelDetailsData)
            File.Delete(Path.Combine(AppDir, LabelDetailsData))
            loaded = True
        Else
            LogEvent("LabelDetails reloaded")
        End If

        If System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed Then
            With System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion
                Me.Text &= " Version " & .Major & "." & .Minor & "." & .Build & "." & .Revision
            End With
            AppVersion = My.Application.Deployment.CurrentVersion.ToString()
        Else
            'Put the assembly version here if the application is not a publish-ready ClickOnce app
            AppVersion = System.Windows.Forms.Application.ProductVersion()
        End If
        Me.Text &= " Version " & AppVersion
        LogEvent("*** App Version " & AppVersion & " ***" & " Me.Text: " & Me.Text)

        ImportXmlInit(xmlfn, status)
        If Not status Then
            End
        End If

        btnProcess.BackColor = Color.LightYellow
        TextBoxPath.BackColor = Color.LightYellow

        TextBoxPath.Text = AppXmlDir
        LogEvent("AppXmlDir: " & AppXmlDir)

        'Reset the ProgressBar
        Me.ProgressBar1.Value = 0
        Me.ProgressBar1.Minimum = 0
        Me.ProgressBar1.Maximum = 0
        Me.ProgressBar1.Visible = False
    End Sub

    Public Sub LogEvent(ByRef msg As String)

        Dim fn As String = Path.Combine(AppLogDir, AppLog)
        Try
            Dim sw As StreamWriter
            If Directory.Exists(AppLogDir) Then
                sw = New StreamWriter(fn, True, Encoding.Default)
                'sw.WriteLine(Date.Now & " - " & msg)
                sw.WriteLine(String.Format("{0:yyyy/MM/dd HH:mm}", Date.Now) & " - " & msg)
                sw.Flush()
                sw.Close()
            Else
                Directory.CreateDirectory(AppLogDir)
                File.Create(fn)
                sw = New StreamWriter(fn, True, Encoding.Default)
                'sw.WriteLine(Date.Now & " - " & msg)
                sw.WriteLine(String.Format("{0:yyyy/MM/dd HH:mm}", Date.Now) & " - " & msg)
                sw.Flush()
                sw.Close()
            End If

        Catch ex As Exception
            MsgBox("Log Event " & ex.Message)
            Dim Response = MsgBox(ex.Message & vbCrLf & "Do you want to retry? ", MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton2 Or MsgBoxStyle.Critical, "LabelDetails")
            If Response = MsgBoxResult.Yes Then
                LogEvent(msg)
            Else
                End
            End If
        End Try

    End Sub

    Public Sub ImportXmlInit(ByRef fn As String, ByRef status As Boolean)
        Dim m_xmld As New XmlDocument
        Dim Parsing As String = "ImportXmlInit"
        Dim iSL2 As Integer = 0

        Try
            If File.Exists(fn) Then
#If DEBUG Then
                fn = Replace(fn, "\\OmegaFS2", "C:")
#End If
                LogEvent("Load File: " & fn)
                m_xmld.Load(fn)
                Dim elemList As XmlNodeList

                LogEvent("ImportXmlInit function")

                Parsing = "AppVersion"
                LogEvent(Parsing)
                elemList = m_xmld.GetElementsByTagName(Parsing)
                Dim AppVer As String = elemList(0).InnerText.ToString
                If Not AppVer.Equals(AppVersion) Then
                    LogEvent("Warning - Application and LabelDetails.xml Versions Do Not Match.")
                    LogEvent("Application Version: " & AppVersion & " XML Version: " & AppVer)
                End If
                LogEvent("AppVer: " & AppVer)

                Parsing = "AppDir"
                LogEvent(Parsing)
                elemList = m_xmld.GetElementsByTagName(Parsing)
                AppDir = elemList(0).InnerText.ToString
                LogEvent("AppDir: " & AppDir)

                Parsing = "AppXmlDir"
                LogEvent(Parsing)
                elemList = m_xmld.GetElementsByTagName(Parsing)
                AppXmlDir = elemList(0).InnerText.ToString
                Try
                    If Directory.Exists(AppXmlDir) Then
                        LogEvent("AppXmlDir exists: " & AppXmlDir)
                    End If

                Catch ex As Exception
                    Dim Response = MessageBox.Show("Archive directory: " & AppXmlDir & " does not exist.", "LabelDetails")
                    MsgBox("Please correct the setting <AppXmlDir> in " & fn)
                    End
                End Try

                status = True
                LogEvent("Return Status: " & status)
            Else
                MsgBox("Settings file not found: " & fn & vbCrLf)
                LogEvent("File not found: " & fn & vbCrLf)
            End If

        Catch ex As Exception
            status = False
            If ex.Message.Contains("Could not find file ") Then
                MsgBox("Settings file not found: " & fn & vbCrLf & ex.Message)
                LogEvent("File not found: " & fn & vbCrLf & ex.Message)

            Else
                MsgBox("A fatal error was caused when parsing " & Parsing & " in file " & fn & vbCrLf & ex.Message & vbCrLf)
                LogEvent("A fatal error was caused when parsing " & Parsing & " in file " & fn & vbCrLf & ex.Message)
            End If

            End
        End Try

    End Sub

    Private Sub LabelArchiveVerify(ByRef LabelArchiveDir As String, ByRef LabelAgeOut As String)
        Dim aDate As Date = Date.Today.AddDays(-LabelAgeOut) 'Started with 30 day ageout
        Dim ArchiveDir As New DirectoryInfo(LabelArchiveDir)

        Try
            LogEvent("aDate " & aDate.ToString("MM/dd/yy"))
            'Check if the target directory exists, if not, create it.
            If Directory.Exists(LabelArchiveDir) = False Then
                Directory.CreateDirectory(LabelArchiveDir)
                LogEvent("Created directory: " & LabelArchiveDir)
            End If

            Dim strLastModified As String
            For Each f_info As FileInfo In ArchiveDir.GetFiles()
                If f_info.IsReadOnly Then Continue For
                strLastModified = System.IO.File.GetLastWriteTime(Path.Combine(LabelArchiveDir, f_info.ToString())).ToShortDateString()
                If strLastModified <= aDate Then
                    If File.Exists(f_info.FullName) Then
                        File.Delete(f_info.FullName)
                        LogEvent("Delete Archive File: " & f_info.FullName)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox("Label Archive" & vbCrLf & ex.Message)
            LogEvent("Label Archive" & vbCrLf & ex.Message)
        End Try
        LogEvent("End Sub Label Archive")
    End Sub

    Private Sub LabelArchiveKit(ByRef LabelArchiveDir As String, ByRef LabelAgeOut As String)
        Dim aDate As Date = Date.Today.AddDays(-LabelAgeOut) 'Started with 30 day ageout
        Dim ArchiveDir As New DirectoryInfo(LabelArchiveDir)

        ' Check if the target directory exists, if not, create it.
        If Directory.Exists(LabelArchiveDir) = False Then
            Directory.CreateDirectory(LabelArchiveDir)
            LogEvent("Created directory: " & LabelArchiveDir)
        End If
        Dim strLastModified As String
        For Each f_info As FileInfo In ArchiveDir.GetFiles()
            If f_info.IsReadOnly Then Continue For
            strLastModified = System.IO.File.GetLastWriteTime(Path.Combine(LabelArchiveDir, f_info.ToString())).ToShortDateString()
            If strLastModified <= aDate Then
                If File.Exists(f_info.FullName) Then
                    File.Delete(f_info.FullName)
#If DEBUG Then
                    LogEvent("Delete Archive File: " & f_info.FullName)
#End If
                End If
            End If
        Next
    End Sub

    Private Sub ReprintArchive(ByVal FilesPath As String)
        'Const vbCrLf = "\n\r"
        Dim XmlDir As New DirectoryInfo(AppXmlDir)
    End Sub

    Private Sub WriteOutLine(ByRef msg As String, ByRef tmpFN As String)
        'WriteLine
        Dim fn As String = tmpFN 'Path.Combine(AppLogDir, AppLog)
        Try
            Dim sw As StreamWriter
            If Directory.Exists(AppLogDir) Then
                sw = New StreamWriter(fn, True, Encoding.Default)
                sw.WriteLine(msg)
                sw.Flush()
                sw.Close()
            Else
                Directory.CreateDirectory(AppLogDir)
                File.Create(fn)
                sw = New StreamWriter(fn, True, Encoding.Default)
                sw.WriteLine(msg)
                sw.Flush()
                sw.Close()
            End If
        Catch ex As Exception
            MsgBox("Write Out Line" & vbCrLf & ex.Message)
            LogEvent("Write Out Line" & vbCrLf & ex.Message)
        End Try
    End Sub

    Private Sub btnBrowse_Click(sender As Object, e As EventArgs) Handles btnBrowse.Click
        Dim fd As OpenFileDialog = New OpenFileDialog()

        fd.Title = "Open File Dialog"
        fd.InitialDirectory = "C:\"
        fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
        fd.FilterIndex = 2
        fd.RestoreDirectory = True

        If fd.ShowDialog() = DialogResult.OK Then
            TextBoxPath.Text = fd.FileName
        End If
    End Sub

    Private Sub btnProcess_Click(sender As Object, e As EventArgs) Handles btnProcess.Click
        Dim fn As String 'Create full filenames
        Dim ArchiveDir As New DirectoryInfo(TextBoxPath.Text)
        Dim sTemplate As String = ""
        Dim tmpFN As String = ""
        Dim Candiates(10000) As String
        Dim inFile As Integer = 0
        Dim numCandiates As Integer = 0

        btnProcess.BackColor = Color.OrangeRed

        'ArchiveDir = TextBoxPath.Text.ToString()

        Try
            'Reset the ProgressBar
            Me.ProgressBar1.Visible = True
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Minimum = 0
            Me.ProgressBar1.Maximum = ArchiveDir.GetFiles.Length
            For Each f_info As FileInfo In ArchiveDir.GetFiles() 'Oldest file is tested first
                fn = Path.Combine(ArchiveDir.ToString(), f_info.ToString())
                If fn.Contains("tmp") Then Continue For
                LogEvent("Filename " & vbLf & fn)
                If fn.Contains("html") Then tmpFN = Replace(fn, "html", "tmp")
                If Not File.Exists(tmpFN) Then
                    tmpFN = fn
                    Dim fStream As FileStream = File.Create(tmpFN)
                    fStream.Flush()
                    fStream.Close()
                    LogEvent("Tmp Filename " & tmpFN)
                End If

                Dim sr As StreamReader = New StreamReader(fn)
                Do While sr.Peek() >= 0
                    sTemplate = sr.ReadLine
                    If sTemplate.Trim().Equals(String.Empty) Then
                        'Set the value of ProgressBar
                        If Me.ProgressBar1.Value < Me.ProgressBar1.Maximum Then
                            Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
                        End If
                        Continue Do
                    End If

                    If sTemplate.Contains("HTTrack") Then 'is HTTrack in file?
                        'LogEvent("sTemplate HTTrack: " & sTemplate)
                        If sTemplate.Contains("Mirrored") Then
                            'Set the value of ProgressBar
                            If Me.ProgressBar1.Value < Me.ProgressBar1.Maximum Then
                                Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
                            End If
                            Continue Do
                        Else
                            sTemplate = Replace(sTemplate, "<!-- Added by HTTrack -->", "")
                            sTemplate = Replace(sTemplate, "<!-- /Added by HTTrack -->", "")
                        End If
                        'WriteLine
                        WriteOutLine(sTemplate.TrimEnd, tmpFN)
                        TextBoxPath.AppendText(sTemplate.Trim() & vbCrLf) 'Append
                        'LogEvent("sTemplate: " & sTemplate)
                    ElseIf sTemplate.Contains("><") Then 'try next row
                        'LogEvent("sTemplate Tags: " & sTemplate)
                        Dim TemplateParts() = Split(sTemplate, "><")
                        LogEvent("TemplateParts Length " & TemplateParts.Length)
                        Dim idx As Integer = 0
                        For idx = 0 To TemplateParts.Length - 1
                            If Not TemplateParts(idx).Trim.Equals(String.Empty) Then
                                LogEvent("TemplateParts (" & idx & ")" & " of " & TemplateParts.Length & " " & TemplateParts(idx).Trim())
                                Dim sTemp = "<" & TemplateParts(idx).Trim() & ">"
                                If idx.Equals(0) Then sTemp = Replace(sTemp, "<<", "<")
                                If TemplateParts.Length > 0 Then
                                    If idx.Equals(TemplateParts.Length - 1) Then sTemp = Replace(sTemp, ">>", ">")
                                End If
                                inFile += 1
                                WriteOutLine(sTemp.TrimEnd, tmpFN)
                                TextBoxPath.AppendText(sTemp.TrimEnd & vbCrLf) 'Append
                            End If
                        Next
                        LogEvent("Remove Lines Candiates(" & numCandiates & ") " & Candiates(numCandiates))
                        numCandiates += 1
                    Else
                        WriteOutLine(sTemplate.TrimEnd, tmpFN)
                        'TextBoxPath.AppendText(sTemplate.TrimEnd & vbCrLf) 'Append
                        'LogEvent("sTemplate: " & sTemplate)
                    End If
                Loop
                sr.Close()
                'Set the value of ProgressBar
                If Me.ProgressBar1.Value < Me.ProgressBar1.Maximum Then
                    Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
                End If
            Next
        Catch ex As Exception
            MsgBox("Button Go Click" & vbCrLf & ex.Message)
            LogEvent("Button Go Click" & vbCrLf & ex.Message)
        Finally
            'Reset the ProgressBar
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Visible = False
        End Try

        btnProcess.BackColor = Color.LightYellow
    End Sub


    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        End
    End Sub
End Class
