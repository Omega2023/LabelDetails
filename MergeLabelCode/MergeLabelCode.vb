Imports System
Imports System.Collections
Imports System.Data
Imports System.Data.Common
Imports System.Drawing
Imports System.EventArgs
Imports System.Globalization
Imports System.IO
Imports System.Net
Imports System.Reflection
Imports System.Text
Imports System.Windows.Forms
Imports System.Windows.Forms.Application
Imports System.Xml

'Imports Microsoft.Office.Interop
'Imports Microsoft.Office.Interop.Excel

Public Class MergeLabelCode
    Inherits System.Windows.Forms.Form

#If DEBUG Then
    'debug use C:\ ...
    Dim xmlfn As String = "C:\Apps\OmegaLabels\MergeLabelCode.xml"
#Else
    Dim xmlfn As String = "\\OmegaFS2\Apps\OmegaLabels\MergeLabelCode.xml"
#End If

    'Global items
    Dim AppLogDir As String = "C:\Log\"
    Dim AppLog As String = "MergeLabelCode.txt"
    Dim MyApplication As String = "MergeLabelCode"
    Dim AppDir As String = "C:\Visual Studio 2019\Projects\"
    Dim MergeLabelCodeData As String = "MergeLabelCodeData.txt"
    Dim MergeLabelCodeKits As String = "MergeLabelCodeKitPage.txt"

    'MergeLabelCode.xml file initializations
    Dim AppVersion As String = "1.0.0.0"

#If DEBUG Then
    'debug use C:\ ...
    Dim AppHtmlDir As String = "C:\Visual Studio 2019\Projects\" 'App Html File Folder
#Else
    Dim AppHtmlDir As String = "C:\Visual Studio 2019\Projects\" 'App Html File Folder
#End If

    Private Sub MergeLabelCode_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim status As Boolean = False
        Static loaded As Boolean = False

        If Not loaded Then
            LogEvent("Delete MergeLabelCode Log file: " & AppLog)
            Dim fnLog As String = Path.Combine(AppLogDir, AppLog)
            If Directory.Exists(AppLogDir) Then
                If File.Exists(fnLog) Then
                    File.Delete(fnLog)
                End If
            Else
                Directory.CreateDirectory(AppLogDir)
                File.Create(fnLog)
            End If
            LogEvent("Delete MergeLabelCode Data file: " & MergeLabelCodeData)
            File.Delete(Path.Combine(AppDir, MergeLabelCodeData))
            loaded = True
        Else
            LogEvent("MergeLabelCode reloaded")
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

        TextBoxPath.Text = AppHtmlDir
        LogEvent("AppHtmlDir: " & AppHtmlDir)

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
            Dim Response = MsgBox(ex.Message & vbCrLf & "Do you want to retry? ", MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton2 Or MsgBoxStyle.Critical, "MergeLabelCode")
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
        Dim n As Integer = 0

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
                    LogEvent("Warning - Application and MergeLabelCode.xml Versions Do Not Match.")
                    LogEvent("Application Version: " & AppVersion & " XML Version: " & AppVer)
                End If
                LogEvent("AppVer: " & AppVer)

                Parsing = "AppDir"
                LogEvent(Parsing)
                elemList = m_xmld.GetElementsByTagName(Parsing)
                AppDir = elemList(0).InnerText.ToString
                LogEvent("AppDir: " & AppDir)

                Parsing = "AppHtmlDir"
                LogEvent(Parsing)
                elemList = m_xmld.GetElementsByTagName(Parsing)
                AppHtmlDir = elemList(0).InnerText.ToString
                Try
                    If Directory.Exists(AppHtmlDir) Then
                        LogEvent("AppHtmlDir exists: " & AppHtmlDir)
                    End If

                Catch ex As Exception
                    Dim Response = MessageBox.Show("Archive directory: " & AppHtmlDir & " does not exist.", "MergeLabelCode")
                    MsgBox("Please correct the setting <AppHtmlDir> in " & fn)
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

    Private Sub ReprintArchive(ByVal FilesPath As String)
        'Const vbCrLf = "\n\r"
        Dim HtmlDir As New DirectoryInfo(AppHtmlDir)
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
        Dim ArchiveDir As New DirectoryInfo(AppHtmlDir)
        Dim sTemplate As String = ""
        Dim tmpFN As String = ""
        Dim Candiates(10000) As String
        Dim inFile As Integer = 0
        Dim numCandiates As Integer = 0

        btnProcess.BackColor = Color.OrangeRed

        Try
            'Reset the ProgressBar
            Me.ProgressBar1.Visible = True
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Minimum = 0
            Me.ProgressBar1.Maximum = ArchiveDir.GetFiles.Length

            fn = TextBoxPath.Text
            LogEvent("Filename " & vbLf & fn)
            If fn.Contains("vb") Then tmpFN = Replace(fn, "vb", "tmp")
            If Not File.Exists(tmpFN) Then
                Dim fStream As FileStream = File.Create(tmpFN)
                fStream.Flush()
                fStream.Close()
                LogEvent("Tmp Filename " & tmpFN)
            End If
            LogEvent("fn: " & fn)
            Dim sr As StreamReader = New StreamReader(fn)
            Me.ProgressBar1.Maximum = File.ReadAllLines(fn).Length
            Do While sr.Peek() >= 0
                sTemplate = sr.ReadLine
                'LogEvent("sTemplate: " & sTemplate)
                'If sTemplate.Trim().Equals(String.Empty) Then
                '    'Set the value of ProgressBar
                '    If Me.ProgressBar1.Value < Me.ProgressBar1.Maximum Then
                '        Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
                '    End If
                '    Continue Do
                'End If
                If sTemplate.Contains("BCLabels.xml") Then
                    sTemplate = Replace(sTemplate, "BCLabels.xml", "BCPartLabels.xml")
                End If
                If sTemplate.Contains("BCLabels.txt") Then
                    sTemplate = Replace(sTemplate, "BCLabels.txt", "BCPartLabels.txt")
                End If
                If sTemplate.Contains("BC Labels") Then
                    sTemplate = Replace(sTemplate, "BC Labels", "Part Labels")
                End If
                If sTemplate.Contains("Software\BCLabels") Then
                    sTemplate = Replace(sTemplate, "Software\BCLabels", "Software\BCPartLabels")
                End If
                If sTemplate.Contains("Sales Order Number") Then
                    sTemplate = Replace(sTemplate, "Sales Order Number", "Part Number")
                End If
                If sTemplate.Contains("Sales Order") Then
                    sTemplate = Replace(sTemplate, "Sales Order", "Part Number")
                End If
                If sTemplate.Contains("RePrint Labels") Then
                    sTemplate = Replace(sTemplate, "RePrint Labels", "Part Labels")
                End If
                If sTemplate.Contains("txbSalesOrder "" & txbSalesOrder.Text") Then
                    sTemplate = Replace(sTemplate, "txbSalesOrder "" & txbSalesOrder.Text", "txbPartNum "" & txbPartNum.Text")
                End If
                If sTemplate.Contains("txbxSalesOrderText "" & txbxSalesOrderText") Then
                    sTemplate = Replace(sTemplate, "txbxSalesOrderText "" & txbxSalesOrderText", "txbxPartNumText "" & txbxPartNumText")
                End If
                If sTemplate.Contains("If (ZoneList.Count() > 1 Or ZoneList.Length > 1) Or") Then
                    sTemplate = Replace(sTemplate, "If (ZoneList.Count() > 1 Or ZoneList.Length > 1) Or", "If")
                End If
                If sTemplate.Contains("Dim SortIndex As Integer = daCol.zone_name") Then
                    sTemplate = Replace(sTemplate, "Dim SortIndex As Integer = daCol.zone_name", "Dim SortIndex As Integer = daCol.Customer")
                End If
                If sTemplate.Contains("Dim SortIndex2 As Integer = daCol.locat") Then
                    sTemplate = Replace(sTemplate, "Dim SortIndex2 As Integer = daCol.locat", "Dim SortIndex2 As Integer = daCol.salesorder")
                End If
                If sTemplate.Contains("1=zone_name, 2=locat") Then
                    sTemplate = Replace(sTemplate, "1=zone_name, 2=locat", "1=Customer, 2=salesorder")
                End If
                If sTemplate.Contains("Me.Text = ""Part Labels """) Then
                    sTemplate = Replace(sTemplate, "Me.Text = ""Part Labels """, "Me.Text = ""Part Labels""")
                End If
                If sTemplate.Contains("#If DEBUG Then") Then
                    'Remove the #Else part keeping only the #If DEBUG part
                    '#If DEBUG Then
                    'Dim xmlfn As String = "C:\Apps\OmegaLabels\BCPartLabels.xml"
                    '#Else
                    'Dim xmlfn As String = ".\BCPartLabels.xml"
                    '#End If
                    'C:\Visual Studio 2019\Projects\BCPartLabels\MainWindow.vb
                    Do While Not sTemplate.Contains("#Else")
                        LogEvent("sTemplate: " & sTemplate)
                        sTemplate = sr.ReadLine
                        If sTemplate.Contains("BCLabels.xml") Then
                            sTemplate = Replace(sTemplate, "BCLabels.xml", "BCPartLabels.xml")
                        End If
                        If sTemplate.Contains("#End If") Then Exit Do
                        If Not sTemplate.Contains("#Else") Then WriteOutLine(sTemplate, tmpFN)
                    Loop
                    If sTemplate.Contains("#End If") Then
                        WriteOutLine(sTemplate, tmpFN)
                    Else
                        If sTemplate.Contains("#Else") Then sTemplate = sr.ReadLine
                    End If
                    LogEvent("sTemplate: " & sTemplate)
                    Do While Not sTemplate.Contains("#End If")
                        LogEvent("sTemplate: " & sTemplate)
                        sTemplate = sr.ReadLine
                    Loop
                    If sTemplate.Contains("#End If") Then sTemplate = sr.ReadLine
                    LogEvent("sTemplate: " & sTemplate)
                End If
                If sTemplate.Contains("If rbtnSO.Checked Then") Then
                    'Remove the rbtnSO.Checked part keeping only the Else part
                    'If rbtnSO.Checked Then
                    'Else
                    'End If
                    'Do While sr.Peek() >= 0
                    'Loop
                    'Do While sr.Peek() >= 0
                    'Loop
                End If
                If sTemplate.Contains("If rbtnPart.Checked Then") Then
                    'Remove the rbtnPart.Checked part keeping only the Else part
                    'If rbtnPart.Checked Then
                    'Else
                    'End If
                    'Do While sr.Peek() >= 0
                    'Loop
                    'Do While sr.Peek() >= 0
                    'Loop
                End If

                WriteOutLine(sTemplate, tmpFN)
                'Set the value of ProgressBar
                If Me.ProgressBar1.Value < Me.ProgressBar1.Maximum Then
                    Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
                End If
            Loop
            sr.Close()
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
