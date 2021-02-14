'==============================================================================================================================================================================================
'
'Copyright 2019 Signalworks Pty Ltd, ABN 26 066 681 598

'Licensed under the Apache License, Version 2.0 (the "License");
'you may not use this file except in compliance with the License.
'You may obtain a copy of the License at
'
'http://www.apache.org/licenses/LICENSE-2.0
'
'Unless required by applicable law or agreed to in writing, software
'distributed under the License is distributed on an "AS IS" BASIS,
''WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'See the License for the specific language governing permissions and
'limitations under the License.
'
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Imports System.Windows.Forms.DataVisualization 'Add reference Assemblies, Framework, System.Windows.Forms.DataVisualization
Imports System.ComponentModel
Imports System.Security.Permissions
Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Timers

<PermissionSet(SecurityAction.Demand, Name:="FullTrust")>
<System.Runtime.InteropServices.ComVisibleAttribute(True)>
Public Class Main
    'The ADVL_Point_Chart application plots point data on a chart.

    'Information about the Chart control used to draw the line charts:
    'https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.datavisualization.charting.chart?view=netframework-4.8
    'Line chart custom attributes:
    'https://msdn.microsoft.com/en-us/data/dd489252(v=vs.95)
    'Chart types:
    'https://msdn.microsoft.com/en-us/data/dd489233(v=vs.95)


#Region " Coding Notes - Notes on the code used in this class." '==============================================================================================================================

    'ADD THE SYSTEM UTILITIES REFERENCE: ==========================================================================================
    'The following references are required by this software: 
    'ADVL_Utilities_Library_1.dll
    'To add the reference, press Project \ Add Reference... 
    '  Select the Browse option then press the Browse button
    '  Find the ADVL_Utilities_Library_1.dll file (it should be located in the directory ...\Projects\ADVL_Utilities_Library_1\ADVL_Utilities_Library_1\bin\Debug\)
    '  Press the Add button. Press the OK button.
    'The Utilities Library is used for Project Management, Archive file management, running XSequence files and running XMessage files.
    'If there are problems with a reference, try deleting it from the references list and adding it again.

    'ADD THE SERVICE REFERENCE: ===================================================================================================
    'A service reference to the Message Service must be added to the source code before this service can be used.
    'This is used to connect to the Application Network.

    'Adding the service reference to a project that includes the Message Service project: -----------------------------------------
    'Project \ Add Service Reference
    'Press the Discover button.
    'Expand the items in the Services window and select IMsgService.
    'Press OK.
    '------------------------------------------------------------------------------------------------------------------------------
    '------------------------------------------------------------------------------------------------------------------------------
    'Adding the service reference to other projects that dont include the Message Service project: -------------------------------
    'Run the ADVL_Network_1 application to start the message service.
    'In Microsoft Visual Studio select: Project \ Add Service Reference
    'Enter the address: http://localhost:8734/ADVLService
    'Press the Go button.
    'MsgService is found.
    'Press OK to add ServiceReference1 to the project.
    '------------------------------------------------------------------------------------------------------------------------------
    '
    'ADD THE MsgServiceCallback CODE: =============================================================================================
    'This is used to connect to the Application Network.
    'In Microsoft Visual Studio select: Project \ Add Class
    'MsgServiceCallback.vb
    'Add the following code to the class:
    'Imports System.ServiceModel
    'Public Class MsgServiceCallback
    '    Implements ServiceReference1.IMsgServiceCallback
    '    Public Sub OnSendMessage(message As String) Implements ServiceReference1.IMsgServiceCallback.OnSendMessage
    '        'A message has been received.
    '        'Set the InstrReceived property value to the message (usually in XMessage format). This will also apply the instructions in the XMessage.
    '        Main.InstrReceived = message
    '    End Sub
    'End Class
    '------------------------------------------------------------------------------------------------------------------------------
    '
    'DEBUGGING TIPS:
    '1. If an application based on the Application Template does not initially run correctly,
    '    check that the copied methods, such as Main_Load, have the correct Handles statement.
    '    For example: the Main_Load method should have the following declaration: Private Sub Main_Load(sender As Object, e As EventArgs) Handles Me.Load
    '      It will not run when the application loads, with this declaration:      Private Sub Main_Load(sender As Object, e As EventArgs)
    '    For example: the Main_FormClosing method should have the following declaration: Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
    '      It will not run when the application closes, with this declaration:     Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs)
    '------------------------------------------------------------------------------------------------------------------------------
    '
    'ADD THE Timer1 Control to the Main Form: =====================================================================================
    'Select the Main.vb [Design] tab.
    'Press Toolbox \ Compnents \ Times and add Timer1 to the Main form.
    '------------------------------------------------------------------------------------------------------------------------------
    '
    'EDIT THE DefaultAppProperties() CODE: ========================================================================================
    'This sets the Application properties that are stored in the Application_Info_ADVL_2.xml settings file.
    'The following properties need to be updated:
    '  ApplicationInfo.Name
    '  ApplicationInfo.Description
    '  ApplicationInfo.CreationDate
    '  ApplicationInfo.Author
    '  ApplicationInfo.Copyright
    '  ApplicationInfo.Trademarks
    '  ApplicationInfo.License
    '  ApplicationInfo.SourceCode          (Optional - Preliminary implemetation coded.)
    '  ApplicationInfo.ModificationSummary (Optional - Preliminary implemetation coded.)
    '  ApplicationInfo.Libraries           (Optional - Preliminary implemetation coded.)
    '------------------------------------------------------------------------------------------------------------------------------
    '
    'ADD THE Application Icon: ====================================================================================================
    'Double-click My Project in the Solution Explorer window to open the project tab.
    'In the Application section press the Icon box and selct Browse.
    'Select an application icon.
    'This icon can also be selected for the Main form icon by editing the properties of this form.
    '------------------------------------------------------------------------------------------------------------------------------
    '
    'EDIT THE Application Info Text: ==============================================================================================
    'The Application Info Text is used to label the appllication icon in the Application Network tree view.
    'This is edited in the SendApplicationInfo() method of the Main form.
    'Edit the line of code: Dim text As New XElement("Text", "Application Template").
    'Replace the default text "Application Template" with the required text.
    'Note that this text can be updated at any time and when the updated executable is run, it will update the Application Network tree view the next time it is connected.
    '------------------------------------------------------------------------------------------------------------------------------
    '
    'Calling JavaScript from VB.NET:
    'The following Imports statement and permissions are required for the Main form:
    'Imports System.Security.Permissions
    '<PermissionSet(SecurityAction.Demand, Name:="FullTrust")> _
    '<System.Runtime.InteropServices.ComVisibleAttribute(True)> _
    'NOTE: the line continuation characters (_) will disappear form the code view after they have been typed!
    '------------------------------------------------------------------------------------------------------------------------------
    'Calling VB.NET from JavaScript
    'Add the following line to the Main.Load method:
    '  Me.WebBrowser1.ObjectForScripting = Me
    '------------------------------------------------------------------------------------------------------------------------------

#End Region 'Coding Notes ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Variable Declarations - All the variables and class objects used in this form and this application." '===============================================================================

    Public WithEvents ApplicationInfo As New ADVL_Utilities_Library_1.ApplicationInfo 'This object is used to store application information.
    Public WithEvents Project As New ADVL_Utilities_Library_1.Project 'This object is used to store Project information.
    Public WithEvents Message As New ADVL_Utilities_Library_1.Message 'This object is used to display messages in the Messages window.
    Public WithEvents ApplicationUsage As New ADVL_Utilities_Library_1.Usage 'This object stores application usage information.

    'Declare Forms used by the application:
    Public WithEvents DesignQuery As frmDesignQuery
    Public WithEvents ViewDatabaseData As frmViewDatabaseData
    Public WithEvents WebPageList As frmWebPageList

    Public WithEvents NewHtmlDisplay As frmHtmlDisplay
    Public HtmlDisplayFormList As New ArrayList 'Used for displaying multiple HtmlDisplay forms.

    Public WithEvents NewWebPage As frmWebPage
    Public WebPageFormList As New ArrayList 'Used for displaying multiple WebView forms.

    Public WithEvents ZoomChart As frmZoomChart 'Used to zoom the chart view.
    Public WithEvents MarkerProperties As frmMarkerProperties 'Used to edit the Chart marker properties.
    Public WithEvents ChartInfoForm As frmChartInfoForm 'Displays additional information about a chart.

    'Declare objects used to connect to the Message Service:
    Public client As ServiceReference1.MsgServiceClient
    Public WithEvents XMsg As New ADVL_Utilities_Library_1.XMessage
    Dim XDoc As New System.Xml.XmlDocument
    Public Status As New System.Collections.Specialized.StringCollection
    Dim ClientProNetName As String = "" 'The name of the client Project Network requesting service. 
    Dim ClientAppName As String = "" 'The name of the client requesting service
    Dim ClientConnName As String = "" 'The name of the client connection requesting service
    Dim MessageXDoc As System.Xml.Linq.XDocument
    Dim xmessage As XElement 'This will contain the message. It will be added to MessageXDoc.
    Dim xlocns As New List(Of XElement) 'A list of locations. Each location forms part of the reply message. The information in the reply message will be sent to the specified location in the client application.
    Dim MessageText As String = "" 'The text of a message sent through the Application Network.

    'Dim CompletionInstruction As String = "Stop" 'The last instruction returned on completion of the processing of an XMessage.
    Public OnCompletionInstruction As String = "Stop" 'The last instruction returned on completion of the processing of an XMessage.
    Public EndInstruction As String = "Stop" 'Another method of specifying the last instruction. This is processed in the EndOfSequence section of XMsg.Instructions.


    Public ConnectionName As String = "" 'The name of the connection used to connect this application to the ComNet (Message Service).

    Public ProNetName As String = "" 'The name of the Project Network
    Public ProNetPath As String = "" 'The path of the Project Network

    Public AdvlNetworkAppPath As String = "" 'The application path of the ADVL Network application (ComNet). This is where the "Application.Lock" file will be while ComNet is running
    Public AdvlNetworkExePath As String = "" 'The executable path of the ADVL Network.

    'Variable for local processing of an XMessage:
    Public WithEvents XMsgLocal As New ADVL_Utilities_Library_1.XMessage
    Dim XDocLocal As New System.Xml.XmlDocument
    Public StatusLocal As New System.Collections.Specialized.StringCollection

    'Main.Load variables:
    Dim ProjectSelected As Boolean = False 'If True, a project has been selected using Command Arguments. Used in Main.Load.
    Dim StartupConnectionName As String = "" 'If not "" the application will be connected to the ComNet using this connection name in  Main.Load.

    'The following variables are used to run JavaScript in Web Pages loaded into the Document View: -------------------
    Public WithEvents XSeq As New ADVL_Utilities_Library_1.XSequence
    'To run an XSequence:
    '  XSeq.RunXSequence(xDoc, Status) 'ImportStatus in Import
    '    Handle events:
    '      XSeq.ErrorMsg
    '      XSeq.Instruction(Info, Locn)

    Private XStatus As New System.Collections.Specialized.StringCollection

    'Variables used to restore Item values on a web page.
    Private FormName As String
    Private ItemName As String
    Private SelectId As String

    'StartProject variables:
    Private StartProject_AppName As String  'The application name
    Private StartProject_ConnName As String 'The connection name
    Private StartProject_ProjID As String   'The project ID

    'Public PointChart As New PointChart

    'Dataset used to hold points for plotting:
    'Public ds As New DataSet 'THIS IS NOW STORED IN ChartInfo.

    Private WithEvents bgwComCheck As New System.ComponentModel.BackgroundWorker 'Used to perform communication checks on a separate thread.

    Dim cboFieldSelections As New DataGridViewComboBoxColumn 'Used for selecting Y Value fields in the Chart Settings tab (This is declared here because it may be modified later.)

    Dim WithEvents Zip As ADVL_Utilities_Library_1.ZipComp

    Public WithEvents bgwSendMessage As New System.ComponentModel.BackgroundWorker 'Used to send a message through the Message Service.
    Dim SendMessageParams As New clsSendMessageParams 'This hold the Send Message parameters: .ProjectNetworkName, .ConnectionName & .Message

    'Alternative SendMessage background worker - needed to send a message while instructions are being processed.
    Public WithEvents bgwSendMessageAlt As New System.ComponentModel.BackgroundWorker 'Used to send a message through the Message Service - alternative backgound worker.
    Dim SendMessageParamsAlt As New clsSendMessageParams 'This hold the Send Message parameters: .ProjectNetworkName, .ConnectionName & .Message - for the alternative background worker.

    Public WithEvents ChartInfo As New ChartInfo 'Stores information about the chart. Contains methods to Save, Load and Clear the chart.

    'Chart1 Variables:
    'Chart Pan variables:
    Dim MouseDownXMin As Double
    Dim MouseDownXMax As Double
    Dim MouseDownXRange As Double
    Dim MouseDownXPixels As Integer
    'Dim MouseDownX As Double
    Dim MouseDownXPixel As Integer
    Dim MouseDownYMin As Double
    Dim MouseDownYMax As Double
    Dim MouseDownYRange As Double
    Dim MouseDownYPixels As Integer
    'Dim MouseDownY As Double
    Dim MouseDownYPixel As Integer

    'Dim XPosn As Double
    'Dim YPosn As Double

    'Dim WithEvents PanTimer As New System.Timers.Timer 'Used to update a chart while panning only every 250ms.

    'Selected Points:
    Public SelPtBorderColor As Color = Color.Black
    Public SelPtBorderWidth As Integer = 1
    Public SelPtColor As Color = Color.Red
    Public SelPtSize As Integer = 8
    Public SelPtStyle As MarkerStyle = MarkerStyle.Circle

    Public selPoints As New List(Of DataPoint) 'Use GetPointsInSelectionArea() to select points


#End Region 'Variable Declarations ------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Properties - All the properties used in this form and this application" '============================================================================================================

    Private _connectionHashcode As Integer 'The Message Service connection hashcode. This is used to identify a connection in the Message Service when reconnecting.
    Property ConnectionHashcode As Integer
        Get
            Return _connectionHashcode
        End Get
        Set(value As Integer)
            _connectionHashcode = value
        End Set
    End Property

    Private _connectedToComNet As Boolean = False  'True if the application is connected to the Communication Network (Message Service).
    Property ConnectedToComNet As Boolean
        Get
            Return _connectedToComNet
        End Get
        Set(value As Boolean)
            _connectedToComNet = value
        End Set
    End Property

    Private _instrReceived As String = "" 'Contains Instructions received via the Message Service.
    Property InstrReceived As String
        Get
            Return _instrReceived
        End Get
        Set(value As String)
            If value = Nothing Then
                Message.Add("Empty message received!")
            Else
                _instrReceived = value
                ProcessInstructions(_instrReceived)
            End If
        End Set
    End Property

    Private Sub ProcessInstructions(ByVal Instructions As String)
        'Process the XMessage instructions.

        Dim MsgType As String
        If Instructions.StartsWith("<XMsg>") Then
            MsgType = "XMsg"
            If ShowXMessages Then
                'Add the message header to the XMessages window:
                Message.XAddText("Message received: " & vbCrLf, "XmlReceivedNotice")
            End If
        ElseIf Instructions.StartsWith("<XSys>") Then
            MsgType = "XSys"
            If ShowSysMessages Then
                'Add the message header to the XMessages window:
                Message.XAddText("System Message received: " & vbCrLf, "XmlReceivedNotice")
            End If
        Else
            MsgType = "Unknown"
        End If

        'If ShowXMessages Then
        '    'Add the message header to the XMessages window:
        '    Message.XAddText("Message received: " & vbCrLf, "XmlReceivedNotice")
        'End If

        'If Instructions.StartsWith("<XMsg>") Then 'This is an XMessage set of instructions.
        If MsgType = "XMsg" Or MsgType = "XSys" Then 'This is an XMessage or XSystem set of instructions.
            Try
                'Inititalise the reply message:
                ClientProNetName = ""
                ClientConnName = ""
                ClientAppName = ""
                xlocns.Clear() 'Clear the list of locations in the reply message.
                Dim Decl As New XDeclaration("1.0", "utf-8", "yes")
                MessageXDoc = New XDocument(Decl, Nothing) 'Reply message - this will be sent to the Client App.
                'xmessage = New XElement("XMsg")
                xmessage = New XElement(MsgType)
                xlocns.Add(New XElement("Main")) 'Initially set the location in the Client App to Main.

                'Run the received message:
                Dim XmlHeader As String = "<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>"

                XDoc.LoadXml(XmlHeader & vbCrLf & Instructions.Replace("&", "&amp;")) 'Replace "&" with "&amp:" before loading the XML text.
                'If ShowXMessages Then
                '    Message.XAddXml(XDoc)   'Add the message to the XMessages window.
                '    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                'End If
                If (MsgType = "XMsg") And ShowXMessages Then
                    Message.XAddXml(XDoc)  'Add the message to the XMessages window.
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                ElseIf (MsgType = "XSys") And ShowSysMessages Then
                    Message.XAddXml(XDoc)  'Add the message to the XMessages window.
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                End If
                XMsg.Run(XDoc, Status)
            Catch ex As Exception
                Message.Add("Error running XMsg: " & ex.Message & vbCrLf)
            End Try

            'XMessage has been run.
            'Reply to this message:
            'Add the message reply to the XMessages window:
            'Complete the MessageXDoc:
            xmessage.Add(xlocns(xlocns.Count - 1)) 'Add the last location reply instructions to the message.
            MessageXDoc.Add(xmessage)
            MessageText = MessageXDoc.ToString

            If ClientConnName = "" Then
                'No client to send a message to - process the message locally.
                'If ShowXMessages Then
                '    Message.XAddText("Message processed locally:" & vbCrLf, "XmlSentNotice")
                '    Message.XAddXml(MessageText)
                '    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                'End If
                If (MsgType = "XMsg") And ShowXMessages Then
                    Message.XAddText("Message processed locally:" & vbCrLf, "XmlSentNotice")
                    Message.XAddXml(MessageText)
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                ElseIf (MsgType = "XSys") And ShowSysMessages Then
                    Message.XAddText("System Message processed locally:" & vbCrLf, "XmlSentNotice")
                    Message.XAddXml(MessageText)
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                End If
                ProcessLocalInstructions(MessageText)
            Else
                'If ShowXMessages Then
                '    Message.XAddText("Message sent to [" & ClientProNetName & "]." & ClientConnName & ":" & vbCrLf, "XmlSentNotice")
                '    Message.XAddXml(MessageText)
                '    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                'End If
                If (MsgType = "XMsg") And ShowXMessages Then
                    Message.XAddText("Message sent to [" & ClientProNetName & "]." & ClientConnName & ":" & vbCrLf, "XmlSentNotice")   'NOTE: There is no SendMessage code in the Message Service application!
                    Message.XAddXml(MessageText)
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                ElseIf (MsgType = "XSys") And ShowSysMessages Then
                    Message.XAddText("System Message sent to [" & ClientProNetName & "]." & ClientConnName & ":" & vbCrLf, "XmlSentNotice")   'NOTE: There is no SendMessage code in the Message Service application!
                    Message.XAddXml(MessageText)
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                End If

                'Send Message on a new thread:
                SendMessageParams.ProjectNetworkName = ClientProNetName
                SendMessageParams.ConnectionName = ClientConnName
                SendMessageParams.Message = MessageText
                If bgwSendMessage.IsBusy Then
                    Message.AddWarning("Send Message backgroundworker is busy." & vbCrLf)
                Else
                    bgwSendMessage.RunWorkerAsync(SendMessageParams)
                End If
            End If
        Else 'This is not an XMessage!
            If Instructions.StartsWith("<XMsgBlk>") Then 'This is an XMessageBlock.
                'Process the received message:
                Dim XmlHeader As String = "<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>"
                XDoc.LoadXml(XmlHeader & vbCrLf & Instructions.Replace("&", "&amp;")) 'Replace "&" with "&amp:" before loading the XML text.
                If ShowXMessages Then
                    Message.XAddXml(XDoc)   'Add the message to the XMessages window.
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                End If

                'Process the XMessageBlock:
                Dim XMsgBlkLocn As String
                XMsgBlkLocn = XDoc.GetElementsByTagName("ClientLocn")(0).InnerText
                Select Case XMsgBlkLocn
                    Case "DisplayChart"
                        Dim XData As Xml.XmlNodeList = XDoc.GetElementsByTagName("XInfo")
                        Dim ChartXDoc As New Xml.Linq.XDocument
                        Try
                            ChartXDoc = XDocument.Parse("<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>" & vbCrLf & XData(0).InnerXml)
                            ChartInfo.LoadXml(ChartXDoc, Chart1)
                            txtChartFileName.Text = ""
                            txtSeriesName.Text = ChartXDoc.<ChartSettings>.<SeriesCollection>.<Series>.<Name>.Value
                            UpdateInputDataTabSettings()
                            UpdateTitlesTabSettings()
                            UpdateAreasTabSettings() 'Update Areas Tab before Series Tab. This will update the list of Chart Areas on the Series Tab.
                            UpdateSeriesTabSettings()
                            DrawPointChart()

                        Catch ex As Exception
                            Message.Add(ex.Message & vbCrLf)
                        End Try
                    Case Else
                        Message.AddWarning("Unknown XInfo Message location: " & XMsgBlkLocn & vbCrLf)
                End Select
            Else
                Message.XAddText("The message is not an XMessage or XMessageBlock: " & vbCrLf & Instructions & vbCrLf & vbCrLf, "Normal")
            End If
        End If
    End Sub

    Private Sub ProcessLocalInstructions(ByVal Instructions As String)
        'Process the XMessage instructions locally.

        'If Instructions.StartsWith("<XMsg>") Then 'This is an XMessage set of instructions.
        If Instructions.StartsWith("<XMsg>") Or Instructions.StartsWith("<XSys>") Then 'This is an XMessage set of instructions.
            'Run the received message:
            Dim XmlHeader As String = "<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>"
            XDocLocal.LoadXml(XmlHeader & vbCrLf & Instructions)
            XMsgLocal.Run(XDocLocal, StatusLocal)
        Else 'This is not an XMessage!
            Message.XAddText("The message is not an XMessage: " & Instructions & vbCrLf, "Normal")
        End If
    End Sub

    Private _showXMessages As Boolean = True 'If True, XMessages that are sent or received will be shown in the Messages window.
    Property ShowXMessages As Boolean
        Get
            Return _showXMessages
        End Get
        Set(value As Boolean)
            _showXMessages = value
        End Set
    End Property

    Private _showSysMessages As Boolean = True 'If True, System messages that are sent or received will be shown in the messages window.
    Property ShowSysMessages As Boolean
        Get
            Return _showSysMessages
        End Get
        Set(value As Boolean)
            _showSysMessages = value
        End Set
    End Property

    Private _closedFormNo As Integer 'Temporarily holds the number of the form that is being closed. 
    Property ClosedFormNo As Integer
        Get
            Return _closedFormNo
        End Get
        Set(value As Integer)
            _closedFormNo = value
        End Set
    End Property

    Private _inputDatabaseDirectory As String = "" 'The directory of the Input Database. When the Find database button is pressed, the Open File Dialog will open in this directory.
    Property InputDatabaseDirectory As String
        Get
            Return _inputDatabaseDirectory
        End Get
        Set(value As String)
            _inputDatabaseDirectory = value
        End Set
    End Property

    Private _chartWindow As String = "Preview" '(Preview or New Window) Chart can be drawn in the Preview window or a New Window.
    Property ChartWindow As String
        Get
            Return _chartWindow
        End Get
        Set(value As String)
            _chartWindow = value
            If _chartWindow = "Preview" Then
                rbPreviewChart.Checked = True
            ElseIf _chartWindow = "New Window" Then
                rbNewWindowChart.Checked = True
            End If
        End Set
    End Property

    Private _workflowFileName As String = "" 'The file name of the html document displayed in the Workflow tab.
    Public Property WorkflowFileName As String
        Get
            Return _workflowFileName
        End Get
        Set(value As String)
            _workflowFileName = value
        End Set
    End Property

    'Private _chartMode As String = "Select" 'The mouse mode for the chart: Select - crosshair pointer Pan - Hand pointer
    'Public Property ChartMode As String
    '    Get
    '        Return _chartMode
    '    End Get
    '    Set(value As String)
    '        _chartMode = value
    '    End Set
    'End Property

    'Private _selSeries As String = "" 'The selected Chart Series.
    'Property SelSeries As String
    '    Get
    '        Return _selSeries
    '    End Get
    '    Set(value As String)
    '        _selSeries = value
    '    End Set
    'End Property

    'Private _selChartArea As String = "" 'The selected Chart Area.
    'Property SelChartArea As String
    '    Get
    '        Return _selChartArea
    '    End Get
    '    Set(value As String)
    '        _selChartArea = value
    '    End Set
    'End Property


#End Region 'Properties -----------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Process XML Files - Read and write XML files." '=====================================================================================================================================

    Private Sub SaveFormSettings()
        'Save the form settings in an XML document.
        Dim settingsData = <?xml version="1.0" encoding="utf-8"?>
                           <!---->
                           <!--Form settings for Main form.-->
                           <FormSettings>
                               <Left><%= Me.Left %></Left>
                               <Top><%= Me.Top %></Top>
                               <Width><%= Me.Width %></Width>
                               <Height><%= Me.Height %></Height>
                               <AdvlNetworkAppPath><%= AdvlNetworkAppPath %></AdvlNetworkAppPath>
                               <AdvlNetworkExePath><%= AdvlNetworkExePath %></AdvlNetworkExePath>
                               <ShowXMessages><%= ShowXMessages %></ShowXMessages>
                               <ShowSysMessages><%= ShowSysMessages %></ShowSysMessages>
                               <!---->
                               <SelectedTabIndex><%= TabControl1.SelectedIndex %></SelectedTabIndex>
                           </FormSettings>

        'Add code to include other settings to save after the comment line <!---->

        Dim SettingsFileName As String = "FormSettings_" & ApplicationInfo.Name & " - Main.xml"
        Project.SaveXmlSettings(SettingsFileName, settingsData)
    End Sub

    Private Sub RestoreFormSettings()
        'Read the form settings from an XML document.

        Dim SettingsFileName As String = "FormSettings_" & ApplicationInfo.Name & " - Main.xml"

        If Project.SettingsFileExists(SettingsFileName) Then
            Dim Settings As System.Xml.Linq.XDocument
            Project.ReadXmlSettings(SettingsFileName, Settings)

            If IsNothing(Settings) Then 'There is no Settings XML data.
                Exit Sub
            End If

            'Restore form position and size:
            If Settings.<FormSettings>.<Left>.Value <> Nothing Then Me.Left = Settings.<FormSettings>.<Left>.Value
            If Settings.<FormSettings>.<Top>.Value <> Nothing Then Me.Top = Settings.<FormSettings>.<Top>.Value
            If Settings.<FormSettings>.<Height>.Value <> Nothing Then Me.Height = Settings.<FormSettings>.<Height>.Value
            If Settings.<FormSettings>.<Width>.Value <> Nothing Then Me.Width = Settings.<FormSettings>.<Width>.Value

            If Settings.<FormSettings>.<AdvlNetworkAppPath>.Value <> Nothing Then AdvlNetworkAppPath = Settings.<FormSettings>.<AdvlNetworkAppPath>.Value
            If Settings.<FormSettings>.<AdvlNetworkExePath>.Value <> Nothing Then AdvlNetworkExePath = Settings.<FormSettings>.<AdvlNetworkExePath>.Value

            If Settings.<FormSettings>.<ShowXMessages>.Value <> Nothing Then ShowXMessages = Settings.<FormSettings>.<ShowXMessages>.Value
            If Settings.<FormSettings>.<ShowSysMessages>.Value <> Nothing Then ShowSysMessages = Settings.<FormSettings>.<ShowSysMessages>.Value

            'Add code to read other saved setting here:
            If Settings.<FormSettings>.<SelectedTabIndex>.Value <> Nothing Then TabControl1.SelectedIndex = Settings.<FormSettings>.<SelectedTabIndex>.Value

            CheckFormPos()
        End If
    End Sub

    Private Sub CheckFormPos()
        'Check that the form can be seen on a screen.

        Dim MinWidthVisible As Integer = 192 'Minimum number of X pixels visible. The form will be moved if this many form pixels are not visible.
        Dim MinHeightVisible As Integer = 64 'Minimum number of Y pixels visible. The form will be moved if this many form pixels are not visible.

        Dim FormRect As New Rectangle(Me.Left, Me.Top, Me.Width, Me.Height)
        Dim WARect As Rectangle = Screen.GetWorkingArea(FormRect) 'The Working Area rectangle - the usable area of the screen containing the form.

        'Check if the top of the form is above the top of the Working Area:
        If Me.Top < WARect.Top Then
            Me.Top = WARect.Top
        End If

        'Check if the top of the form is too close to the bottom of the Working Area:
        If (Me.Top + MinHeightVisible) > (WARect.Top + WARect.Height) Then
            Me.Top = WARect.Top + WARect.Height - MinHeightVisible
        End If

        'Check if the left edge of the form is too close to the right edge of the Working Area:
        If (Me.Left + MinWidthVisible) > (WARect.Left + WARect.Width) Then
            Me.Left = WARect.Left + WARect.Width - MinWidthVisible
        End If

        'Check if the right edge of the form is too close to the left edge of the Working Area:
        If (Me.Left + Me.Width - MinWidthVisible) < WARect.Left Then
            Me.Left = WARect.Left - Me.Width + MinWidthVisible
        End If
    End Sub

    Private Sub ReadApplicationInfo()
        'Read the Application Information.

        If ApplicationInfo.FileExists Then
            ApplicationInfo.ReadFile()
        Else
            'There is no Application_Info_ADVL_2.xml file.
            DefaultAppProperties() 'Create a new Application Info file with default application properties:
            ApplicationInfo.WriteFile() 'Write the file now. The file information may be used by other applications.
        End If
    End Sub

    Private Sub DefaultAppProperties()
        'These properties will be saved in the Application_Info.xml file in the application directory.
        'If this file is deleted, it will be re-created using these default application properties.

        'Change this to show your application Name, Description and Creation Date.
        ApplicationInfo.Name = "ADVL_Point_Chart_1"

        'ApplicationInfo.ApplicationDir is set when the application is started.
        ApplicationInfo.ExecutablePath = Application.ExecutablePath
        ApplicationInfo.Description = "The Point Chart application plots point data on a chart."
        ApplicationInfo.CreationDate = "16-Jun-2019 20:44:00"

        'Author -----------------------------------------------------------------------------------------------------------
        'Change this to show your Name, Description and Contact information.
        ApplicationInfo.Author.Name = "Signalworks Pty Ltd"
        ApplicationInfo.Author.Description = "Signalworks Pty Ltd" & vbCrLf &
            "Australian Proprietary Company" & vbCrLf &
            "ABN 26 066 681 598" & vbCrLf &
            "Registration Date 05/10/1994"

        ApplicationInfo.Author.Contact = "http://www.andorville.com.au/"

        'File Associations: -----------------------------------------------------------------------------------------------
        'Add any file associations here.
        'The file extension and a description of files that can be opened by this application are specified.
        'The example below specifies a coordinate system parameter file type with the file extension .ADVLCoord.
        'Dim Assn1 As New ADVL_System_Utilities.FileAssociation
        'Assn1.Extension = "ADVLCoord"
        'Assn1.Description = "Andorville™ software coordinate system parameter file"
        'ApplicationInfo.FileAssociations.Add(Assn1)

        'Version ----------------------------------------------------------------------------------------------------------
        ApplicationInfo.Version.Major = My.Application.Info.Version.Major
        ApplicationInfo.Version.Minor = My.Application.Info.Version.Minor
        ApplicationInfo.Version.Build = My.Application.Info.Version.Build
        ApplicationInfo.Version.Revision = My.Application.Info.Version.Revision

        'Copyright --------------------------------------------------------------------------------------------------------
        'Add your copyright information here.
        ApplicationInfo.Copyright.OwnerName = "Signalworks Pty Ltd, ABN 26 066 681 598"
        ApplicationInfo.Copyright.PublicationYear = "2019"

        'Trademarks -------------------------------------------------------------------------------------------------------
        'Add your trademark information here.
        Dim Trademark1 As New ADVL_Utilities_Library_1.Trademark
        Trademark1.OwnerName = "Signalworks Pty Ltd, ABN 26 066 681 598"
        Trademark1.Text = "Andorville"
        Trademark1.Registered = False
        Trademark1.GenericTerm = "software"
        ApplicationInfo.Trademarks.Add(Trademark1)
        Dim Trademark2 As New ADVL_Utilities_Library_1.Trademark
        Trademark2.OwnerName = "Signalworks Pty Ltd, ABN 26 066 681 598"
        Trademark2.Text = "AL-H7"
        Trademark2.Registered = False
        Trademark2.GenericTerm = "software"
        ApplicationInfo.Trademarks.Add(Trademark2)
        Dim Trademark3 As New ADVL_Utilities_Library_1.Trademark
        Trademark3.OwnerName = "Signalworks Pty Ltd, ABN 26 066 681 598"
        Trademark3.Text = "AL-M7"
        Trademark3.Registered = False
        Trademark3.GenericTerm = "software"
        ApplicationInfo.Trademarks.Add(Trademark3)

        'License -------------------------------------------------------------------------------------------------------
        'Add your license information here.
        ApplicationInfo.License.CopyrightOwnerName = "Signalworks Pty Ltd, ABN 26 066 681 598"
        ApplicationInfo.License.PublicationYear = "2019"

        'License Links:
        'http://choosealicense.com/
        'http://www.apache.org/licenses/
        'http://opensource.org/

        'Apache License 2.0 ---------------------------------------------
        ApplicationInfo.License.Code = ADVL_Utilities_Library_1.License.Codes.Apache_License_2_0
        ApplicationInfo.License.Notice = ApplicationInfo.License.ApacheLicenseNotice 'Get the pre-defined Aapche license notice.
        ApplicationInfo.License.Text = ApplicationInfo.License.ApacheLicenseText     'Get the pre-defined Apache license text.

        'Code to use other pre-defined license types is shown below:

        'GNU General Public License, version 3 --------------------------
        'ApplicationInfo.License.Type = ADVL_Utilities_Library_1.License.Types.GNU_GPL_V3_0
        'ApplicationInfo.License.Notice = 'Add the License Notice to ADVL_Utilities_Library_1 License class.
        'ApplicationInfo.License.Text = 'Add the License Text to ADVL_Utilities_Library_1 License class.

        'The MIT License ------------------------------------------------
        'ApplicationInfo.License.Type = ADVL_Utilities_Library_1.License.Types.MIT_License
        'ApplicationInfo.License.Notice = ApplicationInfo.License.MITLicenseNotice
        'ApplicationInfo.License.Text = ApplicationInfo.License.MITLicenseText

        'No License Specified -------------------------------------------
        'ApplicationInfo.License.Type = ADVL_Utilities_Library_1.License.Types.None
        'ApplicationInfo.License.Notice = ""
        'ApplicationInfo.License.Text = ""

        'The Unlicense --------------------------------------------------
        'ApplicationInfo.License.Type = ADVL_Utilities_Library_1.License.Types.The_Unlicense
        'ApplicationInfo.License.Notice = ApplicationInfo.License.UnLicenseNotice
        'ApplicationInfo.License.Text = ApplicationInfo.License.UnLicenseText

        'Unknown License ------------------------------------------------
        'ApplicationInfo.License.Type = ADVL_Utilities_Library_1.License.Types.Unknown
        'ApplicationInfo.License.Notice = ""
        'ApplicationInfo.License.Text = ""

        'Source Code: --------------------------------------------------------------------------------------------------
        'Add your source code information here if required.
        'THIS SECTION WILL BE UPDATED TO ALLOW A GITHUB LINK.
        ApplicationInfo.SourceCode.Language = "Visual Basic 2015"
        ApplicationInfo.SourceCode.FileName = ""
        ApplicationInfo.SourceCode.FileSize = 0
        ApplicationInfo.SourceCode.FileHash = ""
        ApplicationInfo.SourceCode.WebLink = ""
        ApplicationInfo.SourceCode.Contact = ""
        ApplicationInfo.SourceCode.Comments = ""

        'ModificationSummary: -----------------------------------------------------------------------------------------
        'Add any source code modification here is required.
        ApplicationInfo.ModificationSummary.BaseCodeName = ""
        ApplicationInfo.ModificationSummary.BaseCodeDescription = ""
        ApplicationInfo.ModificationSummary.BaseCodeVersion.Major = 0
        ApplicationInfo.ModificationSummary.BaseCodeVersion.Minor = 0
        ApplicationInfo.ModificationSummary.BaseCodeVersion.Build = 0
        ApplicationInfo.ModificationSummary.BaseCodeVersion.Revision = 0
        ApplicationInfo.ModificationSummary.Description = "This is the first released version of the application. No earlier base code used."

        'Library List: ------------------------------------------------------------------------------------------------
        'Add the ADVL_Utilties_Library_1 library:
        Dim NewLib As New ADVL_Utilities_Library_1.LibrarySummary
        NewLib.Name = "ADVL_System_Utilities"
        NewLib.Description = "System Utility classes used in Andorville™ software development system applications"
        NewLib.CreationDate = "7-Jan-2016 12:00:00"
        NewLib.LicenseNotice = "Copyright 2016 Signalworks Pty Ltd, ABN 26 066 681 598" & vbCrLf &
                               vbCrLf &
                               "Licensed under the Apache License, Version 2.0 (the ""License"");" & vbCrLf &
                               "you may not use this file except in compliance with the License." & vbCrLf &
                               "You may obtain a copy of the License at" & vbCrLf &
                               vbCrLf &
                               "http://www.apache.org/licenses/LICENSE-2.0" & vbCrLf &
                               vbCrLf &
                               "Unless required by applicable law or agreed to in writing, software" & vbCrLf &
                               "distributed under the License is distributed on an ""AS IS"" BASIS," & vbCrLf &
                               "WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied." & vbCrLf &
                               "See the License for the specific language governing permissions and" & vbCrLf &
                               "limitations under the License." & vbCrLf

        NewLib.CopyrightNotice = "Copyright 2016 Signalworks Pty Ltd, ABN 26 066 681 598"

        NewLib.Version.Major = 1
        NewLib.Version.Minor = 0
        NewLib.Version.Build = 1
        NewLib.Version.Revision = 0

        NewLib.Author.Name = "Signalworks Pty Ltd"
        NewLib.Author.Description = "Signalworks Pty Ltd" & vbCrLf &
            "Australian Proprietary Company" & vbCrLf &
            "ABN 26 066 681 598" & vbCrLf &
            "Registration Date 05/10/1994"

        NewLib.Author.Contact = "http://www.andorville.com.au/"

        Dim NewClass1 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass1.Name = "ZipComp"
        NewClass1.Description = "The ZipComp class is used to compress files into and extract files from a zip file."
        NewLib.Classes.Add(NewClass1)
        Dim NewClass2 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass2.Name = "XSequence"
        NewClass2.Description = "The XSequence class is used to run an XML property sequence (XSequence) file. XSequence files are used to record and replay processing sequences in Andorville™ software applications."
        NewLib.Classes.Add(NewClass2)
        Dim NewClass3 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass3.Name = "XMessage"
        NewClass3.Description = "The XMessage class is used to read an XML Message (XMessage). An XMessage is a simplified XSequence used to exchange information between Andorville™ software applications."
        NewLib.Classes.Add(NewClass3)
        Dim NewClass4 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass4.Name = "Location"
        NewClass4.Description = "The Location class consists of properties and methods to store data in a location, which is either a directory or archive file."
        NewLib.Classes.Add(NewClass4)
        Dim NewClass5 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass5.Name = "Project"
        NewClass5.Description = "An Andorville™ software application can store data within one or more projects. Each project stores a set of related data files. The Project class contains properties and methods used to manage a project."
        NewLib.Classes.Add(NewClass5)
        Dim NewClass6 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass6.Name = "ProjectSummary"
        NewClass6.Description = "ProjectSummary stores a summary of a project."
        NewLib.Classes.Add(NewClass6)
        Dim NewClass7 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass7.Name = "DataFileInfo"
        NewClass7.Description = "The DataFileInfo class stores information about a data file."
        NewLib.Classes.Add(NewClass7)
        Dim NewClass8 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass8.Name = "Message"
        NewClass8.Description = "The Message class contains text properties and methods used to display messages in an Andorville™ software application."
        NewLib.Classes.Add(NewClass8)
        Dim NewClass9 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass9.Name = "ApplicationSummary"
        NewClass9.Description = "The ApplicationSummary class stores a summary of an Andorville™ software application."
        NewLib.Classes.Add(NewClass9)
        Dim NewClass10 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass10.Name = "LibrarySummary"
        NewClass10.Description = "The LibrarySummary class stores a summary of a software library used by an application."
        NewLib.Classes.Add(NewClass10)
        Dim NewClass11 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass11.Name = "ClassSummary"
        NewClass11.Description = "The ClassSummary class stores a summary of a class contained in a software library."
        NewLib.Classes.Add(NewClass11)
        Dim NewClass12 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass12.Name = "ModificationSummary"
        NewClass12.Description = "The ModificationSummary class stores a summary of any modifications made to an application or library."
        NewLib.Classes.Add(NewClass12)
        Dim NewClass13 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass13.Name = "ApplicationInfo"
        NewClass13.Description = "The ApplicationInfo class stores information about an Andorville™ software application."
        NewLib.Classes.Add(NewClass13)
        Dim NewClass14 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass14.Name = "Version"
        NewClass14.Description = "The Version class stores application, library or project version information."
        NewLib.Classes.Add(NewClass14)
        Dim NewClass15 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass15.Name = "Author"
        NewClass15.Description = "The Author class stores information about an Author."
        NewLib.Classes.Add(NewClass15)
        Dim NewClass16 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass16.Name = "FileAssociation"
        NewClass16.Description = "The FileAssociation class stores the file association extension and description. An application can open files on its file association list."
        NewLib.Classes.Add(NewClass16)
        Dim NewClass17 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass17.Name = "Copyright"
        NewClass17.Description = "The Copyright class stores copyright information."
        NewLib.Classes.Add(NewClass17)
        Dim NewClass18 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass18.Name = "License"
        NewClass18.Description = "The License class stores license information."
        NewLib.Classes.Add(NewClass18)
        Dim NewClass19 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass19.Name = "SourceCode"
        NewClass19.Description = "The SourceCode class stores information about the source code for the application."
        NewLib.Classes.Add(NewClass19)
        Dim NewClass20 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass20.Name = "Usage"
        NewClass20.Description = "The Usage class stores information about application or project usage."
        NewLib.Classes.Add(NewClass20)
        Dim NewClass21 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass21.Name = "Trademark"
        NewClass21.Description = "The Trademark class stored information about a trademark used by the author of an application or data."
        NewLib.Classes.Add(NewClass21)

        ApplicationInfo.Libraries.Add(NewLib)

        'Add other library information here: --------------------------------------------------------------------------

    End Sub

    'Save the form settings if the form is being minimised:
    Protected Overrides Sub WndProc(ByRef m As Message)
        If m.Msg = &H112 Then 'SysCommand
            If m.WParam.ToInt32 = &HF020 Then 'Form is being minimised
                SaveFormSettings()
            End If
        End If
        MyBase.WndProc(m)
    End Sub

    Private Sub SaveProjectSettings()
        'Save the project settings in an XML file.

        'Add any Project Settings to be saved into the settingsData XDocument.
        Dim settingsData = <?xml version="1.0" encoding="utf-8"?>
                           <!---->
                           <!--Project settings for ADVL_Stock_Chart_1 application.-->
                           <ProjectSettings>
                               <!--Plot Settings:-->
                               <ChartWindow><%= ChartWindow %></ChartWindow>
                               <AutoDraw><%= chkAutoDraw.Checked %></AutoDraw>
                               <ChartFileName><%= txtChartFileName.Text %></ChartFileName>
                           </ProjectSettings>

        Dim SettingsFileName As String = "ProjectSettings_" & ApplicationInfo.Name & ".xml"
        Project.SaveXmlSettings(SettingsFileName, settingsData)

        'Dim SettingsFileName As String = "ProjectSettings_" & ApplicationInfo.Name & "_" & ".xml"
        'Project.SaveXmlSettings(SettingsFileName, PointChart.ToXDoc)

        'Save the last chart:
        'Project.SaveXmlData("LastChart.xml", PointChart.ToXDoc)

    End Sub

    Private Sub RestoreProjectSettings()
        'Restore the project settings from an XML document.

        'Dim SettingsFileName As String = "ProjectSettings_" & ApplicationInfo.Name & "_" & ".xml"
        Dim SettingsFileName As String = "ProjectSettings_" & ApplicationInfo.Name & ".xml"

        If Project.SettingsFileExists(SettingsFileName) Then
            Dim Settings As System.Xml.Linq.XDocument
            Project.ReadXmlSettings(SettingsFileName, Settings)

            If IsNothing(Settings) Then 'There is no Settings XML data.
                'Exit Sub
            Else
                ''Restore a Project Setting example:
                'If Settings.<ProjectSettings>.<Setting1>.Value = Nothing Then
                '    'Project setting not saved.
                '    'Setting1 = ""
                'Else
                '    'Setting1 = Settings.<ProjectSettings>.<Setting1>.Value
                'End If

                'Restore Plot Settings: ==============================================================================================
                If Settings.<ProjectSettings>.<ChartWindow>.Value <> Nothing Then ChartWindow = Settings.<ProjectSettings>.<ChartWindow>.Value
                If Settings.<ProjectSettings>.<AutoDraw>.Value <> Nothing Then chkAutoDraw.Checked = Settings.<ProjectSettings>.<AutoDraw>.Value
                If Settings.<ProjectSettings>.<ChartFileName>.Value <> Nothing Then txtChartFileName.Text = Settings.<ProjectSettings>.<ChartFileName>.Value
                '---------------------------------------------------------------------------------------------------------------------

                'Continue restoring saved settings.

            End If
        End If

        ''Restore the last chart:
        'PointChart.LoadFile("LastChart.xml")
        'PointChart.FileName = txtChartFileName.Text 'When LastChart.xml was loaded, the FileName would have been set to LastChart
        'txtDatabasePath.Text = PointChart.InputDatabasePath
        'FillLstTables()
        'If PointChart.InputDatabasePath = "" Then
        '    InputDatabaseDirectory = ""
        'Else
        '    InputDatabaseDirectory = System.IO.Path.GetDirectoryName(PointChart.InputDatabasePath)
        'End If

        'txtDataDescription.Text = PointChart.InputDataDescr
        'txtInputQuery.Text = PointChart.InputQuery
        'ApplyQuery()
        'GetChartOptionsFromDataset() 'This updates the chart options using Dataset fields.
        'UpdatePointChartForm() 'Update the Chart Type, Titles, X Axis and Y Axis tabs with the settings stored in PointChart. 'APPLYING THIS NOW PRODUCES AN ERROR!

    End Sub

#End Region 'Process XML Files ----------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Form Display Methods - Code used to display this form." '============================================================================================================================

    Private Sub Main_Load(sender As Object, e As EventArgs) Handles Me.Load
        'Loading the Main form.

        'Set the Application Directory path: ------------------------------------------------
        Project.ApplicationDir = My.Application.Info.DirectoryPath.ToString

        'Read the Application Information file: ---------------------------------------------
        ApplicationInfo.ApplicationDir = My.Application.Info.DirectoryPath.ToString 'Set the Application Directory property
        'Get the Application Version Information:
        ApplicationInfo.Version.Major = My.Application.Info.Version.Major
        ApplicationInfo.Version.Minor = My.Application.Info.Version.Minor
        ApplicationInfo.Version.Build = My.Application.Info.Version.Build
        ApplicationInfo.Version.Revision = My.Application.Info.Version.Revision

        If ApplicationInfo.ApplicationLocked Then
            MessageBox.Show("The application is locked. If the application is not already in use, remove the 'Application_Info.lock file from the application directory: " & ApplicationInfo.ApplicationDir, "Notice", MessageBoxButtons.OK)
            Dim dr As System.Windows.Forms.DialogResult
            dr = MessageBox.Show("Press 'Yes' to unlock the application", "Notice", MessageBoxButtons.YesNo)
            If dr = System.Windows.Forms.DialogResult.Yes Then
                ApplicationInfo.UnlockApplication()
            Else
                Application.Exit()
                Exit Sub
            End If
        End If

        ReadApplicationInfo()

        'Read the Application Usage information: --------------------------------------------
        ApplicationUsage.StartTime = Now
        ApplicationUsage.SaveLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
        ApplicationUsage.SaveLocn.Path = Project.ApplicationDir
        ApplicationUsage.RestoreUsageInfo()

        'Restore Project information: -------------------------------------------------------
        'Project.ApplicationName = ApplicationInfo.Name
        Project.Application.Name = ApplicationInfo.Name

        'Project.ReadProjectInfoFile()

        'Set up Message object:
        Message.ApplicationName = ApplicationInfo.Name

        'Set up a temporary initial settings location:
        Dim TempLocn As New ADVL_Utilities_Library_1.FileLocation
        TempLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
        TempLocn.Path = ApplicationInfo.ApplicationDir
        Message.SettingsLocn = TempLocn

        Me.Show() 'Show this form before showing the Message form - This will show the App icon on top in the TaskBar.

        'Start showing messages here - Message system is set up.
        'Message.AddText("------------------- Starting Application: ADVL Application Template ----------------- " & vbCrLf, "Heading")
        Message.AddText("------------------- Starting Application: ADVL Point Chart -------------------------- " & vbCrLf, "Heading")
        Message.AddText("Application usage: Total duration = " & Format(ApplicationUsage.TotalDuration.TotalHours, "#.##") & " hours" & vbCrLf, "Normal")

        'https://msdn.microsoft.com/en-us/library/z2d603cy(v=vs.80).aspx#Y550
        'Process any command line arguments:
        Try
            For Each s As String In My.Application.CommandLineArgs
                Message.Add("Command line argument: " & vbCrLf)
                Message.AddXml(s & vbCrLf & vbCrLf)
                InstrReceived = s
            Next
        Catch ex As Exception
            Message.AddWarning("Error processing command line arguments: " & ex.Message & vbCrLf)
        End Try

        If ProjectSelected = False Then
            'Read the Settings Location for the last project used:
            Project.ReadLastProjectInfo()
            'The Last_Project_Info.xml file contains:
            '  Project Name and Description. Settings Location Type and Settings Location Path.
            Message.Add("Last project details:" & vbCrLf)
            Message.Add("Project Type:  " & Project.Type.ToString & vbCrLf)
            Message.Add("Project Path:  " & Project.Path & vbCrLf)

            'At this point read the application start arguments, if any.
            'The selected project may be changed here.

            'Check if the project is locked:
            If Project.ProjectLocked Then
                Message.AddWarning("The project is locked: " & Project.Name & vbCrLf)
                Dim dr As System.Windows.Forms.DialogResult
                dr = MessageBox.Show("Press 'Yes' to unlock the project", "Notice", MessageBoxButtons.YesNo)
                If dr = System.Windows.Forms.DialogResult.Yes Then
                    Project.UnlockProject()
                    Message.AddWarning("The project has been unlocked: " & Project.Name & vbCrLf)
                    'Read the Project Information file: -------------------------------------------------
                    Message.Add("Reading project info." & vbCrLf)
                    Project.ReadProjectInfoFile()                 'Read the file in the SettingsLocation: ADVL_Project_Info.xml

                    Project.ReadParameters()
                    Project.ReadParentParameters()
                    If Project.ParentParameterExists("ProNetName") Then
                        Project.AddParameter("ProNetName", Project.ParentParameter("ProNetName").Value, Project.ParentParameter("ProNetName").Description) 'AddParameter will update the parameter if it already exists.
                        ProNetName = Project.Parameter("ProNetName").Value
                    Else
                        ProNetName = Project.GetParameter("ProNetName")
                    End If
                    If Project.ParentParameterExists("ProNetPath") Then 'Get the parent parameter value - it may have been updated.
                        Project.AddParameter("ProNetPath", Project.ParentParameter("ProNetPath").Value, Project.ParentParameter("ProNetPath").Description) 'AddParameter will update the parameter if it already exists.
                        ProNetPath = Project.Parameter("ProNetPath").Value
                    Else
                        ProNetPath = Project.GetParameter("ProNetPath") 'If the parameter does not exist, the value is set to ""
                    End If
                    Project.SaveParameters() 'These should be saved now - child projects look for parent parameters in the parameter file.

                    Project.LockProject() 'Lock the project while it is open in this application.
                    'Set the project start time. This is used to track project usage.
                    Project.Usage.StartTime = Now
                    ApplicationInfo.SettingsLocn = Project.SettingsLocn
                    'Set up the Message object:
                    Message.SettingsLocn = Project.SettingsLocn
                    Message.Show() 'Added 18May19
                Else
                    'Continue without any project selected.
                    Project.Name = ""
                    Project.Type = ADVL_Utilities_Library_1.Project.Types.None
                    Project.Description = ""
                    Project.SettingsLocn.Path = ""
                    Project.DataLocn.Path = ""
                End If
            Else
                'Read the Project Information file: -------------------------------------------------
                Message.Add("Reading project info." & vbCrLf)
                Project.ReadProjectInfoFile()  'Read the file in the Project Location: ADVL_Project_Info.xml

                Project.ReadParameters()
                Project.ReadParentParameters()
                If Project.ParentParameterExists("ProNetName") Then
                    Project.AddParameter("ProNetName", Project.ParentParameter("ProNetName").Value, Project.ParentParameter("ProNetName").Description) 'AddParameter will update the parameter if it already exists.
                    ProNetName = Project.Parameter("ProNetName").Value
                Else
                    ProNetName = Project.GetParameter("ProNetName")
                End If
                If Project.ParentParameterExists("ProNetPath") Then 'Get the parent parameter value - it may have been updated.
                    Project.AddParameter("ProNetPath", Project.ParentParameter("ProNetPath").Value, Project.ParentParameter("ProNetPath").Description) 'AddParameter will update the parameter if it already exists.
                    ProNetPath = Project.Parameter("ProNetPath").Value
                Else
                    ProNetPath = Project.GetParameter("ProNetPath") 'If the parameter does not exist, the value is set to ""
                End If
                Project.SaveParameters() 'These should be saved now - child projects look for parent parameters in the parameter file.

                Project.LockProject() 'Lock the project while it is open in this application.
                'Set the project start time. This is used to track project usage.
                Project.Usage.StartTime = Now
                ApplicationInfo.SettingsLocn = Project.SettingsLocn
                'Set up the Message object:
                Message.SettingsLocn = Project.SettingsLocn
                Message.Show() 'Added 18May19
            End If

        Else  'Project has been opened using Command Line arguments.
            Project.ReadParameters()
            Project.ReadParentParameters()
            If Project.ParentParameterExists("ProNetName") Then
                Project.AddParameter("ProNetName", Project.ParentParameter("ProNetName").Value, Project.ParentParameter("ProNetName").Description) 'AddParameter will update the parameter if it already exists.
                ProNetName = Project.Parameter("ProNetName").Value
            Else
                ProNetName = Project.GetParameter("ProNetName")
            End If
            If Project.ParentParameterExists("ProNetPath") Then 'Get the parent parameter value - it may have been updated.
                Project.AddParameter("ProNetPath", Project.ParentParameter("ProNetPath").Value, Project.ParentParameter("ProNetPath").Description) 'AddParameter will update the parameter if it already exists.
                ProNetPath = Project.Parameter("ProNetPath").Value
            Else
                ProNetPath = Project.GetParameter("ProNetPath") 'If the parameter does not exist, the value is set to ""
            End If
            Project.SaveParameters() 'These should be saved now - child projects look for parent parameters in the parameter file.

            Project.LockProject() 'Lock the project while it is open in this application.
            ProjectSelected = False 'Reset the Project Selected flag.

            'Set up the Message object:
            Message.SettingsLocn = Project.SettingsLocn
            Message.Show() 'Added 18May19
        End If

        'START Initialise the form: ===============================================================

        Me.WebBrowser1.ObjectForScripting = Me
        'IF THE LINE ABOVE PRODUCES AN ERROR ON STARTUP, CHECK THAT THE CODE ON THE FOLLOWING THREE LINES IS INSERTED JUST ABOVE THE Public Class Main STATEMENT.
        'Imports System.Security.Permissions
        '<PermissionSet(SecurityAction.Demand, Name:="FullTrust")>
        '<System.Runtime.InteropServices.ComVisibleAttribute(True)>


        'Initialise Input Data Tab ----------------------------------------------------------
        cmbDatabaseType.Items.Add("Access2007To2013")
        cmbDatabaseType.SelectedIndex = 0 'Select the first item

        'Initialise Titles Tab --------------------------------------------------------------
        Dim items As Array
        items = System.Enum.GetNames(GetType(ContentAlignment))
        Dim item As String
        For Each item In items
            cmbAlignment.Items.Add(item)
        Next

        items = System.Enum.GetNames(GetType(DataVisualization.Charting.TextOrientation))
        For Each item In items
            cmbOrientation.Items.Add(item)
        Next


        'Initialise Chart Settings Tab ------------------------------------------------------

        '  Titles Tab
        txtTitlesRecordNo.Text = "0"
        txtNTitlesRecords.Text = "0"
        '  Series Tab
        txtSeriesRecordNo.Text = "0"
        txtNSeriesRecords.Text = "0"
        '  Areas Tab
        txtAreaRecordNo.Text = "0"
        txtNAreaRecords.Text = "0"


        'Set up the Y Values grid:
        DataGridView1.ColumnCount = 1
        DataGridView1.RowCount = 1
        DataGridView1.Columns(0).HeaderText = "Y Value"
        DataGridView1.Columns(0).Width = 120
        'Dim cboFieldSelections As New DataGridViewComboBoxColumn 'Used for selecting Y Value fields in the Chart Settings tab (This is declared in the Variables Decl section obecause it may be modified later.)
        DataGridView1.Columns.Insert(1, cboFieldSelections)
        DataGridView1.Columns(1).HeaderText = "Field"
        DataGridView1.Columns(1).Width = 120
        DataGridView1.AllowUserToResizeColumns = True

        'Set up the Custom Attributes grid:
        DataGridView2.ColumnCount = 3
        DataGridView2.RowCount = 1
        DataGridView2.Columns(0).HeaderText = "Custom Attribute"
        DataGridView2.Columns(0).Width = 120
        DataGridView2.Columns(1).HeaderText = "Value Range"
        DataGridView2.Columns(1).Width = 120
        DataGridView2.Columns(2).HeaderText = "Value"
        DataGridView2.Columns(2).Width = 120
        DataGridView2.AllowUserToResizeColumns = True

        cmbXAxisType.Items.Add("Primary")
        cmbXAxisType.Items.Add("Secondary")

        For Each item In [Enum].GetNames(GetType(DataVisualization.Charting.ChartValueType))
            cmbXAxisValueType.Items.Add(item)
        Next

        For Each item In [Enum].GetNames(GetType(DataVisualization.Charting.ChartValueType))
            cmbYAxisValueType.Items.Add(item)
        Next

        cmbYAxisType.Items.Add("Primary")
        cmbYAxisType.Items.Add("Secondary")

        cmbXAxisTitleAlignment.Items.Add("Center")
        cmbXAxisTitleAlignment.Items.Add("Far")
        cmbXAxisTitleAlignment.Items.Add("Near")

        cmbX2AxisTitleAlignment.Items.Add("Center")
        cmbX2AxisTitleAlignment.Items.Add("Far")
        cmbX2AxisTitleAlignment.Items.Add("Near")

        cmbYAxisTitleAlignment.Items.Add("Center")
        cmbYAxisTitleAlignment.Items.Add("Far")
        cmbYAxisTitleAlignment.Items.Add("Near")

        cmbY2AxisTitleAlignment.Items.Add("Center")
        cmbY2AxisTitleAlignment.Items.Add("Far")
        cmbY2AxisTitleAlignment.Items.Add("Near")

        SetupPointChartSeriesTab()

        bgwSendMessage.WorkerReportsProgress = True
        bgwSendMessage.WorkerSupportsCancellation = True


        InitialiseForm() 'Initialise the form for a new project.

        'PointChart.DataLocation = Project.DataLocn 'This value is required for PointChart.LoadFile - used in RestoreProjectSettings.

        'END   Initialise the form: ---------------------------------------------------------------

        RestoreFormSettings() 'Restore the form settings
        Message.ShowXMessages = ShowXMessages
        Message.ShowSysMessages = ShowSysMessages
        RestoreProjectSettings() 'Restore the Project settings

        'RestoreProjectSettings also updates PointChart with the last Chart Settings.
        'Some more form settings can be restored:

        ''Set up Titles tab:  ------------------------------------------------------------------------------------------------------
        ''BottomCenter BottomLeft BottomRight MiddleCenter MiddleLeft MiddleRight TopCenter TopLeft TopRight
        'cmbAlignment.Items.Clear()

        ''Show the list of ContentAlignment enumerations in the cmbAlignment combobox:
        'Dim Alignment As String = [Enum].GetName(GetType(ContentAlignment), PointChart.ChartLabel.Alignment) 'MODIFY LATER TO SHOW SETTING FOR CURRENT CHART TYPE.
        'For Each item In System.Enum.GetValues(GetType(ContentAlignment))
        '    cmbAlignment.Items.Add(item)
        '    If item.ToString = Alignment Then
        '        cmbAlignment.SelectedIndex = cmbAlignment.Items.Count - 1
        '    End If
        'Next

        ''Set up the X Axis tab: ---------------------------------------------------------------------------------------------------
        'cmbXAxisTitleAlignment.Items.Clear()
        'Dim XAxisTitleAlignment As String = [Enum].GetName(GetType(StringAlignment), PointChart.XAxis.TitleAlignment) 'MODIFY LATER TO SHOW SETTING FOR CURRENT CHART TYPE.
        'For Each item In System.Enum.GetValues(GetType(StringAlignment))
        '    cmbXAxisTitleAlignment.Items.Add(item)
        '    If item.ToString = XAxisTitleAlignment Then
        '        cmbXAxisTitleAlignment.SelectedIndex = cmbXAxisTitleAlignment.Items.Count - 1
        '    End If
        'Next

        ''Set up the Y Axis tab: ---------------------------------------------------------------------------------------------------
        'cmbYAxisTitleAlignment.Items.Clear()
        'Dim YAxisTitleAlignment As String = [Enum].GetName(GetType(StringAlignment), PointChart.YAxis.TitleAlignment) 'MODIFY LATER TO SHOW SETTING FOR CURRENT CHART TYPE.
        'For Each item In System.Enum.GetValues(GetType(StringAlignment))
        '    cmbYAxisTitleAlignment.Items.Add(item)
        '    If item.ToString = YAxisTitleAlignment Then
        '        cmbYAxisTitleAlignment.SelectedIndex = cmbYAxisTitleAlignment.Items.Count - 1
        '    End If
        'Next

        'SetUpPointChartForm() 'Set up the Settings forms for a Point Chart.
        ''UpdatePointChartForm() 'Update the chart settings on the form with the settings stored in PointChart. THIS IS CALLED FROM SetUpPointChartForm()!


        ShowProjectInfo() 'Show the project information.

        If chkAutoDraw.Checked Then DrawPointChart()

        Message.AddText("------------------- Started OK -------------------------------------------------------------------------- " & vbCrLf & vbCrLf, "Heading")

        If StartupConnectionName = "" Then
            If Project.ConnectOnOpen Then
                ConnectToComNet() 'The Project is set to connect when it is opened.
            ElseIf ApplicationInfo.ConnectOnStartup Then
                ConnectToComNet() 'The Application is set to connect when it is started.
            Else
                'Don't connect to ComNet.
            End If
        Else
            'Connect to ComNet using the connection name StartupConnectionName.
            ConnectToComNet(StartupConnectionName)
        End If

    End Sub

    Private Sub InitialiseForm()
        'Initialise the form for a new project.
        OpenStartPage()

        ChartInfo.DataLocation = Project.DataLocn
        If Project.DataLocn.FileExists("LastChart.PointChart") Then
            Try
                ChartInfo.LoadFile("LastChart.PointChart", Chart1)
                UpdateInputDataTabSettings()
                UpdateTitlesTabSettings()
                UpdateSeriesTabSettings()
                UpdateAreasTabSettings()

            Catch ex As Exception
                Message.AddWarning("Error loading LastChart.PointChart: " & ex.Message & vbCrLf & vbCrLf)
            End Try
        End If

    End Sub

    Private Sub ShowProjectInfo()
        'Show the project information:

        txtParentProject.Text = Project.ParentProjectName
        txtProNetName.Text = Project.GetParameter("ProNetName")
        txtProjectName.Text = Project.Name
        txtProjectDescription.Text = Project.Description
        Select Case Project.Type
            Case ADVL_Utilities_Library_1.Project.Types.Directory
                txtProjectType.Text = "Directory"
            Case ADVL_Utilities_Library_1.Project.Types.Archive
                txtProjectType.Text = "Archive"
            Case ADVL_Utilities_Library_1.Project.Types.Hybrid
                txtProjectType.Text = "Hybrid"
            Case ADVL_Utilities_Library_1.Project.Types.None
                txtProjectType.Text = "None"
        End Select

        txtCreationDate.Text = Format(Project.Usage.FirstUsed, "d-MMM-yyyy H:mm:ss")
        txtLastUsed.Text = Format(Project.Usage.LastUsed, "d-MMM-yyyy H:mm:ss")

        txtProjectPath.Text = Project.Path

        Select Case Project.SettingsLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                txtSettingsLocationType.Text = "Directory"
            Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
                txtSettingsLocationType.Text = "Archive"
        End Select
        txtSettingsPath.Text = Project.SettingsLocn.Path

        Select Case Project.DataLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                txtDataLocationType.Text = "Directory"
            Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
                txtDataLocationType.Text = "Archive"
        End Select
        txtDataPath.Text = Project.DataLocn.Path

        Select Case Project.SystemLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                txtSystemLocationType.Text = "Directory"
            Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
                txtSystemLocationType.Text = "Archive"
        End Select
        txtSystemPath.Text = Project.SystemLocn.Path

        If Project.ConnectOnOpen Then
            chkConnect.Checked = True
        Else
            chkConnect.Checked = False
        End If

        txtTotalDuration.Text = Project.Usage.TotalDuration.Days.ToString.PadLeft(5, "0"c) & ":" &
                                Project.Usage.TotalDuration.Hours.ToString.PadLeft(2, "0"c) & ":" &
                                Project.Usage.TotalDuration.Minutes.ToString.PadLeft(2, "0"c) & ":" &
                                Project.Usage.TotalDuration.Seconds.ToString.PadLeft(2, "0"c)

        txtCurrentDuration.Text = Project.Usage.CurrentDuration.Days.ToString.PadLeft(5, "0"c) & ":" &
                                  Project.Usage.CurrentDuration.Hours.ToString.PadLeft(2, "0"c) & ":" &
                                  Project.Usage.CurrentDuration.Minutes.ToString.PadLeft(2, "0"c) & ":" &
                                  Project.Usage.CurrentDuration.Seconds.ToString.PadLeft(2, "0"c)

    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Exit the Application

        DisconnectFromComNet() 'Disconnect from the Communication Network (Message Service).

        SaveProjectSettings() 'Save project settings.

        ApplicationInfo.WriteFile() 'Update the Application Information file.

        Project.SaveLastProjectInfo() 'Save information about the last project used.

        Project.SaveParameters()
        ChartInfo.SaveFile("LastChart.PointChart", Chart1)

        'Project.SaveProjectInfoFile() 'Update the Project Information file. This is not required unless there is a change made to the project.

        Project.Usage.SaveUsageInfo() 'Save Project usage information.

        Project.UnlockProject() 'Unlock the project.

        ApplicationUsage.SaveUsageInfo() 'Save Application usage information.
        ApplicationInfo.UnlockApplication()

        Application.Exit()

    End Sub

    Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        'Save the form settings if the form state is normal. (A minimised form will have the incorrect size and location.)
        If WindowState = FormWindowState.Normal Then
            SaveFormSettings()
        End If
    End Sub


#End Region 'Form Display Methods -------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Open and Close Forms - Code used to open and close other forms." '===================================================================================================================

    Private Sub btnMessages_Click(sender As Object, e As EventArgs) Handles btnMessages.Click
        'Show the Messages form.
        Message.ApplicationName = ApplicationInfo.Name
        Message.SettingsLocn = Project.SettingsLocn
        Message.Show()
        Message.ShowXMessages = ShowXMessages
        Message.MessageForm.BringToFront()
    End Sub

    Private Sub btnWebPages_Click(sender As Object, e As EventArgs) Handles btnWebPages.Click
        'Open the Web Pages form.
        If IsNothing(WebPageList) Then
            WebPageList = New frmWebPageList
            WebPageList.Show()
        Else
            WebPageList.Show()
            WebPageList.BringToFront()
        End If
    End Sub

    Private Sub WebPageList_FormClosed(sender As Object, e As FormClosedEventArgs) Handles WebPageList.FormClosed
        WebPageList = Nothing
    End Sub

    Public Function OpenNewWebPage() As Integer
        'Open a new HTML Web View window, or reuse an existing one if avaiable.
        'The new forms index number in WebViewFormList is returned.

        NewWebPage = New frmWebPage
        If WebPageFormList.Count = 0 Then
            WebPageFormList.Add(NewWebPage)
            WebPageFormList(0).FormNo = 0
            WebPageFormList(0).Show
            Return 0 'The new HTML Display is at position 0 in WebViewFormList()
        Else
            Dim I As Integer
            Dim FormAdded As Boolean = False
            For I = 0 To WebPageFormList.Count - 1 'Check if there are closed forms in WebViewFormList. They can be re-used.
                If IsNothing(WebPageFormList(I)) Then
                    WebPageFormList(I) = NewWebPage
                    WebPageFormList(I).FormNo = I
                    WebPageFormList(I).Show
                    FormAdded = True
                    Return I 'The new Html Display is at position I in WebViewFormList()
                    Exit For
                End If
            Next
            If FormAdded = False Then 'Add a new form to WebViewFormList
                Dim FormNo As Integer
                WebPageFormList.Add(NewWebPage)
                FormNo = WebPageFormList.Count - 1
                WebPageFormList(FormNo).FormNo = FormNo
                WebPageFormList(FormNo).Show
                Return FormNo 'The new WebPage is at position FormNo in WebPageFormList()
            End If
        End If
    End Function

    Public Sub WebPageFormClosed()
        'This subroutine is called when the Web Page form has been closed.
        'The subroutine is usually called from the FormClosed event of the WebPage form.
        'The WebPage form may have multiple instances.
        'The ClosedFormNumber property should contain the number of the instance of the WebPage form.
        'This property should be updated by the WebPage form when it is being closed.
        'The ClosedFormNumber property value is used to determine which element in WebPageList should be set to Nothing.

        If WebPageFormList.Count < ClosedFormNo + 1 Then
            'ClosedFormNo is too large to exist in WebPageFormList
            Exit Sub
        End If

        If IsNothing(WebPageFormList(ClosedFormNo)) Then
            'The form is already set to nothing
        Else
            WebPageFormList(ClosedFormNo) = Nothing
        End If
    End Sub

    Public Function OpenNewHtmlDisplayPage() As Integer
        'Open a new HTML display window, or reuse an existing one if avaiable.
        'The new forms index number in HtmlDisplayFormList is returned.

        NewHtmlDisplay = New frmHtmlDisplay
        If HtmlDisplayFormList.Count = 0 Then
            HtmlDisplayFormList.Add(NewHtmlDisplay)
            HtmlDisplayFormList(0).FormNo = 0
            HtmlDisplayFormList(0).Show
            Return 0 'The new HTML Display is at position 0 in HtmlDisplayFormList()
        Else
            Dim I As Integer
            Dim FormAdded As Boolean = False
            For I = 0 To HtmlDisplayFormList.Count - 1 'Check if there are closed forms in HtmlDisplayFormList. They can be re-used.
                If IsNothing(HtmlDisplayFormList(I)) Then
                    HtmlDisplayFormList(I) = NewHtmlDisplay
                    HtmlDisplayFormList(I).FormNo = I
                    HtmlDisplayFormList(I).Show
                    FormAdded = True
                    Return I 'The new Html Display is at position I in HtmlDisplayFormList()
                    Exit For
                End If
            Next
            If FormAdded = False Then 'Add a new form to HtmlDisplayFormList
                Dim FormNo As Integer
                HtmlDisplayFormList.Add(NewHtmlDisplay)
                FormNo = HtmlDisplayFormList.Count - 1
                HtmlDisplayFormList(FormNo).FormNo = FormNo
                HtmlDisplayFormList(FormNo).Show
                Return FormNo 'The new HtmlDisplay is at position FormNo in HtmlDisplayFormList()
            End If
        End If
    End Function

    Public Sub HtmlDisplayFormClosed()
        'This subroutine is called when the Html Display form has been closed.
        'The subroutine is usually called from the FormClosed event of the HtmlDisplay form.
        'The HtmlDisplay form may have multiple instances.
        'The ClosedFormNumber property should contain the number of the instance of the HtmlDisplay form.
        'This property should be updated by the HtmlDisplay form when it is being closed.
        'The ClosedFormNumber property value is used to determine which element in HtmlDisplayList should be set to Nothing.

        If HtmlDisplayFormList.Count < ClosedFormNo + 1 Then
            'ClosedFormNo is too large to exist in HtmlDisplayFormList
            Exit Sub
        End If

        If IsNothing(HtmlDisplayFormList(ClosedFormNo)) Then
            'The form is already set to nothing
        Else
            HtmlDisplayFormList(ClosedFormNo) = Nothing
        End If
    End Sub



    Private Sub btnDesignQuery_Click(sender As Object, e As EventArgs) Handles btnDesignQuery.Click
        'Open the Design Query form:
        If IsNothing(DesignQuery) Then
            DesignQuery = New frmDesignQuery
            DesignQuery.Show()
            DesignQuery.DatabasePath = ChartInfo.InputDatabasePath
        Else
            DesignQuery.Show()
            DesignQuery.DatabasePath = ChartInfo.InputDatabasePath
        End If
    End Sub

    Private Sub DesignQuery_FormClosed(sender As Object, e As FormClosedEventArgs) Handles DesignQuery.FormClosed
        DesignQuery = Nothing
    End Sub

    Private Sub btnMarkerProps_Click(sender As Object, e As EventArgs) Handles btnMarkerProps.Click
        'Open the Marker Properties form.
        If IsNothing(MarkerProperties) Then
            MarkerProperties = New frmMarkerProperties
            MarkerProperties.Show()
            MarkerProperties.Chart = Chart1
            MarkerProperties.SelectSeries(txtSeriesName.Text)
        Else
            MarkerProperties.Show()
            MarkerProperties.SelectSeries(txtSeriesName.Text)
        End If
    End Sub



    Private Sub btnZoomChart_Click(sender As Object, e As EventArgs) Handles btnZoomChart.Click
        'Open the Zoom Chart form.
        If IsNothing(ZoomChart) Then
            ZoomChart = New frmZoomChart
            ZoomChart.Show()
            ZoomChart.Chart = Chart1
            ZoomChart.SelectAxis(txtAreaName.Text, "X Axis")
        Else
            ZoomChart.Show()
            ZoomChart.SelectAxis(txtAreaName.Text, "X Axis")
        End If
    End Sub

    Private Sub ZoomChart_FormClosed(sender As Object, e As FormClosedEventArgs) Handles ZoomChart.FormClosed
        ZoomChart = Nothing
    End Sub

    Private Sub Chart1_AxisViewChanged(sender As Object, e As ViewEventArgs) Handles Chart1.AxisViewChanged
        If IsNothing(ZoomChart) Then

        Else
            ZoomChart.UpdateSettings() 'Update the Zoom settings. These may have changed if the chart was scrolled.
        End If
    End Sub

    Private Sub btnMarkerProps2_Click(sender As Object, e As EventArgs) Handles btnMarkerProps2.Click
        'Open the Marker Properties form.
        If IsNothing(MarkerProperties) Then
            MarkerProperties = New frmMarkerProperties
            MarkerProperties.Show()
            MarkerProperties.Chart = Chart1
            MarkerProperties.SelectSeries(txtSeriesName.Text)
        Else
            MarkerProperties.Show()
            MarkerProperties.SelectSeries(txtSeriesName.Text)
        End If
    End Sub

    Private Sub MarkerProperties_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MarkerProperties.FormClosed
        MarkerProperties = Nothing
    End Sub


    Private Sub btnChartInfo_Click(sender As Object, e As EventArgs) Handles btnChartInfo.Click
        'Open the Chart Information Form.
        If IsNothing(ChartInfoForm) Then
            ChartInfoForm = New frmChartInfoForm
            ChartInfoForm.Show()

        Else
            ChartInfoForm.Show()
        End If
    End Sub

    Private Sub ChartInfoForm_FormClosed(sender As Object, e As FormClosedEventArgs) Handles ChartInfoForm.FormClosed
        ChartInfoForm = Nothing
    End Sub

#End Region 'Open and Close Forms -------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Form Methods - The main actions performed by this form." '===========================================================================================================================

    Private Sub btnProject_Click(sender As Object, e As EventArgs) Handles btnProject.Click
        Project.SelectProject()
    End Sub

    Private Sub btnParameters_Click(sender As Object, e As EventArgs) Handles btnParameters.Click
        Project.ShowParameters()
    End Sub

    Private Sub btnAppInfo_Click(sender As Object, e As EventArgs) Handles btnAppInfo.Click
        ApplicationInfo.ShowInfo()
    End Sub

    Private Sub btnAndorville_Click(sender As Object, e As EventArgs) Handles btnAndorville.Click
        ApplicationInfo.ShowInfo()
    End Sub

    Private Sub ApplicationInfo_UpdateExePath() Handles ApplicationInfo.UpdateExePath
        'Update the Executable Path.
        ApplicationInfo.ExecutablePath = Application.ExecutablePath
    End Sub

    Private Sub ApplicationInfo_RestoreDefaults() Handles ApplicationInfo.RestoreDefaults
        'Restore the default application settings.
        DefaultAppProperties()
    End Sub

    Public Sub UpdateWebPage(ByVal FileName As String)
        'Update the web page in WebPageFormList if the Web file name is FileName.

        Dim NPages As Integer = WebPageFormList.Count
        Dim I As Integer

        Try
            For I = 0 To NPages - 1
                If IsNothing(WebPageFormList(I)) Then
                    'Web page has been deleted!
                Else
                    If WebPageFormList(I).FileName = FileName Then
                        WebPageFormList(I).OpenDocument
                    End If
                End If
            Next
        Catch ex As Exception
            Message.AddWarning(ex.Message & vbCrLf)
        End Try
    End Sub


#Region " Start Page Code" '=========================================================================================================================================

    Public Sub OpenStartPage()
        'Open the StartPage.html file and display in the Workflow tab.

        If Project.DataFileExists("StartPage.html") Then
            WorkflowFileName = "StartPage.html"
            DisplayWorkflow()
        Else
            CreateStartPage()
            WorkflowFileName = "StartPage.html"
            DisplayWorkflow()
        End If
    End Sub

    Public Sub DisplayWorkflow()
        'Display the StartPage.html file in the Start Page tab.

        If Project.DataFileExists(WorkflowFileName) Then
            Dim rtbData As New IO.MemoryStream
            Project.ReadData(WorkflowFileName, rtbData)
            rtbData.Position = 0
            Dim sr As New IO.StreamReader(rtbData)
            WebBrowser1.DocumentText = sr.ReadToEnd()
        Else
            Message.AddWarning("Web page file not found: " & WorkflowFileName & vbCrLf)
        End If
    End Sub

    Private Sub CreateStartPage()
        'Create a new default StartPage.html file.

        Dim htmData As New IO.MemoryStream
        Dim sw As New IO.StreamWriter(htmData)
        sw.Write(AppInfoHtmlString("Application Information")) 'Create a web page providing information about the application.
        sw.Flush()
        Project.SaveData("StartPage.html", htmData)
    End Sub

    Public Function AppInfoHtmlString(ByVal DocumentTitle As String) As String
        'Create an Application Information Web Page.

        'This function should be edited to provide a brief description of the Application.

        Dim sb As New System.Text.StringBuilder

        sb.Append("<!DOCTYPE html>" & vbCrLf)
        sb.Append("<html>" & vbCrLf)
        sb.Append("<head>" & vbCrLf)
        sb.Append("<title>" & DocumentTitle & "</title>" & vbCrLf)
        sb.Append("<meta name=""description"" content=""Application information."">" & vbCrLf)
        sb.Append("</head>" & vbCrLf)

        sb.Append("<body style=""font-family:arial;"">" & vbCrLf & vbCrLf)

        sb.Append("<h2>" & "Andorville&trade; Point Chart" & "</h2>" & vbCrLf & vbCrLf) 'Add the page title.
        sb.Append("<hr>" & vbCrLf) 'Add a horizontal divider line.
        sb.Append("<p>The Point Chart application plots point data on a chart.</p>" & vbCrLf) 'Add an application description.
        sb.Append("<hr>" & vbCrLf & vbCrLf) 'Add a horizontal divider line.

        sb.Append(DefaultJavaScriptString)

        sb.Append("</body>" & vbCrLf)
        sb.Append("</html>" & vbCrLf)

        Return sb.ToString

    End Function

    Public Function DefaultJavaScriptString() As String
        'Generate the default JavaScript section of an Andorville(TM) Workflow Web Page.

        Dim sb As New System.Text.StringBuilder

        'Add JavaScript section:
        sb.Append("<script>" & vbCrLf & vbCrLf)

        'START: User defined JavaScript functions ==========================================================================
        'Add functions to implement the main actions performed by this web page.
        sb.Append("//START: User defined JavaScript functions ==========================================================================" & vbCrLf)
        sb.Append("//  Add functions to implement the main actions performed by this web page." & vbCrLf & vbCrLf)

        sb.Append("//END:   User defined JavaScript functions __________________________________________________________________________" & vbCrLf & vbCrLf & vbCrLf)
        'END:   User defined JavaScript functions --------------------------------------------------------------------------


        'START: User modified JavaScript functions ==========================================================================
        'Modify these function to save all required web page settings and process all expected XMessage instructions.
        sb.Append("//START: User modified JavaScript functions ==========================================================================" & vbCrLf)
        sb.Append("//  Modify these function to save all required web page settings and process all expected XMessage instructions." & vbCrLf & vbCrLf)

        'Add the Start Up code section.
        sb.Append("//Code to execute on Start Up:" & vbCrLf)
        sb.Append("function StartUpCode() {" & vbCrLf)
        sb.Append("  RestoreSettings() ;" & vbCrLf)
        'sb.Append("  GetCalcsDbPath() ;" & vbCrLf)
        sb.Append("}" & vbCrLf & vbCrLf)

        'Add the SaveSettings function - This is used to save web page settings between sessions.
        sb.Append("//Save the web page settings." & vbCrLf)
        sb.Append("function SaveSettings() {" & vbCrLf)
        sb.Append("  var xSettings = ""<Settings>"" + "" \n"" ; //String containing the web page settings in XML format." & vbCrLf)
        sb.Append("  //Add xml lines to save each setting." & vbCrLf & vbCrLf)
        sb.Append("  xSettings +=    ""</Settings>"" + ""\n"" ; //End of the Settings element." & vbCrLf)
        sb.Append(vbCrLf)
        sb.Append("  //Save the settings as an XML file in the project." & vbCrLf)
        sb.Append("  window.external.SaveHtmlSettings(xSettings) ;" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        'Process a single XMsg instruction (Information:Location pair)
        sb.Append("//Process an XMessage instruction:" & vbCrLf)
        sb.Append("function XMsgInstruction(Info, Locn) {" & vbCrLf)
        sb.Append("  switch(Locn) {" & vbCrLf)
        sb.Append("  //Insert case statements here." & vbCrLf)
        sb.Append("  case ""Status"" :" & vbCrLf)
        sb.Append("    if (Info = ""OK"") { " & vbCrLf)
        sb.Append("      //Instruction processing completed OK:" & vbCrLf)
        sb.Append("      } else {" & vbCrLf)
        sb.Append("      window.external.AddWarning(""Error: Unknown Status information: "" + "" Info: "" + Info + ""\r\n"") ;" & vbCrLf)
        sb.Append("     }" & vbCrLf)
        sb.Append("    break ;" & vbCrLf)
        sb.Append(vbCrLf)
        sb.Append("  default:" & vbCrLf)
        sb.Append("    window.external.AddWarning(""Unknown location: "" + Locn + ""\r\n"") ;" & vbCrLf)
        sb.Append("  }" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        sb.Append("//END:   User modified JavaScript functions __________________________________________________________________________" & vbCrLf & vbCrLf & vbCrLf)
        'END:   User modified JavaScript functions --------------------------------------------------------------------------

        'START: Required Document Library Web Page JavaScript functions ==========================================================================
        sb.Append("//START: Required Document Library Web Page JavaScript functions ==========================================================================" & vbCrLf & vbCrLf)

        'Add the AddText function - This sends a message to the message window using a named text type.
        sb.Append("//Add text to the Message window using a named txt type:" & vbCrLf)
        sb.Append("function AddText(Msg, TextType) {" & vbCrLf)
        sb.Append("  window.external.AddText(Msg, TextType) ;" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        'Add the AddMessage function - This sends a message to the message window using default black text.
        sb.Append("//Add a message to the Message window using the default black text:" & vbCrLf)
        sb.Append("function AddMessage(Msg) {" & vbCrLf)
        sb.Append("  window.external.AddMessage(Msg) ;" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        'Add the AddWarning function - This sends a red, bold warning message to the message window.
        sb.Append("//Add a warning message to the Message window using bold red text:" & vbCrLf)
        sb.Append("function AddWarning(Msg) {" & vbCrLf)
        sb.Append("  window.external.AddWarning(Msg) ;" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        'Add the RestoreSettings function - This is used to restore web page settings.
        sb.Append("//Restore the web page settings." & vbCrLf)
        sb.Append("function RestoreSettings() {" & vbCrLf)
        sb.Append("  window.external.RestoreHtmlSettings() " & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        'This line runs the RestoreSettings function when the web page is loaded.
        sb.Append("//Restore the web page settings when the page loads." & vbCrLf)
        'sb.Append("window.onload = RestoreSettings; " & vbCrLf)
        sb.Append("window.onload = StartUpCode ; " & vbCrLf)
        sb.Append(vbCrLf)

        'Restores a single setting on the web page.
        sb.Append("//Restore a web page setting." & vbCrLf)
        sb.Append("  function RestoreSetting(FormName, ItemName, ItemValue) {" & vbCrLf)
        sb.Append("  document.forms[FormName][ItemName].value = ItemValue ;" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        'Add the RestoreOption function - This is used to add an option to a Select list.
        sb.Append("//Restore a Select control Option." & vbCrLf)
        sb.Append("function RestoreOption(SelectId, OptionText) {" & vbCrLf)
        sb.Append("  var x = document.getElementById(SelectId) ;" & vbCrLf)
        sb.Append("  var option = document.createElement(""Option"") ;" & vbCrLf)
        sb.Append("  option.text = OptionText ;" & vbCrLf)
        sb.Append("  x.add(option) ;" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        sb.Append("//END:   Required Document Library Web Page JavaScript functions __________________________________________________________________________" & vbCrLf & vbCrLf)
        'END:   Required Document Library Web Page JavaScript functions --------------------------------------------------------------------------

        sb.Append("</script>" & vbCrLf & vbCrLf)

        Return sb.ToString

    End Function

    Public Function DefaultHtmlString(ByVal DocumentTitle As String) As String
        'Create a blank HTML Web Page.

        Dim sb As New System.Text.StringBuilder

        sb.Append("<!DOCTYPE html>" & vbCrLf)
        sb.Append("<html>" & vbCrLf)
        sb.Append("<!-- Andorville(TM) Workflow File -->" & vbCrLf)
        sb.Append("<!-- Application Name:    " & ApplicationInfo.Name & " -->" & vbCrLf)
        sb.Append("<!-- Application Version: " & My.Application.Info.Version.ToString & " -->" & vbCrLf)
        sb.Append("<!-- Creation Date:          " & Format(Now, "dd MMMM yyyy") & " -->" & vbCrLf)
        sb.Append("<head>" & vbCrLf)
        sb.Append("<title>" & DocumentTitle & "</title>" & vbCrLf)
        sb.Append("<meta name=""description"" content=""Workflow description."">" & vbCrLf)
        sb.Append("</head>" & vbCrLf)

        sb.Append("<body style=""font-family:arial;"">" & vbCrLf & vbCrLf)

        sb.Append("<h2>" & DocumentTitle & "</h2>" & vbCrLf & vbCrLf)

        sb.Append(DefaultJavaScriptString)

        sb.Append("</body>" & vbCrLf)
        sb.Append("</html>" & vbCrLf)

        Return sb.ToString

    End Function

    Public Function DefaultHtmlString_Old(ByVal DocumentTitle As String) As String
        'Create a blank HTML Web Page.

        Dim sb As New System.Text.StringBuilder

        sb.Append("<!DOCTYPE html>" & vbCrLf)
        sb.Append("<html>" & vbCrLf & "<head>" & vbCrLf & "<title>" & DocumentTitle & "</title>" & vbCrLf)
        sb.Append("</head>" & vbCrLf & "<body>" & vbCrLf & vbCrLf)
        sb.Append("<h1>" & DocumentTitle & "</h1>" & vbCrLf & vbCrLf)

        'Add JavaScript section:
        sb.Append("<script>" & vbCrLf & vbCrLf)

        'START: User defined JavaScript functions ==========================================================================
        'Add functions to implement the main actions performed by this web page.
        sb.Append("//START: User defined JavaScript functions ==========================================================================" & vbCrLf)
        sb.Append("//  Add functions to implement the main actions performed by this web page." & vbCrLf & vbCrLf)

        sb.Append("//END:   User defined JavaScript functions __________________________________________________________________________" & vbCrLf & vbCrLf & vbCrLf)
        'END:   User defined JavaScript functions --------------------------------------------------------------------------


        'START: User modified JavaScript functions ==========================================================================
        'Modify these function to save all required web page settings and process all expected XMessage instructions.
        sb.Append("//START: User modified JavaScript functions ==========================================================================" & vbCrLf)
        sb.Append("//  Modify these function to save all required web page settings and process all expected XMessage instructions." & vbCrLf & vbCrLf)

        'Add the SaveSettings function - This is used to save web page settings between sessions.
        sb.Append("//Save the web page settings." & vbCrLf)
        sb.Append("function SaveSettings() {" & vbCrLf)
        sb.Append("  var xSettings = ""<Settings>"" + "" \n"" ; //String containing the web page settings in XML format." & vbCrLf)
        sb.Append("  //Add xml lines to save each setting." & vbCrLf & vbCrLf)
        sb.Append("  xSettings +=    ""</Settings>"" + ""\n"" ; //End of the Settings element." & vbCrLf)
        sb.Append(vbCrLf)
        sb.Append("  //Save the settings as an XML file in the project." & vbCrLf)
        sb.Append("  window.external.SaveHtmlSettings(xSettings) ;" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        'Process a single XMsg instruction (Information:Location pair)
        sb.Append("//Process an XMessage instruction:" & vbCrLf)
        sb.Append("function XMsgInstruction(Info, Locn) {" & vbCrLf)
        sb.Append("  switch(Locn) {" & vbCrLf)
        sb.Append("  //Insert case statements here." & vbCrLf)
        sb.Append("  default:" & vbCrLf)
        sb.Append("    window.external.AddWarning(""Unknown location: "" + Locn + ""\r\n"") ;" & vbCrLf)
        sb.Append("  }" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        sb.Append("//END:   User modified JavaScript functions __________________________________________________________________________" & vbCrLf & vbCrLf & vbCrLf)
        'END:   User modified JavaScript functions --------------------------------------------------------------------------

        'START: Required Document Library Web Page JavaScript functions ==========================================================================
        sb.Append("//START: Required Document Library Web Page JavaScript functions ==========================================================================" & vbCrLf & vbCrLf)

        'Add the AddText function - This sends a message to the message window using a named text type.
        sb.Append("//Add text to the Message window using a named txt type:" & vbCrLf)
        sb.Append("function AddText(Msg, TextType) {" & vbCrLf)
        sb.Append("  window.external.AddText(Msg, TextType) ;" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        'Add the AddMessage function - This sends a message to the message window using default black text.
        sb.Append("//Add a message to the Message window using the default black text:" & vbCrLf)
        sb.Append("function AddMessage(Msg) {" & vbCrLf)
        sb.Append("  window.external.AddMessage(Msg) ;" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        'Add the AddWarning function - This sends a red, bold warning message to the message window.
        sb.Append("//Add a warning message to the Message window using bold red text:" & vbCrLf)
        sb.Append("function AddWarning(Msg) {" & vbCrLf)
        sb.Append("  window.external.AddWarning(Msg) ;" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        'Add the RestoreSettings function - This is used to restore web page settings.
        sb.Append("//Restore the web page settings." & vbCrLf)
        sb.Append("function RestoreSettings() {" & vbCrLf)
        sb.Append("  window.external.RestoreHtmlSettings() " & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        'This line runs the RestoreSettings function when the web page is loaded.
        sb.Append("//Restore the web page settings when the page loads." & vbCrLf)
        sb.Append("window.onload = RestoreSettings; " & vbCrLf)
        sb.Append(vbCrLf)

        'Restores a single setting on the web page.
        sb.Append("//Restore a web page setting." & vbCrLf)
        sb.Append("  function RestoreSetting(FormName, ItemName, ItemValue) {" & vbCrLf)
        sb.Append("  document.forms[FormName][ItemName].value = ItemValue ;" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        'Add the RestoreOption function - This is used to add an option to a Select list.
        sb.Append("//Restore a Select control Option." & vbCrLf)
        sb.Append("function RestoreOption(SelectId, OptionText) {" & vbCrLf)
        sb.Append("  var x = document.getElementById(SelectId) ;" & vbCrLf)
        sb.Append("  var option = document.createElement(""Option"") ;" & vbCrLf)
        sb.Append("  option.text = OptionText ;" & vbCrLf)
        sb.Append("  x.add(option) ;" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        sb.Append("//END:   Required Document Library Web Page JavaScript functions __________________________________________________________________________" & vbCrLf & vbCrLf)
        'END:   Required Document Library Web Page JavaScript functions --------------------------------------------------------------------------

        sb.Append("</script>" & vbCrLf & vbCrLf)

        sb.Append("</body>" & vbCrLf & "</html>" & vbCrLf)

        Return sb.ToString

    End Function

#End Region 'Start Page Code ------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Methods Called by JavaScript - A collection of methods that can be called by JavaScript in a web page shown in WebBrowser1" '==================================
    'These methods are used to display HTML pages in the Document tab.
    'The same methods can be found in the WebView form, which displays web pages on seprate forms.


    'Display Messages ==============================================================================================

    Public Sub AddMessage(ByVal Msg As String)
        'Add a normal text message to the Message window.
        Message.Add(Msg)
    End Sub

    Public Sub AddWarning(ByVal Msg As String)
        'Add a warning text message to the Message window.
        Message.AddWarning(Msg)
    End Sub

    Public Sub AddTextTypeMessage(ByVal Msg As String, ByVal TextType As String)
        'Add a message with the specified Text Type to the Message window.
        Message.AddText(Msg, TextType)
    End Sub

    Public Sub AddXmlMessage(ByVal XmlText As String)
        'Add an Xml message to the Message window.
        Message.AddXml(XmlText)
    End Sub

    'END Display Messages ------------------------------------------------------------------------------------------


    'Run an XSequence ==============================================================================================

    Public Sub RunClipboardXSeq()
        'Run the XSequence instructions in the clipboard.

        Dim XDocSeq As System.Xml.Linq.XDocument
        Try
            XDocSeq = XDocument.Parse(My.Computer.Clipboard.GetText)
        Catch ex As Exception
            Message.AddWarning("Error reading Clipboard data. " & ex.Message & vbCrLf)
            Exit Sub
        End Try

        If IsNothing(XDocSeq) Then
            Message.Add("No XSequence instructions were found in the clipboard.")
        Else
            Dim XmlSeq As New System.Xml.XmlDocument
            Try
                XmlSeq.LoadXml(XDocSeq.ToString) 'Convert XDocSeq to an XmlDocument to process with XSeq.
                'Run the sequence:
                XSeq.RunXSequence(XmlSeq, Status)
            Catch ex As Exception
                Message.AddWarning("Error restoring HTML settings. " & ex.Message & vbCrLf)
            End Try
        End If
    End Sub

    Public Sub RunXSequence(ByVal XSequence As String)
        'Run the XMSequence
        Dim XmlSeq As New System.Xml.XmlDocument
        XmlSeq.LoadXml(XSequence)
        XSeq.RunXSequence(XmlSeq, Status)
    End Sub

    Private Sub XSeq_ErrorMsg(ErrMsg As String) Handles XSeq.ErrorMsg
        Message.AddWarning(ErrMsg & vbCrLf)
    End Sub

    Private Sub XSeq_Instruction(Data As String, Locn As String) Handles XSeq.Instruction
        'Execute each instruction produced by running the XSeq file.

        Select Case Locn
            Case "Settings:Form:Name"
                FormName = Data

            Case "Settings:Form:Item:Name"
                ItemName = Data

            Case "Settings:Form:Item:Value"
                RestoreSetting(FormName, ItemName, Data)

            Case "Settings:Form:SelectId"
                SelectId = Data

            Case "Settings:Form:OptionText"
                RestoreOption(SelectId, Data)

            Case "Settings"

            Case "EndOfSequence"
                'Main.Message.Add("End of processing sequence" & Data & vbCrLf)

            Case Else
                Message.AddWarning("Unknown location: " & Locn & "  Data: " & Data & vbCrLf)

        End Select
    End Sub

    'END Run an XSequence ------------------------------------------------------------------------------------------


    'Run an XMessage ===============================================================================================

    Public Sub RunXMessage(ByVal XMsg As String)
        'Run the XMessage by sending it to InstrReceived.
        InstrReceived = XMsg
    End Sub

    Public Sub SendXMessage(ByVal ConnName As String, ByVal XMsg As String)
        'Send the XMessage to the application with the connection name ConnName.
        If IsNothing(client) Then
            Message.Add("No client connection available!" & vbCrLf)
        Else
            If client.State = ServiceModel.CommunicationState.Faulted Then
                Message.Add("client state is faulted. Message not sent!" & vbCrLf)
            Else
                If bgwSendMessage.IsBusy Then
                    Message.AddWarning("Send Message backgroundworker is busy." & vbCrLf)
                Else
                    Dim SendMessageParams As New clsSendMessageParams
                    SendMessageParams.ProjectNetworkName = ProNetName
                    SendMessageParams.ConnectionName = ConnName
                    SendMessageParams.Message = XMsg
                    bgwSendMessage.RunWorkerAsync(SendMessageParams)
                    If ShowXMessages Then
                        Message.XAddText("Message sent to " & "[" & ProNetName & "]." & ConnName & ":" & vbCrLf, "XmlSentNotice")
                        Message.XAddXml(XMsg)
                        Message.XAddText(vbCrLf, "Normal") 'Add extra line
                    End If
                End If
            End If
        End If
    End Sub

    Public Sub SendXMessageExt(ByVal ProNetName As String, ByVal ConnName As String, ByVal XMsg As String)
        'Send the XMsg to the application with the connection name ConnName and Project Network Name ProNetname.
        'This version can send the XMessage to a connection external to the current Project Network.
        If IsNothing(client) Then
            Message.Add("No client connection available!" & vbCrLf)
        Else
            If client.State = ServiceModel.CommunicationState.Faulted Then
                Message.Add("client state is faulted. Message not sent!" & vbCrLf)
            Else
                If bgwSendMessage.IsBusy Then
                    Message.AddWarning("Send Message backgroundworker is busy." & vbCrLf)
                Else
                    Dim SendMessageParams As New clsSendMessageParams
                    SendMessageParams.ProjectNetworkName = ProNetName
                    SendMessageParams.ConnectionName = ConnName
                    SendMessageParams.Message = XMsg
                    bgwSendMessage.RunWorkerAsync(SendMessageParams)
                    If ShowXMessages Then
                        Message.XAddText("Message sent to " & "[" & ProNetName & "]." & ConnName & ":" & vbCrLf, "XmlSentNotice")
                        Message.XAddXml(XMsg)
                        Message.XAddText(vbCrLf, "Normal") 'Add extra line
                    End If
                End If
            End If
        End If
    End Sub

    Public Sub SendXMessageWait(ByVal ConnName As String, ByVal XMsg As String)
        'Send the XMsg to the application with the connection name ConnName.
        'Wait for the connection to be made.
        If IsNothing(client) Then
            Message.Add("No client connection available!" & vbCrLf)
        Else
            Try
                'Application.DoEvents() 'TRY THE METHOD WITHOUT THE DOEVENTS
                If client.State = ServiceModel.CommunicationState.Faulted Then
                    Message.Add("client state is faulted. Message not sent!" & vbCrLf)
                Else
                    Dim StartTime As Date = Now
                    Dim Duration As TimeSpan
                    'Wait up to 16 seconds for the connection ConnName to be established
                    While client.ConnectionExists(ProNetName, ConnName) = False 'Wait until the required connection is made.
                        System.Threading.Thread.Sleep(1000) 'Pause for 1000ms
                        Duration = Now - StartTime
                        If Duration.Seconds > 16 Then Exit While
                    End While

                    If client.ConnectionExists(ProNetName, ConnName) = False Then
                        Message.AddWarning("Connection not available: " & ConnName & " in application network: " & ProNetName & vbCrLf)
                    Else
                        If bgwSendMessage.IsBusy Then
                            Message.AddWarning("Send Message backgroundworker is busy." & vbCrLf)
                        Else
                            Dim SendMessageParams As New clsSendMessageParams
                            SendMessageParams.ProjectNetworkName = ProNetName
                            SendMessageParams.ConnectionName = ConnName
                            SendMessageParams.Message = XMsg
                            bgwSendMessage.RunWorkerAsync(SendMessageParams)
                            If ShowXMessages Then
                                Message.XAddText("Message sent to " & "[" & ProNetName & "]." & ConnName & ":" & vbCrLf, "XmlSentNotice")
                                Message.XAddXml(XMsg)
                                Message.XAddText(vbCrLf, "Normal") 'Add extra line
                            End If
                        End If
                    End If
                End If
            Catch ex As Exception
                Message.AddWarning(ex.Message & vbCrLf)
            End Try
        End If
    End Sub

    Public Sub SendXMessageExtWait(ByVal ProNetName As String, ByVal ConnName As String, ByVal XMsg As String)
        'Send the XMsg to the application with the connection name ConnName and Project Network Name ProNetName.
        'Wait for the connection to be made.
        'This version can send the XMessage to a connection external to the current Project Network.
        If IsNothing(client) Then
            Message.Add("No client connection available!" & vbCrLf)
        Else
            If client.State = ServiceModel.CommunicationState.Faulted Then
                Message.Add("client state is faulted. Message not sent!" & vbCrLf)
            Else
                Dim StartTime As Date = Now
                Dim Duration As TimeSpan
                'Wait up to 16 seconds for the connection ConnName to be established
                While client.ConnectionExists(ProNetName, ConnName) = False
                    System.Threading.Thread.Sleep(1000) 'Pause for 1000ms
                    Duration = Now - StartTime
                    If Duration.Seconds > 16 Then Exit While
                End While

                If client.ConnectionExists(ProNetName, ConnName) = False Then
                    Message.AddWarning("Connection not available: " & ConnName & " in application network: " & ProNetName & vbCrLf)
                Else
                    If bgwSendMessage.IsBusy Then
                        Message.AddWarning("Send Message backgroundworker is busy." & vbCrLf)
                    Else
                        Dim SendMessageParams As New clsSendMessageParams
                        SendMessageParams.ProjectNetworkName = ProNetName
                        SendMessageParams.ConnectionName = ConnName
                        SendMessageParams.Message = XMsg
                        bgwSendMessage.RunWorkerAsync(SendMessageParams)
                        If ShowXMessages Then
                            Message.XAddText("Message sent to " & "[" & ProNetName & "]." & ConnName & ":" & vbCrLf, "XmlSentNotice")
                            Message.XAddXml(XMsg)
                            Message.XAddText(vbCrLf, "Normal") 'Add extra line
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Public Sub XMsgInstruction(ByVal Info As String, ByVal Locn As String)
        'Send the XMessage Instruction to the JavaScript function XMsgInstruction for processing.
        Me.WebBrowser1.Document.InvokeScript("XMsgInstruction", New String() {Info, Locn})
    End Sub

    'END Run an XMessage -------------------------------------------------------------------------------------------


    'Get Information ===============================================================================================

    Public Function GetFormNo() As String
        'Return FormNo.ToString
        Return "-1"
    End Function

    Public Function GetParentFormNo() As String
        'Return the Form Number of the Parent Form (that called this form).
        'Return ParentWebPageFormNo.ToString
        Return "-1" 'The Main Form does not have a Parent Web Page.
    End Function

    Public Function GetConnectionName() As String
        'Return the Connection Name of the Project.
        Return ConnectionName
    End Function

    Public Function GetProNetName() As String
        'Return the Project Network Name of the Project.
        Return ProNetName
    End Function

    Public Sub ParentProjectName(ByVal FormName As String, ByVal ItemName As String)
        'Return the Parent Project name:
        RestoreSetting(FormName, ItemName, Project.ParentProjectName)
    End Sub

    Public Sub ParentProjectPath(ByVal FormName As String, ByVal ItemName As String)
        'Return the Parent Project path:
        RestoreSetting(FormName, ItemName, Project.ParentProjectPath)
    End Sub

    Public Sub ParentProjectParameterValue(ByVal FormName As String, ByVal ItemName As String, ByVal ParameterName As String)
        'Return the specified Parent Project parameter value:
        RestoreSetting(FormName, ItemName, Project.ParentParameter(ParameterName).Value)
    End Sub

    Public Sub ProjectParameterValue(ByVal FormName As String, ByVal ItemName As String, ByVal ParameterName As String)
        'Return the specified Project parameter value:
        RestoreSetting(FormName, ItemName, Project.Parameter(ParameterName).Value)
    End Sub

    Public Sub ProjectNetworkName(ByVal FormName As String, ByVal ItemName As String)
        'Return the name of the Application Network:
        RestoreSetting(FormName, ItemName, Project.Parameter("ProNetName").Value)
    End Sub

    'END Get Information -------------------------------------------------------------------------------------------


    'Open a Web Page ===============================================================================================

    Public Sub OpenWebPage(ByVal FileName As String)
        'Open the web page with the specified File Name.

        If FileName = "" Then

        Else
            'First check if the HTML file is already open:
            Dim FileFound As Boolean = False
            If WebPageFormList.Count = 0 Then

            Else
                Dim I As Integer
                For I = 0 To WebPageFormList.Count - 1
                    If WebPageFormList(I) Is Nothing Then

                    Else
                        If WebPageFormList(I).FileName = FileName Then
                            FileFound = True
                            WebPageFormList(I).BringToFront
                        End If
                    End If
                Next
            End If

            If FileFound = False Then
                Dim FormNo As Integer = OpenNewWebPage()
                WebPageFormList(FormNo).FileName = FileName
                WebPageFormList(FormNo).OpenDocument
                WebPageFormList(FormNo).BringToFront
            End If
        End If
    End Sub

    'Public Sub OpenWebPage(ByVal FileName As String)
    '    'Open the web page with the specified File Name.

    '    If FileName = "" Then

    '    Else
    '        'First check if the HTML file is already open:
    '        Dim FileFound As Boolean = False
    '        If WebPageFormList.Count = 0 Then

    '        Else
    '            Dim I As Integer
    '            For I = 0 To WebPageFormList.Count - 1
    '                If WebPageFormList(I) Is Nothing Then

    '                Else
    '                    If WebPageFormList(I).FileName = FileName Then
    '                        FileFound = True
    '                        WebPageFormList(I).BringToFront
    '                    End If
    '                End If
    '            Next
    '        End If

    '        If FileFound = False Then
    '            Dim FormNo As Integer = OpenNewWebPage()
    '            WebPageFormList(FormNo).FileName = FileName
    '            WebPageFormList(FormNo).OpenDocument
    '            WebPageFormList(FormNo).BringToFront
    '        End If
    '    End If
    'End Sub

    'END Open a Web Page -------------------------------------------------------------------------------------------


    'Open and Close Projects =======================================================================================

    Public Sub OpenProjectAtRelativePath(ByVal RelativePath As String, ByVal ConnectionName As String)
        'Open the Project at the specified Relative Path using the specified Connection Name.

        Dim ProjectPath As String
        If RelativePath.StartsWith("\") Then
            ProjectPath = Project.Path & RelativePath
            client.StartProjectAtPath(ProjectPath, ConnectionName)
        Else
            ProjectPath = Project.Path & "\" & RelativePath
            client.StartProjectAtPath(ProjectPath, ConnectionName)
        End If
    End Sub

    Public Sub CheckOpenProjectAtRelativePath(ByVal RelativePath As String, ByVal ConnectionName As String)
        'Check if the project at the specified Relative Path is open.
        'Open it if it is not already open.
        'Open the Project at the specified Relative Path using the specified Connection Name.

        Dim ProjectPath As String
        If RelativePath.StartsWith("\") Then
            ProjectPath = Project.Path & RelativePath
            If client.ProjectOpen(ProjectPath) Then
                'Project is already open.
            Else
                client.StartProjectAtPath(ProjectPath, ConnectionName)
            End If
        Else
            ProjectPath = Project.Path & "\" & RelativePath
            If client.ProjectOpen(ProjectPath) Then
                'Project is already open.
            Else
                client.StartProjectAtPath(ProjectPath, ConnectionName)
            End If
        End If
    End Sub

    Public Sub OpenProjectAtProNetPath(ByVal RelativePath As String, ByVal ConnectionName As String)
        'Open the Project at the specified Path (relative to the Project Network Path) using the specified Connection Name.

        Dim ProjectPath As String
        If RelativePath.StartsWith("\") Then
            If Project.ParameterExists("ProNetPath") Then
                ProjectPath = Project.GetParameter("ProNetPath") & RelativePath
                client.StartProjectAtPath(ProjectPath, ConnectionName)
            Else
                Message.AddWarning("The Project Network Path is not known." & vbCrLf)
            End If
        Else
            If Project.ParameterExists("ProNetPath") Then
                ProjectPath = Project.GetParameter("ProNetPath") & "\" & RelativePath
                client.StartProjectAtPath(ProjectPath, ConnectionName)
            Else
                Message.AddWarning("The Project Network Path is not known." & vbCrLf)
            End If
        End If
    End Sub

    Public Sub CheckOpenProjectAtProNetPath(ByVal RelativePath As String, ByVal ConnectionName As String)
        'Check if the project at the specified Path (relative to the Project Network Path) is open.
        'Open it if it is not already open.
        'Open the Project at the specified Path using the specified Connection Name.

        Dim ProjectPath As String
        If RelativePath.StartsWith("\") Then
            If Project.ParameterExists("ProNetPath") Then
                ProjectPath = Project.GetParameter("ProNetPath") & RelativePath
                'client.StartProjectAtPath(ProjectPath, ConnectionName)
                If client.ProjectOpen(ProjectPath) Then
                    'Project is already open.
                Else
                    client.StartProjectAtPath(ProjectPath, ConnectionName)
                End If
            Else
                Message.AddWarning("The Project Network Path is not known." & vbCrLf)
            End If
        Else
            If Project.ParameterExists("ProNetPath") Then
                ProjectPath = Project.GetParameter("ProNetPath") & "\" & RelativePath
                'client.StartProjectAtPath(ProjectPath, ConnectionName)
                If client.ProjectOpen(ProjectPath) Then
                    'Project is already open.
                Else
                    client.StartProjectAtPath(ProjectPath, ConnectionName)
                End If
            Else
                Message.AddWarning("The Project Network Path is not known." & vbCrLf)
            End If
        End If
    End Sub

    Public Sub CloseProjectAtConnection(ByVal ProNetName As String, ByVal ConnectionName As String)
        'Close the Project at the specified connection.

        If IsNothing(client) Then
            Message.Add("No client connection available!" & vbCrLf)
        Else
            If client.State = ServiceModel.CommunicationState.Faulted Then
                Message.Add("client state is faulted. Message not sent!" & vbCrLf)
            Else
                'Create the XML instructions to close the application at the connection.
                Dim decl As New XDeclaration("1.0", "utf-8", "yes")
                Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
                Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class

                'NOTE: No reply expected. No need to provide the following client information(?)
                'Dim clientConnName As New XElement("ClientConnectionName", Me.ConnectionName)
                'xmessage.Add(clientConnName)

                Dim command As New XElement("Command", "Close")
                xmessage.Add(command)
                doc.Add(xmessage)

                'Show the message sent to AppNet:
                Message.XAddText("Message sent to: [" & ProNetName & "]." & ConnectionName & ":" & vbCrLf, "XmlSentNotice")
                Message.XAddXml(doc.ToString)
                Message.XAddText(vbCrLf, "Normal") 'Add extra line

                client.SendMessage(ProNetName, ConnectionName, doc.ToString)
            End If
        End If
    End Sub

    'END Open and Close Projects -----------------------------------------------------------------------------------


    'System Methods ================================================================================================

    Public Sub SaveHtmlSettings(ByVal xSettings As String, ByVal FileName As String)
        'Save the Html settings for a web page.

        'Convert the XSettings to XML format:
        Dim XmlHeader As String = "<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>"
        Dim XDocSettings As New System.Xml.Linq.XDocument

        Try
            XDocSettings = System.Xml.Linq.XDocument.Parse(XmlHeader & vbCrLf & xSettings)
        Catch ex As Exception
            Message.AddWarning("Error saving HTML settings file. " & ex.Message & vbCrLf)
        End Try

        Project.SaveXmlData(FileName, XDocSettings)
    End Sub


    Public Sub RestoreHtmlSettings()
        'Restore the Html settings for a web page.

        Dim SettingsFileName As String = WorkflowFileName & "Settings"
        Dim XDocSettings As New System.Xml.Linq.XDocument
        Project.ReadXmlData(SettingsFileName, XDocSettings)

        If XDocSettings Is Nothing Then
            'Message.Add("No HTML Settings file : " & SettingsFileName & vbCrLf)
        Else
            Dim XSettings As New System.Xml.XmlDocument
            Try
                XSettings.LoadXml(XDocSettings.ToString)
                'Run the Settings file:
                XSeq.RunXSequence(XSettings, Status)
            Catch ex As Exception
                Message.AddWarning("Error restoring HTML settings. " & ex.Message & vbCrLf)
            End Try
        End If
    End Sub

    Public Sub RestoreSetting(ByVal FormName As String, ByVal ItemName As String, ByVal ItemValue As String)
        'Restore the setting value with the specified Form Name and Item Name.
        Me.WebBrowser1.Document.InvokeScript("RestoreSetting", New String() {FormName, ItemName, ItemValue})
    End Sub

    Public Sub RestoreOption(ByVal SelectId As String, ByVal OptionText As String)
        'Restore the Option text in the Select control with the Id SelectId.
        Me.WebBrowser1.Document.InvokeScript("RestoreOption", New String() {SelectId, OptionText})
    End Sub

    Private Sub SaveWebPageSettings()
        'Call the SaveSettings JavaScript function:
        Try
            Me.WebBrowser1.Document.InvokeScript("SaveSettings")
        Catch ex As Exception
            Message.AddWarning("Web page settings not saved: " & ex.Message & vbCrLf)
        End Try

    End Sub

    'END System Methods --------------------------------------------------------------------------------------------


    'Legacy Code (These methods should no longer be used) ==========================================================

    Public Sub JSMethodTest1()
        'Test method that is called from JavaScript.
        Message.Add("JSMethodTest1 called OK." & vbCrLf)
    End Sub

    Public Sub JSMethodTest2(ByVal Var1 As String, ByVal Var2 As String)
        'Test method that is called from JavaScript.
        Message.Add("Var1 = " & Var1 & " Var2 = " & Var2 & vbCrLf)
    End Sub

    Public Sub JSDisplayXml(ByRef XDoc As XDocument)
        Message.Add(XDoc.ToString & vbCrLf & vbCrLf)
    End Sub

    Public Sub ShowMessage(ByVal Msg As String)
        Message.Add(Msg)
    End Sub

    Public Sub AddText(ByVal Msg As String, ByVal TextType As String)
        Message.AddText(Msg, TextType)
    End Sub


    'END Legacy Code -----------------------------------------------------------------------------------------------


#End Region 'Methods Called by JavaScript -------------------------------------------------------------------------------------------------------------------------------


#Region " Project Events Code"

    Private Sub Project_Message(Msg As String) Handles Project.Message
        'Display the Project message:
        Message.Add(Msg & vbCrLf)
    End Sub

    Private Sub Project_ErrorMessage(Msg As String) Handles Project.ErrorMessage
        'Display the Project error message:
        Message.AddWarning(Msg & vbCrLf)
    End Sub

    Private Sub Project_Closing() Handles Project.Closing
        'The current project is closing.
        SaveFormSettings() 'Save the form settings - they are saved in the Project before is closes.
        SaveProjectSettings() 'Update this subroutine if project settings need to be saved.
        Project.Usage.SaveUsageInfo() 'Save the current project usage information.
        Project.UnlockProject() 'Unlock the current project before it Is closed.
        If ConnectedToComNet Then DisconnectFromComNet()
    End Sub

    Private Sub Project_Selected() Handles Project.Selected
        'A new project has been selected.

        RestoreFormSettings()
        Project.ReadProjectInfoFile()

        Project.ReadParameters()
        Project.ReadParentParameters()
        If Project.ParentParameterExists("ProNetName") Then
            Project.AddParameter("ProNetName", Project.ParentParameter("ProNetName").Value, Project.ParentParameter("ProNetName").Description) 'AddParameter will update the parameter if it already exists.
            ProNetName = Project.Parameter("ProNetName").Value
        Else
            ProNetName = Project.GetParameter("ProNetName")
        End If
        If Project.ParentParameterExists("ProNetPath") Then 'Get the parent parameter value - it may have been updated.
            Project.AddParameter("ProNetPath", Project.ParentParameter("ProNetPath").Value, Project.ParentParameter("ProNetPath").Description) 'AddParameter will update the parameter if it already exists.
            ProNetPath = Project.Parameter("ProNetPath").Value
        Else
            ProNetPath = Project.GetParameter("ProNetPath") 'If the parameter does not exist, the value is set to ""
        End If
        Project.SaveParameters() 'These should be saved now - child projects look for parent parameters in the parameter file.

        Project.LockProject() 'Lock the project while it is open in this application.

        Project.Usage.StartTime = Now

        ApplicationInfo.SettingsLocn = Project.SettingsLocn
        Message.SettingsLocn = Project.SettingsLocn
        Message.Show() 'Added 18May19

        'Restore the new project settings:
        RestoreProjectSettings() 'Update this subroutine if project settings need to be restored.

        ShowProjectInfo()

        ''Show the project information:
        'txtProjectName.Text = Project.Name
        'txtProjectDescription.Text = Project.Description
        'Select Case Project.Type
        '    Case ADVL_Utilities_Library_1.Project.Types.Directory
        '        txtProjectType.Text = "Directory"
        '    Case ADVL_Utilities_Library_1.Project.Types.Archive
        '        txtProjectType.Text = "Archive"
        '    Case ADVL_Utilities_Library_1.Project.Types.Hybrid
        '        txtProjectType.Text = "Hybrid"
        '    Case ADVL_Utilities_Library_1.Project.Types.None
        '        txtProjectType.Text = "None"
        'End Select

        'txtCreationDate.Text = Format(Project.CreationDate, "d-MMM-yyyy H:mm:ss")
        'txtLastUsed.Text = Format(Project.Usage.LastUsed, "d-MMM-yyyy H:mm:ss")
        'Select Case Project.SettingsLocn.Type
        '    Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
        '        txtSettingsLocationType.Text = "Directory"
        '    Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
        '        txtSettingsLocationType.Text = "Archive"
        'End Select
        'txtSettingsPath.Text = Project.SettingsLocn.Path
        'Select Case Project.DataLocn.Type
        '    Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
        '        txtDataLocationType.Text = "Directory"
        '    Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
        '        txtDataLocationType.Text = "Archive"
        'End Select
        'txtDataPath.Text = Project.DataLocn.Path

        If Project.ConnectOnOpen Then
            ConnectToComNet() 'The Project is set to connect when it is opened.
        ElseIf ApplicationInfo.ConnectOnStartup Then
            ConnectToComNet() 'The Application is set to connect when it is started.
        Else
            'Don't connect to ComNet.
        End If

    End Sub

    Private Sub chkConnect_LostFocus(sender As Object, e As EventArgs) Handles chkConnect.LostFocus
        If chkConnect.Checked Then
            Project.ConnectOnOpen = True
        Else
            Project.ConnectOnOpen = False
        End If
        Project.SaveProjectInfoFile()
    End Sub

#End Region 'Project Events Code

#Region " Online/Offline Code" '=========================================================================================================================================

    Private Sub btnOnline_Click(sender As Object, e As EventArgs) Handles btnOnline.Click
        'Connect to or disconnect from the Message System (ComNet).
        If ConnectedToComNet = False Then
            ConnectToComNet()
        Else
            DisconnectFromComNet()
        End If
    End Sub

    Private Sub ConnectToComNet()
        'Connect to the Message Service. (ComNet)

        If IsNothing(client) Then
            client = New ServiceReference1.MsgServiceClient(New System.ServiceModel.InstanceContext(New MsgServiceCallback))
        End If

        If ComNetRunning() Then
            'The Application.Lock file has been found at AdvlNetworkAppPath
            'The Message Service is Running.
        Else 'The Message Service is NOT running'
            'Start the Andorville™ Network:
            If AdvlNetworkAppPath = "" Then
                Message.AddWarning("Andorville™ Network application path is unknown." & vbCrLf)
            Else
                If System.IO.File.Exists(AdvlNetworkExePath) Then 'OK to start the Message Service application:
                    Shell(Chr(34) & AdvlNetworkExePath & Chr(34), AppWinStyle.NormalFocus) 'Start Message Service application with no argument
                Else
                    'Incorrect Message Service Executable path.
                    Message.AddWarning("Andorville™ Network exe file not found. Service not started." & vbCrLf)
                End If
            End If
        End If

        'Try to fix a faulted client state:
        If client.State = ServiceModel.CommunicationState.Faulted Then
            client = Nothing
            client = New ServiceReference1.MsgServiceClient(New System.ServiceModel.InstanceContext(New MsgServiceCallback))
        End If

        If client.State = ServiceModel.CommunicationState.Faulted Then
            Message.AddWarning("Client state is faulted. Connection not made!" & vbCrLf)
        Else
            Try
                client.Endpoint.Binding.SendTimeout = New System.TimeSpan(0, 0, 16) 'Temporarily set the send timeaout to 16 seconds (8 seconds is too short for a slow computer!)

                ConnectionName = ApplicationInfo.Name 'This name will be modified if it is already used in an existing connection.
                ConnectionName = client.Connect(ProNetName, ApplicationInfo.Name, ConnectionName, Project.Name, Project.Description, Project.Type, Project.Path, False, False)

                If ConnectionName <> "" Then
                    Message.Add("Connected to the Andorville™ Network with Connection Name: [" & ProNetName & "]." & ConnectionName & vbCrLf)
                    client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeout to 1 hour
                    btnOnline.Text = "Online"
                    btnOnline.ForeColor = Color.ForestGreen
                    ConnectedToComNet = True
                    SendApplicationInfo()
                    SendProjectInfo()
                    client.GetAdvlNetworkAppInfoAsync() 'Update the Exe Path in case it has changed. This path may be needed in the future to start the ComNet (Message Service).

                    bgwComCheck.WorkerReportsProgress = True
                    bgwComCheck.WorkerSupportsCancellation = True
                    If bgwComCheck.IsBusy Then
                        'The ComCheck thread is already running.
                    Else
                        bgwComCheck.RunWorkerAsync() 'Start the ComCheck thread.
                    End If
                Else
                    Message.Add("Connection to the Andorville™ Network failed!" & vbCrLf)
                    client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeout to 1 hour
                End If
            Catch ex As System.TimeoutException
                Message.Add("Timeout error. Check if the Andorville™ Network (Message Service) is running." & vbCrLf)
            Catch ex As Exception
                Message.Add("Error message: " & ex.Message & vbCrLf)
                client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeout to 1 hour
            End Try
        End If
    End Sub

    Private Sub ConnectToComNet(ByVal ConnName As String)
        'Connect to the Message Service (ComNet) with the connection name ConnName.

        If ConnectedToComNet = False Then
            If IsNothing(client) Then
                client = New ServiceReference1.MsgServiceClient(New System.ServiceModel.InstanceContext(New MsgServiceCallback))
            End If

            'Try to fix a faulted client state:
            If client.State = ServiceModel.CommunicationState.Faulted Then
                client = Nothing
                client = New ServiceReference1.MsgServiceClient(New System.ServiceModel.InstanceContext(New MsgServiceCallback))
            End If

            If client.State = ServiceModel.CommunicationState.Faulted Then
                Message.AddWarning("client state is faulted. Connection not made!" & vbCrLf)
            Else
                Try
                    client.Endpoint.Binding.SendTimeout = New System.TimeSpan(0, 0, 16) 'Temporarily set the send timeout to 16 seconds
                    ConnectionName = ConnName 'This name will be modified if it is already used in an existing connection.
                    ConnectionName = client.Connect(ProNetName, ApplicationInfo.Name, ConnectionName, Project.Name, Project.Description, Project.Type, Project.Path, False, False)

                    If ConnectionName <> "" Then
                        Message.Add("Connected to the Andorville™ Network with Connection Name: [" & ProNetName & "]." & ConnectionName & vbCrLf)
                        client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeout to 1 hour
                        btnOnline.Text = "Online"
                        btnOnline.ForeColor = Color.ForestGreen
                        ConnectedToComNet = True
                        SendApplicationInfo()
                        SendProjectInfo()
                        client.GetAdvlNetworkAppInfoAsync() 'Update the Exe Path in case it has changed. This path may be needed in the future to start the ComNet (Message Service).

                        bgwComCheck.WorkerReportsProgress = True
                        bgwComCheck.WorkerSupportsCancellation = True
                        If bgwComCheck.IsBusy Then
                            'The ComCheck thread is already running.
                        Else
                            bgwComCheck.RunWorkerAsync() 'Start the ComCheck thread.
                        End If
                    Else
                        Message.Add("Connection to the Andorville™ Network failed!" & vbCrLf)
                        client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeout to 1 hour
                    End If
                Catch ex As System.TimeoutException
                    Message.Add("Timeout error. Check if the Andorville™ Network (Message Service) is running." & vbCrLf)
                Catch ex As Exception
                    Message.Add("Error message: " & ex.Message & vbCrLf)
                    client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeout to 1 hour
                End Try
            End If
        Else
            Message.AddWarning("Already connected to the Andorville™ Network (Message Service)." & vbCrLf)
        End If

    End Sub

    Private Sub DisconnectFromComNet()
        'Disconnect from the Communication Network (Message Service).

        If ConnectedToComNet = True Then
            If IsNothing(client) Then
                'Message.Add("Already disconnected from the Communication Network." & vbCrLf)
                Message.Add("Already disconnected from the Andorville™ Network (Message Service)." & vbCrLf)
                btnOnline.Text = "Offline"
                btnOnline.ForeColor = Color.Red
                ConnectedToComNet = False
                ConnectionName = ""
            Else
                If client.State = ServiceModel.CommunicationState.Faulted Then
                    Message.Add("client state is faulted." & vbCrLf)
                    ConnectionName = ""
                Else
                    Try
                        'client.Disconnect(AppNetName, ConnectionName)
                        client.Disconnect(ProNetName, ConnectionName)
                        btnOnline.Text = "Offline"
                        btnOnline.ForeColor = Color.Red
                        ConnectedToComNet = False
                        ConnectionName = ""
                        'Message.Add("Disconnected from the Communication Network." & vbCrLf)
                        Message.Add("Disconnected from the Andorville™ Network (Message Service)." & vbCrLf)

                        If bgwComCheck.IsBusy Then
                            bgwComCheck.CancelAsync()
                        End If
                    Catch ex As Exception
                        'Message.AddWarning("Error disconnecting from Communication Network: " & ex.Message & vbCrLf)
                        Message.AddWarning("Error disconnecting from Andorville™ Network (Message Service): " & ex.Message & vbCrLf)
                    End Try
                End If
            End If
        End If
    End Sub

    Private Sub SendApplicationInfo()
        'Send the application information to the Administrator connections.

        If IsNothing(client) Then
            Message.Add("No client connection available!" & vbCrLf)
        Else
            If client.State = ServiceModel.CommunicationState.Faulted Then
                Message.Add("Client state is faulted. Message not sent!" & vbCrLf)
            Else
                'Create the XML instructions to send application information.
                Dim decl As New XDeclaration("1.0", "utf-8", "yes")
                Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
                Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class
                Dim applicationInfo As New XElement("ApplicationInfo")
                Dim name As New XElement("Name", Me.ApplicationInfo.Name)
                applicationInfo.Add(name)

                Dim text As New XElement("Text", "Point Chart")
                applicationInfo.Add(text)

                Dim exePath As New XElement("ExecutablePath", Me.ApplicationInfo.ExecutablePath)
                applicationInfo.Add(exePath)

                Dim directory As New XElement("Directory", Me.ApplicationInfo.ApplicationDir)
                applicationInfo.Add(directory)
                Dim description As New XElement("Description", Me.ApplicationInfo.Description)
                applicationInfo.Add(description)
                xmessage.Add(applicationInfo)
                doc.Add(xmessage)

                'Show the message sent to ComNet:
                Message.XAddText("Message sent to " & "Message Service" & ":" & vbCrLf, "XmlSentNotice")
                Message.XAddXml(doc.ToString)
                Message.XAddText(vbCrLf, "Normal") 'Add extra line

                client.SendMessage("", "MessageService", doc.ToString)
            End If
        End If

    End Sub

    Private Sub SendProjectInfo()
        'Send the project information to the Network application.

        If ConnectedToComNet = False Then
            Message.AddWarning("The application is not connected to the Message Service." & vbCrLf)
        Else 'Connected to the Message Service (ComNet).
            If IsNothing(client) Then
                Message.Add("No client connection available!" & vbCrLf)
            Else
                If client.State = ServiceModel.CommunicationState.Faulted Then
                    Message.Add("Client state is faulted. Message not sent!" & vbCrLf)
                Else
                    'Construct the XMessage to send to AppNet:
                    Dim decl As New XDeclaration("1.0", "utf-8", "yes")
                    Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
                    Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class
                    Dim projectInfo As New XElement("ProjectInfo")

                    Dim Path As New XElement("Path", Project.Path)
                    projectInfo.Add(Path)
                    xmessage.Add(projectInfo)
                    doc.Add(xmessage)

                    'Show the message sent to the Message Service:
                    Message.XAddText("Message sent to " & "Message Service" & ":" & vbCrLf, "XmlSentNotice")
                    Message.XAddXml(doc.ToString)
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                    client.SendMessage("", "MessageService", doc.ToString)
                End If
            End If
        End If
    End Sub

    Public Sub SendProjectInfo(ByVal ProjectPath As String)
        'Send the project information to the Network application.
        'This version of SendProjectInfo uses the ProjectPath argument.

        If ConnectedToComNet = False Then
            Message.AddWarning("The application is not connected to the Message Service." & vbCrLf)
        Else 'Connected to the Message Service (ComNet).
            If IsNothing(client) Then
                Message.Add("No client connection available!" & vbCrLf)
            Else
                If client.State = ServiceModel.CommunicationState.Faulted Then
                    Message.Add("Client state is faulted. Message not sent!" & vbCrLf)
                Else
                    'Construct the XMessage to send to AppNet:
                    Dim decl As New XDeclaration("1.0", "utf-8", "yes")
                    Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
                    Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class
                    Dim projectInfo As New XElement("ProjectInfo")

                    Dim Path As New XElement("Path", ProjectPath)
                    projectInfo.Add(Path)
                    xmessage.Add(projectInfo)
                    doc.Add(xmessage)

                    'Show the message sent to the Message Service:
                    Message.XAddText("Message sent to " & "Message Service" & ":" & vbCrLf, "XmlSentNotice")
                    Message.XAddXml(doc.ToString)
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                    client.SendMessage("", "MessageService", doc.ToString)
                End If
            End If
        End If
    End Sub

    Private Function ComNetRunning() As Boolean
        'Return True if ComNet (Message Service) is running.
        ''If System.IO.File.Exists(MsgServiceAppPath & "\Application.Lock") Then
        'If System.IO.File.Exists(AdvlNetworkAppPath & "\Application.Lock") Then
        '    Return True
        'Else
        '    Return False
        'End If

        'If MsgServiceAppPath = "" Then
        If AdvlNetworkAppPath = "" Then
            'Message.Add("Message Service application path is not known." & vbCrLf)
            Message.Add("Andorville™ Network application path is not known." & vbCrLf)
            'Message.Add("Run the Message Service before connecting to update the path." & vbCrLf)
            Message.Add("Run the Andorville™ Network before connecting to update the path." & vbCrLf)
            Return False
        Else
            'If System.IO.File.Exists(MsgServiceAppPath & "\Application.Lock") Then
            If System.IO.File.Exists(AdvlNetworkAppPath & "\Application.Lock") Then
                'Message.Add("AppLock found - ComNet is running." & vbCrLf)
                Return True
            Else
                'Message.Add("AppLock not found - ComNet is running." & vbCrLf)
                Return False
            End If
        End If

    End Function

#End Region 'Online/Offline code ----------------------------------------------------------------------------------------------------------------------------------------

    Private Sub TabPage2_Enter(sender As Object, e As EventArgs) Handles TabPage2.Enter
        'Update the current duration:

        txtCurrentDuration.Text = Project.Usage.CurrentDuration.Days.ToString.PadLeft(5, "0"c) & ":" &
                                   Project.Usage.CurrentDuration.Hours.ToString.PadLeft(2, "0"c) & ":" &
                                   Project.Usage.CurrentDuration.Minutes.ToString.PadLeft(2, "0"c) & ":" &
                                   Project.Usage.CurrentDuration.Seconds.ToString.PadLeft(2, "0"c)

        Timer2.Interval = 5000 '5 seconds
        Timer2.Enabled = True
        Timer2.Start()

    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        'Update the current duration:

        txtCurrentDuration.Text = Project.Usage.CurrentDuration.Days.ToString.PadLeft(5, "0"c) & ":" &
                           Project.Usage.CurrentDuration.Hours.ToString.PadLeft(2, "0"c) & ":" &
                           Project.Usage.CurrentDuration.Minutes.ToString.PadLeft(2, "0"c) & ":" &
                           Project.Usage.CurrentDuration.Seconds.ToString.PadLeft(2, "0"c)
    End Sub

    Private Sub TabPage2_Leave(sender As Object, e As EventArgs) Handles TabPage2.Leave
        Timer2.Enabled = False
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        'Add the current project to the Message Service list.

        If Project.ParentProjectName <> "" Then
            Message.AddWarning("This project has a parent: " & Project.ParentProjectName & vbCrLf)
            Message.AddWarning("Child projects can not be added to the list." & vbCrLf)
            Exit Sub
        End If

        If ConnectedToComNet = False Then
            Message.AddWarning("The application is not connected to the Message Service." & vbCrLf)
        Else 'Connected to the Message Service (ComNet).
            If IsNothing(client) Then
                Message.Add("No client connection available!" & vbCrLf)
            Else
                If client.State = ServiceModel.CommunicationState.Faulted Then
                    Message.Add("Client state is faulted. Message not sent!" & vbCrLf)
                Else
                    'Construct the XMessage to send to AppNet:
                    Dim decl As New XDeclaration("1.0", "utf-8", "yes")
                    Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
                    Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class
                    Dim projectInfo As New XElement("ProjectInfo")

                    Dim Path As New XElement("Path", Project.Path)
                    projectInfo.Add(Path)
                    xmessage.Add(projectInfo)
                    doc.Add(xmessage)

                    'Show the message sent to AppNet:
                    Message.XAddText("Message sent to " & "Message Service" & ":" & vbCrLf, "XmlSentNotice")
                    Message.XAddXml(doc.ToString)
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                    client.SendMessage("", "MessageService", doc.ToString)
                End If
            End If
        End If
    End Sub

    Private Sub btnOpenProject_Click(sender As Object, e As EventArgs) Handles btnOpenProject.Click

        If Project.Type = ADVL_Utilities_Library_1.Project.Types.Archive Then

        Else
            Process.Start(Project.Path)
        End If

    End Sub

    Private Sub btnOpenSettings_Click(sender As Object, e As EventArgs) Handles btnOpenSettings.Click
        If Project.SettingsLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory Then
            Process.Start(Project.SettingsLocn.Path)
        End If
    End Sub

    Private Sub btnOpenData_Click(sender As Object, e As EventArgs) Handles btnOpenData.Click
        If Project.DataLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory Then
            Process.Start(Project.DataLocn.Path)
        End If
    End Sub

    Private Sub btnOpenSystem_Click(sender As Object, e As EventArgs) Handles btnOpenSystem.Click
        If Project.SystemLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory Then
            Process.Start(Project.SystemLocn.Path)
        End If
    End Sub

    Private Sub btnOpenAppDir_Click(sender As Object, e As EventArgs) Handles btnOpenAppDir.Click
        Process.Start(ApplicationInfo.ApplicationDir)
    End Sub


#Region " Process XMessages" '===========================================================================================================================================

    Private Sub XMsg_Instruction(Data As String, Locn As String) Handles XMsg.Instruction
        'Process an XMessage instruction.
        'An XMessage is a simplified XSequence. It is used to exchange information between Andorville™ applications.
        '
        'An XSequence file is an AL-H7™ Information Sequence stored in an XML format.
        'AL-H7™ is the name of a programming system that uses sequences of data and location value pairs to store information or processing steps.
        'Any program, mathematical expression or data set can be expressed as an Information Sequence.

        'Add code here to process the XMessage instructions.
        'See other Andorville™ applications for examples.

        If IsDBNull(Data) Then
            Data = ""
        End If

        'Intercept instructions with the prefix "WebPage_"
        If Locn.StartsWith("WebPage_") Then 'Send the Data, Location data to the correct Web Page:
            'Message.Add("Web Page Location: " & Locn & vbCrLf)
            If Locn.Contains(":") Then
                Dim EndOfWebPageNoString As Integer = Locn.IndexOf(":")
                If Locn.Contains("-") Then
                    Dim HyphenLocn As Integer = Locn.IndexOf("-")
                    If HyphenLocn < EndOfWebPageNoString Then 'Web Page Location contains a sub-location in the web page - WebPage_1-SubLocn:Locn - SubLocn:Locn will be sent to Web page 1
                        EndOfWebPageNoString = HyphenLocn
                    End If
                End If
                Dim PageNoLen As Integer = EndOfWebPageNoString - 8
                Dim WebPageNoString As String = Locn.Substring(8, PageNoLen)
                Dim WebPageNo As Integer = CInt(WebPageNoString)
                Dim WebPageData As String = Data
                Dim WebPageLocn As String = Locn.Substring(EndOfWebPageNoString + 1)

                'Message.Add("WebPageData = " & WebPageData & "  WebPageLocn = " & WebPageLocn & vbCrLf)

                WebPageFormList(WebPageNo).XMsgInstruction(WebPageData, WebPageLocn)
            Else
                Message.AddWarning("XMessage instruction location is not complete: " & Locn & vbCrLf)
            End If
        Else

            Select Case Locn

            'Case "ClientAppNetName"
            '    ClientAppNetName = Data 'The name of the Client Application Network requesting service. ADDED 2Feb19.
                Case "ClientProNetName"
                    ClientProNetName = Data 'The name of the Client Application Network requesting service. AD

                Case "ClientName"
                    ClientAppName = Data 'The name of the Client application requesting service.

                Case "ClientConnectionName"
                    ClientConnName = Data 'The name of the client connection requesting service.

                Case "ClientLocn" 'The Location within the Client requesting service.
                    Dim statusOK As New XElement("Status", "OK") 'Add Status OK element when the Client Location is changed
                    xlocns(xlocns.Count - 1).Add(statusOK)

                    xmessage.Add(xlocns(xlocns.Count - 1)) 'Add the instructions for the last location to the reply xmessage
                    xlocns.Add(New XElement(Data)) 'Start the new location instructions

                'Case "OnCompletion" 'Specify the last instruction to be returned on completion of the XMessage processing.
                '    CompletionInstruction = Data

                    'UPDATE:
                Case "OnCompletion"
                    OnCompletionInstruction = Data

                Case "Main"
                 'Blank message - do nothing.

                'Case "Main:OnCompletion"
                '    Select Case "Stop"
                '        'Stop on completion of the instruction sequence.
                '    End Select

                Case "Main:EndInstruction"
                    Select Case Data
                        Case "Stop"
                            'Stop at the end of the instruction sequence.

                            'Add other cases here:
                    End Select

                Case "Main:Status"
                    Select Case Data
                        Case "OK"
                            'Main instructions completed OK
                    End Select

                Case "Command"
                    Select Case Data
                        Case "ConnectToComNet" 'Startup Command
                            If ConnectedToComNet = False Then
                                ConnectToComNet()
                            End If

                        Case "AppComCheck"
                            'Add the Appplication Communication info to the reply message:
                            Dim clientProNetName As New XElement("ClientProNetName", ProNetName) 'The Project Network Name
                            xlocns(xlocns.Count - 1).Add(clientProNetName)
                            Dim clientName As New XElement("ClientName", "ADVL_Point_Chart_1") 'The name of this application.
                            xlocns(xlocns.Count - 1).Add(clientName)
                            Dim clientConnectionName As New XElement("ClientConnectionName", ConnectionName)
                            xlocns(xlocns.Count - 1).Add(clientConnectionName)
                            '<Status>OK</Status> will be automatically appended to the XMessage before it is sent.

                    End Select

                Case "GetChartSettings"
                    'This is the new code used to send the current chart settings to the client.
                    '  The chart settings will be returned to the client in an XDataMsg file (instead of an XMsg file).
                    '  Data contains the Client Location to include in the XDataMsg file.
                    GetPointChartSettings(Data)


            'Startup Command Arguments ================================================
            'Case "AppNetName"
            '    'This is currently not used.
            '    'The AppNetName is determined elsewhere.
                Case "ProNetName"
                'This is currently not used.
                'The ProNetName is determined elsewhere.

                Case "ProjectName"
                    If Project.OpenProject(Data) = True Then
                        ProjectSelected = True 'Project has been opened OK.
                    Else
                        ProjectSelected = False 'Project could not be opened.
                    End If

                Case "ProjectID"
                    Message.AddWarning("Add code to handle ProjectID parameter at StartUp!" & vbCrLf)
                'Note the ComNet will usually select a project using ProjectPath.

                Case "ProjectPath"
                    If Project.OpenProjectPath(Data) = True Then
                        ProjectSelected = True 'Project has been opened OK.
                        'THE PROJECT IS LOCKED IN THE Form.Load EVENT:

                        ApplicationInfo.SettingsLocn = Project.SettingsLocn
                        Message.SettingsLocn = Project.SettingsLocn 'Set up the Message object
                        Message.Show() 'Added 18May19

                        txtTotalDuration.Text = Project.Usage.TotalDuration.Days.ToString.PadLeft(5, "0"c) & ":" &
                                      Project.Usage.TotalDuration.Hours.ToString.PadLeft(2, "0"c) & ":" &
                                      Project.Usage.TotalDuration.Minutes.ToString.PadLeft(2, "0"c) & ":" &
                                      Project.Usage.TotalDuration.Seconds.ToString.PadLeft(2, "0"c)

                        txtCurrentDuration.Text = Project.Usage.CurrentDuration.Days.ToString.PadLeft(5, "0"c) & ":" &
                                       Project.Usage.CurrentDuration.Hours.ToString.PadLeft(2, "0"c) & ":" &
                                       Project.Usage.CurrentDuration.Minutes.ToString.PadLeft(2, "0"c) & ":" &
                                       Project.Usage.CurrentDuration.Seconds.ToString.PadLeft(2, "0"c)

                    Else
                        ProjectSelected = False 'Project could not be opened.
                        Message.AddWarning("Project could not be opened at path: " & Data & vbCrLf)
                    End If

                Case "ConnectionName"
                    StartupConnectionName = Data
            '--------------------------------------------------------------------------

            'Application Information  =================================================
            'returned by client.GetAdvlNetworkAppInfoAsync()
            'Case "MessageServiceAppInfo:Name"
            '    'The name of the Message Service Application. (Not used.)
                Case "AdvlNetworkAppInfo:Name"
                'The name of the Andorville™ Network Application. (Not used.)

            'Case "MessageServiceAppInfo:ExePath"
            '    'The executable file path of the Message Service Application.
            '    MsgServiceExePath = Data
                Case "AdvlNetworkAppInfo:ExePath"
                    'The executable file path of the Andorville™ Network Application.
                    AdvlNetworkExePath = Data

            'Case "MessageServiceAppInfo:Path"
            '    'The path of the Message Service Application (ComNet). (This is where an Application.Lock file will be found while ComNet is running.)
            '    MsgServiceAppPath = Data
                Case "AdvlNetworkAppInfo:Path"
                    'The path of the Andorville™ Network Application (ComNet). (This is where an Application.Lock file will be found while ComNet is running.)
                    AdvlNetworkAppPath = Data
           '---------------------------------------------------------------------------

           'Message Window Instructions  ==============================================
                Case "MessageWindow:Left"
                    If IsNothing(Message.MessageForm) Then
                        Message.ApplicationName = ApplicationInfo.Name
                        Message.SettingsLocn = Project.SettingsLocn
                        Message.Show()
                    End If
                    Message.MessageForm.Left = Data
                Case "MessageWindow:Top"
                    If IsNothing(Message.MessageForm) Then
                        Message.ApplicationName = ApplicationInfo.Name
                        Message.SettingsLocn = Project.SettingsLocn
                        Message.Show()
                    End If
                    Message.MessageForm.Top = Data
                Case "MessageWindow:Width"
                    If IsNothing(Message.MessageForm) Then
                        Message.ApplicationName = ApplicationInfo.Name
                        Message.SettingsLocn = Project.SettingsLocn
                        Message.Show()
                    End If
                    Message.MessageForm.Width = Data
                Case "MessageWindow:Height"
                    If IsNothing(Message.MessageForm) Then
                        Message.ApplicationName = ApplicationInfo.Name
                        Message.SettingsLocn = Project.SettingsLocn
                        Message.Show()
                    End If
                    Message.MessageForm.Height = Data
                Case "MessageWindow:Command"
                    Select Case Data
                        Case "BringToFront"
                            If IsNothing(Message.MessageForm) Then
                                Message.ApplicationName = ApplicationInfo.Name
                                Message.SettingsLocn = Project.SettingsLocn
                                Message.Show()
                            End If
                            'Message.MessageForm.BringToFront()
                            Message.MessageForm.Activate()
                            Message.MessageForm.TopMost = True
                            Message.MessageForm.TopMost = False
                        Case "SaveSettings"
                            Message.MessageForm.SaveFormSettings()
                    End Select

            '---------------------------------------------------------------------------

            'Command to bring the Application window to the front:
                Case "ApplicationWindow:Command"
                    Select Case Data
                        Case "BringToFront"
                            Me.Activate()
                            Me.TopMost = True
                            Me.TopMost = False
                    End Select


                Case "EndOfSequence"
                    'End of Information Vector Sequence reached.
                    'Add Status OK element at the end of the sequence:
                    Dim statusOK As New XElement("Status", "OK")
                    xlocns(xlocns.Count - 1).Add(statusOK)

                    Select Case EndInstruction
                        Case "Stop"
                            'No instructions.

                            'Add any other Cases here:

                        Case Else
                            Message.AddWarning("Unknown End Instruction: " & EndInstruction & vbCrLf)
                    End Select
                    EndInstruction = "Stop"

                    ''Add the final OnCompletion instruction:
                    'Dim onCompletion As New XElement("OnCompletion", CompletionInstruction) '
                    'xlocns(xlocns.Count - 1).Add(onCompletion)
                    'CompletionInstruction = "Stop" 'Reset the Completion Instruction

                    ''Final Version:
                    ''Add the final EndInstruction:
                    'Dim xEndInstruction As New XElement("EndInstruction", OnCompletionInstruction)
                    'xlocns(xlocns.Count - 1).Add(xEndInstruction)
                    'OnCompletionInstruction = "Stop" 'Reset the OnCompletion Instruction

                    'Add the final EndInstruction:
                    If OnCompletionInstruction = "Stop" Then
                        'Final EndInstruction is not required.
                    Else
                        Dim xEndInstruction As New XElement("EndInstruction", OnCompletionInstruction)
                        xlocns(xlocns.Count - 1).Add(xEndInstruction)
                        OnCompletionInstruction = "Stop" 'Reset the OnCompletion Instruction
                    End If

                Case Else
                    Message.AddWarning("Unknown location: " & Locn & vbCrLf)
                    Message.AddWarning("            info: " & Data & vbCrLf)
            End Select
        End If
    End Sub

    Private Sub GetPointChartSettings(ByVal ClientLocn As String)
        'Return the current Point Chart settings data in an XMsgBlk (XMessageBlock) file.

        If ConnectedToComNet = False Then
            Message.AddWarning("The application is not connected to the Message Service." & vbCrLf)
        Else
            If IsNothing(client) Then
                Message.Add("No client connection available!" & vbCrLf)
            Else
                If client.State = ServiceModel.CommunicationState.Faulted Then
                    Message.Add("Client state is faulted. Message not sent!" & vbCrLf)
                Else
                    Dim ChartSettings As System.Xml.Linq.XDocument = ChartInfo.ToXDoc(Chart1)

                    Dim Decl As New XDeclaration("1.0", "utf-8", "yes")
                    Dim xMsgBlkDoc As New XDocument(Decl, Nothing)
                    Dim xMsgBlk As New XElement("XMsgBlk")

                    Dim clientLocation As New XElement("ClientLocn", ClientLocn)
                    xMsgBlk.Add(clientLocation)

                    Dim xInfo As New XElement("XInfo")
                    xInfo.Add(ChartSettings.<ChartSettings>)
                    xMsgBlk.Add(xInfo)
                    xMsgBlkDoc.Add(xMsgBlk)

                    'Show the message sent to the Message Service:
                    Message.XAddText("Message sent to [" & ClientProNetName & "]." & ClientConnName & ":" & vbCrLf, "XmlSentNotice")
                    Message.XAddXml(xMsgBlkDoc.ToString)
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line

                    'Because Instructiosn are still being processed, use Alternative SendMessage background worker
                    SendMessageParamsAlt.ProjectNetworkName = ClientProNetName
                    SendMessageParamsAlt.ConnectionName = ClientConnName
                    SendMessageParamsAlt.Message = xMsgBlkDoc.ToString
                    If bgwSendMessageAlt.IsBusy Then
                        Message.AddWarning("Send Message backgroundworker is busy." & vbCrLf)
                    Else
                        bgwSendMessageAlt.RunWorkerAsync(SendMessageParamsAlt)
                    End If
                End If
            End If
        End If
    End Sub

    'Private Sub SendMessage()
    '    'Code used to send a message after a timer delay.
    '    'The message destination is stored in MessageDest
    '    'The message text is stored in MessageText
    '    Timer1.Interval = 100 '100ms delay
    '    Timer1.Enabled = True 'Start the timer.
    'End Sub

    'Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

    '    If IsNothing(client) Then
    '        Message.AddWarning("No client connection available!" & vbCrLf)
    '    Else
    '        If client.State = ServiceModel.CommunicationState.Faulted Then
    '            Message.AddWarning("client state is faulted. Message not sent!" & vbCrLf)
    '        Else
    '            Try
    '                'Message.Add("Sending a message. Number of characters: " & MessageText.Length & vbCrLf)
    '                'client.SendMessage(ClientAppNetName, ClientConnName, MessageText)
    '                client.SendMessage(ClientProNetName, ClientConnName, MessageText)
    '                MessageText = "" 'Clear the message after it has been sent.
    '                ClientAppName = "" 'Clear the Client Application Name after the message has been sent.
    '                'ClientAppLocn = "" 'Clear the Client Application Location after the message has been sent.
    '                'ClientAppNetName = "" 'Clear the Client Application Network Name after the message has been sent. ADDED 2Feb19.
    '                ClientProNetName = "" 'Clear the Client Application Network Name after the message has been sent.
    '                ClientConnName = "" 'Clear the Client Application Name after the message has been sent.
    '                xlocns.Clear()
    '            Catch ex As Exception
    '                Message.AddWarning("Error sending message: " & ex.Message & vbCrLf)
    '            End Try
    '        End If
    '    End If

    '    'Stop timer:
    '    Timer1.Enabled = False
    'End Sub

    'Private Sub Timer3_Tick(sender As Object, e As EventArgs)
    '    'Keet the connection awake with each tick:

    '    If ConnectedToComNet = True Then
    '        Try
    '            If client.IsAlive() Then
    '                Message.Add(Format(Now, "HH:mm:ss") & " Connection OK." & vbCrLf)
    '                Timer3.Interval = TimeSpan.FromMinutes(55).TotalMilliseconds '55 minute interval
    '            Else
    '                Message.Add(Format(Now, "HH:mm:ss") & " Connection Fault." & vbCrLf)
    '                Timer3.Interval = TimeSpan.FromMinutes(55).TotalMilliseconds '55 minute interval
    '            End If
    '        Catch ex As Exception
    '            Message.AddWarning(ex.Message & vbCrLf)
    '            'Set interval to five minutes - try again in five minutes:
    '            Timer3.Interval = TimeSpan.FromMinutes(5).TotalMilliseconds '5 minute interval
    '        End Try
    '    Else
    '        Message.Add(Format(Now, "HH:mm:ss") & " Not connected." & vbCrLf)
    '    End If
    'End Sub

    Private Sub XMsgLocal_Instruction(Data As String, Locn As String) Handles XMsgLocal.Instruction
        'Process an XMessage instruction locally.

        If IsDBNull(Data) Then
            Data = ""
        End If

        'Intercept instructions with the prefix "WebPage_"
        If Locn.StartsWith("WebPage_") Then 'Send the Data, Location data to the correct Web Page:
            'Message.Add("Web Page Location: " & Locn & vbCrLf)
            If Locn.Contains(":") Then
                Dim EndOfWebPageNoString As Integer = Locn.IndexOf(":")
                If Locn.Contains("-") Then
                    Dim HyphenLocn As Integer = Locn.IndexOf("-")
                    If HyphenLocn < EndOfWebPageNoString Then 'Web Page Location contains a sub-location in the web page - WebPage_1-SubLocn:Locn - SubLocn:Locn will be sent to Web page 1
                        EndOfWebPageNoString = HyphenLocn
                    End If
                End If
                Dim PageNoLen As Integer = EndOfWebPageNoString - 8
                Dim WebPageNoString As String = Locn.Substring(8, PageNoLen)
                Dim WebPageNo As Integer = CInt(WebPageNoString)
                Dim WebPageData As String = Data
                Dim WebPageLocn As String = Locn.Substring(EndOfWebPageNoString + 1)

                'Message.Add("WebPageData = " & WebPageData & "  WebPageLocn = " & WebPageLocn & vbCrLf)

                WebPageFormList(WebPageNo).XMsgInstruction(WebPageData, WebPageLocn)
            Else
                Message.AddWarning("XMessage instruction location is not complete: " & Locn & vbCrLf)
            End If
        Else

            Select Case Locn
                Case "ClientName"
                    ClientAppName = Data 'The name of the Client requesting service.

                         'UPDATE:
                Case "OnCompletion"
                    OnCompletionInstruction = Data

                Case "Main"
                 'Blank message - do nothing.

                'Case "Main:OnCompletion"
                '    Select Case "Stop"
                '        'Stop on completion of the instruction sequence.
                '    End Select

                Case "Main:EndInstruction"
                    Select Case Data
                        Case "Stop"
                            'Stop at the end of the instruction sequence.

                            'Add other cases here:
                    End Select

                Case "Main:Status"
                    Select Case Data
                        Case "OK"
                            'Main instructions completed OK
                    End Select

                Case "EndOfSequence"
                    'End of Information Sequence reached.

                Case Else
                    Message.AddWarning("Local XMessage: " & Locn & vbCrLf)
                    Message.AddWarning("Unknown location: " & Locn & vbCrLf)
                    Message.AddWarning("            data: " & Data & vbCrLf)
            End Select
        End If
    End Sub



#End Region 'Process XMessages ------------------------------------------------------------------------------------------------------------------------------------------


    Private Sub ToolStripMenuItem1_EditWorkflowTabPage_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1_EditWorkflowTabPage.Click
        'Edit the Workflow Web Page:

        If WorkflowFileName = "" Then
            Message.AddWarning("No page to edit." & vbCrLf)
        Else
            Dim FormNo As Integer = OpenNewHtmlDisplayPage()
            HtmlDisplayFormList(FormNo).FileName = WorkflowFileName
            HtmlDisplayFormList(FormNo).OpenDocument
        End If

    End Sub

    Private Sub ToolStripMenuItem1_ShowStartPageInWorkflowTab_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1_ShowStartPageInWorkflowTab.Click
        'Show the Start Page in the Workflow Tab:
        OpenStartPage()
    End Sub

    Private Sub Label7_Click(sender As Object, e As EventArgs) Handles Label7.Click

    End Sub

    Private Sub Main_Move(sender As Object, e As EventArgs) Handles Me.Move
        txtLeft.Text = Me.Left
        txtTop.Text = Me.Top
    End Sub

#Region " Chart Input Data Tab" '========================================================================================================================================

    Private Sub UpdateInputDataTabSettings()
        'Update the Input Data tab settings from the Chart control.

        txtDatabasePath.Text = ChartInfo.InputDatabasePath
        Select Case ChartInfo.InputDataType
            Case "Database"
                rbDatabase.Checked = True
            Case "Dataset"
                rbDataset.Checked = True
            Case Else
                rbDatabase.Checked = True
        End Select
        FillLstTables()
        txtDataDescription.Text = ChartInfo.InputDataDescr
        txtInputQuery.Text = ChartInfo.InputQuery

        ChartInfo.ApplyQuery()
        'GetChartOptionsFromDataset() 'This updates the chart options using Dataset fields.
        UpdateSeriesTabSettings()
        GetChartOptionsFromDataset() 'This updates the chart options using Dataset fields.

    End Sub



    Private Sub btnDatabase_Click(sender As Object, e As EventArgs) Handles btnDatabase.Click
        'Select a database file:

        OpenFileDialog1.Filter = "Access Database |*.accdb"
        OpenFileDialog1.FileName = ""

        If InputDatabaseDirectory <> "" Then
            OpenFileDialog1.InitialDirectory = InputDatabaseDirectory
        End If

        OpenFileDialog1.ShowDialog()

        'PointChart.InputDatabasePath = OpenFileDialog1.FileName
        ChartInfo.InputDatabasePath = OpenFileDialog1.FileName
        'txtDatabasePath.Text = PointChart.InputDatabasePath
        txtDatabasePath.Text = ChartInfo.InputDatabasePath
        'FillLstTables()

        InputDatabaseDirectory = System.IO.Path.GetDirectoryName(OpenFileDialog1.FileName)

        Message.Add("InputDatabaseDirectory = " & InputDatabaseDirectory & vbCrLf)

        FillLstTables()
    End Sub

    Private Sub FillLstTables()
        'Fill the lstSelectTable listbox with the availalble tables in the selected database.

        lstTables.Items.Clear()

        'If PointChart.InputDatabasePath = "" Then Exit Sub
        If ChartInfo.InputDatabasePath = "" Then Exit Sub

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim dt As DataTable

        lstTables.Items.Clear()

        'Specify the connection string:
        'Access 2003
        'connectionString = "provider=Microsoft.Jet.OLEDB.4.0;" + _
        '"data source = " + txtDatabase.Text

        'Access 2007:
        connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        "data source = " + txtDatabasePath.Text

        'Connect to the Access database:
        conn = New System.Data.OleDb.OleDbConnection(connectionString)

        conn.Open()

        Dim restrictions As String() = New String() {Nothing, Nothing, Nothing, "TABLE"} 'This restriction removes system tables
        dt = conn.GetSchema("Tables", restrictions)

        'Fill lstTables
        Dim dr As DataRow
        Dim I As Integer 'Loop index
        Dim MaxI As Integer

        MaxI = dt.Rows.Count
        For I = 0 To MaxI - 1
            dr = dt.Rows(0)
            lstTables.Items.Add(dt.Rows(I).Item(2).ToString)
        Next I

        conn.Close()

    End Sub

    Private Sub lstTables_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstTables.SelectedIndexChanged
        FillLstFields()
    End Sub

    Private Sub FillLstFields()
        'Fill the lstFields listbox with the availalble fields in the selected table.

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim commandString As String 'Declare a command string - contains the query to be passed to the database.
        Dim ds As DataSet 'Declate a Dataset.
        Dim dt As DataTable

        If lstTables.SelectedIndex = -1 Then 'No item is selected
            lstFields.Items.Clear()

        Else 'A table has been selected. List its fields:
            lstFields.Items.Clear()

            'Specify the connection string (Access 2007):
            connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
            "data source = " + txtDatabasePath.Text

            'Connect to the Access database:
            conn = New System.Data.OleDb.OleDbConnection(connectionString)
            conn.Open()

            'Specify the commandString to query the database:
            commandString = "SELECT TOP 500 * FROM " + lstTables.SelectedItem.ToString
            Dim dataAdapter As New System.Data.OleDb.OleDbDataAdapter(commandString, conn)
            ds = New DataSet
            dataAdapter.Fill(ds, "SelTable") 'ds was defined earlier as a DataSet
            dt = ds.Tables("SelTable")

            Dim NFields As Integer
            NFields = dt.Columns.Count
            Dim I As Integer
            For I = 0 To NFields - 1
                lstFields.Items.Add(dt.Columns(I).ColumnName.ToString)
            Next

            conn.Close()

        End If
    End Sub

    Private Sub btnViewData_Click(sender As Object, e As EventArgs) Handles btnViewData.Click
        'Open the View Database Data form:
        If IsNothing(ViewDatabaseData) Then
            ViewDatabaseData = New frmViewDatabaseData
            ViewDatabaseData.Show()
            ViewDatabaseData.Update()
        Else
            ViewDatabaseData.Show()
            ViewDatabaseData.Update()
        End If
    End Sub

    Private Sub ViewDatabaseData_FormClosed(sender As Object, e As FormClosedEventArgs) Handles ViewDatabaseData.FormClosed
        ViewDatabaseData = Nothing
    End Sub

    Private Sub btnApplyQuery_Click(sender As Object, e As EventArgs) Handles btnApplyQuery.Click
        'PointChart.InputQuery = txtInputQuery.Text
        ChartInfo.InputQuery = txtInputQuery.Text
        ApplyQuery() 'This fills the Dataset with the Query data.

        'THEFOLLOWING CODE is now included in ApplyQuery():
        'GetChartOptionsFromDataset() 'This updates the chart options using Dataset fields.
        'UpdateSeriesTabSettings()
    End Sub

    Private Sub GetChartOptionsFromDataset()
        'Update the Chart display options from the Dataset.

        'First check that data to chart has been loaded into the dataset:
        'If ds.Tables.Count = 0 Then
        If ChartInfo.ds.Tables.Count = 0 Then
            Message.AddWarning("No data has been selected for charting." & vbCrLf)
        Else
            'Get the list of available fields from the dataset:
            GetFieldListFromDataset()
        End If

        'Show the selected XValues field: --------------------------------------------------------------------------------
        'Dim I As Integer 'Loop index
        'For I = 1 To cmbXValues.Items.Count
        '    If cmbXValues.Items(I - 1) = PointChart.XValuesFieldName Then
        '        cmbXValues.SelectedIndex = I - 1
        '    End If
        'Next

        Dim SeriesName As String = txtSeriesName.Text.Trim
        If SeriesName = "" Then
            'Message.AddWarning("The Series Name is blank" & vbCrLf)
            'ElseIf ChartInfo.dictFields.ContainsKey(SeriesName) Then
        ElseIf ChartInfo.dictSeriesInfo.ContainsKey(SeriesName) Then
            Dim I As Integer 'Loop index
            For I = 1 To cmbXValues.Items.Count
                'The Series Name is in txtSeriesName.Text
                If cmbXValues.Items(I - 1) = ChartInfo.dictSeriesInfo(txtSeriesName.Text).XValuesFieldName Then
                    cmbXValues.SelectedIndex = I - 1
                End If
            Next
        Else
            Message.AddWarning("The Series Name is not in the Chart Info dictionary: " & SeriesName & vbCrLf)
        End If

    End Sub

    Private Sub GetFieldListFromDataset()
        'Update the available list of fields for plotting on the X and Y axes.

        'cboFieldSelections.Items.Clear()
        cmbXValues.Items.Clear()
        cboFieldSelections.Items.Clear()

        'If ds.Tables(0).Columns.Count > 0 Then
        If ChartInfo.ds.Tables(0).Columns.Count > 0 Then
            Dim I As Integer 'Loop index
            'For I = 1 To ds.Tables(0).Columns.Count
            For I = 1 To ChartInfo.ds.Tables(0).Columns.Count
                cboFieldSelections.Items.Add(ChartInfo.ds.Tables(0).Columns(I - 1).ColumnName) 'ComboBox column used in DataGridView1 (Y Values to chart)
                cmbXValues.Items.Add(ChartInfo.ds.Tables(0).Columns(I - 1).ColumnName) 'ComboBox used to select the Field to use along the X Axis
            Next
        End If
    End Sub

    Public Sub ApplyQuery()
        'Use the query to fill the ds dataset

        'If PointChart.InputDatabasePath = "" Then
        If ChartInfo.InputDatabasePath = "" Then
            Message.AddWarning("InputDatabasePath is not defined!" & vbCrLf)
            Exit Sub
        End If

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim commandString As String 'Declare a command string - contains the query to be passed to the database.

        'Specify the connection string (Access 2007):
        connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        "data source = " + ChartInfo.InputDatabasePath

        'Connect to the Access database:
        conn = New System.Data.OleDb.OleDbConnection(connectionString)
        conn.Open()

        'Specify the commandString to query the database:
        commandString = ChartInfo.InputQuery
        Dim dataAdapter As New System.Data.OleDb.OleDbDataAdapter(commandString, conn)

        ChartInfo.ds.Clear()
        ChartInfo.ds.Reset()

        dataAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey

        Try
            dataAdapter.Fill(ChartInfo.ds, "SelTable")
            'UpdateChartQuery() 'NOT NEEDED??? 'This was originally used to set PointChart or StockChart .Input Query to the property InputQuery. (See the Chart app code.)
            GetChartOptionsFromDataset() 'This updates the chart options using Dataset fields.
            UpdateSeriesTabSettings()
        Catch ex As Exception
            Message.AddWarning("Error applying query." & vbCrLf)
            Message.AddWarning(ex.Message & vbCrLf)
        End Try

        conn.Close()

    End Sub

    Private Sub txtDataDescription_LostFocus(sender As Object, e As EventArgs) Handles txtDataDescription.LostFocus
        ChartInfo.InputDataDescr = txtDataDescription.Text
    End Sub

    Private Sub txtInputQuery_LostFocus(sender As Object, e As EventArgs) Handles txtInputQuery.LostFocus
        ChartInfo.InputQuery = txtInputQuery.Text
    End Sub

#End Region 'Chart Input Data Tab ---------------------------------------------------------------------------------------------------------------------------------------

#Region " Chart Titles Tab" '============================================================================================================================================

    Private Sub UpdateTitlesTabSettings()
        'Update the Titles tab settings from the Chart control.

        Dim NTitlesRecords As Integer = Chart1.Titles.Count
        txtNTitlesRecords.Text = NTitlesRecords

        Dim TitleNo As Integer = Val(txtTitlesRecordNo.Text)

        If TitleNo + 1 > NTitlesRecords Then TitleNo = NTitlesRecords - 1

        If TitleNo >= 0 Then
            txtTitlesRecordNo.Text = TitleNo + 1
            txtChartTitle.Text = Chart1.Titles(TitleNo).Text
            txtTitleName.Text = Chart1.Titles(TitleNo).Name

            cmbAlignment.SelectedIndex = cmbAlignment.FindStringExact(Chart1.Titles(TitleNo).Alignment.ToString)
            cmbOrientation.SelectedIndex = cmbOrientation.FindStringExact(Chart1.Titles(TitleNo).TextOrientation.ToString)

            txtChartTitle.Font = Chart1.Titles(TitleNo).Font
            txtChartTitle.ForeColor = Chart1.Titles(TitleNo).ForeColor
        Else
            txtTitlesRecordNo.Text = 0
            txtChartTitle.Text = ""
            txtTitleName.Text = ""
        End If
    End Sub

    Private Sub ApplyTitlesTabSettings()
        'Apply the Titles tab settings to the Chart control.

        Dim TitleNo As Integer = Val(txtTitlesRecordNo.Text) - 1

        Dim TitleCount As Integer = Chart1.Titles.Count

        If TitleNo > TitleCount Then
            'Dim I As Integer
            'For I = 0 To TitleNo - TitleCount
            '    Chart1.Titles.Add("Label" & I)
            'Next
            Message.AddWarning("The title number is larger than to number of titles in the chart!" & vbCrLf)
            Exit Sub
        End If

        Chart1.Titles(TitleNo).Name = txtTitleName.Text.Trim
        Chart1.Titles(TitleNo).Text = txtChartTitle.Text
        Chart1.Titles(TitleNo).ForeColor = txtChartTitle.ForeColor
        Chart1.Titles(TitleNo).Font = txtChartTitle.Font

        Select Case cmbAlignment.SelectedItem.ToString
            Case "BottomCenter"
                Chart1.Titles(TitleNo).Alignment = ContentAlignment.BottomCenter
            Case "BottomLeft"
                Chart1.Titles(TitleNo).Alignment = ContentAlignment.BottomLeft
            Case "BottomRight"
                Chart1.Titles(TitleNo).Alignment = ContentAlignment.BottomRight
            Case "MiddleCenter"
                Chart1.Titles(TitleNo).Alignment = ContentAlignment.MiddleCenter
            Case "MiddleLeft"
                Chart1.Titles(TitleNo).Alignment = ContentAlignment.MiddleLeft
            Case "MiddleRight"
                Chart1.Titles(TitleNo).Alignment = ContentAlignment.MiddleRight
            Case "TopCenter"
                Chart1.Titles(TitleNo).Alignment = ContentAlignment.TopCenter
            Case "TopLeft"
                Chart1.Titles(TitleNo).Alignment = ContentAlignment.TopLeft
            Case "TopRight"
                Chart1.Titles(TitleNo).Alignment = ContentAlignment.TopRight
            Case "BottomRight"
                Chart1.Titles(TitleNo).Alignment = ContentAlignment.BottomRight
            Case Else
                Chart1.Titles(TitleNo).Alignment = ContentAlignment.TopCenter
        End Select

        Select Case cmbOrientation.SelectedItem.ToString
            Case "Auto"
                Chart1.Titles(TitleNo).TextOrientation = DataVisualization.Charting.TextOrientation.Auto
            Case "Horizontal"
                Chart1.Titles(TitleNo).TextOrientation = DataVisualization.Charting.TextOrientation.Horizontal
            Case "Rotated270"
                Chart1.Titles(TitleNo).TextOrientation = DataVisualization.Charting.TextOrientation.Rotated270
            Case "Rotated90"
                Chart1.Titles(TitleNo).TextOrientation = DataVisualization.Charting.TextOrientation.Rotated90
            Case "Stacked"
                Chart1.Titles(TitleNo).TextOrientation = DataVisualization.Charting.TextOrientation.Stacked
            Case Else
                Chart1.Titles(TitleNo).TextOrientation = DataVisualization.Charting.TextOrientation.Auto
        End Select

    End Sub

    Private Sub btnAddTitle_Click(sender As Object, e As EventArgs) Handles btnAddTitle.Click
        'Add a new Chart Title:

        Dim NewTitleNo As Integer = Chart1.Titles.Count + 1
        If NewTitleNo = 0 Then NewTitleNo = 1 'The first title should be named Title1

        Message.Add("Chart1.Titles.Count = " & Chart1.Titles.Count & vbCrLf)
        Dim NewTitleName As String = "Label" & NewTitleNo
        Chart1.Titles.Add(NewTitleName)

        txtTitleName.Text = NewTitleName
        'txtTitlesRecordNo.Text = NewTitleNo + 1
        txtTitlesRecordNo.Text = NewTitleNo
        txtNTitlesRecords.Text = Chart1.Titles.Count
        txtChartTitle.Text = ""

    End Sub

    Private Sub btnApplyTitlesSettings_Click(sender As Object, e As EventArgs) Handles btnApplyTitlesSettings.Click
        'Apply the Titles settings:
        ApplyTitlesTabSettings()
    End Sub


#End Region 'Chart Titles Tab -------------------------------------------------------------------------------------------------------------------------------------------

#Region " Chart Series Tab" '============================================================================================================================================

    'Private Sub SetupLineChartSeriesTab()
    Private Sub SetupPointChartSeriesTab()
        'Set up the series tab for a Point Chart:

        'List of Chart Types:
        'https://msdn.microsoft.com/en-us/data/dd489233(v=vs.95)
        'Point Chart Characteristics:
        'https://msdn.microsoft.com/en-us/data/dd456684(v=vs.95)


        'Set up Point chart:
        txtChartDescr.Text = "Line Chart" & vbCrLf
        txtChartDescr.Text = txtChartDescr.Text & "The Point chart type uses value points to represent its data." & vbCrLf
        txtChartDescr.Text = txtChartDescr.Text & "Custom Attributes:" & vbCrLf
        txtChartDescr.Text = txtChartDescr.Text & "EmptyPointValue - Specifies the value to be used for empty points. This property determines how an empty point is treated when the chart is drawn. (Average, Zero)" & vbCrLf
        txtChartDescr.Text = txtChartDescr.Text & "LabelStyle - Specifies the label position of the data point. (Auto, Top, Bottom, Right, Left, TopLeft, TopRight, BottomLeft, BottomRight, Center)" & vbCrLf
        txtChartDescr.Text = txtChartDescr.Text & "PixelPointDepth - Specifies the 3D series depth in pixels. (Any integer > 0)" & vbCrLf
        txtChartDescr.Text = txtChartDescr.Text & "PixelPointGapDepth - Specifies the 3D gap depth in pixels. (Any integer > 0)" & vbCrLf
        'txtChartDescr.Text = txtChartDescr.Text & "ShowMarkerLines - Specifies whether marker lines are displayed when rendered in 3D. (True, False)" & vbCrLf

        'Y Values:
        DataGridView1.Rows.Clear()
        DataGridView1.Rows.Add(1)
        DataGridView1.Rows(0).Cells(0).Value = "Yvalue" 'Y Value Parameter Name

        'Custom Attributes:
        DataGridView2.Rows.Clear()
        DataGridView2.Rows.Add(5)

        '  EmptyPointValue:
        DataGridView2.Rows(0).Cells(0).Value = "EmptyPointValue"
        DataGridView2.Rows(0).Cells(1).Value = "Average, Zero"
        Dim cbc0 As New DataGridViewComboBoxCell
        cbc0.Items.Add(" ")
        cbc0.Items.Add("Average")
        cbc0.Items.Add("Zero")
        DataGridView2.Rows(0).Cells(2) = cbc0

        '  LabelStyle:
        DataGridView2.Rows(1).Cells(0).Value = "LabelStyle"
        DataGridView2.Rows(1).Cells(1).Value = "Auto, Top, Bottom, Right, Left, TopLeft, TopRight, BottomLeft, BottomRight, Center"
        Dim cbc1 As New DataGridViewComboBoxCell
        cbc1.Items.Add(" ")
        cbc1.Items.Add("Auto")
        cbc1.Items.Add("Top")
        cbc1.Items.Add("Bottom")
        cbc1.Items.Add("Right")
        cbc1.Items.Add("Left")
        cbc1.Items.Add("TopLeft")
        cbc1.Items.Add("TopRight")
        cbc1.Items.Add("BottomLeft")
        cbc1.Items.Add("BottomRight")
        cbc1.Items.Add("Center")
        DataGridView2.Rows(1).Cells(2) = cbc1

        '  PixelPointDepth:
        DataGridView2.Rows(2).Cells(0).Value = "PixelPointDepth"
        DataGridView2.Rows(2).Cells(1).Value = "Any integer > 0"

        '  PixelPointGapDepth:
        DataGridView2.Rows(3).Cells(0).Value = "PixelPointGapDepth"
        DataGridView2.Rows(3).Cells(1).Value = "Any integer > 0"

        ''  ShowMarkerLines:
        'DataGridView2.Rows(4).Cells(0).Value = "ShowMarkerLines"
        'DataGridView2.Rows(4).Cells(1).Value = "True, False"
        'Dim cbc4 As New DataGridViewComboBoxCell
        'cbc4.Items.Add(" ")
        'cbc4.Items.Add("True")
        'cbc4.Items.Add("False")
        'DataGridView2.Rows(4).Cells(2) = cbc4

    End Sub

    Private Sub btnApplySeriesSettings_Click(sender As Object, e As EventArgs) Handles btnApplySeriesSettings.Click
        ApplySeriesTabSettings()
    End Sub

    Private Sub ApplySeriesTabSettings()
        'Update the Chart control settings from the Series tab.

        Dim SeriesNo As Integer = Val(txtSeriesRecordNo.Text) - 1
        Dim SeriesCount As Integer = Chart1.Series.Count

        If SeriesNo > SeriesCount Then
            Message.AddWarning("The series number is larger than to number of series in the chart!" & vbCrLf)
            Exit Sub
        End If

        Dim SeriesName As String = txtSeriesName.Text.Trim
        Chart1.Series(SeriesNo).Name = SeriesName
        If ChartInfo.dictSeriesInfo.ContainsKey(SeriesName) Then
            'SeriesName is already in the dictionary of database fields.
        Else
            ChartInfo.dictSeriesInfo.Add(SeriesName, New SeriesInfo)
        End If

        'Chart1.Series(SeriesNo).ChartType = DataVisualization.Charting.SeriesChartType.Line
        Chart1.Series(SeriesNo).ChartType = DataVisualization.Charting.SeriesChartType.Point
        'Chart1.Series(SeriesNo).Points.DataBindXY(ChartInfo.ds.Tables(0).DefaultView, ChartInfo.dictFields(SeriesName).XValuesFieldName, ChartInfo.ds.Tables(0).DefaultView, ChartInfo.dictFields(SeriesName).YValuesFieldName)

        'Chart1.Series(SeriesNo).ChartArea = cmbChartArea.SelectedItem.ToString
        If IsNothing(cmbChartArea.SelectedItem) Then
        Else
            ChartInfo.dictSeriesInfo(SeriesName).ChartArea = cmbChartArea.SelectedItem.ToString
            Chart1.Series(SeriesNo).ChartArea = cmbChartArea.SelectedItem.ToString
        End If
        'ChartInfo.dictSeriesInfo(SeriesName).ChartArea = cmbChartArea.SelectedItem.ToString

        'ChartInfo.dictFields(SeriesName).XValuesFieldName = cmbXValues.SelectedText
        'Message.Add("cmbXValues.SelectedText" & cmbXValues.SelectedText & vbCrLf)
        'ChartInfo.dictFields(SeriesName).XValuesFieldName = cmbXValues.SelectedItem.ToString
        ChartInfo.dictSeriesInfo(SeriesName).XValuesFieldName = cmbXValues.SelectedItem.ToString
        'Message.Add("cmbXValues.SelectedItem.ToString = " & cmbXValues.SelectedItem.ToString & vbCrLf)

        'ChartInfo.dictFields(SeriesName).YValuesFieldName = DataGridView1.Rows(0).Cells(1).Value
        ChartInfo.dictSeriesInfo(SeriesName).YValuesFieldName = DataGridView1.Rows(0).Cells(1).Value

        'Select Case cmbXAxisType.SelectedText
        Select Case cmbXAxisType.SelectedItem.ToString
            Case "Primary"
                Chart1.Series(SeriesName).XAxisType = DataVisualization.Charting.AxisType.Primary
            Case "Secondary"
                Chart1.Series(SeriesName).XAxisType = DataVisualization.Charting.AxisType.Secondary
            Case Else
                Chart1.Series(SeriesName).XAxisType = DataVisualization.Charting.AxisType.Primary
        End Select

        Select Case cmbXAxisValueType.SelectedItem.ToString
            Case "Auto"
                Chart1.Series(SeriesName).XValueType = DataVisualization.Charting.ChartValueType.Auto
            Case "Date"
                Chart1.Series(SeriesName).XValueType = DataVisualization.Charting.ChartValueType.Date
            Case "DateTime"
                Chart1.Series(SeriesName).XValueType = DataVisualization.Charting.ChartValueType.DateTime
            Case "DateTimeOffset"
                Chart1.Series(SeriesName).XValueType = DataVisualization.Charting.ChartValueType.DateTimeOffset
            Case "Double"
                Chart1.Series(SeriesName).XValueType = DataVisualization.Charting.ChartValueType.Double
            Case "Int32"
                Chart1.Series(SeriesName).XValueType = DataVisualization.Charting.ChartValueType.Int32
            Case "Int64"
                Chart1.Series(SeriesName).XValueType = DataVisualization.Charting.ChartValueType.Int64
            Case "Single"
                Chart1.Series(SeriesName).XValueType = DataVisualization.Charting.ChartValueType.Single
            Case "String"
                Chart1.Series(SeriesName).XValueType = DataVisualization.Charting.ChartValueType.String
            Case "Time"
                Chart1.Series(SeriesName).XValueType = DataVisualization.Charting.ChartValueType.Time
            Case "UInt32"
                Chart1.Series(SeriesName).XValueType = DataVisualization.Charting.ChartValueType.UInt32
            Case "UInt64"
                Chart1.Series(SeriesName).XValueType = DataVisualization.Charting.ChartValueType.UInt64
            Case Else
                Chart1.Series(SeriesName).XValueType = DataVisualization.Charting.ChartValueType.Auto
        End Select

        Select Case cmbYAxisType.SelectedItem.ToString
            Case "Primary"
                Chart1.Series(SeriesName).YAxisType = DataVisualization.Charting.AxisType.Primary
            Case "Secondary"
                Chart1.Series(SeriesName).YAxisType = DataVisualization.Charting.AxisType.Secondary
            Case Else
                Chart1.Series(SeriesName).YAxisType = DataVisualization.Charting.AxisType.Primary
        End Select

        Select Case cmbYAxisValueType.SelectedItem.ToString
            Case "Auto"
                Chart1.Series(SeriesName).YValueType = DataVisualization.Charting.ChartValueType.Auto
            Case "Date"
                Chart1.Series(SeriesName).YValueType = DataVisualization.Charting.ChartValueType.Date
            Case "DateTime"
                Chart1.Series(SeriesName).YValueType = DataVisualization.Charting.ChartValueType.DateTime
            Case "DateTimeOffset"
                Chart1.Series(SeriesName).YValueType = DataVisualization.Charting.ChartValueType.DateTimeOffset
            Case "Double"
                Chart1.Series(SeriesName).YValueType = DataVisualization.Charting.ChartValueType.Double
            Case "Int32"
                Chart1.Series(SeriesName).YValueType = DataVisualization.Charting.ChartValueType.Int32
            Case "Int64"
                Chart1.Series(SeriesName).YValueType = DataVisualization.Charting.ChartValueType.Int64
            Case "Single"
                Chart1.Series(SeriesName).YValueType = DataVisualization.Charting.ChartValueType.Single
            Case "String"
                Chart1.Series(SeriesName).YValueType = DataVisualization.Charting.ChartValueType.String
            Case "Time"
                Chart1.Series(SeriesName).YValueType = DataVisualization.Charting.ChartValueType.Time
            Case "UInt32"
                Chart1.Series(SeriesName).YValueType = DataVisualization.Charting.ChartValueType.UInt32
            Case "UInt64"
                Chart1.Series(SeriesName).YValueType = DataVisualization.Charting.ChartValueType.UInt64
            Case Else
                Chart1.Series(SeriesName).YValueType = DataVisualization.Charting.ChartValueType.Auto
        End Select


        'Save custom attributes:
        If DataGridView2.Rows(0).Cells(2).Value = "" Then
            Chart1.Series(SeriesName).SetCustomProperty("EmptyPointValue", "Average")
        Else
            Chart1.Series(SeriesName).SetCustomProperty("EmptyPointValue", DataGridView2.Rows(0).Cells(2).Value)
        End If
        If DataGridView2.Rows(1).Cells(2).Value = "" Then
            Chart1.Series(SeriesName).SetCustomProperty("LabelStyle", "Auto")
        Else
            Chart1.Series(SeriesName).SetCustomProperty("LabelStyle", DataGridView2.Rows(1).Cells(2).Value)
        End If
        If DataGridView2.Rows(2).Cells(2).Value = "" Then
            Chart1.Series(SeriesName).SetCustomProperty("PixelPointDepth", "1")
        Else
            Chart1.Series(SeriesName).SetCustomProperty("PixelPointDepth", DataGridView2.Rows(2).Cells(2).Value)
        End If
        If DataGridView2.Rows(3).Cells(2).Value = "" Then
            Chart1.Series(SeriesName).SetCustomProperty("PixelPointGapDepth", "1")
        Else
            Chart1.Series(SeriesName).SetCustomProperty("PixelPointGapDepth", DataGridView2.Rows(3).Cells(2).Value)
        End If

        'NOTE: THE FOLLOWING IS NOT USED FOR A POINT CHART:
        'If DataGridView2.Rows(4).Cells(2).Value = "" Then
        '    Chart1.Series(SeriesName).SetCustomProperty("ShowMarkerLines", "True")
        'Else
        '    Chart1.Series(SeriesName).SetCustomProperty("ShowMarkerLines", DataGridView2.Rows(4).Cells(2).Value)
        'End If

        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.Area
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.Bar
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.BoxPlot
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.Bubble
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.Candlestick
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.Column
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.Doughnut
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.ErrorBar
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.FastLine
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.FastPoint
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.Funnel
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.Kagi
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.Line
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.Pie
        Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.Point
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.PointAndFigure
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.Polar
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.Pyramid
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.Radar
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.Range
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.RangeBar
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.RangeColumn
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.Renko
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.Spline
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.SplineArea
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.SplineRange
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.StackedArea
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.StackedArea100
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.StackedBar
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.StackedBar100
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.StackedColumn
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.StackedColumn100
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.StepLine
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.Stock
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.ThreeLineBreak

        'Chart1.Series(SeriesName).ChartArea = "Area1"



        'Chart1.Series(SeriesName).MarkerStyle = DataVisualization.Charting.MarkerStyle.Circle
        'Chart1.Series(SeriesName).MarkerBorderColor = Color.Black
        'Chart1.Series(SeriesName).BorderWidth = 2
        'Chart1.Series(SeriesName).BorderColor = Color.Blue



        'Chart1.Series(SeriesName).XValueType = DataVisualization.Charting.ChartValueType.Auto
        'Chart1.Series(SeriesName).XValueType = DataVisualization.Charting.ChartValueType.Date
        'Chart1.Series(SeriesName).XValueType = DataVisualization.Charting.ChartValueType.Date

        'Check the Series colors:
        If Chart1.Series(SeriesName).MarkerBorderColor = Color.FromArgb(0) Then Chart1.Series(SeriesName).MarkerBorderColor = Color.Black
        If Chart1.Series(SeriesName).MarkerColor = Color.FromArgb(0) Then Chart1.Series(SeriesName).MarkerColor = Color.Gray
        If Chart1.Series(SeriesName).Color = Color.FromArgb(0) Then Chart1.Series(SeriesName).Color = Color.Blue

    End Sub

    Private Sub btnAddSeries_Click(sender As Object, e As EventArgs) Handles btnAddSeries.Click
        'Add a new series to the Chart:

        If txtSeriesRecordNo.Text = "0" Then
            'No Point Chart data has been loaded.
            'Check if there is a default Series already in the Chart control:
            If Chart1.Series.Count > 0 Then
                txtSeriesRecordNo.Text = "1"
                txtNSeriesRecords.Text = Chart1.Series.Count
                txtSeriesName.Text = Chart1.Series(0).Name
            Else
                'To Do!!!

            End If

        Else 'Add a new Series to the Chart control:
            'Dim NewSeriesNo As Integer = Chart1.Series.Count
            Dim NewSeriesNo As Integer = Chart1.Series.Count + 1

            Dim NewSeriesName As String = "Series" & NewSeriesNo
            Chart1.Series.Add(NewSeriesName)

            txtSeriesName.Text = NewSeriesName
            txtSeriesRecordNo.Text = NewSeriesNo + 1
            txtNSeriesRecords.Text = Chart1.Series.Count
        End If
    End Sub

    Private Sub UpdateSeriesTabSettings()
        'Update the Series tab settings from the Chart control and ChartInfo.

        Dim NSeries As Integer = Chart1.Series.Count
        'txtNSeriesRecords.Text = Chart1.Series.Count
        txtNSeriesRecords.Text = NSeries

        'If txtSeriesRecordNo.Text.Trim = "" Then
        '    txtSeriesRecordNo.Text = "1"
        'End If

        'Dim SeriesNo As Integer = Val(txtSeriesRecordNo.Text) - 1
        Dim SeriesNo As Integer

        'Dim SeriesName As String = Chart1.Series(SeriesNo).Name
        Dim SeriesName As String

        If NSeries = 0 Then
            txtSeriesRecordNo.Text = "0"
            SeriesNo = 0
            SeriesName = ""
            txtSeriesName.Text = ""
            cmbXValues.SelectedIndex = 0
        Else
            If txtSeriesRecordNo.Text.Trim = "" Then
                txtSeriesRecordNo.Text = "1"
            End If
            'SeriesNo = 1
            SeriesNo = Val(txtSeriesRecordNo.Text)
            If SeriesNo < 1 Then
                SeriesNo = 1
                'If SeriesNo > NSeries Then
                '    SeriesNo = NSeries
                'End If
            End If
            If SeriesNo > NSeries Then SeriesNo = NSeries
            txtSeriesRecordNo.Text = SeriesNo
            'txtSeriesName.Text = Chart1.Series(0).Name
            'SeriesName = Chart1.Series(SeriesNo).Name
            SeriesName = Chart1.Series(SeriesNo - 1).Name
            txtSeriesName.Text = SeriesName

            'Update list of areas in Series tab:
            cmbChartArea.Items.Clear()
            For Each item In Chart1.ChartAreas
                cmbChartArea.Items.Add(item.Name)
                'Message.Add("Adding Chare Area: " & item.Name & vbCrLf)
            Next

            If SeriesName = "" Then
                Message.AddWarning("The Series Name is blank" & vbCrLf)
                'ElseIf ChartInfo.dictFields.ContainsKey(SeriesName) Then
            ElseIf ChartInfo.dictSeriesInfo.ContainsKey(SeriesName) Then
                'Apply X Values Field Name:
                'cmbXValues.SelectedIndex = cmbXValues.FindStringExact(ChartInfo.dictFields(SeriesName).XValuesFieldName)
                cmbXValues.SelectedIndex = cmbXValues.FindStringExact(ChartInfo.dictSeriesInfo(SeriesName).XValuesFieldName)

                cmbXAxisType.SelectedIndex = cmbXAxisType.FindStringExact(Chart1.Series(SeriesName).XAxisType.ToString)
                cmbXAxisValueType.SelectedIndex = cmbXAxisValueType.FindStringExact(Chart1.Series(SeriesName).XValueType.ToString)

                'cmbChartArea.SelectedIndex = cmbChartArea.FindStringExact(ChartInfo.dictSeriesInfo(SeriesName).ChartArea)
                cmbChartArea.SelectedIndex = cmbChartArea.FindStringExact(Chart1.Series(SeriesName).ChartArea)
                'Apply Y Values selections:
                For I = 1 To cboFieldSelections.Items.Count
                    'If PointChart.YValuesFieldName = cboFieldSelections.Items(I - 1) Then
                    'If ChartInfo.dictFields(SeriesName).YValuesFieldName = cboFieldSelections.Items(I - 1) Then
                    If ChartInfo.dictSeriesInfo(SeriesName).YValuesFieldName = cboFieldSelections.Items(I - 1) Then
                        DataGridView1.Rows(0).Cells(1).Value = cboFieldSelections.Items(I - 1)
                    End If
                Next

                cmbYAxisType.SelectedIndex = cmbYAxisType.FindStringExact(Chart1.Series(SeriesName).YAxisType.ToString)
                cmbYAxisValueType.SelectedIndex = cmbYAxisValueType.FindStringExact(Chart1.Series(SeriesName).YValueType.ToString)

                'Apply Custom Attributes selections:
                'If Chart1.Series(SeriesName).CustomProperties("EmptyPointValue") <> "" Then
                If Chart1.Series(SeriesName).GetCustomProperty("EmptyPointValue") <> "" Then
                    'DataGridView2.Rows(0).Cells(2).Value = Chart1.Series(SeriesName).CustomProperties("EmptyPointValue")
                    DataGridView2.Rows(0).Cells(2).Value = Chart1.Series(SeriesName).GetCustomProperty("EmptyPointValue")
                End If
                'If Chart1.Series(SeriesName).CustomProperties("LabelStyle") <> "" Then
                If Chart1.Series(SeriesName).GetCustomProperty("LabelStyle") <> "" Then
                    'DataGridView2.Rows(1).Cells(2).Value = Chart1.Series(SeriesName).CustomProperties("LabelStyle")
                    DataGridView2.Rows(1).Cells(2).Value = Chart1.Series(SeriesName).GetCustomProperty("LabelStyle")
                End If
                'DataGridView2.Rows(2).Cells(2).Value = Chart1.Series(SeriesName).CustomProperties("PixelPointDepth")
                DataGridView2.Rows(2).Cells(2).Value = Chart1.Series(SeriesName).GetCustomProperty("PixelPointDepth")
                'DataGridView2.Rows(3).Cells(2).Value = Chart1.Series(SeriesName).CustomProperties("PixelPointGapDepth")
                DataGridView2.Rows(3).Cells(2).Value = Chart1.Series(SeriesName).GetCustomProperty("PixelPointGapDepth")

                'DataGridView2.Rows(4).Cells(2).Value = Chart1.Series(SeriesName).GetCustomProperty("ShowMarkerLines") 'NOTE USED IN A POINT CHART
            Else
                Message.AddWarning("The Series Name is not in the Chart Info dictionary: " & SeriesName & vbCrLf)
            End If
        End If


        ''Apply Custom Attributes selections:
        ''LabelValueType (High, Low, Open, Close)
        ''If PointChart.EmptyPointValue <> "" Then
        'If Chart1.Series(SeriesName).CustomProperties("EmptyPointValue") <> "" Then
        '    'DataGridView2.Rows(0).Cells(2).Value = PointChart.EmptyPointValue
        '    DataGridView2.Rows(0).Cells(2).Value = Chart1.Series(SeriesName).CustomProperties("EmptyPointValue")
        'End If
        ''If PointChart.LabelStyle <> "" Then
        'If Chart1.Series(SeriesName).CustomProperties("LabelStyle") <> "" Then
        '    'DataGridView2.Rows(1).Cells(2).Value = PointChart.LabelStyle
        '    DataGridView2.Rows(1).Cells(2).Value = Chart1.Series(SeriesName).CustomProperties("LabelStyle")
        'End If
        ''PixelPointDepth (Any integer > 0)
        ''DataGridView2.Rows(2).Cells(2).Value = PointChart.PixelPointDepth
        'DataGridView2.Rows(2).Cells(2).Value = Chart1.Series(SeriesName).CustomProperties("PixelPointDepth")
        ''PixelPointGapDepth (Any integer > 0)
        ''DataGridView2.Rows(3).Cells(2).Value = PointChart.PixelPointGapDepth
        'DataGridView2.Rows(3).Cells(2).Value = Chart1.Series(SeriesName).CustomProperties("PixelPointGapDepth")

        'DataGridView2.Rows(4).Cells(2).Value = Chart1.Series(SeriesName).CustomProperties("ShowMarkerLines")

    End Sub



#End Region 'Chart Series Tab -------------------------------------------------------------------------------------------------------------------------------------------

#Region " Chart Areas Tab" '=============================================================================================================================================

    Private Sub UpdateAreasTabSettings()
        'Update the Areas tab settings from the Chart control.

        Dim NAreas As Integer = Chart1.ChartAreas.Count
        txtNAreaRecords.Text = NAreas

        Dim AreaNo As Integer
        Dim AreaName As String

        If NAreas = 0 Then
            txtAreaRecordNo.Text = "0"
            AreaNo = 0
            AreaName = ""
            txtAreaName.Text = ""

        Else
            If txtAreaRecordNo.Text.Trim = "" Then
                txtAreaRecordNo.Text = "1"
            End If

            'AreaNo = Val(txtAreaRecordNo.Text)
            AreaNo = Val(txtAreaRecordNo.Text) - 1 'Zero-based area number.
            'If AreaNo < 1 Then
            '    AreaNo = 1
            'End If
            If AreaNo < 0 Then
                AreaNo = 0
            End If
            'If AreaNo > NAreas Then AreaNo = NAreas
            If AreaNo + 1 > NAreas Then AreaNo = NAreas - 1
            'txtAreaRecordNo.Text = AreaNo
            'txtAreaRecordNo.Text = AreaNo + 1
            ShowArea(AreaNo)

            'AreaName = Chart1.ChartAreas(AreaNo - 1).Name
            'txtAreaName.Text = AreaName

            'If AreaName = "" Then
            '    Message.AddWarning("The Area Name is blank" & vbCrLf)
            'End If


            'txtXAxisTitle.Text = Chart1.ChartAreas(AreaNo - 1).AxisX.Title
            'txtX2AxisTitle.Text = Chart1.ChartAreas(AreaNo - 1).AxisX2.Title
            'txtYAxisTitle.Text = Chart1.ChartAreas(AreaNo - 1).AxisY.Title
            'txtY2AxisTitle.Text = Chart1.ChartAreas(AreaNo - 1).AxisY2.Title

            'txtXAxisTitle.Font = Chart1.ChartAreas(AreaNo - 1).AxisX.TitleFont
            'txtX2AxisTitle.Font = Chart1.ChartAreas(AreaNo - 1).AxisX2.TitleFont
            'txtYAxisTitle.Font = Chart1.ChartAreas(AreaNo - 1).AxisY.TitleFont
            'txtY2AxisTitle.Font = Chart1.ChartAreas(AreaNo - 1).AxisY2.TitleFont

            'txtXAxisTitle.ForeColor = Chart1.ChartAreas(AreaNo - 1).AxisX.TitleForeColor
            'txtX2AxisTitle.ForeColor = Chart1.ChartAreas(AreaNo - 1).AxisX2.TitleForeColor
            'txtYAxisTitle.ForeColor = Chart1.ChartAreas(AreaNo - 1).AxisY.TitleForeColor
            'txtY2AxisTitle.ForeColor = Chart1.ChartAreas(AreaNo - 1).AxisY2.TitleForeColor

            'cmbXAxisTitleAlignment.SelectedIndex = cmbXAxisTitleAlignment.FindStringExact(Chart1.ChartAreas(AreaNo - 1).AxisX.TitleAlignment.ToString)
            'cmbX2AxisTitleAlignment.SelectedIndex = cmbX2AxisTitleAlignment.FindStringExact(Chart1.ChartAreas(AreaNo - 1).AxisX2.TitleAlignment.ToString)
            'cmbYAxisTitleAlignment.SelectedIndex = cmbYAxisTitleAlignment.FindStringExact(Chart1.ChartAreas(AreaNo - 1).AxisY.TitleAlignment.ToString)
            'cmbY2AxisTitleAlignment.SelectedIndex = cmbY2AxisTitleAlignment.FindStringExact(Chart1.ChartAreas(AreaNo - 1).AxisY2.TitleAlignment.ToString)


            ''Axis Minimum values: --------------------------------------------
            ''If ChartInfo.dictAreas.ContainsKey(AreaName) Then
            'If ChartInfo.dictAreaInfo.ContainsKey(AreaName) Then
            '    chkXAxisAutoMin.Checked = ChartInfo.dictAreaInfo(AreaName).AutoXAxisMinimum
            '    txtXAxisMin.Text = Chart1.ChartAreas(AreaNo - 1).AxisX.Minimum
            '    txtXAxisZoomFrom.Text = Chart1.ChartAreas(AreaNo - 1).AxisX.Minimum
            '    chkX2AxisAutoMin.Checked = ChartInfo.dictAreaInfo(AreaName).AutoX2AxisMinimum
            '    txtX2AxisMin.Text = Chart1.ChartAreas(AreaNo - 1).AxisX2.Minimum
            '    chkYAxisAutoMin.Checked = ChartInfo.dictAreaInfo(AreaName).AutoYAxisMinimum
            '    txtYAxisMin.Text = Chart1.ChartAreas(AreaNo - 1).AxisY.Minimum
            '    chkY2AxisAutoMin.Checked = ChartInfo.dictAreaInfo(AreaName).AutoY2AxisMinimum
            '    txtY2AxisMin.Text = Chart1.ChartAreas(AreaNo - 1).AxisY2.Minimum
            'End If


            ''Axis Maximum values: -----------------------------------------
            ''If ChartInfo.dictAreas.ContainsKey(AreaName) Then
            'If ChartInfo.dictAreaInfo.ContainsKey(AreaName) Then
            '    chkXAxisAutoMax.Checked = ChartInfo.dictAreaInfo(AreaName).AutoXAxisMaximum
            '    txtXAxisMax.Text = Chart1.ChartAreas(AreaNo - 1).AxisX.Maximum
            '    txtXAxisZoomTo.Text = Chart1.ChartAreas(AreaNo - 1).AxisX.Maximum
            '    txtXAxisZoomInterval.Text = Chart1.ChartAreas(AreaNo - 1).AxisX.Maximum - Chart1.ChartAreas(AreaNo - 1).AxisX.Minimum
            '    chkX2AxisAutoMax.Checked = ChartInfo.dictAreaInfo(AreaName).AutoX2AxisMaximum
            '    txtX2AxisMax.Text = Chart1.ChartAreas(AreaNo - 1).AxisX2.Maximum
            '    chkYAxisAutoMax.Checked = ChartInfo.dictAreaInfo(AreaName).AutoYAxisMaximum
            '    txtYAxisMax.Text = Chart1.ChartAreas(AreaNo - 1).AxisY.Maximum
            '    chkY2AxisAutoMax.Checked = ChartInfo.dictAreaInfo(AreaName).AutoY2AxisMaximum
            '    txtY2AxisMax.Text = Chart1.ChartAreas(AreaNo - 1).AxisY2.Maximum
            'End If

            ''Axis Intervals: -------------------------------------------------------------
            'If Chart1.ChartAreas(AreaNo - 1).AxisX.Interval = 0 Then 'Auto mode.
            '    chkXAxisAutoAnnotInt.Checked = True
            'Else
            '    chkXAxisAutoAnnotInt.Checked = False
            'End If
            'txtXAxisAnnotInt.Text = Chart1.ChartAreas(AreaNo - 1).AxisX.Interval

            'If Chart1.ChartAreas(AreaNo - 1).AxisX2.Interval = 0 Then 'Auto mode.
            '    chkX2AxisAutoAnnotInt.Checked = True
            'Else
            '    chkX2AxisAutoAnnotInt.Checked = False
            'End If
            'txtX2AxisAnnotInt.Text = Chart1.ChartAreas(AreaNo - 1).AxisX2.Interval
            'If Chart1.ChartAreas(AreaNo - 1).AxisY.Interval = 0 Then 'Auto mode.
            '    chkYAxisAutoAnnotInt.Checked = True
            'Else
            '    chkYAxisAutoAnnotInt.Checked = False
            'End If
            'txtYAxisAnnotInt.Text = Chart1.ChartAreas(AreaNo - 1).AxisY.Interval
            'If Chart1.ChartAreas(AreaNo - 1).AxisY2.Interval = 0 Then 'Auto mode.
            '    chkY2AxisAutoAnnotInt.Checked = True
            'Else
            '    chkY2AxisAutoAnnotInt.Checked = False
            'End If
            'txtY2AxisAnnotInt.Text = Chart1.ChartAreas(AreaNo - 1).AxisY2.Interval


            ''Axis Major Grid Intervals: -----------------------------------------------------------
            ''If ChartInfo.dictAreas.ContainsKey(AreaName) Then
            'If ChartInfo.dictAreaInfo.ContainsKey(AreaName) Then
            '    chkXAxisAutoMajGridInt.Checked = ChartInfo.dictAreaInfo(AreaName).AutoXAxisMajorGridInterval
            '    txtXAxisMajGridInt.Text = Chart1.ChartAreas(AreaNo - 1).AxisX.MajorGrid.Interval
            '    chkX2AxisAutoMajGridInt.Checked = ChartInfo.dictAreaInfo(AreaName).AutoX2AxisMajorGridInterval
            '    txtX2AxisMajGridInt.Text = Chart1.ChartAreas(AreaNo - 1).AxisX2.MajorGrid.Interval
            '    chkYAxisAutoMajGridInt.Checked = ChartInfo.dictAreaInfo(AreaName).AutoYAxisMajorGridInterval
            '    txtYAxisMajGridInt.Text = Chart1.ChartAreas(AreaNo - 1).AxisY.MajorGrid.Interval
            '    chkY2AxisAutoMajGridInt.Checked = ChartInfo.dictAreaInfo(AreaName).AutoY2AxisMajorGridInterval
            '    txtY2AxisMajGridInt.Text = Chart1.ChartAreas(AreaNo - 1).AxisY2.MajorGrid.Interval
            'End If

            ''Axis Label Style Format:
            'txtXAxisLabelStyleFormat.Text = Chart1.ChartAreas(AreaNo - 1).AxisX.LabelStyle.Format
            'txtX2AxisLabelStyleFormat.Text = Chart1.ChartAreas(AreaNo - 1).AxisX2.LabelStyle.Format
            'txtYAxisLabelStyleFormat.Text = Chart1.ChartAreas(AreaNo - 1).AxisY.LabelStyle.Format
            'txtY2AxisLabelStyleFormat.Text = Chart1.ChartAreas(AreaNo - 1).AxisY2.LabelStyle.Format

            ''Update list of areas in Series tab:
            'cmbChartArea.Items.Clear()
            'For Each item In Chart1.ChartAreas
            '    cmbChartArea.Items.Add(item.Name)
            '    'Message.Add("Adding Chare Area: " & item.Name & vbCrLf)
            'Next
            'Dim SeriesName As String = txtSeriesName.Text
            'cmbChartArea.SelectedItem = cmbChartArea.FindStringExact(Chart1.Series(SeriesName).ChartArea)

        End If

    End Sub

    Private Sub ApplyAreasTabSettings()
        'Apply the settings in the Areas tab to the Chart control.

        Dim AreaNo As Integer = Val(txtAreaRecordNo.Text)
        Dim AreaCount As Integer = Chart1.ChartAreas.Count

        If AreaNo - 1 > AreaCount Then
            Message.AddWarning("The area number is larger than to number of areas in the chart!" & vbCrLf)
            Exit Sub
        End If

        Dim AreaName As String = txtAreaName.Text.Trim
        Chart1.ChartAreas(AreaNo - 1).Name = AreaName

        'If ChartInfo.dictAreas.ContainsKey(AreaName) Then
        If ChartInfo.dictAreaInfo.ContainsKey(AreaName) Then
            'AreaName is already in the dictionary of Area Auto settings.
        Else
            ChartInfo.dictAreaInfo.Add(AreaName, New AreaInfo)
        End If


        'Message.Add("1- Chart1.ChartAreas(0).AxisX.Title = " & Chart1.ChartAreas(0).AxisX.Title & vbCrLf) 'NaN



        'X Axis: -------------------------------------------------------------------------
        Chart1.ChartAreas(AreaNo - 1).AxisX.Title = txtXAxisTitle.Text
        Chart1.ChartAreas(AreaNo - 1).AxisX.TitleFont = txtXAxisTitle.Font
        Chart1.ChartAreas(AreaNo - 1).AxisX.TitleForeColor = txtXAxisTitle.ForeColor
        If cmbXAxisTitleAlignment.SelectedItem IsNot Nothing Then Chart1.ChartAreas(AreaNo - 1).AxisX.TitleAlignment = [Enum].Parse(GetType(StringAlignment), cmbXAxisTitleAlignment.SelectedItem.ToString)
        Chart1.ChartAreas(AreaNo - 1).AxisX.LabelStyle.Format = txtXAxisLabelStyleFormat.Text


        If txtXAxisMin.Text.Trim = "" Then
            chkXAxisAutoMin.Checked = True
            Chart1.ChartAreas(AreaNo - 1).AxisX.Minimum = Double.NaN 'Auto minimum.
            ChartInfo.dictAreaInfo(AreaName).AutoXAxisMinimum = True
        ElseIf chkXAxisAutoMin.Checked Then
            Chart1.ChartAreas(AreaNo - 1).AxisX.Minimum = Double.NaN 'Auto minimum.
            ChartInfo.dictAreaInfo(AreaName).AutoXAxisMinimum = True
        Else
            Chart1.ChartAreas(AreaNo - 1).AxisX.Minimum = Val(txtXAxisMin.Text)
            ChartInfo.dictAreaInfo(AreaName).AutoXAxisMinimum = False
        End If
        'Message.Add("2- Chart1.ChartAreas(0).AxisX.Minimum = " & Chart1.ChartAreas(0).AxisX.Minimum & vbCrLf) 'NaN

        If txtXAxisMax.Text.Trim = "" Then
            chkXAxisAutoMax.Checked = True
            Chart1.ChartAreas(AreaNo - 1).AxisX.Maximum = Double.NaN 'Auto maximum.
            ChartInfo.dictAreaInfo(AreaName).AutoXAxisMaximum = True
        ElseIf chkXAxisAutoMax.Checked Then
            Chart1.ChartAreas(AreaNo - 1).AxisX.Maximum = Double.NaN 'Auto maximum.
            ChartInfo.dictAreaInfo(AreaName).AutoXAxisMaximum = True
        Else
            Chart1.ChartAreas(AreaNo - 1).AxisX.Maximum = Val(txtXAxisMax.Text)
            ChartInfo.dictAreaInfo(AreaName).AutoXAxisMaximum = False
        End If
        'Message.Add("3- Chart1.ChartAreas(0).AxisX.Maximum = " & Chart1.ChartAreas(0).AxisX.Maximum & vbCrLf)

        If txtXAxisAnnotInt.Text.Trim = "" Then
            chkXAxisAutoAnnotInt.Checked = True
            Chart1.ChartAreas(AreaNo - 1).AxisX.Interval = 0 'Zero indicates Auto mode.
            Chart1.ChartAreas(AreaNo - 1).AxisX.IntervalAutoMode = DataVisualization.Charting.IntervalAutoMode.VariableCount
        ElseIf chkXAxisAutoAnnotInt.Checked Then
            Chart1.ChartAreas(AreaNo - 1).AxisX.Interval = 0 'Zero indicates Auto mode.
            Chart1.ChartAreas(AreaNo - 1).AxisX.IntervalAutoMode = DataVisualization.Charting.IntervalAutoMode.VariableCount
        Else
            Chart1.ChartAreas(AreaNo - 1).AxisX.Interval = Val(txtXAxisAnnotInt.Text)
        End If
        'Message.Add("4- Chart1.ChartAreas(0).AxisX.Interval = " & Chart1.ChartAreas(0).AxisX.Interval & vbCrLf)

        If txtXAxisMajGridInt.Text.Trim = "" Then
            chkXAxisAutoMajGridInt.Checked = True
            Chart1.ChartAreas(AreaNo - 1).AxisX.MajorGrid.Interval = Double.NaN 'Indicates Not Set - use Axis Interval value.
            ChartInfo.dictAreaInfo(AreaName).AutoXAxisMajorGridInterval = True
        ElseIf chkXAxisAutoMajGridInt.Checked Then
            Chart1.ChartAreas(AreaNo - 1).AxisX.MajorGrid.Interval = Double.NaN 'Indicates Not Set - use Axis Interval value.
            ChartInfo.dictAreaInfo(AreaName).AutoXAxisMajorGridInterval = True
        Else
            Chart1.ChartAreas(AreaNo - 1).AxisX.MajorGrid.Interval = Val(txtXAxisMajGridInt.Text)
            ChartInfo.dictAreaInfo(AreaName).AutoXAxisMajorGridInterval = False
        End If
        'Message.Add("5- Chart1.ChartAreas(0).AxisX.MajorGrid.Interval = " & Chart1.ChartAreas(0).AxisX.MajorGrid.Interval & vbCrLf)

        If chkXAxisScrollBar.Checked Then
            Chart1.ChartAreas(AreaNo - 1).AxisX.ScrollBar.Enabled = True
            Chart1.ChartAreas(AreaNo - 1).AxisX.ScrollBar.Size = 16
        End If


        'X2 Axis: ------------------------------------------------------------------------
        Chart1.ChartAreas(AreaNo - 1).AxisX2.Title = txtX2AxisTitle.Text
        Chart1.ChartAreas(AreaNo - 1).AxisX2.TitleFont = txtX2AxisTitle.Font
        Chart1.ChartAreas(AreaNo - 1).AxisX2.TitleForeColor = txtX2AxisTitle.ForeColor
        If cmbX2AxisTitleAlignment.SelectedItem IsNot Nothing Then Chart1.ChartAreas(AreaNo - 1).AxisX2.TitleAlignment = [Enum].Parse(GetType(StringAlignment), cmbX2AxisTitleAlignment.SelectedItem.ToString)
        Chart1.ChartAreas(AreaNo - 1).AxisX2.LabelStyle.Format = txtX2AxisLabelStyleFormat.Text

        If txtX2AxisMin.Text.Trim = "" Then
            chkX2AxisAutoMin.Checked = True
            Chart1.ChartAreas(AreaNo - 1).AxisX2.Minimum = Double.NaN 'Auto minimum.
            ChartInfo.dictAreaInfo(AreaName).AutoX2AxisMinimum = True
        ElseIf chkX2AxisAutoMin.Checked Then
            Chart1.ChartAreas(AreaNo - 1).AxisX2.Minimum = Double.NaN 'Auto minimum.
            ChartInfo.dictAreaInfo(AreaName).AutoX2AxisMinimum = True
        Else
            Chart1.ChartAreas(AreaNo - 1).AxisX2.Minimum = Val(txtX2AxisMin.Text)
            ChartInfo.dictAreaInfo(AreaName).AutoX2AxisMinimum = False
        End If

        If txtX2AxisMax.Text.Trim = "" Then
            chkX2AxisAutoMax.Checked = True
            Chart1.ChartAreas(AreaNo - 1).AxisX2.Maximum = Double.NaN 'Auto maximum.
            ChartInfo.dictAreaInfo(AreaName).AutoX2AxisMaximum = True
        ElseIf chkX2AxisAutoMax.Checked Then
            Chart1.ChartAreas(AreaNo - 1).AxisX2.Maximum = Double.NaN 'Auto maximum.
            ChartInfo.dictAreaInfo(AreaName).AutoX2AxisMaximum = True
        Else
            Chart1.ChartAreas(AreaNo - 1).AxisX2.Maximum = Val(txtX2AxisMax.Text)
            ChartInfo.dictAreaInfo(AreaName).AutoX2AxisMaximum = False
        End If

        If txtX2AxisAnnotInt.Text.Trim = "" Then
            chkX2AxisAutoAnnotInt.Checked = True
            Chart1.ChartAreas(AreaNo - 1).AxisX2.Interval = 0 'Zero indicates Auto mode.
            Chart1.ChartAreas(AreaNo - 1).AxisX2.IntervalAutoMode = DataVisualization.Charting.IntervalAutoMode.VariableCount
        ElseIf chkX2AxisAutoAnnotInt.Checked Then
            Chart1.ChartAreas(AreaNo - 1).AxisX2.Interval = 0 'Zero indicates Auto mode.
            Chart1.ChartAreas(AreaNo - 1).AxisX2.IntervalAutoMode = DataVisualization.Charting.IntervalAutoMode.VariableCount
        Else
            Chart1.ChartAreas(AreaNo - 1).AxisX2.Interval = Val(txtX2AxisAnnotInt.Text)
        End If

        If txtX2AxisMajGridInt.Text.Trim = "" Then
            chkX2AxisAutoMajGridInt.Checked = True
            Chart1.ChartAreas(AreaNo - 1).AxisX2.MajorGrid.Interval = Double.NaN 'Indicates Not Set - use Axis Interval value.
            ChartInfo.dictAreaInfo(AreaName).AutoX2AxisMajorGridInterval = True
        ElseIf chkX2AxisAutoMajGridInt.Checked Then
            Chart1.ChartAreas(AreaNo - 1).AxisX2.MajorGrid.Interval = Double.NaN 'Indicates Not Set - use Axis Interval value.
            ChartInfo.dictAreaInfo(AreaName).AutoX2AxisMajorGridInterval = True
        Else
            Chart1.ChartAreas(AreaNo - 1).AxisX2.MajorGrid.Interval = Val(txtX2AxisMajGridInt.Text)
            ChartInfo.dictAreaInfo(AreaName).AutoX2AxisMajorGridInterval = False
        End If

        'Y Axis: -------------------------------------------------------------------------
        Chart1.ChartAreas(AreaNo - 1).AxisY.Title = txtYAxisTitle.Text
        Chart1.ChartAreas(AreaNo - 1).AxisY.TitleFont = txtYAxisTitle.Font
        Chart1.ChartAreas(AreaNo - 1).AxisY.TitleForeColor = txtYAxisTitle.ForeColor
        If cmbYAxisTitleAlignment.SelectedItem IsNot Nothing Then Chart1.ChartAreas(AreaNo - 1).AxisY.TitleAlignment = [Enum].Parse(GetType(StringAlignment), cmbYAxisTitleAlignment.SelectedItem.ToString)
        Chart1.ChartAreas(AreaNo - 1).AxisY.LabelStyle.Format = txtYAxisLabelStyleFormat.Text

        If txtYAxisMin.Text.Trim = "" Then
            chkYAxisAutoMin.Checked = True
            Chart1.ChartAreas(AreaNo - 1).AxisY.Minimum = Double.NaN 'Auto minimum.
            ChartInfo.dictAreaInfo(AreaName).AutoYAxisMinimum = True
        ElseIf chkYAxisAutoMin.Checked Then
            Chart1.ChartAreas(AreaNo - 1).AxisY.Minimum = Double.NaN 'Auto minimum.
            ChartInfo.dictAreaInfo(AreaName).AutoYAxisMinimum = True
        Else
            Chart1.ChartAreas(AreaNo - 1).AxisY.Minimum = Val(txtYAxisMin.Text)
            ChartInfo.dictAreaInfo(AreaName).AutoYAxisMinimum = False
        End If

        If txtYAxisMax.Text.Trim = "" Then
            chkYAxisAutoMax.Checked = True
            Chart1.ChartAreas(AreaNo - 1).AxisY.Maximum = Double.NaN 'Auto maximum.
            ChartInfo.dictAreaInfo(AreaName).AutoYAxisMaximum = True
        ElseIf chkYAxisAutoMax.Checked Then
            Chart1.ChartAreas(AreaNo - 1).AxisY.Maximum = Double.NaN 'Auto maximum.
            ChartInfo.dictAreaInfo(AreaName).AutoYAxisMaximum = True
        Else
            Chart1.ChartAreas(AreaNo - 1).AxisY.Maximum = Val(txtYAxisMax.Text)
            ChartInfo.dictAreaInfo(AreaName).AutoYAxisMaximum = False
        End If

        If txtYAxisAnnotInt.Text.Trim = "" Then
            chkYAxisAutoAnnotInt.Checked = True
            Chart1.ChartAreas(AreaNo - 1).AxisY.Interval = 0 'Zero indicates Auto mode.
            Chart1.ChartAreas(AreaNo - 1).AxisY.IntervalAutoMode = DataVisualization.Charting.IntervalAutoMode.VariableCount
        ElseIf chkYAxisAutoAnnotInt.Checked Then
            Chart1.ChartAreas(AreaNo - 1).AxisY.Interval = 0 'Zero indicates Auto mode.
            Chart1.ChartAreas(AreaNo - 1).AxisY.IntervalAutoMode = DataVisualization.Charting.IntervalAutoMode.VariableCount
        Else
            Chart1.ChartAreas(AreaNo - 1).AxisY.Interval = Val(txtYAxisAnnotInt.Text)
        End If

        If txtYAxisMajGridInt.Text.Trim = "" Then
            chkYAxisAutoMajGridInt.Checked = True
            Chart1.ChartAreas(AreaNo - 1).AxisY.MajorGrid.Interval = Double.NaN 'Indicates Not Set - use Axis Interval value.
            ChartInfo.dictAreaInfo(AreaName).AutoYAxisMajorGridInterval = True
        ElseIf chkYAxisAutoMajGridInt.Checked Then
            Chart1.ChartAreas(AreaNo - 1).AxisY.MajorGrid.Interval = Double.NaN 'Indicates Not Set - use Axis Interval value.
            ChartInfo.dictAreaInfo(AreaName).AutoYAxisMajorGridInterval = True
        Else
            Chart1.ChartAreas(AreaNo - 1).AxisY.MajorGrid.Interval = Val(txtYAxisMajGridInt.Text)
            ChartInfo.dictAreaInfo(AreaName).AutoYAxisMajorGridInterval = False
        End If




        'Y2 Axis: ------------------------------------------------------------------------
        Chart1.ChartAreas(AreaNo - 1).AxisY2.Title = txtY2AxisTitle.Text
        Chart1.ChartAreas(AreaNo - 1).AxisY2.TitleFont = txtY2AxisTitle.Font
        Chart1.ChartAreas(AreaNo - 1).AxisY2.TitleForeColor = txtY2AxisTitle.ForeColor
        If cmbY2AxisTitleAlignment.SelectedItem IsNot Nothing Then Chart1.ChartAreas(AreaNo - 1).AxisY2.TitleAlignment = [Enum].Parse(GetType(StringAlignment), cmbY2AxisTitleAlignment.SelectedItem.ToString)
        Chart1.ChartAreas(AreaNo - 1).AxisY2.LabelStyle.Format = txtY2AxisLabelStyleFormat.Text


        If txtY2AxisMin.Text.Trim = "" Then
            chkY2AxisAutoMin.Checked = True
            Chart1.ChartAreas(AreaNo - 1).AxisY2.Minimum = Double.NaN 'Auto minimum.
            ChartInfo.dictAreaInfo(AreaName).AutoY2AxisMinimum = True
        ElseIf chkY2AxisAutoMin.Checked Then
            Chart1.ChartAreas(AreaNo - 1).AxisY2.Minimum = Double.NaN 'Auto minimum.
            ChartInfo.dictAreaInfo(AreaName).AutoY2AxisMinimum = True
        Else
            Chart1.ChartAreas(AreaNo - 1).AxisY2.Minimum = Val(txtY2AxisMin.Text)
            ChartInfo.dictAreaInfo(AreaName).AutoY2AxisMinimum = False
        End If

        If txtY2AxisMax.Text.Trim = "" Then
            chkY2AxisAutoMax.Checked = True
            Chart1.ChartAreas(AreaNo - 1).AxisY2.Maximum = Double.NaN 'Auto maximum.
            ChartInfo.dictAreaInfo(AreaName).AutoY2AxisMaximum = True
        ElseIf chkY2AxisAutoMax.Checked Then
            Chart1.ChartAreas(AreaNo - 1).AxisY2.Maximum = Double.NaN 'Auto maximum.
            ChartInfo.dictAreaInfo(AreaName).AutoY2AxisMaximum = True
        Else
            Chart1.ChartAreas(AreaNo - 1).AxisY2.Maximum = Val(txtY2AxisMax.Text)
            ChartInfo.dictAreaInfo(AreaName).AutoY2AxisMaximum = False
        End If

        If txtY2AxisAnnotInt.Text.Trim = "" Then
            chkY2AxisAutoAnnotInt.Checked = True
            Chart1.ChartAreas(AreaNo - 1).AxisY2.Interval = 0 'Zero indicates Auto mode.
            Chart1.ChartAreas(AreaNo - 1).AxisY2.IntervalAutoMode = DataVisualization.Charting.IntervalAutoMode.VariableCount
        ElseIf chkY2AxisAutoAnnotInt.Checked Then
            Chart1.ChartAreas(AreaNo - 1).AxisY2.Interval = 0 'Zero indicates Auto mode.
            Chart1.ChartAreas(AreaNo - 1).AxisY2.IntervalAutoMode = DataVisualization.Charting.IntervalAutoMode.VariableCount
        Else
            Chart1.ChartAreas(AreaNo - 1).AxisY2.Interval = Val(txtY2AxisAnnotInt.Text)
        End If

        If txtY2AxisMajGridInt.Text.Trim = "" Then
            chkY2AxisAutoMajGridInt.Checked = True
            Chart1.ChartAreas(AreaNo - 1).AxisY2.MajorGrid.Interval = Double.NaN 'Indicates Not Set - use Axis Interval value.
            ChartInfo.dictAreaInfo(AreaName).AutoY2AxisMajorGridInterval = True
        ElseIf chkY2AxisAutoMajGridInt.Checked Then
            Chart1.ChartAreas(AreaNo - 1).AxisY2.MajorGrid.Interval = Double.NaN 'Indicates Not Set - use Axis Interval value.
            ChartInfo.dictAreaInfo(AreaName).AutoY2AxisMajorGridInterval = True
        Else
            Chart1.ChartAreas(AreaNo - 1).AxisY2.MajorGrid.Interval = Val(txtY2AxisMajGridInt.Text)
            ChartInfo.dictAreaInfo(AreaName).AutoY2AxisMajorGridInterval = False
        End If


    End Sub

    Private Sub btnAddArea_Click(sender As Object, e As EventArgs) Handles btnAddArea.Click
        'Add a new area to the Chart:

        If txtAreaRecordNo.Text = "0" Then
            'No Point Chart data hase been loaded.
            'Check if there is a default Area already in the Chart control:
            If Chart1.ChartAreas.Count > 0 Then
                txtAreaRecordNo.Text = "1"
                txtNAreaRecords.Text = Chart1.ChartAreas.Count
                txtAreaName.Text = Chart1.ChartAreas(0).Name
            Else
                'To Do!!!
            End If
        Else 'Add a new Area to the Chart control:
            'Dim NewAreaNo As Integer = Chart1.ChartAreas.Count
            Dim NewAreaNo As Integer = Chart1.ChartAreas.Count + 1

            'Dim NewAreaName As String = "Area" & NewAreaNo
            Dim NewAreaName As String = "ChartArea" & NewAreaNo
            Chart1.ChartAreas.Add(NewAreaName)

            txtAreaName.Text = NewAreaName
            txtAreaRecordNo.Text = NewAreaNo + 1
            txtNAreaRecords.Text = Chart1.ChartAreas.Count
        End If

    End Sub

    Private Sub btnApplyAreaSettings_Click(sender As Object, e As EventArgs) Handles btnApplyAreaSettings.Click
        ApplyAreasTabSettings()
    End Sub

    Private Sub btnXAxisTitleFont_Click(sender As Object, e As EventArgs) Handles btnXAxisTitleFont.Click
        FontDialog1.Font = txtXAxisTitle.Font
        FontDialog1.ShowDialog()
        txtXAxisTitle.Font = FontDialog1.Font
    End Sub
    Private Sub btnXAxisTitleColor_Click(sender As Object, e As EventArgs) Handles btnXAxisTitleColor.Click
        ColorDialog1.Color = txtXAxisTitle.ForeColor
        ColorDialog1.ShowDialog()
        txtXAxisTitle.ForeColor = ColorDialog1.Color
    End Sub

    'Private Sub txtXAxisZoomInterval_LostFocus(sender As Object, e As EventArgs)
    '    'Update ZoomTo value using the ZoomFrom and ZoomInterval values:

    '    txtXAxisZoomTo.Text = Val(txtXAxisZoomFrom.Text) + Val(txtXAxisZoomInterval.Text)
    'End Sub

    'Private Sub btnZoomOK_Click(sender As Object, e As EventArgs) Handles btnZoomOK.Click
    '    'Zoom to the X Axis range shown:

    '    Dim ZoomStart As Double
    '    Dim ZoomEnd As Double

    '    If txtXAxisZoomFrom.Text = "" Then

    '    Else
    '        ZoomStart = Val(txtXAxisZoomFrom.Text)
    '        If txtXAxisZoomTo.Text = "" Then

    '        Else
    '            ZoomEnd = Val(txtXAxisZoomTo.Text)
    '            Dim AreaNo As Integer = Val(txtAreaRecordNo.Text)
    '            Dim AreaCount As Integer = Chart1.ChartAreas.Count

    '            If AreaNo - 1 > AreaCount Then
    '                Message.AddWarning("The area number is larger than to number of areas in the chart!" & vbCrLf)
    '                Exit Sub
    '            End If

    '            Dim AreaName As String = txtAreaName.Text.Trim
    '            If chkXAxisScrollBar.Checked Then
    '                Chart1.ChartAreas(AreaNo - 1).AxisX.ScrollBar.Enabled = True
    '                Chart1.ChartAreas(AreaNo - 1).AxisX.ScrollBar.Size = 16
    '            End If
    '            Chart1.ChartAreas(AreaNo - 1).AxisX.ScaleView.Zoom(ZoomStart, ZoomEnd)
    '        End If
    '    End If

    'End Sub

    Private Sub btnX2AxisTitleColor_Click(sender As Object, e As EventArgs) Handles btnX2AxisTitleColor.Click
        ColorDialog1.Color = txtX2AxisTitle.ForeColor
        ColorDialog1.ShowDialog()
        txtX2AxisTitle.ForeColor = ColorDialog1.Color
    End Sub

    Private Sub btnYAxisTitleColor_Click(sender As Object, e As EventArgs) Handles btnYAxisTitleColor.Click
        ColorDialog1.Color = txtYAxisTitle.ForeColor
        ColorDialog1.ShowDialog()
        txtYAxisTitle.ForeColor = ColorDialog1.Color
    End Sub

    Private Sub btnY2AxisTitleColor_Click(sender As Object, e As EventArgs) Handles btnY2AxisTitleColor.Click
        ColorDialog1.Color = txtY2AxisTitle.ForeColor
        ColorDialog1.ShowDialog()
        txtY2AxisTitle.ForeColor = ColorDialog1.Color
    End Sub

    Private Sub btnX2AxisTitleFont_Click(sender As Object, e As EventArgs) Handles btnX2AxisTitleFont.Click
        FontDialog1.Font = txtX2AxisTitle.Font
        FontDialog1.ShowDialog()
        txtX2AxisTitle.Font = FontDialog1.Font
    End Sub

    Private Sub btnYAxisTitleFont_Click(sender As Object, e As EventArgs) Handles btnYAxisTitleFont.Click
        FontDialog1.Font = txtYAxisTitle.Font
        FontDialog1.ShowDialog()
        txtYAxisTitle.Font = FontDialog1.Font
    End Sub

    Private Sub btnY2AxisTitleFont_Click(sender As Object, e As EventArgs) Handles btnY2AxisTitleFont.Click
        FontDialog1.Font = txtY2AxisTitle.Font
        FontDialog1.ShowDialog()
        txtY2AxisTitle.Font = FontDialog1.Font
    End Sub

    Private Sub btnDeleteArea_Click(sender As Object, e As EventArgs) Handles btnDeleteArea.Click
        'Delete the selected chart area:

        Dim AreaNo = Val(txtAreaRecordNo.Text) - 1 'Zero-based area number.

        If Chart1.ChartAreas.Count > 1 Then
            Chart1.ChartAreas.RemoveAt(AreaNo)
            UpdateAreasTabSettings()
        Else
            'Only one chart area. This should be edited rather than deleted.
        End If
    End Sub

    Private Sub btnPrevArea_Click(sender As Object, e As EventArgs) Handles btnPrevArea.Click
        'Show the previous Area:

        Dim AreaNo = Val(txtAreaRecordNo.Text) - 1 'Zero-based area number.

        If AreaNo = 0 Then
            'Already at the first Area.
        Else
            'Show the previous area:
            AreaNo = AreaNo - 1
            'txtAreaRecordNo.Text = AreaNo + 1
            ShowArea(AreaNo)
        End If

    End Sub

    Private Sub btnNextArea_Click(sender As Object, e As EventArgs) Handles btnNextArea.Click
        'Show the next area:

        Dim AreaNo = Val(txtAreaRecordNo.Text) - 1 'Zero-based area number.

        If AreaNo + 1 >= Chart1.ChartAreas.Count Then
            'Already at the last area.
        Else
            'Snow the next area:
            AreaNo = AreaNo + 1
            ShowArea(AreaNo)
        End If

    End Sub

    Private Sub ShowArea(ByVal AreaNo As Integer)
        'Show the Area for AreaNo (Zero-based index).

        If AreaNo < 0 Then
            'AreaNo is too small!
        ElseIf AreaNo + 1 > Chart1.ChartAreas.Count Then
            'AreaNo is too large!
        Else
            txtAreaRecordNo.Text = AreaNo + 1
            Dim AreaName As String = Chart1.ChartAreas(AreaNo).Name
            txtAreaName.Text = AreaName

            txtXAxisTitle.Text = Chart1.ChartAreas(AreaNo).AxisX.Title
            txtX2AxisTitle.Text = Chart1.ChartAreas(AreaNo).AxisX2.Title
            txtYAxisTitle.Text = Chart1.ChartAreas(AreaNo).AxisY.Title
            txtY2AxisTitle.Text = Chart1.ChartAreas(AreaNo).AxisY2.Title

            txtXAxisTitle.Font = Chart1.ChartAreas(AreaNo).AxisX.TitleFont
            txtX2AxisTitle.Font = Chart1.ChartAreas(AreaNo).AxisX2.TitleFont
            txtYAxisTitle.Font = Chart1.ChartAreas(AreaNo).AxisY.TitleFont
            txtY2AxisTitle.Font = Chart1.ChartAreas(AreaNo).AxisY2.TitleFont

            txtXAxisTitle.ForeColor = Chart1.ChartAreas(AreaNo).AxisX.TitleForeColor
            txtX2AxisTitle.ForeColor = Chart1.ChartAreas(AreaNo).AxisX2.TitleForeColor
            txtYAxisTitle.ForeColor = Chart1.ChartAreas(AreaNo).AxisY.TitleForeColor
            txtY2AxisTitle.ForeColor = Chart1.ChartAreas(AreaNo).AxisY2.TitleForeColor

            cmbXAxisTitleAlignment.SelectedIndex = cmbXAxisTitleAlignment.FindStringExact(Chart1.ChartAreas(AreaNo).AxisX.TitleAlignment.ToString)
            cmbX2AxisTitleAlignment.SelectedIndex = cmbX2AxisTitleAlignment.FindStringExact(Chart1.ChartAreas(AreaNo).AxisX2.TitleAlignment.ToString)
            cmbYAxisTitleAlignment.SelectedIndex = cmbYAxisTitleAlignment.FindStringExact(Chart1.ChartAreas(AreaNo).AxisY.TitleAlignment.ToString)
            cmbY2AxisTitleAlignment.SelectedIndex = cmbY2AxisTitleAlignment.FindStringExact(Chart1.ChartAreas(AreaNo).AxisY2.TitleAlignment.ToString)

            'Axis Minimum values: --------------------------------------------
            If ChartInfo.dictAreaInfo.ContainsKey(AreaName) Then
                chkXAxisAutoMin.Checked = ChartInfo.dictAreaInfo(AreaName).AutoXAxisMinimum
                txtXAxisMin.Text = Chart1.ChartAreas(AreaNo).AxisX.Minimum
                'txtXAxisZoomFrom.Text = Chart1.ChartAreas(AreaNo).AxisX.Minimum
                chkX2AxisAutoMin.Checked = ChartInfo.dictAreaInfo(AreaName).AutoX2AxisMinimum
                txtX2AxisMin.Text = Chart1.ChartAreas(AreaNo).AxisX2.Minimum
                chkYAxisAutoMin.Checked = ChartInfo.dictAreaInfo(AreaName).AutoYAxisMinimum
                txtYAxisMin.Text = Chart1.ChartAreas(AreaNo).AxisY.Minimum
                chkY2AxisAutoMin.Checked = ChartInfo.dictAreaInfo(AreaName).AutoY2AxisMinimum
                txtY2AxisMin.Text = Chart1.ChartAreas(AreaNo).AxisY2.Minimum
            End If


            'Axis Maximum values: -----------------------------------------
            If ChartInfo.dictAreaInfo.ContainsKey(AreaName) Then
                chkXAxisAutoMax.Checked = ChartInfo.dictAreaInfo(AreaName).AutoXAxisMaximum
                txtXAxisMax.Text = Chart1.ChartAreas(AreaNo).AxisX.Maximum
                'txtXAxisZoomTo.Text = Chart1.ChartAreas(AreaNo).AxisX.Maximum
                'txtXAxisZoomInterval.Text = Chart1.ChartAreas(AreaNo).AxisX.Maximum - Chart1.ChartAreas(AreaNo).AxisX.Minimum
                chkX2AxisAutoMax.Checked = ChartInfo.dictAreaInfo(AreaName).AutoX2AxisMaximum
                txtX2AxisMax.Text = Chart1.ChartAreas(AreaNo).AxisX2.Maximum
                chkYAxisAutoMax.Checked = ChartInfo.dictAreaInfo(AreaName).AutoYAxisMaximum
                txtYAxisMax.Text = Chart1.ChartAreas(AreaNo).AxisY.Maximum
                chkY2AxisAutoMax.Checked = ChartInfo.dictAreaInfo(AreaName).AutoY2AxisMaximum
                txtY2AxisMax.Text = Chart1.ChartAreas(AreaNo).AxisY2.Maximum
            End If

            'Axis Intervals: -------------------------------------------------------------
            If Chart1.ChartAreas(AreaNo).AxisX.Interval = 0 Then 'Auto mode.
                chkXAxisAutoAnnotInt.Checked = True
            Else
                chkXAxisAutoAnnotInt.Checked = False
            End If
            txtXAxisAnnotInt.Text = Chart1.ChartAreas(AreaNo).AxisX.Interval

            If Chart1.ChartAreas(AreaNo).AxisX2.Interval = 0 Then 'Auto mode.
                chkX2AxisAutoAnnotInt.Checked = True
            Else
                chkX2AxisAutoAnnotInt.Checked = False
            End If
            txtX2AxisAnnotInt.Text = Chart1.ChartAreas(AreaNo).AxisX2.Interval
            If Chart1.ChartAreas(AreaNo).AxisY.Interval = 0 Then 'Auto mode.
                chkYAxisAutoAnnotInt.Checked = True
            Else
                chkYAxisAutoAnnotInt.Checked = False
            End If
            txtYAxisAnnotInt.Text = Chart1.ChartAreas(AreaNo).AxisY.Interval
            If Chart1.ChartAreas(AreaNo).AxisY2.Interval = 0 Then 'Auto mode.
                chkY2AxisAutoAnnotInt.Checked = True
            Else
                chkY2AxisAutoAnnotInt.Checked = False
            End If
            txtY2AxisAnnotInt.Text = Chart1.ChartAreas(AreaNo).AxisY2.Interval

            'Axis Major Grid Intervals: -----------------------------------------------------------
            If ChartInfo.dictAreaInfo.ContainsKey(AreaName) Then
                chkXAxisAutoMajGridInt.Checked = ChartInfo.dictAreaInfo(AreaName).AutoXAxisMajorGridInterval
                txtXAxisMajGridInt.Text = Chart1.ChartAreas(AreaNo).AxisX.MajorGrid.Interval
                chkX2AxisAutoMajGridInt.Checked = ChartInfo.dictAreaInfo(AreaName).AutoX2AxisMajorGridInterval
                txtX2AxisMajGridInt.Text = Chart1.ChartAreas(AreaNo).AxisX2.MajorGrid.Interval
                chkYAxisAutoMajGridInt.Checked = ChartInfo.dictAreaInfo(AreaName).AutoYAxisMajorGridInterval
                txtYAxisMajGridInt.Text = Chart1.ChartAreas(AreaNo).AxisY.MajorGrid.Interval
                chkY2AxisAutoMajGridInt.Checked = ChartInfo.dictAreaInfo(AreaName).AutoY2AxisMajorGridInterval
                txtY2AxisMajGridInt.Text = Chart1.ChartAreas(AreaNo).AxisY2.MajorGrid.Interval
            End If

            'Axis Label Style Format:
            txtXAxisLabelStyleFormat.Text = Chart1.ChartAreas(AreaNo).AxisX.LabelStyle.Format
            txtX2AxisLabelStyleFormat.Text = Chart1.ChartAreas(AreaNo).AxisX2.LabelStyle.Format
            txtYAxisLabelStyleFormat.Text = Chart1.ChartAreas(AreaNo).AxisY.LabelStyle.Format
            txtY2AxisLabelStyleFormat.Text = Chart1.ChartAreas(AreaNo).AxisY2.LabelStyle.Format

            'Update list of areas in Series tab:
            cmbChartArea.Items.Clear()
            For Each item In Chart1.ChartAreas
                cmbChartArea.Items.Add(item.Name)
                'Message.Add("Adding Chare Area: " & item.Name & vbCrLf)
            Next
            Dim SeriesName As String = txtSeriesName.Text
            cmbChartArea.SelectedItem = cmbChartArea.FindStringExact(Chart1.Series(SeriesName).ChartArea)
        End If
    End Sub

#End Region 'Chart Areas Tab --------------------------------------------------------------------------------------------------------------------------------------------






    'Private Sub UpdateChartQuery()
    '    'Update the Chart Query with the current InputQuery 
    '    'PointChart.InputQuery = PointChart.InputQuery
    '    PointChart.InputQuery = txtInputQuery.Text
    'End Sub

    Private Sub bgwComCheck_DoWork(sender As Object, e As DoWorkEventArgs) Handles bgwComCheck.DoWork
        'The communications check thread.
        While ConnectedToComNet
            Try
                If client.IsAlive() Then
                    bgwComCheck.ReportProgress(1, Format(Now, "HH:mm:ss") & " Connection OK." & vbCrLf)
                Else
                    bgwComCheck.ReportProgress(1, Format(Now, "HH:mm:ss") & " Connection Fault.")
                End If
            Catch ex As Exception
                bgwComCheck.ReportProgress(1, "Error in bgeComCheck_DoWork!" & vbCrLf)
                bgwComCheck.ReportProgress(1, ex.Message & vbCrLf)
            End Try

            'System.Threading.Thread.Sleep(60000) 'Sleep time in milliseconds (60 seconds) - For testing only.
            'System.Threading.Thread.Sleep(3600000) 'Sleep time in milliseconds (60 minutes)
            System.Threading.Thread.Sleep(1800000) 'Sleep time in milliseconds (30 minutes)
        End While
    End Sub

    Private Sub bgwComCheck_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles bgwComCheck.ProgressChanged
        Message.Add(e.UserState.ToString) 'Show the ComCheck message 
    End Sub


    'Private Sub SetUpPointChartForm()
    '    'Set up the settings forms for a Point Chart

    '    'Set up Point chart:
    '    'ChartType = [Enum].Parse(GetType(DataVisualization.Charting.SeriesChartType), "Point")
    '    txtChartDescr.Text = "Point Chart" & vbCrLf
    '    txtChartDescr.Text = txtChartDescr.Text & "The Point chart type uses value points to represent its data." & vbCrLf
    '    txtChartDescr.Text = txtChartDescr.Text & "Custom Attributes:" & vbCrLf
    '    txtChartDescr.Text = txtChartDescr.Text & "EmptyPointValue - Specifies the value to be used for empty points. This property determines how an empty point is treated when the chart is drawn. (Average, Zero)" & vbCrLf
    '    txtChartDescr.Text = txtChartDescr.Text & "LabelStyle - Specifies the label position of the data point. (Auto, Top, Bottom, Right, Left, TopLeft, TopRight, BottomLeft, BottomRight, Center)" & vbCrLf
    '    txtChartDescr.Text = txtChartDescr.Text & "PixelPointDepth - Specifies the 3D series depth in pixels. (Any integer > 0)" & vbCrLf
    '    txtChartDescr.Text = txtChartDescr.Text & "PixelPointGapDepth - Specifies the 3D gap depth in pixels. (Any integer > 0)" & vbCrLf

    '    'Y Values:
    '    DataGridView1.Rows.Clear()
    '    DataGridView1.Rows.Add(1)
    '    DataGridView1.Rows(0).Cells(0).Value = "Yvalue" 'Y Value Parameter Name

    '    'Custom Attributes:
    '    DataGridView2.Rows.Clear()
    '    DataGridView2.Rows.Add(4)
    '    DataGridView2.Rows(0).Cells(0).Value = "EmptyPointValue"
    '    DataGridView2.Rows(0).Cells(1).Value = "Average, Zero"
    '    Dim cbc0 As New DataGridViewComboBoxCell
    '    cbc0.Items.Add(" ")
    '    cbc0.Items.Add("Average")
    '    cbc0.Items.Add("Zero")
    '    DataGridView2.Rows(0).Cells(2) = cbc0
    '    DataGridView2.Rows(1).Cells(0).Value = "LabelStyle"
    '    DataGridView2.Rows(1).Cells(1).Value = "Auto, Top, Bottom, Right, Left, TopLeft, TopRight, BottomLeft, BottomRight, Center"
    '    Dim cbc1 As New DataGridViewComboBoxCell
    '    cbc1.Items.Add(" ")
    '    cbc1.Items.Add("Auto")
    '    cbc1.Items.Add("Top")
    '    cbc1.Items.Add("Bottom")
    '    cbc1.Items.Add("Right")
    '    cbc1.Items.Add("Left")
    '    cbc1.Items.Add("TopLeft")
    '    cbc1.Items.Add("TopRight")
    '    cbc1.Items.Add("BottomLeft")
    '    cbc1.Items.Add("BottomRight")
    '    cbc1.Items.Add("Center")
    '    DataGridView2.Rows(1).Cells(2) = cbc1
    '    DataGridView2.Rows(2).Cells(0).Value = "PixelPointDepth"
    '    DataGridView2.Rows(2).Cells(1).Value = "Any integer > 0"
    '    DataGridView2.Rows(3).Cells(0).Value = "PixelPointGapDepth"
    '    DataGridView2.Rows(3).Cells(1).Value = "Any integer > 0"

    '    'UpdateCurrentSettings()
    '    'UpdateChartForm()
    '    UpdatePointChartForm()
    'End Sub

    'NOTE: REPLACE THIS WITH UpdateSeriesTabSettings() !!!
    'Private Sub UpdatePointChartForm()
    '    'Update the Chart Type, Titles, X Axis and Y Axis tabs with the settings stored in PointChart.

    '    ''Apply Input Data settings:
    '    'InputDataType = PointChart.InputDataType
    '    'InputDatabasePath = PointChart.InputDatabasePath
    '    'InputQuery = PointChart.InputQuery
    '    'InputDataDescr = PointChart.InputDataDescr

    '    txtSeriesName.Text = PointChart.SeriesName

    '    'Apply X Values Field Name:
    '    cmbXValues.SelectedIndex = cmbXValues.FindStringExact(PointChart.XValuesFieldName)

    '    'Apply Y Values selections:
    '    For I = 1 To cboFieldSelections.Items.Count
    '        If PointChart.YValuesFieldName = cboFieldSelections.Items(I - 1) Then
    '            DataGridView1.Rows(0).Cells(1).Value = cboFieldSelections.Items(I - 1)
    '        End If
    '    Next

    '    'Apply Custom Attributes selections:
    '    'LabelValueType (High, Low, Open, Close)
    '    If PointChart.EmptyPointValue <> "" Then
    '        DataGridView2.Rows(0).Cells(2).Value = PointChart.EmptyPointValue
    '    End If
    '    If PointChart.LabelStyle <> "" Then
    '        DataGridView2.Rows(1).Cells(2).Value = PointChart.LabelStyle
    '    End If
    '    'PixelPointDepth (Any integer > 0)
    '    DataGridView2.Rows(2).Cells(2).Value = PointChart.PixelPointDepth
    '    'PixelPointGapDepth (Any integer > 0)
    '    DataGridView2.Rows(3).Cells(2).Value = PointChart.PixelPointGapDepth

    '    'Update the ChartLabel settings: -------------------------------------------------------------------------
    '    txtChartTitle.Text = PointChart.ChartLabel.Text
    '    'txtChartTitle.ForeColor = Color.FromName(PointChart.ChartLabel.Color)
    '    txtChartTitle.ForeColor = PointChart.ChartLabel.Color

    '    txtChartTitle.Text = PointChart.ChartLabel.Text
    '    'txtChartTitle.ForeColor = Color.FromName(PointChart.ChartLabel.Color)
    '    txtChartTitle.ForeColor = PointChart.ChartLabel.Color
    '    Dim myFontStyle As FontStyle = FontStyle.Regular
    '    If PointChart.ChartLabel.Bold Then
    '        myFontStyle = myFontStyle Or FontStyle.Bold
    '    End If
    '    If PointChart.ChartLabel.Italic Then
    '        myFontStyle = myFontStyle Or FontStyle.Italic
    '    End If
    '    If PointChart.ChartLabel.Strikeout Then
    '        myFontStyle = myFontStyle Or FontStyle.Strikeout
    '    End If
    '    If PointChart.ChartLabel.Underline Then
    '        myFontStyle = myFontStyle Or FontStyle.Underline
    '    End If

    '    txtChartTitle.Font = New Font("Arial", PointChart.ChartLabel.Size, myFontStyle)

    '    'Update the XAxis settings: -------------------------------------------------------------------------
    '    txtXAxisTitle.Text = PointChart.XAxis.Title.Text
    '    txtXAxisTitle.ForeColor = Color.FromName(PointChart.XAxis.Title.Color)

    '    myFontStyle = FontStyle.Regular
    '    If PointChart.XAxis.Title.Bold Then
    '        myFontStyle = myFontStyle Or FontStyle.Bold
    '    End If
    '    If PointChart.XAxis.Title.Italic Then
    '        myFontStyle = myFontStyle Or FontStyle.Italic
    '    End If
    '    If PointChart.XAxis.Title.Strikeout Then
    '        myFontStyle = myFontStyle Or FontStyle.Strikeout
    '    End If
    '    If PointChart.XAxis.Title.Underline Then
    '        myFontStyle = myFontStyle Or FontStyle.Underline
    '    End If

    '    txtXAxisTitle.Font = New Font(PointChart.XAxis.Title.FontName, PointChart.XAxis.Title.Size, myFontStyle)

    '    chkXAxisAutoMin.Checked = PointChart.XAxis.AutoMinimum
    '    chkXAxisAutoMax.Checked = PointChart.XAxis.AutoMaximum

    '    txtXAxisMin.Text = PointChart.XAxis.Minimum
    '    txtXAxisMax.Text = PointChart.XAxis.Maximum

    '    chkXAxisAutoAnnotInt.Checked = PointChart.XAxis.AutoInterval
    '    chkXAxisAutoMajGridInt.Checked = PointChart.XAxis.AutoMajorGridInterval

    '    txtXAxisAnnotInt.Text = PointChart.XAxis.Interval
    '    txtXAxisMajGridInt.Text = PointChart.XAxis.MajorGridInterval



    '    'Update the YAxis settings: -----------------------------------------------------------------------------
    '    txtYAxisTitle.Text = PointChart.YAxis.Title.Text
    '    txtYAxisTitle.ForeColor = Color.FromName(PointChart.YAxis.Title.Color)

    '    myFontStyle = FontStyle.Regular
    '    If PointChart.YAxis.Title.Bold Then
    '        myFontStyle = myFontStyle Or FontStyle.Bold
    '    End If
    '    If PointChart.YAxis.Title.Italic Then
    '        myFontStyle = myFontStyle Or FontStyle.Italic
    '    End If
    '    If PointChart.YAxis.Title.Strikeout Then
    '        myFontStyle = myFontStyle Or FontStyle.Strikeout
    '    End If
    '    If PointChart.YAxis.Title.Underline Then
    '        myFontStyle = myFontStyle Or FontStyle.Underline
    '    End If

    '    txtYAxisTitle.Font = New Font(PointChart.YAxis.Title.FontName, PointChart.YAxis.Title.Size, myFontStyle)

    '    chkYAxisAutoMin.Checked = PointChart.YAxis.AutoMinimum
    '    chkYAxisAutoMax.Checked = PointChart.YAxis.AutoMaximum

    '    txtYAxisMin.Text = PointChart.YAxis.Minimum
    '    txtYAxisMax.Text = PointChart.YAxis.Maximum

    '    chkYAxisAutoAnnotInt.Checked = PointChart.YAxis.AutoInterval
    '    chkYAxisAutoMajGridInt.Checked = PointChart.YAxis.AutoMajorGridInterval

    '    txtYAxisAnnotInt.Text = PointChart.YAxis.Interval
    '    txtYAxisMajGridInt.Text = PointChart.YAxis.MajorGridInterval

    '    'Update chart File Name:
    '    txtChartFileName.Text = PointChart.FileName

    'End Sub

    Private Sub btnDrawChart_Click(sender As Object, e As EventArgs) Handles btnDrawChart.Click
        DrawPointChart()
    End Sub

    Private Sub DrawPointChart()
        'Draw the Point Chart:
        Try
            Dim SeriesName As String
            Dim ChartArea As String
            For Each item In Chart1.Series
                SeriesName = item.Name
                ChartArea = Chart1.Series(SeriesName).ChartArea
                If ChartInfo.dictAreaInfo(ChartArea).AutoXAxisMinimum Then Chart1.ChartAreas(ChartArea).AxisX.Minimum = Double.NaN
                If ChartInfo.dictAreaInfo(ChartArea).AutoXAxisMaximum Then Chart1.ChartAreas(ChartArea).AxisX.Maximum = Double.NaN
                If ChartInfo.dictAreaInfo(ChartArea).AutoXAxisMajorGridInterval Then Chart1.ChartAreas(ChartArea).AxisX.MajorGrid.Interval = Double.NaN
                Chart1.ChartAreas(ChartArea).AxisX.IntervalAutoMode = DataVisualization.Charting.IntervalAutoMode.VariableCount
                Chart1.ChartAreas(ChartArea).AxisX.RoundAxisValues()
                Chart1.ChartAreas(ChartArea).AxisY.RoundAxisValues()
                If item.ChartType = DataVisualization.Charting.SeriesChartType.Point Then
                    Chart1.Series(SeriesName).Points.DataBindXY(ChartInfo.ds.Tables(0).DefaultView, ChartInfo.dictSeriesInfo(SeriesName).XValuesFieldName, ChartInfo.ds.Tables(0).DefaultView, ChartInfo.dictSeriesInfo(SeriesName).YValuesFieldName)
                End If
            Next
        Catch ex As Exception
            Message.AddWarning(ex.Message & vbCrLf)
        End Try
    End Sub

    'Private Sub DrawPointChart()
    '    'Draw the Point Chart:

    '    Try

    '        Chart1.Series.Clear()
    '        'Chart1.Series.Add("Series1")
    '        Chart1.Series.Add(PointChart.SeriesName)
    '        Chart1.Series(PointChart.SeriesName).YValuesPerPoint = 1
    '        Chart1.Series(PointChart.SeriesName).Points.DataBindXY(ds.Tables(0).DefaultView, PointChart.XValuesFieldName, ds.Tables(0).DefaultView, PointChart.YValuesFieldName)
    '        Chart1.Series(PointChart.SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.Point

    '        'Chart1.Series(PointChart.SeriesName).Points(200).Label = "Test"
    '        'Chart1.Series(PointChart.SeriesName).Points(200).MarkerColor = Color.PaleGreen
    '        'Chart1.Series(PointChart.SeriesName).Points(210).Label = "Test"
    '        'Chart1.Series(PointChart.SeriesName).Points(210).MarkerColor = Color.Azure
    '        'Chart1.Series(PointChart.SeriesName).Points(220).Label = "Test"
    '        'Chart1.Series(PointChart.SeriesName).Points(220).MarkerColor = Color.Bisque
    '        'Chart1.Series(PointChart.SeriesName).Points(230).Label = "Test"
    '        'Chart1.Series(PointChart.SeriesName).Points(230).MarkerColor = Color.Black
    '        'Chart1.Series(PointChart.SeriesName).Points(240).Label = "Test"
    '        'Chart1.Series(PointChart.SeriesName).Points(240).MarkerColor = Color.BlueViolet
    '        'Chart1.Series(PointChart.SeriesName).Points(250).Label = "Test"
    '        'Chart1.Series(PointChart.SeriesName).Points(250).MarkerColor = Color.CadetBlue

    '        ''Chart1.Series(PointChart.SeriesName).Points(300).Label = 

    '        'For Each item In Chart1.Series(PointChart.SeriesName).Points()

    '        'Next

    '        ''NOTE: This code was used to color code the points according to the "Total_Profit_pct" data field in the Greenblatt analysis table.
    '        'Dim NRows As Integer = ds.Tables(0).Rows.Count
    '        'Dim I As Integer
    '        'For I = 0 To NRows - 1
    '        '    'If ds.Tables(0).Rows(I).Item("Total_Profit_pct") = DBNull Then
    '        '    If IsDBNull(ds.Tables(0).Rows(I).Item("Total_Profit_pct")) Then
    '        '    Else
    '        '        'If ds.Tables(0).Rows(I).Item("Total_Profit_pct") > 0 Then
    '        '        '    Chart1.Series(PointChart.SeriesName).Points(I).MarkerColor = Color.Black
    '        '        'Else
    '        '        '    Chart1.Series(PointChart.SeriesName).Points(I).MarkerColor = Color.Red
    '        '        'End If
    '        '        If ds.Tables(0).Rows(I).Item("Total_Profit_pct") > 100 Then
    '        '            Chart1.Series(PointChart.SeriesName).Points(I).MarkerColor = Color.Black
    '        '        ElseIf ds.Tables(0).Rows(I).Item("Total_Profit_pct") > 50 Then
    '        '            Chart1.Series(PointChart.SeriesName).Points(I).MarkerColor = Color.DarkGray
    '        '        ElseIf ds.Tables(0).Rows(I).Item("Total_Profit_pct") > 10 Then
    '        '            Chart1.Series(PointChart.SeriesName).Points(I).MarkerColor = Color.Gray
    '        '        ElseIf ds.Tables(0).Rows(I).Item("Total_Profit_pct") > 0 Then
    '        '            Chart1.Series(PointChart.SeriesName).Points(I).MarkerColor = Color.LightGray
    '        '        ElseIf ds.Tables(0).Rows(I).Item("Total_Profit_pct") > -10 Then
    '        '            Chart1.Series(PointChart.SeriesName).Points(I).MarkerColor = Color.LightPink
    '        '        ElseIf ds.Tables(0).Rows(I).Item("Total_Profit_pct") > -50 Then
    '        '            Chart1.Series(PointChart.SeriesName).Points(I).MarkerColor = Color.Pink
    '        '        Else
    '        '            Chart1.Series(PointChart.SeriesName).Points(I).MarkerColor = Color.Red
    '        '        End If
    '        '    End If
    '        'Next

    '        If PointChart.EmptyPointValue <> "" Then Chart1.Series(PointChart.SeriesName).SetCustomProperty("EmptyPointValue", PointChart.EmptyPointValue)
    '        If PointChart.LabelStyle <> "" Then Chart1.Series(PointChart.SeriesName).SetCustomProperty("LabelStyle", PointChart.LabelStyle)
    '        Chart1.Series(PointChart.SeriesName).SetCustomProperty("PixelPointDepth", PointChart.PixelPointDepth)
    '        Chart1.Series(PointChart.SeriesName).SetCustomProperty("PixelPointGapDepth", PointChart.PixelPointGapDepth)

    '        'Specify Y Axis range: -------------------------------------------------------------------------------
    '        If PointChart.YAxis.AutoMinimum = True Then
    '            Chart1.ChartAreas(0).AxisY.Minimum = [Double].NaN
    '        Else
    '            Chart1.ChartAreas(0).AxisY.Minimum = PointChart.YAxis.Minimum
    '        End If
    '        If PointChart.YAxis.AutoMaximum = True Then
    '            Chart1.ChartAreas(0).AxisY.Maximum = [Double].NaN
    '        Else
    '            Chart1.ChartAreas(0).AxisY.Maximum = PointChart.YAxis.Maximum
    '        End If

    '        'Chart1.ChartAreas(0).AxisY.MajorGrid.Interval = 5 'Set the grid interval 
    '        'Chart1.ChartAreas(0).AxisY.Interval = 5 'Set the annotation interval

    '        'Chart1.ChartAreas(0).AxisY.MajorGrid.Interval = [Double].NaN 'Default?
    '        'Chart1.ChartAreas(0).AxisY.Interval = [Double].NaN 'Default?

    '        'Chart1.ChartAreas(0).AxisY.MajorGrid.Interval = 0 'Default
    '        'Chart1.ChartAreas(0).AxisY.Interval = 0 'Default

    '        'Specify Y Axis annotation and major grid intervals: -----------------------------------------------------
    '        If PointChart.YAxis.AutoInterval = True Then
    '            Chart1.ChartAreas(0).AxisY.Interval = 0
    '        Else
    '            Chart1.ChartAreas(0).AxisY.Interval = PointChart.YAxis.Interval
    '        End If

    '        If PointChart.YAxis.AutoMajorGridInterval = True Then
    '            Chart1.ChartAreas(0).AxisY.MajorGrid.Interval = 0
    '            Message.Add("Y Axis major grid interval is automatic." & vbCrLf)
    '        Else
    '            Chart1.ChartAreas(0).AxisY.MajorGrid.Interval = PointChart.YAxis.MajorGridInterval
    '        End If

    '        Chart1.ChartAreas(0).AxisY.LineWidth = 2 'Increase the line width of the Y Axis.
    '        'Chart1.ChartAreas(0).AxisY(0).LineWidth = 2
    '        'Chart1.ChartAreas(0).AxisY("0").Line = 2
    '        'Chart1.ChartAreas(0).AxisY.CustomLabels.
    '        Chart1.ChartAreas(0).AxisY.Crossing = 0

    '        'Add a bold origin axis line:
    '        'Dim am1 As New Charting.


    '        'Specify X Axis range: ------------------------------------------------------------------------------
    '        If PointChart.XAxis.AutoMinimum = True Then
    '            Chart1.ChartAreas(0).AxisX.Minimum = [Double].NaN
    '        Else
    '            Chart1.ChartAreas(0).AxisX.Minimum = PointChart.XAxis.Minimum
    '        End If
    '        If PointChart.XAxis.AutoMaximum = True Then
    '            Chart1.ChartAreas(0).AxisX.Maximum = [Double].NaN
    '        Else
    '            Chart1.ChartAreas(0).AxisX.Maximum = PointChart.XAxis.Maximum
    '        End If

    '        'Specify X Axis annotation and major grid intervals: -----------------------------------------------------
    '        Chart1.ChartAreas(0).AxisX.IntervalType = Charting.DateTimeIntervalType.Auto

    '        If PointChart.XAxis.AutoInterval = True Then
    '            Chart1.ChartAreas(0).AxisX.Interval = 0
    '        Else
    '            Chart1.ChartAreas(0).AxisX.Interval = PointChart.XAxis.Interval
    '        End If

    '        If PointChart.XAxis.AutoMajorGridInterval = True Then
    '            Chart1.ChartAreas(0).AxisX.MajorGrid.Interval = 0
    '            Message.Add("X Axis major grid interval is automatic." & vbCrLf)
    '        Else
    '            Chart1.ChartAreas(0).AxisX.MajorGrid.Interval = PointChart.XAxis.MajorGridInterval
    '        End If

    '        Chart1.ChartAreas(0).AxisX.LineWidth = 2 'Increase the line width of the X Axis.
    '        Chart1.ChartAreas(0).AxisX.Crossing = 0

    '        'Specify X Axis label: ------------------------------------------------------------------------------------
    '        Chart1.ChartAreas(0).AxisX.TitleAlignment = PointChart.XAxis.TitleAlignment

    '        Dim myFontStyle As FontStyle = FontStyle.Regular
    '        If PointChart.XAxis.Title.Bold Then
    '            myFontStyle = myFontStyle Or FontStyle.Bold
    '        End If
    '        If PointChart.XAxis.Title.Italic Then
    '            myFontStyle = myFontStyle Or FontStyle.Italic
    '        End If
    '        If PointChart.XAxis.Title.Strikeout Then
    '            myFontStyle = myFontStyle Or FontStyle.Strikeout
    '        End If
    '        If PointChart.XAxis.Title.Underline Then
    '            myFontStyle = myFontStyle Or FontStyle.Underline
    '        End If

    '        Chart1.ChartAreas(0).AxisX.TitleFont = New Font("Arial", PointChart.XAxis.Title.Size, myFontStyle)
    '        Chart1.ChartAreas(0).AxisX.Title = PointChart.XAxis.Title.Text

    '        'Specify Y Axis label: ------------------------------------------------------------------------------------
    '        Chart1.ChartAreas(0).AxisY.TitleAlignment = PointChart.YAxis.TitleAlignment
    '        myFontStyle = FontStyle.Regular
    '        If PointChart.YAxis.Title.Bold Then
    '            myFontStyle = myFontStyle Or FontStyle.Bold
    '        End If
    '        If PointChart.YAxis.Title.Italic Then
    '            myFontStyle = myFontStyle Or FontStyle.Italic
    '        End If
    '        If PointChart.YAxis.Title.Strikeout Then
    '            myFontStyle = myFontStyle Or FontStyle.Strikeout
    '        End If
    '        If PointChart.YAxis.Title.Underline Then
    '            myFontStyle = myFontStyle Or FontStyle.Underline
    '        End If

    '        Chart1.ChartAreas(0).AxisY.TitleFont = New Font("Arial", PointChart.YAxis.Title.Size, myFontStyle)
    '        Chart1.ChartAreas(0).AxisY.Title = PointChart.YAxis.Title.Text

    '        'Draw Chart Label:
    '        'Check if "Label1" is already in the list of titles:
    '        If Chart1.Titles.IndexOf("Label1") = -1 Then 'Label "Label1" doesnt exist
    '            Chart1.Titles.Add("Label1").Name = "Label1" 'The name needs to be explicitly declared!
    '        End If

    '        Chart1.Titles("Label1").Text = PointChart.ChartLabel.Text

    '        Dim myFontStyle2 As FontStyle = FontStyle.Regular
    '        If PointChart.ChartLabel.Bold Then
    '            myFontStyle2 = myFontStyle2 Or FontStyle.Bold
    '        End If
    '        If PointChart.ChartLabel.Italic Then
    '            myFontStyle2 = myFontStyle2 Or FontStyle.Italic
    '        End If
    '        If PointChart.ChartLabel.Strikeout Then
    '            myFontStyle2 = myFontStyle2 Or FontStyle.Strikeout
    '        End If
    '        If PointChart.ChartLabel.Underline Then
    '            myFontStyle2 = myFontStyle2 Or FontStyle.Underline
    '        End If

    '        Chart1.Titles("Label1").Font = New Font("Arial", PointChart.ChartLabel.Size, myFontStyle2)
    '        Chart1.Titles("Label1").Alignment = PointChart.ChartLabel.Alignment

    '        'Chart1.

    '        'Display selected chart information:
    '        Message.Add(vbCrLf & "Main.Chart1.ChartAreas.Count: " & Chart1.ChartAreas.Count & vbCrLf) '1
    '        Message.Add("Main.Chart1.ChartAreas(0).Name: " & Chart1.ChartAreas(0).Name & vbCrLf) 'ChartArea1
    '        Message.Add("Main.Chart1.ChartAreas(0).AxisX.Minimum: " & Chart1.ChartAreas(0).AxisX.Minimum & vbCrLf) '0
    '        Message.Add("Main.Chart1.ChartAreas(0).AxisX.Maximum: " & Chart1.ChartAreas(0).AxisX.Maximum & vbCrLf) '1
    '        Message.Add("Main.Chart1.ChartAreas(0).AxisY.Minimum: " & Chart1.ChartAreas(0).AxisY.Minimum & vbCrLf) 'NaN
    '        Message.Add("Main.Chart1.ChartAreas(0).AxisY.Maximum: " & Chart1.ChartAreas(0).AxisY.Maximum & vbCrLf) 'NaN
    '        Message.Add("Main.Chart1.Series(0).Name: " & Chart1.Series(0).Name & vbCrLf) 'Series1
    '        Message.Add("Main.Chart1.Series(0).Legend: " & Chart1.Series(0).Legend & vbCrLf) 'Legend1
    '        Message.Add("Main.Chart1.Series(0).YValueType: " & Chart1.Series(0).YValueType & vbCrLf) '2
    '        'Main.MessageAdd( "Main.Chart1.Series(0).AxisLabel(0): " & Main.Chart1.Series(0).AxisLabel(0) & vbCrLf) 'Index out of range
    '        'Main.MessageAdd( "Main.Chart1.Series(0).AxisLabel(1): " & Main.Chart1.Series(0).AxisLabel(1) & vbCrLf) 'Index out of range
    '        Message.Add("Main.Chart1.ChartAreas(0).AxisY2.Minimum: " & Chart1.ChartAreas(0).AxisY2.Minimum & vbCrLf) 'NaN unless specified prior
    '        Message.Add("Main.Chart1.ChartAreas(0).AxisY2.Maximum: " & Chart1.ChartAreas(0).AxisY2.Maximum & vbCrLf) 'NaN unless specified prior
    '        Message.Add("Main.Chart1.ChartAreas(0).AxisY.AxisName: " & Chart1.ChartAreas(0).AxisY.AxisName & vbCrLf) '1
    '        Message.Add("Main.Chart1.ChartAreas(0).AxisX.MajorGrid.Interval: " & Chart1.ChartAreas(0).AxisX.MajorGrid.Interval & vbCrLf) '1
    '        Message.Add("Main.Chart1.ChartAreas(0).AxisY.MajorGrid.Interval: " & Chart1.ChartAreas(0).AxisY.MajorGrid.Interval & vbCrLf) 'NaN unless specified prior

    '    Catch ex As Exception
    '        Message.AddWarning("Error drawing point chart: " & ex.Message & vbCrLf)
    '    End Try

    'End Sub

    'Private Sub TabPage4_Leave(sender As Object, e As EventArgs) Handles TabPage4.Leave
    '    'Leaving the chart Settigs tab page
    '    UpdatePointChartSettings()
    'End Sub


    'Private Sub UpdatePointChartSettings()
    '    'Update PointChart with the settings selected on the Chart Type, Titles, X Axis and Y Axis tabs.

    '    ''Update the Input Data settings:
    '    'PointChart.InputDataType = InputDataType
    '    'PointChart.InputDatabasePath = InputDatabasePath
    '    'PointChart.InputDataDescr = InputDataDescr
    '    'PointChart.InputQuery = InputQuery

    '    'Update chart properties:
    '    If txtSeriesName.Text <> "" Then
    '        PointChart.SeriesName = Trim(txtSeriesName.Text)
    '    End If

    '    If IsNothing(cmbXValues.SelectedItem) Then
    '        Message.AddWarning("The Field containing the XValues for the chart has not been selected." & vbCrLf)
    '    Else
    '        PointChart.XValuesFieldName = cmbXValues.SelectedItem.ToString
    '    End If

    '    If Trim(DataGridView1.Rows(0).Cells(1).Value) = "" Then
    '        Message.AddWarning("The Field containing the YValues for the chart has not been selected." & vbCrLf)
    '    Else
    '        PointChart.YValuesFieldName = DataGridView1.Rows(0).Cells(1).Value
    '    End If

    '    If Trim(DataGridView2.Rows(0).Cells(2).Value) = "" Then 'EmptyPointValue not specified
    '    Else
    '        PointChart.EmptyPointValue = DataGridView2.Rows(0).Cells(2).Value
    '    End If

    '    If Trim(DataGridView2.Rows(1).Cells(2).Value) = "" Then 'LabelStyle not specified
    '    Else
    '        PointChart.LabelStyle = DataGridView2.Rows(1).Cells(2).Value
    '    End If

    '    If Trim(DataGridView2.Rows(2).Cells(2).Value) = "" Then 'PixelPointDepth not specified
    '    Else
    '        PointChart.PixelPointDepth = DataGridView2.Rows(2).Cells(2).Value
    '    End If

    '    If Trim(DataGridView2.Rows(3).Cells(2).Value) = "" Then 'PixelPointGapDepth not specified
    '    Else
    '        PointChart.PixelPointGapDepth = DataGridView2.Rows(3).Cells(2).Value
    '    End If

    '    'Update Point Chart label settings:
    '    'PointChart.ChartLabel.FontName = FontDialog1.Font.Name
    '    'PointChart.ChartLabel.Size = FontDialog1.Font.Size
    '    'PointChart.ChartLabel.Bold = FontDialog1.Font.Bold
    '    'PointChart.ChartLabel.Italic = FontDialog1.Font.Italic
    '    'PointChart.ChartLabel.Strikeout = FontDialog1.Font.Strikeout
    '    'PointChart.ChartLabel.Underline = FontDialog1.Font.Underline
    '    'PointChart.ChartLabel.Text = txtChartTitle.Text
    '    PointChart.ChartLabel.FontName = txtChartTitle.Font.Name
    '    PointChart.ChartLabel.Size = txtChartTitle.Font.Size
    '    PointChart.ChartLabel.Bold = txtChartTitle.Font.Bold
    '    PointChart.ChartLabel.Italic = txtChartTitle.Font.Italic
    '    PointChart.ChartLabel.Strikeout = txtChartTitle.Font.Strikeout
    '    PointChart.ChartLabel.Underline = txtChartTitle.Font.Underline
    '    PointChart.ChartLabel.Text = txtChartTitle.Text
    '    If IsNothing(cmbAlignment.SelectedItem) Then
    '    Else
    '        PointChart.ChartLabel.Alignment = [Enum].Parse(GetType(ContentAlignment), cmbAlignment.SelectedItem.ToString)
    '    End If


    '    'Update X Axis settings:
    '    PointChart.XAxis.Title.FontName = txtXAxisTitle.Font.Name
    '    PointChart.XAxis.Title.Size = txtXAxisTitle.Font.Size
    '    PointChart.XAxis.Title.Bold = txtXAxisTitle.Font.Bold
    '    PointChart.XAxis.Title.Italic = txtXAxisTitle.Font.Italic
    '    PointChart.XAxis.Title.Strikeout = txtXAxisTitle.Font.Strikeout
    '    PointChart.XAxis.Title.Underline = txtXAxisTitle.Font.Underline
    '    PointChart.XAxis.Title.Text = txtXAxisTitle.Text

    '    If chkXAxisAutoMin.Checked = True Then
    '        PointChart.XAxis.AutoMinimum = True
    '    Else
    '        PointChart.XAxis.AutoMinimum = False
    '    End If

    '    If chkXAxisAutoMax.Checked = True Then
    '        PointChart.XAxis.AutoMaximum = True
    '    Else
    '        PointChart.XAxis.AutoMaximum = False
    '    End If

    '    PointChart.XAxis.Minimum = Val(txtXAxisMin.Text)
    '    PointChart.XAxis.Maximum = Val(txtXAxisMax.Text)

    '    If chkXAxisAutoAnnotInt.Checked = True Then
    '        PointChart.XAxis.Interval = 0 '0 indicates auto annotation.
    '        PointChart.XAxis.AutoInterval = True
    '    Else
    '        PointChart.XAxis.Interval = Val(txtXAxisAnnotInt.Text)
    '        PointChart.XAxis.AutoInterval = False
    '    End If

    '    If chkXAxisAutoMajGridInt.Checked = True Then
    '        PointChart.XAxis.MajorGridInterval = 0
    '        PointChart.XAxis.AutoMajorGridInterval = True
    '        'Message.Add("X Axis major grid interval set to auto." & vbCrLf)
    '    Else
    '        PointChart.XAxis.MajorGridInterval = Val(txtXAxisMajGridInt.Text)
    '        PointChart.XAxis.AutoMajorGridInterval = False
    '        'Message.Add("X Axis major grid interval set to: " & txtXAxisMajGridInt.Text & vbCrLf)
    '    End If

    '    'Update Y Axis settings:
    '    PointChart.YAxis.Title.FontName = txtYAxisTitle.Font.Name
    '    PointChart.YAxis.Title.Size = txtYAxisTitle.Font.Size
    '    PointChart.YAxis.Title.Bold = txtYAxisTitle.Font.Bold
    '    PointChart.YAxis.Title.Italic = txtYAxisTitle.Font.Italic
    '    PointChart.YAxis.Title.Strikeout = txtYAxisTitle.Font.Strikeout
    '    PointChart.YAxis.Title.Underline = txtYAxisTitle.Font.Underline

    '    PointChart.YAxis.Title.Text = txtYAxisTitle.Text

    '    If chkYAxisAutoMin.Checked = True Then
    '        PointChart.YAxis.AutoMinimum = True
    '    Else
    '        PointChart.YAxis.AutoMinimum = False
    '    End If

    '    If chkYAxisAutoMax.Checked = True Then
    '        PointChart.YAxis.AutoMaximum = True
    '    Else
    '        PointChart.YAxis.AutoMaximum = False
    '    End If

    '    PointChart.YAxis.Minimum = Val(txtYAxisMin.Text)
    '    PointChart.YAxis.Maximum = Val(txtYAxisMax.Text)

    '    If chkYAxisAutoAnnotInt.Checked = True Then
    '        PointChart.YAxis.Interval = 0 '0 indicates auto annotation.
    '        PointChart.YAxis.AutoInterval = True
    '    Else
    '        PointChart.YAxis.Interval = Val(txtYAxisAnnotInt.Text)
    '        PointChart.YAxis.AutoInterval = False
    '    End If

    '    If chkYAxisAutoMajGridInt.Checked = True Then
    '        PointChart.YAxis.MajorGridInterval = 0
    '        PointChart.YAxis.AutoMajorGridInterval = True
    '        'Message.Add("Y Axis major grid interval set to auto." & vbCrLf)
    '    Else
    '        PointChart.YAxis.MajorGridInterval = Val(txtYAxisMajGridInt.Text)
    '        PointChart.YAxis.AutoMajorGridInterval = False
    '        'Message.Add("Y Axis major grid interval set to:" & txtYAxisMajGridInt.Text & vbCrLf)
    '    End If

    'End Sub

    Private Sub bgwSendMessage_DoWork(sender As Object, e As DoWorkEventArgs) Handles bgwSendMessage.DoWork
        'Send a message on a separate thread:
        Try
            If IsNothing(client) Then
                bgwSendMessage.ReportProgress(1, "No Connection available. Message not sent!")
            Else
                If client.State = ServiceModel.CommunicationState.Faulted Then
                    bgwSendMessage.ReportProgress(1, "Connection state is faulted. Message not sent!")
                Else
                    Dim SendMessageParams As clsSendMessageParams = e.Argument
                    client.SendMessage(SendMessageParams.ProjectNetworkName, SendMessageParams.ConnectionName, SendMessageParams.Message)
                End If
            End If
        Catch ex As Exception
            bgwSendMessage.ReportProgress(1, ex.Message)
        End Try
    End Sub

    Private Sub bgwSendMessage_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles bgwSendMessage.ProgressChanged
        'Display an error message:
        Message.AddWarning("Send Message error: " & e.UserState.ToString & vbCrLf) 'Show the bgwSendMessage message 
    End Sub

    Private Sub bgwSendMessageAlt_DoWork(sender As Object, e As DoWorkEventArgs) Handles bgwSendMessageAlt.DoWork
        'Alternative SendMessage background worker - to send a message while instructions are being processed. 
        'Send a message on a separate thread
        Try
            If IsNothing(client) Then
                bgwSendMessageAlt.ReportProgress(1, "No Connection available. Message not sent!")
            Else
                If client.State = ServiceModel.CommunicationState.Faulted Then
                    bgwSendMessageAlt.ReportProgress(1, "Connection state is faulted. Message not sent!")
                Else
                    Dim SendMessageParamsAlt As clsSendMessageParams = e.Argument
                    client.SendMessage(SendMessageParamsAlt.ProjectNetworkName, SendMessageParamsAlt.ConnectionName, SendMessageParamsAlt.Message)
                End If
            End If
        Catch ex As Exception
            bgwSendMessageAlt.ReportProgress(1, ex.Message)
        End Try
    End Sub

    Private Sub bgwSendMessageAlt_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles bgwSendMessageAlt.ProgressChanged
        'Display an error message:
        Message.AddWarning("Send Message error: " & e.UserState.ToString & vbCrLf) 'Show the bgwSendMessage message 
    End Sub

    Private Sub XMsg_ErrorMsg(ErrMsg As String) Handles XMsg.ErrorMsg
        Message.AddWarning(ErrMsg & vbCrLf)
    End Sub

    Private Sub Message_ShowXMessagesChanged(Show As Boolean) Handles Message.ShowXMessagesChanged
        ShowXMessages = Show
    End Sub

    Private Sub Message_ShowSysMessagesChanged(Show As Boolean) Handles Message.ShowSysMessagesChanged
        ShowSysMessages = Show
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        'Save Point Chart file.

        Dim FileName As String = Trim(txtChartFileName.Text)

        If FileName = "" Then
            Message.AddWarning("Please enter a file name." & vbCrLf)
            Exit Sub
        End If

        If LCase(FileName).EndsWith(".pointchart") Then
            FileName = IO.Path.GetFileNameWithoutExtension(FileName) & ".PointChart"
        ElseIf FileName.Contains(".") Then
            Message.AddWarning("Unknown file extension: " & IO.Path.GetExtension(FileName) & vbCrLf)
            Exit Sub
        Else
            FileName = FileName & ".PointChart"
        End If

        'Message.Add("6- Chart1.ChartAreas(0).AxisX.Title = " & Chart1.ChartAreas(0).AxisX.Title & vbCrLf)
        'Message.Add("7- Chart1.ChartAreas(0).AxisX.Minimum = " & Chart1.ChartAreas(0).AxisX.Minimum & vbCrLf)
        'Message.Add("8- Chart1.ChartAreas(0).AxisX.Maximum = " & Chart1.ChartAreas(0).AxisX.Maximum & vbCrLf)
        'Message.Add("9- Chart1.ChartAreas(0).AxisX.Interval = " & Chart1.ChartAreas(0).AxisX.Interval & vbCrLf)
        'Message.Add("10- Chart1.ChartAreas(0).AxisX.MajorGrid.Interval = " & Chart1.ChartAreas(0).AxisX.MajorGrid.Interval & vbCrLf)


        txtChartFileName.Text = FileName
        Project.SaveXmlData(FileName, ChartInfo.ToXDoc(Chart1))
    End Sub

    Private Sub btnOpen_Click(sender As Object, e As EventArgs) Handles btnOpen.Click
        'Open Point Chart file.

        'Find and open a Line Chart file:
        Select Case Project.DataLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                'Select a Line Chart file from the project directory:
                OpenFileDialog1.InitialDirectory = Project.DataLocn.Path
                OpenFileDialog1.Filter = "Line Chart files | *.PointChart"
                If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                    Dim FileName As String = System.IO.Path.GetFileName(OpenFileDialog1.FileName)
                    txtChartFileName.Text = FileName
                    Try
                        ChartInfo.LoadFile(FileName, Chart1)
                    Catch ex As Exception
                        Message.AddWarning("Error in ChartInfo.LoadFile. FileName = " & FileName & vbCrLf)
                        Message.AddWarning(ex.Message & vbCrLf & vbCrLf)
                    End Try
                    UpdateInputDataTabSettings()
                    UpdateTitlesTabSettings()
                    UpdateAreasTabSettings() 'Update Areas Tab before Series Tab. This will update the list of Chart Areas on the Series Tab.
                    UpdateSeriesTabSettings()
                    'If chkAutoDraw.Checked Then DrawLineChart()
                    If chkAutoDraw.Checked Then DrawPointChart()
                End If

            Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
                'Select a Line Chart file from the project archive:
                'Show the zip archive file selection form:
                Zip = New ADVL_Utilities_Library_1.ZipComp
                Zip.ArchivePath = Project.DataLocn.Path
                Zip.SelectFile() 'Show the Select File form
                Zip.SelectFileForm.ApplicationName = Project.Application.Name
                Zip.SelectFileForm.SettingsLocn = Project.SettingsLocn
                Zip.SelectFileForm.Show()
                Zip.SelectFileForm.RestoreFormSettings()
                Zip.SelectFileForm.FileExtensions = {".PointChart"}
                Zip.SelectFileForm.GetFileList()
                If Zip.SelectedFile <> "" Then
                    'A file has been selected
                    txtChartFileName.Text = Zip.SelectedFile
                    Try
                        ChartInfo.LoadFile(Zip.SelectedFile, Chart1)
                    Catch ex As Exception
                        Message.AddWarning("Error in ChartInfo.LoadFile. Zip.SelectedFile = " & Zip.SelectedFile & vbCrLf)
                        Message.AddWarning(ex.Message & vbCrLf & vbCrLf)
                    End Try
                    UpdateInputDataTabSettings()
                    UpdateTitlesTabSettings()
                    UpdateAreasTabSettings()  'Update Areas Tab before Series Tab. This will update the list of Chart Areas on the Series Tab.
                    UpdateSeriesTabSettings()
                    'If chkAutoDraw.Checked Then DrawLineChart()
                    If chkAutoDraw.Checked Then DrawPointChart()
                End If
        End Select
    End Sub

    Private Sub Chart1_Click(sender As Object, e As EventArgs) Handles Chart1.Click

    End Sub

    Private Sub btnNew_Click(sender As Object, e As EventArgs) Handles btnNew.Click
        'Clear the current chart:

        Chart1.ChartAreas.Clear()
        ChartInfo.dictAreaInfo.Clear()
        Chart1.ChartAreas.Add("ChartArea1")
        ChartInfo.dictAreaInfo.Add("ChartArea1", New AreaInfo)

        Chart1.Series.Clear()
        ChartInfo.dictSeriesInfo.Clear()
        Chart1.Series.Add("Series1")
        ChartInfo.dictSeriesInfo.Add("Series1", New SeriesInfo)

        UpdateChartSettingsTabs()
    End Sub

    Private Sub UpdateChartSettingsTabs()

    End Sub

    Private Sub Project_NewProjectCreated(ProjectPath As String) Handles Project.NewProjectCreated
        SendProjectInfo(ProjectPath) 'Send the path of the new project to the Network application. The new project will be added to the list of projects.
    End Sub

    Private Sub Chart1_MouseDown(sender As Object, e As MouseEventArgs) Handles Chart1.MouseDown

        'Save the MouseDown settings:
        MouseDownXMin = Chart1.ChartAreas(0).AxisX.ScaleView.ViewMinimum
        MouseDownXMax = Chart1.ChartAreas(0).AxisX.ScaleView.ViewMaximum
        'MouseDownXRange = Chart1.ChartAreas(0).AxisX.ScaleView.ViewMaximum - MouseDownXMin
        MouseDownXRange = MouseDownXMax - MouseDownXMin
        'MouseDownX = Chart1.ChartAreas(0).AxisX.PixelPositionToValue(e.Location.X)
        MouseDownXPixel = e.Location.X
        MouseDownYMin = Chart1.ChartAreas(0).AxisY.ScaleView.ViewMinimum
        MouseDownYMax = Chart1.ChartAreas(0).AxisY.ScaleView.ViewMaximum
        'MouseDownYRange = Chart1.ChartAreas(0).AxisY.ScaleView.ViewMaximum - MouseDownYMin
        MouseDownYRange = MouseDownYMax - MouseDownYMin
        'MouseDownY = Chart1.ChartAreas(0).AxisY.PixelPositionToValue(e.Location.Y)
        MouseDownYPixel = e.Location.Y

        MouseDownXPixels = Chart1.ChartAreas(0).AxisX.ValueToPixelPosition(MouseDownXMax) - Chart1.ChartAreas(0).AxisX.ValueToPixelPosition(MouseDownXMin)
        MouseDownYPixels = Chart1.ChartAreas(0).AxisY.ValueToPixelPosition(MouseDownYMax) - Chart1.ChartAreas(0).AxisY.ValueToPixelPosition(MouseDownYMin)



    End Sub

    Private Sub Chart1_MouseMove(sender As Object, e As MouseEventArgs) Handles Chart1.MouseMove

        Dim XPixelPosn As Integer = e.Location.X 'Pixel position
        Dim YPixelPosn As Integer = e.Location.Y 'Pixel position

        Try
            'If ChartMode = "Pan" Then
            If ChartInfo.Mode = "Pan" Then
                'Me.Cursor = Cursors.Hand

                If MouseButtons.HasFlag(MouseButtons.Left) Then

                    Dim DXPixels As Integer = XPixelPosn - MouseDownXPixel
                    Dim DYPixels As Integer = YPixelPosn - MouseDownYPixel
                    Dim DX As Double = DXPixels * MouseDownXRange / MouseDownXPixels
                    Dim DY As Double = DYPixels * MouseDownYRange / MouseDownYPixels

                    Dim XMin As Double = MouseDownXMin - DX
                    Dim XMax As Double = XMin + MouseDownXRange
                    Dim YMin As Double = MouseDownYMin - DY
                    Dim YMax As Double = YMin + MouseDownYRange

                    If XMin < Chart1.ChartAreas(0).AxisX.Minimum Then
                        XMin = Chart1.ChartAreas(0).AxisX.Minimum
                        XMax = XMin + MouseDownXRange
                    End If

                    If XMax > Chart1.ChartAreas(0).AxisX.Maximum Then
                        XMax = Chart1.ChartAreas(0).AxisX.Maximum
                        XMin = XMax - MouseDownXRange
                    End If

                    If YMin < Chart1.ChartAreas(0).AxisY.Minimum Then
                        YMin = Chart1.ChartAreas(0).AxisY.Minimum
                        YMax = YMin + MouseDownYRange
                    End If

                    If YMax > Chart1.ChartAreas(0).AxisY.Maximum Then
                        YMax = Chart1.ChartAreas(0).AxisY.Maximum
                        YMin = YMax - MouseDownYRange
                    End If

                    'Apply the Chart Pan:
                    Chart1.ChartAreas(0).AxisX.ScaleView.Zoom(XMin, XMax)
                    Chart1.ChartAreas(0).AxisY.ScaleView.Zoom(YMin, YMax)

                    'Chart1.ChartAreas(0).AxisX.IntervalAutoMode = IntervalAutoMode.VariableCount
                    'Chart1.ChartAreas(0).AxisY.IntervalAutoMode = IntervalAutoMode.VariableCount


                    'Chart1.ChartAreas(0).AxisX.IntervalOffset = 0
                    'Chart1.ChartAreas(0).AxisY.IntervalOffset = 0

                    'Chart1.ChartAreas(0).AxisX.IntervalOffset = Double.NaN
                    'Chart1.ChartAreas(0).AxisY.IntervalOffset = Double.NaN

                    'Chart1.ChartAreas(0).AxisX.Crossing = 0
                    'Chart1.ChartAreas(0).AxisY.Crossing = 0

                    Chart1.ChartAreas(0).AxisX.Crossing = Double.NaN
                    Chart1.ChartAreas(0).AxisY.Crossing = Double.NaN

                    'Chart1.ChartAreas(0).AxisX.MajorGrid.IntervalOffset = 0
                    'Chart1.ChartAreas(0).AxisY.MajorGrid.IntervalOffset = 0

                    'Chart1.ChartAreas(0).AxisX.MajorGrid.IntervalOffset = Double.NaN
                    'Chart1.ChartAreas(0).AxisY.MajorGrid.IntervalOffset = Double.NaN



                    'Chart1.ChartAreas(0).RecalculateAxesScale()

                    Chart1.ChartAreas(0).AxisX.Interval = 0
                    Chart1.ChartAreas(0).AxisX.MajorGrid.Interval = 0
                    'Chart1.ChartAreas(0).AxisX.RoundAxisValues()
                    'Chart1.ChartAreas(0).AxisX.IsStartedFromZero = False
                    Chart1.ChartAreas(0).AxisX.IsStartedFromZero = True
                    Chart1.ChartAreas(0).AxisX.RoundAxisValues()
                    'Chart1.ChartAreas(0).RecalculateAxesScale()

                    'Chart1.ChartAreas(0).AxisY.Interval = 0
                    'Chart1.ChartAreas(0).AxisY.MajorGrid.Interval = 0
                    'Chart1.ChartAreas(0).AxisY.IsStartedFromZero = True
                    Chart1.ChartAreas(0).AxisY.RoundAxisValues()
                End If

                'ElseIf ChartMode = "Select" Then
            ElseIf ChartInfo.Mode = "Select" Then
                'Me.Cursor = Cursors.Cross

            End If

            'Show the mouse position
            If IsNothing(ChartInfoForm) Then

            Else
                'ChartInfoForm.txtXCursorPosn.Text = Chart1.ChartAreas(0).CursorX.Position.ToString
                'ChartInfoForm.txtYCursorPosn.Text = Chart1.ChartAreas(0).CursorY.Position.ToString

                ChartInfoForm.txtMouseX.Text = Chart1.ChartAreas(0).AxisX.PixelPositionToValue(XPixelPosn)
                ChartInfoForm.txtMouseY.Text = Chart1.ChartAreas(0).AxisY.PixelPositionToValue(YPixelPosn)
                ChartInfoForm.txtPixelX.Text = XPixelPosn
                ChartInfoForm.txtPixelY.Text = YPixelPosn

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Chart1_MouseClick(sender As Object, e As MouseEventArgs) Handles Chart1.MouseClick


        If IsNothing(ChartInfoForm) Then
            'Chart Info Form is not displayed
        Else
            If Chart1.ChartAreas(0).CursorX.IsUserEnabled = True Then
                ChartInfoForm.txtXCursorPosn.Text = Chart1.ChartAreas(0).CursorX.Position.ToString 'Show the X Cursor position
            End If
            If Chart1.ChartAreas(0).CursorY.IsUserEnabled = True Then
                ChartInfoForm.txtYCursorPosn.Text = Chart1.ChartAreas(0).CursorY.Position.ToString  'Show the Y Cursor position
            End If
        End If

        Dim Result As HitTestResult = Chart1.HitTest(e.X, e.Y)
        If Result.ChartElementType = ChartElementType.DataPoint Then
            If IsNothing(ChartInfoForm) Then

            Else
                Dim I As Integer = Result.PointIndex
                ChartInfoForm.txtSelPointX.Text = Chart1.Series(0).Points(I).XValue.ToString
                ChartInfoForm.txtSelPointY.Text = Chart1.Series(0).Points(I).YValues(0).ToString
                ChartInfoForm.txtPointNo.Text = I

                'Add the selected point to selPoints:
                selPoints.Add(Chart1.Series(ChartInfo.SelSeries).Points(I))

                'Change the color of the point:
                Chart1.Series(0).Points(I).MarkerSize = SelPtSize
                Chart1.Series(0).Points(I).MarkerColor = SelPtColor
                Chart1.Series(0).Points(I).MarkerBorderColor = SelPtBorderColor
                Chart1.Series(0).Points(I).MarkerBorderWidth = SelPtBorderWidth
                Chart1.Series(0).Points(I).MarkerStyle = SelPtStyle

                If IsNothing(ChartInfoForm) Then
                Else
                    ChartInfoForm.txtNSelPoints.Text = selPoints.Count
                End If

                If IsNothing(ViewDatabaseData) Then
                    'Data table is not displayed
                Else
                    'Highlight the selected data point in the data table:
                    ViewDatabaseData.DataGridView1.ClearSelection()
                    ViewDatabaseData.DataGridView1.Rows(I).Selected = True
                    ViewDatabaseData.DataGridView1.FirstDisplayedScrollingRowIndex = I
                End If

            End If
        End If
    End Sub

    Private Sub Chart1_MouseWheel(sender As Object, e As MouseEventArgs) Handles Chart1.MouseWheel
        'Zoom the chart

        Dim ZoomPercent As Integer = 10 'Zoom in by this percentage each increment
        Dim XMin As Double = Chart1.ChartAreas(0).AxisX.ScaleView.ViewMinimum
        Dim XMax As Double = Chart1.ChartAreas(0).AxisX.ScaleView.ViewMaximum
        Dim YMin As Double = Chart1.ChartAreas(0).AxisY.ScaleView.ViewMinimum
        Dim YMax As Double = Chart1.ChartAreas(0).AxisY.ScaleView.ViewMaximum
        Dim XInc = (XMax - XMin) * ZoomPercent / 200
        Dim YInc = (YMax - YMin) * ZoomPercent / 200

        If e.Delta < 0 Then
            'Scrolling out
            XMin -= XInc
            XMax += XInc
            YMin -= YInc
            YMax += YInc

            If XMin < Chart1.ChartAreas(0).AxisX.Minimum Then XMin = Chart1.ChartAreas(0).AxisX.Minimum
            If XMax > Chart1.ChartAreas(0).AxisX.Maximum Then XMax = Chart1.ChartAreas(0).AxisX.Maximum
            If YMin < Chart1.ChartAreas(0).AxisY.Minimum Then XMin = Chart1.ChartAreas(0).AxisY.Minimum
            If YMax > Chart1.ChartAreas(0).AxisY.Maximum Then XMax = Chart1.ChartAreas(0).AxisY.Maximum

            Chart1.ChartAreas(0).AxisX.ScaleView.Zoom(XMin, XMax)
            Chart1.ChartAreas(0).AxisY.ScaleView.Zoom(YMin, YMax)

            Chart1.ChartAreas(0).AxisX.Interval = 0
            Chart1.ChartAreas(0).AxisX.MajorGrid.Interval = 0
            Chart1.ChartAreas(0).AxisX.IsStartedFromZero = True
            Chart1.ChartAreas(0).AxisX.RoundAxisValues()

            Chart1.ChartAreas(0).AxisY.Interval = 0
            Chart1.ChartAreas(0).AxisY.MajorGrid.Interval = 0
            Chart1.ChartAreas(0).AxisY.IsStartedFromZero = True
            Chart1.ChartAreas(0).AxisY.RoundAxisValues()

            Chart1.ChartAreas(0).RecalculateAxesScale()

        Else
            'Scrolling in
            XMin += XInc
            XMax -= XInc
            YMin += YInc
            YMax -= YInc

            If XMin > XMax Then
                XMin -= XInc
                XMax += XInc
            End If

            If YMin > YMax Then
                YMin -= YInc
                YMax += YInc
            End If

            Chart1.ChartAreas(0).AxisX.ScaleView.Zoom(XMin, XMax)
            Chart1.ChartAreas(0).AxisY.ScaleView.Zoom(YMin, YMax)

            Chart1.ChartAreas(0).AxisX.Interval = 0
            Chart1.ChartAreas(0).AxisX.MajorGrid.Interval = 0
            Chart1.ChartAreas(0).AxisX.IsStartedFromZero = True
            Chart1.ChartAreas(0).AxisX.RoundAxisValues()

            Chart1.ChartAreas(0).AxisY.Interval = 0
            Chart1.ChartAreas(0).AxisY.MajorGrid.Interval = 0
            Chart1.ChartAreas(0).AxisY.IsStartedFromZero = True
            Chart1.ChartAreas(0).AxisY.RoundAxisValues()

            Chart1.ChartAreas(0).RecalculateAxesScale()
        End If
    End Sub

    Private Sub Chart1_MouseUp(sender As Object, e As MouseEventArgs) Handles Chart1.MouseUp

        'Check if the chart is being zoomed:
        If Chart1.ChartAreas(ChartInfo.SelChartArea).AxisX.ScaleView.Zoomable = False Then
            If Chart1.ChartAreas(ChartInfo.SelChartArea).AxisY.ScaleView.Zoomable = False Then
                'Chart is not being zoomed.
                'Check if points are to be selected:
                If Chart1.ChartAreas(ChartInfo.SelChartArea).CursorX.IsUserSelectionEnabled = True Then
                    If Chart1.ChartAreas(ChartInfo.SelChartArea).CursorY.IsUserSelectionEnabled = True Then
                        'X axis and Y axis selections enabled.
                        GetPointsInSelectionArea()
                    Else
                        'X axis selection enabled only.
                        GetPointsInSelectionArea()
                    End If
                Else
                    If Chart1.ChartAreas(ChartInfo.SelChartArea).CursorY.IsUserSelectionEnabled = True Then
                        'Y axis selection enabled only.
                        GetPointsInSelectionArea()
                    Else
                        'No axis selections enabled.
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub Chart1_MouseEnter(sender As Object, e As EventArgs) Handles Chart1.MouseEnter

        'If ChartMode = "Pan" Then
        If ChartInfo.Mode = "Pan" Then
            Me.Cursor = Cursors.Hand
            'ElseIf ChartMode = "Select" Then
        ElseIf ChartInfo.Mode = "Select" Then
            Me.Cursor = Cursors.Cross
        End If

    End Sub

    Private Sub Chart1_MouseLeave(sender As Object, e As EventArgs) Handles Chart1.MouseLeave
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Chart1_SelectionRangeChanged(sender As Object, e As CursorEventArgs) Handles Chart1.SelectionRangeChanged

        'Message.Add("Selection range changed." & vbCrLf)

        'Message.Add("Chart1.Series(0).Points.Count:  " & Chart1.Series(0).Points.Count & vbCrLf)

        'Message.Add("e.Axis.ToString  " & e.Axis.ToString & vbCrLf)
        'Message.Add("e.Axis.Name  " & e.Axis.Name & vbCrLf)
        'Message.Add("e.NewSelectionStart.ToString  " & e.NewSelectionStart.ToString & vbCrLf)
        'Message.Add("e.NewSelectionEnd.ToString  " & e.NewSelectionEnd.ToString & vbCrLf)

    End Sub

    Public Sub GetPointsInSelectionArea()
        'Selects the Series Points in the Cursor Selection area.

        If ChartInfo.SelChartArea = "" Then
            Message.Add("Select a Chart Area" & vbCrLf)
            Exit Sub
        End If
        If ChartInfo.SelSeries = "" Then
            Message.Add("Select a Chart Series" & vbCrLf)
            Exit Sub
        End If

        Dim XStart As Double = Chart1.ChartAreas(ChartInfo.SelChartArea).AxisX.ValueToPixelPosition(Chart1.ChartAreas(ChartInfo.SelChartArea).CursorX.SelectionStart)
        Dim XEnd As Double = Chart1.ChartAreas(ChartInfo.SelChartArea).AxisX.ValueToPixelPosition(Chart1.ChartAreas(ChartInfo.SelChartArea).CursorX.SelectionEnd)
        Dim YStart As Double = Chart1.ChartAreas(ChartInfo.SelChartArea).AxisY.ValueToPixelPosition(Chart1.ChartAreas(ChartInfo.SelChartArea).CursorY.SelectionStart)
        Dim YEnd As Double = Chart1.ChartAreas(ChartInfo.SelChartArea).AxisY.ValueToPixelPosition(Chart1.ChartAreas(ChartInfo.SelChartArea).CursorY.SelectionEnd)

        'XStart must be less than XEnd:
        If XStart > XEnd Then
            XStart = XEnd
            XEnd = Chart1.ChartAreas(ChartInfo.SelChartArea).AxisX.ValueToPixelPosition(Chart1.ChartAreas(ChartInfo.SelChartArea).CursorX.SelectionStart)
        End If

        'YStart must be less than YEnd:
        If YStart > YEnd Then
            YStart = YEnd
            YEnd = Chart1.ChartAreas(ChartInfo.SelChartArea).AxisY.ValueToPixelPosition(Chart1.ChartAreas(ChartInfo.SelChartArea).CursorY.SelectionStart)
        End If

        Dim selRect As Rectangle
        selRect.X = Int(XStart) 'Upper Left  X coordinate
        selRect.Y = Int(YStart) 'Upper Left Y coordinate
        selRect.Width = Int(XEnd - XStart) 'Width of the rectangle
        selRect.Height = Int(YEnd - YStart) 'Height of the rectangle

        For Each dp As DataPoint In Chart1.Series(ChartInfo.SelSeries).Points
            Dim X As Integer = Chart1.ChartAreas(ChartInfo.SelChartArea).AxisX.ValueToPixelPosition(dp.XValue)
            Dim Y As Integer = Chart1.ChartAreas(ChartInfo.SelChartArea).AxisY.ValueToPixelPosition(dp.YValues(0))
            If selRect.Contains(New Point(X, Y)) Then
                selPoints.Add(dp)
                dp.MarkerBorderColor = SelPtBorderColor
                dp.MarkerBorderWidth = SelPtBorderWidth
                dp.MarkerColor = SelPtColor
                dp.MarkerSize = SelPtSize
                dp.MarkerStyle = SelPtStyle
            End If
        Next

        If IsNothing(ChartInfoForm) Then
        Else
            ChartInfoForm.txtNSelPoints.Text = selPoints.Count
        End If

    End Sub

    Private Sub txtChartTitle_TextChanged(sender As Object, e As EventArgs) Handles txtChartTitle.TextChanged

    End Sub

    Private Sub cmbAlignment_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbAlignment.SelectedIndexChanged

    End Sub

    Private Sub btnChartTitleFont_Click(sender As Object, e As EventArgs) Handles btnChartTitleFont.Click

    End Sub

    Private Sub btnDelTitle_Click(sender As Object, e As EventArgs) Handles btnDelTitle.Click

    End Sub

    Private Sub cmbOrientation_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbOrientation.SelectedIndexChanged

    End Sub













#End Region 'Form Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------


End Class

Public Class clsSendMessageParams
    'Parameters used when sending a message using the Message Service.
    Public ProjectNetworkName As String
    Public ConnectionName As String
    Public Message As String
End Class

