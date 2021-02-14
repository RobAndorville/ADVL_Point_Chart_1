<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmChartInfoForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.btnShowRange = New System.Windows.Forms.Button()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.txtPixelY = New System.Windows.Forms.TextBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.txtPixelX = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtMouseX = New System.Windows.Forms.TextBox()
        Me.txtMouseY = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.btnClearAllSelPoints = New System.Windows.Forms.Button()
        Me.btnClearSelectedPoints = New System.Windows.Forms.Button()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.cmbMarkerStyle = New System.Windows.Forms.ComboBox()
        Me.txtMarkerSize = New System.Windows.Forms.TextBox()
        Me.txtMarkerColor = New System.Windows.Forms.TextBox()
        Me.txtBorderWidth = New System.Windows.Forms.TextBox()
        Me.txtBorderColor = New System.Windows.Forms.TextBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.btnPointProperties = New System.Windows.Forms.Button()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.txtPointNo = New System.Windows.Forms.TextBox()
        Me.btnViewData = New System.Windows.Forms.Button()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.txtSelPointX = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.txtSelPointY = New System.Windows.Forms.TextBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.chkShowXCursor = New System.Windows.Forms.CheckBox()
        Me.chkShowYCursor = New System.Windows.Forms.CheckBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtXInterval = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtYCursorPosn = New System.Windows.Forms.TextBox()
        Me.btnShowCursorInfo = New System.Windows.Forms.Button()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.txtYInterval = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.txtXCursorPosn = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.rbSelectRange = New System.Windows.Forms.RadioButton()
        Me.rbZoomRange = New System.Windows.Forms.RadioButton()
        Me.chkSelectYRange = New System.Windows.Forms.CheckBox()
        Me.chkSelectXRange = New System.Windows.Forms.CheckBox()
        Me.chkAllowZoom = New System.Windows.Forms.CheckBox()
        Me.txtChartArea = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.btnApply = New System.Windows.Forms.Button()
        Me.txtToolTipString = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmbSeriesList = New System.Windows.Forms.ComboBox()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.btnZoomReset = New System.Windows.Forms.Button()
        Me.rbSelect = New System.Windows.Forms.RadioButton()
        Me.rbPan = New System.Windows.Forms.RadioButton()
        Me.txtNSelPoints = New System.Windows.Forms.TextBox()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnExit
        '
        Me.btnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExit.Location = New System.Drawing.Point(731, 12)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(48, 22)
        Me.btnExit.TabIndex = 8
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(12, 40)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(767, 595)
        Me.TabControl1.TabIndex = 9
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.btnShowRange)
        Me.TabPage1.Controls.Add(Me.Label22)
        Me.TabPage1.Controls.Add(Me.Label17)
        Me.TabPage1.Controls.Add(Me.GroupBox4)
        Me.TabPage1.Controls.Add(Me.GroupBox3)
        Me.TabPage1.Controls.Add(Me.GroupBox2)
        Me.TabPage1.Controls.Add(Me.GroupBox1)
        Me.TabPage1.Controls.Add(Me.chkAllowZoom)
        Me.TabPage1.Controls.Add(Me.txtChartArea)
        Me.TabPage1.Controls.Add(Me.Label8)
        Me.TabPage1.Controls.Add(Me.btnApply)
        Me.TabPage1.Controls.Add(Me.txtToolTipString)
        Me.TabPage1.Controls.Add(Me.Label2)
        Me.TabPage1.Controls.Add(Me.Label1)
        Me.TabPage1.Controls.Add(Me.cmbSeriesList)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(759, 569)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Point Information"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'btnShowRange
        '
        Me.btnShowRange.Location = New System.Drawing.Point(375, 113)
        Me.btnShowRange.Name = "btnShowRange"
        Me.btnShowRange.Size = New System.Drawing.Size(83, 22)
        Me.btnShowRange.TabIndex = 63
        Me.btnShowRange.Text = "Show Range"
        Me.btnShowRange.UseVisualStyleBackColor = True
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Font = New System.Drawing.Font("Arial Narrow", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label22.Location = New System.Drawing.Point(6, 95)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(650, 15)
        Me.Label22.TabIndex = 62
        Me.Label22.Text = "Keywords: #VALX, #VAL, #VALY, #VALY2, #VALY3, #SERIESNAME, #AXISLABEL, #INDEX, #P" &
    "ERCENT, #TOTAL, #AVG, #MIN, #MAX, #FIRST, #LAST"
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(86, 82)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(220, 13)
        Me.Label17.TabIndex = 61
        Me.Label17.Text = "Example: X: #VALX{0.000} Y: #VALY{0.000}"
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.txtPixelY)
        Me.GroupBox4.Controls.Add(Me.Label15)
        Me.GroupBox4.Controls.Add(Me.Label14)
        Me.GroupBox4.Controls.Add(Me.txtPixelX)
        Me.GroupBox4.Controls.Add(Me.Label3)
        Me.GroupBox4.Controls.Add(Me.Label4)
        Me.GroupBox4.Controls.Add(Me.txtMouseX)
        Me.GroupBox4.Controls.Add(Me.txtMouseY)
        Me.GroupBox4.Controls.Add(Me.Label5)
        Me.GroupBox4.Location = New System.Drawing.Point(6, 359)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(127, 147)
        Me.GroupBox4.TabIndex = 60
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Mouse Position:"
        '
        'txtPixelY
        '
        Me.txtPixelY.Location = New System.Drawing.Point(29, 114)
        Me.txtPixelY.Name = "txtPixelY"
        Me.txtPixelY.Size = New System.Drawing.Size(87, 20)
        Me.txtPixelY.TabIndex = 18
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(6, 117)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(17, 13)
        Me.Label15.TabIndex = 17
        Me.Label15.Text = "Y:"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(6, 91)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(17, 13)
        Me.Label14.TabIndex = 16
        Me.Label14.Text = "X:"
        '
        'txtPixelX
        '
        Me.txtPixelX.Location = New System.Drawing.Point(29, 88)
        Me.txtPixelX.Name = "txtPixelX"
        Me.txtPixelX.Size = New System.Drawing.Size(87, 20)
        Me.txtPixelX.TabIndex = 15
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(6, 72)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(72, 13)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "Pixel Position:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 22)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(17, 13)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "X:"
        '
        'txtMouseX
        '
        Me.txtMouseX.Location = New System.Drawing.Point(29, 19)
        Me.txtMouseX.Name = "txtMouseX"
        Me.txtMouseX.Size = New System.Drawing.Size(87, 20)
        Me.txtMouseX.TabIndex = 11
        '
        'txtMouseY
        '
        Me.txtMouseY.Location = New System.Drawing.Point(29, 45)
        Me.txtMouseY.Name = "txtMouseY"
        Me.txtMouseY.Size = New System.Drawing.Size(87, 20)
        Me.txtMouseY.TabIndex = 14
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(6, 48)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(17, 13)
        Me.Label5.TabIndex = 13
        Me.Label5.Text = "Y:"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Label23)
        Me.GroupBox3.Controls.Add(Me.txtNSelPoints)
        Me.GroupBox3.Controls.Add(Me.btnClearAllSelPoints)
        Me.GroupBox3.Controls.Add(Me.btnClearSelectedPoints)
        Me.GroupBox3.Controls.Add(Me.GroupBox5)
        Me.GroupBox3.Controls.Add(Me.btnPointProperties)
        Me.GroupBox3.Controls.Add(Me.Label13)
        Me.GroupBox3.Controls.Add(Me.txtPointNo)
        Me.GroupBox3.Controls.Add(Me.btnViewData)
        Me.GroupBox3.Controls.Add(Me.Label10)
        Me.GroupBox3.Controls.Add(Me.txtSelPointX)
        Me.GroupBox3.Controls.Add(Me.Label9)
        Me.GroupBox3.Controls.Add(Me.txtSelPointY)
        Me.GroupBox3.Location = New System.Drawing.Point(139, 216)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(396, 347)
        Me.GroupBox3.TabIndex = 59
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Selected Point(s):"
        '
        'btnClearAllSelPoints
        '
        Me.btnClearAllSelPoints.Location = New System.Drawing.Point(166, 101)
        Me.btnClearAllSelPoints.Name = "btnClearAllSelPoints"
        Me.btnClearAllSelPoints.Size = New System.Drawing.Size(99, 22)
        Me.btnClearAllSelPoints.TabIndex = 66
        Me.btnClearAllSelPoints.Text = "Clear All Points"
        Me.btnClearAllSelPoints.UseVisualStyleBackColor = True
        '
        'btnClearSelectedPoints
        '
        Me.btnClearSelectedPoints.Location = New System.Drawing.Point(166, 73)
        Me.btnClearSelectedPoints.Name = "btnClearSelectedPoints"
        Me.btnClearSelectedPoints.Size = New System.Drawing.Size(99, 22)
        Me.btnClearSelectedPoints.TabIndex = 65
        Me.btnClearSelectedPoints.Text = "Clear Point"
        Me.btnClearSelectedPoints.UseVisualStyleBackColor = True
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.cmbMarkerStyle)
        Me.GroupBox5.Controls.Add(Me.txtMarkerSize)
        Me.GroupBox5.Controls.Add(Me.txtMarkerColor)
        Me.GroupBox5.Controls.Add(Me.txtBorderWidth)
        Me.GroupBox5.Controls.Add(Me.txtBorderColor)
        Me.GroupBox5.Controls.Add(Me.Label16)
        Me.GroupBox5.Controls.Add(Me.Label18)
        Me.GroupBox5.Controls.Add(Me.Label19)
        Me.GroupBox5.Controls.Add(Me.Label20)
        Me.GroupBox5.Controls.Add(Me.Label21)
        Me.GroupBox5.Location = New System.Drawing.Point(6, 45)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(145, 163)
        Me.GroupBox5.TabIndex = 58
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Marker:"
        '
        'cmbMarkerStyle
        '
        Me.cmbMarkerStyle.FormattingEnabled = True
        Me.cmbMarkerStyle.Location = New System.Drawing.Point(6, 129)
        Me.cmbMarkerStyle.Name = "cmbMarkerStyle"
        Me.cmbMarkerStyle.Size = New System.Drawing.Size(124, 21)
        Me.cmbMarkerStyle.TabIndex = 333
        '
        'txtMarkerSize
        '
        Me.txtMarkerSize.Location = New System.Drawing.Point(81, 91)
        Me.txtMarkerSize.Name = "txtMarkerSize"
        Me.txtMarkerSize.Size = New System.Drawing.Size(39, 20)
        Me.txtMarkerSize.TabIndex = 331
        '
        'txtMarkerColor
        '
        Me.txtMarkerColor.Location = New System.Drawing.Point(81, 65)
        Me.txtMarkerColor.Name = "txtMarkerColor"
        Me.txtMarkerColor.Size = New System.Drawing.Size(39, 20)
        Me.txtMarkerColor.TabIndex = 330
        '
        'txtBorderWidth
        '
        Me.txtBorderWidth.Location = New System.Drawing.Point(81, 39)
        Me.txtBorderWidth.Name = "txtBorderWidth"
        Me.txtBorderWidth.Size = New System.Drawing.Size(39, 20)
        Me.txtBorderWidth.TabIndex = 329
        '
        'txtBorderColor
        '
        Me.txtBorderColor.Location = New System.Drawing.Point(80, 13)
        Me.txtBorderColor.Name = "txtBorderColor"
        Me.txtBorderColor.Size = New System.Drawing.Size(40, 20)
        Me.txtBorderColor.TabIndex = 328
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(6, 113)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(31, 13)
        Me.Label16.TabIndex = 327
        Me.Label16.Text = "style:"
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(45, 94)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(28, 13)
        Me.Label18.TabIndex = 325
        Me.Label18.Text = "size:"
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(6, 68)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(69, 13)
        Me.Label19.TabIndex = 324
        Me.Label19.Text = "Marker color:"
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(38, 42)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(35, 13)
        Me.Label20.TabIndex = 323
        Me.Label20.Text = "width:"
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(6, 16)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(67, 13)
        Me.Label21.TabIndex = 322
        Me.Label21.Text = "Border color:"
        '
        'btnPointProperties
        '
        Me.btnPointProperties.Location = New System.Drawing.Point(166, 45)
        Me.btnPointProperties.Name = "btnPointProperties"
        Me.btnPointProperties.Size = New System.Drawing.Size(99, 22)
        Me.btnPointProperties.TabIndex = 64
        Me.btnPointProperties.Text = "Point Properties"
        Me.btnPointProperties.UseVisualStyleBackColor = True
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(6, 21)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(49, 13)
        Me.Label13.TabIndex = 38
        Me.Label13.Text = "Point no."
        '
        'txtPointNo
        '
        Me.txtPointNo.Location = New System.Drawing.Point(61, 19)
        Me.txtPointNo.Name = "txtPointNo"
        Me.txtPointNo.Size = New System.Drawing.Size(76, 20)
        Me.txtPointNo.TabIndex = 39
        '
        'btnViewData
        '
        Me.btnViewData.Location = New System.Drawing.Point(6, 253)
        Me.btnViewData.Name = "btnViewData"
        Me.btnViewData.Size = New System.Drawing.Size(103, 22)
        Me.btnViewData.TabIndex = 57
        Me.btnViewData.Text = "View Data Table"
        Me.btnViewData.UseVisualStyleBackColor = True
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(143, 22)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(17, 13)
        Me.Label10.TabIndex = 28
        Me.Label10.Text = "X:"
        '
        'txtSelPointX
        '
        Me.txtSelPointX.Location = New System.Drawing.Point(166, 18)
        Me.txtSelPointX.Name = "txtSelPointX"
        Me.txtSelPointX.Size = New System.Drawing.Size(76, 20)
        Me.txtSelPointX.TabIndex = 27
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(248, 21)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(17, 13)
        Me.Label9.TabIndex = 29
        Me.Label9.Text = "Y:"
        '
        'txtSelPointY
        '
        Me.txtSelPointY.Location = New System.Drawing.Point(271, 18)
        Me.txtSelPointY.Name = "txtSelPointY"
        Me.txtSelPointY.Size = New System.Drawing.Size(76, 20)
        Me.txtSelPointY.TabIndex = 30
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.chkShowXCursor)
        Me.GroupBox2.Controls.Add(Me.chkShowYCursor)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Controls.Add(Me.txtXInterval)
        Me.GroupBox2.Controls.Add(Me.Label7)
        Me.GroupBox2.Controls.Add(Me.txtYCursorPosn)
        Me.GroupBox2.Controls.Add(Me.btnShowCursorInfo)
        Me.GroupBox2.Controls.Add(Me.Label12)
        Me.GroupBox2.Controls.Add(Me.txtYInterval)
        Me.GroupBox2.Controls.Add(Me.Label11)
        Me.GroupBox2.Controls.Add(Me.txtXCursorPosn)
        Me.GroupBox2.Location = New System.Drawing.Point(6, 111)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(363, 99)
        Me.GroupBox2.TabIndex = 58
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Axis Cursors:"
        '
        'chkShowXCursor
        '
        Me.chkShowXCursor.AutoSize = True
        Me.chkShowXCursor.Location = New System.Drawing.Point(6, 20)
        Me.chkShowXCursor.Name = "chkShowXCursor"
        Me.chkShowXCursor.Size = New System.Drawing.Size(95, 17)
        Me.chkShowXCursor.TabIndex = 15
        Me.chkShowXCursor.Text = "Show X cursor"
        Me.chkShowXCursor.UseVisualStyleBackColor = True
        '
        'chkShowYCursor
        '
        Me.chkShowYCursor.AutoSize = True
        Me.chkShowYCursor.Location = New System.Drawing.Point(6, 44)
        Me.chkShowYCursor.Name = "chkShowYCursor"
        Me.chkShowYCursor.Size = New System.Drawing.Size(95, 17)
        Me.chkShowYCursor.TabIndex = 18
        Me.chkShowYCursor.Text = "Show Y cursor"
        Me.chkShowYCursor.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(107, 21)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(45, 13)
        Me.Label6.TabIndex = 16
        Me.Label6.Text = "Interval:"
        '
        'txtXInterval
        '
        Me.txtXInterval.Location = New System.Drawing.Point(158, 16)
        Me.txtXInterval.Name = "txtXInterval"
        Me.txtXInterval.Size = New System.Drawing.Size(76, 20)
        Me.txtXInterval.TabIndex = 17
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(107, 46)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(45, 13)
        Me.Label7.TabIndex = 19
        Me.Label7.Text = "Interval:"
        '
        'txtYCursorPosn
        '
        Me.txtYCursorPosn.Location = New System.Drawing.Point(275, 42)
        Me.txtYCursorPosn.Name = "txtYCursorPosn"
        Me.txtYCursorPosn.Size = New System.Drawing.Size(76, 20)
        Me.txtYCursorPosn.TabIndex = 34
        '
        'btnShowCursorInfo
        '
        Me.btnShowCursorInfo.Location = New System.Drawing.Point(6, 67)
        Me.btnShowCursorInfo.Name = "btnShowCursorInfo"
        Me.btnShowCursorInfo.Size = New System.Drawing.Size(110, 22)
        Me.btnShowCursorInfo.TabIndex = 31
        Me.btnShowCursorInfo.Text = "Show Cursor Info"
        Me.btnShowCursorInfo.UseVisualStyleBackColor = True
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(252, 46)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(17, 13)
        Me.Label12.TabIndex = 36
        Me.Label12.Text = "Y:"
        '
        'txtYInterval
        '
        Me.txtYInterval.Location = New System.Drawing.Point(158, 42)
        Me.txtYInterval.Name = "txtYInterval"
        Me.txtYInterval.Size = New System.Drawing.Size(76, 20)
        Me.txtYInterval.TabIndex = 20
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(252, 19)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(17, 13)
        Me.Label11.TabIndex = 35
        Me.Label11.Text = "X:"
        '
        'txtXCursorPosn
        '
        Me.txtXCursorPosn.Location = New System.Drawing.Point(275, 16)
        Me.txtXCursorPosn.Name = "txtXCursorPosn"
        Me.txtXCursorPosn.Size = New System.Drawing.Size(76, 20)
        Me.txtXCursorPosn.TabIndex = 33
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.rbSelectRange)
        Me.GroupBox1.Controls.Add(Me.rbZoomRange)
        Me.GroupBox1.Controls.Add(Me.chkSelectYRange)
        Me.GroupBox1.Controls.Add(Me.chkSelectXRange)
        Me.GroupBox1.Location = New System.Drawing.Point(6, 216)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(127, 114)
        Me.GroupBox1.TabIndex = 37
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Chart Selection:"
        '
        'rbSelectRange
        '
        Me.rbSelectRange.AutoSize = True
        Me.rbSelectRange.Location = New System.Drawing.Point(6, 88)
        Me.rbSelectRange.Name = "rbSelectRange"
        Me.rbSelectRange.Size = New System.Drawing.Size(90, 17)
        Me.rbSelectRange.TabIndex = 24
        Me.rbSelectRange.TabStop = True
        Me.rbSelectRange.Text = "Select Range"
        Me.rbSelectRange.UseVisualStyleBackColor = True
        '
        'rbZoomRange
        '
        Me.rbZoomRange.AutoSize = True
        Me.rbZoomRange.Location = New System.Drawing.Point(6, 65)
        Me.rbZoomRange.Name = "rbZoomRange"
        Me.rbZoomRange.Size = New System.Drawing.Size(87, 17)
        Me.rbZoomRange.TabIndex = 23
        Me.rbZoomRange.TabStop = True
        Me.rbZoomRange.Text = "Zoom Range"
        Me.rbZoomRange.UseVisualStyleBackColor = True
        '
        'chkSelectYRange
        '
        Me.chkSelectYRange.AutoSize = True
        Me.chkSelectYRange.Location = New System.Drawing.Point(7, 42)
        Me.chkSelectYRange.Name = "chkSelectYRange"
        Me.chkSelectYRange.Size = New System.Drawing.Size(96, 17)
        Me.chkSelectYRange.TabIndex = 22
        Me.chkSelectYRange.Text = "Select Y range"
        Me.chkSelectYRange.UseVisualStyleBackColor = True
        '
        'chkSelectXRange
        '
        Me.chkSelectXRange.AutoSize = True
        Me.chkSelectXRange.Location = New System.Drawing.Point(6, 19)
        Me.chkSelectXRange.Name = "chkSelectXRange"
        Me.chkSelectXRange.Size = New System.Drawing.Size(96, 17)
        Me.chkSelectXRange.TabIndex = 21
        Me.chkSelectXRange.Text = "Select X range"
        Me.chkSelectXRange.UseVisualStyleBackColor = True
        '
        'chkAllowZoom
        '
        Me.chkAllowZoom.AutoSize = True
        Me.chkAllowZoom.Location = New System.Drawing.Point(6, 336)
        Me.chkAllowZoom.Name = "chkAllowZoom"
        Me.chkAllowZoom.Size = New System.Drawing.Size(81, 17)
        Me.chkAllowZoom.TabIndex = 32
        Me.chkAllowZoom.Text = "Allow Zoom"
        Me.chkAllowZoom.UseVisualStyleBackColor = True
        '
        'txtChartArea
        '
        Me.txtChartArea.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtChartArea.Location = New System.Drawing.Point(89, 33)
        Me.txtChartArea.Name = "txtChartArea"
        Me.txtChartArea.Size = New System.Drawing.Size(664, 20)
        Me.txtChartArea.TabIndex = 25
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(24, 36)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(59, 13)
        Me.Label8.TabIndex = 24
        Me.Label8.Text = "Chart area:"
        '
        'btnApply
        '
        Me.btnApply.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnApply.Location = New System.Drawing.Point(705, 58)
        Me.btnApply.Name = "btnApply"
        Me.btnApply.Size = New System.Drawing.Size(48, 22)
        Me.btnApply.TabIndex = 9
        Me.btnApply.Text = "Apply"
        Me.btnApply.UseVisualStyleBackColor = True
        '
        'txtToolTipString
        '
        Me.txtToolTipString.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtToolTipString.Location = New System.Drawing.Point(89, 59)
        Me.txtToolTipString.Name = "txtToolTipString"
        Me.txtToolTipString.Size = New System.Drawing.Size(610, 20)
        Me.txtToolTipString.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 62)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(77, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Tool Tip string:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(18, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(65, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Chart series:"
        '
        'cmbSeriesList
        '
        Me.cmbSeriesList.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbSeriesList.FormattingEnabled = True
        Me.cmbSeriesList.Location = New System.Drawing.Point(89, 6)
        Me.cmbSeriesList.Name = "cmbSeriesList"
        Me.cmbSeriesList.Size = New System.Drawing.Size(664, 21)
        Me.cmbSeriesList.TabIndex = 0
        '
        'TabPage2
        '
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(759, 569)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Other"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'btnZoomReset
        '
        Me.btnZoomReset.Location = New System.Drawing.Point(123, 12)
        Me.btnZoomReset.Name = "btnZoomReset"
        Me.btnZoomReset.Size = New System.Drawing.Size(83, 22)
        Me.btnZoomReset.TabIndex = 23
        Me.btnZoomReset.Text = "Zoom Reset"
        Me.btnZoomReset.UseVisualStyleBackColor = True
        '
        'rbSelect
        '
        Me.rbSelect.AutoSize = True
        Me.rbSelect.Location = New System.Drawing.Point(12, 12)
        Me.rbSelect.Name = "rbSelect"
        Me.rbSelect.Size = New System.Drawing.Size(55, 17)
        Me.rbSelect.TabIndex = 12
        Me.rbSelect.TabStop = True
        Me.rbSelect.Text = "Select"
        Me.rbSelect.UseVisualStyleBackColor = True
        '
        'rbPan
        '
        Me.rbPan.AutoSize = True
        Me.rbPan.Location = New System.Drawing.Point(73, 12)
        Me.rbPan.Name = "rbPan"
        Me.rbPan.Size = New System.Drawing.Size(44, 17)
        Me.rbPan.TabIndex = 13
        Me.rbPan.TabStop = True
        Me.rbPan.Text = "Pan"
        Me.rbPan.UseVisualStyleBackColor = True
        '
        'txtNSelPoints
        '
        Me.txtNSelPoints.Location = New System.Drawing.Point(6, 227)
        Me.txtNSelPoints.Name = "txtNSelPoints"
        Me.txtNSelPoints.Size = New System.Drawing.Size(145, 20)
        Me.txtNSelPoints.TabIndex = 67
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Location = New System.Drawing.Point(6, 211)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(110, 13)
        Me.Label23.TabIndex = 68
        Me.Label23.Text = "No of points selected:"
        '
        'frmChartInfoForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(791, 647)
        Me.Controls.Add(Me.rbPan)
        Me.Controls.Add(Me.rbSelect)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnZoomReset)
        Me.Name = "frmChartInfoForm"
        Me.Text = "Chart Information"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnExit As Button
    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents TabPage1 As TabPage
    Friend WithEvents Label1 As Label
    Friend WithEvents cmbSeriesList As ComboBox
    Friend WithEvents TabPage2 As TabPage
    Friend WithEvents btnApply As Button
    Friend WithEvents txtToolTipString As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents txtMouseY As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents txtMouseX As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents txtYInterval As TextBox
    Friend WithEvents Label7 As Label
    Friend WithEvents chkShowYCursor As CheckBox
    Friend WithEvents txtXInterval As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents chkShowXCursor As CheckBox
    Friend WithEvents chkSelectXRange As CheckBox
    Friend WithEvents chkSelectYRange As CheckBox
    Friend WithEvents btnZoomReset As Button
    Friend WithEvents txtChartArea As TextBox
    Friend WithEvents Label8 As Label
    Friend WithEvents txtSelPointY As TextBox
    Friend WithEvents Label9 As Label
    Friend WithEvents Label10 As Label
    Friend WithEvents txtSelPointX As TextBox
    Friend WithEvents btnShowCursorInfo As Button
    Friend WithEvents chkAllowZoom As CheckBox
    Friend WithEvents Label12 As Label
    Friend WithEvents Label11 As Label
    Friend WithEvents txtYCursorPosn As TextBox
    Friend WithEvents txtXCursorPosn As TextBox
    Friend WithEvents rbSelect As RadioButton
    Friend WithEvents rbPan As RadioButton
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents rbSelectRange As RadioButton
    Friend WithEvents rbZoomRange As RadioButton
    Friend WithEvents txtPointNo As TextBox
    Friend WithEvents Label13 As Label
    Friend WithEvents btnViewData As Button
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents GroupBox4 As GroupBox
    Friend WithEvents txtPixelY As TextBox
    Friend WithEvents Label15 As Label
    Friend WithEvents Label14 As Label
    Friend WithEvents txtPixelX As TextBox
    Friend WithEvents GroupBox5 As GroupBox
    Friend WithEvents cmbMarkerStyle As ComboBox
    Friend WithEvents txtMarkerSize As TextBox
    Friend WithEvents txtMarkerColor As TextBox
    Friend WithEvents txtBorderWidth As TextBox
    Friend WithEvents txtBorderColor As TextBox
    Friend WithEvents Label16 As Label
    Friend WithEvents Label18 As Label
    Friend WithEvents Label19 As Label
    Friend WithEvents Label20 As Label
    Friend WithEvents Label21 As Label
    Friend WithEvents Label17 As Label
    Friend WithEvents Label22 As Label
    Friend WithEvents btnShowRange As Button
    Friend WithEvents btnPointProperties As Button
    Friend WithEvents btnClearSelectedPoints As Button
    Friend WithEvents btnClearAllSelPoints As Button
    Friend WithEvents Label23 As Label
    Friend WithEvents txtNSelPoints As TextBox
End Class
