Public Class frmChartInfoForm
    'Chart Information Form - Displays additional information about a chart.

    'Links:
    'https://www.codeproject.com/Articles/1276879/MSChart-Extension-Zoom-and-Pan-Control-Version-2-2


#Region " Variable Declarations - All the variables used in this form and this application." '=================================================================================================

#End Region 'Variable Declarations ------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Properties - All the properties used in this form and this application" '============================================================================================================

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

#Region " Process XML files - Read and write XML files." '=====================================================================================================================================

    Private Sub SaveFormSettings()
        'Save the form settings in an XML document.
        Dim settingsData = <?xml version="1.0" encoding="utf-8"?>
                           <!---->
                           <FormSettings>
                               <Left><%= Me.Left %></Left>
                               <Top><%= Me.Top %></Top>
                               <Width><%= Me.Width %></Width>
                               <Height><%= Me.Height %></Height>
                               <!---->
                           </FormSettings>

        'Add code to include other settings to save after the comment line <!---->

        Dim SettingsFileName As String = "FormSettings_" & Main.ApplicationInfo.Name & "_" & Me.Text & ".xml"
        Main.Project.SaveXmlSettings(SettingsFileName, settingsData)
    End Sub


    Private Sub RestoreFormSettings()
        'Read the form settings from an XML document.

        Dim SettingsFileName As String = "FormSettings_" & Main.ApplicationInfo.Name & "_" & Me.Text & ".xml"

        If Main.Project.SettingsFileExists(SettingsFileName) Then
            Dim Settings As System.Xml.Linq.XDocument
            Main.Project.ReadXmlSettings(SettingsFileName, Settings)

            If IsNothing(Settings) Then 'There is no Settings XML data.
                Exit Sub
            End If

            'Restore form position and size:
            If Settings.<FormSettings>.<Left>.Value <> Nothing Then Me.Left = Settings.<FormSettings>.<Left>.Value
            If Settings.<FormSettings>.<Top>.Value <> Nothing Then Me.Top = Settings.<FormSettings>.<Top>.Value
            If Settings.<FormSettings>.<Height>.Value <> Nothing Then Me.Height = Settings.<FormSettings>.<Height>.Value
            If Settings.<FormSettings>.<Width>.Value <> Nothing Then Me.Width = Settings.<FormSettings>.<Width>.Value

            'Add code to read other saved setting here:

            CheckFormPos()
        End If
    End Sub

    Private Sub CheckFormPos()
        'Chech that the form can be seen on a screen.

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

    Protected Overrides Sub WndProc(ByRef m As Message) 'Save the form settings before the form is minimised:
        If m.Msg = &H112 Then 'SysCommand
            If m.WParam.ToInt32 = &HF020 Then 'Form is being minimised
                SaveFormSettings()
            End If
        End If
        MyBase.WndProc(m)
    End Sub

#End Region 'Process XML Files ----------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Form Display Methods - Code used to display this form." '============================================================================================================================

    Private Sub Form_Load(sender As Object, e As EventArgs) Handles Me.Load

        cmbMarkerStyle.Items.Clear()
        For Each item In [Enum].GetNames(GetType(DataVisualization.Charting.MarkerStyle))
            cmbMarkerStyle.Items.Add(item)
        Next

        'Read the marker properties from Main:
        txtBorderColor.BackColor = Main.SelPtBorderColor
        txtBorderWidth.Text = Main.SelPtBorderWidth
        txtMarkerColor.BackColor = Main.SelPtColor
        txtMarkerSize.Text = Main.SelPtSize
        cmbMarkerStyle.SelectedIndex = cmbMarkerStyle.FindStringExact(Main.SelPtStyle.ToString)

        RestoreFormSettings()   'Restore the form settings

        'Show a list of available series in the SeriesList:
        Dim SeriesCount As Integer = Main.Chart1.Series.Count
        Dim I As Integer
        For I = 1 To SeriesCount
            cmbSeriesList.Items.Add(Main.Chart1.Series(I - 1).Name)
        Next
        If cmbSeriesList.Items.Count > 0 Then
            If Main.ChartInfo.SelSeries = "" Then
                cmbSeriesList.SelectedIndex = 0 'Select the first series
            Else
                cmbSeriesList.SelectedIndex = cmbSeriesList.FindStringExact(Main.ChartInfo.SelSeries)
            End If
        End If

        'NOTE: DONT DO THIS UNTIL LATER! IT CHANGES SOME OF THE SETTINGS IN MAIN.CHART1
        'If Main.ChartMode = "Pan" Then
        '    rbPan.Checked = True
        'Else
        '    rbSelect.Checked = True
        'End If

        'Get form settings from Main.Chart1 properties:
        If Main.ChartInfo.SelChartArea = "" Then
            'No Chart Area selected.
        Else
            chkShowXCursor.Checked = Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorX.IsUserEnabled
            chkShowYCursor.Checked = Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorY.IsUserEnabled
            txtXInterval.Text = Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorX.Interval
            txtYInterval.Text = Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorY.Interval

            'Main.Message.Add("Main.Chart1.ChartAreas(" & SelChartArea & ").CursorX.IsUserSelectionEnabled: " & Main.Chart1.ChartAreas(SelChartArea).CursorX.IsUserSelectionEnabled & vbCrLf)
            chkSelectXRange.Checked = Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorX.IsUserSelectionEnabled
            'Main.Message.Add("Main.Chart1.ChartAreas(" & SelChartArea & ").CursorY.IsUserSelectionEnabled: " & Main.Chart1.ChartAreas(SelChartArea).CursorY.IsUserSelectionEnabled & vbCrLf)
            chkSelectYRange.Checked = Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorY.IsUserSelectionEnabled

            'Set rbZoomRange / rbSelectRange:
            If Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).AxisX.ScaleView.Zoomable Then
                If Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).AxisY.ScaleView.Zoomable Then
                    'AxisX and AxisY zoomable.
                    rbZoomRange.Checked = True
                Else
                    'Only Axis X zoomable.
                    rbZoomRange.Checked = True
                End If
            Else
                If Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).AxisY.ScaleView.Zoomable Then
                    'Only AxisY zoomable.
                    rbZoomRange.Checked = True
                Else
                    'No axis is zoomable.
                    'rbZoomRange.Checked = False
                    rbSelectRange.Checked = True
                End If
            End If

            ''Set chkSelectXRange & chkSelectYRange:
            'chkSelectXRange.Checked = Main.Chart1.ChartAreas(0).CursorX.IsUserSelectionEnabled
            'chkSelectYRange.Checked = Main.Chart1.ChartAreas(0).CursorY.IsUserSelectionEnabled

            txtToolTipString.Text = Main.Chart1.Series(cmbSeriesList.SelectedItem.ToString).ToolTip
        End If


        'If Main.ChartMode = "Pan" Then
        If Main.ChartInfo.Mode = "Pan" Then
            rbPan.Checked = True
        Else
            rbSelect.Checked = True
        End If

        txtNSelPoints.Text = Main.selPoints.Count



    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Exit the Form
        Me.Close() 'Close the form
    End Sub

    Private Sub Form_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If WindowState = FormWindowState.Normal Then
            SaveFormSettings()
        Else
            'Dont save settings if the form is minimised.
        End If
    End Sub

    Private Sub btnApply_Click(sender As Object, e As EventArgs) Handles btnApply.Click

        If Main.ChartInfo.SelChartArea = "" Then
            Main.Message.AddWarning("Select a chart area" & vbCrLf)
            Exit Sub
        End If

        'Apply the ToolTip text to Chart1
        If txtToolTipString.Text = "" Then
            Main.Message.AddWarning("No Tool Top string specified." & vbCrLf)
        ElseIf cmbSeriesList.SelectedIndex = -1 Then
            Main.Message.AddWarning("No chart series selected." & vbCrLf)
        Else
            'Main.Chart1.Series(cmbChartList.SelectedItem.ToString).ToolTip = txtToolTipString.Text
            'Main.Chart1.Series(SelChartArea).ToolTip = txtToolTipString.Text
            Main.Chart1.Series(cmbSeriesList.SelectedItem.ToString).ToolTip = txtToolTipString.Text
        End If

        'https://stackoverflow.com/questions/39157387/how-do-you-view-the-value-of-a-chart-point-on-mouse-over
        'https://web.archive.org/web/20160826032118/http://support2.dundas.com/Default.aspx?article=1132

    End Sub

    Private Sub chkShowXCursor_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowXCursor.CheckedChanged

        If Main.ChartInfo.SelChartArea = "" Then
            Main.Message.AddWarning("Select a chart area" & vbCrLf)
            Exit Sub
        End If

        If chkShowXCursor.Checked Then
            Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorX.IsUserEnabled = True
            Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorX.LineColor = Color.Red
            If txtXInterval.Text = "" Then
                txtXInterval.Text = "0.1"
                Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorX.Interval = 0.1
            Else
                Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorX.Interval = Val(txtXInterval.Text)
            End If
        Else
            Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorX.IsUserEnabled = False
            'Main.Chart1.ChartAreas(0).CursorX.SetCursorPosition(Double.NaN)  'Hide the cursor - buggy
            Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorX.LineColor = Color.Transparent 'This hides the line!
        End If
    End Sub

    Private Sub chkShowYCursor_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowYCursor.CheckedChanged

        If Main.ChartInfo.SelChartArea = "" Then
            Main.Message.AddWarning("Select a chart area" & vbCrLf)
            Exit Sub
        End If

        If chkShowYCursor.Checked Then
            Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorY.IsUserEnabled = True
            Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorY.LineColor = Color.Red
            If txtYInterval.Text = "" Then
                txtYInterval.Text = "0.1"
                Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorY.Interval = 0.1
            Else
                Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorY.Interval = Val(txtXInterval.Text)
            End If
        Else
            Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorY.IsUserEnabled = False
            'Main.Chart1.ChartAreas(0).CursorY.SetCursorPosition(Double.NaN) 'Hide the cursor - buggy
            Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorY.LineColor = Color.Transparent 'This hides the line!
        End If
    End Sub

    Private Sub chkSelectXRange_CheckedChanged(sender As Object, e As EventArgs) Handles chkSelectXRange.CheckedChanged

        If Main.ChartInfo.SelChartArea = "" Then
            Main.Message.AddWarning("Select a chart area" & vbCrLf)
            Exit Sub
        End If

        If chkSelectXRange.Checked Then
            Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorX.IsUserSelectionEnabled = True
        Else
            Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorX.IsUserSelectionEnabled = False
        End If
        'Main.Message.Add("chkSelectXRange.Checked: " & chkSelectXRange.Checked & vbCrLf)
        'Main.Message.Add("Main.Chart1.ChartAreas(" & SelChartArea & ").CursorX.IsUserSelectionEnabled: " & Main.Chart1.ChartAreas(SelChartArea).CursorX.IsUserSelectionEnabled & vbCrLf)

    End Sub

    Private Sub chkSelectYRange_CheckedChanged(sender As Object, e As EventArgs) Handles chkSelectYRange.CheckedChanged

        If Main.ChartInfo.SelChartArea = "" Then
            Main.Message.AddWarning("Select a chart area" & vbCrLf)
            Exit Sub
        End If

        If chkSelectYRange.Checked Then
            Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorY.IsUserSelectionEnabled = True
        Else
            Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorY.IsUserSelectionEnabled = False
        End If
    End Sub

    Private Sub btnResetChartRange_Click(sender As Object, e As EventArgs) Handles btnZoomReset.Click

        If Main.ChartInfo.SelChartArea = "" Then
            Main.Message.AddWarning("Select a chart area" & vbCrLf)
            Exit Sub
        End If

        'Reset the chart range
        Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).AxisX.ScaleView.ZoomReset()
        Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).AxisX2.ScaleView.ZoomReset()
        Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).AxisY.ScaleView.ZoomReset()
        Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).AxisY2.ScaleView.ZoomReset()


    End Sub

    Private Sub cmbChartList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbSeriesList.SelectedIndexChanged
        If cmbSeriesList.SelectedIndex = -1 Then
            'No item selected
            'SelSeries = ""
            Main.ChartInfo.SelSeries = ""
            txtChartArea.Text = ""
            'SelChartArea = ""
            Main.ChartInfo.SelChartArea = ""
            'UPDATE: SelSeries and SelChartArea properties are now only stored in Main.ChartInfo.SelSeries and Main.ChartInfo.SelChartArea
            'THE FOLLOWING CODE IS NOT REQUIRED:
            'If cmbSeriesList.Focused Then
            '    'Update the Main form:
            '    'Main.SelSeries = ""
            '    Main.ChartInfo.SelSeries = ""
            '    'Main.SelChartArea = ""
            '    Main.ChartInfo.SelChartArea = ""
            '    Main.Message.Add("Selected Series updated on the Main form." & vbCrLf)
            'End If
        Else
            'SelSeries = cmbSeriesList.SelectedItem.ToString
            Main.ChartInfo.SelSeries = cmbSeriesList.SelectedItem.ToString
            'SelChartArea = Main.Chart1.Series(cmbSeriesList.SelectedItem.ToString).ChartArea
            'SelChartArea = Main.Chart1.Series(SelSeries).ChartArea
            Main.ChartInfo.SelChartArea = Main.Chart1.Series(Main.ChartInfo.SelSeries).ChartArea
            txtChartArea.Text = Main.ChartInfo.SelChartArea
            'UPDATE: SelSeries and SelChartArea properties are now only stored in Main.ChartInfo.SelSeries and Main.ChartInfo.SelChartArea
            'THE FOLLOWING CODE IS NOT REQUIRED:
            'If cmbSeriesList.Focused Then
            '    'Update the Main form:
            '    Main.ChartInfo.SelSeries = SelSeries
            '    Main.ChartInfo.SelChartArea = SelChartArea
            '    Main.Message.Add("Selected Series updated on the Main form." & vbCrLf)
            'End If
        End If
    End Sub



    Private Sub btnShowCursorInfo_Click(sender As Object, e As EventArgs) Handles btnShowCursorInfo.Click

        'If SelChartArea = "" Then
        If Main.ChartInfo.SelChartArea = "" Then
            Main.Message.AddWarning("Select a chart area" & vbCrLf)
            Exit Sub
        End If

        'Show the cursor info:
        Main.Message.Add("Main.Chart1.ChartAreas(" & Main.ChartInfo.SelChartArea & ").CursorX.IsUserEnabled: " & Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorX.IsUserEnabled & vbCrLf)
        Main.Message.Add("Main.Chart1.ChartAreas(" & Main.ChartInfo.SelChartArea & ").CursorX.IsUserSelectionEnabled: " & Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorX.IsUserSelectionEnabled & vbCrLf)
        Main.Message.Add("Main.Chart1.ChartAreas(" & Main.ChartInfo.SelChartArea & ").CursorX.LineColor.Name: " & Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorX.LineColor.Name & vbCrLf)
        Main.Message.Add("Main.Chart1.ChartAreas(" & Main.ChartInfo.SelChartArea & ").CursorX.LineWidth: " & Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorX.LineWidth & vbCrLf)
        Main.Message.Add("Main.Chart1.ChartAreas(" & Main.ChartInfo.SelChartArea & ").CursorX.Position: " & Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorX.Position & vbCrLf)
        Main.Message.Add("Main.Chart1.ChartAreas(" & Main.ChartInfo.SelChartArea & ").CursorX.Interval: " & Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorX.Interval & vbCrLf)
        Main.Message.Add("Main.Chart1.ChartAreas(" & Main.ChartInfo.SelChartArea & ").CursorX.SelectionStart: " & Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorX.SelectionStart & vbCrLf)
        Main.Message.Add(vbCrLf)

    End Sub

    Private Sub chkAllowZoom_CheckedChanged(sender As Object, e As EventArgs) Handles chkAllowZoom.CheckedChanged

        'If SelChartArea = "" Then
        If Main.ChartInfo.SelChartArea = "" Then
            Main.Message.AddWarning("Select a chart area" & vbCrLf)
            Exit Sub
        End If

        If chkAllowZoom.Checked Then
            'Enable zooming of the chart:
            Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).AxisX.ScaleView.Zoomable = True
            Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).AxisX2.ScaleView.Zoomable = True
            Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).AxisY.ScaleView.Zoomable = True
            Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).AxisY2.ScaleView.Zoomable = True
        Else
            'Disable zooming of the chart:
            Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).AxisX.ScaleView.Zoomable = False
            Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).AxisX2.ScaleView.Zoomable = False
            Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).AxisY.ScaleView.Zoomable = False
            Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).AxisY2.ScaleView.Zoomable = False
        End If
    End Sub

    Private Sub rbSelect_CheckedChanged(sender As Object, e As EventArgs) Handles rbSelect.CheckedChanged
        If rbSelect.Checked Then
            Main.ChartInfo.Mode = "Select"
            chkShowXCursor.Enabled = True 'Chart cursor can be used.
            chkShowYCursor.Enabled = True 'Chart cursor can be used.
            chkSelectXRange.Enabled = True 'Range selection can be used
            chkSelectYRange.Enabled = True 'Range selection can be used
            ApplyChartSettings()
        End If
    End Sub

    Private Sub rbPan_CheckedChanged(sender As Object, e As EventArgs) Handles rbPan.CheckedChanged
        If rbPan.Checked Then
            Main.ChartInfo.Mode = "Pan"
            chkShowXCursor.Enabled = False 'Chart cursor cannot be used while panning.
            chkShowYCursor.Enabled = False 'Chart cursor cannot be used while panning.
            chkSelectXRange.Enabled = False 'Range selection cannot be used while panning.
            chkSelectYRange.Enabled = False 'Range selection cannot be used while panning.
            ApplyChartSettings()
        End If
    End Sub

    Private Sub ApplyChartSettings()
        'Apply the selected settings to the chart.

        'If SelChartArea = "" Then
        If Main.ChartInfo.SelChartArea = "" Then
            Main.Message.AddWarning("Select a chart area" & vbCrLf)
            Exit Sub
        End If

        'Cursor X:
        If chkShowXCursor.Enabled = True Then
            If chkShowXCursor.Checked = True Then
                Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorX.IsUserEnabled = True 'Ensable the cursor
                Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorX.LineColor = Color.Red 'and show it.
            Else
                Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorX.IsUserEnabled = False 'Disable the cursor
                Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorX.LineColor = Color.Transparent 'and dont show it.
            End If
        Else
            If chkShowXCursor.Checked = True Then
                Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorX.IsUserEnabled = False 'Disable the cursor
                Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorX.LineColor = Color.Red 'but still show it.
            Else
                Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorX.IsUserEnabled = False 'Disable the cursor
                Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorX.LineColor = Color.Transparent 'and dont show it.
            End If
        End If

        'Cursor Y:
        If chkShowYCursor.Enabled = True Then
            If chkShowYCursor.Checked = True Then
                Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorY.IsUserEnabled = True 'Ensable the cursor
                Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorY.LineColor = Color.Red 'and show it.
            Else
                Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorY.IsUserEnabled = False 'Disable the cursor
                Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorY.LineColor = Color.Transparent 'and dont show it.
            End If
        Else
            If chkShowYCursor.Checked = True Then
                Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorY.IsUserEnabled = False 'Disable the cursor
                Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorY.LineColor = Color.Red 'but still show it.
            Else
                Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorY.IsUserEnabled = False 'Disable the cursor
                Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorY.LineColor = Color.Transparent 'and dont show it.
            End If
        End If

        'Select X Range:
        If chkSelectXRange.Enabled = True Then
            If chkSelectXRange.Checked Then
                Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorX.IsUserSelectionEnabled = True
            Else
                Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorX.IsUserSelectionEnabled = False
            End If
        Else
            If chkSelectXRange.Checked Then
                Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorX.IsUserSelectionEnabled = False
            Else
                Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorX.IsUserSelectionEnabled = False
            End If
        End If

        'Select Y Range:
        If chkSelectYRange.Enabled = True Then
            If chkSelectYRange.Checked Then
                Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorY.IsUserSelectionEnabled = True
            Else
                Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorY.IsUserSelectionEnabled = False
            End If
        Else
            If chkSelectYRange.Checked Then
                Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorY.IsUserSelectionEnabled = False
            Else
                Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorY.IsUserSelectionEnabled = False
            End If
        End If


    End Sub

    Private Sub rbZoomRange_CheckedChanged(sender As Object, e As EventArgs) Handles rbZoomRange.CheckedChanged

        If Main.ChartInfo.SelChartArea = "" Then
            Main.Message.AddWarning("Select a chart area" & vbCrLf)
            Exit Sub
        End If

        If rbZoomRange.Checked Then
            If chkSelectXRange.Checked Then
                Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).AxisX.ScaleView.Zoomable = True
            Else
                Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).AxisX.ScaleView.Zoomable = False
            End If
            If chkSelectYRange.Checked Then
                Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).AxisY.ScaleView.Zoomable = True
            Else
                Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).AxisY.ScaleView.Zoomable = False
            End If
        End If

    End Sub

    Private Sub rbSelectRange_CheckedChanged(sender As Object, e As EventArgs) Handles rbSelectRange.CheckedChanged

        If Main.ChartInfo.SelChartArea = "" Then
            Main.Message.AddWarning("Select a chart area" & vbCrLf)
            Exit Sub
        End If

        If rbSelectRange.Checked Then
            Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).AxisX.ScaleView.Zoomable = False
            Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).AxisY.ScaleView.Zoomable = False
        End If
    End Sub

    Private Sub btnViewData_Click(sender As Object, e As EventArgs) Handles btnViewData.Click
        'Open the View Database Data form:
        If IsNothing(Main.ViewDatabaseData) Then
            Main.ViewDatabaseData = New frmViewDatabaseData
            Main.ViewDatabaseData.Show()
            Main.ViewDatabaseData.Update()
        Else
            Main.ViewDatabaseData.Show()
            Main.ViewDatabaseData.Update()
        End If
    End Sub



    Private Sub txtXInterval_LostFocus(sender As Object, e As EventArgs) Handles txtXInterval.LostFocus
        txtXInterval.Text = Val(txtXInterval.Text)
        If Main.ChartInfo.SelChartArea = "" Then
            Main.Message.AddWarning("Select a chart area" & vbCrLf)
            Exit Sub
        End If
        Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorX.Interval = Val(txtXInterval.Text)

    End Sub



    Private Sub txtYInterval_LostFocus(sender As Object, e As EventArgs) Handles txtYInterval.LostFocus
        txtYInterval.Text = Val(txtYInterval.Text)
        If Main.ChartInfo.SelChartArea = "" Then
            Main.Message.AddWarning("Select a chart area" & vbCrLf)
            Exit Sub
        End If
        Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorY.Interval = Val(txtYInterval.Text)
    End Sub

    Private Sub txtToolTipString_TextChanged(sender As Object, e As EventArgs) Handles txtToolTipString.TextChanged

    End Sub

    Private Sub txtXInterval_TextChanged(sender As Object, e As EventArgs) Handles txtXInterval.TextChanged

    End Sub

    Private Sub btnShowRange_Click(sender As Object, e As EventArgs) Handles btnShowRange.Click
        'Show the selected Range:

        If Main.ChartInfo.SelChartArea = "" Then
            Exit Sub
        End If

        'Main.Message.Add("Main.Chart1.ChartAreas(" & SelChartArea & ").AxisX.ScaleView.Position: " & Main.Chart1.ChartAreas(SelChartArea).AxisX.ScaleView.Position & vbCrLf)

        'Main.Message.Add("Main.Chart1.ChartAreas(" & SelChartArea & ").AxisX.ScaleView.ViewMinimum: " & Main.Chart1.ChartAreas(SelChartArea).AxisX.ScaleView.ViewMinimum & vbCrLf)
        'Main.Message.Add("Main.Chart1.ChartAreas(" & SelChartArea & ").AxisX.ScaleView.ViewMaximum: " & Main.Chart1.ChartAreas(SelChartArea).AxisX.ScaleView.ViewMaximum & vbCrLf)
        'Main.Message.Add("Main.Chart1.ChartAreas(" & SelChartArea & ").AxisY.ScaleView.ViewMinimum: " & Main.Chart1.ChartAreas(SelChartArea).AxisY.ScaleView.ViewMinimum & vbCrLf)
        'Main.Message.Add("Main.Chart1.ChartAreas(" & SelChartArea & ").AxisY.ScaleView.ViewMaximum: " & Main.Chart1.ChartAreas(SelChartArea).AxisY.ScaleView.ViewMaximum & vbCrLf)

        Main.Message.Add("Main.Chart1.ChartAreas(" & Main.ChartInfo.SelChartArea & ").CursorX.SelectionStart: " & Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorX.SelectionStart & vbCrLf)
        Main.Message.Add("Main.Chart1.ChartAreas(" & Main.ChartInfo.SelChartArea & ").CursorX.SelectionEnd: " & Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorX.SelectionEnd & vbCrLf)
        Main.Message.Add("Main.Chart1.ChartAreas(" & Main.ChartInfo.SelChartArea & ").CursorY.SelectionStart: " & Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorY.SelectionStart & vbCrLf)
        Main.Message.Add("Main.Chart1.ChartAreas(" & Main.ChartInfo.SelChartArea & ").CursorY.SelectionEnd: " & Main.Chart1.ChartAreas(Main.ChartInfo.SelChartArea).CursorY.SelectionEnd & vbCrLf)

        'Main.Chart1.Series.Remove(Main.Chart1.Series(0)) 'Remove a series




    End Sub



    Private Sub btnClearSelectedPoints_Click(sender As Object, e As EventArgs) Handles btnClearSelectedPoints.Click
        'Clear the selected points:

        'Main.selPoints.Clear()
        'Main.Chart1.Series(Main.ChartInfo.SelSeries).MarkerColor = Color.BlueViolet
        'Main.Chart1.Series(Main.ChartInfo.SelSeries).MarkerSize = 10
        'Main.Chart1.Update()

        Dim PointNo As Integer = Val(txtPointNo.Text)
        Main.Chart1.Series(Main.ChartInfo.SelSeries).Points(PointNo).SetDefault(True) 'This resets a single point.
        'Dim IndexNo As Integer = Main.selPoints.Find(Function(value As DataVisualization.Charting.DataPoint) Return value. = 0 End Function)
        'Main.selPoints.Find(Function(value As DataVisualization.Charting.DataPoint) Return value.Index = 0)

        Dim SelectedPoint As DataVisualization.Charting.DataPoint = Main.Chart1.Series(Main.ChartInfo.SelSeries).Points(PointNo)
        Main.selPoints.Remove(SelectedPoint)

        txtNSelPoints.Text = Main.selPoints.Count

        'Main.Chart1.Update() 'This does not redraw the chart!
        'Main.Chart1.UpdateAnnotations() 'This does not redraw the chart!
        'Main.Chart1.Series(Main.ChartInfo.SelSeries).SetDefault(True)
        'Main.Chart1.Show() 'This does not redraw the chart!
        Main.Chart1.Invalidate() 'Redraws the Chart control


    End Sub

    Private Sub btnPointProperties_Click(sender As Object, e As EventArgs) Handles btnPointProperties.Click
        'Display the selected point properties:
        Try
            Dim PointNo As Integer = Val(txtPointNo.Text)
            Main.Message.Add(vbCrLf & "Information for selected point number: " & PointNo & vbCrLf)
            Main.Message.Add("MarkerBorderColor: " & Main.Chart1.Series(Main.ChartInfo.SelSeries).Points(PointNo).MarkerBorderColor.ToKnownColor.ToString & vbCrLf)
            Main.Message.Add("MarkerSize: " & Main.Chart1.Series(Main.ChartInfo.SelSeries).Points(PointNo).MarkerSize & vbCrLf)
            'Main.Message.Add("Item(0).ToString: " & Main.Chart1.Series(Main.ChartInfo.SelSeries).Points(PointNo).Item(0).ToString & vbCrLf)
            'Main.Message.Add("IsCustomPropertySet(0).ToString: " & Main.Chart1.Series(Main.ChartInfo.SelSeries).Points(PointNo).IsCustomPropertySet(0).ToString & vbCrLf)

            'Main.Chart1.Series(Main.ChartInfo.SelSeries).Points(PointNo).SetDefault(True)

        Catch ex As Exception
            Main.Message.AddWarning("Error: " & vbCrLf & ex.Message & vbCrLf)
        End Try



    End Sub

    Private Sub btnClearAllSelPoints_Click(sender As Object, e As EventArgs) Handles btnClearAllSelPoints.Click
        'Clear all the selected points.

        For Each Item In Main.selPoints
            Item.SetDefault(True)
        Next
        Main.selPoints.Clear()
        Main.Chart1.Invalidate() 'Redraws the Chart control
        txtNSelPoints.Text = Main.selPoints.Count
    End Sub







#End Region 'Form Display Methods -------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Open and Close Forms - Code used to open and close other forms." '===================================================================================================================

#End Region 'Open and Close Forms -------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Form Methods - The main actions performed by this form." '===========================================================================================================================

#End Region 'Form Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------




End Class