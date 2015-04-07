Imports System.IO
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Runtime.Serialization
<Serializable()> _
Public Class Main_form
    Inherits System.Windows.Forms.Form

    ' Notify icon 

    Friend WithEvents notifyIcon1 As System.Windows.Forms.NotifyIcon

    Public My_icon As Icon = My.Resources.Logo_point

    Public Icon_green As Icon = My.Resources.Logo_point_green

    Public Icon_red As Icon = My.Resources.Logo_point_red

    Public contextMenu1 As New System.Windows.Forms.ContextMenu

    Public WithEvents menuItem1 As New System.Windows.Forms.MenuItem

    ' end of notify icon

    Public App_Name As String = "T.H.O.R.I.U.M."

    Public view As Visualisation.Visualisation

    Public Ready As Boolean = False

    Public Loaded As Boolean = False

    Public Current_step As Long = 0

    Public Model_for_visualisation_number = 1

    ''' <summary>
    ''' Model visualisation mode.
    ''' Can be: Wireframe, With contour, Witout contour. 
    ''' </summary>
    ''' <remarks></remarks>
    Public V_mode As New Visualisation.Mode

    Public FEM As Model

    Public WithEvents Solver As TOP_solver


    ' Visualisation fields

    ''' <summary>
    ''' The Min_T and Max_T are taken from current step
    ''' </summary>
    ''' <remarks></remarks>
    Public Min_T_Max_T_from_current_step As Boolean = False

    ''' <summary>
    '''Application visualisation mode. Can be "NORMAL" or "HIDDEN"
    ''' </summary>
    ''' <remarks></remarks>
    Public Mode As String

    ''' <summary>
    ''' Computer identificator. Allows to save same file in one dir. Random 8-byte number.
    ''' </summary>
    ''' <remarks></remarks>
    Public ID As Long

    ''' <summary>
    ''' File name of input model. *.bdf or *.model file.
    ''' </summary>
    ''' <remarks></remarks>
    Public Input_filename As String

    ''' <summary>
    ''' File name of output model, *.model file.
    ''' </summary>
    ''' <remarks></remarks>
    Public Output_filename As String

    ''' <summary>
    ''' Name of file for saving text results.
    ''' </summary>
    ''' <remarks></remarks>
    Public Output_txt_filename As String

    ''' <summary>
    ''' Should we use input *.model file for the analysis resumption or not?
    ''' </summary>
    ''' <remarks></remarks>
    Public Resumption As Boolean = False

    ''' <summary>
    ''' If we should use input *.model file for the analysis resumption, then which moment of time should we use to begin?
    ''' If Resumption_start_time  = Double.MaxValue then this parameter is not set and we should resume analysis 
    ''' from the last but one (предпоследний) step.
    ''' If Resumption_start_time is not Double.MaxValue then we should find step before given Resumption_start_time
    ''' and we have to begin analysis from this step.
    ''' </summary>
    ''' <remarks></remarks>
    Public Resumption_start_time As Double = Double.MaxValue

    ''' <summary>
    ''' The solution finish time, i.e. heat transfer problem will be solved for time interval from Resumption_start_time to Resumption_start_time + Resumption_length_time.
    ''' If Resumption_length_time = Double.MaxValue then this parameter is not set and we should perfome analysis
    ''' until time set in *.model file.
    ''' If Resumption_time is not Double.MaxValue then we should perfome analysis until set time
    ''' </summary>
    ''' <remarks></remarks>
    Public Resumption_length_time As Double = Double.MaxValue

    ''' <summary>
    ''' Path to directiry for temprorary file storing
    ''' </summary>
    ''' <remarks></remarks>
    Public tmp_Directory As String = System.Environment.GetEnvironmentVariable("temp") & "\"

    ''' <summary>
    ''' Whether Work_directory is on FTP server?
    ''' </summary>
    ''' <remarks></remarks>
    Public FTP_mode As Boolean

    ''' <summary>
    ''' Hierarchy of directories on FTP server. For example, ftp.narod.ru/www/111/222/
    ''' </summary>
    ''' <remarks></remarks>
    Public FTP_Dir_Hierarchy() As String


    Private Const Start_Height As Integer = 50

    Private Const Start_Width As Integer = 200


    Private Const Progress_Height As Integer = 50

    Private Const Progress_Width As Integer = 800


    Private Const Result_view_Height As Integer = 120

    Private Const Result_view_Width As Integer = 800



#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        Me.components = New System.ComponentModel.Container

        Me.contextMenu1 = New System.Windows.Forms.ContextMenu
        Me.menuItem1 = New System.Windows.Forms.MenuItem

        ' Initialize contextMenu1
        Me.contextMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() _
                            {Me.menuItem1})

        ' Initialize menuItem1
        Me.menuItem1.Index = 0
        Me.menuItem1.Text = "E&xit"


        ' Create the NotifyIcon.
        Me.notifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)

        ' The Icon property sets the icon that will appear
        ' in the systray for this application.
        notifyIcon1.Icon = Me.My_icon


        ' The ContextMenu property sets the menu that will
        ' appear when the systray icon is right clicked.
        notifyIcon1.ContextMenu = Me.contextMenu1

        ' The Text property sets the text that will be displayed,
        ' in a tooltip, when the mouse hovers over the systray icon.
        notifyIcon1.Text = App_Name
        notifyIcon1.Visible = True

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose( _
                                 ByVal disposing As Boolean)

        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)

    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required 
    'by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Main_form))
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.OpenToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.SaveToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.Save_results_as_txt = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.RunToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ThermalAnalysisToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.Solution_Progressbar = New System.Windows.Forms.ProgressBar
        Me.Step_selector = New System.Windows.Forms.ComboBox
        Me.Value_strip_status = New System.Windows.Forms.CheckBox
        Me.Animation = New System.Windows.Forms.Button
        Me.Next_step = New System.Windows.Forms.Button
        Me.To_last_step = New System.Windows.Forms.Button
        Me.To_first_step = New System.Windows.Forms.Button
        Me.Previous_step = New System.Windows.Forms.Button
        Me.Stop_button = New System.Windows.Forms.Button
        Me.Decimate_results = New System.Windows.Forms.Button
        Me.Colors_from_step = New System.Windows.Forms.RadioButton
        Me.Colors_from_entire_solution = New System.Windows.Forms.RadioButton
        Me.Visual_mode = New System.Windows.Forms.ComboBox
        Me.Show_model = New System.Windows.Forms.Button
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.RunToolStripMenuItem, Me.HelpToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(794, 24)
        Me.MenuStrip1.TabIndex = 4
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.OpenToolStripMenuItem, Me.SaveToolStripMenuItem, Me.Save_results_as_txt, Me.ToolStripSeparator3, Me.ExitToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.FileToolStripMenuItem.Text = "File"
        '
        'OpenToolStripMenuItem
        '
        Me.OpenToolStripMenuItem.AutoSize = False
        Me.OpenToolStripMenuItem.Name = "OpenToolStripMenuItem"
        Me.OpenToolStripMenuItem.Size = New System.Drawing.Size(184, 22)
        Me.OpenToolStripMenuItem.Text = "Open"
        '
        'SaveToolStripMenuItem
        '
        Me.SaveToolStripMenuItem.Name = "SaveToolStripMenuItem"
        Me.SaveToolStripMenuItem.Size = New System.Drawing.Size(165, 22)
        Me.SaveToolStripMenuItem.Text = "Save"
        '
        'Save_results_as_txt
        '
        Me.Save_results_as_txt.Name = "Save_results_as_txt"
        Me.Save_results_as_txt.Size = New System.Drawing.Size(165, 22)
        Me.Save_results_as_txt.Text = "Save results as txt"
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(162, 6)
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(165, 22)
        Me.ExitToolStripMenuItem.Text = "Exit"
        '
        'RunToolStripMenuItem
        '
        Me.RunToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ThermalAnalysisToolStripMenuItem})
        Me.RunToolStripMenuItem.Name = "RunToolStripMenuItem"
        Me.RunToolStripMenuItem.Size = New System.Drawing.Size(62, 20)
        Me.RunToolStripMenuItem.Text = "Analysis"
        '
        'ThermalAnalysisToolStripMenuItem
        '
        Me.ThermalAnalysisToolStripMenuItem.Name = "ThermalAnalysisToolStripMenuItem"
        Me.ThermalAnalysisToolStripMenuItem.Size = New System.Drawing.Size(162, 22)
        Me.ThermalAnalysisToolStripMenuItem.Text = "Thermal analysis"
        '
        'HelpToolStripMenuItem
        '
        Me.HelpToolStripMenuItem.Name = "HelpToolStripMenuItem"
        Me.HelpToolStripMenuItem.Size = New System.Drawing.Size(44, 20)
        Me.HelpToolStripMenuItem.Text = "Help"
        '
        'Solution_Progressbar
        '
        Me.Solution_Progressbar.Location = New System.Drawing.Point(148, 7)
        Me.Solution_Progressbar.Name = "Solution_Progressbar"
        Me.Solution_Progressbar.Size = New System.Drawing.Size(637, 10)
        Me.Solution_Progressbar.TabIndex = 5
        Me.Solution_Progressbar.Visible = False
        '
        'Step_selector
        '
        Me.Step_selector.FormattingEnabled = True
        Me.Step_selector.Location = New System.Drawing.Point(127, 65)
        Me.Step_selector.Name = "Step_selector"
        Me.Step_selector.Size = New System.Drawing.Size(159, 21)
        Me.Step_selector.TabIndex = 12
        '
        'Value_strip_status
        '
        Me.Value_strip_status.AutoSize = True
        Me.Value_strip_status.Location = New System.Drawing.Point(421, 71)
        Me.Value_strip_status.Name = "Value_strip_status"
        Me.Value_strip_status.Size = New System.Drawing.Size(75, 17)
        Me.Value_strip_status.TabIndex = 11
        Me.Value_strip_status.Text = "Value strip"
        Me.Value_strip_status.UseVisualStyleBackColor = True
        '
        'Animation
        '
        Me.Animation.Location = New System.Drawing.Point(110, 36)
        Me.Animation.Name = "Animation"
        Me.Animation.Size = New System.Drawing.Size(95, 23)
        Me.Animation.TabIndex = 10
        Me.Animation.Text = "Animate!"
        Me.Animation.UseVisualStyleBackColor = True
        '
        'Next_step
        '
        Me.Next_step.Location = New System.Drawing.Point(306, 36)
        Me.Next_step.Name = "Next_step"
        Me.Next_step.Size = New System.Drawing.Size(95, 23)
        Me.Next_step.TabIndex = 9
        Me.Next_step.Text = "Next step >"
        Me.Next_step.UseVisualStyleBackColor = True
        '
        'To_last_step
        '
        Me.To_last_step.Location = New System.Drawing.Point(306, 64)
        Me.To_last_step.Name = "To_last_step"
        Me.To_last_step.Size = New System.Drawing.Size(95, 23)
        Me.To_last_step.TabIndex = 7
        Me.To_last_step.Text = "To last step >>"
        Me.To_last_step.UseVisualStyleBackColor = True
        '
        'To_first_step
        '
        Me.To_first_step.Location = New System.Drawing.Point(12, 65)
        Me.To_first_step.Name = "To_first_step"
        Me.To_first_step.Size = New System.Drawing.Size(95, 23)
        Me.To_first_step.TabIndex = 8
        Me.To_first_step.Text = "<< To first step"
        Me.To_first_step.UseVisualStyleBackColor = True
        '
        'Previous_step
        '
        Me.Previous_step.Location = New System.Drawing.Point(12, 36)
        Me.Previous_step.Name = "Previous_step"
        Me.Previous_step.Size = New System.Drawing.Size(95, 23)
        Me.Previous_step.TabIndex = 6
        Me.Previous_step.Text = "< Previous step"
        Me.Previous_step.UseVisualStyleBackColor = True
        '
        'Stop_button
        '
        Me.Stop_button.Location = New System.Drawing.Point(209, 36)
        Me.Stop_button.Name = "Stop_button"
        Me.Stop_button.Size = New System.Drawing.Size(95, 23)
        Me.Stop_button.TabIndex = 10
        Me.Stop_button.Text = "Stop!"
        Me.Stop_button.UseVisualStyleBackColor = True
        '
        'Decimate_results
        '
        Me.Decimate_results.Location = New System.Drawing.Point(421, 36)
        Me.Decimate_results.Name = "Decimate_results"
        Me.Decimate_results.Size = New System.Drawing.Size(95, 23)
        Me.Decimate_results.TabIndex = 10
        Me.Decimate_results.Text = "Decimate results"
        Me.Decimate_results.UseVisualStyleBackColor = True
        '
        'Colors_from_step
        '
        Me.Colors_from_step.AutoSize = True
        Me.Colors_from_step.Location = New System.Drawing.Point(501, 71)
        Me.Colors_from_step.Name = "Colors_from_step"
        Me.Colors_from_step.Size = New System.Drawing.Size(136, 17)
        Me.Colors_from_step.TabIndex = 13
        Me.Colors_from_step.Text = "Colors from current step"
        Me.Colors_from_step.UseVisualStyleBackColor = True
        '
        'Colors_from_entire_solution
        '
        Me.Colors_from_entire_solution.AutoSize = True
        Me.Colors_from_entire_solution.Location = New System.Drawing.Point(635, 71)
        Me.Colors_from_entire_solution.Name = "Colors_from_entire_solution"
        Me.Colors_from_entire_solution.Size = New System.Drawing.Size(145, 17)
        Me.Colors_from_entire_solution.TabIndex = 13
        Me.Colors_from_entire_solution.Text = "Colors from entire solution"
        Me.Colors_from_entire_solution.UseVisualStyleBackColor = True
        '
        'Visual_mode
        '
        Me.Visual_mode.FormattingEnabled = True
        Me.Visual_mode.Location = New System.Drawing.Point(616, 36)
        Me.Visual_mode.Name = "Visual_mode"
        Me.Visual_mode.Size = New System.Drawing.Size(159, 21)
        Me.Visual_mode.TabIndex = 12
        '
        'Show_model
        '
        Me.Show_model.Location = New System.Drawing.Point(527, 36)
        Me.Show_model.Name = "Show_model"
        Me.Show_model.Size = New System.Drawing.Size(74, 23)
        Me.Show_model.TabIndex = 10
        Me.Show_model.Text = "Show model"
        Me.Show_model.UseVisualStyleBackColor = True
        '
        'Main_form
        '
        Me.ClientSize = New System.Drawing.Size(794, 94)
        Me.Controls.Add(Me.Colors_from_entire_solution)
        Me.Controls.Add(Me.Colors_from_step)
        Me.Controls.Add(Me.Visual_mode)
        Me.Controls.Add(Me.Step_selector)
        Me.Controls.Add(Me.Value_strip_status)
        Me.Controls.Add(Me.Show_model)
        Me.Controls.Add(Me.Decimate_results)
        Me.Controls.Add(Me.Stop_button)
        Me.Controls.Add(Me.Animation)
        Me.Controls.Add(Me.Next_step)
        Me.Controls.Add(Me.To_last_step)
        Me.Controls.Add(Me.To_first_step)
        Me.Controls.Add(Me.Previous_step)
        Me.Controls.Add(Me.Solution_Progressbar)
        Me.Controls.Add(Me.MenuStrip1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.MaximumSize = New System.Drawing.Size(800, 300)
        Me.Name = "Main_form"
        Me.Text = "T.H.O.R.I.U.M."
        Me.TopMost = True
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Me.Height = Start_Height

        Me.Width = Start_Width

        Me.Hide()

        ' Initializing

        'ID generation

        Randomize()

        ID = 100000000 * Rnd()

        Me.Text = App_Name

        Solution_Progressbar.Visible = False

        ThermalAnalysisToolStripMenuItem.Enabled = False

        SaveToolStripMenuItem.Enabled = False

        Save_results_as_txt.Enabled = False

        Me.FEM = New Model

        Me.Solver = New TOP_solver(FEM)

        Me.V_mode.Status = Visualisation.Mode.WITH_CONTOUR

        'Me.view.frm.Hide()

        Me.Colors_from_entire_solution.Checked = True

        ' Visialisation modes selector

        Visual_mode.Hide()

        Visual_mode.DropDownStyle = ComboBoxStyle.DropDownList

        Visual_mode.Items.Add("With contour")

        Visual_mode.Items.Add("Without contour")

        Visual_mode.Items.Add("Wireframe")

        Visual_mode.SelectedIndex = 0

        Visual_mode.Show()


        Me.Loaded = True

        Me.Ready = True



        'Me.BackColor = Color.White

        Me.Refresh()

        'Me.view.frm.Hide()


        ' Reading command line

        Dim Arg As String

        For i As Integer = 0 To My.Application.CommandLineArgs.Count - 1

            Do

                Arg = My.Application.CommandLineArgs(i)

                If Arg.ToUpper.StartsWith("MODE=") Then

                    Me.Mode = Arg.Remove(0, "MODE=".Length).ToUpper

                End If


                If Arg.ToUpper.StartsWith("INPUT=") Then

                    Me.Input_filename = Arg.Remove(0, "INPUT=".Length).ToUpper

                    For j As Integer = i + 1 To My.Application.CommandLineArgs.Count - 1

                        Arg = My.Application.CommandLineArgs(j)

                        If InStr(Arg, "=") > 0 Then Exit Do

                        Me.Input_filename &= " " & Arg

                    Next j

                End If

                If Arg.ToUpper.StartsWith("OUTPUT=") Then

                    Me.Output_filename = Arg.Remove(0, "OUTPUT=".Length).ToUpper

                    For j As Integer = i + 1 To My.Application.CommandLineArgs.Count - 1

                        Arg = My.Application.CommandLineArgs(j)

                        If InStr(Arg, "=") > 0 Then Exit Do

                        Me.Output_filename &= " " & Arg

                    Next j

                End If

                If Arg.ToUpper.StartsWith("OUTPUT_TXT=") Then

                    Me.Output_txt_filename = Arg.Remove(0, "OUTPUT_TXT=".Length).ToUpper

                    For j As Integer = i + 1 To My.Application.CommandLineArgs.Count - 1

                        Arg = My.Application.CommandLineArgs(j)

                        If InStr(Arg, "=") > 0 Then Exit Do

                        Me.Output_txt_filename &= " " & Arg

                    Next j

                End If


                If Arg.ToUpper.StartsWith("RESUMPTION") Then

                    Me.Resumption = True

                    If Arg.ToUpper.StartsWith("RESUMPTION_START=") Then

                        Dim tmp_str As String = Arg.ToUpper.Remove(0, "RESUMPTION_START=".Length)

                        Dim tmp_Finish As Integer = InStr(tmp_str, "_LENGTH=")

                        If tmp_Finish > 0 Then

                            Me.Resumption_start_time = Val(Mid(tmp_str, 1, tmp_str.Length - tmp_Finish))

                            Me.Resumption_length_time = Val(Mid(tmp_str, tmp_Finish + "_LENGTH=".Length))

                        Else

                            Me.Resumption_start_time = Val(tmp_str)

                        End If

                    End If

                    If Arg.ToUpper.StartsWith("RESUMPTION_LENGTH=") Then

                        Me.Resumption_length_time = Val(Arg.ToUpper.Remove(0, "RESUMPTION_LENGTH=".Length))

                    End If

                End If

                Exit Do

            Loop

        Next i

        Me.Show()

    End Sub

    Private Function User_file_selection() As String

        Dim ofd As New OpenFileDialog

        ofd.Multiselect = False

        ofd.RestoreDirectory = True

        ofd.Title = "Select model"

        ofd.Filter = "*.bdf files (*.bdf)|*.bdf|compiled model files (*.model)|*.model|*.dat files (*.dat)|*.dat|All files (*.*)|*.*"

        ofd.FilterIndex = 0

        ofd.ShowDialog()

        User_file_selection = ofd.FileName

    End Function

    Private Sub OpenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenToolStripMenuItem.Click

        Me.Ready = False

        FEM = New Model

        Me.Input_filename = User_file_selection()

        Me.Output_filename = ""

        Open_model()

        If FEM.Polygon Is Nothing Then Exit Sub

        Me.Refresh()


        Dim Center As New Visualisation.Vertex

        Center.x = FEM.Center.Coord(0)

        Center.y = FEM.Center.Coord(1)

        Center.z = FEM.Center.Coord(2)

        If Not (Me.view Is Nothing) Then

            Me.view.frm.Close()

            Me.view = Nothing

        End If

        Me.view = New Visualisation.Visualisation(V_mode, Model_for_visualisation_number, FEM.Polygon, UBound(FEM.Polygon) + 1, 500, 500, 32, True, Center, FEM.Maximum_length, FEM.Minimum_length, 60)

        Me.view.frm.Text = Me.Input_filename

        Me.view.frm.Icon = Me.My_icon

        ' Detecting is results in file or not

        Me.Solver.Detect_Min_Max_T()

        If Me.Solver.T Is Nothing Then

            ' No results in file

            Value_strip_status.Checked = False

            Save_results_as_txt.Enabled = False

        Else

            ' The results are in file

            Fill_Step_selector()

            Solver.Current_step = 0

            Draw_current_result_step()

            Value_strip_status.Checked = True

            view.frm.Value_strip_activated = True

            Solver.Create_value_strip(view, 10, Solver.Min_T, Solver.Max_T, "Temp., K")

            Me.Height = Result_view_Height

            Me.Width = Result_view_Width

            Save_results_as_txt.Enabled = True

        End If

        Me.Ready = True

        ThermalAnalysisToolStripMenuItem.Enabled = True

        SaveToolStripMenuItem.Enabled = True

        Dim b As New System.Windows.Forms.MouseEventArgs(System.Windows.Forms.MouseButtons.Left, 1, 1, 1, 1)

        view.frm.Form_MouseMove(Nothing, b)


    End Sub
    ''' <summary>
    ''' This subrutine save model with results to Me.Output_filename
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub Save_model()

        If Me.Output_filename = "" Then

            ' Проверяем не с ФТП ли мы работаем?
            If InStr(Input_filename.ToLower, "ftp") > 0 Then

                Exit Sub

            Else

                'Me.Output_filename = FEM.File_path & FEM.File_name & "_" & Now.Date & "_" & Now.Hour & "_" & Now.Minute & "_" & Now.Second & ".model"
                Me.Output_filename = FEM.File_path & FEM.File_name & "_" & ID.ToString & ".model"

            End If

        End If

        Dim tmp_File_name As String = Me.tmp_Directory & ID.ToString & "_server_model.tmp"

        ' Сериализация
        Dim fs As New FileStream(tmp_File_name, FileMode.Create)
        Dim bf As New BinaryFormatter
        bf.Serialize(fs, Me.Solver.FEM)
        bf.Serialize(fs, Me.Solver.HT_element)
        bf.Serialize(fs, Me.Solver.time)
        bf.Serialize(fs, Me.Solver.T)
        bf.Serialize(fs, Me.Solver.Q)
        fs.Close()

        ' если работаем локально, то все просто

        Try

            If System.IO.File.Exists(Output_filename) Then System.IO.File.Delete(Output_filename)

            System.IO.File.Move(tmp_File_name, Output_filename)

        Catch ex As Exception

        End Try

        

    End Sub

    Friend Sub Save_txt_results()

        Dim String_to_file As String

        If Me.Output_txt_filename = "" Then Exit Sub

        Dim Txt_writer As System.IO.StreamWriter = New System.IO.StreamWriter(Me.Output_txt_filename)

        Txt_writer.WriteLine("Element_number Time,s Temperature,K  Input_heat_power,W")

        For i As Integer = 0 To UBound(Solver.T, 1)

            For j As Integer = 0 To UBound(Solver.time)

                With Solver.HT_element(i)

                    String_to_file = .Element.Number & " " & Math.Round(Solver.time(j), 6) & " " & Math.Round(Solver.T(i, j), 6) & " " & Math.Round(Solver.Q(i, j), 6)

                    String_to_file = Replace(String_to_file, ",", ".")

                    Txt_writer.WriteLine(String_to_file)

                End With

            Next j

        Next i

        Txt_writer.Close()

    End Sub


    ''' <summary>
    ''' This subrutine reads model from me.Input_filename file
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Open_model()

        Me.Ready = False

        Me.FEM = New Model

        ' проверка наличия файла 

        If Not (System.IO.File.Exists(Input_filename)) Then Exit Sub

        Dim File_extension As String

        File_extension = Mid(Me.Input_filename, InStrRev(Me.Input_filename, ".") + 1).ToLower

        Select Case File_extension

            Case "model"

                Dim fs As FileStream
                Dim bf As New BinaryFormatter

                Me.Solver = Nothing

                Dim HT_element() As HT_Element

                Dim time(), T(,), Q(,) As Double

                ReDim time(0), T(0, 0), Q(0, 0)


                ' Десериализация.
                fs = New FileStream(Me.Input_filename, FileMode.Open)
                Me.FEM = New Model
                Me.FEM = Convert.ChangeType(bf.Deserialize(fs), Me.FEM.GetType())
                Me.Solver = New TOP_solver(Me.FEM)
                HT_element = Convert.ChangeType(bf.Deserialize(fs), Me.Solver.HT_element.GetType())
                time = Convert.ChangeType(bf.Deserialize(fs), time.GetType())
                T = Convert.ChangeType(bf.Deserialize(fs), T.GetType())
                Q = Convert.ChangeType(bf.Deserialize(fs), Q.GetType())
                fs.Close()

                ' Connection

                Me.Solver.HT_element = HT_element
                Me.Solver.time = time
                Me.Solver.T = T
                Me.Solver.Q = Q


                'For i As Integer = 0 To UBound(Me.Solver.HT_element)

                '    Me.Solver.HT_element(i).HT_Step = HT_element(i).HT_Step

                'Next i

                Solution_Progressbar.Visible = False

                Me.Text = App_Name

            Case "bdf", "dat"

                FEM = New Model(Me.Input_filename)

                'For i As Integer = 0 To UBound(FEM.Shining_face)

                '    FEM.Shining_face(i).Code = New Face_code

                '    FEM.Shining_face(i).Code.R = 0.1 * 255

                '    FEM.Shining_face(i).Code.G = 0.2 * 255

                '    FEM.Shining_face(i).Code.B = 0.5 * 255

                'Next i

                Me.Solver = New TOP_solver(FEM)

                'Solver.Fill_HT_element()

                'Solver.Init_temperature_field()

                Solver.Prepare_Elements_to_show_before_analysis()

                Me.FEM.Polygon = Me.FEM.Extract_Polygon_from_Face(Me.FEM.Shining_face)

                Me.Height = Start_Height

                Me.Width = Start_Width

            Case Else

                Exit Sub

        End Select

        Me.Solver.Calc_Elements_Faces_codes()

    End Sub


    Public Sub Clear_Dir(ByRef Directory As String, ByRef File_pattern As String)

        Dim File_for_delete_list() As String

        File_for_delete_list = System.IO.Directory.GetFiles(Directory, File_pattern)

        If File_for_delete_list Is Nothing Then Exit Sub

        For i As Integer = 0 To UBound(File_for_delete_list)

            Try

                System.IO.File.Delete(File_for_delete_list(i))

            Catch

            End Try

        Next i

    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click

        Application.Exit()

    End Sub

    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OpenToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RunToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem

    'Private Function NASTRAN_file_selection() As String

    '    Dim ofd As New OpenFileDialog

    '    ofd.Multiselect = False

    '    ofd.RestoreDirectory = True

    '    ofd.Title = "Select NASTRAN input file"

    '    ofd.Filter = "*.bdf files (*.bdf)|*.bdf|*.dat files (*.dat)|*.dat|All files (*.*)|*.*"

    '    ofd.FilterIndex = 0

    '    ofd.ShowDialog()

    '    NASTRAN_file_selection = ofd.FileName

    'End Function

    'Private Function RHT_Matrix_file_selection() As String

    '    Dim ofd As New OpenFileDialog

    '    ofd.Multiselect = False

    '    ofd.RestoreDirectory = True

    '    ofd.Title = "Select radiation heat transfer matrix file"

    '    ofd.Filter = "Radiation heat transfer matrix file (*.mx)|*.mx|All files (*.*)|*.*"

    '    ofd.FilterIndex = 0

    '    ofd.ShowDialog()

    '    RHT_Matrix_file_selection = ofd.FileName

    'End Function

    'Private Function RHT_Matrix_file_for_saving() As String

    '    Dim ofd As New SaveFileDialog

    '    ofd.RestoreDirectory = True

    '    ofd.Title = "Select radiation heat transfer matrix file"

    '    ofd.Filter = "Radiation heat transfer matrix file (*.mx)|*.mx|All files (*.*)|*.*"

    '    ofd.FilterIndex = 0

    '    ofd.ShowDialog()

    '    RHT_Matrix_file_for_saving = ofd.FileName

    'End Function

    'Private Function Binary_model_file_selection() As String

    '    Dim ofd As New OpenFileDialog

    '    ofd.Multiselect = False

    '    ofd.RestoreDirectory = True

    '    ofd.Title = "Select binary model file"

    '    ofd.Filter = "Binary model file (*.bm)|*.bm|All files (*.*)|*.*"

    '    ofd.FilterIndex = 0

    '    ofd.ShowDialog()

    '    Binary_model_file_selection = ofd.FileName

    'End Function

    'Private Function Results_file_selection() As String

    '    Dim ofd As New OpenFileDialog

    '    ofd.Multiselect = False

    '    ofd.RestoreDirectory = True

    '    ofd.Title = "Select analysys result file"

    '    ofd.Filter = "Analysys result file (*.ar)|*.ar|All files (*.*)|*.*"

    '    ofd.FilterIndex = 0

    '    ofd.ShowDialog()

    '    Results_file_selection = ofd.FileName

    'End Function

    Private Function Results_file_for_saving() As String

        Dim ofd As New SaveFileDialog

        ofd.RestoreDirectory = True

        ofd.Title = "Select model file for saving"

        ofd.Filter = "compiled model files (*.model)|*.model|all files (*.*)|*.*"

        ofd.FilterIndex = 0

        ofd.ShowDialog()

        Results_file_for_saving = ofd.FileName

    End Function

    Private Function Txt_file_for_saving() As String

        Dim ofd As New SaveFileDialog

        ofd.RestoreDirectory = True

        ofd.Title = "Select text file for saving results"

        ofd.Filter = "text files (*.txt)|*.txt|all files (*.*)|*.*"

        ofd.FilterIndex = 0

        ofd.ShowDialog()

        Txt_file_for_saving = ofd.FileName

    End Function
    ''' <summary>
    ''' Fill Pult.Step_selector with step numbers and times
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Fill_Step_selector()

        If Solver.time Is Nothing Then Exit Sub

        Step_selector.Hide()

        Step_selector.Items.Clear()

        Step_selector.DropDownStyle = ComboBoxStyle.DropDownList

        Dim Step_string As String

        For i As Integer = 0 To UBound(Solver.time)

            Step_string = "Step " & i.ToString & ". Time " & Str(Solver.time(i)) & " s."

            Step_selector.Items.Add(Step_string)

        Next i

        Step_selector.SelectedIndex = 0

        Step_selector.Show()

    End Sub

    Public Sub Show_results()

        ' Отрисовываем результаты

        Dim Is_form_saved As Boolean

        Dim F_S As New Visualisation.Form_saver

        If Not (view Is Nothing) Then

            ' Сохраняем форму визуализации

            F_S.Save(view)

            Is_form_saved = True

        Else


        End If

        Dim Center As New Visualisation.Vertex

        Center.x = FEM.Center.Coord(0)

        Center.y = FEM.Center.Coord(1)

        Center.z = FEM.Center.Coord(2)

        'Me.FEM.Polygon = Me.FEM.Extract_Polygon_from_Face(Me.FEM.Shining_face)

        Solver.Prepare_Draw_Temperature(Solver.Current_step)

        'Solver.FEM.Polygon = FEM.Extract_Polygon_from_Face(Solver.FEM.Shining_face)

        'view.Polygon = Solver.FEM.Polygon

        Me.view = New Visualisation.Visualisation(V_mode, Model_for_visualisation_number, FEM.Polygon, UBound(FEM.Polygon) + 1, 500, 500, 32, True, Center, FEM.Maximum_length, FEM.Minimum_length, 60)

        Me.view.frm.Icon = Me.My_icon

        If Is_form_saved Then

            F_S.Load(Me.view)

        End If

        Me.Show()

        Solver.Create_value_strip(view, 10, Solver.Min_T, Solver.Max_T, "Temp., K")

        Value_strip_status.Checked = True

        Me.view.frm.Value_strip_activated = True

        Me.view.frm.Show()

        Fill_Step_selector()

        Solution_Progressbar.Visible = False

        Me.Height = Result_view_Height

        Me.Width = Result_view_Width

        Me.Text = App_Name

    End Sub
    Friend WithEvents ThermalAnalysisToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem

    Private Sub ThermalAnalysisToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ThermalAnalysisToolStripMenuItem.Click

        view.frm.Hide()

        Solution_Progressbar.Visible = True

        ' Решаем задачу

        Solver.Solve_problem()

        Show_results()

        Me.notifyIcon1.Text = App_Name

        Save_results_as_txt.Enabled = True

        '' Отрисовываем результаты

        '' Сохраняем форму визуализации
        'Dim F_S As New Visualisation.Form_saver

        'F_S.Save(view)

        'Dim Center As New Visualisation.Vertex

        'Center.x = FEM.Center.Coord(0)

        'Center.y = FEM.Center.Coord(1)

        'Center.z = FEM.Center.Coord(2)

        ''Me.FEM.Polygon = Me.FEM.Extract_Polygon_from_Face(Me.FEM.Shining_face)

        'Solver.Prepare_Draw_Temperature(Solver.Current_step)

        'Solver.FEM.Polygon = FEM.Extract_Polygon_from_Face(Solver.FEM.Shining_face)

        'view.Polygon = Solver.FEM.Polygon

        'Me.view = New Visualisation.Visualisation(V_mode, Model_for_visualisation_number, FEM.Polygon, UBound(FEM.Polygon) + 1, 500, 500, 32, True, Center, 10 * FEM.Maximum_length, FEM.Minimum_length, 60)

        'F_S.Load(Me.view)

        'Me.Show()

        'Solver.Create_value_strip(view, 10, Solver.Min_T, Solver.Max_T, "Temp., K")

        'Me.view.frm.Show()

        'Fill_Step_selector()

        'Pult.Show()

        'Solution_Progressbar.Visible = False

        'Me.Text = App_Name

    End Sub
    Friend WithEvents SaveToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator3 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents Solution_Progressbar As System.Windows.Forms.ProgressBar

    Private Sub Solution_progress_visualisation(ByRef Progress As Integer) Handles Solver.Solution_progress

        Solution_Progressbar.Value = Progress

        Solution_Progressbar.Refresh()

        Me.Text = App_Name & " Current solution progress " & Progress & "%."

        If Me.notifyIcon1.Icon Is Icon_green Then

            Me.notifyIcon1.Icon = Icon_red

        Else

            Me.notifyIcon1.Icon = Icon_green

        End If

        Me.notifyIcon1.Text = App_Name & Chr(13) _
        & "Current solution progress " & Progress & "%."

        Application.DoEvents()

    End Sub

    Private Sub B_matrix_progress_visualisation(ByRef Progress As Integer) Handles Solver.B_matrix_progress

        Solution_Progressbar.Value = Progress

        Solution_Progressbar.Refresh()

        Me.Text = App_Name & " Current matrix calculation progress " & Progress & "%."

        Application.DoEvents()

    End Sub
    Friend WithEvents HelpToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem

    Private Sub HelpToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HelpToolStripMenuItem.Click

        MsgBox("T.H.O.R.I.U.M." & Chr(10) & Chr(13) & "Thermal-optical radiation iteration universal module" & Chr(10) & Chr(13) & "Alexander Shaenko, Moscow, Russia, 2010", , App_Name)

    End Sub

    Private Sub SaveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripMenuItem.Click

        Me.Output_filename = Results_file_for_saving()

        Save_model()

    End Sub
    Friend WithEvents Save_results_as_txt As System.Windows.Forms.ToolStripMenuItem

    Private Sub Save_results_as_txt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Save_results_as_txt.Click

        Me.Output_txt_filename = Txt_file_for_saving()

        Save_txt_results()

    End Sub

    Private Sub menuItem1_Click(ByVal Sender As Object, ByVal e As EventArgs) Handles menuItem1.Click
        ' Close the form, which closes the application.
        Me.Close()

        End
    End Sub


    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
    Friend WithEvents Step_selector As System.Windows.Forms.ComboBox
    Friend WithEvents Value_strip_status As System.Windows.Forms.CheckBox
    Friend WithEvents Animation As System.Windows.Forms.Button
    Friend WithEvents Next_step As System.Windows.Forms.Button
    Friend WithEvents To_last_step As System.Windows.Forms.Button
    Friend WithEvents To_first_step As System.Windows.Forms.Button
    Friend WithEvents Previous_step As System.Windows.Forms.Button
    Friend WithEvents Stop_button As System.Windows.Forms.Button
    Friend WithEvents Decimate_results As System.Windows.Forms.Button
    Friend WithEvents Colors_from_step As System.Windows.Forms.RadioButton
    Friend WithEvents Colors_from_entire_solution As System.Windows.Forms.RadioButton
    Friend WithEvents Visual_mode As System.Windows.Forms.ComboBox

    Private Sub Solution_Progressbar_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Solution_Progressbar.VisibleChanged

        If Me.Solution_Progressbar.Visible Then

            Me.Height = Progress_Height

            Me.Width = Progress_Width

        Else

            Me.Height = Start_Height

            Me.Width = Start_Width

        End If

    End Sub

    Private Sub Decimate_results_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Decimate_results.Click

        Dim str_N As String = InputBox("Please, enter the decimation ratio", )

        Dim N As Integer

        If str_N <> "" Then

            N = Val(str_N)

            If N <= 0 Then Exit Sub

        Else

            Exit Sub

        End If

        Me.Solver.Decimate_results(N)

    End Sub

    Private b As New System.Windows.Forms.MouseEventArgs(System.Windows.Forms.MouseButtons.Left, 1, 1, 1, 1)

    Public Sub Draw_current_result_step()

        If Solver.T Is Nothing Then Exit Sub

        If Solver.Colors_from_entire_solution Then

        Else

            Solver.Detect_Min_Max_T_on_current_step()

            If Solver.Min_T <> Solver.Max_T Then

                Solver.Create_value_strip(view, 10, Solver.Min_T, Solver.Max_T, "Temp., K")

            End If

        End If

        Step_selector.SelectedIndex = Solver.Current_step

        Solver.Prepare_Draw_Temperature(Solver.Current_step)

        'Solver.FEM.Polygon = FEM.Extract_Polygon_from_Face(Solver.FEM.Shining_face)

        Solver.FEM.Polygon = FEM.Extract_Polygon_colors_from_Face(Solver.FEM.Polygon)

        view.Polygon = Solver.FEM.Polygon

        view.frm.Polygon = Solver.FEM.Polygon

        view.Create_display_list(Model_for_visualisation_number, V_mode.Status)

        view.frm.Text = "Step " & Solver.Current_step.ToString & ". Time " & Solver.time(Solver.Current_step).ToString & " s."

        view.frm.Form_MouseMove(Nothing, b)

    End Sub

    Private Sub Next_step_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Next_step.Click

        Solver.Current_step += 1

        If Solver.Current_step > UBound(Solver.time) Then

            Solver.Current_step -= 1

        End If

        Draw_current_result_step()


    End Sub

    Private Sub Previous_step_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Previous_step.Click

        Solver.Current_step -= 1

        If Solver.Current_step < 0 Then

            Solver.Current_step += 1

        End If

        Draw_current_result_step()

    End Sub

    Private Sub Animation_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Animation.Click

        For i As Long = Solver.Current_step To UBound(Solver.time)

            Solver.Current_step = i

            Draw_current_result_step()

            Application.DoEvents()

            If Stop_button.Tag = "Was pressed" Then

                Stop_button.Tag = ""

                Exit For

            End If

        Next i

    End Sub

    Private Sub To_first_step_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles To_first_step.Click

        Solver.Current_step = 0

        Draw_current_result_step()

    End Sub



    Private Sub To_last_step_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles To_last_step.Click

        Solver.Current_step = UBound(Solver.time)

        Draw_current_result_step()

    End Sub

    Private Sub Value_strip_status_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Value_strip_status.CheckedChanged

        view.frm.Value_strip_activated = Value_strip_status.Checked

        If Ready Then

            Dim b As New System.Windows.Forms.MouseEventArgs(System.Windows.Forms.MouseButtons.Left, 1, 1, 1, 1)

            view.frm.Form_MouseMove(Me, b)

        End If

    End Sub

    Private Sub Step_selector_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Step_selector.SelectedIndexChanged

        Solver.Current_step = Step_selector.SelectedIndex

        Draw_current_result_step()

    End Sub

    Private Sub Colors_from_step_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Colors_from_step.CheckedChanged

        If Not (Me.Ready) Then Exit Sub

        If Colors_from_step.Checked Then

            Me.Solver.Colors_from_entire_solution = False

            Solver.Calc_Elements_Faces_codes()

            Draw_current_result_step()

            view.frm.Form_MouseMove(Nothing, b)

        End If

    End Sub

    Private Sub Colors_from_entire_solution_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Colors_from_entire_solution.CheckedChanged

        If Not (Me.Ready) Then Exit Sub

        If Colors_from_entire_solution.Checked Then

            Me.Solver.Colors_from_entire_solution = True

            Solver.Calc_Elements_Faces_codes()

            If Solver.Min_T <> Solver.Max_T Then

                Solver.Create_value_strip(view, 10, Solver.Min_T, Solver.Max_T, "Temp., K")

            End If

            Draw_current_result_step()

            view.frm.Form_MouseMove(Nothing, b)

        End If

    End Sub

    Private Sub Stop_button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Stop_button.Click

        Stop_button.Tag = "Was pressed"

    End Sub

    Private Sub Visual_mode_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Visual_mode.SelectedIndexChanged

        If Not (Ready) Then Exit Sub

        Select Case Visual_mode.SelectedIndex

            Case 0

                view.Mode.Status = Visualisation.Mode.WITH_CONTOUR

            Case 1

                view.Mode.Status = Visualisation.Mode.WITHOUT_CONTOUR

            Case 2

                view.Mode.Status = Visualisation.Mode.WIREFRAME

        End Select

        Draw_current_result_step()

    End Sub
    Friend WithEvents Show_model As System.Windows.Forms.Button

    Private Sub Show_model_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Show_model.Click

        view.frm.Hide()

        view = Nothing

        Show_results()

    End Sub
End Class






