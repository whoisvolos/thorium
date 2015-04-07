''' <summary>
''' This class solves system of heat transfer equations for specifed period of time with defined time step.
''' It saves temperature and inocming heat power vectors for all of step.
''' </summary>
'''<remarks></remarks>
<Serializable()> _
Public Class TOP_solver

    Inherits TOP_item

    ''' <summary>
    ''' The finite element model for heat transfer problem.
    ''' </summary>
    ''' <remarks></remarks>
    Public FEM As Model

    ''' <summary>
    ''' Array of elements, which can involve in heat transfer.
    ''' Elements not included in this array are excluded from heat transfer model.
    ''' </summary>
    ''' <remarks></remarks>
    Public HT_element() As HT_Element


    ''' <summary>
    ''' Matrix of angle coefficients for elements.
    ''' </summary>
    ''' <remarks></remarks>
    Public B_matrix(,) As Double

    ''' <summary>
    ''' The solution start time, i.e. heat transfer problem will be solved for time interval from Start_time to Start_time + Lenght_time.
    ''' </summary>
    ''' <remarks></remarks>
    Public Start_time As Double

    ''' <summary>
    ''' The solution finish time, i.e. heat transfer problem will be solved for time interval from Start_time to Start_time + Lenght_time.
    ''' </summary>
    ''' <remarks></remarks>
    Public Lenght_time As Double

    ''' <summary>
    ''' The duration of solution time step. 
    ''' </summary>
    ''' <remarks></remarks>
    Public Step_time As Double


    ''' <summary>
    ''' The element's available temperature in case of cooling coefficient 
    ''' </summary>
    ''' <remarks></remarks>
    Public K_T_min As Double

    ''' <summary>
    ''' The element's available temperature in case of heating coefficient 
    ''' </summary>
    ''' <remarks></remarks>
    Public K_T_max As Double


    ''' <summary>
    ''' The smoothing type used for incoming heat power smoothing
    ''' The available values "LINEAR" and "CUBIC"
    ''' </summary>
    ''' <remarks></remarks>
    Public Smoothing_type As String

    ''' <summary>
    ''' The number of last steps, used for incoming heat power smoothing in case of "LINEAR" smoothing
    ''' </summary>
    ''' <remarks></remarks>
    Public N_steps As Integer

    ''' <summary>
    ''' Time during solution process. It is changed from step to step
    ''' </summary>
    ''' <remarks></remarks>
    Public Current_time As Double


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
    ''' Maximum relative error Newton method temperature calculation
    ''' </summary>
    ''' <remarks></remarks>
    Const Calc_T_Epsilon = 0.001


    ''' <summary>
    ''' Number of photons, emitted from face at time of one solution step
    ''' </summary>
    ''' <remarks></remarks>
    Public N_photon As Integer

    ''' <summary>
    ''' Minimum and maximum temperatures from all of steps, all of elements
    ''' </summary>
    ''' <remarks></remarks>
    Public Min_T, Max_T As Double

    '''' <summary>
    '''' Auxuluary visualisator form. Used for OpenGL operations.
    '''' </summary>
    '''' <remarks></remarks>
    'Public aux_Visual As Visualisation.Visualisation

    ''' <summary>
    ''' Current step of thermal analysis visualisation. Used only for visualisation.
    ''' </summary>
    ''' <remarks></remarks>
    Public Current_step As Long

    Public Glob_x As New Vector

    Public Glob_y As New Vector

    Public Glob_z As New Vector

    Public Global_medium As MEDIUM1


    Public Event Solution_progress(ByRef Progress As Integer)

    Public Event B_matrix_progress(ByRef Progress As Integer)

    ''' <summary>
    ''' FTP protocol procedures
    ''' </summary>
    ''' <remarks></remarks>
    Public FTP As FTP = New FTP

    ''' <summary>
    ''' Whether Work_directory is on FTP server?
    ''' </summary>
    ''' <remarks></remarks>
    Public FTP_mode As Boolean

    ''' <summary>
    ''' Local directory to store files 
    ''' </summary>
    ''' <remarks></remarks>
    Public tmp_Dir As String


    Public Solution_integrity As Boolean

    ''' <summary>
    ''' If Colors_from_entire_solution is TRUE, then minimal and maximal temparatures are taken from entire analysis.
    ''' If Colors_from_entire_solution is FALSE, then minimal and maximal temparatures are taken from current step tempatature distribution.
    ''' </summary>
    ''' <remarks></remarks>
    Public Colors_from_entire_solution As Boolean = True

    ''' <summary>
    ''' This member  implements the Mersenne Twister (MT), MT19937ar, pseudo-random number generator (PRNG) 
    ''' </summary>
    ''' <remarks></remarks>
    Public MT_RND As New MTRandom(True)

    'Функция timeGetTime() измеряет промежуток времени с момента запуска Windows
    Public Declare Function timeGetTime Lib "winmm.dll" () As Long

    Public Sub New(ByRef FEM As Model)

        Me.FEM = FEM

        Me.Start_time = FEM.Start_time

        Me.Lenght_time = FEM.Lenght_time

        Me.Step_time = FEM.Step_time

        'Me.TOP_Epsilon = FEM.TOP_Epsilon

        Me.N_photon = FEM.N_photon

        Me.K_T_min = FEM.K_T_min

        Me.K_T_max = FEM.K_T_max

        Me.Smoothing_type = FEM.Smoothing_type

        Me.N_steps = FEM.N_steps

        Me.Solution_integrity = True


        Glob_x.Coord(0) = 1

        Glob_x.Coord(1) = 0

        Glob_x.Coord(2) = 0


        Glob_y.Coord(0) = 0

        Glob_y.Coord(1) = 1

        Glob_y.Coord(2) = 0


        Glob_z.Coord(0) = 0

        Glob_z.Coord(1) = 0

        Glob_z.Coord(2) = 1


        Global_medium = FEM.Global_medium


        Fill_HT_element()

        Init_temperature_field()

    End Sub

    Public Overrides Sub Load()

    End Sub
    Public Overrides Sub Prepare_save()

    End Sub

    ''' <summary>
    ''' This sub fill HT_element array with 2D and 3D element.
    ''' Only these types of elements can involve in heat transfering.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Fill_HT_element()

        Dim i_HT_element As Integer = -1

        Dim i_Face As Integer

        Dim N_shining_faces As Integer

        Dim Is_HT_element As Boolean

        If FEM.Element Is Nothing Then Exit Sub

        ReDim HT_element(UBound(FEM.Element))

        For i As Integer = 0 To UBound(FEM.Element)

            Is_HT_element = False

            N_shining_faces = -1

            i_Face = -1

            If Not (FEM.Element(i).Face Is Nothing) Then

                Is_HT_element = True

                For j As Integer = 0 To UBound(FEM.Element(i).Face)

                    If FEM.Element(i).Face(j).Can_Radiate Then

                        N_shining_faces += 1

                    End If

                Next j

            End If

            If Is_HT_element Then

                ' заполняем поля элемента теплопроводности

                i_HT_element += 1

                HT_element(i_HT_element) = New HT_Element

                HT_element(i_HT_element).Element = FEM.Element(i)

                FEM.Element(i).HT_element = HT_element(i_HT_element)

                HT_element(i_HT_element).Number = i_HT_element

                ' заполняем массив излучающих граней

                ReDim HT_element(i_HT_element).Face(N_shining_faces)

                For j As Integer = 0 To UBound(FEM.Element(i).Face)

                    If FEM.Element(i).Face(j).Can_Radiate Then

                        i_Face += 1

                        HT_element(i_HT_element).Face(i_Face) = FEM.Element(i).Face(j)

                        HT_element(i_HT_element).Face(i_Face).HT_element = HT_element(i_HT_element)

                    End If

                Next j

            End If

        Next i

        ReDim Preserve HT_element(i_HT_element)

        For i As Integer = 0 To UBound(HT_element)

            ' заполняем элементы теплопроводности

            ReDim HT_element(i).Heat_conductivity(HT_element(i).Element.N_Neighbor)

            For j As Integer = 0 To HT_element(i).Element.N_Neighbor

                HT_element(i).Heat_conductivity(j) = New Heat_conductivity

            Next j

            For j As Integer = 0 To HT_element(i).Element.N_Neighbor

                'HT_element(i).Heat_conductivity(j).Area = HT_element(i).Element.Neighbor(j).Area * Math.Abs(Cos_angle_between_vectors(HT_element(i).Element.Neighbor(j).Face.oz, HT_element(i).Element.Neighbor(j).Face.Center - HT_element(i).Element.Field.Centroid))

                HT_element(i).Heat_conductivity(j).Area = HT_element(i).Element.Neighbor(j).Area

                'HT_element(i_HT_element).Heat_conductivity(j).Radius_vector = FEM.Element(i).Field.Centroid - FEM.Element(i).Neighbor(j).Element.Field.Centroid

                'HT_element(i).Heat_conductivity(j).Distance = L_vector(HT_element(i).Element.Field.Centroid - HT_element(i).Element.Neighbor(j).Face.Center) + L_vector(HT_element(i).Element.Neighbor(j).Face.Center - HT_element(i).Element.Neighbor(j).Element.Field.Centroid)

                HT_element(i).Heat_conductivity(j).Distance_1 = L_vector(HT_element(i).Element.Field.Centroid - HT_element(i).Element.Neighbor(j).Face.Center)

                HT_element(i).Heat_conductivity(j).Distance_2 = L_vector(HT_element(i).Element.Neighbor(j).Face.Center - HT_element(i).Element.Neighbor(j).Element.Field.Centroid)

                HT_element(i).Heat_conductivity(j).HT_Element = HT_element(i).Element.Neighbor(j).Element.HT_Element

                ' осталось заполнить теплопроводность, но она зависит от температуры, 
                ' поэтому ее заполняем во время расчета

            Next j

            Application.DoEvents()

        Next i

        ' заполняем массив источников тепла, которые действуют на элемент

        For i As Integer = 0 To UBound(HT_element)

            ReDim HT_element(i).Heat_source(0)

        Next i


        For i As Integer = 0 To UBound(FEM.HSOURCE)

            Dim Element_collection As Collection = FEM.Convert_array_to_collection(FEM.HSOURCE(i).Element)

            For j As Integer = 0 To UBound(HT_element)

                If Element_collection.Contains(HT_element(j).Element.Number.ToString) Then

                    HT_element(j).Heat_source(UBound(HT_element(j).Heat_source)) = FEM.HSOURCE(i)

                    ReDim Preserve HT_element(j).Heat_source(UBound(HT_element(j).Heat_source) + 1)

                End If

            Next j

        Next i

        For i As Integer = 0 To UBound(HT_element)

            ReDim Preserve HT_element(i).Heat_source(UBound(HT_element(i).Heat_source) - 1)

        Next i

        ' записываем, есть ли у элемента T_SET

        For i As Integer = 0 To UBound(FEM.T_SET)

            For j As Integer = 0 To UBound(FEM.T_SET(i).Element)

                FEM.T_SET(i).Element(j).HT_element.T_SET = FEM.T_SET(i)

                'FEM.T_SET(i).Element(j) = FEM.T_SET(i).Element(j).HT_element

            Next j

        Next i

    End Sub

    ''' <summary>
    ''' This sub initialized initial temperature field over all of model and
    ''' prepare FTOP for processing.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Init_temperature_field()

        If FEM.Element Is Nothing Then Exit Sub

        For i As Integer = 0 To UBound(FEM.Element)

            FEM.Element(i).Field.Temperature = FEM.T_ALL

        Next i

        If UBound(FEM.T_INIT) > -1 Then

            For i As Long = 0 To UBound(FEM.T_INIT)

                For j As Integer = 0 To UBound(FEM.T_INIT(j).Element)

                    FEM.T_INIT(i).Element(j).Field.Temperature = FEM.T_INIT(i).Temp(j)

                Next j

            Next i

        End If

        ' We should create argument before calculation

        Dim Argument As Vector = New Vector

        ReDim Argument.Coord(5)

        Argument.Coord(1) = Start_time

        If UBound(FEM.T_SET) > -1 Then

            For i As Long = 0 To UBound(FEM.T_SET)

                For j As Integer = 0 To UBound(FEM.T_SET(i).Element)

                    If FEM.T_SET(i).Temp(j).GetType.ToString = "System.Double" Then

                        FEM.T_SET(i).Element(j).Field.Temperature = FEM.T_SET(i).Temp(j)

                    Else

                        FEM.T_SET(i).Element(j).Field.Temperature = FEM.T_SET(i).Temp(j).Return_value(Argument).Coord(0)

                    End If

                Next j

            Next i

        End If

        If Not (FEM.Is_solution_constant) Then

            For i As Integer = 0 To UBound(Me.FEM.FTOP)

                Me.FEM.FTOP(i).Calc_Integral_Eps_DEPEND()

                'Me.FEM.FTOP(i).Calc_Integral_A_DEPEND()

            Next i

        End If

    End Sub

    ''' <summary>
    ''' This sub fill matrices of angle coefficients for faces of elements and elements.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Fill_B_matrices()

        Dim Progress As Integer = 0

        'Dim i_Photon As Integer

        Dim Temperature, Area, eps, B As Double

        ' подготавливаем OpenGL для работы

        Dim Center As New Visualisation.Vertex

        Center.x = FEM.Center.Coord(0)

        Center.y = FEM.Center.Coord(1)

        Center.z = FEM.Center.Coord(2)

        Dim V_mode As New Visualisation.Mode

        V_mode.Status = Visualisation.Mode.WITHOUT_CONTOUR

        Dim aux_Visual As Visualisation.Visualisation

        Dim Model_list_number As Integer = 100

        aux_Visual = New Visualisation.Visualisation(V_mode, Model_list_number, Me.FEM.Polygon, UBound(Me.FEM.Polygon) + 1, 400, 400, 16, False, Center, 2 * FEM.Maximum_length, FEM.Minimum_length, 0.1)


        '' создаем визуализационную форму для расчета теплопередачи

        '' подготавливаем модель

        'For i As Integer = 0 To UBound(FEM.Shining_face)

        '    FEM.Shining_face(i).code = New Face_code

        'Next i

        'FEM.Fill_Shining_faces_colours()

        'FEM.Polygon = FEM.Extract_Polygon_from_Face(FEM.Shining_face)

        'Init_temperature_field()

        '' подготавливаем OpenGL для работы

        'Dim Center As New Visualisation.Vertex

        'Center.x = FEM.Center.Coord(0)

        'Center.y = FEM.Center.Coord(1)

        'Center.z = FEM.Center.Coord(2)

        'Dim V_mode As New Visualisation.Mode

        'V_mode.Status = Visualisation.Mode.WITHOUT_CONTOUR

        'Dim aux_Visual As Visualisation.Visualisation

        'Dim Model_list_number As Integer = 100

        'aux_Visual = New Visualisation.Visualisation(V_mode, Model_list_number, FEM.Polygon, UBound(Me.FEM.Polygon) + 1, 400, 400, 16, False, Center, 1.1 * FEM.Maximum_length, 0.5 * FEM.Minimum_length, 0.1)


        ReDim B_matrix(UBound(HT_element), UBound(HT_element))

        ' Аргумент для расчета свойств грани, для которой рассчитывается строчка

        Dim Argument As New Vector

        ReDim Argument.Coord(5)

        ' Вот это и есть нижеописанная строчка матрицы B
        Dim B_row(UBound(HT_element)) As Double

        ' Одна строчка матрицы B для граней.
        ' Если строчка имеет номер i, то она предствляет собой доли энергии, попадащие с i-го грани
        ' на все остальные грани.
        ' Для получения долей энергии со всех граней на i-ю грань необходимо брать i-й столбец матрицы B.

        ' проходим по всем граням всех элементов и строим для них строчки матрицы B
        For i As Integer = 0 To UBound(HT_element)

            Fill_HT_element_argument(HT_element(i))

            ' Рассчитываем температуру элемента

            Temperature = HT_element(i).Element.Field.Temperature

            For j As Integer = 0 To UBound(HT_element(i).Face)

                ' Рассчитываем площадь 

                Area = HT_element(i).Face(j).Area

                ' Рассчитываем коэффициент черноты для длины волны, на которую приходится максимум теплового потока

                If HT_element(i).Face(j).FTOP.A.GetType.ToString <> "System.Double" Then

                    ReDim Argument.Coord(5)

                    Argument.Coord(0) = Temperature

                    'Argument.Coord(3) = HT_element(i).Face(j).c / HT_element(i).Face(j).Max_emittance_frequency(Temperature)

                    eps = HT_element(i).Face(j).FTOP.A.return_value(Argument).Coord(0)

                End If

                If HT_element(i).Face(j).FTOP.A.GetType.ToString = "System.Double" Then

                    eps = HT_element(i).Face(j).FTOP.A

                End If

                If HT_element(i).Face(j).Enabled Then


                    B_row = Illuminate_model(HT_element(i).Face(j), HT_element(i).Face(j).oz, Math.PI / 2, _
                                            HT_element(i).Face(j).FTOP.CDF, HT_element(i).Argument, _
                                            HT_element(i).Face(j).Global_medium, aux_Visual, N_photon)

                    For l As Integer = 0 To UBound(B_row)

                        B = B_row(l)

                        B_matrix(i, Me.FEM.Shining_face(l).HT_element.Number) += Area * eps * B

                    Next l

                End If


            Next j

            If Progress + 1 <= i / UBound(HT_element) * 100 Then

                Progress += 1

                RaiseEvent B_matrix_progress(Progress)

            End If

            Application.DoEvents()

        Next i

        Dim Element_B_row(UBound(HT_element)) As Double

        For i As Integer = 0 To UBound(FEM.RSOURCE)

            Fill_RSOURCE_argument(FEM.RSOURCE(i))

            With FEM.RSOURCE(i)

                .B_row = Illuminate_model(.Calc_position(.Argument), Glob_z, Math.PI, .CDF, .Argument, Global_medium, aux_Visual, .N_photon)

                For j As Integer = 0 To UBound(.B_row)

                    Element_B_row(FEM.Shining_face(j).HT_element.Number) += .B_row(j)

                Next j

                ReDim .B_row(UBound(HT_element))

                For j As Integer = 0 To UBound(.B_row)

                    .B_row(j) = Element_B_row(j)

                Next j

            End With

        Next i


        't_1 = timeGetTime

        'delta_t = t_1 - t_0

        'delta_t /= UBound(Me.HT_element)

        aux_Visual.frm.Close()

        aux_Visual = Nothing

        'MsgBox(delta_t)

    End Sub

    ''' <summary>
    ''' This subrutine prepares array for heat transfer calculation on Step_Number.
    ''' </summary>
    ''' <param name="Step_Number"></param>
    ''' <remarks></remarks>
    Public Sub Prepare_heat_transfer(ByRef Step_Number As Integer)

        For i As Integer = 0 To UBound(Me.HT_element)

            Fill_HT_element_argument(HT_element(i))

            ReDim Preserve HT_element(i).HT_Step(Step_Number)

            'HT_element(i).HT_Step(Step_Number).Q = 0

            'HT_element(i).HT_Step(Step_Number).T = 0

        Next i

        ' Prepare faces for radiation heat transfer calculation

        For i As Integer = 0 To UBound(FEM.Shining_face)

            FEM.Shining_face(i).Initialize()

        Next i

        For i As Integer = 0 To UBound(FEM.T_SET)

            For j As Integer = 0 To UBound(FEM.T_SET(i).Element)

                Number = FEM.T_SET(i).Element(j).HT_element.Number

                If FEM.T_SET(i).Temp(j).GetType.ToString = "System.Double" Then

                    HT_element(Number).HT_Step(Step_Number).T = FEM.T_SET(i).Temp(j)

                Else

                    HT_element(Number).HT_Step(Step_Number).T = FEM.T_SET(i).Temp(j).Return_value(HT_element(Number).Argument).Coord(0)

                End If

                HT_element(Number).Element.Field.Temperature = HT_element(Number).HT_Step(Step_Number).T

            Next j

        Next i

    End Sub

    ''' <summary>
    ''' This sub is used for raduation heat power caclulation on a given step of intergation.
    ''' The termalotical properties are variable at time of solution process. 
    ''' </summary>
    ''' <param name="Step_Number">Temperatures on this step are used for element emission calculation.</param>
    ''' <param name="N_photon">Numbers of photons emitted from each face on each step of radiation heat transfer calculation</param>
    ''' <param name="vf">Visualisation form used for radiaion heat transfer calculation.</param>
    ''' <remarks></remarks>
    Public Sub Calc_R_transfer_N_photon_old(ByRef Step_Number As Integer, ByRef N_photon As Integer, ByRef vf As Visualisation.Visualisation)

        Dim is_Double As Double

        Dim Empty_photon As New Photon

        ' Effective temperature of element
        Dim T_eff As Double

        Dim Heat_power_by_faces(,) As Double

        ' Faces outcoming heat power calculation 
        For i As Integer = 0 To UBound(FEM.Shining_face)

            With FEM.Shining_face(i)

                ' If face emits like a black body then
                If .FTOP.CDF Is Nothing Then

                    If .FTOP.A.GetType.ToString = is_Double.GetType.ToString Then

                        .Q_R_outcome = .Area * .FTOP.Eps * sigma * .HT_element.HT_step(Step_Number).T ^ 4

                    Else

                        .Element.Field.Temperature = .HT_element.HT_step(Step_Number).T

                        Fill_Face_argument(FEM.Shining_face(i), Empty_photon)

                        .Q_R_outcome = .Area * .FTOP.Calc_intergal_Eps(.Argument) * sigma * .HT_element.HT_step(Step_Number).T ^ 4

                    End If

                    ' If face emits NOT like a black body then 
                Else

                    ' It is needs to caclulate effective temperature of face 
                    .Element.Field.Temperature = .HT_element.HT_step(Step_Number).T

                    Fill_Face_argument(FEM.Shining_face(i), Empty_photon)

                    T_eff = .FTOP.T_eff.Return_value(.Argument)

                    ' and after it calculate outcoming heat power 

                    .Q_R_outcome = .Area * sigma * T_eff * T_eff * T_eff * T_eff

                End If

            End With

            Application.DoEvents()

        Next i

        ' Faces incoming heat power calculation 

        For i As Integer = 0 To UBound(FEM.Shining_face)

            With FEM.Shining_face(i)

                .Element.Field.Temperature = .HT_element.HT_step(Step_Number).T

                Fill_Face_argument(FEM.Shining_face(i), Empty_photon)

                'Heat_power_by_faces = Illuminate_model_and_calc_radiation_Q(.Center, .oz, Math.PI / 2, .FTOP.CDF, .Argument, .Global_medium, vf, N_photon, .Q_R_outcome)

                If .Q_R_outcome > 0 Then

                    Heat_power_by_faces = Illuminate_model_from_face_and_calc_radiation_Q(FEM.Shining_face(i), .oz, Math.PI / 2, .FTOP.CDF, .Argument, .Global_medium, vf, N_photon, .Q_R_outcome)

                    If Not (Heat_power_by_faces Is Nothing) Then

                        For j As Integer = 0 To UBound(Heat_power_by_faces, 2)

                            FEM.Shining_face(Heat_power_by_faces(0, j)).Q_R_income += Heat_power_by_faces(1, j)

                        Next j

                    End If


                End If


            End With

            Application.DoEvents()

        Next i

        ' Faces incoming from radiation heat source heat power calculation


        Dim Q_R_outcome As Double

        Dim R_source_position As New Vector

        For i As Integer = 0 To UBound(FEM.RSOURCE)

            Fill_RSOURCE_argument(FEM.RSOURCE(i))

            With FEM.RSOURCE(i)

                ' If radiation source emits like a black body then
                If .CDF Is Nothing Then

                    If .Q.GetType.ToString = is_Double.GetType.ToString Then

                        Q_R_outcome = .Q

                    Else

                        Q_R_outcome = .Calc_Q(.Argument)

                    End If

                    ' If radiation source emits NOT like a black body then
                Else

                    ' It is needs to caclulate effective temperature of face 

                    T_eff = .T.Return_value(.Argument)

                    ' and after it calculate outcoming heat power 

                    Q_R_outcome = sigma * T_eff * T_eff * T_eff * T_eff

                End If


                'Heat_power_by_faces = Illuminate_model_and_calc_radiation_Q(.Calc_position(.Argument), Glob_z, Math.PI, .CDF, .Argument, Global_medium, vf, .N_photon, Q_R_outcome)

                R_source_position = .Calc_position(.Argument)

                Heat_power_by_faces = Illuminate_model_and_calc_radiation_Q(R_source_position, Norm_vector(Me.FEM.Center - R_source_position), Math.PI, .CDF, .Argument, Global_medium, vf, .N_photon, Q_R_outcome)

                If Not (Heat_power_by_faces Is Nothing) Then

                    For j As Integer = 0 To UBound(Heat_power_by_faces, 2)

                        FEM.Shining_face(Heat_power_by_faces(0, j)).Q_R_income += Heat_power_by_faces(1, j)

                    Next j

                End If

            End With

            Application.DoEvents()

        Next i


        ' Gavering heat powers from faces to element

        Dim HT_element_Number As Integer

        For i As Integer = 0 To UBound(FEM.Shining_face)

            With FEM.Shining_face(i)

                HT_element_Number = .HT_element.Number

                HT_element(HT_element_Number).HT_Step(Step_Number + 1).Q += -FEM.Shining_face(i).Q_R_outcome + FEM.Shining_face(i).Q_R_income

            End With

        Next i

    End Sub


    ''' <summary>
    ''' This sub is used for raduation heat power caclulation on a given step of intergation.
    ''' The termalotical properties are variable at time of solution process. 
    ''' </summary>
    ''' <param name="Step_Number">Temperatures on this step are used for element emission calculation.</param>
    ''' <param name="N_photon">Numbers of photons emitted from each face on each step of radiation heat transfer calculation</param>
    ''' <param name="vf">Visualisation form used for radiaion heat transfer calculation.</param>
    ''' <remarks></remarks>
    Public Sub Calc_R_transfer_N_photon(ByRef Step_Number As Integer, ByRef N_photon As Integer, ByRef vf As Visualisation.Visualisation)


        'Dim writer_1 As System.IO.StreamWriter = New System.IO.StreamWriter("C:\tmp\Eps_v_T_300K.txt", False)

        'Dim Arg As New Vector

        'ReDim Arg.Coord(5)

        'Arg.Coord(0) = 300

        'Dim Eps As Double

        'For v As Double = 5 * 10 ^ 12 To 2 * 10 ^ 14 Step 5 * 10 ^ 12

        '    Arg.Coord(3) = v

        '    Eps = FEM.Shining_face(3).FTOP.Eps.Return_value(Arg).coord(0)

        '    writer_1.WriteLine(v.ToString.Replace(",", ".") & " " & Eps.ToString.Replace(",", "."))

        'Next v

        'writer_1.Close()



        Dim is_Double As Double

        'Illuminate_model_from_face_and_save_absorbed_power_in_faces(FEM.Shining_face(0), New Vector, 0.0, FEM.Dependency(0), New Vector, New MEDIUM1, New Visualisation.Visualisation, 1, 0.0)

        Dim Empty_photon As New Photon

        ' Effective temperature of element
        Dim T_eff As Double

        ' Clearning faces before usage

        Reset_fases_R_Q()

        ' Faces outcoming heat power calculation 
        For i As Integer = 0 To UBound(FEM.Shining_face)

            With FEM.Shining_face(i)

                ' If face emits like a black body then
                If .FTOP.CDF Is Nothing Then

                    If .FTOP.A.GetType.ToString = is_Double.GetType.ToString Then

                        .Q_R_outcome = .Area * .FTOP.Eps * sigma * .HT_element.HT_step(Step_Number).T ^ 4

                    Else

                        .Element.Field.Temperature = .HT_element.HT_step(Step_Number).T

                        Fill_Face_argument(FEM.Shining_face(i), Empty_photon)

                        .Q_R_outcome = .Area * .FTOP.Calc_intergal_Eps(.Argument) * sigma * .HT_element.HT_step(Step_Number).T ^ 4

                    End If

                    ' If face emits NOT like a black body then 
                Else

                    ' It is needs to caclulate effective temperature of face 
                    .Element.Field.Temperature = .HT_element.HT_step(Step_Number).T

                    Fill_Face_argument(FEM.Shining_face(i), Empty_photon)

                    T_eff = .FTOP.T_eff.Return_value(.Argument)

                    ' and after it calculate outcoming heat power 

                    .Q_R_outcome = .Area * sigma * T_eff * T_eff * T_eff * T_eff

                End If

            End With

            Application.DoEvents()

        Next i

        ' Faces incoming heat power calculation 

        For i As Integer = 0 To UBound(FEM.Shining_face)

            With FEM.Shining_face(i)

                .Element.Field.Temperature = .HT_element.HT_step(Step_Number).T

                Fill_Face_argument(FEM.Shining_face(i), Empty_photon)

                Illuminate_model_and_save_absorbed_power_in_faces(FEM.Shining_face(i), .oz, Math.PI / 2, .FTOP.CDF, .Argument, .Global_medium, vf, N_photon, .Q_R_outcome)

            End With

            Application.DoEvents()

        Next i

        ' Faces incoming from radiation heat source heat power calculation

        Dim Q_R_outcome As Double

        Dim R_source_position As New Vector

        For i As Integer = 0 To UBound(FEM.RSOURCE)

            Fill_RSOURCE_argument(FEM.RSOURCE(i))

            With FEM.RSOURCE(i)

                ' If radiation source emits like a black body then
                If .CDF Is Nothing Then

                    If .Q.GetType.ToString = is_Double.GetType.ToString Then

                        Q_R_outcome = .Q

                    Else

                        Q_R_outcome = .Calc_Q(.Argument)

                    End If

                    ' If radiation source emits NOT like a black body then
                Else

                    ' It is needs to caclulate effective temperature of face 

                    T_eff = .T.Return_value(.Argument)

                    ' and after it calculate outcoming heat power 

                    Q_R_outcome = sigma * T_eff * T_eff * T_eff * T_eff

                End If

                R_source_position = .Calc_position(.Argument)

                Illuminate_model_and_save_absorbed_power_in_faces(FEM.RSOURCE(i), Norm_vector(Me.FEM.Center - R_source_position), Math.PI, .CDF, .Argument, Global_medium, vf, .N_photon, Q_R_outcome)

            End With

            Application.DoEvents()

        Next i


        ' Gavering heat powers from faces to element

        Dim HT_element_Number As Integer

        For i As Integer = 0 To UBound(FEM.Shining_face)

            With FEM.Shining_face(i)

                HT_element_Number = .HT_element.Number

                HT_element(HT_element_Number).HT_Step(Step_Number + 1).Q += -.Q_R_outcome + (.Q_R_income - .Q_R_diffuse_reflected - .Q_R_mirror_reflected - .Q_R_transmitted)

            End With

        Next i

    End Sub

    ''' <summary>
    ''' !!!! Передалать в соотвествие Calc_R_transfer_N_photon с другим набором аргументов
    ''' </summary>
    ''' <param name="N_1">First HT_element number to calculate radiation heat transfer</param>
    ''' <param name="N_2">Last HT_element number to calculate radiation heat transfer</param>
    ''' <param name="Step_Number">Temperatures on this step are used for element emission calculation.</param>
    ''' <param name="N_photon">Numbers of photons emitted from each face on each step of radiation heat transfer calculation</param>
    ''' <param name="vf">Visualisation form used for radiaion heat transfer calculation.</param>
    ''' <remarks></remarks>
    Public Function Calc_R_transfer_N_photon(ByRef N_1 As Long, ByRef N_2 As Long, ByRef Step_Number As Integer, ByRef N_photon As Integer, ByRef vf As Visualisation.Visualisation) As Double(,)

        If HT_element Is Nothing Then

            Solution_integrity = False

            Return Nothing

        End If


        Dim is_Double As Double

        Dim Empty_photon As New Photon

        ' Effective temperature of element
        Dim T_eff As Double

        Dim Heat_power_by_faces(,) As Double

        ' Reseting faces heat power
        Reset_fases_R_Q()

        ' Faces outcoming heat power calculation 
        For i As Integer = N_1 To N_2

            For j As Integer = 0 To UBound(HT_element(i).Face)

                With HT_element(i).Face(j)

                    ' If face emits like a black body then
                    If .FTOP.CDF Is Nothing Then

                        If .FTOP.A.GetType.ToString = is_Double.GetType.ToString Then

                            .Q_R_outcome = .Area * .FTOP.Eps * .Sigma * .HT_element.HT_step(Step_Number).T * .HT_element.HT_step(Step_Number).T * .HT_element.HT_step(Step_Number).T * .HT_element.HT_step(Step_Number).T

                        Else

                            .Element.Field.Temperature = .HT_element.HT_step(Step_Number).T

                            Fill_Face_argument(HT_element(i).Face(j), Empty_photon)

                            .Q_R_outcome = .Area * .FTOP.Eps.Return_value(.Argument) * .Sigma * .HT_element.HT_step(Step_Number).T * .HT_element.HT_step(Step_Number).T * .HT_element.HT_step(Step_Number).T * .HT_element.HT_step(Step_Number).T

                        End If

                        ' If face emits NOT like a black body then 
                    Else

                        ' It is needs to caclulate effective temperature of face 
                        .Element.Field.Temperature = .HT_element.HT_step(Step_Number).T

                        Fill_Face_argument(HT_element(i).Face(j), Empty_photon)

                        T_eff = .FTOP.T_eff.Return_value(.Argument)

                        ' and after it calculate outcoming heat power 

                        .Q_R_outcome = .Area * .Sigma * T_eff * T_eff * T_eff * T_eff

                    End If

                End With

                Application.DoEvents()

            Next j

        Next i

        ' Faces incoming heat power calculation 

        For i As Integer = N_1 To N_2

            For j As Integer = 0 To UBound(HT_element(i).Face)

                With HT_element(i).Face(j)

                    .Element.Field.Temperature = .HT_element.HT_step(Step_Number).T

                    Fill_Face_argument(HT_element(i).Face(j), Empty_photon)

                    Heat_power_by_faces = Illuminate_model_from_face_and_calc_radiation_Q(HT_element(i).Face(j), .oz, Math.PI / 2, .FTOP.CDF, .Argument, .Global_medium, vf, N_photon, .Q_R_outcome)

                    If Not (Heat_power_by_faces Is Nothing) Then

                        For k As Integer = 0 To UBound(Heat_power_by_faces, 2)

                            FEM.Shining_face(Heat_power_by_faces(0, k)).Q_R_income += Heat_power_by_faces(1, k)

                        Next k

                    End If

                End With

                Application.DoEvents()

            Next j

        Next i

        ' Gavering heat powers from faces to element and forming array with heat powers

        Dim R_value(1, UBound(HT_element)) As Double

        Dim i_HT_element As Long = -1

        Dim Q_Final As Double

        For i As Integer = 0 To UBound(HT_element)

            Q_Final = 0

            For j As Integer = 0 To UBound(HT_element(i).Face)

                Q_Final += -HT_element(i).Face(j).Q_R_outcome + HT_element(i).Face(j).Q_R_income

            Next j

            If Q_Final <> 0 Then

                i_HT_element += 1

                R_value(0, i_HT_element) = i

                R_value(1, i_HT_element) = Q_Final

            End If

        Next i

        Application.DoEvents()

        ReDim Preserve R_value(1, i_HT_element)

        Return R_value

    End Function

    '''' <summary>
    '''' 
    '''' </summary>
    '''' <param name="N_1">First HT_element number to calculate radiation heat transfer</param>
    '''' <param name="N_2">Last HT_element number to calculate radiation heat transfer</param>
    '''' <param name="Step_Number">Temperatures on this step are used for element emission calculation.</param>
    '''' <param name="Step_Epsilon">Maximum available error in view-factor-vector calculation process</param>
    '''' <param name="vf">Visualisation form used for radiaion heat transfer calculation.</param>
    '''' <remarks></remarks>
    'Public Function Calc_R_transfer_TOP_Epsilon(ByRef N_1 As Long, ByRef N_2 As Long, ByRef Step_Number As Integer, ByRef Step_Epsilon As Double, ByRef vf As Visualisation.Visualisation) As Double(,)

    '    Dim is_Double As Double

    '    Dim Empty_photon As New Photon

    '    ' Effective temperature of element
    '    Dim T_eff As Double

    '    Dim Heat_power_by_faces(,) As Double

    '    ' Reseting faces heat power
    '    Reset_fases_R_Q()

    '    ' Faces outcoming heat power calculation 
    '    For i As Integer = N_1 To N_2

    '        For j As Integer = 0 To UBound(HT_element(i).Face)

    '            With HT_element(i).Face(j)

    '                ' If face emits like a black body then
    '                If .FTOP.CDF Is Nothing Then

    '                    If .FTOP.A.GetType.ToString = is_Double.GetType.ToString Then

    '                        .Q_R_outcome = .Area * .FTOP.A * .Sigma * .HT_element.HT_step(Step_Number).T * .HT_element.HT_step(Step_Number).T * .HT_element.HT_step(Step_Number).T * .HT_element.HT_step(Step_Number).T

    '                    Else

    '                        .Element.Field.Temperature = .HT_element.HT_step(Step_Number).T

    '                        Fill_Face_argument(HT_element(i).Face(j), Empty_photon)

    '                        .Q_R_outcome = .Area * .FTOP.A.Return_value(.Argument) * .Sigma * .HT_element.HT_step(Step_Number).T * .HT_element.HT_step(Step_Number).T * .HT_element.HT_step(Step_Number).T * .HT_element.HT_step(Step_Number).T

    '                    End If

    '                    ' If face emits NOT like a black body then 
    '                Else

    '                    ' It is needs to caclulate effective temperature of face 
    '                    .Element.Field.Temperature = .HT_element.HT_step(Step_Number).T

    '                    Fill_Face_argument(HT_element(i).Face(j), Empty_photon)

    '                    T_eff = .FTOP.T_eff.Return_value(.Argument)

    '                    ' and after it calculate outcoming heat power 

    '                    .Q_R_outcome = .Area * .Sigma * T_eff * T_eff * T_eff * T_eff

    '                End If

    '            End With

    '            Application.DoEvents()

    '        Next j

    '    Next i

    '    ' Faces incoming heat power calculation 

    '    For i As Integer = N_1 To N_2

    '        For j As Integer = 0 To UBound(HT_element(i).Face)

    '            With HT_element(i).Face(j)

    '                .Element.Field.Temperature = .HT_element.HT_step(Step_Number).T

    '                Fill_Face_argument(HT_element(i).Face(j), Empty_photon)

    '                Heat_power_by_faces = Illuminate_model_and_calc_radiation_Q(.Center, .oz, Math.PI / 2, .FTOP.CDF, .Argument, .Global_medium, vf, Step_Epsilon, .Q_R_outcome)

    '                If Not (Heat_power_by_faces Is Nothing) Then

    '                    For k As Integer = 0 To UBound(Heat_power_by_faces, 2)

    '                        FEM.Shining_face(Heat_power_by_faces(0, k)).Q_R_income += Heat_power_by_faces(1, k)

    '                    Next k

    '                End If

    '            End With

    '            Application.DoEvents()

    '        Next j

    '    Next i

    '    ' Gavering heat powers from faces to element and forming array with heat powers

    '    Dim R_value(1, UBound(HT_element)) As Double

    '    Dim i_HT_element As Long = -1

    '    Dim Q_Final As Double

    '    For i As Integer = 0 To UBound(HT_element)

    '        Q_Final = 0

    '        For j As Integer = 0 To UBound(HT_element(i).Face)

    '            Q_Final += -HT_element(i).Face(j).Q_R_outcome + HT_element(i).Face(j).Q_R_income

    '        Next j

    '        If Q_Final <> 0 Then

    '            i_HT_element += 1

    '            R_value(0, i_HT_element) = i

    '            R_value(1, i_HT_element) = Q_Final

    '        End If

    '    Next i

    '    ReDim Preserve R_value(1, i_HT_element)

    '    Return R_value

    'End Function

    ''' <summary>
    ''' This sub calculates heat power from radiation heat source to faces.
    ''' Each radiation source emits N_photon photons.
    ''' </summary>
    ''' <param name="N_photon"></param>
    ''' <param name="vf"></param>
    ''' <remarks></remarks>
    Public Function Calc_R_transfer_R_heat_source_N_photon(ByRef N_photon As Integer, ByRef vf As Visualisation.Visualisation) As Double(,)

        Dim is_Double As Double

        Dim Empty_photon As New Photon

        ' Effective temperature of element
        Dim T_eff As Double

        Dim Heat_power_by_faces(,) As Double

        Dim R_value(1, UBound(HT_element)) As Double

        Dim i_HT_element As Long = -1

        Dim Q_Final As Double

        ' Reseting faces heat power
        Reset_fases_R_Q()

        ' Faces incoming from radiation heat source heat power calculation

        Dim Q_R_outcome As Double

        For i As Integer = 0 To UBound(FEM.RSOURCE)

            Fill_RSOURCE_argument(FEM.RSOURCE(i))

            With FEM.RSOURCE(i)

                ' If radiation source emits like a black body then
                If .CDF Is Nothing Then

                    If .Q.GetType.ToString = is_Double.GetType.ToString Then

                        Q_R_outcome = .Q

                    Else

                        Q_R_outcome = .Calc_Q(.Argument)

                    End If

                    ' If radiation source emits NOT like a black body then
                Else

                    ' It is needs to caclulate effective temperature of face 

                    T_eff = .T.Return_value(.Argument)

                    ' and after it calculate outcoming heat power 

                    Q_R_outcome = sigma * T_eff * T_eff * T_eff * T_eff

                End If


                Heat_power_by_faces = Illuminate_model_and_calc_radiation_Q(.Calc_position(.Argument), Glob_z, Math.PI, .CDF, .Argument, Global_medium, vf, N_photon, Q_R_outcome)

                If Not (Heat_power_by_faces Is Nothing) Then

                    For j As Integer = 0 To UBound(Heat_power_by_faces, 2)

                        FEM.Shining_face(Heat_power_by_faces(0, j)).Q_R_income += Heat_power_by_faces(1, j)

                    Next j

                End If

            End With

            Application.DoEvents()

        Next i

        ' Gavering heat powers from faces to element and forming array with heat powers

        For i As Integer = 0 To UBound(HT_element)

            Q_Final = 0

            For j As Integer = 0 To UBound(HT_element(i).Face)

                Q_Final += -HT_element(i).Face(j).Q_R_outcome + HT_element(i).Face(j).Q_R_income

            Next j

            If Q_Final <> 0 Then

                i_HT_element += 1

                R_value(0, i_HT_element) = i

                R_value(1, i_HT_element) = Q_Final

            End If

        Next i

        ReDim Preserve R_value(1, i_HT_element)

        Return R_value

    End Function


    '''' <summary>
    '''' This sub calculates heat power from radiation heat source to faces.
    '''' Step_epsilon i used to exit photon emission cycle.
    '''' </summary>
    '''' <param name="Step_epsilon"></param>
    '''' <param name="vf"></param>
    '''' <remarks></remarks>
    'Public Function Calc_R_transfer_R_heat_source_TOP_Epsilon(ByRef Step_epsilon As Double, ByRef vf As Visualisation.Visualisation)

    '    Dim is_Double As Double

    '    Dim Empty_photon As New Photon

    '    ' Effective temperature of element
    '    Dim T_eff As Double

    '    Dim Heat_power_by_faces(,) As Double

    '    Dim R_value(1, UBound(HT_element)) As Double

    '    Dim i_HT_element As Long = -1

    '    Dim Q_Final As Double

    '    ' Reseting faces heat power
    '    Reset_fases_R_Q()

    '    ' Faces incoming from radiation heat source heat power calculation


    '    Dim Q_R_outcome As Double

    '    For i As Integer = 0 To UBound(FEM.RSOURCE)

    '        Fill_RSOURCE_argument(FEM.RSOURCE(i))

    '        With FEM.RSOURCE(i)

    '            ' If radiation source emits like a black body then
    '            If .CDF Is Nothing Then

    '                If .Q.GetType.ToString = is_Double.GetType.ToString Then

    '                    Q_R_outcome = .Q

    '                Else

    '                    Q_R_outcome = .Calc_Q(.Argument)

    '                End If

    '                ' If radiation source emits NOT like a black body then
    '            Else

    '                ' It is needs to caclulate effective temperature of face 

    '                T_eff = .T.Return_value(.Argument)

    '                ' and after it calculate outcoming heat power 

    '                Q_R_outcome = .sigma * T_eff * T_eff * T_eff * T_eff

    '            End If


    '            Heat_power_by_faces = Illuminate_model_and_calc_radiation_Q(.Calc_position(.Argument), Glob_z, Math.PI, .CDF, .Argument, Global_medium, vf, Step_epsilon, Q_R_outcome)

    '            If Not (Heat_power_by_faces Is Nothing) Then

    '                For j As Integer = 0 To UBound(Heat_power_by_faces, 2)

    '                    FEM.Shining_face(Heat_power_by_faces(0, j)).Q_R_income += Heat_power_by_faces(1, j)

    '                Next j

    '            End If

    '        End With

    '        Application.DoEvents()

    '    Next i

    '    ' Gavering heat powers from faces to element and forming array with heat powers

    '    For i As Integer = 0 To UBound(HT_element)

    '        Q_Final = 0

    '        For j As Integer = 0 To UBound(HT_element(i).Face)

    '            Q_Final += -HT_element(i).Face(j).Q_R_outcome + HT_element(i).Face(j).Q_R_income

    '        Next j

    '        If Q_Final <> 0 Then

    '            i_HT_element += 1

    '            R_value(0, i_HT_element) = i

    '            R_value(1, i_HT_element) = Q_Final

    '        End If

    '    Next i

    '    ReDim Preserve R_value(1, i_HT_element)

    '    Return R_value


    'End Function

    ''' <summary>
    ''' This sub adds heat powers from input array to Heat Transferring Element.
    ''' About Step_number. Heat power in real is addet to HT_Step(Step_number + 1).Q.
    ''' The array structure is:
    ''' Heat_power(0,x) - HT_element number,
    ''' Heat_power(1,x) - HT_element heat power.
    ''' </summary>
    ''' <param name="Heat_power">Heat power by elements</param>
    ''' <remarks></remarks>
    Public Sub Add_heat_power_from_array_to_HT_elements(ByRef Heat_power(,) As Double, ByRef Step_number As Integer)

        If Heat_power Is Nothing Then Exit Sub

        If Heat_power.Length = 0 Then Exit Sub

        For i As Integer = 0 To UBound(Heat_power, 2)

            HT_element(Heat_power(0, i)).HT_Step(Step_number + 1).Q += Heat_power(1, i)

        Next i

    End Sub


    'Public Sub Calc_R_transfer_TOP_Epsilon(ByRef Step_Number As Integer, ByRef Step_epsilon As Double, ByRef vf As Visualisation.Visualisation)

    '    Dim is_Double As Double

    '    Dim Empty_photon As New Photon

    '    ' Effective temperature of element
    '    Dim T_eff As Double

    '    Dim Heat_power_by_faces(,) As Double

    '    ' Faces outcoming heat power calculation 
    '    For i As Integer = 0 To UBound(FEM.Shining_face)

    '        With FEM.Shining_face(i)

    '            ' If face emits like a black body then
    '            If .FTOP.CDF Is Nothing Then

    '                If .FTOP.A.GetType.ToString = is_Double.GetType.ToString Then

    '                    .Q_R_outcome = .Area * .FTOP.A * .Sigma * .HT_element.HT_step(Step_Number).T * .HT_element.HT_step(Step_Number).T * .HT_element.HT_step(Step_Number).T * .HT_element.HT_step(Step_Number).T

    '                Else

    '                    .Element.Field.Temperature = .HT_element.HT_step(Step_Number).T

    '                    Fill_Face_argument(FEM.Shining_face(i), Empty_photon)

    '                    .Q_R_outcome = .Area * .FTOP.A.Return_value(.Argument) * .Sigma * .HT_element.HT_step(Step_Number).T * .HT_element.HT_step(Step_Number).T * .HT_element.HT_step(Step_Number).T * .HT_element.HT_step(Step_Number).T

    '                End If

    '                ' If face emits NOT like a black body then 
    '            Else

    '                ' It is needs to caclulate effective temperature of face 
    '                .Element.Field.Temperature = .HT_element.HT_step(Step_Number).T

    '                Fill_Face_argument(FEM.Shining_face(i), Empty_photon)

    '                T_eff = .FTOP.T_eff.Return_value(.Argument)

    '                ' and after it calculate outcoming heat power 

    '                .Q_R_outcome = .Area * .Sigma * T_eff * T_eff * T_eff * T_eff

    '            End If

    '        End With

    '        Application.DoEvents()

    '    Next i

    '    ' Faces incoming heat power calculation 

    '    For i As Integer = 0 To UBound(FEM.Shining_face)

    '        With FEM.Shining_face(i)

    '            .Element.Field.Temperature = .HT_element.HT_step(Step_Number).T

    '            Fill_Face_argument(FEM.Shining_face(i), Empty_photon)

    '            'Heat_power_by_faces = Illuminate_model_and_calc_radiation_Q(.Center, .oz, Math.PI / 2, .FTOP.CDF, .Argument, .Global_medium, vf, Step_epsilon, .Q_R_outcome)

    '            Heat_power_by_faces = Illuminate_model_from_face_and_calc_radiation_Q(FEM.Shining_face(i), .oz, Math.PI / 2, .FTOP.CDF, .Argument, .Global_medium, vf, Step_epsilon, .Q_R_outcome)

    '            If Not (Heat_power_by_faces Is Nothing) Then

    '                For j As Integer = 0 To UBound(Heat_power_by_faces, 2)

    '                    FEM.Shining_face(Heat_power_by_faces(0, j)).Q_R_income += Heat_power_by_faces(1, j)

    '                Next j

    '            End If

    '        End With

    '        Application.DoEvents()

    '    Next i

    '    ' Faces incoming from radiation heat source heat power calculation


    '    Dim Q_R_outcome As Double

    '    Dim R_source_position As New Vector

    '    For i As Integer = 0 To UBound(FEM.RSOURCE)

    '        Fill_RSOURCE_argument(FEM.RSOURCE(i))

    '        With FEM.RSOURCE(i)

    '            ' If radiation source emits like a black body then
    '            If .CDF Is Nothing Then

    '                If .Q.GetType.ToString = is_Double.GetType.ToString Then

    '                    Q_R_outcome = .Q

    '                Else

    '                    Q_R_outcome = .Calc_Q(.Argument)

    '                End If

    '                ' If radiation source emits NOT like a black body then
    '            Else

    '                ' It is needs to caclulate effective temperature of face 

    '                T_eff = .T.Return_value(.Argument)

    '                ' and after it calculate outcoming heat power 

    '                Q_R_outcome = .sigma * T_eff * T_eff * T_eff * T_eff

    '            End If


    '            'Heat_power_by_faces = Illuminate_model_and_calc_radiation_Q(.Calc_position(.Argument), Glob_z, Math.PI, .CDF, .Argument, Global_medium, vf, .Epsilon, Q_R_outcome)

    '            R_source_position = .Calc_position(.Argument)

    '            Heat_power_by_faces = Illuminate_model_and_calc_radiation_Q(R_source_position, Norm_vector(Me.FEM.Center - R_source_position), Math.PI / 2, .CDF, .Argument, Global_medium, vf, .N_photon, Q_R_outcome)

    '            If Not (Heat_power_by_faces Is Nothing) Then

    '                For j As Integer = 0 To UBound(Heat_power_by_faces, 2)

    '                    FEM.Shining_face(Heat_power_by_faces(0, j)).Q_R_income += Heat_power_by_faces(1, j)

    '                Next j

    '            End If

    '        End With

    '        Application.DoEvents()

    '    Next i


    '    ' Gavering heat powers from faces to element

    '    Dim HT_element_Number As Integer

    '    For i As Integer = 0 To UBound(FEM.Shining_face)

    '        With FEM.Shining_face(i)

    '            HT_element_Number = .HT_element.Number

    '            HT_element(HT_element_Number).HT_Step(Step_Number + 1).Q += -FEM.Shining_face(i).Q_R_outcome + FEM.Shining_face(i).Q_R_income

    '        End With

    '    Next i

    'End Sub
    ''' <summary>
    ''' This sub set incoming and outcoming radiation heat power of all of shinig faces
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Reset_fases_R_Q()

        For i As Integer = 0 To UBound(FEM.Shining_face)

            With FEM.Shining_face(i)

                .Q_R_outcome = 0

                .Q_R_income = 0

                .Q_R_absorbed = 0

                .Q_R_diffuse_reflected = 0

                .Q_R_mirror_reflected = 0

                .Q_R_transmitted = 0

            End With

        Next i


    End Sub

    ''' <summary>
    ''' This sub calculates heat powers caused by thermal conductivity
    ''' </summary>
    ''' <param name="Step_Number"></param>
    ''' <remarks></remarks>
    Public Sub Calc_C_transfer(ByRef Step_Number As Integer)

        Dim Temperature_1, Temperature_2 As Double

        Dim lambda_1, lambda_2 As Double

        For i As Integer = 0 To UBound(HT_element)

            If HT_element(i).T_SET Is Nothing Then

                Temperature_1 = HT_element(i).HT_Step(Step_Number).T

                ' Рассчитываем тепловой поток от соседних элементов

                For j As Integer = 0 To UBound(HT_element(i).Heat_conductivity)

                    With HT_element(i).Heat_conductivity(j)

                        ' Рассчитываем аргумент соседа

                        Fill_HT_element_argument(.HT_Element)

                        ' Рассчитываем температуру элемента соседа

                        Temperature_2 = HT_element(i).Heat_conductivity(j).HT_Element.HT_Step(Step_Number).T

                        ' Рассчитываем коэффциент теплопроводности для теплобмена

                        lambda_1 = .HT_Element.Element.Medium.Calc_lambda(HT_element(i).Argument)

                        lambda_2 = .HT_Element.Element.Medium.Calc_lambda(.HT_Element.Argument)

                        ' Расcчитываем тепловую мощность приходящую от соседнего элемента

                        HT_element(i).HT_Step(Step_Number + 1).Q += .Area * (lambda_1 * .Distance_1 / (.Distance_1 + .Distance_2) + lambda_2 * .Distance_2 / (.Distance_1 + .Distance_2)) * (Temperature_2 - Temperature_1) / (.Distance_1 + .Distance_2)

                    End With

                Next j

                ' рассчитываем тепловой поток от источников тепла

                If UBound(HT_element(i).Heat_source) >= 0 Then

                    For j As Integer = 0 To UBound(HT_element(i).Heat_source)

                        ' Тепловой поток от источника может зависеть только от времени, поэтому можно передавать аргумент от элемента
                        HT_element(i).HT_Step(Step_Number + 1).Q += HT_element(i).Heat_source(j).Heat(HT_element(i).Argument)

                    Next j

                End If

            End If

        Next i

    End Sub

    '''' <summary>
    '''' This function returns one row of radiation heat transfer matrix
    '''' </summary>
    '''' <param name="Face"> The source of photons</param>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public Function Fill_one_B_row(ByRef Face As Face, ByRef vf As Visualisation.Visualisation) As Integer

    '    Dim i_Photon_Min As Integer = 1 / Me.TOP_Epsilon

    '    Dim Maximum_founded As Boolean = False

    '    Dim Max_old, Max_new As Double

    '    Max_new = Double.MinValue

    '    Dim Max_i As Long

    '    'Dim row(UBound(Me.FEM.Shining_face)) As Double

    '    'объявляем переменную, в которой будем хранить погрешность вычисления строчки
    '    Dim Row_error As Double

    '    Dim i_Photon As Integer = 0

    '    ' номер грани, в которую летит фотон
    '    Dim Face_number As Long

    '    Dim Photon As Photon

    '    ' сбрасываем все счетчики

    '    For i As Integer = 0 To UBound(Me.FEM.Shining_face)

    '        Me.FEM.Shining_face(i).Initialize()

    '    Next i

    '    ' начинаем цикл излучения фотонов

    '    Do

    '        i_Photon += 1

    '        ' генерируем случайный фотон

    '        'Photon = Face.Emit_random_photon()

    '        ' начинаем цикл его распространения
    '        Do

    '            ' смотрим есть ли какая либо грань в его направлении

    '            'Face_number = View_to_vector(vf, Photon.Last_face.Center + Me.FEM.Minimum_length * 0.001 * Photon.Last_face.oz, Photon.Last_face.Center + Photon.Direction)

    '            'Face_number = View_to_vector(vf, Photon.Last_face.Center + 0.001 * Photon.Last_face.oz, Me.FEM.Shining_face(201).Center)

    '            If Face_number <> 0 Then

    '                Fill_Face_argument(FEM.Shining_face(Face_number - 1), Photon)

    '                Photon = FEM.Shining_face(Face_number - 1).Face_photon_interaction(Photon)

    '            End If

    '            ' если взаимодействие не вернуло фотона (он поглотился) или фотон улетел в бесконечность,
    '            ' то значит нужно вычислять строчку
    '            If Photon Is Nothing Or Face_number = 0 Then

    '                If i_Photon < i_Photon_Min Then

    '                    Row_error = Double.MaxValue

    '                Else

    '                    ' ищем максимальный элемент и по нему вычисляем ошибку

    '                    If Not (Maximum_founded) Then

    '                        For i As Integer = 0 To UBound(Me.FEM.Shining_face)

    '                            If Max_new < Me.FEM.Shining_face(i).Interaction.N_A / i_Photon Then

    '                                Max_new = Me.FEM.Shining_face(i).Interaction.N_A / i_Photon

    '                                Max_i = i

    '                            End If

    '                        Next i

    '                        Max_old = Max_new

    '                        Maximum_founded = True

    '                        If Max_new = 0 Then Row_error = 0

    '                    Else

    '                        Max_new = Me.FEM.Shining_face(Max_i).Interaction.N_A / i_Photon

    '                        If Max_new = Max_old Then

    '                            Row_error = Double.MaxValue

    '                        Else

    '                            Row_error = Math.Abs((Max_new - Max_old) / Max_new)

    '                            Max_old = Max_new

    '                        End If

    '                    End If

    '                End If

    '                Exit Do

    '            End If

    '        Loop

    '        If Row_error <= TOP_Epsilon Then

    '            Exit Do

    '        End If

    '    Loop


    '    Return i_Photon


    'End Function

    ''' <summary>
    ''' This function illiminate model with photons, emutted from point Source_location,
    ''' in the line of Direction with Opening_angle.
    ''' The function returns array of parts of energy absorbed in model's elements.    
    ''' </summary>
    ''' <param name="Source_location">Location of radiotion source</param>
    ''' <param name="Direction">Direction of radiation propagation</param>
    ''' <param name="Opening_angle">Opening angle in radiation propagation cone
    ''' If Opening_angle = PI, photons emitted in all of directions, spherically.
    ''' If Opening_angle = PI/2, photons emitted in one half-space.
    ''' If Opening_angle more then 0 and less than  PI/2, photons emitted in cone.</param>
    ''' <param name="CDF">Radiation source cumulative distribution of photon frequency</param>
    ''' <param name="Argument">Radiation source paratemers, such as a temperature, time and so on.</param>
    ''' <param name="Emitter_medium">Medium, surrounded radiation source.</param>
    ''' <param name="vf">Visualisation form, used for photon-face interaction detection</param>
    ''' <param name="N_photon">Number of photons emitted from source</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Illuminate_model(ByRef Source_location As Vector, ByRef Direction As Vector, ByRef Opening_angle As Double, ByRef CDF As DEPEND, ByRef Argument As Vector, ByRef Emitter_medium As MEDIUM1, ByRef vf As Visualisation.Visualisation, ByRef N_photon As Integer) As Double()

        Dim Direction_should_be_reverted As Boolean = False

        ' Vector of  current absorption coefficients
        Dim Steady_state_vector_1 As Vector

        Steady_state_vector_1 = New Vector

        ReDim Steady_state_vector_1.Coord(UBound(FEM.Shining_face))

        ' Previous and current vector difference. Vector compopent, whose difference is less than set maximum
        ' are excluded from error difference calculation
        Dim Vector_error As Double = Double.MaxValue


        ' Feed, illimination angle calculation

        Dim Limit_distance As Double = 2 * FEM.Maximum_length

        ' If radiation source is located farther then Limit_distance, then it can be shifted to it distance.

        ' If radiation source is located farther then Limit_distance, then its opening angle should be shifted to Math.Atan(FEM.Maximum_length/Limit_distance).

        ' If radiation source is shifted to new position, it is needed to correct absorption in model elements.
        ' It is caused by part of photons guarantily passed away from model and can't be absorbed.

        Dim to_Center As Vector = FEM.Center - Source_location

        Dim Source_distance As Double = L_vector(to_Center)

        Dim New_direction As New Vector

        Dim New_opening_angle As Double

        ' This coefficient is used for absorption correction due to illimination angle changing
        Dim Correction_coef As Double

        If Source_distance > Limit_distance Then

            New_opening_angle = Math.Atan(FEM.Maximum_length / 2 / Source_distance)

            Correction_coef = Math.Sin(0.5 * New_opening_angle) ^ 2

            New_direction = to_Center

            If Math.Abs(New_direction.Coord(0)) < Epsilon Then New_direction.Coord(0) = 0

            If Math.Abs(New_direction.Coord(1)) < Epsilon Then New_direction.Coord(1) = 0

            If Math.Abs(New_direction.Coord(2)) < Epsilon Then New_direction.Coord(2) = 0


        Else

            'New_source_location = Source_location

            New_opening_angle = Opening_angle

            If New_opening_angle = Math.PI Then

                New_opening_angle = 0.5 * Math.PI

                Direction_should_be_reverted = True

            End If

            New_direction = Direction

            Correction_coef = 1

        End If

        ' Coordinate system building

        Dim ox, oy, oz As Vector

        oz = Norm_vector(New_direction)

        If oz.Coord(0) = 0 And oz.Coord(1) = 0 And Math.Abs(oz.Coord(2)) = 1 Then

            ox = New Vector

            ox.Coord(0) = 1

            ox.Coord(1) = 0

            ox.Coord(2) = 0


            oy = New Vector

            oy.Coord(0) = 0

            oy.Coord(1) = 1

            oy.Coord(2) = 0

        Else

            oy = New Vector

            oy.Coord(0) = 0

            oy.Coord(1) = 0

            oy.Coord(2) = 1

            ox = oy ^ oz

            oy = oz ^ ox

            ox = Norm_vector(ox)

            oy = Norm_vector(oy)

            oz = Norm_vector(oz)

        End If

        ' Model initializatioan and clearing

        For i As Integer = 0 To UBound(FEM.Shining_face)

            'ReDim FEM.Shining_face(i).Absorption_act(Max_items_default)

            'FEM.Shining_face(i).i_Absorption_act = -1

            FEM.Shining_face(i).Interaction = New Face_photon_interaction

        Next i


        ' Photon emission cycle
        Dim Photon As Photon = New Photon

        ' Face, hitted by photon
        Dim Face_number As Integer

        Dim i_Face_with_absorbtion_number_0 As Integer = -1

        Dim i_Face_with_absorbtion_number_1 As Integer = -1

        Dim i_Photon As Integer = 0

        Dim view_from As Vector

        Dim view_to As Vector

        ' Array of photons to emitt

        Dim Photons(N_photon - 1) As Photon

        Photons = Photon.Generate_N_random_photons(Source_location, N_photon, New_opening_angle, CDF, Argument, Emitter_medium, 0)

        For i As Integer = 0 To N_photon - 1

            ' Emitting random photon

            i_Photon += 1

            Photon = Photons(i)

            'Photon = New Photon

            'Photon = Photon.Emit_random_photon(Source_location, New_opening_angle, CDF, Argument)

            'Photon.Number = i_Photon

            'Photon.Medium = Emitter_medium

            ' Transforming photon coordinates from emitter coordinate system to global coordinate system

            Photon.Direction = Photon.Direction.Coord(0) * ox + Photon.Direction.Coord(1) * oy + Photon.Direction.Coord(2) * oz

            ' Propagation cycle

            ' First photon emitted from corrected emitter position

            If Source_distance > Limit_distance Then

                Photon.Center = Source_location + (Source_distance - FEM.Maximum_length) * Photon.Direction

            End If

            ' if we have to emit spherically  
            If Direction_should_be_reverted Then

                ' we have to reverse emission direction after each proton to fill the sphere 
                ox = -1 * ox

                oy = -1 * oy

                oz = -1 * oz

            End If

            Do

                view_from = Photon.Center + FEM.Minimum_length * 0.001 * Photon.Direction

                view_to = Photon.Center + Photon.Direction

                'Dim writer_1 As System.IO.StreamWriter = New System.IO.StreamWriter("C:\tmp\View_from_view_to.txt", True)

                'writer_1.WriteLine(Str(view_from.Coord(0)) & " " & Str(view_from.Coord(1)) & " " & Str(view_from.Coord(2)) & " " & Str(view_to.Coord(0)) & " " & Str(view_to.Coord(1)) & " " & Str(view_to.Coord(2)))

                ''writer_1.WriteLine(Str(Photon.Center.Coord(0)) & " " & Str(Photon.Center.Coord(1)) & " " & Str(Photon.Center.Coord(2)))

                ''writer_1.WriteLine(Str(Photon.Direction.Coord(0)) & " " & Str(Photon.Direction.Coord(1)) & " " & Str(Photon.Direction.Coord(2)))


                'writer_1.Close()


                'Face_number = View_to_vector(vf, Photon.Center + FEM.Minimum_length * 0.001 * Photon.Direction, Photon.Center + Photon.Direction)

                Face_number = vf.View_to_vector(view_from.Coord(0), view_from.Coord(1), view_from.Coord(2), view_to.Coord(0), view_to.Coord(1), view_to.Coord(2))

                If Face_number <> 0 Then


                    'Dim writer As System.IO.StreamWriter = New System.IO.StreamWriter("C:\tmp\Faces_numbers.txt", True)

                    'writer.WriteLine(Face_number)

                    'writer.Close()

                    Fill_Face_argument(FEM.Shining_face(Face_number - 1), Photon)

                    Photon = FEM.Shining_face(Face_number - 1).Face_photon_interaction(Photon)

                    ' If photon absorbed, exit cycle
                    If Photon Is Nothing Then

                        Exit Do

                    End If

                Else

                    ' If photon is passed away from model, exit cycle

                    Exit Do

                End If

            Loop

        Next i


        For i As Integer = 0 To UBound(FEM.Shining_face)

            With FEM.Shining_face(i)

                Steady_state_vector_1.Coord(i) = Correction_coef * (.Interaction.N_R_dif + .Interaction.N_R_mir + .Interaction.N_D + .Interaction.N_A) / i_Photon * .FTOP.A

            End With

        Next i

        Return Steady_state_vector_1.Coord

    End Function

    ''' <summary>
    ''' This function illiminate model with photons, emutted from point Source_location,
    ''' in the line of Direction with Opening_angle.
    ''' The function returns array of parts of energy absorbed in model's elements.    
    ''' </summary>
    ''' <param name="Face">Face to emit photons</param>
    ''' <param name="Direction">Direction of radiation propagation</param>
    ''' <param name="Opening_angle">Opening angle in radiation propagation cone
    ''' If Opening_angle = PI, photons emitted in all of directions, spherically.
    ''' If Opening_angle = PI/2, photons emitted in one half-space.
    ''' If Opening_angle more then 0 and less than  PI/2, photons emitted in cone.</param>
    ''' <param name="CDF">Radiation source cumulative distribution of photon frequency</param>
    ''' <param name="Argument">Radiation source paratemers, such as a temperature, time and so on.</param>
    ''' <param name="Emitter_medium">Medium, surrounded radiation source.</param>
    ''' <param name="vf">Visualisation form, used for photon-face interaction detection</param>
    ''' <param name="N_photon">Number of photons emitted from source</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Illuminate_model(ByRef Face As Face, ByRef Direction As Vector, ByRef Opening_angle As Double, ByRef CDF As DEPEND, ByRef Argument As Vector, ByRef Emitter_medium As MEDIUM1, ByRef vf As Visualisation.Visualisation, ByRef N_photon As Integer) As Double()

        Dim Direction_should_be_reverted As Boolean = False

        ' Vector of  current absorption coefficients
        Dim Steady_state_vector_1 As Vector

        Steady_state_vector_1 = New Vector

        ReDim Steady_state_vector_1.Coord(UBound(FEM.Shining_face))

        ' Previous and current vector difference. Vector compopent, whose difference is less than set maximum
        ' are excluded from error difference calculation
        Dim Vector_error As Double = Double.MaxValue


        ' Feed, illimination angle calculation

        Dim Limit_distance As Double = 2 * FEM.Maximum_length

        ' If radiation source is located farther then Limit_distance, then it can be shifted to it distance.

        ' If radiation source is located farther then Limit_distance, then its opening angle should be shifted to Math.Atan(FEM.Maximum_length/Limit_distance).

        ' If radiation source is shifted to new position, it is needed to correct absorption in model elements.
        ' It is caused by part of photons guarantily passed away from model and can't be absorbed.

        Dim to_Center As Vector = FEM.Center - Face.Center

        Dim Source_distance As Double = L_vector(to_Center)

        Dim New_direction As New Vector

        Dim New_opening_angle As Double

        ' This coefficient is used for absorption correction due to illimination angle changing
        Dim Correction_coef As Double

        If Source_distance > Limit_distance Then

            New_opening_angle = Math.Atan(FEM.Maximum_length / 2 / Source_distance)

            Correction_coef = Math.Sin(0.5 * New_opening_angle) ^ 2

            New_direction = to_Center

            If Math.Abs(New_direction.Coord(0)) < Epsilon Then New_direction.Coord(0) = 0

            If Math.Abs(New_direction.Coord(1)) < Epsilon Then New_direction.Coord(1) = 0

            If Math.Abs(New_direction.Coord(2)) < Epsilon Then New_direction.Coord(2) = 0


        Else

            'New_source_location = Source_location

            New_opening_angle = Opening_angle

            If New_opening_angle = Math.PI Then

                New_opening_angle = 0.5 * Math.PI

                Direction_should_be_reverted = True

            End If

            New_direction = Direction

            Correction_coef = 1

        End If

        ' Coordinate system building

        Dim ox, oy, oz As Vector

        oz = Norm_vector(New_direction)

        If oz.Coord(0) = 0 And oz.Coord(1) = 0 And Math.Abs(oz.Coord(2)) = 1 Then

            ox = New Vector

            ox.Coord(0) = 1

            ox.Coord(1) = 0

            ox.Coord(2) = 0


            oy = New Vector

            oy.Coord(0) = 0

            oy.Coord(1) = 1

            oy.Coord(2) = 0

        Else

            oy = New Vector

            oy.Coord(0) = 0

            oy.Coord(1) = 0

            oy.Coord(2) = 1

            ox = oy ^ oz

            oy = oz ^ ox

            ox = Norm_vector(ox)

            oy = Norm_vector(oy)

            oz = Norm_vector(oz)

        End If

        ' Model initializatioan and clearing

        For i As Integer = 0 To UBound(FEM.Shining_face)

            'ReDim FEM.Shining_face(i).Absorption_act(Max_items_default)

            'FEM.Shining_face(i).i_Absorption_act = -1

            FEM.Shining_face(i).Interaction = New Face_photon_interaction

        Next i


        ' Photon emission cycle
        Dim Photon As Photon = New Photon

        ' Face, hitted by photon
        Dim Face_number As Integer

        Dim i_Face_with_absorbtion_number_0 As Integer = -1

        Dim i_Face_with_absorbtion_number_1 As Integer = -1

        Dim i_Photon As Integer = 0

        Dim view_from As Vector

        Dim view_to As Vector

        Dim Photon_emission_center As Vector = New Vector

        ' Array of photons to emitt

        Dim Photons(N_photon - 1) As Photon

        Photons = Photon.Generate_N_random_photons(Face, N_photon, New_opening_angle, CDF, Argument, Emitter_medium, 0)

        For i As Integer = 0 To N_photon - 1

            ' Emitting random photon

            i_Photon += 1

            Photon = Photons(i)

            'Photon = New Photon

            'Photon = Photon.Emit_random_photon(Source_location, New_opening_angle, CDF, Argument)

            'Photon.Number = i_Photon

            'Photon.Medium = Emitter_medium

            ' Transforming photon coordinates from emitter coordinate system to global coordinate system

            Photon.Direction = Photon.Direction.Coord(0) * ox + Photon.Direction.Coord(1) * oy + Photon.Direction.Coord(2) * oz

            ' Propagation cycle

            ' First photon emitted from corrected emitter position

            If Source_distance > Limit_distance Then

                Photon.Center = Photon_emission_center + (Source_distance - FEM.Maximum_length) * Photon.Direction

            End If

            ' if we have to emit spherically  
            If Direction_should_be_reverted Then

                ' we have to reverse emission direction after each proton to fill the sphere 
                ox = -1 * ox

                oy = -1 * oy

                oz = -1 * oz

            End If

            Do

                view_from = Photon.Center + FEM.Minimum_length * 0.001 * Photon.Direction

                view_to = Photon.Center + Photon.Direction

                'Dim writer_1 As System.IO.StreamWriter = New System.IO.StreamWriter("C:\tmp\View_from_view_to.txt", True)

                'writer_1.WriteLine(Str(view_from.Coord(0)) & " " & Str(view_from.Coord(1)) & " " & Str(view_from.Coord(2)) & " " & Str(view_to.Coord(0)) & " " & Str(view_to.Coord(1)) & " " & Str(view_to.Coord(2)))

                ''writer_1.WriteLine(Str(Photon.Center.Coord(0)) & " " & Str(Photon.Center.Coord(1)) & " " & Str(Photon.Center.Coord(2)))

                ''writer_1.WriteLine(Str(Photon.Direction.Coord(0)) & " " & Str(Photon.Direction.Coord(1)) & " " & Str(Photon.Direction.Coord(2)))


                'writer_1.Close()


                'Face_number = View_to_vector(vf, Photon.Center + FEM.Minimum_length * 0.001 * Photon.Direction, Photon.Center + Photon.Direction)

                Face_number = vf.View_to_vector(view_from.Coord(0), view_from.Coord(1), view_from.Coord(2), view_to.Coord(0), view_to.Coord(1), view_to.Coord(2))

                If Face_number <> 0 Then


                    'Dim writer As System.IO.StreamWriter = New System.IO.StreamWriter("C:\tmp\Faces_numbers.txt", True)

                    'writer.WriteLine(Face_number)

                    'writer.Close()

                    Fill_Face_argument(FEM.Shining_face(Face_number - 1), Photon)

                    Photon = FEM.Shining_face(Face_number - 1).Face_photon_interaction(Photon)

                    ' If photon absorbed, exit cycle
                    If Photon Is Nothing Then

                        Exit Do

                    End If

                Else

                    ' If photon is passed away from model, exit cycle

                    Exit Do

                End If

            Loop

        Next i


        For i As Integer = 0 To UBound(FEM.Shining_face)

            With FEM.Shining_face(i)

                Steady_state_vector_1.Coord(i) = Correction_coef * (.Interaction.N_R_dif + .Interaction.N_R_mir + .Interaction.N_D + .Interaction.N_A) / i_Photon * .FTOP.A

            End With

        Next i

        Return Steady_state_vector_1.Coord

    End Function


    '''' <summary>
    '''' This function illiminate model with photons, emutted from point Source_location,
    '''' in the line of Direction with Opening_angle.
    '''' The function returns array of energies absorbed in model's elements.    
    '''' </summary>
    '''' <param name="Source_location">Location of radiotion source</param>
    '''' <param name="Direction">Direction of radiation propagation</param>
    '''' <param name="Opening_angle">Opening angle in radiation propagation cone
    '''' If Opening_angle = PI, photons emitted in all of directions, spherically.
    '''' If Opening_angle = PI/2, photons emitted in one half-space.
    '''' If Opening_angle more then 0 and less than  PI/2, photons emitted in cone.</param>
    '''' <param name="CDF">Radiation source cumulative distribution of photon frequency</param>
    '''' <param name="Argument">Radiation source paratemers, such as a temperature, time and so on.</param>
    '''' <param name="Emitter_medium">Medium, surrounded radiation source.</param>
    '''' <param name="vf">Visualisation form, used for photon-face interaction detection</param>
    '''' <param name="Step_epsilon">Maximum available error in view-factor-vector calculation process</param>
    '''' <param name="Total_outcoming_heat_power">Total heat power outcoming from source</param>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public Function Illuminate_model_and_calc_radiation_Q(ByRef Source_location As Vector, ByRef Direction As Vector, ByRef Opening_angle As Double, ByRef CDF As DEPEND, ByRef Argument As Vector, ByRef Emitter_medium As MEDIUM1, ByRef vf As Visualisation.Visualisation, ByRef Step_epsilon As Double, ByRef Total_outcoming_heat_power As Double) As Double(,)

    '    Illuminate_model_and_calc_radiation_Q = Nothing

    '    ' Number of protons, emitted before steady-state attainment definition
    '    Dim Steady_state_photon_number As Integer = 1

    '    ' Number of photons, emitted before steady-state attainment checks
    '    Dim Photons_between_checks As Integer = 10

    '    Dim Next_check As Integer = Steady_state_photon_number

    '    ' Vectors with previous and current absorption coefficients
    '    Dim Steady_state_vector_0, Steady_state_vector_1 As Vector

    '    ' Length of vectors with previous and current absorption coefficients
    '    Dim L_Steady_state_vector_0, L_Steady_state_vector_1 As Double

    '    L_Steady_state_vector_0 = 0

    '    L_Steady_state_vector_1 = 0

    '    Steady_state_vector_0 = New Vector

    '    ReDim Steady_state_vector_0.Coord(UBound(FEM.Shining_face))

    '    Steady_state_vector_1 = New Vector

    '    ReDim Steady_state_vector_1.Coord(UBound(FEM.Shining_face))

    '    ' Previous and current vector difference. Vector compopent, whose difference is less than set maximum
    '    ' are excluded from error difference calculation
    '    Dim Vector_error As Double = Double.MaxValue


    '    ' Feed, illimination angle calculation

    '    Dim Limit_distance As Double = 2 * FEM.Maximum_length

    '    ' If radiation source is located farther then Limit_distance, then it can be shifted to it distance.

    '    ' If radiation source is located farther then Limit_distance, then its opening angle should be shifted to Math.Atan(FEM.Maximum_length/Limit_distance).

    '    ' If radiation source is shifted to new position, it is needed to correct absorption in model elements.
    '    ' It is caused by part of photons guarantily passed away from model and can't be absorbed.

    '    Dim to_Center As Vector = FEM.Center - Source_location

    '    Dim Source_distance As Double = L_vector(to_Center)

    '    'Dim New_source_location As New Vector

    '    Dim New_direction As New Vector

    '    Dim New_opening_angle As Double

    '    ' This coefficient is used for absorption correction due to illimination angle changing
    '    Dim Correction_coef As Double

    '    Dim Direction_should_be_reverted As Boolean = False

    '    If Source_distance > Limit_distance Then


    '        New_opening_angle = Math.Atan(FEM.Maximum_length / 2 / Source_distance)

    '        Correction_coef = Math.Sin(0.5 * New_opening_angle) ^ 2

    '        New_direction = to_Center

    '        If Math.Abs(New_direction.Coord(0)) < Epsilon Then New_direction.Coord(0) = 0

    '        If Math.Abs(New_direction.Coord(1)) < Epsilon Then New_direction.Coord(1) = 0

    '        If Math.Abs(New_direction.Coord(2)) < Epsilon Then New_direction.Coord(2) = 0


    '    Else

    '        'New_source_location = Source_location

    '        New_opening_angle = Opening_angle

    '        If New_opening_angle = Math.PI Then

    '            New_opening_angle = 0.5 * Math.PI

    '            Direction_should_be_reverted = True

    '        End If

    '        New_direction = Direction

    '        Correction_coef = 1

    '    End If


    '    ' Coordinate system building

    '    Dim ox, oy, oz As Vector

    '    oz = Norm_vector(New_direction)

    '    If oz.Coord(0) = 0 And oz.Coord(1) = 0 And Math.Abs(oz.Coord(2)) = 1 Then

    '        ox = New Vector

    '        ox.Coord(0) = 1

    '        ox.Coord(1) = 0

    '        ox.Coord(2) = 0


    '        oy = New Vector

    '        oy.Coord(0) = 0

    '        oy.Coord(1) = 1

    '        oy.Coord(2) = 0

    '    Else

    '        oy = New Vector

    '        oy.Coord(0) = 0

    '        oy.Coord(1) = 0

    '        oy.Coord(2) = 1

    '        ox = oy ^ oz

    '        oy = oz ^ ox

    '    End If

    '    ' Model initializatioan and clearing

    '    For i As Integer = 0 To UBound(FEM.Shining_face)

    '        'ReDim FEM.Shining_face(i).Absorption_act(Max_items_default)

    '        'FEM.Shining_face(i).i_Absorption_act = -1

    '        FEM.Shining_face(i).Interaction = New Face_photon_interaction

    '    Next i


    '    ' Photon emission cycle
    '    Dim Photon As Photon

    '    ' Face, hitted by photon
    '    Dim Face_number As Integer

    '    Const N_Photon_max = 10000

    '    Dim i_Face_with_absorbtion_number_0 As Integer = -1

    '    Dim i_Face_with_absorbtion_number_1 As Integer = -1

    '    Dim Face_with_absorbtion_number_0(N_Photon_max) As Integer

    '    Dim Face_with_absorbtion_number_1(N_Photon_max) As Integer

    '    Dim i_Photon As Integer = 0

    '    Do

    '        ' Emitting random photon

    '        i_Photon += 1

    '        Photon = New Photon

    '        Photon = Photon.Emit_random_photon(Source_location, New_opening_angle, CDF, Argument)

    '        Photon.Number = i_Photon

    '        Photon.Medium = Emitter_medium

    '        ' Transforming photon coordinates from emitter coordinate system to global coordinate system

    '        Photon.Direction = Photon.Direction.Coord(0) * ox + Photon.Direction.Coord(1) * oy + Photon.Direction.Coord(2) * oz

    '        ' Propagation cycle

    '        ' First photon emitted from corrected emitter position

    '        If Source_distance > Limit_distance Then

    '            'Photon.Center = Source_location + (Source_distance - 0.9 * Limit_distance) * Photon.Direction

    '            Photon.Center = Source_location + (Source_distance - FEM.Maximum_length) * Photon.Direction

    '        End If

    '        'if we have to emit spherically  
    '        If Direction_should_be_reverted Then

    '            ' we have to reverse emission direction after each proton to fill the sphere 
    '            ox = -1 * ox

    '            oy = -1 * oy

    '            oz = -1 * oz

    '        End If

    '        Do

    '            Face_number = View_to_vector(vf, Photon.Center + FEM.Minimum_length * 0.001 * Photon.Direction, Photon.Center + Photon.Direction)

    '            If Face_number <> 0 Then

    '                Fill_Face_argument(FEM.Shining_face(Face_number - 1), Photon)

    '                Photon = FEM.Shining_face(Face_number - 1).Face_photon_interaction(Photon)

    '                ' If photon absorbed, exit cycle
    '                If Photon Is Nothing Then

    '                    i_Face_with_absorbtion_number_0 += 1

    '                    Face_with_absorbtion_number_0(i_Face_with_absorbtion_number_0) = FEM.Shining_face(Face_number - 1).Number

    '                    Exit Do

    '                End If

    '            Else

    '                ' If photon is passed away from model, exit cycle

    '                Exit Do

    '            End If

    '        Loop


    '        '' Absorbtion coeffients error evaluation

    '        '' If we emitted adequate number of photons, we can check for steady-state mode 
    '        'If i_Photon >= Steady_state_photon_number Then

    '        '    ' If it is time for steady-state mode check
    '        '    If i_Photon = Next_check Then

    '        '        Next_check += Photons_between_checks

    '        '        ''If i_Face_with_absorbtion_number_0 = -1 Then Exit Do

    '        '        ''Forming array of faces absorbed the photons numbers

    '        '        'Dim Is_in_array As Boolean

    '        '        ''i_Face_with_absorbtion_number_1 = -1

    '        '        'For i As Integer = 0 To i_Face_with_absorbtion_number_0

    '        '        '    Is_in_array = False

    '        '        '    For j As Integer = 0 To i_Face_with_absorbtion_number_1

    '        '        '        If Face_with_absorbtion_number_0(i) = Face_with_absorbtion_number_1(j) Then

    '        '        '            Is_in_array = True

    '        '        '            Exit For

    '        '        '        End If

    '        '        '    Next j

    '        '        '    If Not (Is_in_array) Then

    '        '        '        i_Face_with_absorbtion_number_1 += 1

    '        '        '        Face_with_absorbtion_number_1(i_Face_with_absorbtion_number_1) = Face_with_absorbtion_number_0(i)

    '        '        '    End If

    '        '        'Next i

    '        '        'i_Face_with_absorbtion_number_0 = -1


    '        '        ' if it is first steady state vector, we can't calculate difference.
    '        '        If i_Photon = Steady_state_photon_number Then

    '        '            ' Current vector calculation
    '        '            For i As Integer = 0 To UBound(Steady_state_vector_1.Coord)

    '        '                Steady_state_vector_1.Coord(i) = Correction_coef * (FEM.Shining_face(i).Interaction.N_A) / i_Photon

    '        '                L_Steady_state_vector_1 += Steady_state_vector_1.Coord(i) ^ 2

    '        '            Next i

    '        '            L_Steady_state_vector_1 = Math.Sqrt(L_Steady_state_vector_1)

    '        '        Else


    '        '            ' Saving previous vector

    '        '            For i As Integer = 0 To UBound(Steady_state_vector_1.Coord)

    '        '                Steady_state_vector_0.Coord(i) = Steady_state_vector_1.Coord(i)

    '        '            Next i

    '        '            L_Steady_state_vector_0 = L_Steady_state_vector_1

    '        '            ' Current vector calculation and vector difference evaluation

    '        '            Dim Component_error As Double = Double.MaxValue

    '        '            For i As Integer = 0 To UBound(Steady_state_vector_1.Coord)

    '        '                Steady_state_vector_1.Coord(i) = Correction_coef * (FEM.Shining_face(i).Interaction.N_A) / i_Photon

    '        '                Component_error = Math.Abs(Steady_state_vector_1.Coord(i) - Steady_state_vector_0.Coord(i))

    '        '                ' If vector component error is more than set error this component is included into error calculation
    '        '                If Component_error > Step_epsilon Then

    '        '                    L_Steady_state_vector_1 += Steady_state_vector_1.Coord(i) ^ 2

    '        '                End If

    '        '            Next i

    '        '            L_Steady_state_vector_1 = Math.Sqrt(L_Steady_state_vector_1)

    '        '            ' Error evaluation

    '        '            Vector_error = Math.Abs(L_Steady_state_vector_1 - L_Steady_state_vector_0)

    '        '            ' If steady-state mode is archived 
    '        '            If Vector_error <= Step_epsilon Then

    '        '                ' then calculate incoming heat power to model's elements

    '        '                Dim Return_array(1, UBound(Steady_state_vector_1.Coord)) As Double

    '        '                Dim i_eluminated_face As Integer = -1

    '        '                For i As Integer = 0 To UBound(Steady_state_vector_1.Coord)

    '        '                    If Steady_state_vector_1.Coord(i) > 0 Then

    '        '                        i_eluminated_face += 1

    '        '                        Return_array(0, i_eluminated_face) = i

    '        '                        Return_array(1, i_eluminated_face) = Steady_state_vector_1.Coord(i) * Total_outcoming_heat_power

    '        '                    End If

    '        '                Next i

    '        '                If i_eluminated_face = -1 Then

    '        '                    Return Nothing

    '        '                Else

    '        '                    ReDim Preserve Return_array(1, i_eluminated_face)

    '        '                    Return Return_array

    '        '                End If


    '        '            End If

    '        '        End If

    '        '    End If

    '        'End If

    '        ' If we emitted adequate number of photons, we can check for steady-state mode 
    '        If i_Photon >= Steady_state_photon_number Then

    '            ' If it is time for steady-state mode check
    '            If i_Photon = Next_check Then

    '                ' Begin steady-state mode check

    '                Next_check += Photons_between_checks

    '                'If i_Face_with_absorbtion_number_0 = -1 Then Exit Do

    '                'Forming array of faces absorbed the photons numbers

    '                Dim Is_in_array As Boolean

    '                'i_Face_with_absorbtion_number_1 = -1

    '                For i As Integer = 0 To i_Face_with_absorbtion_number_0

    '                    Is_in_array = False

    '                    For j As Integer = 0 To i_Face_with_absorbtion_number_1

    '                        If Face_with_absorbtion_number_0(i) = Face_with_absorbtion_number_1(j) Then

    '                            Is_in_array = True

    '                            Exit For

    '                        End If

    '                    Next j

    '                    If Not (Is_in_array) Then

    '                        i_Face_with_absorbtion_number_1 += 1

    '                        Face_with_absorbtion_number_1(i_Face_with_absorbtion_number_1) = Face_with_absorbtion_number_0(i)

    '                    End If

    '                Next i

    '                i_Face_with_absorbtion_number_0 = -1

    '                ' if it is first steady state vector, we can't calculate difference.
    '                If i_Photon = Steady_state_photon_number Then

    '                    ReDim Steady_state_vector_1.Coord(i_Face_with_absorbtion_number_1)

    '                    ' Current vector calculation
    '                    For i As Integer = 0 To i_Face_with_absorbtion_number_1

    '                        Steady_state_vector_1.Coord(i) = Correction_coef * (FEM.Shining_face(Face_with_absorbtion_number_1(i)).Interaction.N_A) / i_Photon

    '                        L_Steady_state_vector_1 += Steady_state_vector_1.Coord(i) ^ 2

    '                    Next i

    '                    L_Steady_state_vector_1 = Math.Sqrt(L_Steady_state_vector_1)

    '                Else


    '                    ' Saving previous vector

    '                    ReDim Steady_state_vector_0.Coord(i_Face_with_absorbtion_number_1)

    '                    For i As Integer = 0 To UBound(Steady_state_vector_1.Coord)

    '                        Steady_state_vector_0.Coord(i) = Steady_state_vector_1.Coord(i)

    '                    Next i

    '                    L_Steady_state_vector_0 = L_Steady_state_vector_1

    '                    ' Current vector calculation and vector difference evaluation

    '                    Dim Component_error As Double = Double.MaxValue

    '                    ' Current vector calculation

    '                    ReDim Steady_state_vector_1.Coord(i_Face_with_absorbtion_number_1)

    '                    For i As Integer = 0 To i_Face_with_absorbtion_number_1

    '                        Steady_state_vector_1.Coord(i) = Correction_coef * (FEM.Shining_face(Face_with_absorbtion_number_1(i)).Interaction.N_A) / i_Photon

    '                        Component_error = Math.Abs(Steady_state_vector_1.Coord(i) - Steady_state_vector_0.Coord(i))

    '                        ' If vector component error is more than set error this component is included into error calculation
    '                        If Component_error > Step_epsilon Then

    '                            L_Steady_state_vector_1 += Steady_state_vector_1.Coord(i) ^ 2

    '                        End If

    '                    Next i

    '                    L_Steady_state_vector_1 = Math.Sqrt(L_Steady_state_vector_1)

    '                    ' Error evaluation

    '                    Vector_error = Math.Abs(L_Steady_state_vector_1 - L_Steady_state_vector_0)

    '                    ' If steady-state mode is archived 
    '                    If Vector_error <= Step_epsilon Then

    '                        ' then calculate incoming heat power to model's elements

    '                        Dim Return_array(1, i_Face_with_absorbtion_number_1) As Double

    '                        For i As Integer = 0 To i_Face_with_absorbtion_number_1

    '                            Return_array(0, i) = Face_with_absorbtion_number_1(i)

    '                            Return_array(1, i) = Steady_state_vector_1.Coord(i) * Total_outcoming_heat_power

    '                        Next i

    '                        Return Return_array

    '                    End If

    '                End If

    '            End If

    '        End If


    '        Application.DoEvents()

    '    Loop

    'End Function

    ''' <summary>
    ''' This function illiminate model with photons, emutted from point Source_location,
    ''' in the line of Direction with Opening_angle.
    ''' The function returns array of energies absorbed in model's elements.    
    ''' </summary>
    ''' <param name="Source_location">Location of radiotion source</param>
    ''' <param name="Direction">Direction of radiation propagation</param>
    ''' <param name="Opening_angle">Opening angle in radiation propagation cone
    ''' If Opening_angle = PI, photons emitted in all of directions, spherically.
    ''' If Opening_angle = PI/2, photons emitted in one half-space.
    ''' If Opening_angle more then 0 and less than  PI/2, photons emitted in cone.</param>
    ''' <param name="CDF">Radiation source cumulative distribution of photon frequency</param>
    ''' <param name="Argument">Radiation source paratemers, such as a temperature, time and so on.</param>
    ''' <param name="Emitter_medium">Medium, surrounded radiation source.</param>
    ''' <param name="vf">Visualisation form, used for photon-face interaction detection</param>
    ''' <param name="N_photon">Number of photons emitted from source</param>
    ''' <param name="Total_outcoming_heat_power">Total heat power outcoming from source</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Illuminate_model_and_calc_radiation_Q(ByRef Source_location As Vector, ByRef Direction As Vector, ByRef Opening_angle As Double, ByRef CDF As DEPEND, ByRef Argument As Vector, ByRef Emitter_medium As MEDIUM1, ByRef vf As Visualisation.Visualisation, ByRef N_photon As Integer, ByRef Total_outcoming_heat_power As Double) As Double(,)

        Dim t_0, t_1 As Long

        t_0 = timeGetTime

        Illuminate_model_and_calc_radiation_Q = Nothing

        ' Number of protons, emitted before steady-state attainment definition
        Dim Steady_state_photon_number As Integer = 1

        ' Number of photons, emitted before steady-state attainment checks
        'Dim Photons_between_checks As Integer = 10

        Dim Next_check As Integer = Steady_state_photon_number

        ' Vectors with previous and current absorption coefficients
        Dim Steady_state_vector_0, Steady_state_vector_1 As Vector

        ' Length of vectors with previous and current absorption coefficients
        Dim L_Steady_state_vector_0, L_Steady_state_vector_1 As Double

        L_Steady_state_vector_0 = 0

        L_Steady_state_vector_1 = 0

        Steady_state_vector_0 = New Vector

        ReDim Steady_state_vector_0.Coord(UBound(FEM.Shining_face))

        Steady_state_vector_1 = New Vector

        ReDim Steady_state_vector_1.Coord(UBound(FEM.Shining_face))

        ' Previous and current vector difference. Vector compopent, whose difference is less than set maximum
        ' are excluded from error difference calculation
        'Dim Vector_error As Double = Double.MaxValue


        ' Feed, illimination angle calculation

        Dim Limit_distance As Double = vf.Max_length

        ' If radiation source is located farther then Limit_distance, then it can be shifted to it distance.

        ' If radiation source is located farther then Limit_distance, then its opening angle should be shifted to Math.Atan(FEM.Maximum_length/Limit_distance).

        ' If radiation source is shifted to new position, it is needed to correct absorption in model elements.
        ' It is caused by part of photons guarantily passed away from model and can't be absorbed.

        Dim to_Center As Vector = FEM.Center - Source_location

        Dim Source_distance As Double = L_vector(to_Center)

        'Dim New_source_location As New Vector

        Dim New_direction As New Vector

        Dim New_opening_angle As Double

        Dim Direction_should_be_reverted As Boolean = False

        ' This coefficient is used for absorption correction due to illimination angle changing
        Dim Correction_coef As Double

        If Source_distance > Limit_distance Then


            New_opening_angle = Math.Atan(FEM.Maximum_length / 2 / Source_distance)

            Correction_coef = Math.Sin(0.5 * New_opening_angle) ^ 2

            New_direction = to_Center

            If Math.Abs(New_direction.Coord(0)) < Epsilon Then New_direction.Coord(0) = 0

            If Math.Abs(New_direction.Coord(1)) < Epsilon Then New_direction.Coord(1) = 0

            If Math.Abs(New_direction.Coord(2)) < Epsilon Then New_direction.Coord(2) = 0


        Else

            'New_source_location = Source_location

            New_opening_angle = Opening_angle

            If New_opening_angle = Math.PI Then

                New_opening_angle = 0.5 * Math.PI

                Direction_should_be_reverted = True

            End If

            New_direction = Direction

            Correction_coef = 1

        End If


        'New_opening_angle = Opening_angle

        'New_direction = Direction

        'Correction_coef = 1


        ' Coordinate system building

        Dim ox, oy, oz As Vector

        oz = Norm_vector(New_direction)

        If oz.Coord(0) = 0 And oz.Coord(1) = 0 And Math.Abs(oz.Coord(2)) = 1 Then

            ox = New Vector

            ox.Coord(0) = 1

            ox.Coord(1) = 0

            ox.Coord(2) = 0


            oy = New Vector

            oy.Coord(0) = 0

            oy.Coord(1) = 1

            oy.Coord(2) = 0

        Else

            oy = New Vector

            oy.Coord(0) = 0

            oy.Coord(1) = 0

            oy.Coord(2) = 1

            ox = oy ^ oz

            oy = oz ^ ox

            ox = Norm_vector(ox)

            oy = Norm_vector(oy)

            oz = Norm_vector(oz)

        End If

        ' Model initializatioan and clearing

        For i As Integer = 0 To UBound(FEM.Shining_face)

            'ReDim FEM.Shining_face(i).Absorption_act(Max_items_default)

            'FEM.Shining_face(i).i_Absorption_act = -1

            FEM.Shining_face(i).Interaction = New Face_photon_interaction

        Next i


        ' Photon emission cycle
        Dim Photon As Photon = New Photon

        ' Face, hitted by photon
        Dim Face_number As Integer

        Const N_Photon_max = 1000000

        Dim i_Face_with_absorbtion_number_0 As Integer = -1

        Dim i_Face_with_absorbtion_number_1 As Integer = -1

        Dim Face_with_absorbtion_number_0(N_Photon_max) As Integer

        Dim Face_with_absorbtion_number_1(N_Photon_max) As Integer

        Dim i_Photon As Integer = 0

        Dim view_from As Vector

        Dim view_to As Vector

        ' Array of photons to emitt

        Dim Photons(N_photon - 1) As Photon

        Photons = Photon.Generate_N_random_photons(Source_location, N_photon, New_opening_angle, CDF, Argument, Emitter_medium, Total_outcoming_heat_power)

        For i As Integer = 0 To N_photon - 1

            ' Emitting random photon

            i_Photon += 1

            Photon = Photons(i)

            'Photon = New Photon

            'Photon = Photon.Emit_random_photon(Source_location, New_opening_angle, CDF, Argument)

            'Photon.Number = i_Photon

            'Photon.Medium = Emitter_medium

            ' Transforming photon coordinates from emitter coordinate system to global coordinate system

            Photon.Direction = Photon.Direction.Coord(0) * ox + Photon.Direction.Coord(1) * oy + Photon.Direction.Coord(2) * oz



            ' Propagation cycle

            ' First photon emitted from corrected emitter position

            If Source_distance > Limit_distance Then

                'Photon.Center = Source_location + (Source_distance - 0.9 * Limit_distance) * Photon.Direction

                Photon.Center = Source_location + (Source_distance - FEM.Maximum_length) * Photon.Direction

            End If

            ' if we have to emit spherically  
            If Direction_should_be_reverted Then

                ' we have to reverse emission direction after each proton to fill the sphere 
                ox = -1 * ox

                oy = -1 * oy

                oz = -1 * oz

            End If

            Do

                view_from = Photon.Center + FEM.Minimum_length * 0.001 * Photon.Direction

                view_to = Photon.Center + Photon.Direction

                'Dim writer_1 As System.IO.StreamWriter = New System.IO.StreamWriter("C:\tmp\View_from_view_to.txt", True)

                'writer_1.WriteLine(Str(view_from.Coord(0)) & " " & Str(view_from.Coord(1)) & " " & Str(view_from.Coord(2)) & " " & Str(view_to.Coord(0)) & " " & Str(view_to.Coord(1)) & " " & Str(view_to.Coord(2)))

                ''writer_1.WriteLine(Str(Photon.Center.Coord(0)) & " " & Str(Photon.Center.Coord(1)) & " " & Str(Photon.Center.Coord(2)))

                ''writer_1.WriteLine(Str(Photon.Direction.Coord(0)) & " " & Str(Photon.Direction.Coord(1)) & " " & Str(Photon.Direction.Coord(2)))


                'writer_1.Close()


                'Face_number = View_to_vector(vf, Photon.Center + FEM.Minimum_length * 0.001 * Photon.Direction, Photon.Center + Photon.Direction)

                Face_number = vf.View_to_vector(view_from.Coord(0), view_from.Coord(1), view_from.Coord(2), view_to.Coord(0), view_to.Coord(1), view_to.Coord(2))

                If Face_number <> 0 Then


                    'Dim writer As System.IO.StreamWriter = New System.IO.StreamWriter("C:\tmp\Faces_numbers.txt", True)

                    'writer.WriteLine(Face_number)

                    'writer.Close()

                    Fill_Face_argument(FEM.Shining_face(Face_number - 1), Photon)

                    Photon = FEM.Shining_face(Face_number - 1).Face_photon_interaction(Photon)

                    i_Face_with_absorbtion_number_0 += 1

                    Face_with_absorbtion_number_0(i_Face_with_absorbtion_number_0) = FEM.Shining_face(Face_number - 1).Number

                    ' If photon absorbed, exit cycle
                    If Photon Is Nothing Then

                        'i_Face_with_absorbtion_number_0 += 1

                        'Face_with_absorbtion_number_0(i_Face_with_absorbtion_number_0) = FEM.Shining_face(Face_number - 1).Number

                        Exit Do

                    End If

                Else

                    ' If photon is passed away from model, exit cycle

                    Exit Do

                End If

            Loop

            'Application.DoEvents()

        Next i

        'Forming array of faces absorbed the photons numbers

        Dim Is_in_array As Boolean

        'i_Face_with_absorbtion_number_1 = -1

        For i As Integer = 0 To i_Face_with_absorbtion_number_0

            Is_in_array = False

            For j As Integer = 0 To i_Face_with_absorbtion_number_1

                If Face_with_absorbtion_number_0(i) = Face_with_absorbtion_number_1(j) Then

                    Is_in_array = True

                    Exit For

                End If

            Next j

            If Not (Is_in_array) Then

                i_Face_with_absorbtion_number_1 += 1

                Face_with_absorbtion_number_1(i_Face_with_absorbtion_number_1) = Face_with_absorbtion_number_0(i)


            End If

        Next i


        ReDim Steady_state_vector_1.Coord(i_Face_with_absorbtion_number_1)

        ' Current vector calculation
        For i As Integer = 0 To i_Face_with_absorbtion_number_1

            'Steady_state_vector_1.Coord(i) = Correction_coef * (FEM.Shining_face(Face_with_absorbtion_number_1(i)).Interaction.N_A) / i_Photon

            With FEM.Shining_face(Face_with_absorbtion_number_1(i))

                Steady_state_vector_1.Coord(i) = Correction_coef * (.Interaction.N_R_dif + .Interaction.N_R_mir + .Interaction.N_D + .Interaction.N_A) / i_Photon * .FTOP.A

            End With

            '= Correction_coef * FEM.Shining_face(Face_with_absorbtion_number_1(i)).Interaction.N_R_dif / i_Photon * FEM.Shining_face(Face_with_absorbtion_number_1(i)).FTOP.A
        Next i


        ' then calculate incoming heat power to model's elements

        Dim Return_array(1, i_Face_with_absorbtion_number_1) As Double

        For i As Integer = 0 To i_Face_with_absorbtion_number_1

            'Dim fs As New System.IO.StreamWriter(FEM.File_path & FEM.File_name & "_angle_coefficients.txt", IO.FileMode.Append)

            'fs.WriteLine(Steady_state_vector_1.Coord(0))

            'fs.Close()

            Return_array(0, i) = Face_with_absorbtion_number_1(i)

            Return_array(1, i) = Steady_state_vector_1.Coord(i) * Total_outcoming_heat_power

        Next i

        t_1 = timeGetTime

        'Dim Txt_writer As System.IO.StreamWriter = New System.IO.StreamWriter(FEM.File_path & "Photon_time.txt", True)

        'Txt_writer.WriteLine((t_1 - t_0).ToString)

        'Txt_writer.Close()

        Return Return_array


    End Function

    ''' <summary>
    ''' This function calculates illimination of model from given face with photons.
    ''' The emission center in each act of emission is located randomly on face surface.
    ''' The photon emission is directed by Direction within Opening_angle.
    ''' The function returns array of energies absorbed in model's elements.    
    ''' </summary>
    ''' <param name="Face">Face to emit photons</param>
    ''' <param name="Direction">Direction of radiation propagation</param>
    ''' <param name="Opening_angle">Opening angle in radiation propagation cone
    ''' If Opening_angle = PI, photons emitted in all of directions, spherically.
    ''' If Opening_angle = PI/2, photons emitted in one half-space.
    ''' If Opening_angle more then 0 and less than  PI/2, photons emitted in cone.</param>
    ''' <param name="CDF">Radiation source cumulative distribution of photon frequency</param>
    ''' <param name="Argument">Radiation source paratemers, such as a temperature, time and so on.</param>
    ''' <param name="Emitter_medium">Medium, surrounded radiation source.</param>
    ''' <param name="vf">Visualisation form, used for photon-face interaction detection</param>
    ''' <param name="N_photon">Number of photons emitted from source</param>
    ''' <param name="Total_outcoming_heat_power">Total heat power outcoming from source</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Illuminate_model_from_face_and_calc_radiation_Q(ByRef Face As Face, ByRef Direction As Vector, ByRef Opening_angle As Double, ByRef CDF As DEPEND, ByRef Argument As Vector, ByRef Emitter_medium As MEDIUM1, ByRef vf As Visualisation.Visualisation, ByRef N_photon As Integer, ByRef Total_outcoming_heat_power As Double) As Double(,)

        ' if this face have emissivity equal to zero, then we should exit function

        Dim t_0, t_1 As Long

        t_0 = timeGetTime

        Dim t_1_0, t_1_1, t_1_2 As Long

        Illuminate_model_from_face_and_calc_radiation_Q = Nothing

        ' Number of protons, emitted before steady-state attainment definition
        Dim Steady_state_photon_number As Integer = 1

        ' Number of photons, emitted before steady-state attainment checks
        'Dim Photons_between_checks As Integer = 10

        Dim Next_check As Integer = Steady_state_photon_number

        ' Vectors with previous and current absorption coefficients
        Dim Steady_state_vector_0, Steady_state_vector_1 As Vector

        ' Length of vectors with previous and current absorption coefficients
        Dim L_Steady_state_vector_0, L_Steady_state_vector_1 As Double

        L_Steady_state_vector_0 = 0

        L_Steady_state_vector_1 = 0

        Steady_state_vector_0 = New Vector

        ReDim Steady_state_vector_0.Coord(UBound(FEM.Shining_face))

        Steady_state_vector_1 = New Vector

        ReDim Steady_state_vector_1.Coord(UBound(FEM.Shining_face))

        ' Previous and current vector difference. Vector compopent, whose difference is less than set maximum
        ' are excluded from error difference calculation
        Dim Vector_error As Double = Double.MaxValue


        ' Feed, illimination angle calculation

        Dim Limit_distance As Double = vf.Max_length

        ' The random source location calculation

        Dim Source_location As Vector = New Vector


        ' If radiation source is located farther then Limit_distance, then it can be shifted to it distance.

        ' If radiation source is located farther then Limit_distance, then its opening angle should be shifted to Math.Atan(FEM.Maximum_length/Limit_distance).

        ' If radiation source is shifted to new position, it is needed to correct absorption in model elements.
        ' It is caused by part of photons guarantily passed away from model and can't be absorbed.

        Dim to_Center As Vector = FEM.Center - Face.Center

        Dim Source_distance As Double = L_vector(to_Center)

        ''Dim New_source_location As New Vector

        Dim New_direction As New Vector

        Dim New_opening_angle As Double

        Dim Direction_should_be_reverted As Boolean = False

        '' This coefficient is used for absorption correction due to illimination angle changing
        Dim Correction_coef As Double

        'If Source_distance > Limit_distance Then


        '    New_opening_angle = Math.Atan(FEM.Maximum_length / 2 / Source_distance)

        '    Correction_coef = Math.Sin(0.5 * New_opening_angle) ^ 2

        '    New_direction = to_Center

        '    If Math.Abs(New_direction.Coord(0)) < Epsilon Then New_direction.Coord(0) = 0

        '    If Math.Abs(New_direction.Coord(1)) < Epsilon Then New_direction.Coord(1) = 0

        '    If Math.Abs(New_direction.Coord(2)) < Epsilon Then New_direction.Coord(2) = 0


        'Else

        '    'New_source_location = Source_location

        New_opening_angle = Opening_angle

        If New_opening_angle = Math.PI Then

            New_opening_angle = 0.5 * Math.PI

            Direction_should_be_reverted = True

        End If

        New_direction = Direction

        Correction_coef = 1

        'End If


        'New_opening_angle = Opening_angle

        'New_direction = Direction

        'Correction_coef = 1


        ' Coordinate system building

        Dim ox, oy, oz As Vector

        oz = Norm_vector(New_direction)

        If oz.Coord(0) = 0 And oz.Coord(1) = 0 And Math.Abs(oz.Coord(2)) = 1 Then

            ox = New Vector

            ox.Coord(0) = 1

            ox.Coord(1) = 0

            ox.Coord(2) = 0


            oy = New Vector

            oy.Coord(0) = 0

            oy.Coord(1) = 1

            oy.Coord(2) = 0

        Else

            oy = New Vector

            oy.Coord(0) = 0

            oy.Coord(1) = 0

            oy.Coord(2) = 1

            ox = oy ^ oz

            oy = oz ^ ox

            ox = Norm_vector(ox)

            oy = Norm_vector(oy)

            oz = Norm_vector(oz)

        End If

        ' Model initializatioan and clearing

        For i As Integer = 0 To UBound(FEM.Shining_face)

            'ReDim FEM.Shining_face(i).Absorption_act(Max_items_default)

            'FEM.Shining_face(i).i_Absorption_act = -1

            FEM.Shining_face(i).Interaction = New Face_photon_interaction

        Next i


        ' Photon emission cycle
        Dim Photon As Photon = New Photon

        ' Face, hitted by photon
        Dim Face_number As Integer

        Const N_Photon_max = 1000000

        Dim i_Face_with_absorbtion_number_0 As Integer = -1

        Dim i_Face_with_absorbtion_number_1 As Integer = -1

        Dim Face_with_absorbtion_number_0(N_Photon_max) As Integer

        Dim Face_with_absorbtion_number_1(N_Photon_max) As Integer

        Dim i_Photon As Integer = 0

        Dim view_from As Vector

        Dim view_to As Vector

        t_1_0 = timeGetTime

        Dim t_2_0, t_2_1, t_2_2 As Long

        ' Array of photons to emitt

        Dim Photons(N_photon - 1) As Photon

        Photons = Photon.Generate_N_random_photons(Face, N_photon, New_opening_angle, CDF, Argument, Emitter_medium, Total_outcoming_heat_power)

        For i As Integer = 0 To N_photon - 1

            t_2_0 = timeGetTime

            ' Emitting random photon

            i_Photon += 1

            Photon = Photons(i)

            'Photon = New Photon

            'Photon = Photon.Emit_random_photon(Source_location, New_opening_angle, CDF, Argument)

            'Photon.Number = i_Photon

            'Photon.Medium = Emitter_medium

            ' Transforming photon coordinates from emitter coordinate system to global coordinate system

            Photon.Direction = Photon.Direction.Coord(0) * ox + Photon.Direction.Coord(1) * oy + Photon.Direction.Coord(2) * oz


            ' Propagation cycle

            ' First photon emitted from corrected emitter position

            If Source_distance > Limit_distance Then

                'Photon.Center = Source_location + (Source_distance - 0.9 * Limit_distance) * Photon.Direction

                Photon.Center = Source_location + (L_vector(FEM.Center - Source_location) - FEM.Maximum_length) * Photon.Direction

            End If

            ' if we have to emit spherically  
            If Direction_should_be_reverted Then

                ' we have to reverse emission direction after each proton to fill the sphere 
                ox = -1 * ox

                oy = -1 * oy

                oz = -1 * oz

            End If


            t_2_1 = timeGetTime

            Do

                view_from = Photon.Center + FEM.Minimum_length * 0.001 * Photon.Direction

                view_to = Photon.Center + Photon.Direction

                'Face_number = View_to_vector(vf, Photon.Center + FEM.Minimum_length * 0.001 * Photon.Direction, Photon.Center + Photon.Direction)

                Face_number = vf.View_to_vector(view_from.Coord(0), view_from.Coord(1), view_from.Coord(2), view_to.Coord(0), view_to.Coord(1), view_to.Coord(2))

                If Face_number <> 0 Then

                    Fill_Face_argument(FEM.Shining_face(Face_number - 1), Photon)

                    Photon = FEM.Shining_face(Face_number - 1).Face_photon_interaction(Photon)

                    i_Face_with_absorbtion_number_0 += 1

                    Face_with_absorbtion_number_0(i_Face_with_absorbtion_number_0) = FEM.Shining_face(Face_number - 1).Number

                    ' If photon absorbed, exit cycle
                    If Photon Is Nothing Then

                        'i_Face_with_absorbtion_number_0 += 1

                        'Face_with_absorbtion_number_0(i_Face_with_absorbtion_number_0) = FEM.Shining_face(Face_number - 1).Number

                        Exit Do

                    End If

                Else

                    ' If photon is passed away from model, exit cycle

                    Exit Do

                End If

            Loop

            t_2_2 = timeGetTime

            'Application.DoEvents()

        Next i

        t_1_1 = timeGetTime

        'Forming array of faces absorbed the photons numbers

        Dim Is_in_array As Boolean

        'i_Face_with_absorbtion_number_1 = -1

        For i As Integer = 0 To i_Face_with_absorbtion_number_0

            Is_in_array = False

            For j As Integer = 0 To i_Face_with_absorbtion_number_1

                If Face_with_absorbtion_number_0(i) = Face_with_absorbtion_number_1(j) Then

                    Is_in_array = True

                    Exit For

                End If

            Next j

            If Not (Is_in_array) Then

                i_Face_with_absorbtion_number_1 += 1

                Face_with_absorbtion_number_1(i_Face_with_absorbtion_number_1) = Face_with_absorbtion_number_0(i)

            End If

        Next i

        t_1_2 = timeGetTime


        ReDim Steady_state_vector_1.Coord(i_Face_with_absorbtion_number_1)

        ' Current vector calculation
        For i As Integer = 0 To i_Face_with_absorbtion_number_1

            Steady_state_vector_1.Coord(i) = Correction_coef * (FEM.Shining_face(Face_with_absorbtion_number_1(i)).Interaction.N_A) / i_Photon

            'With FEM.Shining_face(Face_with_absorbtion_number_1(i))

            '    Steady_state_vector_1.Coord(i) = Correction_coef * (.Interaction.N_R_dif + .Interaction.N_R_mir + .Interaction.N_D + .Interaction.N_A) / i_Photon * .FTOP.Calc_intergal_A(.Argument)

            'End With

        Next i


        ' then calculate incoming heat power to model's elements

        Dim Return_array(1, i_Face_with_absorbtion_number_1) As Double

        For i As Integer = 0 To i_Face_with_absorbtion_number_1

            'Dim fs As New System.IO.StreamWriter(FEM.File_path & FEM.File_name & "_angle_coefficients.txt", IO.FileMode.Append)

            'fs.WriteLine(Steady_state_vector_1.Coord(0))

            'fs.Close()

            Return_array(0, i) = Face_with_absorbtion_number_1(i)

            Return_array(1, i) = Steady_state_vector_1.Coord(i) * Total_outcoming_heat_power

        Next i

        t_1 = timeGetTime

        'Dim Txt_writer As System.IO.StreamWriter = New System.IO.StreamWriter(FEM.File_path & "Photon_time.txt", True)

        'Txt_writer.WriteLine((t_1 - t_0).ToString)

        'Txt_writer.Close()

        Return Return_array


    End Function

    ''' <summary>
    ''' This sub calculates illimination of model from given face with photons.
    ''' The emission center in each act of emission is located randomly on face surface.
    ''' The photon emission is directed by Direction within Opening_angle.
    ''' The function returns array of energies absorbed in model's elements.    
    ''' </summary>
    ''' <param name="Source">Face or radiation source to emit photons</param>
    ''' <param name="Direction">Direction of radiation propagation</param>
    ''' <param name="Opening_angle">Opening angle in radiation propagation cone
    ''' If Opening_angle = PI, photons emitted in all of directions, spherically.
    ''' If Opening_angle = PI/2, photons emitted in one half-space.
    ''' If Opening_angle more then 0 and less than  PI/2, photons emitted in cone.</param>
    ''' <param name="CDF">Radiation source cumulative distribution of photon frequency</param>
    ''' <param name="Argument">Radiation source paratemers, such as a temperature, time and so on.</param>
    ''' <param name="Emitter_medium">Medium, surrounded radiation source.</param>
    ''' <param name="vf">Visualisation form, used for photon-face interaction detection</param>
    ''' <param name="N_photon">Number of photons emitted from source</param>
    ''' <param name="Total_outcoming_heat_power">Total heat power outcoming from source</param>
    ''' <remarks></remarks>
    Public Sub Illuminate_model_and_save_absorbed_power_in_faces(ByRef Source As TOP_item, ByRef Direction As Vector, ByRef Opening_angle As Double, ByRef CDF As DEPEND, ByRef Argument As Vector, ByRef Emitter_medium As MEDIUM1, ByRef vf As Visualisation.Visualisation, ByRef N_photon As Integer, ByRef Total_outcoming_heat_power As Double)

        Dim Face As Face = New Face

        Dim R_source As RSOURCE1 = New RSOURCE1

        Dim to_Center As Vector


        If UCase(Source.Type) = "FACE" Then

            Face = Source

            R_source = Nothing

            to_Center = FEM.Center - Face.Center

        Else

            If UCase(Source.Type) = "RSOURCE1" Then

                R_source = Source

                Face = Nothing

                Fill_RSOURCE_argument(R_source)

                R_source.POS = R_source.Calc_position

                to_Center = FEM.Center - R_source.POS

            Else

                Throw New Exception("You used have used " & Source.GetType.ToString & " as a radiation heat source. It is incorrect!")

            End If

        End If



        ' if this face have emissivity equal to zero, then we should exit function

        If Total_outcoming_heat_power = 0 Then Exit Sub

        ' Feed, illumination angle calculation

        Dim Limit_distance As Double = vf.Max_length

        ' The random source location calculation

        Dim Source_location As Vector = New Vector


        ' If radiation source is located farther then Limit_distance, then its photons can be shifted to Limit_distance distance.

        ' If radiation source is located farther then Limit_distance, then its opening angle should be shifted to Math.Atan(FEM.Maximum_length/Limit_distance).

        ' If radiation source is shifted to new position, it is needed to correct absorption in model elements.
        ' It is caused by part of photons guarantily passed away from model and can't be absorbed.

        
        Dim Source_distance As Double = L_vector(to_Center)

        Dim New_direction As New Vector

        Dim New_opening_angle As Double

        Dim Direction_should_be_reverted As Boolean = False

        '' This coefficient is used for absorption correction due to illimination angle changing
        Dim Correction_coef As Double

        If Source_distance > Limit_distance Then

            New_opening_angle = Math.Atan(FEM.Maximum_length / 2 / Source_distance)

            Correction_coef = Math.Sin(0.5 * New_opening_angle) ^ 2

            New_direction = to_Center

            If Math.Abs(New_direction.Coord(0)) < Epsilon Then New_direction.Coord(0) = 0

            If Math.Abs(New_direction.Coord(1)) < Epsilon Then New_direction.Coord(1) = 0

            If Math.Abs(New_direction.Coord(2)) < Epsilon Then New_direction.Coord(2) = 0


        Else


            New_opening_angle = Opening_angle

            If New_opening_angle = Math.PI Then

                New_opening_angle = 0.5 * Math.PI

                Direction_should_be_reverted = True

            End If

            New_direction = Direction

            Correction_coef = 1

        End If

        ' Coordinate system building

        Dim ox, oy, oz As Vector

        oz = Norm_vector(New_direction)

        If oz.Coord(0) = 0 And oz.Coord(1) = 0 And Math.Abs(oz.Coord(2)) = 1 Then

            ox = New Vector

            ox.Coord(0) = 1

            ox.Coord(1) = 0

            ox.Coord(2) = 0


            oy = New Vector

            oy.Coord(0) = 0

            oy.Coord(1) = 1

            oy.Coord(2) = 0

        Else

            oy = New Vector

            oy.Coord(0) = 0

            oy.Coord(1) = 0

            oy.Coord(2) = 1

            ox = oy ^ oz

            oy = oz ^ ox

            ox = Norm_vector(ox)

            oy = Norm_vector(oy)

            oz = Norm_vector(oz)

        End If


        ' Photon emission cycle
        Dim Photon As Photon = New Photon

        ' Face, hitted by photon
        Dim Face_number As Integer

        Dim i_Photon As Integer = 0

        Dim view_from As Vector

        Dim view_to As Vector

        ' Array of photons to emit

        Dim Photons(N_photon - 1) As Photon

        If Face Is Nothing Then

            Fill_RSOURCE_argument(R_source)

            Dim RSOURCE_position As Vector = R_source.Calc_position

            Photons = Photon.Generate_N_random_photons(RSOURCE_position, N_photon, New_opening_angle, CDF, Argument, Emitter_medium, Correction_coef * Total_outcoming_heat_power)

        Else

            Photons = Photon.Generate_N_random_photons(Face, N_photon, New_opening_angle, CDF, Argument, Emitter_medium, Total_outcoming_heat_power)

            'Photons = Photon.Generate_N_random_photons(Face, 1000000, New_opening_angle, CDF, Argument, Emitter_medium, Total_outcoming_heat_power)

        End If


        For i As Integer = 0 To N_photon - 1

            ' Emitting random photon

            i_Photon += 1

            Photon = Photons(i)

            ' Transforming photon coordinates from emitter coordinate system to global coordinate system

            Photon.Direction = Photon.Direction.Coord(0) * ox + Photon.Direction.Coord(1) * oy + Photon.Direction.Coord(2) * oz

            ' Propagation cycle

            ' First photon emitted from corrected emitter position

            If Source_distance > Limit_distance Then

                Photon.Center = Source_location + (L_vector(FEM.Center - Source_location) - FEM.Maximum_length) * Photon.Direction

            End If

            ' if we have to emit spherically  
            If Direction_should_be_reverted Then

                ' we have to reverse emission direction after each proton to fill the sphere 
                ox = -1 * ox

                oy = -1 * oy

                oz = -1 * oz

            End If


            Do

                view_from = Photon.Center + FEM.Minimum_length * 0.001 * Photon.Direction

                view_to = Photon.Center + Photon.Direction

                Face_number = vf.View_to_vector(view_from.Coord(0), view_from.Coord(1), view_from.Coord(2), view_to.Coord(0), view_to.Coord(1), view_to.Coord(2))

                If Face_number <> 0 Then

                    Fill_Face_argument(FEM.Shining_face(Face_number - 1), Photon)

                    Photon = FEM.Shining_face(Face_number - 1).Face_photon_interaction(Photon)

                    ' If photon absorbed, exit cycle
                    If Photon Is Nothing Then

                        Exit Do

                    End If

                Else

                    ' If photon is passed away from model, exit cycle

                    Exit Do

                End If

            Loop

        Next i

    End Sub

    '''' <summary>
    '''' This function calculates illimination of model from given face with photons.
    '''' The emission center in each act of emission is located randomly on face surface.
    '''' The photon emission is directed by Direction within Opening_angle.
    '''' The function returns array of energies absorbed in model's elements.    
    '''' </summary>
    '''' <param name="Face">Face to emit photons</param>
    '''' <param name="Direction">Direction of radiation propagation</param>
    '''' <param name="Opening_angle">Opening angle in radiation propagation cone
    '''' If Opening_angle = PI, photons emitted in all of directions, spherically.
    '''' If Opening_angle = PI/2, photons emitted in one half-space.
    '''' If Opening_angle more then 0 and less than  PI/2, photons emitted in cone.</param>
    '''' <param name="CDF">Radiation source cumulative distribution of photon frequency</param>
    '''' <param name="Argument">Radiation source paratemers, such as a temperature, time and so on.</param>
    '''' <param name="Emitter_medium">Medium, surrounded radiation source.</param>
    '''' <param name="vf">Visualisation form, used for photon-face interaction detection</param>
    '''' <param name="Step_epsilon">Maximum available error in view-factor-vector calculation process</param>
    '''' <param name="Total_outcoming_heat_power">Total heat power outcoming from source</param>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public Function Illuminate_model_from_face_and_calc_radiation_Q(ByRef Face As Face, ByRef Direction As Vector, ByRef Opening_angle As Double, ByRef CDF As DEPEND, ByRef Argument As Vector, ByRef Emitter_medium As MEDIUM1, ByRef vf As Visualisation.Visualisation, ByRef Step_epsilon As Double, ByRef Total_outcoming_heat_power As Double) As Double(,)

    '    Illuminate_model_from_face_and_calc_radiation_Q = Nothing

    '    ' Number of protons, emitted before steady-state attainment definition
    '    Dim Steady_state_photon_number As Integer = 1

    '    ' Number of photons, emitted before steady-state attainment checks
    '    Dim Photons_between_checks As Integer = 10

    '    Dim Next_check As Integer = Steady_state_photon_number

    '    ' Vectors with previous and current absorption coefficients
    '    Dim Steady_state_vector_0, Steady_state_vector_1 As Vector

    '    ' Length of vectors with previous and current absorption coefficients
    '    Dim L_Steady_state_vector_0, L_Steady_state_vector_1 As Double

    '    L_Steady_state_vector_0 = 0

    '    L_Steady_state_vector_1 = 0

    '    Steady_state_vector_0 = New Vector

    '    ReDim Steady_state_vector_0.Coord(UBound(FEM.Shining_face))

    '    Steady_state_vector_1 = New Vector

    '    ReDim Steady_state_vector_1.Coord(UBound(FEM.Shining_face))

    '    ' Previous and current vector difference. Vector compopent, whose difference is less than set maximum
    '    ' are excluded from error difference calculation
    '    Dim Vector_error As Double = Double.MaxValue


    '    ' Feed, illimination angle calculation

    '    Dim Limit_distance As Double = vf.Max_length

    '    ' The random source location calculation

    '    Dim Source_location As Vector


    '    ' If radiation source is located farther then Limit_distance, then it can be shifted to it distance.

    '    ' If radiation source is located farther then Limit_distance, then its opening angle should be shifted to Math.Atan(FEM.Maximum_length/Limit_distance).

    '    ' If radiation source is shifted to new position, it is needed to correct absorption in model elements.
    '    ' It is caused by part of photons guarantily passed away from model and can't be absorbed.

    '    Dim to_Center As Vector = FEM.Center - Face.Center

    '    Dim Source_distance As Double = L_vector(to_Center)

    '    'Dim New_source_location As New Vector

    '    Dim New_direction As New Vector

    '    Dim New_opening_angle As Double

    '    Dim Direction_should_be_reverted As Boolean = False

    '    ' This coefficient is used for absorption correction due to illimination angle changing
    '    Dim Correction_coef As Double

    '    If Source_distance > Limit_distance Then


    '        New_opening_angle = Math.Atan(FEM.Maximum_length / 2 / Source_distance)

    '        Correction_coef = Math.Sin(0.5 * New_opening_angle) ^ 2

    '        New_direction = to_Center

    '        If Math.Abs(New_direction.Coord(0)) < Epsilon Then New_direction.Coord(0) = 0

    '        If Math.Abs(New_direction.Coord(1)) < Epsilon Then New_direction.Coord(1) = 0

    '        If Math.Abs(New_direction.Coord(2)) < Epsilon Then New_direction.Coord(2) = 0


    '    Else

    '        'New_source_location = Source_location

    '        New_opening_angle = Opening_angle

    '        If New_opening_angle = Math.PI Then

    '            New_opening_angle = 0.5 * Math.PI

    '            Direction_should_be_reverted = True

    '        End If

    '        New_direction = Direction

    '        Correction_coef = 1

    '    End If


    '    'New_opening_angle = Opening_angle

    '    'New_direction = Direction

    '    'Correction_coef = 1


    '    ' Coordinate system building

    '    Dim ox, oy, oz As Vector

    '    oz = Norm_vector(New_direction)

    '    If oz.Coord(0) = 0 And oz.Coord(1) = 0 And Math.Abs(oz.Coord(2)) = 1 Then

    '        ox = New Vector

    '        ox.Coord(0) = 1

    '        ox.Coord(1) = 0

    '        ox.Coord(2) = 0


    '        oy = New Vector

    '        oy.Coord(0) = 0

    '        oy.Coord(1) = 1

    '        oy.Coord(2) = 0

    '    Else

    '        oy = New Vector

    '        oy.Coord(0) = 0

    '        oy.Coord(1) = 0

    '        oy.Coord(2) = 1

    '        ox = oy ^ oz

    '        oy = oz ^ ox

    '        ox = Norm_vector(ox)

    '        oy = Norm_vector(oy)

    '        oz = Norm_vector(oz)

    '    End If

    '    ' Model initializatioan and clearing

    '    For i As Integer = 0 To UBound(FEM.Shining_face)

    '        'ReDim FEM.Shining_face(i).Absorption_act(Max_items_default)

    '        'FEM.Shining_face(i).i_Absorption_act = -1

    '        FEM.Shining_face(i).Interaction = New Face_photon_interaction

    '    Next i


    '    ' Photon emission cycle
    '    Dim Photon As Photon

    '    ' Face, hitted by photon
    '    Dim Face_number As Integer

    '    Const N_Photon_max = 1000000

    '    Dim i_Face_with_absorbtion_number_0 As Integer = -1

    '    Dim i_Face_with_absorbtion_number_1 As Integer = -1

    '    Dim Face_with_absorbtion_number_0(N_Photon_max) As Integer

    '    Dim Face_with_absorbtion_number_1(N_Photon_max) As Integer

    '    Dim i_Photon As Integer = 0

    '    Do

    '        ' Emitting random photon

    '        i_Photon += 1

    '        Photon = New Photon

    '        Source_location = Face.Calc_random_emission_center

    '        Photon = Photon.Emit_random_photon(Source_location, New_opening_angle, CDF, Argument)

    '        Photon.Number = i_Photon

    '        Photon.Medium = Emitter_medium

    '        ' Transforming photon coordinates from emitter coordinate system to global coordinate system

    '        Photon.Direction = Photon.Direction.Coord(0) * ox + Photon.Direction.Coord(1) * oy + Photon.Direction.Coord(2) * oz

    '        ' Propagation cycle

    '        ' First photon emitted from corrected emitter position

    '        If Source_distance > Limit_distance Then

    '            'Photon.Center = Source_location + (Source_distance - 0.9 * Limit_distance) * Photon.Direction

    '            Photon.Center = Source_location + (L_vector(FEM.Center - Source_location) - FEM.Maximum_length) * Photon.Direction

    '        End If

    '        'if we have to emit spherically  
    '        If Direction_should_be_reverted Then

    '            ' we have to reverse emission direction after each proton to fill the sphere 
    '            ox = -1 * ox

    '            oy = -1 * oy

    '            oz = -1 * oz

    '        End If

    '        Do

    '            Face_number = View_to_vector(vf, Photon.Center + FEM.Minimum_length * 0.001 * Photon.Direction, Photon.Center + Photon.Direction)

    '            If Face_number <> 0 Then

    '                Fill_Face_argument(FEM.Shining_face(Face_number - 1), Photon)

    '                Photon = FEM.Shining_face(Face_number - 1).Face_photon_interaction(Photon)

    '                ' If photon absorbed, exit cycle
    '                If Photon Is Nothing Then

    '                    i_Face_with_absorbtion_number_0 += 1

    '                    Face_with_absorbtion_number_0(i_Face_with_absorbtion_number_0) = FEM.Shining_face(Face_number - 1).Number

    '                    Exit Do

    '                End If

    '            Else

    '                ' If photon is passed away from model, exit cycle

    '                Exit Do

    '            End If

    '        Loop


    '        '' Absorbtion coeffients error evaluation

    '        '' If we emitted adequate number of photons, we can check for steady-state mode 
    '        'If i_Photon >= Steady_state_photon_number Then

    '        '    ' If it is time for steady-state mode check
    '        '    If i_Photon = Next_check Then

    '        '        Next_check += Photons_between_checks

    '        '        ''If i_Face_with_absorbtion_number_0 = -1 Then Exit Do

    '        '        ''Forming array of faces absorbed the photons numbers

    '        '        'Dim Is_in_array As Boolean

    '        '        ''i_Face_with_absorbtion_number_1 = -1

    '        '        'For i As Integer = 0 To i_Face_with_absorbtion_number_0

    '        '        '    Is_in_array = False

    '        '        '    For j As Integer = 0 To i_Face_with_absorbtion_number_1

    '        '        '        If Face_with_absorbtion_number_0(i) = Face_with_absorbtion_number_1(j) Then

    '        '        '            Is_in_array = True

    '        '        '            Exit For

    '        '        '        End If

    '        '        '    Next j

    '        '        '    If Not (Is_in_array) Then

    '        '        '        i_Face_with_absorbtion_number_1 += 1

    '        '        '        Face_with_absorbtion_number_1(i_Face_with_absorbtion_number_1) = Face_with_absorbtion_number_0(i)

    '        '        '    End If

    '        '        'Next i

    '        '        'i_Face_with_absorbtion_number_0 = -1


    '        '        ' if it is first steady state vector, we can't calculate difference.
    '        '        If i_Photon = Steady_state_photon_number Then

    '        '            ' Current vector calculation
    '        '            For i As Integer = 0 To UBound(Steady_state_vector_1.Coord)

    '        '                Steady_state_vector_1.Coord(i) = Correction_coef * (FEM.Shining_face(i).Interaction.N_A) / i_Photon

    '        '                L_Steady_state_vector_1 += Steady_state_vector_1.Coord(i) ^ 2

    '        '            Next i

    '        '            L_Steady_state_vector_1 = Math.Sqrt(L_Steady_state_vector_1)

    '        '        Else


    '        '            ' Saving previous vector

    '        '            For i As Integer = 0 To UBound(Steady_state_vector_1.Coord)

    '        '                Steady_state_vector_0.Coord(i) = Steady_state_vector_1.Coord(i)

    '        '            Next i

    '        '            L_Steady_state_vector_0 = L_Steady_state_vector_1

    '        '            ' Current vector calculation and vector difference evaluation

    '        '            Dim Component_error As Double = Double.MaxValue

    '        '            For i As Integer = 0 To UBound(Steady_state_vector_1.Coord)

    '        '                Steady_state_vector_1.Coord(i) = Correction_coef * (FEM.Shining_face(i).Interaction.N_A) / i_Photon

    '        '                Component_error = Math.Abs(Steady_state_vector_1.Coord(i) - Steady_state_vector_0.Coord(i))

    '        '                ' If vector component error is more than set error this component is included into error calculation
    '        '                If Component_error > Step_epsilon Then

    '        '                    L_Steady_state_vector_1 += Steady_state_vector_1.Coord(i) ^ 2

    '        '                End If

    '        '            Next i

    '        '            L_Steady_state_vector_1 = Math.Sqrt(L_Steady_state_vector_1)

    '        '            ' Error evaluation

    '        '            Vector_error = Math.Abs(L_Steady_state_vector_1 - L_Steady_state_vector_0)

    '        '            ' If steady-state mode is archived 
    '        '            If Vector_error <= Step_epsilon Then

    '        '                ' then calculate incoming heat power to model's elements

    '        '                Dim Return_array(1, UBound(Steady_state_vector_1.Coord)) As Double

    '        '                Dim i_eluminated_face As Integer = -1

    '        '                For i As Integer = 0 To UBound(Steady_state_vector_1.Coord)

    '        '                    If Steady_state_vector_1.Coord(i) > 0 Then

    '        '                        i_eluminated_face += 1

    '        '                        Return_array(0, i_eluminated_face) = i

    '        '                        Return_array(1, i_eluminated_face) = Steady_state_vector_1.Coord(i) * Total_outcoming_heat_power

    '        '                    End If

    '        '                Next i

    '        '                If i_eluminated_face = -1 Then

    '        '                    Return Nothing

    '        '                Else

    '        '                    ReDim Preserve Return_array(1, i_eluminated_face)

    '        '                    Return Return_array

    '        '                End If


    '        '            End If

    '        '        End If

    '        '    End If

    '        'End If

    '        ' If we emitted adequate number of photons, we can check for steady-state mode 
    '        If i_Photon >= Steady_state_photon_number Then

    '            ' If it is time for steady-state mode check
    '            If i_Photon = Next_check Then

    '                ' Begin steady-state mode check

    '                Next_check += Photons_between_checks

    '                'If i_Face_with_absorbtion_number_0 = -1 Then Exit Do

    '                'Forming array of faces absorbed the photons numbers

    '                Dim Is_in_array As Boolean

    '                'i_Face_with_absorbtion_number_1 = -1

    '                For i As Integer = 0 To i_Face_with_absorbtion_number_0

    '                    Is_in_array = False

    '                    For j As Integer = 0 To i_Face_with_absorbtion_number_1

    '                        If Face_with_absorbtion_number_0(i) = Face_with_absorbtion_number_1(j) Then

    '                            Is_in_array = True

    '                            Exit For

    '                        End If

    '                    Next j

    '                    If Not (Is_in_array) Then

    '                        i_Face_with_absorbtion_number_1 += 1

    '                        Face_with_absorbtion_number_1(i_Face_with_absorbtion_number_1) = Face_with_absorbtion_number_0(i)

    '                    End If

    '                Next i

    '                i_Face_with_absorbtion_number_0 = -1

    '                ' if it is first steady state vector, we can't calculate difference.
    '                If i_Photon = Steady_state_photon_number Then

    '                    ReDim Steady_state_vector_1.Coord(i_Face_with_absorbtion_number_1)

    '                    ' Current vector calculation
    '                    For i As Integer = 0 To i_Face_with_absorbtion_number_1

    '                        Steady_state_vector_1.Coord(i) = Correction_coef * (FEM.Shining_face(Face_with_absorbtion_number_1(i)).Interaction.N_A) / i_Photon

    '                        L_Steady_state_vector_1 += Steady_state_vector_1.Coord(i) ^ 2

    '                    Next i

    '                    L_Steady_state_vector_1 = Math.Sqrt(L_Steady_state_vector_1)

    '                Else


    '                    ' Saving previous vector

    '                    ReDim Steady_state_vector_0.Coord(i_Face_with_absorbtion_number_1)

    '                    For i As Integer = 0 To UBound(Steady_state_vector_1.Coord)

    '                        Steady_state_vector_0.Coord(i) = Steady_state_vector_1.Coord(i)

    '                    Next i

    '                    L_Steady_state_vector_0 = L_Steady_state_vector_1

    '                    ' Current vector calculation and vector difference evaluation

    '                    Dim Component_error As Double = Double.MaxValue

    '                    ' Current vector calculation

    '                    ReDim Steady_state_vector_1.Coord(i_Face_with_absorbtion_number_1)

    '                    For i As Integer = 0 To i_Face_with_absorbtion_number_1

    '                        Steady_state_vector_1.Coord(i) = Correction_coef * (FEM.Shining_face(Face_with_absorbtion_number_1(i)).Interaction.N_A) / i_Photon

    '                        Component_error = Math.Abs(Steady_state_vector_1.Coord(i) - Steady_state_vector_0.Coord(i))

    '                        ' If vector component error is more than set error this component is included into error calculation
    '                        If Component_error > Step_epsilon Then

    '                            L_Steady_state_vector_1 += Steady_state_vector_1.Coord(i) ^ 2

    '                        End If

    '                    Next i

    '                    L_Steady_state_vector_1 = Math.Sqrt(L_Steady_state_vector_1)

    '                    ' Error evaluation

    '                    Vector_error = Math.Abs(L_Steady_state_vector_1 - L_Steady_state_vector_0)

    '                    ' If steady-state mode is archived 
    '                    If Vector_error <= Step_epsilon Then

    '                        ' then calculate incoming heat power to model's elements

    '                        Dim Return_array(1, i_Face_with_absorbtion_number_1) As Double

    '                        For i As Integer = 0 To i_Face_with_absorbtion_number_1

    '                            Return_array(0, i) = Face_with_absorbtion_number_1(i)

    '                            Return_array(1, i) = Steady_state_vector_1.Coord(i) * Total_outcoming_heat_power

    '                        Next i

    '                        Return Return_array

    '                    End If

    '                End If

    '            End If

    '        End If


    '        Application.DoEvents()

    '    Loop

    'End Function

    '''' <summary>
    '''' This function returns emission center coordinates ramdomly located on face surface
    '''' </summary>
    '''' <param name="Face">Face to emit</param>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public Function Calc_random_emission_center_on_face(ByRef Face As Face) As Vector


    '    'Dim Grid(UBound(Face.Grid)) As Vector

    '    'Dim Xpi(UBound(Face.Grid)) As Double

    '    'Dim Ypi(UBound(Face.Grid)) As Double

    '    'Dim x, y, Xmax, Xmin, Ymax, Ymin As Double

    '    'Xmax = -Double.MaxValue

    '    'Ymax = -Double.MaxValue

    '    'Xmin = Double.MaxValue

    '    'Ymin = Double.MaxValue

    '    'For i As Integer = 0 To UBound(Grid)

    '    '    Grid(i) = New Vector

    '    '    Grid(i).Coord(0) = Face.Grid(i).Position.Coord(0)

    '    '    Grid(i).Coord(1) = Face.Grid(i).Position.Coord(1)

    '    '    Grid(i).Coord(2) = Face.Grid(i).Position.Coord(2)

    '    '    Grid(i) -= Face.Center

    '    '    Grid(i) = Face.Transform_to_face(Grid(i))

    '    '    Xpi(i) = Grid(i).Coord(0)

    '    '    Ypi(i) = Grid(i).Coord(1)

    '    '    ' Max and min coordinates determination

    '    '    If Xpi(i) > Xmax Then

    '    '        Xmax = Xpi(i)

    '    '    End If

    '    '    If Xpi(i) < Xmin Then

    '    '        Xmin = Xpi(i)

    '    '    End If

    '    '    If Ypi(i) > Ymax Then

    '    '        Ymax = Ypi(i)

    '    '    End If

    '    '    If Ypi(i) < Ymin Then

    '    '        Ymin = Ypi(i)

    '    '    End If

    '    'Next i

    '    ' Random point generation

    '    Dim x, y As Double

    '    Do

    '        x = (Xmax - Xmin) * MT_RND.genrand_real1 + Xmin

    '        y = (Ymax - Ymin) * MT_RND.genrand_real1 + Ymin

    '        ' if point is located on the face, then exit do
    '        If IsPointInPolygon(x, y, Xpi, Ypi) Then Exit Do

    '    Loop

    '    ' Transform random point back to global coordinate system

    '    Dim RND_point As New Vector

    '    RND_point.Coord(0) = x

    '    RND_point.Coord(1) = y

    '    RND_point.Coord(2) = 0

    '    RND_point = Face.Transform_from_face(RND_point)

    '    RND_point += Face.Center

    '    Return RND_point

    'End Function


    Public Function View_to_vector(ByRef vf As Visualisation.Visualisation, ByRef View_from As Vector, ByRef View_to As Vector) As Long
        'Public Function View_to_vector(ByVal vf As Visualisation.Visualisation, ByVal View_from As Vector, ByRef View_to As Vector) As Long

        Return vf.View_to_vector(View_from.Coord(0), View_from.Coord(1), View_from.Coord(2), View_to.Coord(0), View_to.Coord(1), View_to.Coord(2))

    End Function

    ''' <summary>
    ''' This function returns one element of radiation heat transfer matrix with indexes (i,j)
    ''' </summary>
    ''' <param name="i"></param>
    ''' <param name="j"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function B_matrix_element(ByVal i As Long, ByVal j As Long) As Double

        ' пока мы не освоили метод Ланцоша, поэтому просто возвращаем элемент

        B_matrix_element = Me.B_matrix(i, j)

    End Function
    ''' <summary>
    ''' This sub calculates vector-argument for a heat transferring element.
    ''' Table of parameters (vector-Argument components)
    ''' 0 - temperature,
    ''' 1 - time,
    ''' 2 - random value,
    ''' 3 - incoming radiation frequency,
    ''' 4 - incoming radiation angle of elevation in local coordinate system,
    ''' 5 - incoming radiation azimuth in local coordinate system.
    ''' </summary>
    ''' <param name="HT_element"></param>
    ''' <remarks></remarks>
    Public Sub Fill_HT_element_argument(ByRef HT_element As HT_Element)

        HT_element.Argument.Coord(0) = HT_element.Element.Field.Temperature

        HT_element.Argument.Coord(1) = Current_time

        HT_element.Argument.Coord(2) = 0

        HT_element.Argument.Coord(3) = 0

        HT_element.Argument.Coord(4) = 0

        HT_element.Argument.Coord(5) = 0

    End Sub
    ''' <summary>
    ''' This sub calculates vector-argument for a face.
    ''' Table of parameters (vector-Argument components)
    ''' 0 - temperature,
    ''' 1 - time,
    ''' 2 - random value,
    ''' 3 - incoming radiation frequency,
    ''' 4 - incoming radiation angle of elevation in local coordinate system,
    ''' 5 - incoming radiation azimuth in local coordinate system.
    ''' </summary>
    ''' <param name="Face">Interacting face</param>
    ''' <param name="in_Photon">Incoming photon</param>
    ''' <remarks></remarks>
    Public Sub Fill_Face_argument(ByRef Face As Face, ByRef in_Photon As Photon)

        Face.Argument.Coord(0) = Face.Element.Field.Temperature

        Face.Argument.Coord(1) = Current_time

        Face.Argument.Coord(2) = 0

        Face.Argument.Coord(3) = in_Photon.Frequency

        Face.Argument.Coord(4) = Math.Acos(Math.Sqrt(in_Photon.Direction.Coord(0) ^ 2 + in_Photon.Direction.Coord(1) ^ 2) / Math.Sqrt(in_Photon.Direction.Coord(0) ^ 2 + in_Photon.Direction.Coord(1) ^ 2 + in_Photon.Direction.Coord(2) ^ 2 + Epsilon))

        Face.Argument.Coord(5) = Math.Atan2(in_Photon.Direction.Coord(1), in_Photon.Direction.Coord(0))

    End Sub


    ''' <summary>
    ''' This sub calculates vector-argument for a radiation heat source.
    ''' Table of parameters (vector-Argument components)
    ''' 0 - temperature,
    ''' 1 - time,
    ''' 2 - random value,
    ''' 3 - incoming radiation frequency,
    ''' 4 - incoming radiation angle of elevation in local coordinate system,
    ''' 5 - incoming radiation azimuth in local coordinate system.
    ''' </summary>
    ''' <param name="RSOURCE">Interacting face</param>
    ''' <remarks></remarks>
    Public Sub Fill_RSOURCE_argument(ByRef RSOURCE As RSOURCE1)

        With RSOURCE

            .Argument.Coord(0) = 0

            .Argument.Coord(1) = Current_time

            .Argument.Coord(2) = 0

            .Argument.Coord(3) = 0

            .Argument.Coord(4) = 0

            .Argument.Coord(5) = 0

            .Argument.Coord(0) = .Calc_T(.Argument)

        End With

    End Sub


    ''' <summary>
    ''' This sub calculates vector-argument for a heat source.
    ''' Table of parameters (vector-Argument components)
    ''' 0 - temperature,
    ''' 1 - time,
    ''' 2 - random value,
    ''' 3 - incoming radiation frequency,
    ''' 4 - incoming radiation angle of elevation in local coordinate system,
    ''' 5 - incoming radiation azimuth in local coordinate system.
    ''' </summary>
    ''' <param name="HSOURCE">Interacting face</param>
    ''' <remarks></remarks>
    Public Sub Fill_HSOURCE_argument(ByRef HSOURCE As HSOURCE1)

        With HSOURCE

            .Argument.Coord(0) = 0

            .Argument.Coord(1) = Current_time

            .Argument.Coord(2) = 0

            .Argument.Coord(3) = 0

            .Argument.Coord(4) = 0

            .Argument.Coord(5) = 0

        End With

    End Sub


    '''' <summary>
    '''' This sub calculate total thermal power to element ( if power to element - it is >0)
    '''' </summary>
    '''' <param name="HT_element"></param>
    '''' <remarks></remarks>
    'Public Sub Calc_Q(ByRef HT_element As HT_Element)

    '    Dim Q_R_income, Q_R_outcome, Q_lambda, Temperature_1, Temperature_2, Area, eps As Double

    '    Dim Q_Heat_source, Q_Radio_source As Double

    '    Dim lambda_1, lambda_2, lambda As Double

    '    Dim mass_1, mass_2, mass As Double


    '    ' Рассчитываем температуру элемента

    '    Temperature_1 = HT_element.Element.Field.Temperature

    '    ' Рассчитываем вектор-аргумент для последующей передачи в зависимости

    '    HT_element.Argument.Coord(0) = Temperature_1

    '    HT_element.Argument.Coord(1) = Current_time

    '    HT_element.Argument.Coord(2) = 0

    '    HT_element.Argument.Coord(3) = 0

    '    HT_element.Argument.Coord(4) = 0

    '    HT_element.Argument.Coord(5) = 0

    '    ' Рассчитываем тепловой поток во все стороны элемента за счет излучения

    '    For i As Integer = 0 To UBound(HT_element.Face)

    '        ' Рассчитываем площадь 

    '        Area = HT_element.Face(i).Area

    '        ' Рассчитываем коэффициент черноты для длины волны, на которую приходится максимум теплового потока

    '        If HT_element.Face(i).FTOP.A.GetType.ToString <> "System.Double" Then

    '            Dim Argument As New Vector

    '            ReDim Argument.Coord(5)

    '            Argument.Coord(0) = Temperature_1

    '            Argument.Coord(1) = 0

    '            Argument.Coord(2) = 0

    '            'Argument.Coord(3) = HT_element.Face(i).c / HT_element.Face(i).Max_emittance_frequency(Temperature_1)

    '            Argument.Coord(4) = 0

    '            Argument.Coord(5) = 0

    '            eps = HT_element.Face(i).FTOP.A.return_value(Argument).Coord(0)

    '        End If

    '        If HT_element.Face(i).FTOP.Eps.GetType.ToString = "System.Double" Then

    '            eps = HT_element.Face(i).FTOP.Eps

    '        End If

    '        Q_R_outcome += Area * eps * HT_element.Face(0).sigma * Temperature_1 ^ 4

    '    Next i

    '    ' Рассчитываем тепловой поток к элементу от других элементов за счет излучения

    '    For i As Integer = 0 To UBound(Me.HT_element)

    '        Q_R_income += B_matrix_element(i, HT_element.Number) * HT_element.Face(0).sigma * Me.HT_element(i).Element.Field.Temperature ^ 4

    '    Next i

    '    ' Рассчитываем тепловой поток от соседних элементов

    '    For i As Integer = 0 To UBound(HT_element.Heat_conductivity)

    '        With HT_element.Heat_conductivity(i)

    '            ' Рассчитываем аргумент соседа

    '            Fill_HT_element_argument(.HT_Element)

    '            ' Рассчитываем температуру элемента соседа

    '            Temperature_2 = HT_element.Heat_conductivity(i).HT_Element.Element.Field.Temperature

    '            ' Рассчитываем коэффциент теплопроводности для теплобмена

    '            lambda_1 = HT_element.Element.Medium.Calc_lambda(HT_element.Argument)

    '            lambda_2 = .HT_Element.Element.Medium.Calc_lambda(.HT_Element.Argument)

    '            mass_1 = HT_element.Element.Field.Mass

    '            mass_2 = .HT_Element.Element.Field.Mass

    '            mass = mass_1 + mass_2

    '            lambda = lambda_1 * mass_1 / mass + lambda_2 * mass_2 / mass

    '            ' Расcчитываем тепловую мощность

    '            Q_lambda += .Area * lambda * (Temperature_2 - Temperature_1) / .Distance

    '        End With

    '    Next i

    '    ' рассчитываем тепловой поток от источников тепла

    '    Q_Heat_source = 0

    '    If UBound(HT_element.Heat_source) >= 0 Then

    '        For i As Integer = 0 To UBound(HT_element.Heat_source)

    '            Q_Heat_source += HT_element.Heat_source(i).Heat(HT_element.Argument)

    '        Next i

    '    End If

    '    ' рассчитываем тепловой поток от источников излучения

    '    Q_Radio_source = 0

    '    For i As Integer = 0 To UBound(FEM.RSOURCE)

    '        If FEM.RSOURCE(i).B_row(HT_element.Number) > 0 Then

    '            With FEM.RSOURCE(i)

    '                Fill_RSOURCE_argument(FEM.RSOURCE(i))

    '                Q_Heat_source += .Calc_Q(.Argument) * .B_row(HT_element.Number)

    '            End With

    '        End If

    '    Next i

    '    HT_element.Q_summary = Q_R_income - Q_R_outcome + Q_lambda + Q_Heat_source + Q_Radio_source

    'End Sub

    ''' <summary>
    ''' This sub calculate total thermal power to element ( if power to element - it is >0)
    ''' </summary>
    ''' <param name="HT_element"></param>
    ''' <remarks></remarks>
    Public Sub Calc_Q_const(ByRef HT_element As HT_Element)

        Dim Q_R_income, Q_R_outcome, Q_lambda, Temperature_1, Temperature_2, Area, eps As Double

        Dim Q_Heat_source, Q_Radio_source As Double

        Dim lambda_1, lambda_2 As Double

        ' Рассчитываем температуру элемента

        Temperature_1 = HT_element.Element.Field.Temperature

        ' Рассчитываем тепловой поток во все стороны элемента за счет излучения

        For i As Integer = 0 To UBound(HT_element.Face)

            ' Рассчитываем площадь 

            Area = HT_element.Face(i).Area

            ' Рассчитываем коэффициент черноты для длины волны, на которую приходится максимум теплового потока

            eps = HT_element.Face(i).FTOP.A

            Q_R_outcome += Area * eps * HT_element.Face(0).sigma * Temperature_1 ^ 4

        Next i

        ' Рассчитываем тепловой поток к элементу от других элементов за счет излучения

        For i As Integer = 0 To UBound(Me.HT_element)

            Q_R_income += B_matrix_element(i, HT_element.Number) * HT_element.Face(0).sigma * Me.HT_element(i).Element.Field.Temperature ^ 4

        Next i

        ' Рассчитываем тепловой поток от соседних элементов

        Q_lambda = 0

        For i As Integer = 0 To UBound(HT_element.Heat_conductivity)

            With HT_element.Heat_conductivity(i)

                ' Рассчитываем аргумент соседа

                Fill_HT_element_argument(.HT_Element)

                ' Рассчитываем температуру элемента соседа

                Temperature_2 = HT_element.Heat_conductivity(i).HT_Element.Element.Field.Temperature

                ' Рассчитываем коэффциент теплопроводности для теплобмена

                lambda_1 = HT_element.Element.Medium.Calc_lambda(HT_element.Argument)

                lambda_2 = .HT_Element.Element.Medium.Calc_lambda(.HT_Element.Argument)

                ' Расcчитываем тепловую мощность

                Q_lambda += .Area * (lambda_1 * .Distance_1 / (.Distance_1 + .Distance_2) + lambda_2 * .Distance_2 / (.Distance_1 + .Distance_2)) * (Temperature_2 - Temperature_1) / (.Distance_1 + .Distance_2)

            End With

        Next i

        ' рассчитываем тепловой поток от источников тепла

        Q_Heat_source = 0

        If UBound(HT_element.Heat_source) >= 0 Then

            For i As Integer = 0 To UBound(HT_element.Heat_source)

                Q_Heat_source += HT_element.Heat_source(i).Heat(HT_element.Argument)

            Next i

        End If

        ' рассчитываем тепловой поток от источников излучения

        Q_Radio_source = 0

        For i As Integer = 0 To UBound(FEM.RSOURCE)

            If FEM.RSOURCE(i).B_row(HT_element.Number) > 0 Then

                With FEM.RSOURCE(i)

                    Fill_RSOURCE_argument(FEM.RSOURCE(i))

                    Q_Heat_source += .Calc_Q(.Argument) * .B_row(HT_element.Number)

                End With

            End If

        Next i


        If HT_element.Element.Number = 108 Then

            HT_element.Element.Number = 108

        End If


        HT_element.Q_summary = Q_R_income - Q_R_outcome + Q_lambda + Q_Heat_source + Q_Radio_source

    End Sub

    '''' <summary>
    '''' This sub calculate total thermal power to element ( if power to element - it is >0)
    '''' </summary>
    '''' <param name="HT_element"></param>
    '''' <remarks></remarks>
    'Public Sub Calc_Q_const_matrixless(ByRef HT_element As HT_Element)

    '    Dim Q_R_income, Q_R_outcome, Q_lambda, Temperature_1, Temperature_2 As Double

    '    Dim Q_Heat_source, Q_Radio_source As Double

    '    Dim lambda_1, lambda_2, lambda As Double

    '    Dim mass_1, mass_2, mass As Double




    '    ' Рассчитываем тепловой поток от соседних элементов

    '    For i As Integer = 0 To UBound(HT_element.Heat_conductivity)

    '        With HT_element.Heat_conductivity(i)

    '            ' Рассчитываем температуру элемента соседа

    '            Temperature_2 = HT_element.Heat_conductivity(i).HT_Element.Element.Field.Temperature

    '            ' Рассчитываем коэффциент теплопроводности для теплобмена

    '            lambda_1 = HT_element.Element.Medium.lambda(0)

    '            lambda_2 = .HT_Element.Element.Medium.lambda(0)

    '            mass_1 = HT_element.Element.Field.Mass

    '            mass_2 = .HT_Element.Element.Field.Mass

    '            mass = mass_1 + mass_2

    '            lambda = lambda_1 * mass_1 / mass + lambda_2 * mass_2 / mass

    '            ' Расcчитываем тепловую мощность

    '            Q_lambda += .Area * lambda * (Temperature_2 - Temperature_1) / .Distance

    '        End With

    '    Next i

    '    ' рассчитываем тепловой поток от источников тепла

    '    Q_Heat_source = 0

    '    If UBound(HT_element.Heat_source) >= 0 Then

    '        For i As Integer = 0 To UBound(HT_element.Heat_source)

    '            Q_Heat_source += HT_element.Heat_source(i).Heat(HT_element.Argument)

    '        Next i

    '    End If

    '    ' рассчитываем тепловой поток от источников излучения

    '    Q_Radio_source = 0

    '    For i As Integer = 0 To UBound(FEM.RSOURCE)

    '        If FEM.RSOURCE(i).B_row(HT_element.Number) > 0 Then

    '            With FEM.RSOURCE(i)

    '                Fill_RSOURCE_argument(FEM.RSOURCE(i))

    '                Q_Heat_source += .Calc_Q(.Argument) * .B_row(HT_element.Number)

    '            End With

    '        End If

    '    Next i

    '    HT_element.Q_summary = Q_R_income - Q_R_outcome + Q_lambda + Q_Heat_source + Q_Radio_source

    'End Sub
    '''' <summary>
    '''' This sub calculate total thermal power to element ( if power to element - it is >0)
    '''' </summary>
    '''' <param name="HT_element"></param>
    '''' <remarks></remarks>
    'Public Sub Calc_Q_const_NO_RADIATION_HEAT_TRANSFER(ByRef HT_element As HT_Element)

    '    Dim Q_R_outcome, Q_lambda, Temperature_1, Temperature_2, Area, eps As Double

    '    Dim Q_Heat_source, Q_Radio_source As Double

    '    Dim lambda_1, lambda_2, lambda As Double

    '    Dim mass_1, mass_2, mass As Double


    '    ' Рассчитываем температуру элемента

    '    Temperature_1 = HT_element.Element.Field.Temperature

    '    ' Рассчитываем тепловой поток во все стороны элемента за счет излучения

    '    For i As Integer = 0 To UBound(HT_element.Face)

    '        ' Рассчитываем площадь 

    '        Area = HT_element.Face(i).Area

    '        ' Рассчитываем коэффициент черноты для длины волны, на которую приходится максимум теплового потока

    '        eps = HT_element.Face(i).FTOP.A

    '        Q_R_outcome += Area * eps * HT_element.Face(0).sigma * Temperature_1 ^ 4

    '    Next i


    '    ' Рассчитываем тепловой поток от соседних элементов

    '    For i As Integer = 0 To UBound(HT_element.Heat_conductivity)

    '        With HT_element.Heat_conductivity(i)

    '            ' Рассчитываем температуру элемента соседа

    '            Temperature_2 = HT_element.Heat_conductivity(i).HT_Element.Element.Field.Temperature

    '            ' Рассчитываем коэффциент теплопроводности для теплобмена

    '            lambda_1 = HT_element.Element.Medium.lambda(0)

    '            lambda_2 = .HT_Element.Element.Medium.lambda(0)

    '            mass_1 = HT_element.Element.Field.Mass

    '            mass_2 = .HT_Element.Element.Field.Mass

    '            mass = mass_1 + mass_2

    '            lambda = lambda_1 * mass_1 / mass + lambda_2 * mass_2 / mass

    '            ' Расcчитываем тепловую мощность

    '            Q_lambda += .Area * lambda * (Temperature_2 - Temperature_1) / .Distance

    '        End With

    '    Next i

    '    ' рассчитываем тепловой поток от источников тепла

    '    Q_Heat_source = 0

    '    If UBound(HT_element.Heat_source) >= 0 Then

    '        For i As Integer = 0 To UBound(HT_element.Heat_source)

    '            Q_Heat_source += HT_element.Heat_source(i).Heat(HT_element.Argument)

    '        Next i

    '    End If

    '    ' рассчитываем тепловой поток от источников излучения

    '    Q_Radio_source = 0

    '    For i As Integer = 0 To UBound(FEM.RSOURCE)

    '        If FEM.RSOURCE(i).B_row(HT_element.Number) > 0 Then

    '            With FEM.RSOURCE(i)

    '                Fill_RSOURCE_argument(FEM.RSOURCE(i))

    '                Q_Heat_source += .Calc_Q(.Argument) * .B_row(HT_element.Number)

    '            End With

    '        End If

    '    Next i

    '    HT_element.Q_summary = -Q_R_outcome + Q_lambda + Q_Heat_source + Q_Radio_source

    'End Sub
    ''' <summary>
    ''' This subroutine solves problem, using stated solver.
    ''' Also, subrutine selects depends element properties on temperature or not and choose appropriate type of solver,
    ''' variable or constant.
    ''' </summary>
    ''' <remarks></remarks>
    Sub Solve_problem()

        For i As Integer = 0 To UBound(FEM.Shining_face)

            FEM.Shining_face(i).Code = New Face_code

        Next i

        FEM.Fill_Shining_faces_colours()

        FEM.Polygon = FEM.Extract_Polygon_from_Face(FEM.Shining_face)

        Init_temperature_field()


        Select Case FEM.Solver_type

            Case "EULER"

                'Fill_B_matrices()

                'Euler_const()

                'Euler_const_MATRIXLESS()

                If FEM.Is_solution_constant Then

                    Fill_B_matrices()

                    Euler_const()

                    '' if no radiation heat transfer, we exclude it from calculation

                    'FEM.Is_radiation_heat_transfer = Check_for_radiation_heat_transfer()

                    'If FEM.Is_radiation_heat_transfer Then

                    '    Euler_const_MATRIXLESS()

                    'Else

                    '    Euler_const_NO_RADIATION_HEAT_TRANSFER()

                    'End If

                Else

                    Throw New Exception("The Euler integration method with constant step for variable thermo-phisical and thermooptical properties is not implemented yet. Sorry!")

                    Exit Select

                    't_0 = timeGetTime()

                    '' временно
                    'Fill_B_matrices()
                    '' временно

                    'Euler_var()

                End If

            Case "EULER_V"

                If FEM.Is_solution_constant Then

                    Fill_B_matrices()

                    Euler_var_T_step_with_smoothing()

                Else

                    'Euler_var_T_step_MATRIXLESS()

                    Euler_var_T_step_with_smoothing_MATRIXLESS()

                End If


            Case "ADAMS_4"

                If FEM.Is_solution_constant Then

                    Fill_B_matrices()

                    Adams_4_const()

                Else

                    Throw New Exception("The Adams integration method with constant step for variable thermo-phisical and thermooptical properties is not implemented yet. Sorry!")

                    Exit Select

                    '' временно
                    'Fill_B_matrices()
                    '' временно

                    'Adams_4_var()

                End If

        End Select

        Current_step = 0

        Calc_Elements_Faces_codes()

    End Sub

    ''' <summary>
    ''' This subroutine resumes solution of the problem, using stated solver.
    ''' The resumption of the problem begins from the stated moment of time.
    ''' If moment of time is not stated then resumption begins from the next to last (предыдущий) step.
    ''' </summary>
    ''' <remarks></remarks>
    Sub Resume_problem()

        For i As Integer = 0 To UBound(FEM.Shining_face)

            FEM.Shining_face(i).Code = New Face_code

        Next i

        FEM.Fill_Shining_faces_colours()

        FEM.Polygon = FEM.Extract_Polygon_from_Face(FEM.Shining_face)

        ' Detect start time and finish time 

        Dim N_step_start As Integer

        Dim Start_time_old As Double = Me.Start_time

        If Me.Resumption_start_time = Double.MaxValue Then

            ' we should begin from next to last (предпоследний) step

            For i As Integer = UBound(HT_element(0).HT_Step) To 0 Step -1

                If HT_element(0).HT_Step(i).T > 0 Then

                    N_step_start = i

                    Me.Start_time = HT_element(0).HT_Step(i).Time

                    Exit For

                End If

            Next i

        Else

            For i As Integer = 0 To UBound(HT_element(0).HT_Step) - 1

                If HT_element(0).HT_Step(i + 1).Time > Resumption_start_time Then

                    N_step_start = i + 1

                    Me.Start_time = HT_element(0).HT_Step(i).Time

                    Exit For

                End If

            Next i

        End If


        If Me.Resumption_length_time = Double.MaxValue Then

            Me.Lenght_time -= (Me.Start_time - Start_time_old)

            ' we should finish at last step, so we should not change Lenght_time

        Else

            Me.Lenght_time = Me.Resumption_length_time

        End If

        Select Case FEM.Solver_type

            Case "EULER"

                'Fill_B_matrices()

                'Euler_const()

                'Euler_const_MATRIXLESS()

                If FEM.Is_solution_constant Then

                    Fill_B_matrices()

                    Euler_const()

                    '' if no radiation heat transfer, we exclude it from calculation

                    'FEM.Is_radiation_heat_transfer = Check_for_radiation_heat_transfer()

                    'If FEM.Is_radiation_heat_transfer Then

                    '    Euler_const_MATRIXLESS()

                    'Else

                    '    Euler_const_NO_RADIATION_HEAT_TRANSFER()

                    'End If

                Else

                    Throw New Exception("The Euler integration method with constant step for variable thermo-phisical and thermooptical properties is not implemented yet. Sorry!")

                    Exit Select

                    't_0 = timeGetTime()

                    '' временно
                    'Fill_B_matrices()
                    '' временно

                    'Euler_var()

                End If

            Case "EULER_V"

                If FEM.Is_solution_constant Then

                    Fill_B_matrices()

                    Euler_var_T_step_with_smoothing()

                Else

                    'Euler_var_T_step_MATRIXLESS()

                    Euler_var_T_step_with_smoothing_MATRIXLESS()

                End If


            Case "ADAMS_4"

                If FEM.Is_solution_constant Then

                    Fill_B_matrices()

                    Adams_4_const()

                Else

                    Throw New Exception("The Adams integration method with constant step for variable thermo-phisical and thermooptical properties is not implemented yet. Sorry!")

                    Exit Select

                    '' временно
                    'Fill_B_matrices()
                    '' временно

                    'Adams_4_var()

                End If

        End Select


        Current_step = 0

        Calc_Elements_Faces_codes()

    End Sub

    ''' <summary>
    ''' This subrutine solves stated problem with usig distributed computations.
    ''' </summary>
    ''' <param name="Work_dir">The directory, where all of work files are placed.</param>
    ''' <remarks></remarks>
    Sub Solve_problem_DC(ByRef Work_dir As String)

        Euler_var_MATRIXLESS_DC(Work_dir)

        Current_step = 0

        Calc_Elements_Faces_codes()

    End Sub

    ''' <summary>
    ''' This function calculates productivity index - the time needed for one step of heat transfer
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function Calc_productivity_index(ByRef vf As Visualisation.Visualisation) As Double

        Dim T_0, T_1 As Double

        ' Рассчитываем тепловые потоки на первом шаге

        T_0 = Microsoft.VisualBasic.DateAndTime.Timer

        Prepare_heat_transfer(1)

        If Me.N_photon > 0 Then

            Me.Calc_R_transfer_N_photon(0, Me.N_photon, vf)

        Else

            Exit Function

        End If

        'If Me.N_photon > 0 Then

        '    Me.Calc_R_transfer_N_photon(0, Me.N_photon, vf)

        'Else

        '    Me.Calc_R_transfer_TOP_Epsilon(0, Me.TOP_Epsilon, vf)

        'End If

        T_1 = Microsoft.VisualBasic.DateAndTime.Timer

        Return T_1 - T_0

    End Function

    ''' <summary>
    ''' This is main sub for client in distributed computations 
    ''' </summary>
    ''' <param name="ID">Precalculated client's ID</param>
    ''' <param name="Work_dir">Work directory</param>
    ''' <param name="FTP_mode">Whether Work_directory is on FTP server?</param>
    ''' <remarks></remarks>
    Sub Solution_DC_client(ByRef ID As Long, ByRef Work_dir As String, ByRef FTP_mode As Boolean)

        ' flag initialization 

        Solution_integrity = True

        ' Preparations for calculation

        ' visiualisation firm creation

        ' model preparation

        If FEM.Shining_face Is Nothing Then

            Solution_integrity = False

            Exit Sub

        End If

        For i As Integer = 0 To UBound(FEM.Shining_face)

            FEM.Shining_face(i).code = New Face_code

        Next i

        FEM.Fill_Shining_faces_colours()

        FEM.Polygon = FEM.Extract_Polygon_from_Face(FEM.Shining_face)

        Init_temperature_field()

        ' OpenGL preparation

        Dim Center As New Visualisation.Vertex

        Center.x = FEM.Center.Coord(0)

        Center.y = FEM.Center.Coord(1)

        Center.z = FEM.Center.Coord(2)

        Dim V_mode As New Visualisation.Mode

        V_mode.Status = Visualisation.Mode.WITHOUT_CONTOUR

        Dim aux_Visual As Visualisation.Visualisation

        Dim Model_list_number As Integer = 100

        aux_Visual = New Visualisation.Visualisation(V_mode, Model_list_number, FEM.Polygon, UBound(Me.FEM.Polygon) + 1, 400, 400, 16, False, Center, 2 * FEM.Maximum_length, FEM.Minimum_length, 0.1)

        ' visualisation form is ready

        ' array of HT_steps preparation

        For i As Integer = 0 To UBound(HT_element)

            ReDim HT_element(i).HT_Step(Lenght_time / Step_time)

            For j As Integer = 0 To UBound(HT_element(i).HT_Step)

                HT_element(i).HT_Step(j) = New HT_Step

            Next j

            ' Initial conditions
            HT_element(i).HT_Step(0).T = HT_element(i).Element.Field.Temperature

        Next i

        ' Heat power on first step

        'Prepare_heat_transfer(1)

        Dim END_file() As String



        ' The productivity index calculation
        Dim P_index = Calc_productivity_index(aux_Visual)

        Dim Current_step As Integer = -1

        Dim T_0, T_1 As Double

        ' connection to FTP server

        Dim login, server, pass As String

        login = ""

        server = ""

        pass = ""

        Dim hOpen As Integer

        Dim hConnection As Integer

        If FTP_mode = True Then

            ' server name, login and password extraction

            login = Mid(Work_dir, "ftp://".Length + 1, InStrRev(Work_dir, ":") - "ftp://".Length - 1)

            Try

                server = Mid(Work_dir, InStrRev(Work_dir, "@") + 1, InStr("ftp://".Length + 1, Work_dir, "/") - InStrRev(Work_dir, "@") - 1)

            Catch

                server = Mid(Work_dir, InStrRev(Work_dir, "@") + 1)

            End Try

            pass = Work_dir.Replace(login & ":", "")

            pass = pass.Replace("@" & server, "")

            pass = Mid(pass, "ftp://".Length + 1)

            Try

                pass = Mid(pass, 1, InStr(pass, "/") - 1)

            Catch


            End Try


            ' connection to Internet
            hOpen = FTP.InternetOpen("T.H.O.R.I.U.M.", FTP.INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)

            ' connection to FTP-server
            hConnection = FTP.InternetConnect(hOpen, server, FTP.INTERNET_DEFAULT_FTP_PORT, login, pass, FTP.INTERNET_SERVICE_FTP, IIf(FTP.PassiveConnection, FTP.INTERNET_FLAG_PASSIVE, 0), 0)

            ' selection of proper directory

            FTP.FtpSetDirHierarchy(hConnection, Main_form.FTP_Dir_Hierarchy)

        End If


        Do

            ' Check for solution sequence end

            If FTP_mode Then

                END_file = FTP.FtpGetCurrentDirectoryFileList(hConnection, "END")

            Else

                END_file = System.IO.Directory.GetFiles(Work_dir, "END")

            End If


            If END_file.Length <> 0 Then Exit Do

            ' It is not end now, begin calculation.

            'Current_step += 1

            ' Waiting for temperature file

            Dim T_wait_for_temperature_file As Double = 100 'milliseconds

            Dim T_wait_for_defcon_undo As Double = 60000 'milliseconds = 1 minute

            Dim T_file() As String

            T_0 = Microsoft.VisualBasic.DateAndTime.Timer

            FTP.InternetCloseHandle(hConnection)

            FTP.InternetCloseHandle(hOpen)

            Do


                ' connection to Internet
                hOpen = FTP.InternetOpen("T.H.O.R.I.U.M.", FTP.INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)

                ' connection to FTP-server
                hConnection = FTP.InternetConnect(hOpen, server, FTP.INTERNET_DEFAULT_FTP_PORT, login, pass, FTP.INTERNET_SERVICE_FTP, IIf(FTP.PassiveConnection, FTP.INTERNET_FLAG_PASSIVE, 0), 0)

                ' selection of proper directory

                FTP.FtpSetDirHierarchy(hConnection, Main_form.FTP_Dir_Hierarchy)


                System.Threading.Thread.Sleep(T_wait_for_temperature_file)

                If FTP_mode Then

                    T_file = FTP.FtpGetCurrentDirectoryFileList(hConnection, "step_*.T")

                Else

                    T_file = System.IO.Directory.GetFiles(Work_dir, "step_*.T")

                End If


                If T_file.Length = 1 Then

                    T_file(0) = Replace(T_file(0), Work_dir, "")

                    T_file(0) = Replace(T_file(0), "Step_", "")

                    Current_step = Val(Replace(T_file(0), ".T", ""))

                    Exit Do

                End If

                T_1 = Microsoft.VisualBasic.DateAndTime.Timer

                If (T_1 - T_0) * 1000 > T_wait_for_defcon_undo Then Exit Sub

                FTP.InternetCloseHandle(hConnection)

                FTP.InternetCloseHandle(hOpen)

            Loop

            ' the temperature file for current step is in directory. Read it!

            Read_temperature_file_from_dir(ID, Current_step, Work_dir, hConnection)

            If Not (Solution_integrity) Then

                Exit Sub

            End If


            ' It is need to send our ID and productivity index to server

            Dim P_index_file_name As String = ID.ToString & "_" & P_index.ToString & ".station"

            P_index_file_name = Replace(P_index_file_name, ",", ".")

            Dim writer As System.IO.StreamWriter = New System.IO.StreamWriter(tmp_Dir & P_index_file_name)

            writer.Close()


            If FTP_mode Then

                FTP.FtpPutFile(hConnection, tmp_Dir & P_index_file_name, P_index_file_name, FTP.FTP_TRANSFER_TYPE_UNKNOWN, 0)

                System.IO.File.Delete(tmp_Dir & P_index_file_name)

                If Not (FTP.FtpCurrentDirectoryFileExists(hConnection, P_index_file_name)) Then

                    Solution_integrity = False

                    FTP.InternetCloseHandle(hConnection)

                    FTP.InternetCloseHandle(hOpen)

                    Exit Sub

                End If

            Else

                System.IO.File.Move(tmp_Dir & P_index_file_name, Work_dir & P_index_file_name)

            End If

            ' waiting for elements-for-calculation file

            Dim Element_file_name() As String

            Dim T_wait_for_elements_file As Double = 500 'ms

            T_0 = Microsoft.VisualBasic.DateAndTime.Timer

            Do

                System.Threading.Thread.Sleep(T_wait_for_elements_file)

                If FTP_mode Then

                    Element_file_name = FTP.FtpGetCurrentDirectoryFileList(hConnection, ID.ToString & "*.element")

                Else

                    Element_file_name = System.IO.Directory.GetFiles(Work_dir, ID.ToString & "*.element")

                End If

                If Element_file_name.Length = 1 Then

                    Exit Do

                End If

                T_1 = Microsoft.VisualBasic.DateAndTime.Timer

                If (T_1 - T_0) * 1000 > T_wait_for_defcon_undo Then Exit Sub

            Loop

            ' We have got elements-for-calculation file name, extracting number of elements.

            Element_file_name(0) = Replace(Element_file_name(0), Work_dir, "")

            Element_file_name(0) = Replace(Element_file_name(0), ".element", "")

            Element_file_name(0) = Replace(Element_file_name(0), ID.ToString & "_", "")


            Dim N_1 As Long = Val(Mid(Element_file_name(0), 1, InStr(Element_file_name(0), "_") - 1))

            Dim N_2 As Long = Val(Replace(Element_file_name(0), N_1.ToString & "_", ""))

            ' We have got element for calculation numbers, now we should calculate radiation heat power.

            Dim Heat_power(,) As Double

            If Me.N_photon > 0 Then

                Heat_power = Calc_R_transfer_N_photon(N_1, N_2, Current_step, N_photon, aux_Visual)

            Else

                Exit Sub

                'Heat_power = Calc_R_transfer_TOP_Epsilon(N_1, N_2, Current_step, TOP_Epsilon, aux_Visual)

            End If

            ' The heat power is calculated, transfer it to server

            Write_heat_power_file_to_dir(ID, Work_dir, Heat_power, hConnection)

            If Not (Solution_integrity) Then

                FTP.InternetCloseHandle(hConnection)

                FTP.InternetCloseHandle(hOpen)

                Exit Sub

            End If

            ' Indicate progress

            RaiseEvent Solution_progress(Current_step / Me.HT_element(0).HT_Step.GetLength(0) * 100)

        Loop

    End Sub

    'Function Check_for_radiation_heat_transfer() As Boolean

    '    For i As Integer = 0 To UBound(B_matrix, 1)

    '        For j As Integer = 0 To UBound(B_matrix, 2)

    '            If B_matrix(i, j) <> 0 Then

    '                Return True

    '            End If


    '        Next j

    '    Next i

    '    Return False

    'End Function


    ''' <summary>
    ''' This sub solve problem using 4 step Adams–Bashforth method, taking in consideration elemet properties NOT changed at time of solving
    ''' </summary>
    ''' <remarks></remarks>
    Sub Adams_4_const()

        Dim Progress As Integer = 0

        Dim Mass, c As Double

        ' правые части уравнений (Q_n/(c_n*m))
        ' текущего шага по времени, предыдущего шага, предпредыдущего и т.д.

        Dim F_n_1, F_n_2, F_n_3, F_n_4 As Double

        ' Начинаем отчет времени

        Current_time = Start_time


        ' объявляем массивы шагов по времени

        For i As Integer = 0 To UBound(HT_element)

            ReDim HT_element(i).HT_Step(Lenght_time / Step_time)

            For j As Integer = 0 To UBound(HT_element(i).HT_Step)

                HT_element(i).HT_Step(j) = New HT_Step

            Next j

            ' заполняем начальные условия
            HT_element(i).HT_Step(0).T = HT_element(i).Element.Field.Temperature

            'Fill_HT_element_argument(HT_element(i))

            'Calc_Q(HT_element(i))

            'HT_element(i).HT_Step(0).Q = HT_element(i).Q_summary

        Next i

        Dim N_steps = Lenght_time / Step_time


        ' цикл по шагам времени

        If N_steps >= 1 Then

            Dim n As Integer = 1

            ' рассчитываем тепловой поток на текущем шаге
            For i As Integer = 0 To UBound(HT_element)

                Fill_HT_element_argument(HT_element(i))

                Calc_Q_const(HT_element(i))

                HT_element(i).HT_Step(n).Q = HT_element(i).Q_summary

                HT_element(i).HT_Step(n).Time = Start_time + Step_time

            Next i

            ' циклы по элементам
            ' рассчитываем температуру на текущем шаге
            For i As Integer = 0 To UBound(HT_element)

                Fill_HT_element_argument(HT_element(i))

                Mass = HT_element(i).Element.Field.Mass

                c = HT_element(i).Element.Medium.Calc_c(HT_element(i).Argument)

                F_n_1 = HT_element(i).HT_Step(n).Q / (c * Mass)

                HT_element(i).HT_Step(n).T = HT_element(i).HT_Step(n - 1).T + Step_time * F_n_1

                HT_element(i).Element.Field.Temperature = HT_element(i).HT_Step(n).T

            Next i



        End If

        For n As Integer = 1 To N_steps

            Current_time += Step_time



            ' на втором шаге - метод Адамса второго порядка
            If n = 2 Then

                ' циклы по элементам
                ' рассчитываем температуру на текущем шаге
                For i As Integer = 0 To UBound(HT_element)

                    Fill_HT_element_argument(HT_element(i))

                    Mass = HT_element(i).Element.Field.Mass

                    c = HT_element(i).Element.Medium.Calc_c(HT_element(i).Argument)

                    F_n_1 = HT_element(i).HT_Step(1).Q / (c * Mass)

                    c = HT_element(i).Element.Medium.Calc_c(HT_element(i).Argument)

                    F_n_2 = HT_element(i).HT_Step(0).Q / (c * Mass)



                    HT_element(i).HT_Step(2).T = HT_element(i).HT_Step(1).T + Step_time * (1.5 * F_n_1 - 0.5 * F_n_2)

                    HT_element(i).Element.Field.Temperature = HT_element(i).HT_Step(2).T

                Next i

                ' рассчитываем тепловой поток на текущем шаге
                For i As Integer = 0 To UBound(HT_element)

                    Fill_HT_element_argument(HT_element(i))

                    Calc_Q_const(HT_element(i))

                    HT_element(i).HT_Step(2).Q = HT_element(i).Q_summary

                    HT_element(i).HT_Step(2).Time = Current_time

                Next i


            End If

            ' на третьем шаге - метод Адамса третьего порядка
            If n = 3 Then

                ' циклы по элементам
                ' рассчитываем температуру на текущем шаге
                For i As Integer = 0 To UBound(HT_element)

                    Fill_HT_element_argument(HT_element(i))

                    Mass = HT_element(i).Element.Field.Mass

                    c = HT_element(i).Element.Medium.Calc_c(HT_element(i).Argument)

                    F_n_1 = HT_element(i).HT_Step(2).Q / (c * Mass)

                    c = HT_element(i).Element.Medium.Calc_c(HT_element(i).Argument)

                    F_n_2 = HT_element(i).HT_Step(1).Q / (c * Mass)

                    c = HT_element(i).Element.Medium.Calc_c(HT_element(i).Argument)

                    F_n_3 = HT_element(i).HT_Step(0).Q / (c * Mass)



                    HT_element(i).HT_Step(3).T = HT_element(i).HT_Step(2).T + Step_time * (23 / 12 * F_n_1 - 4 / 3 * F_n_2 + 5 / 12 * F_n_3)

                    HT_element(i).Element.Field.Temperature = HT_element(i).HT_Step(3).T

                Next i

                ' рассчитываем тепловой поток на текущем шаге
                For i As Integer = 0 To UBound(HT_element)

                    Fill_HT_element_argument(HT_element(i))

                    Calc_Q_const(HT_element(i))

                    HT_element(i).HT_Step(3).Q = HT_element(i).Q_summary

                    HT_element(i).HT_Step(3).Time = Current_time

                Next i

            End If

            ' на всех последующих шагах - метод Адамса четвертого порядка
            If n >= 4 Then

                ' циклы по элементам
                ' рассчитываем температуру на текущем шаге
                For i As Integer = 0 To UBound(HT_element)

                    Fill_HT_element_argument(HT_element(i))

                    Mass = HT_element(i).Element.Field.Mass

                    c = HT_element(i).Element.Medium.Calc_c(HT_element(i).Argument)

                    F_n_1 = HT_element(i).HT_Step(n - 1).Q / (c * Mass)

                    c = HT_element(i).Element.Medium.Calc_c(HT_element(i).Argument)

                    F_n_2 = HT_element(i).HT_Step(n - 2).Q / (c * Mass)

                    c = HT_element(i).Element.Medium.Calc_c(HT_element(i).Argument)

                    F_n_3 = HT_element(i).HT_Step(n - 3).Q / (c * Mass)

                    c = HT_element(i).Element.Medium.Calc_c(HT_element(i).Argument)

                    F_n_4 = HT_element(i).HT_Step(n - 4).Q / (c * Mass)



                    HT_element(i).HT_Step(n).T = HT_element(i).HT_Step(n - 1).T + Step_time * (55 / 24 * F_n_1 - 59 / 24 * F_n_2 + 37 / 24 * F_n_3 - 3 / 8 * F_n_4)

                    HT_element(i).Element.Field.Temperature = HT_element(i).HT_Step(n).T

                Next i

                ' рассчитываем тепловой поток на текущем шаге
                For i As Integer = 0 To UBound(HT_element)

                    Fill_HT_element_argument(HT_element(i))

                    Calc_Q_const(HT_element(i))

                    HT_element(i).HT_Step(n).Q = HT_element(i).Q_summary

                    HT_element(i).HT_Step(n).Time = Current_time

                Next i

            End If

            ' явно задаем температуру из T_SET

            '  для этого задаемся текущим временем 

            Dim Argument As Vector = New Vector

            ReDim Argument.Coord(5)

            Argument.Coord(1) = Current_time

            Dim Number As Integer

            For i As Integer = 0 To UBound(FEM.T_SET)

                For j As Integer = 0 To UBound(FEM.T_SET(i).Element)

                    Number = FEM.T_SET(i).Element(j).HT_element.Number

                    If FEM.T_SET(i).Temp(j).GetType.ToString = "System.Double" Then

                        HT_element(Number).HT_Step(n).T = FEM.T_SET(i).Temp(j)

                    Else

                        HT_element(Number).HT_Step(n).T = FEM.T_SET(i).Temp(j).Return_value(Argument)

                    End If

                Next j

            Next i


            If Progress + 1 <= n / (Lenght_time / Step_time) * 100 Then

                Progress += 1

                Main_form.Save_model()

                RaiseEvent Solution_progress(Progress)

            End If

            Application.DoEvents()

        Next n

        Main_form.Save_model()
    End Sub

    Sub Adams_4_var()

        Dim Mass As Double

        ' правые части уравнений (Q_n/(c_n*m))
        ' текущего шага по времени, предыдущего шага, предпредыдущего и т.д.

        Dim F_n_1, F_n_2, F_n_3, F_n_4, F As Double

        ' Начинаем отчет времени

        Current_time = Start_time


        ' объявляем массивы шагов по времени

        For i As Integer = 0 To UBound(HT_element)

            ReDim HT_element(i).HT_Step(Lenght_time / Step_time)

            For j As Integer = 0 To UBound(HT_element(i).HT_Step)

                HT_element(i).HT_Step(j) = New HT_Step

            Next j

            ' заполняем начальные условия
            HT_element(i).HT_Step(0).T = HT_element(i).Element.Field.Temperature

            'Fill_HT_element_argument(HT_element(i))

            'Calc_Q(HT_element(i))

            'HT_element(i).HT_Step(0).Q = HT_element(i).Q_summary

        Next i

        Dim N_steps = Lenght_time / Step_time

        ' цикл по шагам времени

        If N_steps >= 1 Then

            Current_time += Step_time

            Dim n As Integer = 1

            ' рассчитываем тепловой поток на текущем шаге
            For i As Integer = 0 To UBound(HT_element)

                Fill_HT_element_argument(HT_element(i))

                Calc_Q_const(HT_element(i))

                HT_element(i).HT_Step(n).Q = HT_element(i).Q_summary

                HT_element(i).HT_Step(n).Time = Current_time

            Next i

            ' циклы по элементам
            ' рассчитываем температуру на текущем шаге
            For i As Integer = 0 To UBound(HT_element)

                Fill_HT_element_argument(HT_element(i))

                Mass = HT_element(i).Element.Field.Mass

                F_n_1 = HT_element(i).HT_Step(n).Q * Step_time / Mass

                F = F_n_1

                HT_element(i).HT_Step(n).T = Calc_T(HT_element(i), HT_element(i).HT_Step(n - 1).T, HT_element(i).HT_Step(n - 1).T, F)

                HT_element(i).Element.Field.Temperature = HT_element(i).HT_Step(n).T

            Next i

        End If

        ' На втором шаге - метод Адамса второго порядка
        If N_steps >= 2 Then

            Current_time += Step_time

            Dim n As Integer = 2

            ' рассчитываем тепловой поток на текущем шаге
            For i As Integer = 0 To UBound(HT_element)

                Fill_HT_element_argument(HT_element(i))

                Calc_Q_const(HT_element(i))

                HT_element(i).HT_Step(n).Q = HT_element(i).Q_summary

            Next i

            ' рассчитываем температуру на текущем шаге
            For i As Integer = 0 To UBound(HT_element)

                Fill_HT_element_argument(HT_element(i))

                Mass = HT_element(i).Element.Field.Mass

                F_n_1 = 1.5 * HT_element(i).HT_Step(n - 1).Q

                F_n_2 = -0.5 * HT_element(i).HT_Step(n - 2).Q

                F = (F_n_1 + F_n_2) * Step_time / Mass

                HT_element(i).HT_Step(n).T = Calc_T(HT_element(i), HT_element(i).HT_Step(n - 1).T, HT_element(i).HT_Step(n - 2).T, F)

                HT_element(i).Element.Field.Temperature = HT_element(i).HT_Step(n).T

            Next i

        End If

        ' На третьем шаге - метод Адамса третьего порядка
        If N_steps >= 3 Then

            Current_time += Step_time

            Dim n As Integer = 3

            ' рассчитываем тепловой поток на текущем шаге
            For i As Integer = 0 To UBound(HT_element)

                Fill_HT_element_argument(HT_element(i))

                Calc_Q_const(HT_element(i))

                HT_element(i).HT_Step(n).Q = HT_element(i).Q_summary

            Next i

            ' рассчитываем температуру на текущем шаге
            For i As Integer = 0 To UBound(HT_element)

                Fill_HT_element_argument(HT_element(i))

                Mass = HT_element(i).Element.Field.Mass

                F_n_1 = 23 / 12 * HT_element(i).HT_Step(n).Q

                F_n_2 = -16 / 12 * HT_element(i).HT_Step(n - 1).Q

                F_n_3 = 5 / 12 * HT_element(i).HT_Step(n - 2).Q

                F = (F_n_1 + F_n_2 + F_n_3) * Step_time / Mass

                HT_element(i).HT_Step(n).T = Calc_T(HT_element(i), HT_element(i).HT_Step(n - 1).T, HT_element(i).HT_Step(n - 2).T, F)

                HT_element(i).Element.Field.Temperature = HT_element(i).HT_Step(n).T

            Next i

        End If


        If N_steps >= 4 Then

            For n As Integer = 4 To N_steps

                Current_time += Step_time

                ' рассчитываем тепловой поток на текущем шаге
                For i As Integer = 0 To UBound(HT_element)

                    Fill_HT_element_argument(HT_element(i))

                    Calc_Q_const(HT_element(i))

                    HT_element(i).HT_Step(n).Q = HT_element(i).Q_summary

                Next i

                ' рассчитываем температуру на текущем шаге
                For i As Integer = 0 To UBound(HT_element)

                    Fill_HT_element_argument(HT_element(i))

                    Mass = HT_element(i).Element.Field.Mass

                    F_n_1 = 55 / 24 * HT_element(i).HT_Step(n).Q

                    F_n_2 = -59 / 24 * HT_element(i).HT_Step(n - 1).Q

                    F_n_3 = 37 / 24 * HT_element(i).HT_Step(n - 2).Q

                    F_n_4 = -9 / 24 * HT_element(i).HT_Step(n - 3).Q

                    F = (F_n_1 + F_n_2 + F_n_3 + F_n_4) * Step_time / Mass

                    HT_element(i).HT_Step(n).T = Calc_T(HT_element(i), HT_element(i).HT_Step(n - 1).T, HT_element(i).HT_Step(n - 2).T, F)

                    HT_element(i).Element.Field.Temperature = HT_element(i).HT_Step(n).T

                Next i

            Next n

        End If

    End Sub
    ''' <summary>
    ''' This function calcutates temperature for next iteration 
    ''' </summary>
    ''' <param name="HT_element"></param>
    ''' <param name="T_i">Current element temperature</param>
    ''' <param name="T_minus_i">Element temperature from previous step</param>
    ''' <param name="F">Right part of equation</param>
    ''' <returns>
    ''' </returns>
    ''' <remarks></remarks>
    Function Calc_T(ByRef HT_element As HT_Element, ByRef T_i As Double, ByRef T_minus_i As Double, ByRef F As Double) As Double


        'Dim writer_1 As System.IO.StreamWriter = New System.IO.StreamWriter("C:\tmp\T_c_lambda_rod.txt", False)

        'For T_ As Double = 290 To 310 Step 1

        '    HT_element.Argument.Coord(0) = T_

        '    writer_1.WriteLine(Str(T_) & " " & Str(HT_element.Element.Medium.Calc_c(HT_element.Argument)) & " " & Str(HT_element.Element.Medium.Calc_lambda(HT_element.Argument)))

        'Next T_

        'writer_1.Close()

        'writer_1 = Nothing

        'Dim A_1, A_0, cT, K, T_i_plus_1, eps As Double

        'Const delta_T = 0.01

        'HT_element.Argument.Coord(0) = T_i

        'cT = HT_element.Element.Medium.Calc_c(HT_element.Argument) * (T_i)

        '' Zero incoming power check

        'If F = 0 Then Return T_i

        '' Zero approximation

        'T_i_plus_1 = (2 + delta_T) * T_i - T_minus_i

        'Do

        '    HT_element.Argument.Coord(0) = T_i_plus_1

        '    A_0 = HT_element.Element.Medium.Calc_c(HT_element.Argument) * (T_i_plus_1)

        '    eps = Math.Abs((A_0 - F - cT) / F)

        '    If eps <= Calc_T_Epsilon Then

        '        Exit Do

        '    End If

        '    HT_element.Argument.Coord(0) = T_i_plus_1 + delta_T * T_i

        '    A_1 = HT_element.Element.Medium.Calc_c(HT_element.Argument) * (T_i_plus_1 + delta_T * T_i)

        '    K = (A_1 - A_0) / (delta_T * T_i)

        '    T_i_plus_1 = (F + cT - A_0) / K + T_i_plus_1

        'Loop

        'Return T_i_plus_1

        If F = 0 Then Return T_i

        Dim T_next, T, delta_T, cT_i, cT, cT_plus As Double

        delta_T = 0.0001 * T_i

        HT_element.Argument.Coord(0) = T_i

        cT_i = HT_element.Element.Medium.Calc_c(HT_element.Argument) * (T_i)

        T = T_i

        Do
            HT_element.Argument.Coord(0) = T

            cT = HT_element.Element.Medium.Calc_c(HT_element.Argument) * T

            HT_element.Argument.Coord(0) = T + delta_T

            cT_plus = HT_element.Element.Medium.Calc_c(HT_element.Argument) * (T + delta_T)

            T_next = T - (cT - (cT_i + F)) * delta_T / (cT_plus - cT)

            If Math.Abs(T_next - T) < Math.Abs(delta_T) Then Return T_next

            T = T_next

        Loop

    End Function


    '''' <summary>
    '''' This sub solve problem using simple Euler integration method
    '''' </summary>
    '''' <remarks></remarks>
    'Sub Euler_var()

    '    Dim Mass As Double

    '    ' правые части уравнений (Q_n/(c_n*m))
    '    ' текущего шага по времени, предыдущего шага, предпредыдущего и т.д.

    '    Dim F_n_1 As Double

    '    ' Начинаем отчет времени

    '    Current_time = Start_time


    '    ' объявляем массивы шагов по времени

    '    For i As Integer = 0 To UBound(HT_element)

    '        ReDim HT_element(i).HT_Step(Lenght_time / Step_time)

    '        For j As Integer = 0 To UBound(HT_element(i).HT_Step)

    '            HT_element(i).HT_Step(j) = New HT_Step

    '        Next j

    '        ' заполняем начальные условия
    '        HT_element(i).HT_Step(0).T = HT_element(i).Element.Field.Temperature

    '    Next i


    '    ' цикл по шагам времени
    '    For n As Integer = 1 To Lenght_time / Step_time

    '        Current_time += Step_time


    '        ' рассчитываем тепловой поток на текущем шаге
    '        For i As Integer = 0 To UBound(HT_element)

    '            Calc_Q(HT_element(i))

    '            HT_element(i).HT_Step(n).Q = HT_element(i).Q_summary

    '        Next i

    '        ' циклы по элементам
    '        ' рассчитываем температуру на текущем шаге
    '        For i As Integer = 0 To UBound(HT_element)

    '            Fill_HT_element_argument(HT_element(i))

    '            Mass = HT_element(i).Element.Field.Mass

    '            F_n_1 = HT_element(i).HT_Step(n).Q * Step_time / Mass

    '            HT_element(i).HT_Step(n).T = Calc_T(HT_element(i), HT_element(i).HT_Step(n - 1).T, HT_element(i).HT_Step(n - 1).T, F_n_1, TOP_Epsilon)

    '            HT_element(i).Element.Field.Temperature = HT_element(i).HT_Step(n).T

    '            If Double.IsInfinity(HT_element(i).Element.Field.Temperature) Or Double.IsNaN(HT_element(i).Element.Field.Temperature) Then

    '                Throw New System.Exception("Solution is instabile. Step number " & n.ToString)

    '                Exit Sub

    '            End If

    '        Next i

    '    Next n

    'End Sub

    ''' <summary>
    ''' This sub solve problem using simple Euler integration method
    ''' </summary>
    ''' <remarks></remarks>
    Sub Euler_const()

        ' правые части уравнений (Q_n/(c_n*m))
        ' текущего шага по времени, предыдущего шага, предпредыдущего и т.д.

        ' Начинаем отчет времени

        Current_time = Start_time

        Dim Progress As Integer = 0


        ' объявляем массивы шагов по времени

        For i As Integer = 0 To UBound(HT_element)

            ReDim HT_element(i).HT_Step(Lenght_time / Step_time)

            For j As Integer = 0 To UBound(HT_element(i).HT_Step)

                HT_element(i).HT_Step(j) = New HT_Step

            Next j

            ' заполняем начальные условия
            HT_element(i).HT_Step(0).T = HT_element(i).Element.Field.Temperature

        Next i


        ' цикл по шагам времени
        For n As Integer = 1 To Lenght_time / Step_time

            Current_time += Step_time

            ' явно задаем температуру из T_SET

            '  для этого задаемся текущим временем 

            Dim Argument As Vector = New Vector

            ReDim Argument.Coord(5)

            Argument.Coord(1) = Current_time

            Dim Number As Integer

            For i As Integer = 0 To UBound(FEM.T_SET)

                For j As Integer = 0 To UBound(FEM.T_SET(i).Element)

                    Number = FEM.T_SET(i).Element(j).HT_element.Number

                    If FEM.T_SET(i).Temp(j).GetType.ToString = "System.Double" Then

                        HT_element(Number).HT_Step(n).T = FEM.T_SET(i).Temp(j)

                    Else

                        HT_element(Number).HT_Step(n).T = FEM.T_SET(i).Temp(j).Return_value(Argument).Coord(0)

                    End If

                    HT_element(Number).Element.Field.Temperature = HT_element(Number).HT_Step(n).T

                Next j

            Next i



            ' рассчитываем тепловой поток на текущем шаге, проверяя не задана ли температура для этого элемента 
            For i As Integer = 0 To UBound(HT_element)

                If HT_element(i).T_SET Is Nothing Then

                    Fill_HT_element_argument(HT_element(i))

                    Calc_Q_const(HT_element(i))

                    HT_element(i).HT_Step(n).Q = HT_element(i).Q_summary

                End If

                HT_element(i).HT_Step(n).Time = Current_time

            Next i

            ' циклы по элементам
            ' рассчитываем температуру на текущем шаге, проверяя не задана ли уже температура для этого элемента 
            For i As Integer = 0 To UBound(HT_element)

                If HT_element(i).T_SET Is Nothing Then

                    With HT_element(i).Element

                        'HT_element(i).HT_Step(n).T = HT_element(i).HT_Step(n - 1).T + Step_time * HT_element(i).HT_Step(n).Q / (.Medium.c * .Field.Mass)

                        'HT_element(i).HT_Step(n).T = HT_element(i).HT_Step(n - 1).T + Step_time * HT_element(i).Q_summary / (.Medium.c * .Field.Mass)

                        HT_element(i).HT_Step(n).T = Calc_T(HT_element(i), HT_element(i).HT_Step(n - 1).T, HT_element(i).HT_Step(n - 1).T, HT_element(i).Q_summary * Step_time / .Field.Mass)

                        .Field.Temperature = HT_element(i).HT_Step(n).T

                        If Double.IsInfinity(.Field.Temperature) Or Double.IsNaN(.Field.Temperature) Then

                            Throw New System.Exception("Solution is instabile. Step number " & n.ToString)

                            Exit Sub

                        End If

                    End With

                End If

            Next i

            If Progress + 1 <= n / (Lenght_time / Step_time) * 100 Then

                Progress += 1

                RaiseEvent Solution_progress(Progress)

                Main_form.Save_model()

            End If

            Application.DoEvents()

        Next n

    End Sub




    '''' <summary>
    '''' This sub solve problem using simple Euler integration method.
    '''' Radiation heat transfer matrix is not used for this method.
    '''' </summary>
    '''' <param name="N_step_start">Step to begin analysis. Can be skipped to set default step =1.
    '''' If N_step_start is set, the analysis begins from this step.</param>
    '''' <remarks></remarks>
    'Sub Euler_const_MATRIXLESS(Optional ByRef N_step_start As Integer = 1)

    '    Dim T_0, T_1 As Double

    '    Dim start, finish As Double

    '    start = Microsoft.VisualBasic.DateAndTime.Timer

    '    T_0 = Microsoft.VisualBasic.DateAndTime.Timer

    '    ' правые части уравнений (Q_n/(c_n*m))
    '    ' текущего шага по времени, предыдущего шага, предпредыдущего и т.д.

    '    ' Начинаем отчет времени

    '    Current_time = Start_time

    '    Dim Progress As Integer = 0

    '    Dim Old_Progress As Integer = -1

    '    ' создаем визуализационную форму для расчета теплопередачи

    '    ' подготавливаем модель

    '    For i As Integer = 0 To UBound(FEM.Shining_face)

    '        FEM.Shining_face(i).code = New Face_code

    '    Next i

    '    FEM.Fill_Shining_faces_colours()

    '    FEM.Polygon = FEM.Extract_Polygon_from_Face(FEM.Shining_face)

    '    Init_temperature_field()

    '    ' подготавливаем OpenGL для работы

    '    Dim Center As New Visualisation.Vertex

    '    Center.x = FEM.Center.Coord(0)

    '    Center.y = FEM.Center.Coord(1)

    '    Center.z = FEM.Center.Coord(2)

    '    Dim V_mode As New Visualisation.Mode

    '    V_mode.Status = Visualisation.Mode.WITHOUT_CONTOUR

    '    Dim aux_Visual As Visualisation.Visualisation

    '    Dim Model_list_number As Integer = 100

    '    aux_Visual = New Visualisation.Visualisation(V_mode, Model_list_number, FEM.Polygon, UBound(Me.FEM.Polygon) + 1, 400, 400, 16, False, Center, 2 * FEM.Maximum_length, FEM.Minimum_length, 0.1)

    '    ' визуализационная форма для работы готова

    '    ' объявляем массивы шагов по времени, если нужно считать с первого шага

    '    If N_step_start = 1 Then

    '        Init_temperature_field()

    '        For i As Integer = 0 To UBound(HT_element)

    '            ReDim HT_element(i).HT_Step(0)

    '            ' заполняем начальные условия
    '            HT_element(i).HT_Step(0).T = HT_element(i).Element.Field.Temperature

    '            HT_element(i).HT_Step(0).Time = Start_time

    '        Next i

    '    End If

    '    ' Рассчитываем тепловые потоки на первом шаге

    '    Prepare_heat_transfer(N_step_start)

    '    Calc_C_transfer(N_step_start - 1)

    '    Calc_R_transfer_N_photon(N_step_start - 1, Me.N_photon, aux_Visual)


    '    ' цикл по шагам времени
    '    For n As Integer = 1 To Lenght_time / Step_time

    '        Current_time += Step_time

    '        ' явно задаем температуру из T_SET

    '        '  для этого задаемся текущим временем 

    '        Dim Argument As Vector = New Vector

    '        ReDim Argument.Coord(5)

    '        Argument.Coord(1) = Current_time

    '        Dim Number As Integer

    '        For i As Integer = 0 To UBound(FEM.T_SET)

    '            For j As Integer = 0 To UBound(FEM.T_SET(i).Element)

    '                Number = FEM.T_SET(i).Element(j).HT_element.Number

    '                If FEM.T_SET(i).Temp(j).GetType.ToString = "System.Double" Then

    '                    HT_element(Number).HT_Step(n).T = FEM.T_SET(i).Temp(j)

    '                Else

    '                    HT_element(Number).HT_Step(n).T = FEM.T_SET(i).Temp(j).Return_value(Argument)

    '                End If

    '                HT_element(Number).Element.Field.Temperature = HT_element(i).HT_Step(n).T

    '            Next j

    '        Next i



    '        ' циклы по элементам
    '        ' рассчитываем температуру на текущем шаге
    '        For i As Integer = 0 To UBound(HT_element)

    '            If HT_element(i).HT_Step(n).T = 0 Then

    '                HT_element(i).HT_Step(n).Time = Current_time

    '                With HT_element(i).Element

    '                    HT_element(i).HT_Step(n).T = HT_element(i).HT_Step(n - 1).T + Step_time * HT_element(i).HT_Step(n).Q / (.Medium.Calc_c(HT_element(i).Argument) * .Field.Mass)
    '                    .Field.Temperature = HT_element(i).HT_Step(n).T

    '                    If Double.IsInfinity(.Field.Temperature) Or Double.IsNaN(.Field.Temperature) Then

    '                        Throw New System.Exception("Solution is instabile. Step number " & n.ToString)

    '                        Exit Sub

    '                    End If

    '                End With

    '            End If

    '        Next i


    '        If n >= Lenght_time / Step_time Then Exit For

    '        ' рассчитываем тепловой поток на следующем шаге шаге

    '        Prepare_heat_transfer(n + 1)

    '        Calc_C_transfer(n)

    '        If Me.N_photon > 0 Then

    '            Me.Calc_R_transfer_N_photon(n, Me.N_photon, aux_Visual)

    '        Else

    '            Exit Sub

    '            'Me.Calc_R_transfer_TOP_Epsilon(n, Me.TOP_Epsilon, aux_Visual)

    '        End If

    '        ' отображаем ход решения

    '        T_1 = Microsoft.VisualBasic.DateAndTime.Timer

    '        If T_1 - T_0 > 0.25 Then

    '            Progress = (Current_time - Start_time) / Lenght_time * 100

    '            RaiseEvent Solution_progress(Progress)

    '            If Progress > Old_Progress Then

    '                Main_form.Save_model()

    '            End If

    '        End If

    '        Application.DoEvents()

    '        T_0 = Microsoft.VisualBasic.DateAndTime.Timer

    '    Next n

    '    finish = Microsoft.VisualBasic.DateAndTime.Timer

    'End Sub

    ''' <summary>
    ''' This sub solve problem using simple Euler integration method with variable time step.
    ''' Radiation heat transfer matrix is not used for this method.
    ''' </summary>
    ''' <param name="N_step_start">Step to begin analysis. Can be skipped to set default step =1.
    ''' If N_step_start is set, the analysis begins from this step.</param>
    ''' <remarks></remarks>
    Sub Euler_var_T_step_with_smoothing_MATRIXLESS(Optional ByRef N_step_start As Integer = 1)

        Dim delta_t_min(UBound(HT_element)) As Double

        Dim delta_t_max(UBound(HT_element)) As Double

        Dim max_delta_t As Double

        Dim T_0, T_1 As Double

        Dim start, finish As Double

        start = Microsoft.VisualBasic.DateAndTime.Timer

        T_0 = Microsoft.VisualBasic.DateAndTime.Timer

        ' правые части уравнений (Q_n/(c_n*m))
        ' текущего шага по времени, предыдущего шага, предпредыдущего и т.д.

        ' Начинаем отчет времени

        Current_time = Start_time

        Dim Progress As Integer = 0

        Dim Old_Progress As Integer = -1

        ' создаем визуализационную форму для расчета теплопередачи

        ' подготавливаем модель

        For i As Integer = 0 To UBound(FEM.Shining_face)

            FEM.Shining_face(i).code = New Face_code

        Next i

        FEM.Fill_Shining_faces_colours()

        FEM.Polygon = FEM.Extract_Polygon_from_Face(FEM.Shining_face)

        ' подготавливаем OpenGL для работы

        Dim Center As New Visualisation.Vertex

        Center.x = FEM.Center.Coord(0)

        Center.y = FEM.Center.Coord(1)

        Center.z = FEM.Center.Coord(2)

        Dim V_mode As New Visualisation.Mode

        V_mode.Status = Visualisation.Mode.WITHOUT_CONTOUR

        Dim aux_Visual As Visualisation.Visualisation

        Dim Model_list_number As Integer = 100

        aux_Visual = New Visualisation.Visualisation(V_mode, Model_list_number, FEM.Polygon, UBound(Me.FEM.Polygon) + 1, 400, 400, 16, False, Center, 2 * FEM.Maximum_length, FEM.Minimum_length, 0.1)

        ' визуализационная форма для работы готова

        ' объявляем массивы шагов по времени, если нужно считать с первого шага

        If N_step_start = 1 Then

            Init_temperature_field()

            For i As Integer = 0 To UBound(HT_element)

                ReDim HT_element(i).HT_Step(0)

                ' заполняем начальные условия
                HT_element(i).HT_Step(0).T = HT_element(i).Element.Field.Temperature

                HT_element(i).HT_Step(0).Time = Start_time

            Next i

        End If

        ' Рассчитываем тепловые потоки на первом шаге

        'HT_element(1).HT_Step(1).Q = 0

        Prepare_heat_transfer(N_step_start)

        Calc_C_transfer(N_step_start - 1)

        Calc_R_transfer_N_photon(N_step_start - 1, Me.N_photon, aux_Visual)


        ' номер шага по времени
        Dim n As Integer = N_step_start

        ' массив, в который кладем сглаженные мощности всех элементов на шаге n
        Dim Q_smoothed(UBound(HT_element)) As Double

        Dim Q_in_to_smoothing() As Double

        Dim t_in_to_smoothing() As Double


        ReDim Q_in_to_smoothing(N_steps - 1)

        ReDim t_in_to_smoothing(N_steps - 1)


        Dim N_for_initialization As Integer = 4

        Dim N_order_for_smoothing As Integer = 3

        Do

            ' цикл по элементам

            ' проводим сглаживание графиков мощностей на элемент.

            If Smoothing_type = "LINEAR" And N_steps > 1 Then
                ' сглаживание проводим линейное. Коэффициенты подбираем методом наименьших квадратов

                For i As Integer = 0 To UBound(HT_element)

                    If n >= N_steps Then

                        For j As Integer = 0 To UBound(t_in_to_smoothing)

                            Q_in_to_smoothing(j) = HT_element(i).HT_Step(n - N_steps + j + 1).Q


                            t_in_to_smoothing(j) = HT_element(i).HT_Step(n - N_steps + j).Time

                        Next j

                        Q_smoothed(i) = Linear_smoothing(HT_element(i).HT_Step(n - 1).Time, t_in_to_smoothing, Q_in_to_smoothing)

                    Else

                        Q_smoothed(i) = HT_element(i).HT_Step(n).Q

                    End If

                Next i

            End If

            If Smoothing_type = "" Then

                For i As Integer = 0 To UBound(HT_element)

                    Q_smoothed(i) = HT_element(i).HT_Step(n).Q

                Next i

            End If


            ' расчитываем максимальный возможный шаг по времени

            max_delta_t = Double.MaxValue

            For i As Integer = 0 To UBound(HT_element)

                HT_element(i).Argument.Coord(0) = HT_element(i).HT_Step(n - 1).T

                HT_element(i).Argument.Coord(1) = Current_time

                With HT_element(i).Element

                    delta_t_min(i) = (K_T_min - 1) * HT_element(i).HT_Step(n - 1).T * .Medium.Calc_c(HT_element(i).Argument) * .Field.Mass / Q_smoothed(i)

                    delta_t_max(i) = (K_T_max - 1) * HT_element(i).HT_Step(n - 1).T * .Medium.Calc_c(HT_element(i).Argument) * .Field.Mass / Q_smoothed(i)

                    If delta_t_min(i) < max_delta_t And delta_t_min(i) > 0 Then

                        max_delta_t = delta_t_min(i)

                    End If

                    If delta_t_max(i) < max_delta_t And delta_t_max(i) > 0 Then

                        max_delta_t = delta_t_max(i)

                    End If

                End With

            Next i

            If Lenght_time + Start_time - Current_time < max_delta_t Then

                max_delta_t = Lenght_time + Start_time - Current_time

            End If

            Step_time = max_delta_t

            Current_time += Step_time

            If n = 1 Then

                ' циклы по элементам
                ' рассчитываем температуру на текущем шаге
                For i As Integer = 0 To UBound(HT_element)

                    HT_element(i).HT_Step(n).Time = Current_time

                    If HT_element(i).T_SET Is Nothing Then

                        With HT_element(i).Element

                            'HT_element(i).HT_Step(n).T = HT_element(i).HT_Step(n - 1).T + Step_time * Q_smoothed(i) / (.Medium.Calc_c(HT_element(i).Argument) * .Field.Mass)
                            '.Field.Temperature = HT_element(i).HT_Step(n).T

                            HT_element(i).HT_Step(n).T = Calc_T(HT_element(i), HT_element(i).HT_Step(n - 1).T, HT_element(i).HT_Step(n - 1).T, Q_smoothed(i) * Step_time / .Field.Mass)

                            If Double.IsInfinity(.Field.Temperature) Or Double.IsNaN(.Field.Temperature) Then

                                Throw New System.Exception("Solution is instabile. Step number " & n.ToString)

                                Exit Sub

                            End If

                        End With

                    End If

                Next i


            Else

                ' циклы по элементам
                ' рассчитываем температуру на текущем шаге
                For i As Integer = 0 To UBound(HT_element)

                    HT_element(i).HT_Step(n).Time = Current_time

                    If HT_element(i).T_SET Is Nothing Then

                        With HT_element(i).Element

                            'HT_element(i).HT_Step(n).T = HT_element(i).HT_Step(n - 1).T + Step_time * Q_smoothed(i) / (.Medium.Calc_c(HT_element(i).Argument) * .Field.Mass)
                            '.Field.Temperature = HT_element(i).HT_Step(n).T

                            HT_element(i).HT_Step(n).T = Calc_T(HT_element(i), HT_element(i).HT_Step(n - 1).T, HT_element(i).HT_Step(n - 2).T, Q_smoothed(i) * Step_time / .Field.Mass)

                            If Double.IsInfinity(.Field.Temperature) Or Double.IsNaN(.Field.Temperature) Then

                                Throw New System.Exception("Solution is instabile. Step number " & n.ToString)

                                Exit Sub

                            End If

                        End With

                    End If

                Next i


            End If


            ' явно задаем температуру из T_SET

            '  для этого задаемся текущим временем 

            Dim Argument As Vector = New Vector

            ReDim Argument.Coord(5)

            Argument.Coord(1) = Current_time

            Dim Number As Integer

            For i As Integer = 0 To UBound(FEM.T_SET)

                For j As Integer = 0 To UBound(FEM.T_SET(i).Element)

                    Number = FEM.T_SET(i).Element(j).HT_element.Number

                    If FEM.T_SET(i).Temp(j).GetType.ToString = "System.Double" Then

                        HT_element(Number).HT_Step(n).T = FEM.T_SET(i).Temp(j)

                    Else

                        HT_element(Number).HT_Step(n).T = FEM.T_SET(i).Temp(j).Return_value(Argument)

                    End If

                Next j

            Next i

            If Current_time >= Start_time + Lenght_time Then Exit Do

            ' рассчитываем тепловой поток на следующем шаге шаге

            Prepare_heat_transfer(n + 1)

            Calc_C_transfer(n)

            Calc_R_transfer_N_photon(n, Me.N_photon, aux_Visual)

            ' отображаем ход решения

            T_1 = Microsoft.VisualBasic.DateAndTime.Timer

            If T_1 - T_0 > 0.25 Then

                Old_Progress = Progress

                Progress = (Current_time - Start_time) / Lenght_time * 100

                RaiseEvent Solution_progress(Progress)

                If Progress > Old_Progress Then

                    Main_form.Save_model()

                End If

            End If

            Application.DoEvents()

            T_0 = Microsoft.VisualBasic.DateAndTime.Timer

            n += 1

        Loop

        Main_form.Save_model()

        finish = Microsoft.VisualBasic.DateAndTime.Timer

    End Sub

    ''' <summary>
    ''' This sub solve problem using simple Euler integration method with variable time step.
    ''' This method uses radiation transfer matrix.
    ''' </summary>
    ''' <param name="N_step_start">Step to begin analysis. Can be skipped to set default step =1.
    ''' If N_step_start is set, the analysis begins from this step.</param>
    ''' <remarks></remarks>
    Sub Euler_var_T_step_with_smoothing(Optional ByRef N_step_start As Integer = 1)

        Dim delta_t_min(UBound(HT_element)) As Double

        Dim delta_t_max(UBound(HT_element)) As Double

        Dim max_delta_t As Double

        Dim T_0, T_1 As Double

        T_0 = Microsoft.VisualBasic.DateAndTime.Timer

        ' правые части уравнений (Q_n/(c_n*m))
        ' текущего шага по времени, предыдущего шага, предпредыдущего и т.д.

        ' Начинаем отчет времени

        Current_time = Start_time

        Dim Progress As Integer = 0

        Dim Old_Progress As Integer = -1

        ' объявляем массивы шагов по времени, если нужно считать с первого шага

        If N_step_start = 1 Then

            Init_temperature_field()

            For i As Integer = 0 To UBound(HT_element)

                ReDim HT_element(i).HT_Step(0)

                ' заполняем начальные условия
                HT_element(i).HT_Step(0).T = HT_element(i).Element.Field.Temperature

                HT_element(i).HT_Step(0).Time = Start_time

            Next i

        End If

        ' Рассчитываем тепловые потоки на первом шаге

        Prepare_heat_transfer(N_step_start)

        ' рассчитываем тепловой поток на текущем шаге
        For i As Integer = 0 To UBound(HT_element)

            Fill_HT_element_argument(HT_element(i))

            Calc_Q_const(HT_element(i))

            HT_element(i).HT_Step(N_step_start).Q = HT_element(i).Q_summary

            'HT_element(i).HT_Step(N_step_start).Time = Current_time

        Next i


        ' номер шага по времени
        Dim n As Integer = N_step_start

        ' массив, в который кладем сглаженные мощности всех элементов на шаге n
        Dim Q_smoothed(UBound(HT_element)) As Double

        Dim Q_in_to_smoothing() As Double

        Dim t_in_to_smoothing() As Double

        ReDim Q_in_to_smoothing(N_steps - 1)

        ReDim t_in_to_smoothing(N_steps - 1)

        Do

            ' цикл по элементам

            ' проводим сглаживание графиков мощностей на элемент.

            If Smoothing_type = "LINEAR" And N_steps > 1 Then

                ' сглаживание проводим линейное. Коэффициенты подбираем методом наименьших квадратов

                For i As Integer = 0 To UBound(HT_element)

                    If n >= N_steps Then

                        For j As Integer = 0 To UBound(t_in_to_smoothing)

                            Q_in_to_smoothing(j) = HT_element(i).HT_Step(n - N_steps + j + 1).Q


                            t_in_to_smoothing(j) = HT_element(i).HT_Step(n - N_steps + j).Time

                        Next j

                        Q_smoothed(i) = Linear_smoothing(HT_element(i).HT_Step(n - 1).Time, t_in_to_smoothing, Q_in_to_smoothing)

                    Else

                        Q_smoothed(i) = HT_element(i).HT_Step(n).Q

                    End If

                Next i

            End If


            If Smoothing_type = "" Then

                For i As Integer = 0 To UBound(HT_element)

                    Q_smoothed(i) = HT_element(i).HT_Step(n).Q

                Next i

            End If


            ' расчитываем максимальный возможный шаг по времени

            max_delta_t = Double.MaxValue

            For i As Integer = 0 To UBound(HT_element)

                HT_element(i).Argument.Coord(0) = HT_element(i).HT_Step(n - 1).T

                HT_element(i).Argument.Coord(1) = Current_time

                With HT_element(i).Element

                    delta_t_min(i) = (K_T_min - 1) * HT_element(i).HT_Step(n - 1).T * .Medium.Calc_c(HT_element(i).Argument) * .Field.Mass / Q_smoothed(i)

                    delta_t_max(i) = (K_T_max - 1) * HT_element(i).HT_Step(n - 1).T * .Medium.Calc_c(HT_element(i).Argument) * .Field.Mass / Q_smoothed(i)

                    If delta_t_min(i) < max_delta_t And delta_t_min(i) > 0 Then

                        max_delta_t = delta_t_min(i)

                    End If

                    If delta_t_max(i) < max_delta_t And delta_t_max(i) > 0 Then

                        max_delta_t = delta_t_max(i)

                    End If

                End With

            Next i

            If Lenght_time + Start_time - Current_time < max_delta_t Then

                max_delta_t = Lenght_time + Start_time - Current_time

            End If

            Step_time = max_delta_t

            Current_time += Step_time

            ' явно задаем температуру из T_SET

            '  для этого задаемся текущим временем 

            Dim Argument As Vector = New Vector

            ReDim Argument.Coord(5)

            Argument.Coord(1) = Current_time

            Dim Number As Integer

            For i As Integer = 0 To UBound(FEM.T_SET)

                For j As Integer = 0 To UBound(FEM.T_SET(i).Element)

                    Number = FEM.T_SET(i).Element(j).HT_element.Number

                    If FEM.T_SET(i).Temp(j).GetType.ToString = "System.Double" Then

                        HT_element(Number).HT_Step(n).T = FEM.T_SET(i).Temp(j)

                    Else

                        HT_element(Number).HT_Step(n).T = FEM.T_SET(i).Temp(j).Return_value(Argument).Coord(0)

                    End If

                    HT_element(Number).Element.Field.Temperature = HT_element(i).HT_Step(n).T

                Next j

            Next i


            If n = 1 Then

                ' циклы по элементам
                ' рассчитываем температуру на текущем шаге
                For i As Integer = 0 To UBound(HT_element)

                    If HT_element(i).T_SET Is Nothing Then

                        HT_element(i).HT_Step(n).Time = Current_time

                        With HT_element(i).Element

                            'HT_element(i).HT_Step(n).T = HT_element(i).HT_Step(n - 1).T + Step_time * Q_smoothed(i) / (.Medium.Calc_c(HT_element(i).Argument) * .Field.Mass)
                            '.Field.Temperature = HT_element(i).HT_Step(n).T

                            HT_element(i).HT_Step(n).T = Calc_T(HT_element(i), HT_element(i).HT_Step(n - 1).T, HT_element(i).HT_Step(n - 1).T, Q_smoothed(i) * Step_time / .Field.Mass)

                            .Field.Temperature = HT_element(i).HT_Step(n).T

                            If Double.IsInfinity(.Field.Temperature) Or Double.IsNaN(.Field.Temperature) Then

                                Throw New System.Exception("Solution is instabile. Step number " & n.ToString)

                                Exit Sub

                            End If

                        End With

                    End If

                Next i


            Else

                ' циклы по элементам
                ' рассчитываем температуру на текущем шаге
                For i As Integer = 0 To UBound(HT_element)

                    If HT_element(i).T_SET Is Nothing Then

                        HT_element(i).HT_Step(n).Time = Current_time

                        With HT_element(i).Element

                            'HT_element(i).HT_Step(n).T = HT_element(i).HT_Step(n - 1).T + Step_time * Q_smoothed(i) / (.Medium.Calc_c(HT_element(i).Argument) * .Field.Mass)
                            '.Field.Temperature = HT_element(i).HT_Step(n).T

                            HT_element(i).HT_Step(n).T = Calc_T(HT_element(i), HT_element(i).HT_Step(n - 1).T, HT_element(i).HT_Step(n - 2).T, Q_smoothed(i) * Step_time / .Field.Mass)

                            .Field.Temperature = HT_element(i).HT_Step(n).T

                            If Double.IsInfinity(.Field.Temperature) Or Double.IsNaN(.Field.Temperature) Then

                                Throw New System.Exception("Solution is instabile. Step number " & n.ToString)

                                Exit Sub

                            End If

                        End With

                    End If

                Next i


            End If

            If Current_time >= Start_time + Lenght_time Then Exit Do

            ' рассчитываем тепловой поток на следующем шаге шаге

            Prepare_heat_transfer(n + 1)


            ' рассчитываем тепловой поток на текущем шаге
            For i As Integer = 0 To UBound(HT_element)

                If HT_element(i).T_SET Is Nothing Then

                    Fill_HT_element_argument(HT_element(i))

                    Calc_Q_const(HT_element(i))

                    HT_element(i).HT_Step(n + 1).Q = HT_element(i).Q_summary

                End If

                HT_element(i).HT_Step(n + 1).Time = Current_time

            Next i

            ' отображаем ход решения

            T_1 = Microsoft.VisualBasic.DateAndTime.Timer

            If T_1 - T_0 > 0.25 Then

                Old_Progress = Progress

                Progress = (Current_time - Start_time) / Lenght_time * 100

                RaiseEvent Solution_progress(Progress)

                If Progress > Old_Progress Then

                    Main_form.Save_model()

                End If

            End If

            Application.DoEvents()

            T_0 = Microsoft.VisualBasic.DateAndTime.Timer

            n += 1

        Loop

        Main_form.Save_model()

    End Sub



    ''' <summary>
    ''' This function realized 2D linear smoothing
    ''' </summary>
    ''' <param name="x">argument to calculate Y</param>
    ''' <param name="X_raw">unsmoothed raw X data </param>
    ''' <param name="Y_raw">unsmoothed raw Y data</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function Linear_smoothing(ByRef x As Double, ByRef X_raw() As Double, ByRef Y_raw() As Double) As Double

        Dim D_0, D_1, Y_0, Y_1, S_0_0, S_1_0, a_0, b_0, a_1, b_1, A__0, A__1 As Double

        For i As Integer = 0 To UBound(X_raw)

            D_0 += 1

            D_1 += X_raw(i) * X_raw(i)

            Y_0 += Y_raw(i)

            Y_1 += Y_raw(i) * X_raw(i)

            S_0_0 += X_raw(i)

            S_1_0 += X_raw(i)

        Next i

        a_0 = Y_0 / D_0

        b_0 = S_0_0 / D_0

        a_1 = Y_1 / D_1

        b_1 = S_1_0 / D_1

        A__0 = (-a_0 + b_0 * a_1) / (-1 + b_1 * b_0)

        A__1 = (-a_1 + b_1 * a_0) / (-1 + b_1 * b_0)



        'Dim writer As System.IO.StreamWriter = New System.IO.StreamWriter("C:\tmp\X_raw.txt")

        'For i As Integer = 0 To UBound(X_raw)

        '    writer.WriteLine(X_raw(i))

        'Next i

        'writer.Close()



        'writer = New System.IO.StreamWriter("C:\tmp\Y_raw.txt")

        'For i As Integer = 0 To UBound(Y_raw)

        '    writer.WriteLine(Y_raw(i))

        'Next i

        'writer.Close()



        'writer = New System.IO.StreamWriter("C:\tmp\Coeff.txt")

        'writer.WriteLine(A__0)

        'writer.WriteLine(A__1)

        'writer.WriteLine(x)

        'writer.WriteLine(A__0 + A__1 * x)

        'writer.Close()

        Return A__0 + A__1 * x

    End Function

    ''' <summary>
    ''' This function realized 2D N-th order smoothing
    ''' </summary>
    ''' <param name="x">argument to calculate Y</param>
    ''' <param name="N ">order of smoothing</param>
    ''' <param name="X_raw">unsmoothed raw X data </param>
    ''' <param name="Y_raw">unsmoothed raw Y data</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function N_th_order_smoothing(ByRef x As Double, ByRef N As Integer, ByRef X_raw() As Double, ByRef Y_raw() As Double) As Double

        Dim A(N), C(N, N), Inv_C(,), Y(N), Y_smoothed As Double

        Dim M_points As Integer = UBound(X_raw)

        Dim i, j, m As Integer

        For i = 0 To N

            For j = i To N

                For m = 0 To M_points

                    ' Эту матрицу можно вычислять перед началом шага. Она одинаковая для всех элементов.
                    ' Вообше, всю эту процедуру можно отдельно не выделять, а встроить прямо в интегратор.
                    C(i, j) += X_raw(m) ^ (i + j)

                    C(j, i) = C(i, j)

                Next m

            Next j

            For m = 0 To M_points

                Y(i) += Y_raw(m) * X_raw(m) ^ (i)

            Next m

        Next i

        Inv_C = Inverse_Matrix(C)

        For i = 0 To N

            For j = 0 To N

                A(i) += Inv_C(i, j) * Y(j)

            Next j

        Next i

        For i = 0 To N

            Y_smoothed += A(i) * x ^ i

        Next i

        Return Y_smoothed

    End Function


    '''' <summary>
    '''' This sub solve problem using simple Euler integration method with variable time step.
    '''' Radiation heat transfer matrix is not used for this method.
    '''' </summary>
    '''' <param name="N_step_start">Step to begin analysis. Can be skipped to set default step =1.
    '''' If N_step_start is set, the analysis begins from this step.</param>
    '''' <remarks></remarks>
    'Sub Euler_var_T_step_MATRIXLESS(Optional ByRef N_step_start As Integer = 1)

    '    Dim delta_t_min(UBound(HT_element)) As Double

    '    Dim delta_t_max(UBound(HT_element)) As Double

    '    Dim max_delta_t As Double

    '    Dim T_0, T_1 As Double

    '    Dim start, finish As Double

    '    start = Microsoft.VisualBasic.DateAndTime.Timer

    '    T_0 = Microsoft.VisualBasic.DateAndTime.Timer

    '    ' правые части уравнений (Q_n/(c_n*m))
    '    ' текущего шага по времени, предыдущего шага, предпредыдущего и т.д.

    '    ' Начинаем отчет времени

    '    Current_time = Start_time

    '    Dim Progress As Integer = 0

    '    Dim Old_Progress As Integer = -1

    '    ' создаем визуализационную форму для расчета теплопередачи

    '    ' подготавливаем модель

    '    For i As Integer = 0 To UBound(FEM.Shining_face)

    '        FEM.Shining_face(i).code = New Face_code

    '    Next i

    '    FEM.Fill_Shining_faces_colours()

    '    FEM.Polygon = FEM.Extract_Polygon_from_Face(FEM.Shining_face)

    '    ' подготавливаем OpenGL для работы

    '    Dim Center As New Visualisation.Vertex

    '    Center.x = FEM.Center.Coord(0)

    '    Center.y = FEM.Center.Coord(1)

    '    Center.z = FEM.Center.Coord(2)

    '    Dim V_mode As New Visualisation.Mode

    '    V_mode.Status = Visualisation.Mode.WITHOUT_CONTOUR

    '    Dim aux_Visual As Visualisation.Visualisation

    '    Dim Model_list_number As Integer = 100

    '    aux_Visual = New Visualisation.Visualisation(V_mode, Model_list_number, FEM.Polygon, UBound(Me.FEM.Polygon) + 1, 400, 400, 16, False, Center, 2 * FEM.Maximum_length, FEM.Minimum_length, 0.1)

    '    ' визуализационная форма для работы готова

    '    ' объявляем массивы шагов по времени, если нужно считать с первого шага

    '    If N_step_start = 1 Then

    '        Init_temperature_field()

    '        For i As Integer = 0 To UBound(HT_element)

    '            ReDim HT_element(i).HT_Step(0)

    '            ' заполняем начальные условия
    '            HT_element(i).HT_Step(0).T = HT_element(i).Element.Field.Temperature

    '            HT_element(i).HT_Step(0).Time = Start_time

    '        Next i

    '    End If

    '    ' Рассчитываем тепловые потоки на первом шаге

    '    Prepare_heat_transfer(N_step_start)

    '    Calc_C_transfer(N_step_start - 1)

    '    Calc_R_transfer_N_photon(N_step_start - 1, Me.N_photon, aux_Visual)


    '    ' номер шага по времени
    '    Dim n As Integer = N_step_start

    '    Do

    '        ' цикл по элементам
    '        ' расчитываем максимальный возможный шаг по времени

    '        max_delta_t = Double.MaxValue

    '        For i As Integer = 0 To UBound(HT_element)

    '            HT_element(i).Argument.Coord(0) = HT_element(i).HT_Step(n - 1).T

    '            HT_element(i).Argument.Coord(1) = Current_time

    '            With HT_element(i).Element

    '                delta_t_min(i) = (K_T_min - 1) * HT_element(i).HT_Step(n - 1).T * .Medium.Calc_c(HT_element(i).Argument) * .Field.Mass / HT_element(i).HT_Step(n).Q

    '                delta_t_max(i) = (K_T_max - 1) * HT_element(i).HT_Step(n - 1).T * .Medium.Calc_c(HT_element(i).Argument) * .Field.Mass / HT_element(i).HT_Step(n).Q

    '                If delta_t_min(i) < max_delta_t And delta_t_min(i) > 0 Then

    '                    max_delta_t = delta_t_min(i)

    '                End If

    '                If delta_t_max(i) < max_delta_t And delta_t_max(i) > 0 Then

    '                    max_delta_t = delta_t_max(i)

    '                End If

    '            End With

    '        Next i

    '        If Lenght_time + Start_time - Current_time < max_delta_t Then

    '            max_delta_t = Lenght_time + Start_time - Current_time

    '        End If

    '        Step_time = max_delta_t

    '        Current_time += Step_time


    '        ' циклы по элементам
    '        ' рассчитываем температуру на текущем шаге
    '        For i As Integer = 0 To UBound(HT_element)

    '            HT_element(i).HT_Step(n).Time = Current_time

    '            With HT_element(i).Element

    '                HT_element(i).HT_Step(n).T = HT_element(i).HT_Step(n - 1).T + Step_time * HT_element(i).HT_Step(n).Q / (.Medium.Calc_c(HT_element(i).Argument) * .Field.Mass)
    '                .Field.Temperature = HT_element(i).HT_Step(n).T

    '                If Double.IsInfinity(.Field.Temperature) Or Double.IsNaN(.Field.Temperature) Then

    '                    Throw New System.Exception("Solution is instabile. Step number " & n.ToString)

    '                    Exit Sub

    '                End If

    '            End With

    '        Next i

    '        ' явно задаем температуру из T_SET

    '        '  для этого задаемся текущим временем 

    '        Dim Argument As Vector = New Vector

    '        ReDim Argument.Coord(5)

    '        Argument.Coord(1) = Current_time

    '        Dim Number As Integer

    '        For i As Integer = 0 To UBound(FEM.T_SET)

    '            For j As Integer = 0 To UBound(FEM.T_SET(i).Element)

    '                Number = FEM.T_SET(i).Element(j).HT_element.Number

    '                If FEM.T_SET(i).Temp(j).GetType.ToString = "System.Double" Then

    '                    HT_element(Number).HT_Step(n).T = FEM.T_SET(i).Temp(j)

    '                Else

    '                    HT_element(Number).HT_Step(n).T = FEM.T_SET(i).Temp(j).Return_value(Argument)

    '                End If

    '            Next j

    '        Next i

    '        If Current_time >= Start_time + Lenght_time Then Exit Do

    '        ' рассчитываем тепловой поток на следующем шаге шаге

    '        Prepare_heat_transfer(n + 1)

    '        Calc_C_transfer(n)

    '        Calc_R_transfer_N_photon(n, Me.N_photon, aux_Visual)

    '        ' отображаем ход решения

    '        T_1 = Microsoft.VisualBasic.DateAndTime.Timer

    '        If T_1 - T_0 > 0.25 Then

    '            Old_Progress = Progress

    '            Progress = (Current_time - Start_time) / Lenght_time * 100

    '            RaiseEvent Solution_progress(Progress)

    '            If Progress > Old_Progress Then

    '                Main_form.Save_model()

    '            End If

    '        End If

    '        Application.DoEvents()

    '        T_0 = Microsoft.VisualBasic.DateAndTime.Timer

    '        n += 1

    '    Loop

    '    Main_form.Save_model()

    '    finish = Microsoft.VisualBasic.DateAndTime.Timer

    'End Sub


    ''' <summary>
    ''' !!!!
    ''' !!!! Необходимо добавить элементы при постоянной температуре и переделать под переменные теплофизические и термооптические коэффициенты!!!
    ''' !!!!
    ''' This sub solve problem using simple Euler integration method.
    ''' Radiation heat transfer matrix is not used for this method.
    ''' Distributed computing is used for this method.
    ''' </summary>
    ''' <param name="Work_dir">The directory, where all of work files are placed.</param>
    ''' <remarks></remarks>
    Sub Euler_var_MATRIXLESS_DC(ByRef Work_dir As String)

        ' Ставим флаг стабильного решения

        Solution_integrity = True

        ' Начинаем отcчет времени

        Dim T_0, T_1 As Double

        T_0 = Microsoft.VisualBasic.DateAndTime.Timer

        Current_time = Start_time

        Dim Progress As Integer = 0

        ' создаем визуализационную форму для расчета теплопередачи (на всякий случай, вдруг в процессе расчета исчезнут все клиенты)

        ' подготавливаем модель

        For i As Integer = 0 To UBound(FEM.Shining_face)

            FEM.Shining_face(i).code = New Face_code

        Next i

        FEM.Fill_Shining_faces_colours()

        FEM.Polygon = FEM.Extract_Polygon_from_Face(FEM.Shining_face)

        Init_temperature_field()

        ' подготавливаем OpenGL для работы

        Dim Center As New Visualisation.Vertex

        Center.x = FEM.Center.Coord(0)

        Center.y = FEM.Center.Coord(1)

        Center.z = FEM.Center.Coord(2)

        Dim V_mode As New Visualisation.Mode

        V_mode.Status = Visualisation.Mode.WITHOUT_CONTOUR

        Dim aux_Visual As Visualisation.Visualisation

        Dim Model_list_number As Integer = 100

        aux_Visual = New Visualisation.Visualisation(V_mode, Model_list_number, FEM.Polygon, UBound(Me.FEM.Polygon) + 1, 400, 400, 16, False, Center, 2 * FEM.Maximum_length, FEM.Minimum_length, 0.1)

        ' визуализационная форма для работы готова


        ' объявляем массивы шагов по времени

        For i As Integer = 0 To UBound(HT_element)

            ReDim HT_element(i).HT_Step(Lenght_time / Step_time)

            For j As Integer = 0 To UBound(HT_element(i).HT_Step)

                HT_element(i).HT_Step(j) = New HT_Step

            Next j

            ' заполняем начальные условия
            HT_element(i).HT_Step(0).T = HT_element(i).Element.Field.Temperature

        Next i


        ' если мы работаем в режиме ФТП, то подсоединяемся к ФТП-серверу

        Dim login, server, pass As String

        login = ""

        server = ""

        pass = ""

        Dim hOpen As Integer

        Dim hConnection As Integer

        If FTP_mode = True Then

            ' выделяем имя сервера, логин и пароль

            login = Mid(Work_dir, "ftp://".Length + 1, InStrRev(Work_dir, ":") - "ftp://".Length - 1)

            Try

                server = Mid(Work_dir, InStrRev(Work_dir, "@") + 1, InStr("ftp://".Length + 1, Work_dir, "/") - InStrRev(Work_dir, "@") - 1)

            Catch

                server = Mid(Work_dir, InStrRev(Work_dir, "@") + 1)

            End Try

            pass = Work_dir.Replace(login & ":", "")

            pass = pass.Replace("@" & server, "")

            pass = Mid(pass, "ftp://".Length + 1)

            Try

                pass = Mid(pass, 1, InStr(pass, "/") - 1)

            Catch


            End Try


            ' устанавливаем соединение с Интернетом
            hOpen = FTP.InternetOpen("T.H.O.R.I.U.M.", FTP.INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)

            ' устанавливаем соединение с сервером
            hConnection = FTP.InternetConnect(hOpen, server, FTP.INTERNET_DEFAULT_FTP_PORT, login, pass, FTP.INTERNET_SERVICE_FTP, IIf(FTP.PassiveConnection, FTP.INTERNET_FLAG_PASSIVE, 0), 0)

            ' устанавливаем правильную директорию

            FTP.FtpSetDirHierarchy(hConnection, Main_form.FTP_Dir_Hierarchy)

        End If

        ' если в рабочей директории есть файл END, то удаляем его.

        If FTP_mode Then

            If FTP.FtpCurrentDirectoryFileExists(hConnection, "END") Then

                FTP.FtpDeleteFile(hConnection, "END")

            End If

        Else

            If System.IO.File.Exists(Work_dir & "END") Then

                System.IO.File.Delete(Work_dir & "END")

            End If

        End If


        Prepare_heat_transfer(1)

        Solution_integrity = True

        Calc_Q_DC(Work_dir, 0, aux_Visual, hConnection)


        ' цикл по шагам времени
        For n As Integer = 1 To Lenght_time / Step_time

            Current_time += Step_time

            ' проверяем есть ли соединение. Если нет, переподсоединяемся

            If FTP_mode Then

                If Not Solution_integrity Then

                    FTP.InternetCloseHandle(hConnection)

                    FTP.InternetCloseHandle(hOpen)

                    ' устанавливаем соединение с Интернетом
                    hOpen = FTP.InternetOpen("T.H.O.R.I.U.M.", FTP.INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)

                    ' устанавливаем соединение с сервером
                    hConnection = FTP.InternetConnect(hOpen, server, FTP.INTERNET_DEFAULT_FTP_PORT, login, pass, FTP.INTERNET_SERVICE_FTP, IIf(FTP.PassiveConnection, FTP.INTERNET_FLAG_PASSIVE, 0), 0)

                    ' устанавливаем правильную директорию

                    FTP.FtpSetDirHierarchy(hConnection, Main_form.FTP_Dir_Hierarchy)

                End If

            End If

            ' циклы по элементам
            ' рассчитываем температуру на текущем шаге
            For i As Integer = 0 To UBound(HT_element)

                With HT_element(i).Element

                    HT_element(i).HT_Step(n).T = HT_element(i).HT_Step(n - 1).T + Step_time * HT_element(i).HT_Step(n).Q / (.Medium.c * .Field.Mass)

                    .Field.Temperature = HT_element(i).HT_Step(n).T

                    If Double.IsInfinity(.Field.Temperature) Or Double.IsNaN(.Field.Temperature) Then

                        Throw New System.Exception("Solution is instabile. Step number " & n.ToString)

                        Exit Sub

                    End If

                End With

            Next i

            If n >= Lenght_time / Step_time Then Exit For

            ' рассчитываем тепловой поток на следующем шаге шаге

            Prepare_heat_transfer(n + 1)

            Solution_integrity = True

            Calc_Q_DC(Work_dir, n, aux_Visual, hConnection)

            ' отображаем ход решения

            T_1 = Microsoft.VisualBasic.DateAndTime.Timer

            If T_1 - T_0 > 0.25 Then

                If Progress + 1 <= n / (Lenght_time / Step_time) * 100 Then

                    Progress = n * Step_time / Lenght_time * 100

                    RaiseEvent Solution_progress(Progress)

                End If

            End If

            Application.DoEvents()

            T_0 = Microsoft.VisualBasic.DateAndTime.Timer

        Next n

        ' Завершаем распределенные вычисления

        If FTP_mode Then

            Dim writer As System.IO.StreamWriter = New System.IO.StreamWriter(tmp_Dir & "END")

            writer.Close()

            FTP.FtpPutFile(hConnection, tmp_Dir & "END", "END", FTP.FTP_TRANSFER_TYPE_UNKNOWN, 0)

            FTP.InternetCloseHandle(hConnection)

            FTP.InternetCloseHandle(hOpen)

        Else

            Dim writer As System.IO.StreamWriter = New System.IO.StreamWriter(Work_dir & "END")

            writer.Close()


        End If

    End Sub
    ''' <summary>
    ''' This sub calculates heat power on selected step using distributed computing
    ''' </summary>
    ''' <param name="Work_dir"></param>
    ''' <param name="Step_number"></param>
    ''' <param name="aux_Visual"></param>
    ''' <param name="hConnection">Connection to FTP server handle</param>
    ''' <remarks></remarks>
    Public Sub Calc_Q_DC(ByRef Work_dir As String, ByRef Step_number As Integer, ByRef aux_Visual As Visualisation.Visualisation, ByRef hConnection As Integer)

        ' Рассчитываем тепловые потоки на первом шаге

        '   Очищаем рабочую директорию

        If FTP_mode Then

            FTP.FtpCurrentDirectoryFileListDelete(hConnection, "*.T")

            FTP.FtpCurrentDirectoryFileListDelete(hConnection, "*.Q")

            FTP.FtpCurrentDirectoryFileListDelete(hConnection, "*.station")

            FTP.FtpCurrentDirectoryFileListDelete(hConnection, "*.element")

            FTP.FtpCurrentDirectoryFileListDelete(hConnection, "*.tmp")

        End If

        '   Формируем файл температур

        Write_temperature_file_to_dir(Step_number, Work_dir, hConnection)

        '   Ждем заданное время

        Dim T_wait As Double = 1000 ' время ожидания файлов с идентификаторами

        System.Threading.Thread.Sleep(T_wait)


        '   Выясняем сколько клиентов готовы принять участие в вычислениях на текущем шаге

        Dim Station() As String

        If FTP_mode Then

            Station = FTP.FtpGetCurrentDirectoryFileList(hConnection, "*.station")

        Else

            Station = System.IO.Directory.GetFiles(Work_dir, "*.station")

        End If


        ' удаляем файл температур


        If FTP_mode Then

            'Solution_integrity = FTP.FtpDeleteFile(hConnection, "Step_" & Trim(Step_number.ToString) & ".T")

            FTP.FtpCurrentDirectoryFileListDelete(hConnection, "*.T")

        Else

            System.IO.File.Delete(Work_dir & "Step_" & Trim(Step_number.ToString) & ".T")

        End If



        ' Получили информацию о готовых клиентах, удаляем файлы *.station
        If Not (Station Is Nothing) Then

            If Station.Length > 0 Then

                For i As Integer = 0 To UBound(Station)

                    If FTP_mode Then

                        Solution_integrity = FTP.FtpDeleteFile(hConnection, Station(i))

                    Else

                        System.IO.File.Delete(Station(i))

                    End If



                Next i

            End If

        End If

        Dim P_index(UBound(Station)) As Double

        Dim ID(UBound(Station)) As Long

        Dim Heat_power(,) As Double

        Dim N_1(UBound(Station)), N_2(UBound(Station)) As Long

        Do ' 'этот цикл для того, чтобы можно было завершать расчет и не переходить по Goto

            '   Если никто не согласился принять участие, то рассчитываем шаг самостоятельно
            If Station.Length = 0 Then

                ' Рассчитываем шаг самостоятельно

                ' Расчет теплопередачи

                Calc_C_transfer(Step_number)

                ' Расчет тепловвых мощностей от источников излучения

                If Me.N_photon > 0 Then

                    Heat_power = Calc_R_transfer_R_heat_source_N_photon(Me.N_photon, aux_Visual)

                Else

                    Exit Sub

                    'Heat_power = Calc_R_transfer_R_heat_source_TOP_Epsilon(Me.TOP_Epsilon, aux_Visual)

                End If

                Add_heat_power_from_array_to_HT_elements(Heat_power, Step_number)

                ' Расчет теплопредачи в модели

                If Me.N_photon > 0 Then

                    Heat_power = Calc_R_transfer_N_photon(0, UBound(HT_element), Step_number, Me.N_photon, aux_Visual)

                Else

                    Exit Sub

                    'Heat_power = Calc_R_transfer_TOP_Epsilon(0, UBound(HT_element), Step_number, Me.TOP_Epsilon, aux_Visual)

                End If

                Add_heat_power_from_array_to_HT_elements(Heat_power, Step_number)

                Exit Do

            Else

                ' Ура! Мы работаем не одни!
                ' Формируем массивы индексов производительности и идентификаторов

                For i As Integer = 0 To UBound(Station)

                    Station(i) = Replace(Station(i), Work_dir, "")

                    ID(i) = Val(Mid(Station(i), 1, InStr(Station(i), "_")))

                    P_index(i) = Val(Replace(Mid(Station(i), ID(i).ToString.Length + 2), ".station", ""))

                Next i

            End If

            ' Думаем, сколько элементов рассчитывать каждому клиенту

            Dim DC_part() As Double = Detect_part_of_model_for_clients(P_index)

            ' Рассчитываем продолжительность расчета одного шага

            Dim T_step As Double = 0

            For i As Integer = 0 To UBound(P_index)

                T_step += P_index(i) * DC_part(i)

            Next i

            ' Вводим запас в 10% и добавляем секунду на скачивание файлов с сервера и секунду на передачу файла на сервер.

            T_step = 1.1 * T_step + 1 + 1

            ' Придумали сколько элементов. Теперь создаем в рабочей директории файлы с указанием диапазона элементов для обработки.

            N_1(0) = 0

            N_2(0) = -1

            For i As Integer = 0 To UBound(Station)

                If i - 1 >= 0 Then

                    N_1(i) = N_2(i - 1) + 1

                    If N_1(i) > UBound(HT_element) Then

                        ReDim Preserve N_1(UBound(N_1) - 1)

                        ReDim Preserve N_2(UBound(N_2) - 1)

                        ReDim Preserve ID(UBound(ID) - 1)

                        Exit For

                    End If


                Else

                    N_1(i) = 0

                End If

                N_2(i) = N_1(i) + DC_part(i) * (UBound(HT_element))

                If N_2(i) > UBound(HT_element) Then N_2(i) = UBound(HT_element)

                'Dim writer As System.IO.StreamWriter = New System.IO.StreamWriter(Work_dir & ID(i).ToString & "_" & N_1(i).ToString & "_" & N_2(i).ToString & ".element")

                Dim writer As System.IO.StreamWriter = New System.IO.StreamWriter(tmp_Dir & ID(i).ToString & "_" & N_1(i).ToString & "_" & N_2(i).ToString & ".element")

                writer.Close()

                If FTP_mode Then

                    Solution_integrity = FTP.FtpPutFile(hConnection, tmp_Dir & ID(i).ToString & "_" & N_1(i).ToString & "_" & N_2(i).ToString & ".element", ID(i).ToString & "_" & N_1(i).ToString & "_" & N_2(i).ToString & ".element", FTP.FTP_TRANSFER_TYPE_UNKNOWN, 0)

                Else

                    System.IO.File.Copy(tmp_Dir & ID(i).ToString & "_" & N_1(i).ToString & "_" & N_2(i).ToString & ".element", Work_dir & ID(i).ToString & "_" & N_1(i).ToString & "_" & N_2(i).ToString & ".element")

                End If

                System.IO.File.Delete(tmp_Dir & ID(i).ToString & "_" & N_1(i).ToString & "_" & N_2(i).ToString & ".element")

            Next i

            ' Создали файлы. Клиенты их скачали и начали расчет. Ждем время завершения одного шага и собираем информацию.
            Dim T_0 As Double = Microsoft.VisualBasic.DateAndTime.Timer

            ' А пока клиенты считают, рассчитываем теплопроводность и мощности от источников излучения

            Calc_C_transfer(Step_number)

            If Me.N_photon > 0 Then

                Heat_power = Calc_R_transfer_R_heat_source_N_photon(Me.N_photon, aux_Visual)

            Else

                Exit Sub

                'Heat_power = Calc_R_transfer_R_heat_source_TOP_Epsilon(Me.TOP_Epsilon, aux_Visual)

            End If

            Add_heat_power_from_array_to_HT_elements(Heat_power, Step_number)

            ' Теперь ждем окончания расчетного времени шага

            Dim T_1 As Double = Microsoft.VisualBasic.DateAndTime.Timer

            If T_1 - T_0 < T_step Then

                System.Threading.Thread.Sleep((T_step - (T_1 - T_0)) * 1000)

            End If



            ' Подождали. Собираем файлы.

            Dim ID_calculation_succeseful As Boolean

            For i As Integer = 0 To UBound(ID)

                ' Если клиент все-таки посчитал, то прибавляем тепловые мощности от него в модель
                ID_calculation_succeseful = False

                If FTP_mode Then

                    If FTP.FtpCurrentDirectoryFileExists(hConnection, ID(i).ToString & ".Q") Then

                        ID_calculation_succeseful = True

                    End If

                Else

                    If System.IO.File.Exists(Work_dir & ID(i).ToString & ".Q") Then

                        ID_calculation_succeseful = True

                    End If

                End If


                If ID_calculation_succeseful Then

                    ' Считываем файл
                    Read_heat_power_from_file_in_dir_to_HT_element(ID(i), Work_dir, Step_number, hConnection)

                    ' Удаляем его
                    'System.IO.File.Delete(Work_dir & ID(i).ToString & ".Q")

                Else

                    ' Увы, клиент не посчитал. Прийдется считать самим.

                    If Me.N_photon > 0 Then

                        Heat_power = Calc_R_transfer_N_photon(N_1(i), N_2(i), Step_number, Me.N_photon, aux_Visual)

                    Else

                        Exit Sub

                        'Heat_power = Calc_R_transfer_TOP_Epsilon(N_1(i), N_2(i), Step_number, Me.TOP_Epsilon, aux_Visual)

                    End If

                    Add_heat_power_from_array_to_HT_elements(Heat_power, Step_number)

                End If

            Next i

            Exit Do

        Loop

        ' Тепловые мощности рассчитаны.

        ' Удаляем файл температур
        'System.IO.File.Delete(Work_dir & "Step_" & Trim(Step_number.ToString) & ".T")

        ' Удаляем файлы с номерами элементов

        For i As Integer = 0 To UBound(ID)

            If FTP_mode Then

                Solution_integrity = FTP.FtpDeleteFile(hConnection, ID(i).ToString & "_" & N_1(i).ToString & "_" & N_2(i).ToString & ".element")

            Else

                System.IO.File.Delete(Work_dir & ID(i).ToString & "_" & N_1(i).ToString & "_" & N_2(i).ToString & ".element")

            End If

        Next i


    End Sub


    ''' <summary>
    ''' This function calculated part of model to calculate for every client according to it's productivity index 
    ''' </summary>
    ''' <param name="P_index">Array productivity indexes</param>
    ''' <returns>Array of part for clients</returns>
    ''' <remarks></remarks>
    Function Detect_part_of_model_for_clients(ByRef P_index() As Double) As Double()

        Dim A(UBound(P_index)) As Double

        Dim I_summ, A_summ As Double

        I_summ = 0

        A_summ = 0

        For i As Integer = 0 To UBound(P_index)

            I_summ += P_index(i)

        Next i

        For i As Integer = 0 To UBound(P_index)

            A(i) = I_summ / P_index(i)

        Next i

        For i As Integer = 0 To UBound(P_index)

            A_summ += A(i)

        Next i

        Dim R_value(UBound(P_index)) As Double

        For i As Integer = 0 To UBound(P_index)

            R_value(i) = A(i) / A_summ

        Next i

        Return R_value

    End Function

    ''' <summary>
    ''' This sub creates file with temperatures on selected step and place it to work directory.
    ''' </summary>
    ''' <param name="Step_number"></param>
    ''' <param name="Work_dir"></param>
    ''' <remarks></remarks>
    Sub Write_temperature_file_to_dir(ByRef Step_number As Integer, ByRef Work_dir As String, ByRef hConnection As Integer)

        'Dim writer As System.IO.StreamWriter = New System.IO.StreamWriter(Work_dir & "Step_" & Trim(Step_number.ToString) & ".T")

        Dim writer As System.IO.StreamWriter = New System.IO.StreamWriter(tmp_Dir & "Step_" & Trim(Step_number.ToString) & ".T")

        For i As Integer = 0 To UBound(HT_element)

            writer.WriteLine(HT_element(i).HT_Step(Step_number).T)

        Next i

        writer.Close()

        If Me.FTP_mode Then

            FTP.FtpPutFile(hConnection, tmp_Dir & "Step_" & Trim(Step_number.ToString) & ".T", "Step_" & Trim(Step_number.ToString) & ".tmp", FTP.FTP_TRANSFER_TYPE_UNKNOWN, 0)

            System.IO.File.Delete(tmp_Dir & "Step_" & Trim(Step_number.ToString) & ".T")

            FTP.FtpRenameFile(hConnection, "Step_" & Trim(Step_number.ToString) & ".tmp", "Step_" & Trim(Step_number.ToString) & ".T")

            If Not (FTP.FtpCurrentDirectoryFileExists(hConnection, "Step_" & Trim(Step_number.ToString) & ".T")) Then Solution_integrity = False

        Else

            System.IO.File.Copy(tmp_Dir & "Step_" & Trim(Step_number.ToString) & ".T", Work_dir & "Step_" & Trim(Step_number.ToString) & ".tmp")

            System.IO.File.Delete(tmp_Dir & "Step_" & Trim(Step_number.ToString) & ".T")

            System.IO.File.Move(Work_dir & "Step_" & Trim(Step_number.ToString) & ".tmp", Work_dir & "Step_" & Trim(Step_number.ToString) & ".T")

            If Not (System.IO.File.Exists(Work_dir & "Step_" & Trim(Step_number.ToString) & ".T")) Then Solution_integrity = False

        End If

    End Sub

    ''' <summary>
    ''' This sub reads file with temperatures on selected step from work directory and fill HT_element array with these temperatures.
    ''' </summary>
    ''' <param name="Step_number"></param>
    ''' <param name="Work_dir"></param>
    ''' <remarks></remarks>
    Sub Read_temperature_file_from_dir(ByRef ID As Long, ByRef Step_number As Integer, ByRef Work_dir As String, ByRef hConnection As Integer)

        If FTP_mode Then

            FTP.FtpGetFile(hConnection, "Step_" & Trim(Step_number.ToString) & ".T", tmp_Dir & ID.ToString & "current_step.T", True, 0, FTP.FTP_TRANSFER_TYPE_UNKNOWN, 0)

        Else

            System.IO.File.Copy(Work_dir & "Step_" & Trim(Step_number.ToString) & ".T", tmp_Dir & ID.ToString & "current_step.T")

        End If

        If System.IO.File.Exists(tmp_Dir & ID.ToString & "current_step.T") Then

            Dim tmp_str As String

            Dim reader As System.IO.StreamReader = New System.IO.StreamReader(tmp_Dir & ID.ToString & "current_step.T")

            For i As Integer = 0 To UBound(HT_element)

                tmp_str = reader.ReadLine()

                If tmp_str Is Nothing Then Exit For

                tmp_str = tmp_str.Replace(",", ".")

                HT_element(i).HT_Step(Step_number).T = Val(tmp_str)

            Next i

            reader.Close()

            System.IO.File.Delete(tmp_Dir & ID.ToString & "current_step.T")

        Else

            Solution_integrity = False

        End If

    End Sub


    ''' <summary>
    ''' This sub writes heat power to model's elements to file
    ''' </summary>
    ''' <param name="ID"></param>
    ''' <param name="Work_dir"></param>
    ''' <param name="Q_array">2-dimension array for write to file. 
    ''' Q_array(0,x) - number of HT_element to add heat_power.
    ''' Q_array(1,x) - heat_power to add.
    ''' </param>
    ''' <remarks></remarks>
    Sub Write_heat_power_file_to_dir(ByRef ID As Long, ByRef Work_dir As String, ByRef Q_array(,) As Double, ByRef hConnection As Integer)

        'Dim writer As System.IO.FileStream = New System.IO.FileStream(Work_dir & ID.ToString & ".Q", IO.FileMode.Create, IO.FileAccess.Write)

        Dim writer As System.IO.FileStream = New System.IO.FileStream(tmp_Dir & ID.ToString & ".Q", IO.FileMode.Create, IO.FileAccess.Write)

        Dim w As New System.IO.BinaryWriter(writer)

        ' number of elements to store in file
        w.Write(CLng(UBound(Q_array, 2)))

        For i As Integer = 0 To UBound(Q_array, 2)

            w.Write(CLng(Q_array(0, i)))

            w.Write(CDbl(Q_array(1, i)))

        Next i

        w.Close()

        writer.Close()

        ' copy file from local temp directory to work directory

        If FTP_mode Then

            FTP.FtpPutFile(hConnection, tmp_Dir & ID.ToString & ".Q", ID.ToString & ".Q", FTP.FTP_TRANSFER_TYPE_UNKNOWN, 0)

            System.IO.File.Delete(tmp_Dir & ID.ToString & ".Q")

            If Not (FTP.FtpCurrentDirectoryFileExists(hConnection, ID.ToString & ".Q")) Then

                Solution_integrity = False

                Exit Sub

            End If

        Else

            Try

                System.IO.File.Move(tmp_Dir & ID.ToString & ".Q", Work_dir & ID.ToString & ".Q")

            Catch

                Write_heat_power_file_to_dir(ID, Work_dir, Q_array, hConnection)

            End Try

        End If

    End Sub

    ''' <summary>
    ''' This reads heat power from file and adds it to model's heat transfering elements on selected step number
    ''' </summary>
    ''' <param name="ID"></param>
    ''' <param name="Work_dir"></param>
    ''' <param name="Step_number">Number of step to add heat power</param>
    ''' <remarks></remarks>
    Sub Read_heat_power_from_file_in_dir_to_HT_element(ByRef ID As Long, ByRef Work_dir As String, ByRef Step_number As Integer, ByRef hConnection As Integer)

        ' copy file to local temp directory

        If FTP_mode Then

            FTP.FtpGetFile(hConnection, ID.ToString & ".Q", tmp_Dir & "Server.Q", True, 0, FTP.FTP_TRANSFER_TYPE_UNKNOWN, 0)

            FTP.FtpDeleteFile(hConnection, ID.ToString & ".Q")

        Else

            Try

                If System.IO.File.Exists(tmp_Dir & "Server.Q") Then

                    System.IO.File.Delete(tmp_Dir & "Server.Q")

                End If

                System.IO.File.Move(Work_dir & ID.ToString & ".Q", tmp_Dir & "Server.Q")

            Catch

                System.Threading.Thread.Sleep(1000)

                If System.IO.File.Exists(tmp_Dir & "Server.Q") Then

                    System.IO.File.Delete(tmp_Dir & "Server.Q")

                End If

                System.IO.File.Move(Work_dir & ID.ToString & ".Q", tmp_Dir & "Server.Q")

            End Try

        End If

        If Not (System.IO.File.Exists(tmp_Dir & "Server.Q")) Then

            Solution_integrity = False

            Exit Sub

        End If


        'Dim reader As System.IO.FileStream = New System.IO.FileStream(Work_dir & ID.ToString & ".Q", IO.FileMode.Open, IO.FileAccess.Read)

        Dim reader As System.IO.FileStream = New System.IO.FileStream(tmp_Dir & "Server.Q", IO.FileMode.Open, IO.FileAccess.Read)

        Dim r As New System.IO.BinaryReader(reader)

        Dim HT_element_Number As Long

        Dim HT_element_Q As Double

        Dim N_elements As Long = r.ReadInt64

        ' reading heat power info from file
        For i As Long = 0 To N_elements

            HT_element_Number = r.ReadInt64()

            HT_element_Q = r.ReadDouble()

            ' adding heat power to element

            HT_element(HT_element_Number).HT_Step(Step_number + 1).Q += HT_element_Q

        Next i

        r.Close()

        reader.Close()

        System.IO.File.Delete(tmp_Dir & "Server.Q")

    End Sub

    '''' <summary>
    '''' This sub solve problem using simple Euler integration method.
    '''' Solution without radiation heat transfer
    '''' </summary>
    '''' <remarks></remarks>
    'Sub Euler_const_NO_RADIATION_HEAT_TRANSFER()

    '    'Dim Mass, c As Double

    '    ' правые части уравнений (Q_n/(c_n*m))
    '    ' текущего шага по времени, предыдущего шага, предпредыдущего и т.д.

    '    'Dim F_n_1 As Double

    '    ' Начинаем отчет времени

    '    Current_time = Start_time

    '    Dim Progress As Integer = 0


    '    ' объявляем массивы шагов по времени

    '    For i As Integer = 0 To UBound(HT_element)

    '        ReDim HT_element(i).HT_Step(Lenght_time / Step_time)

    '        For j As Integer = 0 To UBound(HT_element(i).HT_Step)

    '            HT_element(i).HT_Step(j) = New HT_Step

    '        Next j

    '        ' заполняем начальные условия
    '        HT_element(i).HT_Step(0).T = HT_element(i).Element.Field.Temperature

    '    Next i


    '    ' цикл по шагам времени
    '    For n As Integer = 1 To Lenght_time / Step_time

    '        Current_time += Step_time


    '        ' рассчитываем тепловой поток на текущем шаге
    '        For i As Integer = 0 To UBound(HT_element)

    '            Calc_Q_const_NO_RADIATION_HEAT_TRANSFER(HT_element(i))

    '            HT_element(i).HT_Step(n).Q = HT_element(i).Q_summary

    '        Next i

    '        ' циклы по элементам
    '        ' рассчитываем температуру на текущем шаге
    '        For i As Integer = 0 To UBound(HT_element)

    '            With HT_element(i).Element

    '                'HT_element(i).HT_Step(n).T = HT_element(i).HT_Step(n - 1).T + Step_time * HT_element(i).HT_Step(n).Q / (.Medium.c * .Field.Mass)

    '                HT_element(i).HT_Step(n).T = HT_element(i).HT_Step(n - 1).T + Step_time * HT_element(i).Q_summary / (.Medium.c * .Field.Mass)

    '                .Field.Temperature = HT_element(i).HT_Step(n).T

    '                If Double.IsInfinity(.Field.Temperature) Or Double.IsNaN(.Field.Temperature) Then

    '                    Throw New System.Exception("Solution is instabile. Step number " & n.ToString)

    '                    Exit Sub

    '                End If

    '            End With

    '        Next i

    '        If Progress + 1 <= n / (Lenght_time / Step_time) * 100 Then

    '            Progress += 1

    '            RaiseEvent Solution_progress(Progress)

    '        End If

    '        Application.DoEvents()

    '    Next n

    'End Sub

    ''' <summary>
    ''' This sub calculates element's face codes, with taking in to account element's temperature.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Calc_Elements_Faces_codes()

        Dim T As Double

        Detect_Min_Max_T()

        Dim Glob_Max_T As Double = Max_T

        Dim Glob_Min_T As Double = Min_T

        If Colors_from_entire_solution Then

            If Min_T = Max_T Then Exit Sub

            For i_Step As Integer = 0 To UBound(HT_element(0).HT_Step)

                ' Определяем коды цветов для граней
                For i As Integer = 0 To UBound(HT_element)

                    T = HT_element(i).HT_Step(i_Step).T

                    HT_element(i).HT_Step(i_Step).Code = Field_color(T, Min_T, Max_T)

                Next i

            Next i_Step

        Else

            Dim old_Current_step As Integer = Current_step

            For i_Step As Integer = 0 To UBound(HT_element(0).HT_Step)

                Me.Current_step = i_Step

                Detect_Min_Max_T_on_current_step()

                If Min_T = Max_T Then

                    Min_T = Glob_Min_T

                    Max_T = Glob_Max_T

                End If

                ' Определяем коды цветов для граней
                For i As Integer = 0 To UBound(HT_element)

                    T = HT_element(i).HT_Step(i_Step).T

                    HT_element(i).HT_Step(i_Step).Code = Field_color(T, Min_T, Max_T)

                Next i

            Next i_Step

            Current_step = old_Current_step

        End If

    End Sub
    Public Sub Prepare_Draw_Temperature(ByRef i_Step As Integer)

        Dim Text As String

        Dim Q_text As String

        For i As Integer = 0 To UBound(HT_element)

            Text = "Element " & HT_element(i).Element.Number.ToString & "." & Chr(13) & _
            "Temperature = " & HT_element(i).HT_Step(i_Step).T & " K." & Chr(13) & _
            "Heat input power = "

            If HT_element(i).HT_Step(i_Step).Q > 0 Then

                Q_text = "+" & HT_element(i).HT_Step(i_Step).Q.ToString() & "W."

            Else

                Q_text = HT_element(i).HT_Step(i_Step).Q.ToString() & "W."

            End If


            If HT_element(i).HT_Step(i_Step).Q = 0 Then

                Q_text = "0 W."

            End If

            Text &= Q_text

            For j As Integer = 0 To UBound(HT_element(i).Face)

                HT_element(i).Face(j).Code = HT_element(i).HT_Step(i_Step).Code

                HT_element(i).Face(j).Text = Text

            Next j

        Next i

    End Sub


    Public Sub Prepare_Elements_to_show_before_analysis()

        Dim Text, Text_face As String

        Dim Lambda, c, R_dif, R_mir, A, D, eps, T As Double

        Dim Argument As New Vector

        ReDim Argument.Coord(5)

        For i As Integer = 0 To UBound(HT_element)

            HT_element(i).Argument.Coord(0) = HT_element(i).Element.Field.Temperature

            HT_element(i).Argument.Coord(1) = FEM.Start_time

            Lambda = HT_element(i).Element.Medium.Calc_lambda(HT_element(i).Argument)

            c = HT_element(i).Element.Medium.Calc_c(HT_element(i).Argument)

            T = HT_element(i).Element.Field.Temperature

            Text = "Element " & HT_element(i).Element.Number.ToString & "." & Chr(13) & _
            "Material " & HT_element(i).Element.Element_property.Number.ToString & Chr(13) & _
            "Density (ro) " & HT_element(i).Element.Field.Density.ToString & " kg/m^3 " & Chr(13) & _
            "Thermal conductivity (lambda) " & Lambda & " W/(m*K) " & Chr(13) & _
            "Specific thermal capacity (c) " & c & " J/(kg*K) " & Chr(13) & _
            "Temperature " & T & " K " & Chr(13)

            For j As Integer = 0 To UBound(HT_element(i).Face)

                HT_element(i).Face(j).Code = New Face_code

                Argument.Coord(0) = T

                Argument.Coord(1) = FEM.Start_time

                If HT_element(i).Face(j).FTOP.R_dif.GetType.ToString <> "System.Double" Then

                    R_dif = HT_element(i).Face(j).FTOP.R_dif.return_value(Argument).Coord(0)

                Else

                    R_dif = HT_element(i).Face(j).FTOP.R_dif

                End If


                If HT_element(i).Face(j).FTOP.R_mir.GetType.ToString <> "System.Double" Then

                    R_mir = HT_element(i).Face(j).FTOP.R_mir.return_value(Argument).Coord(0)

                Else

                    R_mir = HT_element(i).Face(j).FTOP.R_mir

                End If


                If HT_element(i).Face(j).FTOP.A.GetType.ToString <> "System.Double" Then

                    A = HT_element(i).Face(j).FTOP.A.return_value(Argument).Coord(0)

                Else

                    A = HT_element(i).Face(j).FTOP.A

                End If


                If HT_element(i).Face(j).FTOP.D.GetType.ToString <> "System.Double" Then

                    D = HT_element(i).Face(j).FTOP.D.Return_value(Argument).Coord(0)

                Else

                    D = HT_element(i).Face(j).FTOP.D

                End If


                If HT_element(i).Face(j).FTOP.eps.GetType.ToString <> "System.Double" Then

                    eps = HT_element(i).Face(j).FTOP.eps.Return_value(Argument).Coord(0)

                Else

                    eps = HT_element(i).Face(j).FTOP.eps

                End If



                Text_face = "Face thermal-optical properties (FTOP) " & HT_element(i).Face(j).FTOP.Number.ToString & Chr(13) & _
                "Diffuse reflection coefficient (R_dif) " & R_dif & Chr(13) & _
                "Mirror reflection coefficient (R_mir) " & R_mir & Chr(13) & _
                "Absorbtion coefficient (A) " & A & Chr(13) & _
                "Transmission coefficient (D) " & D & Chr(13) & _
                "Emissivity coefficient (eps) " & eps & Chr(13)

                HT_element(i).Face(j).Text = Text & Text_face

                Select Case j

                    Case 0

                        HT_element(i).Face(j).Code.B = 255

                    Case 1

                        HT_element(i).Face(j).Code.R = 255

                    Case 2

                        HT_element(i).Face(j).Code.G = 255

                    Case 3

                        HT_element(i).Face(j).Code.R = 255

                        HT_element(i).Face(j).Code.G = 255

                    Case 4

                        HT_element(i).Face(j).Code.B = 255

                        HT_element(i).Face(j).Code.G = 255

                    Case 5

                        HT_element(i).Face(j).Code.R = 255

                        HT_element(i).Face(j).Code.B = 255

                End Select

            Next j

        Next i

    End Sub


    ''' <summary>
    ''' Процедура определяет минимальную и максимальную температуры за весь расчет.
    ''' Заполняет поля Min_T и Max_T
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Detect_Min_Max_T()

        Min_T = 0

        Max_T = 0

        If HT_element Is Nothing Then Exit Sub

        If HT_element(0).HT_Step Is Nothing Then Exit Sub

        ' Определяем максимальную и минимальную температуры для шага по времени

        Dim T As Double

        Min_T = Double.MaxValue

        Max_T = Double.MinValue

        For i As Integer = 0 To UBound(HT_element)

            For j As Integer = 0 To UBound(HT_element(i).HT_Step)

                T = HT_element(i).HT_Step(j).T

                If T < Min_T Then Min_T = T

                If T > Max_T Then Max_T = T

            Next j

        Next i

    End Sub

    ''' <summary>
    ''' Процедура определяет минимальную и максимальную температуры на текущем шаге.
    ''' Заполняет поля Min_T и Max_T
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Detect_Min_Max_T_on_current_step()

        Min_T = 0

        Max_T = 0

        If HT_element Is Nothing Then Exit Sub

        If HT_element(0).HT_Step Is Nothing Then Exit Sub

        ' Определяем максимальную и минимальную температуры для шага по времени

        Dim T As Double

        Min_T = Double.MaxValue

        Max_T = Double.MinValue

        For i As Integer = 0 To UBound(HT_element)

            T = HT_element(i).HT_Step(Me.Current_step).T

            If T < Min_T Then Min_T = T

            If T > Max_T Then Max_T = T

        Next i

        If Min_T = Max_T Then Detect_Min_Max_T()

    End Sub

    Public Sub Create_value_strip(ByRef view As Visualisation.Visualisation, ByRef N_values As Integer, ByRef Min_Value As Double, ByRef Max_Value As Double, ByRef Title As String)

        Dim Value_strip_RGB_code(N_values - 1) As Visualisation.RGB_code

        Dim RGB_Value As New Face_code

        Dim Value_strip_value(N_values - 1) As String

        Dim Converter As New ACCEL


        Dim Value As Double

        view.frm.Value_strip_title = Title

        view.frm.Value_strip_intervals = N_values

        ReDim view.frm.Value_strip_RGB_code(N_values - 1)

        ReDim view.frm.Value_strip_value(N_values - 1)

        For i As Integer = 0 To N_values - 1

            Value = Min_Value + i / (N_values - 1) * (Max_Value - Min_Value)

            RGB_Value = Field_color(Value, Min_Value, Max_Value)

            view.frm.Value_strip_RGB_code(i) = New Visualisation.RGB_code

            view.frm.Value_strip_RGB_code(i).R = RGB_Value.R

            view.frm.Value_strip_RGB_code(i).G = RGB_Value.G

            view.frm.Value_strip_RGB_code(i).B = RGB_Value.B

            view.frm.Value_strip_value(i) = Converter.my_Str(Value)

            'Value_strip_RGB_code(i) = New Visualisation.RGB_code

            'Value_strip_RGB_code(i).R = RGB_Value.R

            'Value_strip_RGB_code(i).G = RGB_Value.G

            'Value_strip_RGB_code(i).B = RGB_Value.B

            'Value_strip_value(i) = Converter.my_Str(Value)

        Next i

    End Sub

    ''' <summary>
    ''' This function return faces' color code which describes field's value, taking in consideration fields's minimum and maximum field's values
    ''' </summary>
    ''' <param name="Value">Field's value. Color codes is calculated for this value.</param>
    ''' <param name="Min_Value">Field's minimum value.</param>
    ''' <param name="Max_Value">Field's maximum value.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Field_color(ByVal Value As Double, ByVal Min_Value As Double, ByVal Max_Value As Double) As Face_code

        Dim V_0_Blue, V_1_Blue, V_2_Blue, V_3_Blue As Double

        Dim V_0_Green, V_1_Green, V_2_Green, V_3_Green As Double

        Dim V_0_Red, V_1_Red, V_2_Red, V_3_Red As Double

        Dim I_Blue, I_Green, I_Red As Double


        Const Max_I_Blue = 0.5

        Const Min_I_Blue = 0.3


        Const Max_I_Green = 0.7

        Const Min_I_Green = 0.0


        Const Max_I_Red = 1.0

        Const Min_I_Red = 0.0


        V_0_Blue = Min_Value

        V_1_Blue = Min_Value + 0.3 * (Max_Value - Min_Value)

        V_2_Blue = Min_Value + 0.3 * (Max_Value - Min_Value)

        V_3_Blue = Min_Value + 0.4 * (Max_Value - Min_Value)


        V_0_Green = Min_Value + 0.1 * (Max_Value - Min_Value)

        V_1_Green = Min_Value + 0.4 * (Max_Value - Min_Value)

        V_2_Green = Min_Value + 0.65 * (Max_Value - Min_Value)

        V_3_Green = Min_Value + 1 * (Max_Value - Min_Value)


        V_0_Red = Min_Value + 0.4 * (Max_Value - Min_Value)

        V_1_Red = Min_Value + 0.75 * (Max_Value - Min_Value)

        V_2_Red = Min_Value + 100 * (Max_Value - Min_Value)

        V_3_Red = Min_Value + 1000 * (Max_Value - Min_Value)



        I_Blue = 0

        If V_0_Blue <= Value And Value < V_1_Blue Then

            I_Blue = (Max_I_Blue - Min_I_Blue) / (V_1_Blue - V_0_Blue) * (Value - V_0_Blue) + Min_I_Blue

        End If

        If V_1_Blue <= Value And Value <= V_2_Blue Then

            I_Blue = Max_I_Blue

        End If

        If V_2_Blue <= Value And Value <= V_3_Blue Then

            I_Blue = -Max_I_Blue / (V_3_Blue - V_2_Blue) * (Value - V_2_Blue) + Max_I_Blue

        End If


        I_Green = 0

        If V_0_Green <= Value And Value < V_1_Green Then

            I_Green = (Max_I_Green - Min_I_Green) / (V_1_Green - V_0_Green) * (Value - V_0_Green) + Min_I_Green

        End If

        If V_1_Green <= Value And Value <= V_2_Green Then

            I_Green = Max_I_Green

        End If

        If V_2_Green <= Value And Value <= V_3_Green Then

            I_Green = (Min_I_Green - Max_I_Green) / (V_3_Green - V_2_Green) * (Value - V_2_Green) + Max_I_Green

        End If



        I_Red = 0

        If V_0_Red <= Value And Value < V_1_Red Then

            I_Red = (Max_I_Red - Min_I_Red) / (V_1_Red - V_0_Red) * (Value - V_0_Red) + Min_I_Red

        End If

        If V_1_Red <= Value And Value <= V_2_Red Then

            I_Red = Max_I_Red

        End If

        If V_2_Red <= Value And Value <= V_3_Red Then

            I_Red = (Min_I_Red - Max_I_Red) / (V_3_Red - V_2_Red) * (Value - V_2_Red) + Max_I_Red

        End If


        Field_color = New Face_code

        Field_color.B = I_Blue * 255

        Field_color.G = I_Green * 255

        Field_color.R = I_Red * 255

    End Function


    ''' <summary>
    ''' This function return faces' color code which describes field's value, taking in consideration fields's minimum and maximum field's values
    ''' </summary>
    ''' <param name="Value">Field's value. Color codes is calculated for this value.</param>
    ''' <param name="Min_Value">Field's minimum value.</param>
    ''' <param name="Max_Value">Field's maximum value.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Field_color_old(ByVal Value As Double, ByVal Min_Value As Double, ByVal Max_Value As Double) As Face_code

        Dim Value_Blue, Value_Green, Value_Red As Double

        Dim I_Blue, I_Green, I_Red As Double

        Const Min_I = 0.3

        Value_Blue = Min_Value + 1 / 3 * (Max_Value - Min_Value)

        Value_Green = Min_Value + 2 / 3 * (Max_Value - Min_Value)

        Value_Red = Max_Value

        If Min_Value <= Value And Value < Value_Blue Then

            I_Blue = (1 - Min_I) / (Value_Blue - Min_Value) * (Value - Min_Value) + Min_I

        Else

            I_Blue = 1

        End If

        I_Green = 0

        If Value_Blue <= Value And Value < Value_Green Then

            I_Green = 1 / (Value_Green - Value_Blue) * (Value - Value_Blue)

        End If

        If Value >= Value_Green Then

            I_Green = 1

        End If

        I_Red = 0

        If Value_Red <= Value And Value <= Max_Value Then

            I_Red = (1 - Min_I) / (Value_Red - Value_Green) * (Value - Value_Green)

        End If

        Field_color_old = New Face_code

        Field_color_old.B = I_Blue * 255

        Field_color_old.G = I_Green * 255

        Field_color_old.R = I_Red * 255

    End Function

    Public Sub Extract_result(ByRef T(,) As Double, ByRef Q(,) As Double)

        'T(i-й элемент, n-й шаг )

        ReDim T(UBound(HT_element), Lenght_time / Step_time)

        ReDim Q(UBound(HT_element), Lenght_time / Step_time)

        For i As Integer = 0 To UBound(HT_element)

            For n As Integer = 0 To UBound(HT_element(i).HT_Step)

                T(i, n) = HT_element(i).HT_Step(n).T

                Q(i, n) = HT_element(i).HT_Step(n).Q

            Next n

        Next i


    End Sub

    Public Sub Insert_result(ByRef T(,) As Double, ByRef Q(,) As Double)

        'T(i-й элемент, n-й шаг )

        For i As Integer = 0 To UBound(HT_element)

            ReDim HT_element(i).HT_Step(UBound(T, 2))

            For n As Integer = 0 To UBound(HT_element(i).HT_Step)

                HT_element(i).HT_Step(n).T = T(i, n)

                HT_element(i).HT_Step(n).Q = Q(i, n)

            Next n

        Next i

    End Sub

    ''' <summary>
    ''' This sub go trow results and leaves only 0..N..2N..3N results item.
    ''' It can be used for results file decrease.
    ''' Процедура прореживает результаты, оставляя каждый N-й. 
    ''' </summary>
    ''' <param name="N"></param>
    ''' <remarks></remarks>
    Public Sub Decimate_results(ByRef N As Integer)

        Dim i_Step As Integer

        Dim Stp() As HT_Step

        For i As Integer = 0 To UBound(HT_element)

            ReDim Stp(UBound(HT_element(i).HT_Step))

            i_Step = -1

            For j As Integer = 0 To UBound(HT_element(i).HT_Step) Step N

                i_Step += 1

                Stp(i_Step) = HT_element(i).HT_Step(j)

            Next j

            ReDim Preserve Stp(i_Step)

            HT_element(i).HT_Step = Stp

        Next i

        Main_form.Solver.Current_step = 0

        Main_form.Fill_Step_selector()


    End Sub


    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class

''' <summary>
''' Heat transfer element
''' </summary>
''' <remarks></remarks>
<Serializable()> _
Public Class HT_Element

    Public Number As Long

    ''' <summary>
    ''' FEM element
    ''' </summary>
    ''' <remarks></remarks>
    Public Element As Object


    ''' <summary>
    ''' Shining faces of heat transfer element
    ''' </summary>
    ''' <remarks></remarks>
    Public Face() As Object

    ''' <summary>
    ''' Array of element's heat conductivity acts.
    ''' For exmaple, if element has 4 neighbours, 
    ''' then Heat_conductivity array for this element will fill with 4 items.
    ''' </summary>
    ''' <remarks></remarks>
    Public Heat_conductivity() As Heat_conductivity


    ''' <summary>
    ''' Array of heat sources, heating element
    ''' </summary>
    ''' <remarks></remarks>
    Public Heat_source() As HSOURCE1

    ''' <summary>
    ''' Array of radiation sources, heating element
    ''' </summary>
    ''' <remarks></remarks>
    Public Radiation_source() As RSOURCE1


    ''' <summary>
    ''' T_SET, setted the element's temperature
    ''' </summary>
    ''' <remarks></remarks>
    Public T_SET As T_SET

    ''' <summary>
    ''' Full heat to/from element
    ''' </summary>
    ''' <remarks></remarks>
    Public Q_summary As Double

    ''' <summary>
    ''' The steps of problem solution
    ''' </summary>
    ''' <remarks></remarks>
    Public HT_Step() As HT_Step

    ''' <summary>
    ''' Table of parameters (vector-Argument components)
    ''' 0 - temperature,
    ''' 1 - time,
    ''' 2 - random value,
    ''' 3 - incoming radiation wavelenght
    ''' 4 - incoming radiation angle of elevation in local coordinate system
    ''' 5 - incoming radiation azimuth in local coordinate system
    ''' </summary>
    ''' <remarks></remarks>
    Public Argument As Vector

    Public Sub New()

        Argument = New Vector()

        ReDim Argument.Coord(5)

    End Sub
End Class

<Serializable()> _
Public Class Heat_conductivity

    ''' <summary>
    ''' The element, that involve in heat transfering
    ''' </summary>
    ''' <remarks></remarks>
    Public HT_Element As HT_Element

    ''' <summary>
    ''' The area, that transmiss heat flow 
    ''' </summary>
    ''' <remarks></remarks>
    Public Area As Double

    ''' <summary>
    ''' The distance from this element to boundary 
    ''' </summary>
    ''' <remarks></remarks>
    Public Distance_1 As Double

    ''' <summary>
    ''' The distance from boundary to neighbouring element
    ''' </summary>
    ''' <remarks></remarks>
    Public Distance_2 As Double

    '''' <summary>
    '''' The radius-vector between two elements
    '''' </summary>
    '''' <remarks></remarks>
    'Public Radius_vector As Vector

End Class

''' <summary>
''' This class contain information about one step of solution solving 
''' </summary>
''' <remarks></remarks>
<Serializable()> _
Public Structure HT_Step

    ''' <summary>
    ''' The moment of time when step was done
    ''' </summary>
    ''' <remarks></remarks>
    Public Time As Double

    ''' <summary>
    ''' Element's temperature
    ''' </summary>
    ''' <remarks></remarks>
    Public T As Double

    ''' <summary>
    ''' Element's incoming heat power (if >0, then power incoming)
    ''' </summary>
    ''' <remarks></remarks>
    Public Q As Double

    ''' <summary>
    ''' Codes of colors represented element's faces in visualisation
    ''' </summary>
    ''' <remarks></remarks>
    Public Code As Face_code

End Structure

<Serializable()> _
Public Structure Element_result

    Public HT_step() As HT_Step

End Structure

<Serializable()> _
Public Structure Total_result

    Public Element_result() As Element_result

End Structure

