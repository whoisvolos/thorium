''' <summary>
''' FEM any object (i.e. any finite element, coordinate system, material, property or something else)
''' </summary>
''' <remarks></remarks>
<Serializable()> _
Public MustInherit Class TOP_item

    Inherits Linear_algebra

    ''' <summary>
    ''' Name of model item. The number identification usually used, but in some reasons name can be more useful.
    ''' </summary>
    ''' <remarks></remarks>
    Public Name As String

    ''' <summary>
    ''' Type of model item. For instance, type can be used for subdividing materials to MAT1, MAT2, MAT8 and etc.
    ''' </summary>
    ''' <remarks></remarks>
    Public Type As String

    ''' <summary>
    ''' Number of model item. The common used identificator.
    ''' </summary>
    ''' <remarks></remarks>
    Public Number As Long

    ''' <summary>
    ''' The string, which represent model item in file.
    ''' </summary>
    ''' <remarks></remarks>
    Public String_in_file As String

    ''' <summary>
    ''' Array of fields, which represent model item in file.
    ''' </summary>
    ''' <remarks></remarks>
    Public Fields_array() As String


    ''' <summary>
    ''' Is model item used or not?
    ''' </summary>
    ''' <remarks></remarks>
    Public Enabled As Boolean


    '''' <summary>
    '''' The collection for some special applications. For example, here allowables can be stored.
    '''' </summary>
    '''' <remarks></remarks>
    'Public Special_collection As New Collection

    ''' <summary>
    ''' Model itemes, which were created with this model item
    ''' </summary>
    ''' <remarks></remarks>
    Public Children() As Object

    ''' <summary>
    ''' Plank constant
    ''' </summary>
    ''' <remarks></remarks>
    Public Const h = 6.626196E-34

    ''' <summary>
    ''' Speed of light
    ''' </summary>
    ''' <remarks></remarks>
    Public Const c = 299792458.0

    ''' <summary>
    ''' Boltzmann constant
    ''' </summary>
    ''' <remarks></remarks>
    Public Const k = 1.380623E-23

    ''' <summary>
    ''' Stefan-Boltzmann law constant
    ''' </summary>
    ''' <remarks></remarks>
    Public Const sigma = 2 * Math.PI ^ 5 * k ^ 4 / (15 * c ^ 2 * h ^ 3)


    ''' <summary>
    ''' This constant is used for maximum emittance of black body frequency (v_max) calculation.
    ''' v_max = v_max_const*T
    ''' For example, for T=100  v_max = v_max_const*100 = 5.878E+12
    ''' </summary>
    ''' <remarks></remarks>
    Public Const v_max_const = 58777311000.0


    ''' <summary>
    ''' Sub can be used for uploading model item fields from array which represent model item in file.
    ''' String_in_file field used as an information source.
    ''' </summary>
    ''' <remarks></remarks>
    Public MustOverride Sub Load()

    ''' <summary>
    ''' Sub can be used for preparing array of model item fields.
    ''' Fields used as an information source.
    ''' </summary>
    ''' <remarks></remarks>
    Public MustOverride Sub Prepare_Save()

    ''' <summary>
    ''' Sub can be used for downloading model item fields to string which represent model item in file.
    ''' String_in_file field used as an information destination.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Save()

        Prepare_Save()

        String_in_file = ""

        Dim Out_array() As String = Convert_array_to_NASTRAN_small_field_format_string(Fields_array)

        For i As Integer = 0 To UBound(Out_array)

            String_in_file = String_in_file & Out_array(i) & vbCrLf

        Next i

    End Sub

    ' функция вставляет в число, записанное строкой, символ "E"
    ' это нужно для того, чтобы Visual Basic правильно распознал порядок числа, записанного в эскпоненциальной форме
    Public Function Insert_E(ByRef Number As String)

        Dim Plus_position, Minus_position, Sign_position As Long

        Number = Trim(Number)

        Plus_position = InStrRev(Number, "+")
        Minus_position = InStrRev(Number, "-")

        If InStr(Number, ".") = 0 Then

            Insert_E = Number

            Exit Function

        End If


        If InStr(Number, "E") > 0 Then

            Insert_E = Number

            Exit Function

        End If


        If Plus_position > Minus_position Then

            Sign_position = Plus_position

        Else

            Sign_position = Minus_position

        End If

        If Sign_position < 2 Then

            Insert_E = Number

        Else

            Insert_E = Mid(Number, 1, Sign_position - 1) & "E" & Mid(Number, Sign_position)

        End If



    End Function
    ''' <summary>
    ''' Function for conversion element in file representation string to array of element's fields
    ''' </summary>
    ''' <param name="in_String">Element in file representation string</param>
    ''' <returns>Array of element's fields</returns>
    ''' <remarks></remarks>
    Public Function Convert_string_in_file_to_array_1(ByRef in_String As String) As String()

        Dim Position, old_Position, i, j, Record_length As Integer

        Dim tmp_String_0, tmp_Output_string() As String

        Position = 1

        'i = 1

        ReDim tmp_Output_string(0)

        Dim Output_collection As New Collection

        Do ' цикл по сторочкам

            old_Position = Position

            Position = InStr(Position, in_String, Chr(13))

            If Position = 0 Then Exit Do

            ' это одна строчка
            tmp_String_0 = Mid(in_String, old_Position, Position - old_Position + 1)

            tmp_String_0 = Replace(tmp_String_0, Chr(10), "")

            tmp_String_0 = Replace(tmp_String_0, Chr(13), "")

            Position = Position + 1

            ' Записываем тип
            Output_collection.Add(Trim(Mid(tmp_String_0, 1, 8)))

            If InStr(Mid(tmp_String_0, 1, 8), "*") > 0 Then

                Record_length = 16

            Else

                Record_length = 8

            End If

            For j = 9 To 80 Step Record_length ' цикл по полям в строчке

                ' это одно поле
                Output_collection.Add(Insert_E(Mid(tmp_String_0, j, Record_length)))

            Next j

        Loop

        ' преобразуем строки в малый формат полей, то есть в одной строчке должно быть 8 значащих полей, 
        ' а не 4 как в длинном формате.
        ' удаляем из строки * и пустое поле перед ними 

        ' Цикл по строчкам

        i = 1 ' позиция начала строчки

        Do

            If i > Output_collection.Count Then Exit Do

            If InStr(Output_collection.Item(i).ToString, "*") > 0 And i + 12 <= Output_collection.Count Then

                Output_collection.Remove(i + 5)

                Output_collection.Remove(i + 5)

                If InStr(Output_collection.Item(i + 10).ToString, "*") = 0 Then

                    Output_collection.Add("", , i + 10)

                    Output_collection.Add("*", , i + 10)

                End If

            End If

            If InStr(Output_collection.Item(i).ToString, "*") > 0 And i + 12 > Output_collection.Count Then

                If i + 5 <= Output_collection.Count Then

                    Output_collection.Remove(i + 5)

                    Output_collection.Remove(i + 5)

                End If

            End If

            i += 10

        Loop

        ReDim tmp_Output_string(Output_collection.Count - 1)

        For i = 0 To UBound(tmp_Output_string)

            tmp_Output_string(i) = Output_collection.Item(i + 1).ToString

        Next i

        Convert_string_in_file_to_array_1 = tmp_Output_string

    End Function

    ''' <summary>
    ''' Function for conversion element in file representation string to array of element's fields
    ''' </summary>
    ''' <param name="in_String">Element in file representation string</param>
    ''' <returns>Array of element's fields</returns>
    ''' <remarks></remarks>
    Public Function Convert_string_in_file_to_array_2(ByRef in_String As String) As String()

        Dim Position, old_Position, i, j, Record_length As Integer

        Dim tmp_String_0, tmp_Output_string(), str_Type As String

        Position = 1

        'i = 1

        ReDim tmp_Output_string(0)

        Dim Output_string() As String


        ReDim Output_string(Len(in_String))

        ' Записываем тип
        str_Type = Trim(Mid(in_String, 1, 8))

        i = -1

        Do ' цикл по сторочкам

            old_Position = Position

            Position = InStr(Position, in_String, Chr(13))

            If Position = 0 Then Exit Do

            ' это одна строчка
            tmp_String_0 = Mid(in_String, old_Position, Position - old_Position + 1)

            tmp_String_0 = Replace(tmp_String_0, Chr(10), "")

            tmp_String_0 = Replace(tmp_String_0, Chr(13), "")

            Position = Position + 1

            If InStr(Mid(tmp_String_0, 1, 8), "*") > 0 Then

                Record_length = 16

            Else

                Record_length = 8

            End If

            For j = 9 To 72 Step Record_length ' цикл по полям в строчке

                i += 1

                ' это одно поле
                Output_string(i) = Insert_E(Mid(tmp_String_0, j, Record_length))

            Next j



        Loop

        ' удаляем лишние поля

        ReDim Preserve Output_string(i)

        ReDim tmp_Output_string(((UBound(Output_string)) / 8) * 2 + UBound(Output_string))

        Dim i_Output_string As Integer = -1

        For i = 0 To (UBound(Output_string) + 1) / 8

            For j = 1 To 8

                i_Output_string += 1

                If i_Output_string > UBound(Output_string) Then GoTo Exit_fors

                tmp_Output_string(i * 10 + j) = Output_string(i_Output_string)

            Next j

        Next i

Exit_fors:

        ' преобразуем строки в малый формат полей, то есть в одной строчке должно быть 8 значащих полей, 
        ' а не 4 как в длинном формате.
        ' удаляем из строки * и пустое поле перед ними 

        tmp_Output_string(0) = str_Type

        Convert_string_in_file_to_array_2 = tmp_Output_string

    End Function

    ''' <summary>
    ''' Function for conversion element in file representation string to array of element's fields
    ''' </summary>
    ''' <param name="in_String">Element in file representation string</param>
    ''' <returns>Array of element's fields</returns>
    ''' <remarks></remarks>
    Public Function Convert_string_in_file_to_array_3(ByRef in_String As String) As String()

        Dim First_character_code As Integer

        Dim Position, old_Position, i, j, Record_length As Integer

        Dim tmp_String_0, tmp_Output_string(), str_Type As String

        Position = InStr(in_String, "BEGIN BULK") + Len("BEGIN BULK")

        ReDim tmp_Output_string(0)

        Dim Output_string() As String


        ReDim Output_string(Len(in_String))


        i = -1

        Do ' цикл по сторочкам

            old_Position = Position

            Position = InStr(Position, in_String, Chr(13))

            If Position = 0 Then Exit Do

            ' это одна строчка
            tmp_String_0 = Mid(in_String, old_Position, Position - old_Position + 1)

            tmp_String_0 = Replace(tmp_String_0, Chr(10), "")

            tmp_String_0 = Replace(tmp_String_0, Chr(13), "")

            Position = Position + 1

            str_Type = Mid(tmp_String_0, 1, 8)

            If str_Type <> "" Then

                First_character_code = Asc(Mid(str_Type, 1, 1))

                If First_character_code >= 65 And First_character_code <= 90 Or First_character_code = 32 Or First_character_code = 42 Or First_character_code = 43 Then ' Or First_character_code = 36

                    If InStr(str_Type, "*") > 0 Then

                        Record_length = 16

                        i += 1

                        ' это поле первое, которое всегда 8 символов
                        Output_string(i) = Insert_E(str_Type)

                        For j = 9 To 72 Step Record_length ' цикл по полям в строчке

                            i += 1

                            ' это одно поле
                            Output_string(i) = Insert_E(Mid(tmp_String_0, j, Record_length))

                        Next j

                        old_Position = Position

                        Position = InStr(Position, in_String, Chr(13))

                        If Position = 0 Then Exit Do

                        ' это одна строчка
                        tmp_String_0 = Mid(in_String, old_Position, Position - old_Position + 1)

                        tmp_String_0 = Replace(tmp_String_0, Chr(10), "")

                        tmp_String_0 = Replace(tmp_String_0, Chr(13), "")

                        Position = Position + 1


                        For j = 9 To 72 Step Record_length ' цикл по полям в строчке

                            i += 1

                            ' это одно поле
                            Output_string(i) = Insert_E(Mid(tmp_String_0, j, Record_length))

                        Next j


                    Else

                        Record_length = 8

                        i += 1

                        ' это поле первое, которое всегда 8 символов
                        Output_string(i) = Insert_E(str_Type)

                        For j = 9 To 72 Step Record_length ' цикл по полям в строчке

                            i += 1

                            ' это одно поле
                            Output_string(i) = Insert_E(Mid(tmp_String_0, j, Record_length))

                        Next j

                        i += 1

                        ' это последнее поле, которое всегда 8 символов
                        Output_string(i) = Insert_E(Mid(tmp_String_0, 73, 8))

                    End If

                End If

            End If




        Loop

        ' удаляем лишние поля

        ReDim Preserve Output_string(i)

        For i = 0 To UBound(Output_string)



        Next i

        Convert_string_in_file_to_array_3 = Output_string

    End Function


    ' Функция получает на входе массив, который предразует в несколько строк малого формата NASTRAN
    ' В каждой сторочке 10 позиций по 8 символов
    ' Для больших карт NASTRAN, там где используется несколько строк.
    Public Function Convert_array_to_NASTRAN_small_field_format_string(ByRef In_array() As String) As String()

        Dim tmp_Array() As String

        Dim Out_array() As String

        ReDim Out_array(0)

        Dim Edge, i_tmp As Long

        For i As Long = 0 To UBound(In_array) Step 10

            ReDim tmp_Array(10)

            For j As Integer = 0 To UBound(tmp_Array)

                tmp_Array(j) = " "

            Next j

            Edge = i + 9

            If Edge > UBound(In_array) Then Edge = UBound(In_array)

            i_tmp = 0

            For j As Long = i To Edge

                tmp_Array(i_tmp) = In_array(j)

                i_tmp = i_tmp + 1

            Next j

            Out_array(UBound(Out_array)) = Convert_array_to_small_field_format_string(tmp_Array)

            ReDim Preserve Out_array(UBound(Out_array) + 1)

        Next i

        ReDim Preserve Out_array(UBound(Out_array) - 1)

        Return Out_array

    End Function

    ' Функция получает на входе массив, который предразует в строчку малого формата NASTRAN
    ' В сторочке 10 позиций по 8 символов
    Public Function Convert_array_to_small_field_format_string(ByRef One_array() As String) As String

        Dim X As Long

        Convert_array_to_small_field_format_string = ""

        For i As Integer = 1 To 80

            Convert_array_to_small_field_format_string = Convert_array_to_small_field_format_string & " "

        Next i

        For i As Long = 0 To UBound(One_array) - 1

            X = i * 8 + 1

            If Len(One_array(i)) > 8 Then Call Err.Raise(1, "", "Convert_array_to_small_field_format_string - Array element length is more whan 8 symbols")

            Mid(Convert_array_to_small_field_format_string, X, Len(One_array(i))) = One_array(i)

        Next i

        Convert_array_to_small_field_format_string = Convert_array_to_small_field_format_string

    End Function
    Public Function my_Val(ByRef A As String) As Object

        If InStr(A, ".") = 0 And InStr(A, "E") = 0 And InStr(A, "e") = 0 Then

            my_Val = CLng(Val(A))

        Else

            my_Val = CDbl(Val(A))

        End If

    End Function



    ' Функция переводит число в малый формат NASTRAN, страясь впихнуть в 8 знаков как можно больше точности
    Public Function my_Str(ByRef A As Double) As String

        my_Str = ""

        If A = 0 Then

            my_Str = "0."

            Exit Function

        End If

        my_Str = My_Str_Conversion(A)

    End Function

    ' Функция переводит число в малый формат NASTRAN, страясь впихнуть в 8 знаков как можно больше точности
    Public Function my_Str(ByRef A As Long) As String

        my_Str = ""

        If A = 0 Then

            my_Str = "0"

            Exit Function

        End If

        'my_Str = Trim(Str(A))

        my_Str = A

    End Function

    ' Функция переводит число в малый формат NASTRAN, страясь впихнуть в 8 знаков как можно больше точности
    Public Function my_Str(ByRef A As String) As String

        my_Str = A

    End Function

    Public Function my_Str(ByRef A As Object) As String

        On Error Resume Next

        Dim tmp_Number As Long = Long.MinValue

        tmp_Number = A.Number

        If tmp_Number = Long.MinValue Then

            If CLng(A) = 0 Then

                my_Str = ""

            Else

                my_Str = my_Str(CLng(A))

            End If

        Else

            my_Str = my_Str(tmp_Number)

        End If

        On Error GoTo 0

    End Function

    Function My_Str_Conversion(ByRef A As Double)

        Dim Var_1, Var_2 As String

        Var_1 = ""

        Var_2 = ""

        Dim Eps_1, Eps_2 As Double

        Eps_1 = Double.MaxValue

        Eps_2 = Double.MaxValue

        Dim Exp As Double

        If A < 0 Then

            Exp = Math.Log10(Math.Abs(A))

        Else

            Exp = Math.Log10(A)

        End If

        Dim Mant As Double

        Exp = Math.Floor(Exp)

        Mant = Math.Round((A / 10 ^ (Exp)), 8)

        Dim Int_mant As Integer

        Dim Real_mant As Integer

        Int_mant = Math.Floor(Mant)

        If Int_mant < 0 Then Int_mant += 1

        Dim Round_to As Integer = 6 - Trim(Str(Int_mant)).Length - Trim(Str(Exp)).Length

        'If A < 0 Then Round_to -= 1

        Mant = Math.Round(Mant, Round_to)

        Real_mant = Math.Abs((Mant - Int_mant) * 10 ^ (Round_to))

        If Round_to - Real_mant.ToString.Length > -1 Then

            Var_1 = Int_mant & "." & StrDup(Round_to - Real_mant.ToString.Length, "0") & Real_mant & "E" & Exp.ToString

            Eps_1 = Math.Abs(A - Val(Var_1))

        End If

        If A >= 0.000001 And A <= 99999999 Then

            Var_2 = Math.Round(A, Round_to).ToString.Replace(",", ".")

            Eps_2 = Math.Abs(A - Val(Var_2))

        End If

        If A <= -0.00001 And A >= -9999999 Then

            Var_2 = Math.Round(A, Round_to).ToString.Replace(",", ".")

            Eps_2 = Math.Abs(A - Val(Var_2))

        End If

        If Eps_1 < Eps_2 Then

            If Eps_1 > Math.Abs(A * 0.01) Then Throw New Exception("Error in My_Str_Conversion. " & Str(A) & " is converted to " & Var_1)

            Return Var_1

        Else

            If Eps_2 > Math.Abs(A * 0.01) Then Throw New Exception("Error in My_Str_Conversion. " & Str(A) & " is converted to " & Var_2)

            Return Var_2

        End If

        'Dim Log_a, Mant_a As Double

        'Dim Exp_a As Integer

        'Log_a = Math.Log(Math.Abs(A), 10)

        'Exp_a = Math.Round(Log_a, 0) + 1

        'Mant_a = 10 ^ (Log_a - Exp_a)


        'If A > 0.000001 And A < 9999999 Then

        '    If Exp_a > 7 Then Exp_a = 7

        '    If Exp_a < 0 Then Exp_a = 0

        '    'My_Str_Conversion = Trim(Str(Math.Round(A, 7 - Exp_a)))
        '    My_Str_Conversion = Math.Round(A, 7 - Exp_a)

        '    If A = Math.Round(A, 0) Then

        '        My_Str_Conversion = My_Str_Conversion & "."

        '    End If

        '    If A > 0 And A < 1 Then

        '        My_Str_Conversion = Mid(My_Str_Conversion, 2).Replace(",", ".")

        '    End If

        '    If Mid(My_Str_Conversion, 1, 1) = "E" Then

        '        My_Str_Conversion = "1" & My_Str_Conversion

        '    End If

        '    If Not (My_Str_Conversion.GetType.ToString = "System.String") Then My_Str_Conversion = Str(My_Str_Conversion)

        '    My_Str_Conversion = My_Str_Conversion.Replace(",", ".")

        '    My_Str_Conversion = My_Str_Conversion.Replace(" ", "")

        '    Exit Function

        'End If

        'If A < -0.00001 And A > -999999 Then

        '    If Exp_a > 6 Then Exp_a = 6

        '    If Exp_a < 0 Then Exp_a = 0

        '    'My_Str_Conversion = Trim(Str(Math.Round(A, 6 - Exp_a)))
        '    My_Str_Conversion = Math.Round(A, 6 - Exp_a)

        '    If A = Math.Round(A, 0) Then

        '        My_Str_Conversion = My_Str_Conversion & "."

        '    End If

        '    If A < 0 And A > -1 Then

        '        My_Str_Conversion = "-." & Mid(My_Str_Conversion, 4)

        '    End If

        '    If Not (My_Str_Conversion.GetType.ToString = "System.String") Then My_Str_Conversion = Str(My_Str_Conversion)

        '    My_Str_Conversion = My_Str_Conversion.Replace(",", ".")

        '    Exit Function


        'End If


        'Log_a = Math.Log(Math.Abs(A), 10)

        'Exp_a = Math.Round(Log_a, 0)

        'If Exp_a > Log_a Then Exp_a -= 1

        'Mant_a = 10 ^ (Log_a - Exp_a)

        'Mant_a = Math.Round(Mant_a, 2)

        'My_Str_Conversion = ""

        'Select Case Math.Sign(A)

        '    Case 1

        '        My_Str_Conversion = My_Str_Conversion

        '    Case 0

        '        My_Str_Conversion = My_Str_Conversion

        '    Case -1

        '        My_Str_Conversion = "-"

        'End Select

        ''My_Str_Conversion = My_Str_Conversion & Trim(Str(Mant_a)) & "E" & Exp_a
        'My_Str_Conversion = My_Str_Conversion & Mant_a & "E" & Exp_a

        'My_Str_Conversion = My_Str_Conversion.Replace(",", ".")


    End Function

    Public Sub New(Optional ByRef String_in_file As String = "")

        Enabled = True

        Name = ""

        Number = 0

        Type = UCase(Mid(ToString(), InStr(ToString(), ".") + 1))


        If String_in_file <> "" Then

            Fields_array = Convert_string_in_file_to_array_2(String_in_file)

            If Fields_array(0) <> Type Then

                Fields_array(0) = Replace(Fields_array(0), "*", "")

                If Fields_array(0) <> Type Then

                    Exit Sub 'MsgBox("Element type " & Type & " and passed string " & String_in_file & "are not compared." & Chr(10) & Chr(13) & "This item loading is terminated.")

                End If

            Else

                Me.String_in_file = String_in_file

                Load()

            End If

        End If

    End Sub

End Class

<Serializable()> _
Public Class Function_point

    Public Argument As Vector

    Public Value As Vector

    Public Sub New(ByRef N_Argument As Long, ByRef N_value As Long)

        Argument = New Vector

        Value = New Vector

        ReDim Argument.Coord(N_Argument)

        ReDim Value.Coord(N_value)

    End Sub
End Class

''' <summary>
''' This class is for dependencies definition, i.e. reflectance from radiation wavelength and from temperature,
''' or heat flux vector from radii-vector and etc.
''' Items can depends on 6 parameters, compacted in 1 vector-Argument.
''' Table of parameters (vector-Argument components)
''' 0 - temperature,
''' 1 - time,
''' 2 - random value,
''' 3 - incoming radiation frequency,
''' 4 - incoming radiation angle of elevation in local coordinate system,
''' 5 - incoming radiation azimuth in local coordinate system.
''' </summary>
''' <remarks></remarks>
<Serializable()> _
Public Class DEPEND

    Inherits TOP_item

    Public Point() As Function_point

    ''' <summary>
    ''' If any component of the Aggument vector is equal to zero in all of DEPEND Points,
    ''' then this component should be excluded from distance calculation.
    ''' For example, if value do not depends on time then Base(1)=0,
    ''' if value do not depends on temperature then Base(0)=0, and so on.
    ''' If value DEPENDS on some component, the Base(i)=1/(Component_max_value - Component_min_value)^2,
    ''' where
    ''' Component_max_value - the largest component value in all of given points,
    ''' Component_min_value - the smallest component value in all of given points.
    ''' </summary>
    ''' <remarks></remarks>
    Public Base As Vector

    ''' <summary>
    ''' This coefficient is used in inverce distance (ID) calculation: ID = 1/(L)^n,
    ''' where L is distance from tabulated point to given point.
    ''' </summary>
    ''' <remarks></remarks>
    Public n As Double = 2

    Public Overrides Sub Load()

        Dim i_Point As Integer = -1

        Dim i_Argument As Integer = -1

        Dim i_Value As Integer = -1

        Dim N_Argument As Integer = -1

        Dim N_value As Integer = -1


        Number = Val(Fields_array(1))

        n = Val(Fields_array(2))

        '' Detecting Argument vector dimension

        'For i As Integer = 11 To 19

        '    If Fields_array(i) <> "" And Not (Fields_array(i) Is Nothing) Then

        '        N_Argument += 1

        '    End If

        'Next i

        N_Argument = 5

        ' Detecting value vector dimension

        For i As Integer = 21 To 29

            If Fields_array(i) <> "" And Not (Fields_array(i) Is Nothing) Then

                N_value += 1

            End If

        Next i

        For i As Integer = 11 To UBound(Fields_array) Step 20

            If Fields_array(i) <> "" And Not (Fields_array(i) Is Nothing) Then
                'If Not (Fields_array(i) Is Nothing) Then

                i_Argument = -1

                i_Value = -1

                i_Point += 1

                ReDim Preserve Point(i_Point)

                Point(i_Point) = New Function_point(N_Argument, N_value)

                For j As Integer = i To i + 7

                    If j > UBound(Fields_array) Then Exit For

                    If Fields_array(j) <> "" And Not (Fields_array(j) Is Nothing) Then

                        i_Argument += 1

                        Point(i_Point).Argument.Coord(i_Argument) = Val(Fields_array(j))

                    End If

                Next j

                For j As Integer = i + 10 To i + 7 + 10

                    If j > UBound(Fields_array) Then Exit For

                    If Fields_array(j) <> "" And Not (Fields_array(j) Is Nothing) Then

                        i_Value += 1

                        Point(i_Point).Value.Coord(i_Value) = Val(Fields_array(j))

                    End If

                Next j

            End If

        Next i

    End Sub

    Public Overrides Sub Prepare_Save()


        ReDim Fields_array(10 + (UBound(Point) + 1) * 20)

        Fields_array(0) = Type

        Fields_array(1) = my_Str(Number)

        Fields_array(2) = my_Str(n)

        For i As Integer = 0 To UBound(Point)

            For j As Integer = 0 To UBound(Point(i).Argument.Coord)

                Fields_array(20 * (i) + 1 + j + 10) = my_Str(Point(i).Argument.Coord(j))

            Next j

            For j As Integer = 0 To UBound(Point(i).Value.Coord)

                Fields_array(20 * (i) + 1 + j + 20) = my_Str(Point(i).Value.Coord(j))
            Next j


        Next i


    End Sub
    Public Sub New(Optional ByRef String_in_file As String = "")

        MyBase.New(String_in_file)

    End Sub

    Public Sub New(ByRef Fields_array() As String)

        Me.Fields_array = Fields_array

        Load()

        Calc_base_vector()

    End Sub

    Public Sub Calc_base_vector()

        Dim Comp_max, Comp_min As Double

        Base = New Vector

        ReDim Base.Coord(5)

        For i As Integer = 0 To 5

            Base.Coord(i) = 0

            Comp_max = Double.MinValue

            Comp_min = Double.MaxValue

            For j As Integer = 0 To UBound(Point)

                If Point(j).Argument.Coord(i) <> 0 Then

                    Base.Coord(i) = 1

                    If Point(j).Argument.Coord(i) > Comp_max Then

                        Comp_max = Point(j).Argument.Coord(i)

                    End If

                    If Point(j).Argument.Coord(i) < Comp_min Then

                        Comp_min = Point(j).Argument.Coord(i)

                    End If


                End If

            Next j

            If Base.Coord(i) = 1 Then

                Base.Coord(i) = 1 / (Comp_max - Comp_min) ^ 2

            End If

        Next i


    End Sub

    ''' <summary>
    ''' This functon return interpolated value of dependency, using inverse distance method with bases
    ''' Возвращает значение функции, заданной множеством точек, используя метод обратного расстояния,
    ''' при этом расстояние между точками вычислеятся с учетом размаха переменной.
    ''' Например, функция зависит от времени и от скорости и определена на интервалах [0 сек,1 сек] и [ 1 км/с , 10 км/с ]
    ''' Расстояние между точками вычисляется как L = math.sqrt( ( (t - t_i)/ (1 сек) )^2  + ((v - v_i)/(9 км/c))^2 )
    ''' Весовой коэффициент вычисляется как (1/(L + 1.0E-30))^n
    ''' </summary>
    ''' <param name="Argument"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Return_value(ByRef Argument As Vector) As Vector

        If UBound(Point) = 0 Then

            Return Point(0).Value

        End If

        Dim inverse_Based_L(UBound(Point)) As Double

        Dim Summa_inverse_Based_L As Double = 0

        For i As Integer = 0 To UBound(Point)

            inverse_Based_L(i) = 1 / (Based_L(Argument, Point(i).Argument) + 1.0E-30) ^ n

        Next i

        For i As Integer = 0 To UBound(Point)

            Summa_inverse_Based_L += inverse_Based_L(i)

        Next i

        Return_value = New Vector

        ReDim Return_value.Coord(UBound(Point(0).Value.Coord))

        ' цикл по всем компонентам выходного вектора
        For i As Integer = 0 To UBound(Point(0).Value.Coord)

            For j As Integer = 0 To UBound(Point)

                Return_value.Coord(i) += Point(j).Value.Coord(i) * inverse_Based_L(j)

            Next j

            Return_value.Coord(i) /= Summa_inverse_Based_L

        Next i

    End Function

    Public Function Based_L(ByRef Argument_1 As Vector, ByRef Argument_2 As Vector) As Double

        Based_L = 0

        For i As Integer = 0 To UBound(Argument_1.Coord)

            'Based_L += ((Argument_1.Coord(i) - Argument_2.Coord(i)) / Base.Coord(i)) ^ 2

            Based_L += ((Argument_1.Coord(i) - Argument_2.Coord(i))) ^ 2 * Base.Coord(i)

        Next i

        Based_L = Math.Sqrt(Based_L)

    End Function

End Class

''' <summary>
''' This class is for face's thermal-optical properties definition
''' </summary>
''' <remarks></remarks>
<Serializable()> _
Public Class FTOP

    Inherits TOP_item

    ''' <summary>
    ''' Diffuse reflectance
    ''' </summary>
    ''' <remarks></remarks>
    Public R_dif As Object

    ''' <summary>
    ''' Mirror reflectance
    ''' </summary>
    ''' <remarks></remarks>
    Public R_mir As Object

    ''' <summary>
    ''' Absorbtance
    ''' </summary>
    ''' <remarks></remarks>
    Public A As Object

    ''' <summary>
    ''' Transmittance
    ''' </summary>
    ''' <remarks></remarks>
    Public D As Object

    ''' <summary>
    ''' Emissivity. Степень черноты.
    ''' </summary>
    ''' <remarks></remarks>
    Public Eps As Object

    ''' <summary>
    ''' Cumulative distribution of photon frequency. Photons' spectrum calcalated on a basis of CDF.
    ''' </summary>
    ''' <remarks></remarks>
    Public CDF As Object

    ''' <summary>
    ''' Face effective temperature
    ''' </summary>
    ''' <remarks></remarks>
    Public T_eff As Object

    ''' <summary>
    ''' The integral emissivity.
    ''' The emissivity can depends on frequency of radiation, but output overal radiation flux do not depends on it.
    ''' So, we should integrate emissivity from zero frequency to infinity to gain integral emissivity 
    ''' for whole frequency band.
    ''' </summary>
    ''' <remarks></remarks>
    Public Integral_Eps_DEPEND As DEPEND

    ''' <summary>
    ''' The integral absorbtivity.
    ''' The absorbtivity can depends on frequency of radiation, but output overal absorbted radiation flux do not depends on it.
    ''' So, we should integrate absobtivity from zero frequency to infinity to gain integral absorbtivity 
    ''' for whole frequency band.
    ''' </summary>
    ''' <remarks></remarks>
    Public Integral_A_DEPEND As DEPEND



    Public Overrides Sub Load()


        Number = Val(Fields_array(1))

        R_dif = my_Val(Fields_array(2))

        R_mir = my_Val(Fields_array(3))

        A = my_Val(Fields_array(4))

        D = my_Val(Fields_array(5))

        Eps = my_Val(Fields_array(6))

        CDF = my_Val(Fields_array(7))

        T_eff = my_Val(Fields_array(8))

    End Sub

    Public Overrides Sub Prepare_Save()


        ReDim Fields_array(5)

        Fields_array(0) = Type

        Fields_array(1) = my_Str(Number)

        Fields_array(2) = my_Str(R_dif)

        Fields_array(3) = my_Str(R_mir)

        Fields_array(4) = my_Str(A)

        Fields_array(5) = my_Str(D)

        Fields_array(6) = my_Str(Eps)

        Fields_array(7) = my_Str(CDF)

        Fields_array(8) = my_Str(T_eff)


    End Sub
    Public Sub New(Optional ByRef String_in_file As String = "")

        MyBase.New(String_in_file)

    End Sub

    Public Sub New(ByRef Fields_array() As String)

        Me.Fields_array = Fields_array

        Load()

    End Sub

    '''' <summary>
    '''' This function calculates diffuse reflectance coefficient for a given Argument
    '''' </summary>
    '''' <param name="Argument">Given Argument</param>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public Function Calc_R_dif(ByRef Argument As Vector) As Double

    '    If R_dif.GetType.ToString <> "System.Double" Then

    '        Return R_dif.Return_value(Argument).Coord(0)

    '    End If

    '    If R_dif.GetType.ToString = "System.Double" Then

    '        Return R_dif

    '    End If

    'End Function
    '''' <summary>
    '''' This function calculates mirror reflectance coefficient for a given Argument
    '''' </summary>
    '''' <param name="Argument">Given Argument</param>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public Function Calc_R_mir(ByRef Argument As Vector) As Double

    '    If R_mir.GetType.ToString <> "System.Double" Then

    '        Return R_mir.Return_value(Argument).Coord(0)

    '    End If

    '    If R_mir.GetType.ToString = "System.Double" Then

    '        Return R_mir

    '    End If

    'End Function

    ''' <summary>
    ''' This function calculates absorbtion coefficient for a given Argument
    ''' </summary>
    ''' <param name="Argument">Given Argument</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Calc_A(ByRef Argument As Vector) As Double

        If A.GetType.ToString <> "System.Double" Then

            Return A.Return_value(Argument).Coord(0)

        End If

        If A.GetType.ToString = "System.Double" Then

            Return A

        End If

    End Function
    ''' <summary>
    ''' This function calculates integral emissivity for a given Argument.
    ''' The emissivity can depends on frequency of radiation, but output overal radiation flux do not depends on it.
    ''' So, we should integrate emissivity from zero frequency to infinity to gain integral emissivity 
    ''' for whole frequency band.
    ''' </summary>
    ''' <param name="Argument">Given Argument</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Calc_intergal_Eps(ByRef Argument As Vector) As Double

        Return Integral_Eps_DEPEND.Return_value(Argument).Coord(0)

    End Function

    ''' <summary>
    ''' The integral absorbtivity.
    ''' The absorbtivity can depends on frequency of radiation, but output overal absorbted radiation flux do not depends on it.
    ''' So, we should integrate absobtivity from zero frequency to infinity to gain integral absorbtivity 
    ''' for whole frequency band.
    ''' </summary>
    ''' <param name="Argument">Given Argument</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Calc_intergal_A(ByRef Argument As Vector) As Double

        Return Integral_A_DEPEND.Return_value(Argument).Coord(0)

    End Function

    ''' <summary>
    ''' This sub calculates integral emissivity values for temperature and time in initial Eps dependency.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Calc_Integral_Eps_DEPEND()

        ' проверяем, вдруг наша степень черноты - константа?

        If Eps.GetType.ToString = "System.Double" Then Exit Sub

        ' Инициализируем зависимость
        Integral_Eps_DEPEND = Calc_Integral_DEPEND(Eps, 10.0)


        '' Определяем координаты точек, в которых мы будем рассчитывать интегральную степень черноты

        'Dim Mode As Integer

        'Const NO_FREQ As Integer = 0, NO_TEMP_NO_TIME As Integer = 1, YES_TEMP_NO_TIME As Integer = 2
        'Const NO_TEMP_YES_TIME As Integer = 3, YES_TEMP_YES_TIME As Integer = 4


        '' Смотрим от чего зависит степень черноты

        'If Eps.Base.Coord(0) = 0 And Eps.Base.Coord(1) = 0 And Eps.Base.Coord(3) = 0 Then

        '    Mode = NO_FREQ

        '    Integral_Eps_DEPEND = Eps

        '    Exit Sub

        'End If

        'If Eps.Base.Coord(0) = 0 And Eps.Base.Coord(1) = 0 Then

        '    Mode = NO_TEMP_NO_TIME

        '    Throw New Exception("The dependency № " & Eps.number & "are is used for for emissivity calculation. This dependency is not depends on temperature or time. It is incorrect.")

        '    Exit Sub

        'End If

        'If Eps.Base.Coord(0) <> 0 And Eps.Base.Coord(1) = 0 Then

        '    Mode = YES_TEMP_NO_TIME

        'End If

        'If Eps.Base.Coord(0) = 0 And Eps.Base.Coord(1) <> 0 Then

        '    Mode = NO_TEMP_YES_TIME

        'End If

        'If Eps.Base.Coord(0) <> 0 And Eps.Base.Coord(1) <> 0 Then

        '    Mode = YES_TEMP_YES_TIME

        'End If

        '' в каком режиме понятно.
        '' Теперь надо определить границы того поля, на котором мы будем строить зависимость

        'Dim Temp_min, Temp_max, Time_min, Time_max As Double

        'Temp_min = Double.MaxValue

        'Temp_max = Double.MinValue

        'Time_min = Double.MaxValue

        'Time_max = Double.MinValue

        'For i As Integer = 0 To UBound(Eps.Point)

        '    If Eps.point(i).Argument.Coord(0) < Temp_min Then

        '        Temp_min = Eps.point(i).Argument.Coord(0)

        '    End If

        '    If Eps.point(i).Argument.Coord(0) > Temp_max Then

        '        Temp_max = Eps.point(i).Argument.Coord(0)

        '    End If



        '    If Eps.point(i).Argument.Coord(1) < Time_min Then

        '        Time_min = Eps.point(i).Argument.Coord(1)

        '    End If

        '    If Eps.point(i).Argument.Coord(1) > Time_max Then

        '        Time_max = Eps.point(i).Argument.Coord(1)

        '    End If

        'Next i

        '' определили поле, на котором будем работать.

        '' определям количество точек в зависимости интегральной степени черноты, которые мы хотим получить.
        'Dim N_point_side As Integer

        'Select Case Mode

        '    Case NO_TEMP_YES_TIME, YES_TEMP_NO_TIME

        '        N_point_side = UBound(Eps.Point) + 1

        '    Case YES_TEMP_YES_TIME

        '        N_point_side = (UBound(Eps.Point) + 1) ^ 0.5

        'End Select

        'Dim Arg As Vector = New Vector

        'ReDim Arg.Coord(5)

        'Select Case Mode

        '    Case YES_TEMP_NO_TIME

        '        ReDim Integral_Eps_DEPEND.Point(N_point_side - 1)

        '        For i_point As Integer = 0 To N_point_side - 1

        '            ReDim Arg.Coord(5)

        '            Arg.Coord(0) = Temp_min + i_point * (Temp_max - Temp_min) / (N_point_side - 1)

        '            Integral_Eps_DEPEND.Point(i_point) = Calc_integral_Eps_Point(Arg)

        '        Next i_point

        '    Case NO_TEMP_YES_TIME

        '        ReDim Integral_Eps_DEPEND.Point(N_point_side - 1)

        '        For i_point As Integer = 0 To N_point_side - 1

        '            Arg.Coord(1) = Time_min + i_point * (Time_max - Time_min) / (N_point_side - 1)

        '            Integral_Eps_DEPEND.Point(i_point) = Calc_integral_Eps_Point(Arg)

        '        Next i_point

        '    Case YES_TEMP_YES_TIME

        '        ReDim Integral_Eps_DEPEND.Point(N_point_side * N_point_side - 1)

        '        Dim i_point As Integer = -1

        '        For i As Integer = 0 To N_point_side - 1

        '            Arg.Coord(0) = Temp_min + i * (Temp_max - Temp_min) / (N_point_side - 1)

        '            For j As Integer = 0 To N_point_side - 1

        '                Arg.Coord(1) = Time_min + j * (Time_max - Time_min) / (N_point_side - 1)

        '                i_point += 1

        '                Integral_Eps_DEPEND.Point(i_point) = Calc_integral_Eps_Point(Arg)

        '            Next j

        '        Next i

        'End Select

        '' Зависимость заполнена, выставляем коэффициент n=10, вычисляем ее базу и она готова к использованию

        'Integral_Eps_DEPEND.Calc_base_vector()

        'Integral_Eps_DEPEND.n = 10

        '' Зависимость готова к использованию


        ''Dim writer_1 As System.IO.StreamWriter = New System.IO.StreamWriter("C:\tmp\int_Eps_v_1.txt", False)


        ''For T As Double = 100.0 To 600.0 Step 1.0

        ''    Arg.Coord(0) = T

        ''    Dim integral_Eps As Double

        ''    integral_Eps = Me.Calc_intergal_Eps(Arg)

        ''    writer_1.WriteLine(T.ToString.Replace(",", ".") & " " & integral_Eps.ToString.Replace(",", "."))

        ''Next T

        ''writer_1.Close()

    End Sub

    ''' <summary>
    ''' This sub calculates integral absorbtion values for temperature and time 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Calc_Integral_A_DEPEND()

        ' проверяем, вдруг наша степень черноты - константа?

        If A.GetType.ToString = "System.Double" Then Exit Sub

        ' Инициализируем зависимость
        Integral_A_DEPEND = Calc_Integral_DEPEND(A, 10.0)

    End Sub

    ''' <summary>
    ''' This sub calculates integral of any DEPEND along incoming photon frequency
    ''' </summary>
    ''' <remarks></remarks>
    Public Function Calc_Integral_DEPEND(ByRef DEPEND_to_integrate As DEPEND, ByRef n As Double) As DEPEND

        ' Инициализируем зависимость
        Dim Integral_DEPEND As DEPEND = New DEPEND

        ' Определяем координаты точек, в которых мы будем рассчитывать интегральную степень черноты

        Dim Mode As Integer

        Const NO_FREQ As Integer = 0, NO_TEMP_NO_TIME As Integer = 1, YES_TEMP_NO_TIME As Integer = 2
        Const NO_TEMP_YES_TIME As Integer = 3, YES_TEMP_YES_TIME As Integer = 4


        ' Смотрим от чего зависит степень черноты

        If Eps.Base.Coord(3) = 0 Then

            Mode = NO_FREQ

            Return DEPEND_to_integrate

        End If

        If Eps.Base.Coord(0) = 0 And Eps.Base.Coord(1) = 0 Then

            Mode = NO_TEMP_NO_TIME

            Throw New Exception("The dependency № " & Eps.number & "are is used for for emissivity calculation. This dependency is not depends on temperature or time. It is incorrect.")

            Exit Function

        End If

        If Eps.Base.Coord(0) <> 0 And Eps.Base.Coord(1) = 0 Then

            Mode = YES_TEMP_NO_TIME

        End If

        If Eps.Base.Coord(0) = 0 And Eps.Base.Coord(1) <> 0 Then

            Mode = NO_TEMP_YES_TIME

        End If

        If Eps.Base.Coord(0) <> 0 And Eps.Base.Coord(1) <> 0 Then

            Mode = YES_TEMP_YES_TIME

        End If

        ' в каком режиме понятно.
        ' Теперь надо определить границы того поля, на котором мы будем строить зависимость

        Dim Temp_min, Temp_max, Time_min, Time_max As Double

        Temp_min = Double.MaxValue

        Temp_max = Double.MinValue

        Time_min = Double.MaxValue

        Time_max = Double.MinValue

        For i As Integer = 0 To UBound(DEPEND_to_integrate.Point)

            If DEPEND_to_integrate.Point(i).Argument.Coord(0) < Temp_min Then

                Temp_min = DEPEND_to_integrate.Point(i).Argument.Coord(0)

            End If

            If DEPEND_to_integrate.Point(i).Argument.Coord(0) > Temp_max Then

                Temp_max = DEPEND_to_integrate.Point(i).Argument.Coord(0)

            End If



            If DEPEND_to_integrate.Point(i).Argument.Coord(1) < Time_min Then

                Time_min = DEPEND_to_integrate.Point(i).Argument.Coord(1)

            End If

            If DEPEND_to_integrate.Point(i).Argument.Coord(1) > Time_max Then

                Time_max = DEPEND_to_integrate.Point(i).Argument.Coord(1)

            End If

        Next i

        ' определили поле, на котором будем работать.

        ' определям количество точек в зависимости интегральной степени черноты, которые мы хотим получить.
        Dim N_point_side As Integer

        Select Case Mode

            Case NO_TEMP_YES_TIME, YES_TEMP_NO_TIME

                N_point_side = UBound(DEPEND_to_integrate.Point) + 1

            Case YES_TEMP_YES_TIME

                N_point_side = (UBound(DEPEND_to_integrate.Point) + 1) ^ 0.5

        End Select

        Dim Arg As Vector = New Vector

        ReDim Arg.Coord(5)

        Select Case Mode

            Case YES_TEMP_NO_TIME

                ReDim Integral_DEPEND.Point(N_point_side - 1)

                For i_point As Integer = 0 To N_point_side - 1

                    ReDim Arg.Coord(5)

                    Arg.Coord(0) = Temp_min + i_point * (Temp_max - Temp_min) / (N_point_side - 1)

                    Integral_DEPEND.Point(i_point) = Calc_integral_Point(Arg, DEPEND_to_integrate) 'Calc_integral_Eps_Point(Arg)

                Next i_point

            Case NO_TEMP_YES_TIME

                ReDim Integral_DEPEND.Point(N_point_side - 1)

                For i_point As Integer = 0 To N_point_side - 1

                    Arg.Coord(1) = Time_min + i_point * (Time_max - Time_min) / (N_point_side - 1)

                    Integral_DEPEND.Point(i_point) = Calc_integral_Point(Arg, DEPEND_to_integrate) 'Calc_integral_Eps_Point(Arg)

                Next i_point

            Case YES_TEMP_YES_TIME

                ReDim Integral_DEPEND.Point(N_point_side * N_point_side - 1)

                Dim i_point As Integer = -1

                For i As Integer = 0 To N_point_side - 1

                    Arg.Coord(0) = Temp_min + i * (Temp_max - Temp_min) / (N_point_side - 1)

                    For j As Integer = 0 To N_point_side - 1

                        Arg.Coord(1) = Time_min + j * (Time_max - Time_min) / (N_point_side - 1)

                        i_point += 1

                        Integral_DEPEND.Point(i_point) = Calc_integral_Point(Arg, DEPEND_to_integrate) 'Calc_integral_Eps_Point(Arg)

                    Next j

                Next i

        End Select

        ' Зависимость заполнена, выставляем коэффициент n=10, вычисляем ее базу и она готова к использованию

        Integral_DEPEND.Calc_base_vector()

        Integral_DEPEND.n = n

        Return Integral_DEPEND

        ' Зависимость готова к использованию

    End Function

    ''' <summary>
    ''' This function returns value in a point of integral Eps DEPEND.
    ''' For a given temperture Arg.Coord(0) and time Arg.Coord(1) function integrates Eps dependency 
    ''' over frequency Arg.Coord(3) from 0 Hz to automatically selected value Hz.
    ''' The assumption. The face with given Eps emitts like black body.
    ''' The simple Euler integrator is used. 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function Calc_integral_Eps_Point(ByVal Arg As Vector) As Function_point

        Dim integral_Eps As Double = 0

        Dim T As Double = Arg.Coord(0) + 0.000001

        Dim Arg_to_Eps As Vector = New Vector

        ReDim Arg_to_Eps.Coord(5)

        Arg_to_Eps.Coord(0) = T

        Arg_to_Eps.Coord(1) = Arg.Coord(1)

        ' the frequency with maximum emission
        Dim v_max As Double = v_max_const * T

        Dim v_start As Double = v_max / 10

        Dim v_step As Double = v_start

        Dim v_finish As Double = 100 * v_step


        For v As Double = v_start To v_finish Step v_step

            Arg_to_Eps.Coord(3) = v

            integral_Eps += v_step * 1 / 6 * h ^ 4 * Math.Exp(-h * v / (k * T)) * v ^ 3 / (k ^ 4 * T ^ 4) * Eps.return_value(Arg_to_Eps).Coord(0)

        Next v

        Dim R_point As Function_point = New Function_point(5, 0)

        R_point.Argument.Coord(0) = T

        R_point.Argument.Coord(1) = Arg.Coord(1)

        R_point.Value.Coord(0) = integral_Eps

        Return R_point

    End Function


    ''' <summary>
    ''' This function returns value in a point of integral DEPEND.
    ''' For a given temperture Arg.Coord(0) and time Arg.Coord(1) function integrates DEPEND dependency 
    ''' over frequency Arg.Coord(3) from 0 Hz to automatically selected value Hz.
    ''' The assumption. The face with given Eps emitts like black body.
    ''' The simple Euler integrator is used. 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function Calc_integral_Point(ByVal Arg As Vector, ByRef DEPEND As DEPEND) As Function_point

        Dim integral_Eps As Double = 0

        Dim T As Double = Arg.Coord(0) + 0.000001

        Dim Arg_to_Eps As Vector = New Vector

        ReDim Arg_to_Eps.Coord(5)

        Arg_to_Eps.Coord(0) = T

        Arg_to_Eps.Coord(1) = Arg.Coord(1)

        ' the frequency with maximum emission
        Dim v_max As Double = v_max_const * T

        Dim v_start As Double = v_max / 10

        Dim v_step As Double = v_start

        Dim v_finish As Double = 100 * v_step


        For v As Double = v_start To v_finish Step v_step

            Arg_to_Eps.Coord(3) = v

            integral_Eps += v_step * 1 / 6 * h ^ 4 * Math.Exp(-h * v / (k * T)) * v ^ 3 / (k ^ 4 * T ^ 4) * DEPEND.Return_value(Arg_to_Eps).Coord(0)

        Next v

        Dim R_point As Function_point = New Function_point(5, 0)

        R_point.Argument.Coord(0) = T

        R_point.Argument.Coord(1) = Arg.Coord(1)

        R_point.Value.Coord(0) = integral_Eps

        Return R_point

    End Function


    '''' <summary>
    '''' This function calculates transmittance coefficient for a given Argument
    '''' </summary>
    '''' <param name="Argument">Given Argument</param>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public Function Calc_D(ByRef Argument As Vector) As Double

    '    If D.GetType.ToString <> "System.Double" Then

    '        Return D.Return_value(Argument).Coord(0)

    '    End If

    '    If D.GetType.ToString = "System.Double" Then

    '        Return D

    '    End If

    'End Function

    '''' <summary>
    '''' This function calculates dark ratio for a given Argument
    '''' </summary>
    '''' <param name="Argument">Given Argument</param>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public Function Calc_Eps(ByRef Argument As Vector) As Double

    '    If Eps.GetType.ToString <> "System.Double" Then

    '        Return Eps.Return_value(Argument).Coord(0)

    '    End If

    '    If Eps.GetType.ToString = "System.Double" Then

    '        Return Eps

    '    End If

    'End Function

End Class

''' <summary>
''' This class for isotropic element's thermal-optical properties determination
''' </summary>
''' <remarks></remarks>
<Serializable()> _
Public Class MEDIUM1

    Inherits TOP_item

    ' коэффициент преломления
    Public n() As Object

    ' показатель поглощения k.
    ' Линейный коэффициент поглощения по глубине определяется как обратная величина к расстоянию,
    ' на котором интенсивность прошедшего потока излучения снижается в e раз.
    'q_1(l) = q_0 * e ^ (-k * l)
    Public Shadows k() As Object

    ' теплопроводность
    Public lambda() As Object

    ' удельная теплоемкость
    Public Shadows c As Object

    ' номера термооптических свойств для граней (максимум может быть 6 граней)
    Public FTOP() As Object


    Public Overrides Sub Load()

        ReDim n(0)

        ReDim k(0)

        ReDim lambda(0)

        ReDim FTOP(5)


        Number = Val(Fields_array(1))

        n(0) = my_Val(Fields_array(2))

        k(0) = my_Val(Fields_array(3))

        lambda(0) = my_Val(Fields_array(4))

        c = my_Val(Fields_array(5))

        For i As Integer = 0 To 5

            FTOP(i) = Val(Fields_array(11 + i))

        Next i


    End Sub

    Public Overrides Sub Prepare_Save()


        ReDim Fields_array(16)

        Fields_array(0) = Type

        Fields_array(1) = my_Str(Number)

        Fields_array(2) = my_Str(n(0))

        Fields_array(3) = my_Str(k(0))

        Fields_array(4) = my_Str(lambda(0))

        Fields_array(5) = my_Str(c)

        For i As Integer = 0 To 5

            Fields_array(11 + i) = my_Str(FTOP(i))

        Next i


    End Sub
    Public Sub New(Optional ByRef String_in_file As String = "")

        MyBase.New(String_in_file)

    End Sub

    Public Sub New(ByRef Fields_array() As String)

        Me.Fields_array = Fields_array

        Load()

    End Sub


    ''' <summary>
    ''' This function calculates refraction factor from incoming photon
    ''' </summary>
    ''' <param name="Argument"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Calc_n(ByRef Argument As Vector) As Double

        If n(0).GetType.ToString <> "System.Double" Then

            Calc_n = n(0).Return_value(Argument).Coord(0)

        End If

        If n(0).GetType.ToString = "System.Double" Then

            Calc_n = Me.n(0)

        End If

    End Function


    ''' <summary>
    '''  This function calculates depletion factor from incoming photon
    ''' </summary>
    ''' <param name="Argument"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Calc_k(ByRef Argument As Vector) As Double

        If k(0).GetType.ToString <> "System.Double" Then

            Calc_k = k(0).Return_value(Argument).Coord(0)

        End If

        If k(0).GetType.ToString = "System.Double" Then

            Calc_k = Me.k(0)

        End If

    End Function


    ''' <summary>
    ''' This function calculates thermal conductivity
    ''' </summary>
    ''' <param name="Argument"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Calc_lambda(ByRef Argument As Vector) As Double

        If lambda(0).GetType.ToString <> "System.Double" Then

            Calc_lambda = lambda(0).Return_value(Argument).Coord(0)

        End If

        If lambda(0).GetType.ToString = "System.Double" Then

            Calc_lambda = Me.lambda(0)

        End If

    End Function

    ''' <summary>
    ''' This function calculates specific heat capacity
    ''' </summary>
    ''' <param name="Argument"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Calc_c(ByRef Argument As Vector) As Double

        If c.GetType.ToString = "System.Double" Then

            Return c

        End If

        If c.GetType.ToString <> "System.Double" Then

            Return c.Return_value(Argument).Coord(0)

        End If

    End Function


End Class

'''' <summary>
'''' This class for anisotropic element's thermal-optical properties determination
'''' </summary>
'''' <remarks></remarks>
'Public Class MEDIUM9

'    Inherits TOP_item

'    ' коэффициент преломления
'    Public n() As Object

'    ' показатель поглощения k.
'    ' Линейный коэффициент поглощения по глубине определяется как обратная величина к расстоянию,
'    ' на котором интенсивность прошедшего потока излучения снижается в e раз.
'    'q_1(l) = q_0 * e ^ (-k * l)
'    Public k() As Object

'    ' теплопроводность
'    Public lambda() As Object

'    ' удельная теплоемкость
'    Public c As Object

'    ' номера термооптических свойств для граней (максимум может быть 6 граней)
'    Public FTOP() As Object


'    Public Overrides Sub Load()

'        ReDim n(2)

'        ReDim k(2)

'        ReDim lambda(2)

'        ReDim FTOP(5)


'        Number = Val(Fields_array(1))

'        n(0) = my_Val(Fields_array(2))

'        n(1) = my_Val(Fields_array(3))

'        n(2) = my_Val(Fields_array(4))

'        k(0) = my_Val(Fields_array(5))

'        k(1) = my_Val(Fields_array(6))

'        k(2) = my_Val(Fields_array(7))




'        lambda(0) = my_Val(Fields_array(11))

'        lambda(1) = my_Val(Fields_array(12))

'        lambda(2) = my_Val(Fields_array(13))

'        c = my_Val(Fields_array(14))


'        For i As Integer = 0 To 5

'            FTOP(i) = Val(Fields_array(21 + i))

'        Next i


'    End Sub

'    Public Overrides Sub Prepare_Save()


'        ReDim Fields_array(26)

'        Fields_array(0) = Type

'        Fields_array(1) = my_Str(Number)

'        Fields_array(2) = my_Str(n(0))

'        Fields_array(3) = my_Str(n(1))

'        Fields_array(4) = my_Str(n(2))

'        Fields_array(5) = my_Str(k(0))

'        Fields_array(6) = my_Str(k(1))

'        Fields_array(7) = my_Str(k(2))



'        Fields_array(11) = my_Str(lambda(0))

'        Fields_array(12) = my_Str(lambda(1))

'        Fields_array(13) = my_Str(lambda(2))

'        Fields_array(14) = my_Str(c)



'        For i As Integer = 0 To 5

'            Fields_array(21 + i) = my_Str(FTOP(i))

'        Next i


'    End Sub
'    Public Sub New(Optional ByRef String_in_file As String = "")

'        MyBase.New(String_in_file)

'    End Sub

'    Public Sub New(ByRef Fields_array() As String)

'        Me.Fields_array = Fields_array

'        Load()

'    End Sub

'    ''' <summary>
'    ''' This function calculates refraction factor from incoming photon
'    ''' </summary>
'    ''' <param name="in_Photon"></param>
'    ''' <param name="Coordinate_system"></param>
'    ''' <param name="Temperature"></param>
'    ''' <returns></returns>
'    ''' <remarks></remarks>
'    Public Function Calc_n(ByVal in_Photon As Photon, ByVal Coordinate_system As Object, ByVal Temperature As Double) As Double

'        If Coordinate_system.GetType.ToString <> "System.Double" Then

'            in_Photon.Direction = Coordinate_system.Transform(in_Photon.Direction, 0, Coordinate_system)

'            in_Photon.Direction = Norm_vector(in_Photon.Direction)

'        End If

'        Dim n_(UBound(n)) As Double

'        Dim Argument As New Vector

'        ReDim Argument.Coord(1)

'        For i As Integer = 0 To UBound(n)

'            If n(i).GetType.ToString <> "System.Double" Then

'                Argument.Coord(0) = Temperature

'                Argument.Coord(1) = in_Photon.Wavelength

'                n_(i) = n(i).Return_value(Argument).Coord(0)

'            End If

'            If n(i).GetType.ToString = "System.Double" Then

'                n_(i) = Me.n(i)

'            End If

'        Next i

'        Calc_n = n_(0) * in_Photon.Direction.Coord(0) + n_(1) * in_Photon.Direction.Coord(1) + n_(2) * in_Photon.Direction.Coord(2)

'    End Function


'    ''' <summary>
'    ''' This function calculates depletion factor for incoming photon
'    ''' </summary>
'    ''' <param name="in_Photon"></param>
'    ''' <param name="Coordinate_system"></param>
'    ''' <param name="Temperature"></param>
'    ''' <returns></returns>
'    ''' <remarks></remarks>
'    Public Function Calc_k(ByVal in_Photon As Photon, ByVal Coordinate_system As Object, ByVal Temperature As Double) As Double

'        If Coordinate_system.GetType.ToString <> "System.Double" Then

'            in_Photon.Direction = Coordinate_system.Transform(in_Photon.Direction, 0, Coordinate_system)

'            in_Photon.Direction = Norm_vector(in_Photon.Direction)

'        End If

'        Dim k_(UBound(n)) As Double

'        Dim Argument As New Vector

'        ReDim Argument.Coord(1)

'        For i As Integer = 0 To UBound(n)

'            If k(i).GetType.ToString <> "System.Double" Then

'                Argument.Coord(0) = Temperature

'                Argument.Coord(1) = in_Photon.Wavelength

'                k_(i) = k(i).Return_value(Argument).Coord(0)

'            End If

'            If k(i).GetType.ToString = "System.Double" Then

'                k_(i) = Me.k(i)

'            End If

'        Next i

'        Calc_k = k_(0) * in_Photon.Direction.Coord(0) + k_(1) * in_Photon.Direction.Coord(1) + k_(2) * in_Photon.Direction.Coord(2)

'    End Function

'    ''' <summary>
'    ''' This function calculates thermal conductivity
'    ''' </summary>
'    ''' <param name="in_Vector"></param>
'    ''' <returns></returns>
'    ''' <remarks></remarks>
'    Public Function Calc_lambda(ByVal in_Vector As Vector, ByVal Coordinate_system As Object, ByVal Temperature As Double) As Double

'        If Coordinate_system.GetType.ToString <> "System.Double" Then

'            in_Vector = Coordinate_system.Transform(in_Vector, 0, Coordinate_system)

'            in_Vector = Norm_vector(in_Vector)

'        End If

'        Dim lambda_(UBound(n)) As Double

'        Dim Argument As New Vector

'        ReDim Argument.Coord(0)

'        For i As Integer = 0 To UBound(n)

'            If lambda(i).GetType.ToString <> "System.Double" Then

'                Argument.Coord(0) = Temperature

'                lambda_(i) = lambda(i).Return_value(Argument).Coord(0)

'            End If

'            If lambda(i).GetType.ToString = "System.Double" Then

'                lambda_(i) = Me.lambda(i)

'            End If

'        Next i

'        Calc_lambda = lambda_(0) * in_Vector.Coord(0) + lambda_(1) * in_Vector.Coord(1) + lambda_(2) * in_Vector.Coord(2)

'    End Function

'    ''' <summary>
'    ''' This function calculates specific heat capacity
'    ''' </summary>
'    ''' <param name="Temperature"></param>
'    ''' <returns></returns>
'    ''' <remarks></remarks>
'    Public Function Calc_c(ByVal Temperature As Double) As Double

'        If c.GetType.ToString <> "System.Double" Then

'            Dim Argument As New Vector

'            ReDim Argument.Coord(0)

'            Argument.Coord(0) = Temperature

'            Calc_c = c.Return_value(Argument).Coord(0)

'        End If

'        If c.GetType.ToString = "System.Double" Then

'            Calc_c = c

'        End If

'    End Function
'End Class


''' <summary>
''' This class is for radiation sources definition
''' </summary>
''' <remarks></remarks>
<Serializable()> _
Public Class RSOURCE1


    Inherits TOP_item

    ''' <summary>
    ''' Radiation heat source coordinates in global coordinate system
    ''' </summary>
    ''' <remarks></remarks>
    Public X, Y, Z As Object

    ''' <summary>
    ''' Radiation heat source position in global coordinate system
    ''' Used for moving radiation source
    ''' </summary>
    ''' <remarks></remarks>
    Public POS As Object

    ''' <summary>
    ''' Radiation heat source emissive power
    ''' </summary>
    ''' <remarks></remarks>
    Public Q As Object

    ''' <summary>
    ''' Radiation heat source temperature
    ''' </summary>
    ''' <remarks></remarks>
    Public T As Object

    ''' <summary>
    ''' Radiation heat source cumulative distribution of photon frequency
    ''' </summary>
    ''' <remarks></remarks>
    Public CDF As Object


    '''' <summary>
    '''' Maximum error during radiation heat source's model illumination.
    '''' </summary>
    '''' <remarks></remarks>
    'Public Epsilon As Double


    ''' <summary>
    ''' Number of photons, emitted from radiation source at time of one solution step
    ''' </summary>
    ''' <remarks></remarks>
    Public N_photon As Integer



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

    ''' <summary>
    ''' Parts of source energy absorbed by model's elements
    ''' </summary>
    ''' <remarks></remarks>
    Public B_row() As Double

    ''' <summary>
    ''' Radiation heat source current position
    ''' </summary>
    ''' <remarks></remarks>
    Public Current_position As New Vector

    Public Overrides Sub Load()

        Number = Val(Fields_array(1))

        X = Val(Fields_array(2))

        POS = X

        Y = my_Val(Fields_array(3))

        Z = my_Val(Fields_array(4))

        Q = my_Val(Fields_array(5))

        T = my_Val(Fields_array(6))

        CDF = my_Val(Fields_array(7))

        N_photon = Val(Fields_array(8))

        '' проверяем целое ли число

        'If InStr(Fields_array(8), ".") > 0 Or InStr(Fields_array(8).ToUpper, "E") > 0 Then

        '    Me.Epsilon = Val(Fields_array(8))

        '    N_photon = -1

        'Else

        '    Me.Epsilon = -1

        '    N_photon = Val(Fields_array(8))

        'End If

    End Sub

    Public Overrides Sub Prepare_Save()


        ReDim Fields_array(7)

        Fields_array(0) = Type

        Fields_array(1) = my_Str(Number)

        If Y Is Nothing And Z Is Nothing Then

            Fields_array(2) = my_Str(POS)

            Fields_array(3) = ""

            Fields_array(4) = ""


        Else

            Fields_array(2) = my_Str(X)

            Fields_array(3) = my_Str(Y)

            Fields_array(4) = my_Str(Z)

        End If

        Fields_array(5) = my_Str(Q)

        Fields_array(6) = my_Str(T)

        Fields_array(7) = my_Str(CDF)

        Fields_array(8) = my_Str(Me.N_photon)

    End Sub
    Public Sub New(Optional ByRef String_in_file As String = "")

        MyBase.New(String_in_file)

        Argument = New Vector

        ReDim Argument.Coord(5)

    End Sub

    Public Sub New(ByRef Fields_array() As String)

        Me.Fields_array = Fields_array

        Load()

        Argument = New Vector

        ReDim Argument.Coord(5)

    End Sub

    ''' <summary>
    ''' This function return current radiation power
    ''' </summary>
    ''' <param name="Argument"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Calc_Q(ByRef Argument As Vector) As Double

        If Q.GetType.ToString = "System.Double" Then

            Return Q

        Else

            Return Q.Return_value(Argument)(0)

        End If

    End Function

    ''' <summary>
    ''' This function return current radiation heat source position
    ''' </summary>
    ''' <param name="Argument"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Calc_position(Optional ByRef Argument As Vector = Nothing) As Vector

        If Argument Is Nothing Then

            Argument = Me.Argument

        End If

        If X Is Nothing Then

            Return POS.Return_value(Argument)

        Else

            Return Current_position

        End If

    End Function

    ''' <summary>
    ''' This function return current radiation heat source temperature
    ''' </summary>
    ''' <param name="Argument"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Calc_T(ByRef Argument As Vector) As Double

        If T.GetType.ToString = "System.Double" Then

            Return T

        Else

            Return T.Return_value(Argument)(0)

        End If
    End Function

End Class

''' <summary>
''' This class is for heat sources definition
''' </summary>
''' <remarks></remarks>
<Serializable()> _
Public Class HSOURCE1

    Inherits TOP_item

    ''' <summary>
    ''' Heat source power. Can be positive or negative
    ''' </summary>
    ''' <remarks></remarks>
    Public Q As Object


    ''' <summary>
    ''' Heat source's elements. Heat source consists of these elements.
    ''' </summary>
    ''' <remarks></remarks>
    Public Element() As Object


    ''' <summary>
    ''' Number of elements, heat source consist of.
    ''' </summary>
    ''' <remarks></remarks>
    Public N_elements As Long

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

    Public Overrides Sub Load()

        Number = Val(Fields_array(1))

        Q = my_Val(Fields_array(2))

        ReDim Element(Max_items_default_2)

        Dim i_Element As Integer = -1

        For i As Integer = 10 To UBound(Fields_array) Step 10

            For j As Integer = 1 To 8

                If i + j <= UBound(Fields_array) Then

                    If Val(Fields_array(i + j)) > 0 Then

                        i_Element += 1

                        Element(i_Element) = my_Val(Fields_array(i + j))

                    End If

                End If

            Next j

        Next i

        ReDim Preserve Element(i_Element)

        N_elements = i_Element + 1

    End Sub

    Public Overrides Sub Prepare_Save()

        Dim N_lines_for_Elements As Integer

        If UBound(Element) = UBound(Element) \ 8 Then

            N_lines_for_Elements = UBound(Element) \ 8

        Else

            N_lines_for_Elements = UBound(Element) \ 8 + 1

        End If

        ReDim Fields_array(10 + N_lines_for_Elements * 10)

        Fields_array(0) = Type

        Fields_array(1) = my_Str(Number)

        Fields_array(2) = my_Str(Q)

        Dim Line_Current As Integer = 1

        Dim i_Current As Integer = 1

        For i As Integer = 0 To UBound(Element)

            Fields_array(Line_Current * 10 + i_Current) = my_Str(Element(i))

            i_Current += 1

            If i_Current > 8 Then

                i_Current = 1

                Line_Current += 1

            End If

        Next i

    End Sub
    Public Sub New(Optional ByRef String_in_file As String = "")

        MyBase.New(String_in_file)

        Argument = New Vector

        ReDim Argument.Coord(5)

    End Sub

    Public Sub New(ByRef Fields_array() As String)

        Me.Fields_array = Fields_array

        Load()

        Argument = New Vector

        ReDim Argument.Coord(5)

    End Sub

    ''' <summary>
    ''' This function returns quatinty of heat, emitted by heat source.
    ''' </summary>
    ''' <param name="Argument">Parameters heat source of depends on.
    ''' Table of parameters (sensor components)
    ''' 0 - temperature,
    ''' 1 - time,
    ''' 2 - random value,
    ''' 3 - incoming radiation wavelenght
    ''' 4 - incoming radiation angle of elevation in local coordinate system
    ''' 5 - incoming radiation azimuth in local coordinate system</param>
    ''' <returns>Quatinty of heat</returns>
    ''' <remarks></remarks>
    Public Function Heat(ByRef Argument As Vector) As Double

        If Q.GetType.ToString = "System.Double" Then

            Return Q / N_elements

        Else

            Return Q.Return_value(Argument).Coord(0) / N_elements

        End If

    End Function

End Class
''' <summary>
''' This class is for initial temperature field definition
''' </summary>
''' <remarks></remarks>
<Serializable()> _
Public Class T_INIT

    Inherits TOP_item

    ''' <summary>
    ''' Initial temperature fields elements. Temperature field defined on these elements. 
    ''' </summary>
    ''' <remarks></remarks>
    Public Element() As Object


    ''' <summary>
    ''' Temperatures of elements.
    ''' </summary>
    ''' <remarks></remarks>
    Public Temp() As Double

    Public Overrides Sub Load()

        Number = Val(Fields_array(1))

        ReDim Element(Max_items_default_2)

        ReDim Temp(Max_items_default_2)

        Dim i_Element As Integer = -1

        For i As Integer = 10 To UBound(Fields_array) Step 10

            For j As Integer = 1 To 7 Step 2

                If i + j <= UBound(Fields_array) Then

                    If Fields_array(i + j) <> "" Then

                        i_Element += 1

                        Element(i_Element) = my_Val(Fields_array(i + j))

                        Temp(i_Element) = Val(Fields_array(i + j + 1))

                    End If

                End If

            Next j

        Next i

        ReDim Preserve Element(i_Element)

        ReDim Preserve Temp(i_Element)

    End Sub

    Public Overrides Sub Prepare_Save()

        Dim N_lines_for_Elements As Integer

        If UBound(Element) = UBound(Element) \ 8 Then

            N_lines_for_Elements = UBound(Element) \ 8

        Else

            N_lines_for_Elements = UBound(Element) \ 8 + 1

        End If

        ReDim Fields_array(10 + N_lines_for_Elements * 10)

        Fields_array(0) = Type

        Fields_array(1) = my_Str(Number)

        Dim Line_Current As Integer = 1

        Dim i_Current As Integer = 1

        For i As Integer = 0 To UBound(Element)

            Fields_array(Line_Current * 10 + i_Current) = my_Str(Element(i))

            Fields_array(Line_Current * 10 + i_Current + 1) = my_Str(Temp(i))

            i_Current += 2

            If i_Current > 7 Then

                i_Current = 1

                Line_Current += 1

            End If

        Next i

    End Sub
    Public Sub New(Optional ByRef String_in_file As String = "")

        MyBase.New(String_in_file)

    End Sub

    Public Sub New(ByRef Fields_array() As String)

        Me.Fields_array = Fields_array

        Load()

    End Sub

End Class

''' <summary>
''' This class is for temperature field definition during the analysis sequence
''' </summary>
''' <remarks></remarks>
<Serializable()> _
Public Class T_SET

    Inherits TOP_item

    ''' <summary>
    ''' Initial temperature fields elements. Temperature field defined on these elements. 
    ''' </summary>
    ''' <remarks></remarks>
    Public Element() As Object


    ''' <summary>
    ''' Temperatures of elements. Can be Double or DEPEND.
    ''' </summary>
    ''' <remarks></remarks>
    Public Temp() As Object

    Public Overrides Sub Load()

        Number = Val(Fields_array(1))

        ReDim Element(Max_items_default_2)

        ReDim Temp(Max_items_default_2)

        Dim i_Element As Integer = -1

        For i As Integer = 10 To UBound(Fields_array) Step 10

            For j As Integer = 1 To 7 Step 2

                If i + j <= UBound(Fields_array) Then

                    If Fields_array(i + j) <> "" Then

                        i_Element += 1

                        Element(i_Element) = my_Val(Fields_array(i + j))

                        Temp(i_Element) = my_Val(Fields_array(i + j + 1))

                    End If

                End If

            Next j

        Next i

        ReDim Preserve Element(i_Element)

        ReDim Preserve Temp(i_Element)

    End Sub

    Public Overrides Sub Prepare_Save()

        Dim N_lines_for_Elements As Integer

        If UBound(Element) = UBound(Element) \ 8 Then

            N_lines_for_Elements = UBound(Element) \ 8

        Else

            N_lines_for_Elements = UBound(Element) \ 8 + 1

        End If

        ReDim Fields_array(10 + N_lines_for_Elements * 10)

        Fields_array(0) = Type

        Fields_array(1) = my_Str(Number)

        Dim Line_Current As Integer = 1

        Dim i_Current As Integer = 1

        For i As Integer = 0 To UBound(Element)

            Fields_array(Line_Current * 10 + i_Current) = my_Str(Element(i))

            Fields_array(Line_Current * 10 + i_Current + 1) = my_Str(Temp(i))

            i_Current += 2

            If i_Current > 7 Then

                i_Current = 1

                Line_Current += 1

            End If

        Next i

    End Sub
    Public Sub New(Optional ByRef String_in_file As String = "")

        MyBase.New(String_in_file)

    End Sub

    Public Sub New(ByRef Fields_array() As String)

        Me.Fields_array = Fields_array

        Load()

    End Sub

End Class

<Serializable()> _
Public Class Face_code

    Public R As Byte

    Public G As Byte

    Public B As Byte

    Public Sub New()

        R = 0

        G = 0

        B = 0

    End Sub

End Class

''' <summary>
''' This class collect information about face-photon interaction
''' </summary>
''' <remarks></remarks>
<Serializable()> _
Public Class Face_photon_interaction

    ''' <summary>
    ''' Number of photons, diffuselly reflected
    ''' </summary>
    ''' <remarks></remarks>
    Public N_R_dif As Double

    ''' <summary>
    ''' Number of photons, mirrorly reflected
    ''' </summary>
    ''' <remarks></remarks>
    Public N_R_mir As Double

    ''' <summary>
    ''' Number of absorpted photons
    ''' </summary>
    ''' <remarks></remarks>
    Public N_A As Double


    ''' <summary>
    ''' Number of transmitted photons
    ''' </summary>
    ''' <remarks></remarks>
    Public N_D As Double

    '''' <summary>
    '''' Face's absorption during global heat transfer.
    '''' This variable forms row in global radiation heat transfer matrix.
    '''' </summary>
    '''' <remarks></remarks>
    'Public Global_Absorbtion As Double

    Public Sub New()

        N_R_dif = 0

        N_R_mir = 0

        N_A = 0

        N_D = 0

        'Global_Absorbtion = 0

    End Sub
End Class
''' <summary>
''' Bundle of photons with same wavelength and same propagation direction
''' It can be used, for example, for Monte-Carlo calculation of radiation heat transfer angle coefficient matrix
''' </summary>
''' <remarks></remarks>
<Serializable()> _
Public Class Photon

    Inherits TOP_item

    ''' <summary>
    ''' Normalized vector of photon propagation direction 
    ''' </summary>
    ''' <remarks></remarks>
    Public Direction As New Vector

    ''' <summary>
    ''' Photon frequency
    ''' </summary>
    ''' <remarks></remarks>
    Public Frequency As Double

    ''' <summary>
    ''' Photon depletion due to propogation through medium
    ''' </summary>
    ''' <remarks></remarks>
    Public Depletion As Double

    ''' <summary>
    ''' The point, where photon was emitted or reflected or passed by.
    ''' </summary>
    ''' <remarks></remarks>
    Public Center As Vector

    ''' <summary>
    ''' The medium, there propagation occuring now
    ''' </summary>
    ''' <remarks></remarks>
    Public Medium As Object

    ''' <summary>
    ''' The photon power
    ''' </summary>
    ''' <remarks></remarks>
    Public Shadows Power As Double


    Public Sub New()

        Frequency = 0

        Depletion = 1

    End Sub

    Public Overrides Sub Load()


    End Sub

    Public Overrides Sub Prepare_save()


    End Sub




    ''' <summary>
    ''' This function return photon with random direction and random wavelength.
    ''' </summary>
    ''' <param name="Source_location">The location of radiation source</param>
    ''' <param name="Emission_angle">Opening angle in emission cone.
    ''' If Emission_angle = PI, photons emitted in all of directions, spherically.
    ''' If Emission_angle = PI/2, photons emitted in one half-space.
    ''' If Emission_angle more then 0 and less than  PI/2, photons emitted in cone.</param>
    ''' <param name="CDF"></param>
    ''' <param name="Argument"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Emit_random_photon(ByRef Source_location As Vector, ByRef Emission_angle As Double, ByRef CDF As DEPEND, ByRef Argument As Vector) As Photon

        Dim Photon As New Photon

        Photon.Depletion = 1

        Photon.Center = Source_location

        ' Field medium must be filled later in calling function with global medium 
        Photon.Medium = Nothing

        If CDF Is Nothing Then

            Photon.Frequency = Random_photon_frequency(Argument.Coord(0))

        Else

            Photon.Frequency = Random_photon_frequency(CDF, Argument)

        End If

        'Dim Angle_1, Angle_2 As Double

        'Randomize()

        'Angle_1 = Rnd() * 2 * Math.PI

        'Angle_2 = Rnd() * Emission_angle

        'Photon.Direction.Coord(0) = Math.Cos(Angle_1) * Math.Sin(Angle_2)

        'Photon.Direction.Coord(1) = Math.Sin(Angle_1) * Math.Sin(Angle_2)

        'Photon.Direction.Coord(2) = Math.Cos(Angle_2)


        Do


            'Photon.Direction.Coord(0) = -Math.Sin(Emission_angle) + Rnd() * 2 * Math.Sin(Emission_angle)

            'Photon.Direction.Coord(1) = -Math.Sin(Emission_angle) + Rnd() * 2 * Math.Sin(Emission_angle)

            Photon.Direction.Coord(0) = -2 * Math.Sin(Emission_angle) + Main_form.Solver.MT_RND.genrand_real1 * 4 * Math.Sin(Emission_angle)

            Photon.Direction.Coord(1) = -2 * Math.Sin(Emission_angle) + Main_form.Solver.MT_RND.genrand_real1 * 4 * Math.Sin(Emission_angle)

            If Math.Sqrt(Photon.Direction.Coord(0) * Photon.Direction.Coord(0) + Photon.Direction.Coord(1) * Photon.Direction.Coord(1)) <= Math.Sin(Emission_angle) Then

                Exit Do

            End If

        Loop

        Photon.Direction.Coord(2) = Math.Sqrt(1 - Photon.Direction.Coord(0) * Photon.Direction.Coord(0) - Photon.Direction.Coord(1) * Photon.Direction.Coord(1))


        Emit_random_photon = Photon

    End Function

    Public Function Generate_N_random_photons(ByRef Source_location As Vector, ByRef N_photons As Integer, ByRef Emission_angle As Double, ByRef CDF As DEPEND, ByRef Argument As Vector, ByRef Emitter_medium As MEDIUM1, ByRef Total_outcoming_heat_power As Double) As Photon()

        Dim Photon(N_photons - 1) As Photon

        ' Доля излучения, переносимая одном фотоном. 
        ' В случае, если термооптические свойства зависят от частоты излучения, то нужно, чтобы доли энергии,
        ' переносимые каждым фотоном, тоже зависели от частоты. R - это и есть доли 
        Dim R(N_photons - 1) As Double

        Dim R_summ As Double = 0

        For i As Integer = 0 To UBound(Photon)

            Photon(i) = Emit_random_photon(Source_location, Emission_angle, CDF, Argument)

            Photon(i).Medium = Emitter_medium

            R(i) = 1 / 6 * h ^ 4 * Math.Exp(-h * Photon(i).Frequency / (k * Argument.Coord(0))) * Photon(i).Frequency ^ 3 / (k ^ 4 * Argument.Coord(0) ^ 4)

            R_summ += R(i)

        Next i

        For i As Integer = 0 To UBound(Photon)

            Photon(i).Depletion = 1 'R(i) / R_summ

            Photon(i).Power = Total_outcoming_heat_power * R(i) / R_summ

        Next i

        Return Photon

    End Function

    Public Function Generate_N_random_photons(ByRef Face As Face, ByRef N_photons As Integer, ByRef Emission_angle As Double, ByRef CDF As DEPEND, ByRef Argument As Vector, ByRef Emitter_medium As MEDIUM1, ByRef Total_outcoming_heat_power As Double) As Photon()

        Dim Photon(N_photons - 1) As Photon

        For i As Integer = 0 To UBound(Photon)

            Photon(i) = Emit_random_photon(Face.Calc_random_emission_center, Emission_angle, CDF, Argument)

            Photon(i).Medium = Emitter_medium

        Next i

        For i As Integer = 0 To UBound(Photon)

            Photon(i).Depletion = 1

            Photon(i).Power = Total_outcoming_heat_power / N_photons
        Next i

        ' строим спектр

        'Draw_spectrum(Photon, Argument.Coord(0), 20)

        Return Photon

    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Photon">Набор фотонов для построение спектра</param>
    ''' <param name="T">Температура излучателя</param>
    ''' <param name="N_bands">Количество диапазонов, в которых суммируется мощность</param>
    ''' <remarks></remarks>
    Private Sub Draw_spectrum(ByRef Photon() As Photon, ByRef T As Double, ByRef N_bands As Integer)

        Dim Band_power(N_bands - 1) As Double

        Dim Band_v(N_bands - 1) As Double

        Dim v_min As Double = v_max_const * T * 0.01

        Dim v_max As Double = v_max_const * T * 10

        Dim i_Band As Integer = 0

        Dim N_photon As Integer = 0

        For i As Integer = 0 To UBound(Photon)

            i_Band = (Photon(i).Frequency - v_min) / (v_max - v_min) * (N_bands - 1)

            If i_Band > -1 And i_Band < N_bands Then

                Band_power(i_Band) += Photon(i).Power

                N_photon += 1

            End If

        Next i

        Dim writer_1 As System.IO.StreamWriter = New System.IO.StreamWriter("C:\tmp\Photon_spectrum.txt", False)

        For i As Integer = 0 To N_bands - 1

            Band_v(i) = i / (N_bands - 1) * (v_max - v_min) + v_min

            writer_1.WriteLine(Band_v(i).ToString & " " & Band_power(i).ToString)

        Next i

        writer_1.Close()

    End Sub


    ''' <summary>
    ''' This function return frequency of photon, emitted from absolute black body with set temperature
    ''' </summary>
    ''' <param name="Temperature">The temperature of body</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Random_photon_frequency(ByRef Temperature As Double) As Double

        If Temperature = 0 Then Return 0

        Dim CDF_Err As Double

        Dim v As Double

        ' Frequency for maximum of emittance
        Dim v_P_max As Double = v_max_const * Temperature

        v = v_P_max

        ' The Planck's CDF
        Dim CDF As Double

        ' The first derivative of Planck CDF
        Dim CDF_der As Double

        Dim CDF_rnd As Double = Main_form.Solver.MT_RND.genrand_real1

        Do

            CDF = Planck_CDF(v, Temperature)

            CDF_Err = Math.Abs(CDF - CDF_rnd) / CDF

            If CDF_Err <= Epsilon / 10000 Then Exit Do

            'If CDF_Err <= Epsilon Then Exit Do

            If CDF = 0 Then Exit Do

            CDF_der = Planck_CDF_der(v, Temperature)

            v = v - (CDF - CDF_rnd) / CDF_der

            If Double.IsNaN(v) Then

                Throw New Exception("The solution is instable")

            End If

        Loop

        Return v

    End Function

    ''' <summary>
    ''' This function return frequency of photon, emitted from body with set CDF, temperature and in set moment of time.
    ''' WARNING!!! CDF must be function of frequency for random value. Not like in absolute black body frequency calculation.
    ''' </summary>
    ''' <param name="CDF">The cumulative distibution function for body radiation frequency</param>
    ''' <param name="Argument">The vector-argument for frequency dependency calculation</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Random_photon_frequency(ByRef CDF As DEPEND, ByRef Argument As Vector) As Double

        Argument.Coord(2) = Main_form.Solver.MT_RND.genrand_real1

        Return CDF.Return_value(Argument).Coord(0)

    End Function

    ''' <summary>
    ''' This function return cumulative distribution of photon frequency depending of photon's frequency and body temperature
    ''' </summary>
    ''' <param name="v">photon frequency</param>
    ''' <param name="T">body temperature</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function Planck_CDF(ByRef v As Double, ByRef T As Double) As Double

        Dim x As Double = h * v / (k * T)

        Planck_CDF = 1 / 6 * (6 - Math.Exp(-x) * (x ^ 3 + 3 * x ^ 2 + 6 * x + 6))

    End Function


    ''' <summary>
    ''' This function return first derivative of cumulative distribution of photon frequency depending of photon's frequency and body temperature
    ''' </summary>
    ''' <param name="v">photon frequency</param>
    ''' <param name="T">body temperature</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function Planck_CDF_der(ByRef v As Double, ByRef T As Double) As Double

        Dim x As Double = h * v / (k * T)

        Planck_CDF_der = 1 / 6 * (h ^ 4) * Math.Exp(-x) * v ^ 3 / ((k * T) ^ 4)

    End Function

End Class

''' <summary>
''' The face of finite element. It can be tree of four-grided. 
''' </summary>
''' <remarks></remarks>
<Serializable()> _
Public Class Face

    Inherits TOP_item

    ' Номер грани
    Public Shadows Number As Long

    ' Массив узлов на которых строится грань
    Public Grid() As Object

    ' Массив координат X вершин грани в системе координат грани
    Public Xpi() As Double

    ' Массив координат Y вершин грани в системе координат грани
    Public Ypi() As Double

    ' Пределы изменения координат вернин граней в системе координат грани
    Public Xmin As Double

    Public Xmax As Double

    Public Ymin As Double

    Public Ymax As Double


    ' Центр системы координат грани. Вектор в глобальной системе координат
    Public Center As New Vector

    ' Ось X системы координат грани. Вектор нормирован.
    Public ox As New Vector

    ' Ось Y системы координат грани Вектор нормирован.
    Public oy As New Vector

    ' Ось Z системы координат грани. Вектор нормирован.
    ' Она же является вектором нормали к элементу
    Public oz As New Vector

    ' Цвета, кодирующие номер грани
    ' Код для грани с номером 0 рассчитывается как для грани с номером 1, для 1 - как для 2 и так далее
    Public Code As Face_code

    ' Термооптические свойства грани
    Public FTOP As FTOP

    Public Interaction As New Face_photon_interaction

    '' В этот массив записывается фотон, который поглотился в грани.
    'Public Absorption_act() As Object

    'Public i_Absorption_act As Integer = -1

    '' Математическое ожидание коэффициента поглощения
    'Public Absortpion_expectation As Double

    ' матрица перехода из системы координат грани в глобальную

    Public from_0(,) As Double

    ' матрица перехода из глобальной системы координат в систему координат грани

    Public to_0(,) As Double

    ' прощадь грани
    Public Area As Double

    ' может ли излучать
    Public Can_Radiate As Boolean

    ' задействована ли
    Public Shadows Enabled As Boolean

    'Элемент, который является предком грани.
    Public Element As Object

    'Элемент, участвующий в теплообмене, в который входит грань
    Public HT_element As Object

    Public Neighbour_face As Face

    ' среда, заполняющая все вне элементов модели
    Public Global_medium As Object

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

    ''' <summary>
    ''' Text associated with a face
    ''' </summary>
    ''' <remarks></remarks>
    Public Text As String

    Const Color_base = 255

    ''' <summary>
    ''' Точность расчетов
    ''' </summary>
    ''' <remarks></remarks>
    Public TOP_Epsilon As Double

    ''' <summary>
    ''' Heat power emitted from face
    ''' </summary>
    ''' <remarks></remarks>
    Public Q_R_outcome As Double

    ''' <summary>
    ''' Heat power hit the face
    ''' </summary>
    ''' <remarks></remarks>
    Public Q_R_income As Double

    ''' <summary>
    ''' Heat power absorbed by face
    ''' </summary>
    ''' <remarks></remarks>
    Public Q_R_absorbed As Double

    ''' <summary>
    ''' Heat power diffuse reflected by face
    ''' </summary>
    ''' <remarks></remarks>
    Public Q_R_diffuse_reflected As Double

    ''' <summary>
    ''' Heat power mirror reflected by face
    ''' </summary>
    ''' <remarks></remarks>
    Public Q_R_mirror_reflected As Double

    ''' <summary>
    ''' Heat power transmitted by face
    ''' </summary>
    ''' <remarks></remarks>
    Public Q_R_transmitted As Double

    Public Overrides Sub Load()

    End Sub

    Public Overrides Sub Prepare_Save()

    End Sub

    Public Sub New()

        ReDim from_0(2, 2), to_0(2, 2)

        Can_Radiate = True

        Code = New Face_code

        Argument = New Vector

        ReDim Argument.Coord(5)

        Me.Type = "FACE"

        'ReDim Absorption_act(Max_items_default)

    End Sub

    ''' <summary>
    ''' Are faces equal or not?
    ''' </summary>
    ''' <param name="A"></param>
    ''' <param name="B"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Shared Operator Like(ByVal A As Face, ByVal B As Face) As Boolean

        Const Eps = 0.000001

        If Math.Sqrt((A.Center.Coord(0) - B.Center.Coord(0)) ^ 2 + (A.Center.Coord(1) - B.Center.Coord(1)) ^ 2 + (A.Center.Coord(2) - B.Center.Coord(2)) ^ 2) < Eps Then

            Return True

        Else

            Return False

        End If

    End Operator

    Public Sub Initialize()

        Interaction = New Face_photon_interaction

        Q_R_income = 0

        Q_R_outcome = 0

    End Sub


    ' процедура заполняет матрицы косинусов перехода в глобальную ск и назад
    Public Sub Calculate_transition_matrices()

        Dim Glob_ox As New Vector
        Dim Glob_oy As New Vector
        Dim Glob_oz As New Vector

        Glob_ox.Coord(0) = 1
        Glob_ox.Coord(1) = 0
        Glob_ox.Coord(2) = 0

        Glob_oy.Coord(0) = 0
        Glob_oy.Coord(1) = 1
        Glob_oy.Coord(2) = 0

        Glob_oz.Coord(0) = 0
        Glob_oz.Coord(1) = 0
        Glob_oz.Coord(2) = 1


        to_0(0, 0) = Cos_angle_between_vectors(Glob_ox, ox)
        to_0(0, 1) = Cos_angle_between_vectors(Glob_ox, oy)
        to_0(0, 2) = Cos_angle_between_vectors(Glob_ox, oz)

        to_0(1, 0) = Cos_angle_between_vectors(Glob_oy, ox)
        to_0(1, 1) = Cos_angle_between_vectors(Glob_oy, oy)
        to_0(1, 2) = Cos_angle_between_vectors(Glob_oy, oz)

        to_0(2, 0) = Cos_angle_between_vectors(Glob_oz, ox)
        to_0(2, 1) = Cos_angle_between_vectors(Glob_oz, oy)
        to_0(2, 2) = Cos_angle_between_vectors(Glob_oz, oz)


        from_0(0, 0) = Cos_angle_between_vectors(ox, Glob_ox)
        from_0(0, 1) = Cos_angle_between_vectors(ox, Glob_oy)
        from_0(0, 2) = Cos_angle_between_vectors(ox, Glob_oz)

        from_0(1, 0) = Cos_angle_between_vectors(oy, Glob_ox)
        from_0(1, 1) = Cos_angle_between_vectors(oy, Glob_oy)
        from_0(1, 2) = Cos_angle_between_vectors(oy, Glob_oz)

        from_0(2, 0) = Cos_angle_between_vectors(oz, Glob_ox)
        from_0(2, 1) = Cos_angle_between_vectors(oz, Glob_oy)
        from_0(2, 2) = Cos_angle_between_vectors(oz, Glob_oz)

    End Sub

    ''' <summary>
    ''' Function transform vector from global coordinate system to face's coordinate system
    ''' </summary>
    ''' <param name="in_Vector"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Transform_to_face(ByVal in_Vector As Vector) As Vector

        Dim tmp_Vector_1 As New Vector

        For i_1 As Integer = 0 To 2

            For j_1 As Integer = 0 To 2

                tmp_Vector_1.Coord(i_1) = tmp_Vector_1.Coord(i_1) + in_Vector.Coord(j_1) * from_0(i_1, j_1)

            Next j_1

        Next i_1

        Transform_to_face = tmp_Vector_1

    End Function
    ''' <summary>
    ''' Function transform vector from face's coordinate system to global coordinate system
    ''' </summary>
    ''' <param name="in_Vector"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Transform_from_face(ByVal in_Vector As Vector) As Vector

        Dim tmp_Vector_1 As New Vector

        For i_1 As Integer = 0 To 2

            For j_1 As Integer = 0 To 2

                tmp_Vector_1.Coord(i_1) = tmp_Vector_1.Coord(i_1) + in_Vector.Coord(j_1) * to_0(i_1, j_1)

            Next j_1

        Next i_1

        Transform_from_face = tmp_Vector_1

    End Function

    ' Процедура определяет соовествие между цветовыми составляющими и номером.
    ' Для перевода используется база Color_base, имеющая смысл максимального значения одной цветовой составляющей
    ' В случае вызова процедуры с ненулевым номером вычисляются цветовые составляющие
    ' В случае вызова процедуры с нулевым номером вычисляется номер
    ' Номера граней начинаются с 1 
    Public Sub Detect_face_number(ByRef Code As Face_code, ByRef Number As Long)

        If Code.R = 0 And Code.G = 0 And Code.B = 0 And Number <= Color_base ^ 3 And Number > 0 Then

            Code.R = Number \ (Color_base ^ 2)

            Code.G = (Number - Code.R * Color_base ^ 2) \ Color_base

            Code.B = Number - Code.R * Color_base ^ 2 - Code.G * Color_base

            'Code.R = Code.R / Color_base

            'Code.G = Code.G / Color_base

            'Code.B = Code.B / Color_base

        End If

        If Number = 0 And Not (Code Is Nothing) Then

            Number = (Code.R * Color_base ^ 2 + Code.G * Color_base + Code.B) '* Color_base

        End If

    End Sub


    ''' <summary>
    ''' This function is for determination incoming photon (incoming to face).
    ''' The photon can be reflected (mirrorly and diffuselly), transmitted or absorbed.
    ''' If photon absorbed, then function returs Nothing, else return photon.
    ''' </summary>
    ''' <param name="in_Photon"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function Face_photon_interaction(ByRef in_Photon As Photon) As Photon

        Dim R_dif As Double

        Dim R_mir As Double

        Dim A As Double

        Dim D As Double

        Dim tmp_Photon As New Photon

        ' определяем ее отражательную, поглощательную и пропускательную способности

        If FTOP.R_dif.GetType.ToString <> "System.Double" Then

            R_dif = FTOP.R_dif.return_value(Argument).Coord(0)

        Else

            R_dif = FTOP.R_dif

        End If


        If FTOP.R_mir.GetType.ToString <> "System.Double" Then

            R_mir = FTOP.R_mir.return_value(Argument).Coord(0)

        Else

            R_mir = FTOP.R_mir

        End If


        If FTOP.A.GetType.ToString <> "System.Double" Then

            A = FTOP.A.return_value(Argument).Coord(0)

        Else

            A = FTOP.A

        End If


        If FTOP.D.GetType.ToString <> "System.Double" Then

            D = FTOP.D.Return_value(Argument).Coord(0)

        Else

            D = FTOP.D

        End If



        '        ' определяем расстояние, пройденое фотоном от прошлой грани

        '        Dim L As Double = L_vector(in_Photon.Center - Me.Center)

        '        ' определяем ослабление фотона при прохождении через среду

        '        Dim K_depletion As Double

        '        K_depletion = in_Photon.Medium.Calc_k(Argument)

        '        tmp_Photon.Depletion = in_Photon.Depletion * Math.E ^ (-K_depletion * L)

        '        ' Определяем среды, в которых распространяется или может быть будет распространяться фотон

        '        Dim Medium_1, Medium_2 As MEDIUM1

        '        ' определяем свободная ли грань

        '        ' если фотон подходит изнутри элемента и у грани нет соседа
        '        If Me.Can_Radiate And Cos_angle_between_vectors(in_Photon.Direction, Me.oz) > 0 Then

        '            Medium_1 = Me.Element.MEDIUM

        '            Medium_2 = Me.Global_medium

        '            GoTo Skip_medium_checks

        '        End If

        '        If Me.Can_Radiate And Cos_angle_between_vectors(in_Photon.Direction, Me.oz) <= 0 Then

        '            Medium_1 = Me.Global_medium

        '            Medium_2 = Me.Element.MEDIUM

        '            GoTo Skip_medium_checks

        '        End If

        '        ' если фотон подходит изнутри элемента и у грани есть сосед
        '        If Not (Me.Can_Radiate) And Cos_angle_between_vectors(in_Photon.Direction, Me.oz) > 0 Then

        '            Medium_1 = Me.Element.MEDIUM

        '            Medium_2 = Me.Neighbour_face.Element.MEDIUM

        '            GoTo Skip_medium_checks

        '        End If

        '        If Not (Me.Can_Radiate) And Cos_angle_between_vectors(in_Photon.Direction, Me.oz) <= 0 Then

        '            Medium_1 = Me.Neighbour_face.Element.MEDIUM

        '            Medium_2 = Me.Element.MEDIUM

        '            GoTo Skip_medium_checks

        '        End If

        'Skip_medium_checks:

        Me.Q_R_income += in_Photon.Power



        ' уровни вероятности, при которых происходят события и случайное число
        Dim Lower_level, Upper_level, rnd_Value As Double


        ' "бросаем кубик" для всех возможных видов взаимодействия и определяем какое из них происходит

        rnd_Value = Main_form.Solver.MT_RND.genrand_real1

        ' проверяем, не случилось ли диффузного отражения
        Lower_level = 0

        Upper_level = Lower_level + R_dif

        If Lower_level <= rnd_Value And rnd_Value < Upper_level Then

            ' если оно случилось, то производим диффузное отражение фотона

            tmp_Photon = Diffuse_reflection(in_Photon)

            Interaction.N_R_dif += tmp_Photon.Depletion

            Me.Q_R_diffuse_reflected += tmp_Photon.Power

            ' принимаем во внимание, что часть фотона поглощается в элементе

            'Interaction.N_A += Me.FTOP.A * tmp_Photon.Depletion

            ' и энергия фотона соотвественно уменьшается

            'tmp_Photon.Depletion *= R_dif

            GoTo Final_interaction

        End If


        ' проверяем, не случилось ли зеркального отражения

        Lower_level = Upper_level

        Upper_level = Lower_level + R_mir

        If Lower_level <= rnd_Value And rnd_Value < Upper_level Then

            ' если оно случилось, то производим зеркальное отражение фотона

            tmp_Photon = Mirror_reflection(in_Photon)

            Interaction.N_R_mir += tmp_Photon.Depletion

            Me.Q_R_mirror_reflected += tmp_Photon.Power

            ' принимаем во внимание, что часть фотона поглощается в элементе

            'Interaction.N_A += Me.FTOP.A * tmp_Photon.Depletion

            ' и энергия фотона соотвественно уменьшается 

            'tmp_Photon.Depletion *= R_mir

            GoTo Final_interaction

        End If


        ' проверяем, не случилось ли преломления
        Lower_level = Upper_level

        Upper_level = Lower_level + D

        If Lower_level <= rnd_Value And rnd_Value < Upper_level Then

            ' если оно случилось, то производим преломление фотона

            'Dim tmp_Type = Mid(Me.Element.Element_property.GetType.ToString, _
            'InStrRev(Me.Element.Element_property.GetType.ToString, ".") + 1)

            'If tmp_Type = "PSOLID" Then

            '    tmp_Photon = Transmission(in_Photon, Medium_1, Medium_2, Temperature)

            '    tmp_Photon.Center = Me.Center

            '    Return tmp_Photon


            'Else

            'tmp_Photon = Transmission_flat(in_Photon)

            'End If

            tmp_Photon = Transmission_flat(in_Photon)

            Interaction.N_D += tmp_Photon.Depletion

            Me.Q_R_transmitted += tmp_Photon.Power

            ' принимаем во внимание, что часть фотона поглощается в элементе

            'Interaction.N_A += Me.FTOP.A * tmp_Photon.Depletion

            ' и энергия фотона соотвественно уменьшается 

            'tmp_Photon.Depletion *= D

            GoTo Final_interaction

        End If


        ' если мы дошли сюда, то надо производить поглощеньице фотончика

        Interaction.N_A += in_Photon.Depletion

        Me.Q_R_absorbed += in_Photon.Power

        ' запоминаем какой фотон поглотился

        'i_Absorption_act += 1

        'Absorption_act(i_Absorption_act) = in_Photon

        Return Nothing

Final_interaction:

        ' сюда мы попадем, если у нас произошло какое-либо отражение или преломление через тонкий элемент.
        ' в этих случаях мы остаемся в той же среде, в какой и входили

        ' рассчитываем координаты точки, в которую попадает фотон

        Dim t As Double = -(Me.oz.Coord(0) * (in_Photon.Center.Coord(0) - Me.Grid(0).Position.Coord(0)) + Me.oz.Coord(1) * (in_Photon.Center.Coord(1) - Me.Grid(0).Position.Coord(1)) + Me.oz.Coord(2) * (in_Photon.Center.Coord(2) - Me.Grid(0).Position.Coord(2))) / (Me.oz.Coord(0) * in_Photon.Direction.Coord(0) + Me.oz.Coord(1) * in_Photon.Direction.Coord(1) + Me.oz.Coord(2) * in_Photon.Direction.Coord(2))

        tmp_Photon.Center.Coord(0) = in_Photon.Center.Coord(0) + in_Photon.Direction.Coord(0) * t

        tmp_Photon.Center.Coord(1) = in_Photon.Center.Coord(1) + in_Photon.Direction.Coord(1) * t

        tmp_Photon.Center.Coord(2) = in_Photon.Center.Coord(2) + in_Photon.Direction.Coord(2) * t

        'tmp_Photon.Center = Me.Center

        tmp_Photon.Medium = in_Photon.Medium

        Return tmp_Photon

    End Function


    Public Function Diffuse_reflection(ByRef in_Photon As Photon) As Photon

        Dim tmp_Photon As New Photon

        'tmp_Photon.Depletion = in_Photon.Depletion

        tmp_Photon.Center = in_Photon.Center

        tmp_Photon.Medium = in_Photon.Medium

        tmp_Photon.Frequency = in_Photon.Frequency

        Dim Angle_1, Angle_2 As Double

        Angle_1 = Main_form.Solver.MT_RND.genrand_real1 * 2 * Math.PI

        Angle_2 = Main_form.Solver.MT_RND.genrand_real1 * Math.PI / 2

        tmp_Photon.Direction.Coord(0) = Math.Cos(Angle_1) * Math.Cos(Angle_2)

        tmp_Photon.Direction.Coord(1) = Math.Sin(Angle_1) * Math.Cos(Angle_2)

        tmp_Photon.Direction.Coord(2) = Math.Sin(Angle_2)

        tmp_Photon.Direction = Transform_from_face(tmp_Photon.Direction)

        tmp_Photon.Power = in_Photon.Power

        tmp_Photon.Number = in_Photon.Number

        Diffuse_reflection = tmp_Photon

    End Function


    Public Function Mirror_reflection(ByRef in_Photon As Photon) As Photon

        Dim tmp_Photon As New Photon

        'tmp_Photon.Depletion = in_Photon.Depletion

        tmp_Photon.Center = in_Photon.Center

        tmp_Photon.Medium = in_Photon.Medium

        tmp_Photon.Frequency = in_Photon.Frequency

        tmp_Photon.Direction = Transform_to_face(in_Photon.Direction)

        tmp_Photon.Direction.Coord(2) = -tmp_Photon.Direction.Coord(2)

        tmp_Photon.Direction = Transform_from_face(tmp_Photon.Direction)

        tmp_Photon.Power = in_Photon.Power

        tmp_Photon.Number = in_Photon.Number

        Mirror_reflection = tmp_Photon

    End Function

    ''' <summary>
    ''' This function return mirrorly transmitted photon, taking into account total reflection
    ''' </summary>
    ''' <param name="in_Photon"> Incoming photon</param>
    ''' <param name="Medium_1">Medium from incoming side</param>
    ''' <param name="Medium_2">Medium from outcoming side</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Transmission(ByRef in_Photon As Photon, ByRef Medium_1 As Object, ByRef Medium_2 As Object, ByRef Temperature As Double) As Photon

        Dim n_1, n_2 As Double

        ' рассчитываем коэффициент преломления во входящей части полусферы

        n_1 = Medium_1.Calc_n(in_Photon.Direction, Element.element_property.coordinate_system, Temperature)

        ' рассчитываем коэффициент преломления в выходящей части полусферы

        n_2 = Medium_2.Calc_n(in_Photon.Direction, Element.element_property.coordinate_system, Temperature)


        Dim tmp_Photon As New Photon

        Dim xy2 As Double

        'tmp_Photon.Depletion = in_Photon.Depletion

        tmp_Photon.Center = in_Photon.Center

        tmp_Photon.Medium = in_Photon.Medium

        tmp_Photon.Frequency = in_Photon.Frequency

        tmp_Photon.Direction = Transform_to_face(in_Photon.Direction)

        tmp_Photon.Direction.Coord(0) = n_1 / n_2 * tmp_Photon.Direction.Coord(0)

        tmp_Photon.Direction.Coord(1) = n_1 / n_2 * tmp_Photon.Direction.Coord(1)

        xy2 = tmp_Photon.Direction.Coord(0) ^ 2 + tmp_Photon.Direction.Coord(1) ^ 2

        If xy2 <= 1 Then

            tmp_Photon.Direction.Coord(2) = Math.Sign(tmp_Photon.Direction.Coord(2)) * Math.Sqrt(1 - xy2)

            tmp_Photon.Medium = Medium_2

        Else

            tmp_Photon.Direction.Coord(2) = -Math.Sign(tmp_Photon.Direction.Coord(2)) * Math.Sqrt(xy2 - 1)

            tmp_Photon.Medium = Medium_1

        End If

        tmp_Photon.Direction = Norm_vector(tmp_Photon.Direction)

        tmp_Photon.Direction = Transform_from_face(tmp_Photon.Direction)

        Transmission = tmp_Photon

    End Function

    ''' <summary>
    ''' This function return photon, mirrorly transmitted from flat element's face 
    ''' </summary>
    ''' <param name="in_Photon"> Incoming photon</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Transmission_flat(ByRef in_Photon As Photon) As Photon

        Dim tmp_Photon As New Photon

        'tmp_Photon.Depletion = in_Photon.Depletion

        tmp_Photon.Center = in_Photon.Center

        tmp_Photon.Medium = in_Photon.Medium

        tmp_Photon.Frequency = in_Photon.Frequency

        tmp_Photon.Direction.Coord(0) = in_Photon.Direction.Coord(0)

        tmp_Photon.Direction.Coord(1) = in_Photon.Direction.Coord(1)

        tmp_Photon.Direction.Coord(2) = in_Photon.Direction.Coord(2)

        tmp_Photon.Power = in_Photon.Power

        tmp_Photon.Number = in_Photon.Number

        Transmission_flat = tmp_Photon

    End Function

    ''' <summary>
    ''' This subroutine caclulates face's grids coordinates in face's coordinate system and detects lower and upper bounds of X- and Y-coordinates of grids.  
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Calc_local_coordinates()

        Dim Local_grid(UBound(Grid)) As Vector

        ReDim Xpi(UBound(Grid))

        ReDim Ypi(UBound(Grid))

        Xmax = -Double.MaxValue

        Ymax = -Double.MaxValue

        Xmin = Double.MaxValue

        Ymin = Double.MaxValue

        For i As Integer = 0 To UBound(Grid)

            Local_grid(i) = New Vector

            Local_grid(i).Coord(0) = Grid(i).Position.Coord(0)

            Local_grid(i).Coord(1) = Grid(i).Position.Coord(1)

            Local_grid(i).Coord(2) = Grid(i).Position.Coord(2)

            Local_grid(i) -= Center

            Local_grid(i) = Transform_to_face(Local_grid(i))

            Xpi(i) = Local_grid(i).Coord(0)

            Ypi(i) = Local_grid(i).Coord(1)

            ' Max and min coordinates determination

            If Xpi(i) > Xmax Then

                Xmax = Xpi(i)

            End If

            If Xpi(i) < Xmin Then

                Xmin = Xpi(i)

            End If

            If Ypi(i) > Ymax Then

                Ymax = Ypi(i)

            End If

            If Ypi(i) < Ymin Then

                Ymin = Ypi(i)

            End If

        Next i

        'If Me.Element.Number = 125 Then

        '    Stop

        'End If

        'If Xmax = -Double.MaxValue Then

        '    Xmax = Xmax

        'End If


    End Sub

    ''' <summary>
    ''' This function returns emission center coordinates ramdomly located on face surface
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Calc_random_emission_center() As Vector
        ' Random point generation



        Dim RND_point As New Vector

        Do

            RND_point.Coord(0) = (Xmax - Xmin) * Main_form.Solver.MT_RND.genrand_real1 + Xmin

            RND_point.Coord(1) = (Ymax - Ymin) * Main_form.Solver.MT_RND.genrand_real1 + Ymin

            ' if point is located on the face, then exit do
            If IsPointInPolygon(RND_point.Coord(0), RND_point.Coord(1), Xpi, Ypi) Then Exit Do

        Loop

        ' Transform random point back to global coordinate system

        RND_point.Coord(2) = 0

        RND_point = Transform_from_face(RND_point)

        RND_point += Center

        Return RND_point

    End Function

    ''' <summary>
    ''' This function return is point with coordinates (x,y) lies in polygon (convex or not), given by array of its 
    ''' vertex coordinates, or not lies.
    ''' </summary>
    ''' <param name="x">X coordinate of investigated point</param>
    ''' <param name="y">Y coordinate of investigated point</param>
    ''' <returns>Is point with coordinates (x,y) lies in polygon (convex or not), given by array of its 
    ''' vertex coordinates or not lies.</returns>
    ''' <remarks></remarks>
    Public Function IsPointInPolygon(ByRef x As Double, ByRef y As Double, ByRef Xpi() As Double, ByRef Ypi() As Double) As Boolean
        'Dim Result As Boolean
        Dim n As Long
        Dim b1 As Boolean
        Dim b2 As Boolean

        n = 0
        'Result = False
        IsPointInPolygon = False
        Do
            b1 = y > Ypi(n)
            b2 = y <= Ypi(n + 1)
            If Not (b1 And Not b2 Or b2 And Not b1) Then
                If x - Xpi(n) < (y - Ypi(n)) * (Xpi(n + 1) - Xpi(n)) / (Ypi(n + 1) - Ypi(n)) Then
                    'Result = Not Result
                    IsPointInPolygon = Not IsPointInPolygon
                End If
            End If
            n = n + 1
        Loop Until Not n <= UBound(Xpi) - 1

        'IsPointInPolygon = Result
    End Function




End Class



