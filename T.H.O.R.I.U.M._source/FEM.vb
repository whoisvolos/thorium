Imports System
Imports System.IO

''' <summary>
''' FEM any object (i.e. any finite element, coordinate system, material, property or something else)
''' </summary>
''' <remarks></remarks>
<Serializable()> _
Public MustInherit Class Model_item

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

        If InStr(in_String, "BEGIN BULK") > 1 Then

            Position = InStr(in_String, "BEGIN BULK") + Len("BEGIN BULK")

        Else

            Position = 1

        End If

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
Public MustInherit Class Material

    Inherits Model_item

    ''' <summary>
    ''' Stiffness matrix of material.
    ''' For each type of material it must be resized.
    ''' </summary>
    ''' <remarks></remarks>
    Public G(,) As Double

    ''' <summary>
    ''' Poisson's ratio
    ''' </summary>
    ''' <remarks></remarks>
    Public Nu As Double

    ''' <summary>
    ''' Thermal expansion coefficients matrix
    ''' For each type of material it must be resized.
    ''' </summary>
    ''' <remarks></remarks>
    Public A(,) As Double

    ''' <summary>
    ''' Density of material
    ''' </summary>
    ''' <remarks></remarks>
    Public Rho As Double

    ''' <summary>
    ''' The properties of material is defined for this temperature
    ''' </summary>
    ''' <remarks></remarks>
    Public T_reference As Double

    ''' <summary>
    ''' Ultimate tension stress
    ''' </summary>
    ''' <remarks></remarks>
    Public Ultimate_tension_stress() As Double

    ''' <summary>
    ''' Ultimate compression stress
    ''' </summary>
    ''' <remarks></remarks>
    Public Ultimate_compression_stress() As Double

    ''' <summary>
    ''' Ultimate shear stress
    ''' </summary>
    ''' <remarks></remarks>
    Public Ultimate_shear_stress As Double

    ''' <summary>
    ''' Damping coefficient
    ''' </summary>
    ''' <remarks></remarks>
    Public Damping_coefficient As Double

    Public Sub New(Optional ByRef String_in_file As String = "")

        MyBase.New(String_in_file)

    End Sub

End Class

''' <summary>
''' G(0,0) - Elasticity modulus, E
''' G(1,1) - Shear modulus, G
''' </summary>
''' <remarks></remarks>
<Serializable()> _
Public Class MAT1

    Inherits Material

    Public Coordinate_system As Long

    Public Overrides Sub Load()

        On Error Resume Next

        ReDim G(1, 1)

        ReDim A(0, 0)

        ReDim Ultimate_tension_stress(0)

        ReDim Ultimate_compression_stress(0)

        Number = Val(Fields_array(1))

        G(0, 0) = Val(Fields_array(2))

        G(1, 1) = Val(Fields_array(3))

        Nu = Val(Fields_array(4))

        Rho = Val(Fields_array(5))

        A(0, 0) = Val(Fields_array(6))

        T_reference = Val(Fields_array(7))

        Damping_coefficient = Val(Fields_array(8))

        Ultimate_tension_stress(0) = Val(Fields_array(11))

        Ultimate_compression_stress(0) = Val(Fields_array(12))

        Ultimate_shear_stress = Val(Fields_array(13))

        Coordinate_system = Val(Fields_array(14))

        On Error GoTo 0

    End Sub
    Public Overrides Sub Prepare_Save()

        ReDim Fields_array(14)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(Number)

        Fields_array(2) = my_Str(G(0, 0))

        Fields_array(3) = my_Str(G(1, 1))

        Fields_array(4) = my_Str(Nu)

        Fields_array(5) = my_Str(Rho)

        Fields_array(6) = my_Str(A(0, 0))

        Fields_array(7) = my_Str(T_reference)

        Fields_array(8) = my_Str(Damping_coefficient)

        Fields_array(11) = my_Str(Ultimate_tension_stress(0))

        Fields_array(12) = my_Str(Ultimate_compression_stress(0))

        Fields_array(13) = my_Str(Ultimate_shear_stress)

        Fields_array(14) = my_Str(Coordinate_system)

        On Error GoTo 0

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
''' G(0,0) - G11
''' G(0,1) - G12
''' G(0,2) - G13
''' G(1,1) - G22
''' G(1,2) - G23
''' G(2,2) - G33
''' A(0,0) - A1
''' A(1,1) - A2
''' A(2,2) - A3
''' </summary>
''' <remarks></remarks>
<Serializable()> _
Public Class MAT2

    Inherits Material

    Public Coordinate_system As Long

    Public Overrides Sub Load()

        On Error Resume Next

        ReDim G(2, 2)

        ReDim A(2, 2)

        ReDim Ultimate_tension_stress(0)

        ReDim Ultimate_compression_stress(0)

        Number = Val(Fields_array(1))

        G(0, 0) = Val(Fields_array(2))

        G(0, 1) = Val(Fields_array(3))

        G(0, 2) = Val(Fields_array(4))

        G(1, 1) = Val(Fields_array(5))

        G(1, 2) = Val(Fields_array(6))

        G(2, 2) = Val(Fields_array(7))

        Rho = Val(Fields_array(8))

        A(0, 0) = Val(Fields_array(11))

        A(1, 1) = Val(Fields_array(12))

        A(2, 2) = Val(Fields_array(13))

        T_reference = Val(Fields_array(14))

        Damping_coefficient = Val(Fields_array(15))

        Ultimate_tension_stress(0) = Val(Fields_array(16))

        Ultimate_compression_stress(0) = Val(Fields_array(17))

        Ultimate_shear_stress = Val(Fields_array(18))

        Coordinate_system = Val(Fields_array(21))

        On Error GoTo 0

    End Sub

    Public Overrides Sub Prepare_Save()

        ReDim Fields_array(21)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(Number)

        Fields_array(2) = my_Str(G(0, 0))

        Fields_array(3) = my_Str(G(0, 1))

        Fields_array(4) = my_Str(G(0, 2))

        Fields_array(5) = my_Str(G(1, 1))

        Fields_array(6) = my_Str(G(1, 2))

        Fields_array(7) = my_Str(G(2, 2))

        Fields_array(8) = my_Str(Rho)

        Fields_array(11) = my_Str(A(0, 0))

        Fields_array(12) = my_Str(A(1, 1))

        Fields_array(13) = my_Str(A(2, 2))

        Fields_array(14) = my_Str(T_reference)

        Fields_array(15) = my_Str(Damping_coefficient)

        Fields_array(16) = my_Str(Ultimate_tension_stress(0))

        Fields_array(17) = my_Str(Ultimate_compression_stress(0))

        Fields_array(18) = my_Str(Ultimate_shear_stress)

        Fields_array(21) = my_Str(Coordinate_system)

        On Error GoTo 0
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
Public Class MAT4

    Inherits Material

    Public lambda As Double

    Public c As Double

    Public Coordinate_system As Long

    Public Overrides Sub Load()

        On Error Resume Next

        Number = Val(Fields_array(1))

        lambda = Val(Fields_array(2))

        c = Val(Fields_array(3))

        Rho = Val(Fields_array(4))

        On Error GoTo 0

    End Sub

    Public Overrides Sub Prepare_Save()

        ReDim Fields_array(4)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(Number)

        Fields_array(2) = my_Str(lambda)

        Fields_array(3) = my_Str(c)

        Fields_array(4) = my_Str(Rho)

        On Error GoTo 0
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
''' G(0,0) - E1 - longitudinal direction
''' G(1,1) - E2 - transverse direction
''' G(0,1) - G12
''' G(0,2) - G1Z
''' G(1,2) - G2Z
''' A(0,0) - A1
''' A(1,1) - A2
''' Ultimate_tension_stress(0)- Xt- longitudinal direction
''' Ultimate_compression_stress(0) - Xc- longitudinal direction
''' Ultimate_tension_stress(1)- Yt- transverse direction
''' Ultimate_compression_stress(1) - Yc- transverse direction
'''  </summary>
''' <remarks></remarks>
<Serializable()> _
Public Class MAT8

    Inherits Material

    ''' <summary>
    ''' Flag, for declaration of theory of limit deformations usage
    ''' </summary> 
    ''' <remarks></remarks>
    Public STRN As String

    ''' <summary>
    ''' For Tsay-Wu theory
    ''' </summary>
    ''' <remarks></remarks>
    Public F12 As Double

    Public Overrides Sub Load()

        On Error Resume Next

        ReDim G(1, 1)

        ReDim A(0, 0)

        ReDim Ultimate_tension_stress(1)

        ReDim Ultimate_compression_stress(1)


        Number = Val(Fields_array(1))

        G(0, 0) = Val(Fields_array(2))

        G(1, 1) = Val(Fields_array(3))

        Nu = Val(Fields_array(4))

        G(0, 1) = Val(Fields_array(5))

        G(0, 2) = Val(Fields_array(6))

        G(1, 2) = Val(Fields_array(7))

        Rho = Val(Fields_array(8))


        A(0, 0) = Val(Fields_array(11))

        A(1, 1) = Val(Fields_array(12))

        T_reference = Val(Fields_array(13))

        Ultimate_tension_stress(0) = Val(Fields_array(14))

        Ultimate_compression_stress(0) = Val(Fields_array(15))

        Ultimate_tension_stress(1) = Val(Fields_array(16))

        Ultimate_compression_stress(1) = Val(Fields_array(17))

        Ultimate_shear_stress = Val(Fields_array(18))


        Damping_coefficient = Val(Fields_array(21))

        F12 = Val(Fields_array(22))

        STRN = Trim(Fields_array(23))

        On Error GoTo 0

    End Sub

    Public Overrides Sub Prepare_Save()

        ReDim Fields_array(23)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(Number)

        Fields_array(2) = my_Str(G(0, 0))

        Fields_array(3) = my_Str(G(1, 1))

        Fields_array(4) = my_Str(Nu)

        Fields_array(5) = my_Str(G(0, 1))

        Fields_array(6) = my_Str(G(0, 2))

        Fields_array(7) = my_Str(G(1, 2))

        Fields_array(8) = my_Str(Rho)


        Fields_array(11) = my_Str(A(0, 0))

        Fields_array(12) = my_Str(A(1, 1))

        Fields_array(13) = my_Str(T_reference)

        Fields_array(14) = my_Str(Ultimate_tension_stress(0))

        Fields_array(15) = my_Str(Ultimate_compression_stress(0))

        Fields_array(16) = my_Str(Ultimate_tension_stress(1))

        Fields_array(17) = my_Str(Ultimate_compression_stress(1))

        Fields_array(18) = my_Str(Ultimate_shear_stress)


        Fields_array(21) = my_Str(Damping_coefficient)

        Fields_array(22) = my_Str(F12)

        Fields_array(23) = STRN

        On Error GoTo 0

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
''' G(0,0) - G11
''' G(0,1) - G12
''' G(0,2) - G13
''' G(0,3) - G14
''' G(0,4) - G15
''' G(0,5) - G16
''' 
''' G(1,1) - G22
''' G(1,2) - G23
''' G(1,3) - G24
''' G(1,4) - G25
''' G(1,5) - G26
''' 
''' G(2,2) - G33
''' G(2,3) - G34
''' G(2,4) - G35
''' G(2,5) - G36
''' 
''' G(3,3) - G44
''' G(3,4) - G45
''' G(3,5) - G46
''' 
''' G(4,4) - G55
''' G(3,5) - G56
''' 
''' G(5,5) - G66
'''
''' A(0,0) - A1
''' A(1,1) - A2
''' A(2,2) - A3
''' A(3,3) - A4
''' A(4,4) - A5
''' A(5,5) - A6
'''  </summary>
''' <remarks></remarks>
<Serializable()> _
Public Class MAT9

    Inherits Material

    Public Overrides Sub Load()

        On Error Resume Next

        ReDim G(5, 5)

        ReDim A(5, 5)

        Number = Val(Fields_array(1))

        G(0, 0) = Val(Fields_array(2))

        G(0, 1) = Val(Fields_array(3))

        G(0, 2) = Val(Fields_array(4))

        G(0, 3) = Val(Fields_array(5))

        G(0, 4) = Val(Fields_array(6))

        G(0, 5) = Val(Fields_array(7))

        G(1, 1) = Val(Fields_array(8))


        G(1, 2) = Val(Fields_array(11))

        G(1, 3) = Val(Fields_array(12))

        G(1, 4) = Val(Fields_array(13))

        G(1, 5) = Val(Fields_array(14))

        G(2, 2) = Val(Fields_array(15))

        G(2, 3) = Val(Fields_array(16))

        G(2, 4) = Val(Fields_array(17))

        G(2, 5) = Val(Fields_array(18))


        G(3, 3) = Val(Fields_array(21))

        G(3, 4) = Val(Fields_array(22))

        G(3, 5) = Val(Fields_array(23))

        G(4, 4) = Val(Fields_array(24))

        G(4, 5) = Val(Fields_array(25))

        G(5, 5) = Val(Fields_array(26))

        Rho = Val(Fields_array(27))

        A(0, 0) = Val(Fields_array(28))



        A(1, 1) = Val(Fields_array(31))

        A(2, 2) = Val(Fields_array(32))

        A(3, 3) = Val(Fields_array(33))

        A(4, 4) = Val(Fields_array(34))

        A(5, 5) = Val(Fields_array(35))

        T_reference = Val(Fields_array(36))

        On Error GoTo 0

    End Sub

    Public Overrides Sub Prepare_Save()

        ReDim Fields_array(36)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(Number)

        Fields_array(2) = my_Str(G(0, 0))

        Fields_array(3) = my_Str(G(0, 1))

        Fields_array(4) = my_Str(G(0, 2))

        Fields_array(5) = my_Str(G(0, 3))

        Fields_array(6) = my_Str(G(0, 4))

        Fields_array(7) = my_Str(G(0, 5))

        Fields_array(8) = my_Str(G(1, 1))


        Fields_array(11) = my_Str(G(1, 2))

        Fields_array(12) = my_Str(G(1, 3))

        Fields_array(13) = my_Str(G(1, 4))

        Fields_array(14) = my_Str(G(1, 5))

        Fields_array(15) = my_Str(G(2, 2))

        Fields_array(16) = my_Str(G(2, 3))

        Fields_array(17) = my_Str(G(2, 4))

        Fields_array(18) = my_Str(G(2, 5))


        Fields_array(21) = my_Str(G(3, 3))

        Fields_array(22) = my_Str(G(3, 4))

        Fields_array(23) = my_Str(G(3, 5))

        Fields_array(24) = my_Str(G(4, 4))

        Fields_array(25) = my_Str(G(4, 5))

        Fields_array(26) = my_Str(G(5, 5))

        Fields_array(27) = my_Str(Rho)

        Fields_array(28) = my_Str(A(0, 0))



        Fields_array(31) = my_Str(A(1, 1))

        Fields_array(32) = my_Str(A(2, 2))

        Fields_array(33) = my_Str(A(3, 3))

        Fields_array(34) = my_Str(A(4, 4))

        Fields_array(35) = my_Str(A(5, 5))

        Fields_array(36) = my_Str(T_reference)

        On Error GoTo 0

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
Public MustInherit Class Coordinate_system

    Inherits Model_item

    ' система координат, с помощью которой определяется указанная система координат
    Public Coordinate_system As Object

    ' координаты точки системы координат
    Public point_A_vector As Vector

    Public point_B_vector As Vector

    Public point_C_vector As Vector

    ' номера опорных точек системы координат
    Public point_A As Object

    Public point_B As Object

    Public point_C As Object

    ' матрица перехода из системы координат в глобальную
    Public from_0(2, 2) As Double

    ' матрица перехода из глобальной системы координат
    Public to_0(2, 2) As Double

    ''' <summary>
    ''' Calculation transition matres from and to Cartesian rectangular coordinates.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Calculate_transition_matrices()

        Dim ox As New Vector
        Dim oy As New Vector
        Dim oz As New Vector

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


        Dim AB As Vector
        Dim AC As Vector

        If Number = 0 Then



            from_0(0, 0) = 1
            from_0(0, 1) = 0
            from_0(0, 2) = 0

            from_0(1, 0) = 0
            from_0(1, 1) = 1
            from_0(1, 2) = 0


            from_0(2, 0) = 0
            from_0(2, 1) = 0
            from_0(2, 2) = 1



            to_0(0, 0) = 1
            to_0(0, 1) = 0
            to_0(0, 2) = 0

            to_0(1, 0) = 0
            to_0(1, 1) = 1
            to_0(1, 2) = 0


            to_0(2, 0) = 0
            to_0(2, 1) = 0
            to_0(2, 2) = 1

            Exit Sub

        End If

        AB = point_B_vector - point_A_vector
        AC = point_C_vector - point_A_vector
        AC = Norm_vector(AC)

        oz = Norm_vector(AB)
        oy = oz ^ AC
        oy = Norm_vector(oy)
        ox = oy ^ oz
        ox = Norm_vector(ox)

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

    Public Function Transform(ByVal in_Vector As Vector, ByVal in_Coord As Coordinate_system, ByVal out_Coord As Coordinate_system) As Vector

        Dim tmp_Vector_1, tmp_Vector_2 As New Vector

        ReDim tmp_Vector_1.Coord(2), tmp_Vector_2.Coord(2)

        in_Vector = in_Coord.Transform_from_special_coordinates_to_local_rectangular(in_Vector)

        For i_1 As Integer = 0 To 2

            For j_1 As Integer = 0 To 2

                tmp_Vector_1.Coord(i_1) = tmp_Vector_1.Coord(i_1) + in_Vector.Coord(j_1) * in_Coord.to_0(i_1, j_1)

            Next j_1


        Next i_1

        If out_Coord.Number = 0 Then

            Transform = out_Coord.Transform_to_special_coordinates_from_local_rectangular(tmp_Vector_1)

            Exit Function

        End If

        For i_1 As Integer = 0 To 2

            For j_1 As Integer = 0 To 2

                tmp_Vector_2.Coord(i_1) = tmp_Vector_2.Coord(i_1) + tmp_Vector_1.Coord(j_1) * out_Coord.from_0(i_1, j_1)

            Next j_1


        Next i_1

        Transform = out_Coord.Transform_to_special_coordinates_from_local_rectangular(tmp_Vector_2)

    End Function

    Public Function Transform(ByVal in_Vector As Vector, ByVal in_Coord As Object, ByVal out_Coord As Object) As Vector

        On Error Resume Next

        Dim tmp_Vector_1, tmp_Vector_2 As New Vector

        ReDim tmp_Vector_1.Coord(2), tmp_Vector_2.Coord(2)

        If in_Coord = 0 Then

            For i_1 As Integer = 0 To 2

                For j_1 As Integer = 0 To 2

                    tmp_Vector_1.Coord(i_1) = tmp_Vector_1.Coord(i_1) + in_Vector.Coord(j_1) * out_Coord.from_0(i_1, j_1)

                Next j_1


            Next i_1

            Return tmp_Vector_1

        End If

        in_Vector = in_Coord.Transform_from_special_coordinates_to_local_rectangular(in_Vector)

        For i_1 As Integer = 0 To 2

            For j_1 As Integer = 0 To 2

                tmp_Vector_1.Coord(i_1) = tmp_Vector_1.Coord(i_1) + in_Vector.Coord(j_1) * in_Coord.to_0(i_1, j_1)

            Next j_1

        Next i_1

        If out_Coord = 0 Then

            Transform = tmp_Vector_1

            Exit Function

        End If

        For i_1 As Integer = 0 To 2

            For j_1 As Integer = 0 To 2

                tmp_Vector_2.Coord(i_1) = tmp_Vector_2.Coord(i_1) + tmp_Vector_1.Coord(j_1) * out_Coord.from_0(i_1, j_1)

            Next j_1


        Next i_1

        Transform = out_Coord.Transform_to_special_coordinates_from_local_rectangular(tmp_Vector_2)

        On Error GoTo 0

    End Function

    Public Function Transform(ByVal in_Matrix(,) As Double, ByVal in_Coord As Coordinate_system, ByVal out_Coord As Coordinate_system) As Double(,)

        If InStr(in_Coord.Type, "R") = 3 And InStr(out_Coord.Type, "R") = 3 Then

            Dim tmp_Array_1(0 To 2, 0 To 2), tmp_Array_2(0 To 2, 0 To 2), tmp_1 As Double

            ' переводим в глобальную систему координат

            tmp_1 = 0

            For i_1 As Integer = 0 To 2
                For j_1 As Integer = 0 To 2

                    For i_2 As Integer = 0 To 2
                        For j_2 As Integer = 0 To 2

                            tmp_1 = tmp_1 + in_Matrix(i_2, j_2) * in_Coord.to_0(i_1, i_2) * in_Coord.to_0(j_1, j_2)

                        Next j_2
                    Next i_2

                    tmp_Array_1(i_1, j_1) = tmp_1

                    tmp_1 = 0

                Next j_1
            Next i_1

            If out_Coord.Number = 0 Then

                Transform = tmp_Array_1

                Exit Function

            End If


            ' переводим из глобальной системы координат в Coord_sys_2

            tmp_1 = 0

            For i_1 As Integer = 0 To 2
                For j_1 As Integer = 0 To 2

                    For i_2 As Integer = 0 To 2
                        For j_2 As Integer = 0 To 2

                            tmp_1 = tmp_1 + tmp_Array_1(i_2, j_2) * out_Coord.from_0(i_1, i_2) * out_Coord.from_0(j_1, j_2)

                        Next j_2
                    Next i_2

                    tmp_Array_2(i_1, j_1) = tmp_1

                    tmp_1 = 0

                Next j_1
            Next i_1

            Transform = tmp_Array_2
        Else

            Transform = in_Matrix

        End If

    End Function

    Public MustOverride Function Transform_to_special_coordinates_from_local_rectangular(ByVal in_Vector As Vector) As Vector

    Public MustOverride Function Transform_from_special_coordinates_to_local_rectangular(ByVal in_Vector As Vector) As Vector

    Public Overrides Sub Load()

        Dim tmp_point_A, tmp_point_B, tmp_point_C As New Vector

        On Error Resume Next

        Number = Val(Fields_array(1))

        Coordinate_system = Val(Fields_array(2))

        If InStr(Type, "1") > 0 Then

            point_A = Val(Fields_array(3))

            point_B = Val(Fields_array(4))

            point_C = Val(Fields_array(5))

        End If

        If InStr(Type, "2") > 0 Then

            tmp_point_A.Coord(0) = Val(Fields_array(3))

            tmp_point_A.Coord(1) = Val(Fields_array(4))

            tmp_point_A.Coord(2) = Val(Fields_array(5))


            tmp_point_B.Coord(0) = Val(Fields_array(6))

            tmp_point_B.Coord(1) = Val(Fields_array(7))

            tmp_point_B.Coord(2) = Val(Fields_array(8))


            tmp_point_C.Coord(0) = Val(Fields_array(11))

            tmp_point_C.Coord(1) = Val(Fields_array(12))

            tmp_point_C.Coord(2) = Val(Fields_array(13))



            'point_A.Coord(0) = Val(Fields_array(3))

            'point_A.Coord(1) = Val(Fields_array(4))

            'point_A.Coord(2) = Val(Fields_array(5))


            'point_B.Coord(0) = Val(Fields_array(6))

            'point_B.Coord(1) = Val(Fields_array(7))

            'point_B.Coord(2) = Val(Fields_array(8))


            'point_C.Coord(0) = Val(Fields_array(11))

            'point_C.Coord(1) = Val(Fields_array(12))

            'point_C.Coord(2) = Val(Fields_array(13))

            point_A_vector = tmp_point_A

            point_B_vector = tmp_point_B

            point_C_vector = tmp_point_C

        End If

        On Error GoTo 0

    End Sub

    Public Overrides Sub Prepare_Save()

        Dim tmp_Number As Long

        ReDim Fields_array(13)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(Number)

        Fields_array(2) = my_Str(Coordinate_system)

        If InStr(Type, "1") > 0 Then

            Fields_array(3) = my_Str(point_A)

            Fields_array(4) = my_Str(point_B)

            Fields_array(5) = my_Str(point_C)

        End If

        If InStr(Type, "2") > 0 Then

            Fields_array(3) = my_Str(point_A.Coord(0))

            Fields_array(4) = my_Str(point_A.Coord(1))

            Fields_array(5) = my_Str(point_A.Coord(2))


            Fields_array(6) = my_Str(point_B.Coord(0))

            Fields_array(7) = my_Str(point_B.Coord(1))

            Fields_array(8) = my_Str(point_B.Coord(2))


            Fields_array(11) = my_Str(point_C.Coord(0))

            Fields_array(12) = my_Str(point_C.Coord(1))

            Fields_array(13) = my_Str(point_C.Coord(2))

        End If

        On Error GoTo 0

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
Public Class CORD1R

    Inherits Coordinate_system



    Public Overrides Function Transform_from_special_coordinates_to_local_rectangular(ByVal in_Vector As Vector) As Vector

        Transform_from_special_coordinates_to_local_rectangular = in_Vector

    End Function

    Public Overrides Function Transform_to_special_coordinates_from_local_rectangular(ByVal in_Vector As Vector) As Vector

        Transform_to_special_coordinates_from_local_rectangular = in_Vector

    End Function

    Public Sub New(Optional ByRef String_in_file As String = "")

        MyBase.New(String_in_file)

    End Sub

    Public Sub New(ByRef Fields_array() As String)

        Me.Fields_array = Fields_array

        Load()

    End Sub
End Class
<Serializable()> _
Public Class CORD2R

    Inherits Coordinate_system

    Public Overrides Function Transform_from_special_coordinates_to_local_rectangular(ByVal in_Vector As Vector) As Vector

        Transform_from_special_coordinates_to_local_rectangular = in_Vector

    End Function

    Public Overrides Function Transform_to_special_coordinates_from_local_rectangular(ByVal in_Vector As Vector) As Vector

        Transform_to_special_coordinates_from_local_rectangular = in_Vector

    End Function


    Public Sub New(Optional ByRef String_in_file As String = "")

        MyBase.New(String_in_file)

    End Sub

    Public Sub New(ByRef Fields_array() As String)

        Me.Fields_array = Fields_array

        Load()

    End Sub
End Class

''' <summary>
''' Cylindrical coordinate system.
''' Vector's component
''' (0) - R
''' (1) - theta
''' (2) - z
''' </summary>
''' <remarks></remarks>
<Serializable()> _
Public Class CORD1C

    Inherits Coordinate_system

    Public Overrides Function Transform_from_special_coordinates_to_local_rectangular(ByVal in_Vector As Vector) As Vector

        Dim tmp_Vector As New Vector

        tmp_Vector.Coord(0) = in_Vector.Coord(0) * Math.Cos(in_Vector.Coord(1))

        tmp_Vector.Coord(1) = in_Vector.Coord(0) * Math.Sin(in_Vector.Coord(1))

        tmp_Vector.Coord(2) = in_Vector.Coord(2)

        Transform_from_special_coordinates_to_local_rectangular = tmp_Vector

    End Function

    Public Overrides Function Transform_to_special_coordinates_from_local_rectangular(ByVal in_Vector As Vector) As Vector

        Dim tmp_Vector As New Vector

        If in_Vector.Coord(0) = in_Vector.Coord(1) And in_Vector.Coord(1) = in_Vector.Coord(2) Then

            tmp_Vector.Coord(0) = 0

            tmp_Vector.Coord(1) = 0

            tmp_Vector.Coord(2) = 0

            Transform_to_special_coordinates_from_local_rectangular = tmp_Vector

            Exit Function


        End If

        tmp_Vector.Coord(0) = Math.Sqrt(in_Vector.Coord(0) ^ 2 + in_Vector.Coord(1) ^ 2)

        tmp_Vector.Coord(1) = Math.Atan2(in_Vector.Coord(1), in_Vector.Coord(0))

        tmp_Vector.Coord(2) = in_Vector.Coord(2)

        Transform_to_special_coordinates_from_local_rectangular = tmp_Vector

    End Function


    Public Sub New(Optional ByRef String_in_file As String = "")

        MyBase.New(String_in_file)

    End Sub

    Public Sub New(ByRef Fields_array() As String)

        Me.Fields_array = Fields_array

        Load()

    End Sub
End Class

''' <summary>
''' Cylindrical coordinate system.
''' Vector's component
''' (0) - R
''' (1) - theta
''' (2) - z
''' </summary>
''' <remarks></remarks>
<Serializable()> _
Public Class CORD2C

    Inherits Coordinate_system

    Public Overrides Function Transform_from_special_coordinates_to_local_rectangular(ByVal in_Vector As Vector) As Vector

        Dim tmp_Vector As New Vector

        tmp_Vector.Coord(0) = in_Vector.Coord(0) * Math.Cos(in_Vector.Coord(1))

        tmp_Vector.Coord(1) = in_Vector.Coord(0) * Math.Sin(in_Vector.Coord(1))

        tmp_Vector.Coord(2) = in_Vector.Coord(2)

        Transform_from_special_coordinates_to_local_rectangular = tmp_Vector

    End Function

    Public Overrides Function Transform_to_special_coordinates_from_local_rectangular(ByVal in_Vector As Vector) As Vector

        Dim tmp_Vector As New Vector

        If in_Vector.Coord(0) = in_Vector.Coord(1) And in_Vector.Coord(1) = in_Vector.Coord(2) Then

            tmp_Vector.Coord(0) = 0

            tmp_Vector.Coord(1) = 0

            tmp_Vector.Coord(2) = 0

            Transform_to_special_coordinates_from_local_rectangular = tmp_Vector

            Exit Function


        End If

        tmp_Vector.Coord(0) = Math.Sqrt(in_Vector.Coord(0) ^ 2 + in_Vector.Coord(1) ^ 2)

        tmp_Vector.Coord(1) = Math.Atan2(in_Vector.Coord(1), in_Vector.Coord(0))

        tmp_Vector.Coord(2) = in_Vector.Coord(2)

        Transform_to_special_coordinates_from_local_rectangular = tmp_Vector

    End Function

    Public Sub New(Optional ByRef String_in_file As String = "")

        MyBase.New(String_in_file)

    End Sub

    Public Sub New(ByRef Fields_array() As String)

        Me.Fields_array = Fields_array

        Load()

    End Sub

End Class

''' <summary>
''' Spherical coordinate system.
''' Vector's component
''' (0) - phi
''' (1) - theta
''' (2) - R
''' </summary>
''' <remarks></remarks>
<Serializable()> _
Public Class CORD1S

    Inherits Coordinate_system

    Public Overrides Function Transform_from_special_coordinates_to_local_rectangular(ByVal in_Vector As Vector) As Vector

        Dim tmp_Vector As New Vector

        tmp_Vector.Coord(0) = in_Vector.Coord(2) * Math.Sin(in_Vector.Coord(1)) * Math.Cos(in_Vector.Coord(0))

        tmp_Vector.Coord(1) = in_Vector.Coord(2) * Math.Sin(in_Vector.Coord(1)) * Math.Sin(in_Vector.Coord(0))

        tmp_Vector.Coord(2) = in_Vector.Coord(2) * Math.Cos(in_Vector.Coord(1))

        Transform_from_special_coordinates_to_local_rectangular = tmp_Vector

    End Function

    Public Overrides Function Transform_to_special_coordinates_from_local_rectangular(ByVal in_Vector As Vector) As Vector

        Dim tmp_Vector As New Vector

        If in_Vector.Coord(0) = in_Vector.Coord(1) And in_Vector.Coord(1) = in_Vector.Coord(2) Then

            tmp_Vector.Coord(0) = 0

            tmp_Vector.Coord(1) = 0

            tmp_Vector.Coord(2) = 0

            Transform_to_special_coordinates_from_local_rectangular = tmp_Vector

            Exit Function

        End If

        tmp_Vector.Coord(2) = Math.Sqrt(in_Vector.Coord(0) ^ 2 + in_Vector.Coord(1) ^ 2 + in_Vector.Coord(2) ^ 2)

        tmp_Vector.Coord(1) = Math.Acos(in_Vector.Coord(2) / tmp_Vector.Coord(2))

        tmp_Vector.Coord(0) = Math.Atan2(in_Vector.Coord(1), in_Vector.Coord(0))

        Transform_to_special_coordinates_from_local_rectangular = tmp_Vector

    End Function

    Public Sub New(Optional ByRef String_in_file As String = "")

        MyBase.New(String_in_file)

    End Sub

    Public Sub New(ByRef Fields_array() As String)

        Me.Fields_array = Fields_array

        Load()

    End Sub
End Class

''' <summary>
''' Spherical coordinate system.
''' Vector's component
''' (0) - phi
''' (1) - theta
''' (2) - R
''' </summary>
''' <remarks></remarks>
<Serializable()> _
Public Class CORD2S

    Inherits Coordinate_system

    Public Overrides Function Transform_from_special_coordinates_to_local_rectangular(ByVal in_Vector As Vector) As Vector

        Dim tmp_Vector As New Vector

        tmp_Vector.Coord(0) = in_Vector.Coord(2) * Math.Sin(in_Vector.Coord(1)) * Math.Cos(in_Vector.Coord(0))

        tmp_Vector.Coord(1) = in_Vector.Coord(2) * Math.Sin(in_Vector.Coord(1)) * Math.Sin(in_Vector.Coord(0))

        tmp_Vector.Coord(2) = in_Vector.Coord(2) * Math.Cos(in_Vector.Coord(1))

        Transform_from_special_coordinates_to_local_rectangular = tmp_Vector

    End Function

    Public Overrides Function Transform_to_special_coordinates_from_local_rectangular(ByVal in_Vector As Vector) As Vector

        Dim tmp_Vector As New Vector

        If in_Vector.Coord(0) = in_Vector.Coord(1) And in_Vector.Coord(1) = in_Vector.Coord(2) Then

            tmp_Vector.Coord(0) = 0

            tmp_Vector.Coord(1) = 0

            tmp_Vector.Coord(2) = 0

            Transform_to_special_coordinates_from_local_rectangular = tmp_Vector

            Exit Function

        End If

        tmp_Vector.Coord(2) = Math.Sqrt(in_Vector.Coord(0) ^ 2 + in_Vector.Coord(1) ^ 2 + in_Vector.Coord(2) ^ 2)

        tmp_Vector.Coord(1) = Math.Asin(in_Vector.Coord(2) / tmp_Vector.Coord(2))

        tmp_Vector.Coord(0) = Math.Atan2(in_Vector.Coord(1), in_Vector.Coord(0))

        Transform_to_special_coordinates_from_local_rectangular = tmp_Vector

    End Function

    Public Sub New(Optional ByRef String_in_file As String = "")

        MyBase.New(String_in_file)

    End Sub

    Public Sub New(ByRef Fields_array() As String)

        Me.Fields_array = Fields_array

        Load()

    End Sub
End Class

<Serializable()> _
Public Class Grid

    Inherits Model_item
    ''' <summary>
    ''' The coordinate system number, there determined grid's coordinates.
    ''' </summary>
    ''' <remarks></remarks>
    Public Coordinate_system As Object

    ''' <summary>
    ''' Grid's coordinates in related coordinate system.
    ''' </summary>
    ''' <remarks></remarks>
    Public Position As Vector

    ''' <summary>
    ''' The coordinate system number, there determined grid's displacement, degrees of freedom, boundary conditions and etc. 
    ''' </summary>
    ''' <remarks></remarks>
    Public CD As Object

    ''' <summary>
    ''' Grid's constant boundary conditions (SPC) associated with.
    ''' </summary>
    ''' <remarks></remarks>
    Public PS As Object

    ''' <summary>
    ''' Super element ID
    ''' </summary>
    ''' <remarks></remarks>
    Public Super_element_ID As Long

    ''' <summary>
    ''' Faces, which are builded on a face.
    ''' </summary>
    ''' <remarks></remarks>
    Public Face() As Object

    ''' <summary>
    ''' Maximum index of faces, contained in Face() array
    ''' </summary>
    ''' <remarks></remarks>
    Public N_Face As Long

    ''' <summary>
    ''' Elements, which are builded on a face.
    ''' </summary>
    ''' <remarks></remarks>
    Public Element() As Object

    ''' <summary>
    ''' Maximum index of element, contained in Element() array
    ''' </summary>
    ''' <remarks></remarks>
    Public N_Element As Long

    Public Overrides Sub Load()

        Position = New Vector

        Number = Val(Fields_array(1))

        Coordinate_system = Val(Fields_array(2))

        Position.Coord(0) = Val(Fields_array(3))

        Position.Coord(1) = Val(Fields_array(4))

        Position.Coord(2) = Val(Fields_array(5))

        CD = Val(Fields_array(6))

        PS = Val(Fields_array(7))

        Super_element_ID = Val(Fields_array(8))

    End Sub

    Public Overrides Sub Prepare_Save()

        ReDim Fields_array(8)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(Number)

        Fields_array(2) = my_Str(Coordinate_system)

        Fields_array(3) = my_Str(Position.Coord(0))

        Fields_array(4) = my_Str(Position.Coord(1))

        Fields_array(5) = my_Str(Position.Coord(2))

        Fields_array(6) = my_Str(CD)

        Fields_array(7) = my_Str(PS)

        Fields_array(8) = my_Str(Super_element_ID)

        On Error GoTo 0

    End Sub

    Public Sub New(Optional ByRef String_in_file As String = "")

        'Me.String_in_file = String_in_file

        'If String_in_file <> "" Then

        '    Fields_array = Convert_string_in_file_to_array_2(String_in_file)

        'End If

        'Load()

        MyBase.New(String_in_file)


    End Sub

    Public Sub New(ByRef Fields_array() As String)

        Me.Fields_array = Fields_array

        Load()

    End Sub
End Class

<Serializable()> _
Public MustInherit Class Element_1D_property

    Inherits Model_item

    Public Material As Object

    Public Area() As Double

    Public I1(), I2(), I12() As Double

    Public J() As Double

    Public NSM() As Double

    Public K1, K2 As Double

    Public C1(), C2(), D1(), D2(), E1(), E2(), F1(), F2() As Double

    Public Sub New(Optional ByRef String_in_file As String = "")

        MyBase.New(String_in_file)

    End Sub

    Public Sub New(ByRef Fields_array() As String)

        Me.Fields_array = Fields_array

        Load()

    End Sub
End Class

<Serializable()> _
Public Class PROD

    Inherits Element_1D_property


    Public Overrides Sub Load()

        ReDim Area(0), J(0), C1(0), NSM(0)

        On Error Resume Next

        Number = Val(Fields_array(1))

        Material = Val(Fields_array(2))

        Area(0) = Val(Fields_array(3))

        J(0) = Val(Fields_array(4))

        C1(0) = Val(Fields_array(5))

        NSM(0) = Val(Fields_array(6))

        On Error GoTo 0

    End Sub

    Public Overrides Sub Prepare_Save()

        ReDim Fields_array(6)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(Number)

        Fields_array(2) = my_Str(Material)

        Fields_array(3) = my_Str(Area(0))

        Fields_array(4) = my_Str(J(0))

        Fields_array(5) = my_Str(C1(0))

        Fields_array(6) = my_Str(NSM(0))

        On Error GoTo 0

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
Public Class PBAR

    Inherits Element_1D_property


    Public Overrides Sub Load()

        ReDim Area(0), I1(0), I2(0), NSM(0), C1(0), C2(0), D1(0), D2(0), E1(0), E2(0), F1(0), F2(0), I12(0)

        On Error Resume Next

        Number = Val(Fields_array(1))

        Material = Val(Fields_array(2))

        Area(0) = Val(Fields_array(3))

        I1(0) = Val(Fields_array(4))

        I2(0) = Val(Fields_array(5))

        J(0) = Val(Fields_array(6))

        NSM(0) = Val(Fields_array(7))



        C1(0) = Val(Fields_array(11))

        C2(0) = Val(Fields_array(12))

        D1(0) = Val(Fields_array(13))

        D2(0) = Val(Fields_array(14))

        E1(0) = Val(Fields_array(15))

        E2(0) = Val(Fields_array(16))

        F1(0) = Val(Fields_array(17))

        F2(0) = Val(Fields_array(18))


        K1 = Val(Fields_array(21))

        K2 = Val(Fields_array(22))

        I12(0) = Val(Fields_array(23))

        On Error GoTo 0

    End Sub

    Public Overrides Sub Prepare_Save()

        ReDim Fields_array(23)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(Number)

        Fields_array(2) = my_Str(Material)

        Fields_array(3) = my_Str(Area(0))

        Fields_array(4) = my_Str(I1(0))

        Fields_array(5) = my_Str(I2(0))

        Fields_array(6) = my_Str(J(0))

        Fields_array(7) = my_Str(NSM(0))



        Fields_array(11) = my_Str(C1(0))

        Fields_array(12) = my_Str(C2(0))

        Fields_array(13) = my_Str(D1(0))

        Fields_array(14) = my_Str(D2(0))

        Fields_array(15) = my_Str(E1(0))

        Fields_array(16) = my_Str(E2(0))

        Fields_array(17) = my_Str(F1(0))

        Fields_array(18) = my_Str(F2(0))


        Fields_array(21) = my_Str(K1)

        Fields_array(22) = my_Str(K2)

        Fields_array(23) = my_Str(I12(0))

        On Error GoTo 0

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
Public Class PBEAM

    Inherits Element_1D_property

    Public SO(0) As String

    Public X_XB(0) As Double

    Public S1, S2 As Double

    Public NSI(1), CW(1), M1(1), M2(1), N1(1), N2(1) As Double

    Public Overrides Sub Load()

        ReDim Area(0), I1(0), I2(0), NSM(0), C1(0), C2(0), D1(0), D2(0), E1(0), E2(0), F1(0), F2(0), I12(0)

        ReDim NSI(1), CW(1), N1(1), N2(1), M1(1), M2(1)

        On Error Resume Next

        Number = Val(Fields_array(1))

        Material = Val(Fields_array(2))

        Area(0) = Val(Fields_array(3))

        I1(0) = Val(Fields_array(4))

        I2(0) = Val(Fields_array(5))

        I12(0) = Val(Fields_array(6))

        J(0) = Val(Fields_array(7))

        NSM(0) = Val(Fields_array(8))



        C1(0) = Val(Fields_array(11))

        C2(0) = Val(Fields_array(12))

        D1(0) = Val(Fields_array(13))

        D2(0) = Val(Fields_array(14))

        E1(0) = Val(Fields_array(15))

        E2(0) = Val(Fields_array(16))

        F1(0) = Val(Fields_array(17))

        F2(0) = Val(Fields_array(18))

        Dim i_Section As Integer = 1

        ReDim SO(i_Section), X_XB(i_Section)

        For i As Integer = 20 To UBound(Fields_array) - 20

            SO(i_Section) = Fields_array(1 + i)

            X_XB(i_Section) = Val(Fields_array(2 + i))

            Area(i_Section) = Val(Fields_array(3 + i))

            I1(i_Section) = Val(Fields_array(4 + i))

            I2(i_Section) = Val(Fields_array(5 + i))

            I12(i_Section) = Val(Fields_array(6 + i))

            J(i_Section) = Val(Fields_array(7 + i))

            NSM(i_Section) = Val(Fields_array(8 + i))



            C1(i_Section) = Val(Fields_array(11 + i))

            C2(i_Section) = Val(Fields_array(12 + i))

            D1(i_Section) = Val(Fields_array(13 + i))

            D2(i_Section) = Val(Fields_array(14 + i))

            E1(i_Section) = Val(Fields_array(15 + i))

            E2(i_Section) = Val(Fields_array(16 + i))

            F1(i_Section) = Val(Fields_array(17 + i))

            F2(i_Section) = Val(Fields_array(18 + i))

            i_Section += 1

            ReDim Preserve SO(i_Section), X_XB(i_Section), Area(i_Section), I1(i_Section), I2(i_Section)

            ReDim Preserve I12(i_Section), J(i_Section)

            ReDim Preserve NSM(i_Section), C1(i_Section), C2(i_Section), D1(i_Section), D2(i_Section)

            ReDim Preserve E1(i_Section), E2(i_Section), F1(i_Section), F2(i_Section)

        Next i

        i_Section -= 1

        ReDim Preserve SO(i_Section), X_XB(i_Section), Area(i_Section), I1(i_Section), I2(i_Section)

        ReDim Preserve I12(i_Section), J(i_Section)

        ReDim Preserve NSM(i_Section), C1(i_Section), C2(i_Section), D1(i_Section), D2(i_Section)

        ReDim Preserve E1(i_Section), E2(i_Section), F1(i_Section), F2(i_Section)


        K1 = Val(Fields_array(UBound(Fields_array) - 18))

        K2 = Val(Fields_array(UBound(Fields_array) - 17))

        S1 = Val(Fields_array(UBound(Fields_array) - 17))

        S2 = Val(Fields_array(UBound(Fields_array) - 16))

        NSI(0) = Val(Fields_array(UBound(Fields_array) - 15))

        NSI(1) = Val(Fields_array(UBound(Fields_array) - 14))

        CW(0) = Val(Fields_array(UBound(Fields_array) - 13))

        CW(1) = Val(Fields_array(UBound(Fields_array) - 12))


        M1(0) = Val(Fields_array(UBound(Fields_array) - 8))

        M2(0) = Val(Fields_array(UBound(Fields_array) - 7))

        M1(1) = Val(Fields_array(UBound(Fields_array) - 6))

        M2(1) = Val(Fields_array(UBound(Fields_array) - 5))

        N1(0) = Val(Fields_array(UBound(Fields_array) - 4))

        N1(1) = Val(Fields_array(UBound(Fields_array) - 3))

        N2(0) = Val(Fields_array(UBound(Fields_array) - 2))

        N2(1) = Val(Fields_array(UBound(Fields_array) - 1))

        On Error GoTo 0

    End Sub

    Public Overrides Sub Prepare_Save()

        ReDim Fields_array(20 + UBound(Area) + 20)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(Number)

        Fields_array(2) = my_Str(Material)

        Fields_array(3) = my_Str(Area(0))

        Fields_array(4) = my_Str(I1(0))

        Fields_array(5) = my_Str(I2(0))

        Fields_array(6) = my_Str(I12(0))

        Fields_array(7) = my_Str(J(0))

        Fields_array(8) = my_Str(NSM(0))


        Fields_array(11) = my_Str(C1(0))

        Fields_array(12) = my_Str(C2(0))

        Fields_array(13) = my_Str(D1(0))

        Fields_array(14) = my_Str(D2(0))

        Fields_array(15) = my_Str(E1(0))

        Fields_array(16) = my_Str(E2(0))

        Fields_array(17) = my_Str(F1(0))

        Fields_array(18) = my_Str(F2(0))


        For i As Integer = 1 To UBound(Area)

            Fields_array(20 * i + 1) = SO(i)

            Fields_array(20 * i + 2) = my_Str(X_XB(i))

            Fields_array(20 * i + 3) = my_Str(Area(i))

            Fields_array(20 * i + 4) = my_Str(I1(i))

            Fields_array(20 * i + 5) = my_Str(I2(i))

            Fields_array(20 * i + 6) = my_Str(I12(i))

            Fields_array(20 * i + 7) = my_Str(J(i))

            Fields_array(20 * i + 8) = my_Str(NSM(i))



            Fields_array(20 * i + 11) = my_Str(C1(i))

            Fields_array(20 * i + 12) = my_Str(C2(i))

            Fields_array(20 * i + 13) = my_Str(D1(i))

            Fields_array(20 * i + 14) = my_Str(D2(i))

            Fields_array(20 * i + 15) = my_Str(E1(i))

            Fields_array(20 * i + 16) = my_Str(E2(i))

            Fields_array(20 * i + 17) = my_Str(F1(i))

            Fields_array(20 * i + 18) = my_Str(F2(i))


            ReDim Preserve Fields_array(20 * (i + 1))

        Next i


        Fields_array(UBound(Fields_array) - 18) = my_Str(K1)

        Fields_array(UBound(Fields_array) - 17) = my_Str(K2)

        Fields_array(UBound(Fields_array) - 16) = my_Str(S1)

        Fields_array(UBound(Fields_array) - 15) = my_Str(S2)

        Fields_array(UBound(Fields_array) - 14) = my_Str(NSI(0))

        Fields_array(UBound(Fields_array) - 13) = my_Str(NSI(1))

        Fields_array(UBound(Fields_array) - 12) = my_Str(CW(0))

        Fields_array(UBound(Fields_array) - 11) = my_Str(CW(1))


        Fields_array(UBound(Fields_array) - 8) = my_Str(M1(0))

        Fields_array(UBound(Fields_array) - 7) = my_Str(M2(0))

        Fields_array(UBound(Fields_array) - 6) = my_Str(M1(1))

        Fields_array(UBound(Fields_array) - 5) = my_Str(M2(1))

        Fields_array(UBound(Fields_array) - 4) = my_Str(N1(0))

        Fields_array(UBound(Fields_array) - 3) = my_Str(N1(1))

        Fields_array(UBound(Fields_array) - 2) = my_Str(N2(0))

        Fields_array(UBound(Fields_array) - 1) = my_Str(N2(1))

        On Error GoTo 0

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
Public Class PSHELL

    Inherits Model_item

    Public Material As Object

    ''' <summary>
    ''' T in NASTRAN formulation
    ''' </summary>
    ''' <remarks></remarks>
    Public Thickness As Double

    ''' <summary>
    ''' MID2 in NASTRAN formulation
    ''' </summary>
    ''' <remarks></remarks>
    Public Material_for_bending As Object

    ''' <summary>
    ''' 12I/T^3 in NASTRAN formulation
    ''' </summary>
    ''' <remarks></remarks>
    Public Bending_thichness_parameter As Double

    ''' <summary>
    ''' MID3 in NASTRAN formulation
    ''' </summary>
    ''' <remarks></remarks>
    Public Material_for_shear As Object

    ''' <summary>
    ''' TS/T in NASTRAN formulation
    ''' </summary>
    ''' <remarks></remarks>
    Public Shear_thickness_to_membrane_thickness_ratio As Double

    Public NSM As Double

    Public Z1, Z2 As Double
    ''' <summary>
    ''' MID4 in NASTRAN formulation
    ''' </summary>
    ''' <remarks></remarks>
    Public Material_for_membrane_bending_coupling As Object
    Public Overrides Sub Load()

        On Error Resume Next

        Number = Val(Fields_array(1))

        Material = Val(Fields_array(2))

        Thickness = Val(Fields_array(3))

        Material_for_bending = Val(Fields_array(4))

        Bending_thichness_parameter = Val(Fields_array(5))

        Material_for_shear = Val(Fields_array(6))

        Shear_thickness_to_membrane_thickness_ratio = Val(Fields_array(7))

        NSM = Val(Fields_array(8))



        Z1 = Val(Fields_array(11))

        Z2 = Val(Fields_array(12))

        Material_for_membrane_bending_coupling = Val(Fields_array(13))

        On Error GoTo 0

    End Sub

    Public Overrides Sub Prepare_Save()

        ReDim Fields_array(13)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(Number)

        Fields_array(2) = my_Str(Material)

        Fields_array(3) = my_Str(Thickness)

        Fields_array(4) = my_Str(Material_for_bending)

        Fields_array(5) = my_Str(Bending_thichness_parameter)

        Fields_array(6) = my_Str(Material_for_shear)

        Fields_array(7) = my_Str(Shear_thickness_to_membrane_thickness_ratio)

        Fields_array(8) = my_Str(NSM)



        Fields_array(11) = my_Str(Z1)

        Fields_array(12) = my_Str(Z2)

        Fields_array(13) = my_Str(Material_for_membrane_bending_coupling)

        On Error GoTo 0

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
Public Class PCOMP

    Inherits Model_item

    Public Class Layer
        ' номер элемента в массиве материалов
        Public Material As Object

        ' толщина слоя
        Public Thickness As Double

        ' угол ориентации
        Public Orient_angle As Double

        ' флаг запроса на вывод результатов
        Public SOUT As String

        ' различные специальные поля

        ' матрица жесткости монослоя
        Public с(,) As Double

    End Class

    ' дистанция от плоскости ссылки до нижней поверхности (???, взято из книжки)
    Public Z0 As Long

    ''' <summary>
    ''' Nonstructural mass per unit area
    ''' </summary>
    ''' <remarks></remarks>
    Public NSM As Double

    ' допускаемое напряжение в связующем
    Public SB As Double

    ' теория прочности
    Public FT As String

    ' ссылочная температура
    Public T_reference As Double

    Public Damping_coefficient As Double

    ' флаг симметричности конструкции
    Public LAM As String

    ' массив слоев
    Public Layers() As Layer


    ' рассчитываемые характеристики
    Public g11 As Double

    Public g22 As Double

    Public g12 As Double

    Public E_x As Double

    Public E_y As Double

    Public Thickness As Double

    ' доли материалов в различных направлениях

    Public Part_0 As Double

    Public Part_45_minus As Double

    Public Part_45_plus As Double

    Public Part_45_effective As Double ' эффективная доля слоев в направлении действия нагрузки. Используются для расчета допускаемых значений

    Public Part_90 As Double

    ' сумма долей материалов в различных направлениях. Должна быть как можно ближе к 1.

    Public Summary_parts As Double

    Public Effective_thickness As Double

    Public Allowable_compression_epsilon As Double

    Public Allowable_tension_epsilon As Double

    Public Angle_true As Boolean ' признак того, что укладка выдержана в углах 0,-45, +45, 90

    Public Max_E1 As Double 'максимальный модуль из всех материалов в укладке

    'Растяжения, мембранная жесткость
    Public A(2, 2) As Double

    ' изгибная матрица жесткости
    Public B(2, 2) As Double

    ' совместная матрица жесткости
    Public D(2, 2) As Double

    Public Overrides Sub Load()

        ReDim Layers(0)

        Layers(0) = New Layer

        ReDim Layers(0).с(2, 2)

        On Error Resume Next

        Number = Val(Fields_array(1))

        Z0 = Val(Fields_array(2))

        NSM = Val(Fields_array(3))

        SB = Val(Fields_array(4))

        FT = Fields_array(5)

        T_reference = Val(Fields_array(6))

        Damping_coefficient = Val(Fields_array(7))

        LAM = Fields_array(8)


        Dim i_Layer As Integer = -1

        Dim Layer_Material As Long

        For i As Integer = 10 To UBound(Fields_array) Step 10

            Layer_Material = Val(Fields_array(i + 1))

            If Layer_Material > 0 Then

                i_Layer += 1

                ReDim Preserve Layers(i_Layer)

                Layers(i_Layer) = New Layer

                ReDim Layers(i_Layer).с(2, 2)

                Layers(i_Layer).Material = Layer_Material

                Layers(i_Layer).Thickness = Val(Fields_array(i + 2))

                Layers(i_Layer).Orient_angle = Val(Fields_array(i + 3))

                Layers(i_Layer).SOUT = Fields_array(i + 4)

            End If

            Layer_Material = Val(Fields_array(i + 5))

            If Layer_Material > 0 Then

                i_Layer += 1

                ReDim Preserve Layers(i_Layer)

                Layers(i_Layer) = New Layer

                ReDim Layers(i_Layer).с(2, 2)

                Layers(i_Layer).Material = Layer_Material

                Layers(i_Layer).Thickness = Val(Fields_array(i + 6))

                Layers(i_Layer).Orient_angle = Val(Fields_array(i + 7))

                Layers(i_Layer).SOUT = Fields_array(i + 8)

            End If

        Next i

        On Error GoTo 0

    End Sub

    Public Overrides Sub Prepare_Save()

        Dim N_lines_for_layers As Integer = 0

        Dim i_Layer As Integer = 0

        If UBound(Layers) = UBound(Layers) \ 2 Then

            N_lines_for_layers = UBound(Layers) \ 2

        Else

            N_lines_for_layers = UBound(Layers) \ 2 + 1

        End If

        ReDim Fields_array(10 + N_lines_for_layers * 10)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(Number)

        Fields_array(2) = my_Str(Z0)

        Fields_array(3) = my_Str(NSM)

        Fields_array(4) = my_Str(SB)

        Fields_array(5) = FT

        Fields_array(6) = my_Str(T_reference)

        Fields_array(7) = my_Str(Damping_coefficient)

        Fields_array(8) = my_Str(LAM)



        For i As Integer = 1 To N_lines_for_layers

            If i_Layer <= UBound(Layers) Then

                Fields_array(10 * i + 1) = my_Str(Layers(i_Layer).Material)

                Fields_array(10 * i + 2) = my_Str(Layers(i_Layer).Thickness)

                Fields_array(10 * i + 3) = my_Str(Layers(i_Layer).Orient_angle)

                Fields_array(10 * i + 4) = Layers(i_Layer).SOUT

                i_Layer += 1

            Else

                Exit For

            End If

            If i_Layer <= UBound(Layers) Then

                Fields_array(10 * i + 5) = my_Str(Layers(i_Layer).Material)

                Fields_array(10 * i + 6) = my_Str(Layers(i_Layer).Thickness)

                Fields_array(10 * i + 7) = my_Str(Layers(i_Layer).Orient_angle)

                Fields_array(10 * i + 8) = Layers(i_Layer).SOUT

                i_Layer += 1

            Else

                Exit For

            End If


        Next i

        On Error GoTo 0

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
Public Class PSOLID

    Inherits Model_item

    Public Material As Object

    Public Coordinate_system As Object

    Public Integration_mesh As String

    Public Stress As String

    Public ISOP As String

    Public FCNT As String


    Public Overrides Sub Load()

        On Error Resume Next

        Number = Val(Fields_array(1))

        If Fields_array(2) = "" Then

            Material = -1

        Else

            Material = Val(Fields_array(2))

        End If

        Coordinate_system = Val(Fields_array(3))

        Integration_mesh = Fields_array(4)

        Stress = Fields_array(5)

        ISOP = Fields_array(6)

        FCNT = Fields_array(7)

        On Error GoTo 0

    End Sub

    Public Overrides Sub Prepare_Save()

        ReDim Fields_array(7)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(Number)

        Fields_array(2) = my_Str(Material)

        Fields_array(3) = my_Str(Coordinate_system)

        Fields_array(4) = Integration_mesh

        Fields_array(5) = Stress

        Fields_array(6) = ISOP

        Fields_array(7) = FCNT

        On Error GoTo 0

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
Public Class PSHEAR

    Inherits Model_item

    Public Material As Object

    Public Thickness As Double

    Public NSM As Double

    Public F1, F2 As Double

    Public Overrides Sub Load()

        On Error Resume Next

        Number = Val(Fields_array(1))

        Material = Val(Fields_array(2))

        Thickness = Val(Fields_array(3))

        NSM = Val(Fields_array(4))

        F1 = Val(Fields_array(5))

        F2 = Val(Fields_array(6))

        On Error GoTo 0

    End Sub
    Public Overrides Sub Prepare_Save()

        ReDim Fields_array(6)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(Number)

        Fields_array(2) = my_Str(Material)

        Fields_array(3) = my_Str(Thickness)

        Fields_array(4) = my_Str(NSM)

        Fields_array(5) = my_Str(F1)

        Fields_array(6) = my_Str(F2)

        On Error GoTo 0

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
Public Class Neighbour

    ''' <summary>
    ''' Neightbouring element
    ''' </summary>
    ''' <remarks></remarks>
    Public Element As Object

    ''' <summary>
    ''' Coupling face
    ''' </summary>
    ''' <remarks></remarks>
    Public Face As Object

    ''' <summary>
    ''' Heat transfer area of first element in neighbourhood
    ''' </summary>
    ''' <remarks></remarks>
    Public Area_1 As Double

    ''' <summary>
    ''' Heat transfer area of first element in neighbourhood
    ''' </summary>
    ''' <remarks></remarks>
    Public Area_2 As Double

End Class

''' <summary>
''' This class is to describe variables, calculated for many purposes.
''' I.e. stress, deformation, displasement tensors for stress analysis,
''' temperature for thermal analysis,
''' velocity and acceleration for dymanic analysis and etc. 
''' </summary>
''' <remarks></remarks>
<Serializable()> _
Public Class Finite_element_variables

    ' General properties

    Public Volume As Double

    Public Density As Double

    Public Mass As Double

    Public Centroid As New Vector

    Public Inertia(,) As Double

    ' Stress properties 

    Public Displacement(,) As Double

    Public Strain(,) As Double

    Public Force(,) As Double

    Public Moment(,) As Double

    Public Stress(,) As Double

    ' Thermal properties

    Public Temperature As Double

    Public Specific_heat_capacity As Double

    Public Heat_conductivity() As Double

    ' Dynamic properties

    Public Velocity(,) As Double

    Public Acceleration(,) As Double


    Public Sub New()

        'Volume = 0

        'Density = 0

        'Mass = 0



        'Temperature = 0

        'Specific_heat_capacity = 0

        'Heat_conductivity = New Double() {0, 0, 0}



        'Displacement = New Double(,) {{0, 0, 0}, {0, 0, 0}, {0, 0, 0}}

        'Strain = New Double(,) {{0, 0, 0}, {0, 0, 0}, {0, 0, 0}}

        'Force = New Double(,) {{0, 0, 0}, {0, 0, 0}, {0, 0, 0}}

        'Moment = New Double(,) {{0, 0, 0}, {0, 0, 0}, {0, 0, 0}}

        'Stress = New Double(,) {{0, 0, 0}, {0, 0, 0}, {0, 0, 0}}



        'Velocity = New Double(,) {{0, 0, 0}, {0, 0, 0}, {0, 0, 0}}

        'Acceleration = New Double(,) {{0, 0, 0}, {0, 0, 0}, {0, 0, 0}}



    End Sub
End Class



<Serializable()> _
Public MustInherit Class Finite_element

    Inherits Model_item

    ''' <summary>
    ''' Finite element property number, means material and something else.
    ''' PID in NASTRAN formulation
    ''' </summary>
    ''' <remarks></remarks>
    Public Element_property As Object

    ''' <summary>
    ''' Grid, which finite element consist of. 
    ''' </summary>
    ''' <remarks></remarks>
    Public Grid() As Object


    ''' <summary>
    ''' The state of element, stress, temperature, velocity and etc.
    ''' </summary>
    ''' <remarks></remarks>
    Public Field As New Finite_element_variables

    ''' <summary>
    ''' Faces, which forms element surface.
    ''' </summary>
    ''' <remarks></remarks>
    Public Face() As Object

    ''' <summary>
    ''' Neightbors of element
    ''' </summary>
    ''' <remarks></remarks>
    Public Neighbor() As Object

    ''' <summary>
    ''' Numbers of neighbors, contained in Neighbor() array
    ''' </summary>
    ''' <remarks></remarks>
    Public N_Neighbor As Long

    ''' <summary>
    ''' Element's thermal-optical properties
    ''' </summary>
    ''' <remarks></remarks>
    Public MEDIUM As Object

    ''' <summary>
    ''' Heat transfer element, representing finite element in heat transfer analysis
    ''' </summary>
    ''' <remarks></remarks>
    Public HT_Element As Object

    Public MustOverride Sub Calc_volume_mass_inertia()


    Public Sub New(Optional ByRef String_in_file As String = "")

        MyBase.New(String_in_file)

    End Sub

    Public Sub New(ByRef Fields_array() As String)

        Me.Fields_array = Fields_array

        Load()

    End Sub
End Class

<Serializable()> _
Public Class CROD

    Inherits Finite_element

    ''' <summary>
    ''' Length of element
    ''' </summary>
    ''' <remarks></remarks>
    Public length As Double

    Public Overrides Sub Load()

        ReDim Grid(1)

        On Error Resume Next

        Number = Val(Fields_array(1))

        Element_property = Val(Fields_array(2))

        Grid(0) = Val(Fields_array(3))

        Grid(1) = Val(Fields_array(4))

        On Error GoTo 0

    End Sub
    Public Overrides Sub Prepare_Save()

        ReDim Fields_array(6)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(Number)

        Fields_array(2) = my_Str(Element_property)

        Fields_array(3) = my_Str(Grid(0))

        Fields_array(4) = my_Str(Grid(1))

        On Error GoTo 0

    End Sub

    Public Sub New(Optional ByRef String_in_file As String = "")

        MyBase.New(String_in_file)

    End Sub

    Public Sub New(ByRef Fields_array() As String)

        Me.Fields_array = Fields_array

        Load()

    End Sub

    ''' <summary>
    ''' This sub calculates element mass, volume, centroid and moment of inertia in element's grids coordinate system.
    ''' WARNING!!! 1D-element moments of inertia are calculated with error up to 15% 
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Calc_volume_mass_inertia()

        Dim MIP As New MIP

        Dim Vertex(UBound(Grid)) As Vector

        For i As Integer = 0 To UBound(Grid)

            Vertex(i) = Grid(i).Position

        Next i

        Field.Inertia = MIP.Inertia_2_node(Vertex, Element_property.Area, Field.Volume, Field.Centroid)

        Field.Density = Element_property.Material.Rho

        Field.Mass = Field.Volume * Field.Density

        ' Adding nonstructural mass

        Field.Mass += length * Element_property.NSM

    End Sub

End Class

<Serializable()> _
Public Class CBAR

    Inherits Finite_element

    ''' <summary>
    ''' Length of element
    ''' </summary>
    ''' <remarks></remarks>
    Public length As Double

    Public G0 As Long

    Public X(2), WA(2), WB(2) As Double

    Public Pin_flag(1) As Long

    Public Overrides Sub Load()

        ReDim Grid(1)

        On Error Resume Next

        Number = Val(Fields_array(1))

        Element_property = Val(Fields_array(2))

        Grid(0) = Val(Fields_array(3))

        Grid(1) = Val(Fields_array(4))

        If Fields_array(7) = "" And Fields_array(8) = "" Then

            G0 = Val(Fields_array(5))

        Else

            X(0) = Val(Fields_array(5))

            X(1) = Val(Fields_array(6))

            X(2) = Val(Fields_array(7))

        End If

        Pin_flag(0) = Val(Fields_array(11))

        Pin_flag(1) = Val(Fields_array(12))

        WA(0) = Val(Fields_array(13))

        WA(1) = Val(Fields_array(14))

        WA(2) = Val(Fields_array(15))

        WB(0) = Val(Fields_array(16))

        WB(1) = Val(Fields_array(17))

        WB(2) = Val(Fields_array(18))

        On Error GoTo 0

    End Sub
    Public Overrides Sub Prepare_Save()

        ReDim Fields_array(18)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(Number)

        Fields_array(2) = my_Str(Element_property)

        Fields_array(3) = my_Str(Grid(0))

        Fields_array(4) = my_Str(Grid(1))

        If G0 > 0 Then

            Fields_array(5) = my_Str(G0)

        Else

            Fields_array(5) = my_Str(X(0))

            Fields_array(6) = my_Str(X(1))

            Fields_array(7) = my_Str(X(2))

        End If

        Fields_array(11) = my_Str(Pin_flag(0))

        Fields_array(12) = my_Str(Pin_flag(1))

        Fields_array(13) = my_Str(WA(0))

        Fields_array(14) = my_Str(WA(1))

        Fields_array(15) = my_Str(WA(2))

        Fields_array(16) = my_Str(WB(0))

        Fields_array(17) = my_Str(WB(1))

        Fields_array(18) = my_Str(WB(2))

        On Error GoTo 0

    End Sub

    Public Sub New(Optional ByRef String_in_file As String = "")

        MyBase.New(String_in_file)

    End Sub

    Public Sub New(ByRef Fields_array() As String)

        Me.Fields_array = Fields_array

        Load()

    End Sub
    ''' <summary>
    ''' This sub calculates element mass, volume, centroid and moment of inertia in element's grids coordinate system.
    ''' WARNING!!! 1D-element moments of inertia are calculated with error up to 15% 
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Calc_volume_mass_inertia()

        Dim MIP As New MIP

        Dim Vertex(UBound(Grid)) As Vector

        For i As Integer = 0 To UBound(Grid)

            Vertex(i) = Grid(i).Position

        Next i

        Field.Inertia = MIP.Inertia_2_node(Vertex, Element_property.Area, Field.Volume, Field.Centroid)

        Field.Density = Element_property.Material.Rho

        Field.Mass = Field.Volume * Field.Density

        ' Adding nonstructural mass

        Field.Mass += length * Element_property.NSM

    End Sub

End Class

<Serializable()> _
Public Class CBEAM

    Inherits Finite_element

    ''' <summary>
    ''' Length of element
    ''' </summary>
    ''' <remarks></remarks>
    Public length As Double

    Public G0 As Long

    Public X(2), WA(2), WB(2) As Double

    Public Pin_flag(1), S(1) As Long

    Public Overrides Sub Load()

        ReDim Grid(1)

        On Error Resume Next

        Number = Val(Fields_array(1))

        Element_property = Val(Fields_array(2))

        Grid(0) = Val(Fields_array(3))

        Grid(1) = Val(Fields_array(4))

        X(0) = Val(Fields_array(5))

        X(1) = Val(Fields_array(6))

        X(2) = Val(Fields_array(7))


        Pin_flag(0) = Val(Fields_array(11))

        Pin_flag(1) = Val(Fields_array(12))

        WA(0) = Val(Fields_array(13))

        WA(1) = Val(Fields_array(14))

        WA(2) = Val(Fields_array(15))

        WB(0) = Val(Fields_array(16))

        WB(1) = Val(Fields_array(17))

        WB(2) = Val(Fields_array(18))


        S(0) = Val(Fields_array(21))

        S(1) = Val(Fields_array(22))

        On Error GoTo 0

    End Sub
    Public Overrides Sub Prepare_Save()

        ReDim Fields_array(18)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(Number)

        Fields_array(2) = my_Str(Element_property)

        Fields_array(3) = my_Str(Grid(0))

        Fields_array(4) = my_Str(Grid(1))

        Fields_array(5) = my_Str(X(0))

        Fields_array(6) = my_Str(X(1))

        Fields_array(7) = my_Str(X(2))



        Fields_array(11) = my_Str(Pin_flag(0))

        Fields_array(12) = my_Str(Pin_flag(1))

        Fields_array(13) = my_Str(WA(0))

        Fields_array(14) = my_Str(WA(1))

        Fields_array(15) = my_Str(WA(2))

        Fields_array(16) = my_Str(WB(0))

        Fields_array(17) = my_Str(WB(1))

        Fields_array(18) = my_Str(WB(2))


        Fields_array(21) = my_Str(S(0))

        Fields_array(22) = my_Str(S(1))

        On Error GoTo 0

    End Sub

    Public Sub New(Optional ByRef String_in_file As String = "")

        MyBase.New(String_in_file)

    End Sub

    Public Sub New(ByRef Fields_array() As String)

        Me.Fields_array = Fields_array

        Load()

    End Sub
    ''' <summary>
    ''' This sub calculates element mass, volume, centroid and moment of inertia in element's grids coordinate system.
    ''' WARNING!!! 1D-element moments of inertia are calculated with error up to 15% 
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Calc_volume_mass_inertia()

        Dim MIP As New MIP

        Dim Vertex(UBound(Grid)) As Vector

        Field.Inertia = Nothing

        Field.Density = Element_property.Material.Rho

        Dim Volume(UBound(Element_property.Area)) As Double

        Dim Centroid(UBound(Element_property.Area)) As Vector

        ReDim Field.Inertia(2, 2)

        Dim tmp_Inertia(,) As Double

        Field.Volume = 0

        Field.Mass = 0

        Dim oz As New Vector

        oz = Norm_vector(Vertex(1) - Vertex(0))

        Vertex(0) = Grid(0).Position



        Dim L As Double

        For i As Integer = 0 To UBound(Element_property.Area)

            Vertex(1) = Vertex(0) + Element_property.X_XB(i) * oz - Vertex(1)

            tmp_Inertia = MIP.Inertia_2_node(Vertex, Element_property.Area(i), Volume(i), Centroid(i))

            Field.Mass += Volume(i) * Field.Density

            Field.Volume += Volume(i)

            For m As Integer = 0 To 2

                For n As Integer = 0 To 2

                    Field.Inertia(m, n) += tmp_Inertia(m, n)

                Next n

            Next m

            L = L_vector(Vertex(1) - Vertex(0))

            Field.Mass += L * Element_property.NSM(i)

        Next i

        For i As Integer = 0 To UBound(Element_property.Area)

            Field.Centroid.Coord(0) += Centroid(i).Coord(0) * Volume(i) / Field.Volume

            Field.Centroid.Coord(1) += Centroid(i).Coord(1) * Volume(i) / Field.Volume

            Field.Centroid.Coord(2) += Centroid(i).Coord(2) * Volume(i) / Field.Volume

        Next i

    End Sub
End Class

<Serializable()> _
Public Class CTRIA3

    Inherits Finite_element

    ''' <summary>
    ''' Area of element
    ''' </summary>
    ''' <remarks></remarks>
    Public Area As Double

    ''' <summary>
    ''' THETA in Nastran formulation
    ''' </summary>
    ''' <remarks></remarks>
    Public Material_orientation_angle As Double

    ''' <summary>
    ''' MCID in Nastran formulation
    ''' </summary>
    ''' <remarks></remarks>
    Public Coordinate_system As Object

    ''' <summary>
    ''' ZOFFS in Nastran formulation
    ''' </summary>
    ''' <remarks></remarks>
    Public Offset As Double

    Public TFLAG As Long

    Public T() As Double

    Public Overrides Sub Load()

        ReDim Grid(2)

        ReDim T(2)

        Number = Val(Fields_array(1))

        Element_property = Val(Fields_array(2))

        Grid(0) = Val(Fields_array(3))

        Grid(1) = Val(Fields_array(4))

        Grid(2) = Val(Fields_array(5))

        If InStr(Fields_array(6), ".") > 0 Then

            Material_orientation_angle = Val(Fields_array(6))

        Else

            Coordinate_system = Val(Fields_array(6))

        End If

        Offset = Val(Fields_array(7))


        If UBound(Fields_array) > 11 Then

            TFLAG = Val(Fields_array(12))

            T(0) = Val(Fields_array(13))

            T(1) = Val(Fields_array(14))

            T(2) = Val(Fields_array(15))

        End If

    End Sub
    Public Overrides Sub Prepare_Save()

        ReDim Fields_array(15)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(Number)

        Fields_array(2) = my_Str(Element_property)

        Fields_array(3) = my_Str(Grid(0))

        Fields_array(4) = my_Str(Grid(1))

        Fields_array(5) = my_Str(Grid(2))

        If Material_orientation_angle = Double.MinValue Then

            Fields_array(6) = my_Str(Coordinate_system)

        Else

            Fields_array(6) = my_Str(Material_orientation_angle)

        End If

        Fields_array(7) = my_Str(Offset)



        Fields_array(12) = my_Str(TFLAG)

        Fields_array(13) = my_Str(T(0))

        Fields_array(14) = my_Str(T(1))

        Fields_array(15) = my_Str(T(2))

        On Error GoTo 0

    End Sub

    Public Sub New(Optional ByRef String_in_file As String = "")

        MyBase.New(String_in_file)

        Coordinate_system = Long.MinValue

        Material_orientation_angle = Double.MinValue

    End Sub

    Public Sub New(ByRef Fields_array() As String)

        Me.Fields_array = Fields_array

        Coordinate_system = Long.MinValue

        Material_orientation_angle = Double.MinValue

        Load()

    End Sub

    ''' <summary>
    ''' This sub calculates element mass, volume, centroid and moment of inertia in element's grids coordinate system.
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Calc_volume_mass_inertia()

        Dim MIP As New MIP

        Dim Vertex(2 * UBound(Grid) + 1) As Vector

        Dim oz As Vector = Face(0).oz


        If Element_property.Type = "PSHELL" Then

            For i As Integer = 0 To UBound(Grid)

                Vertex(i) = Grid(i).Position + Element_property.Thickness * 0.5 * oz

            Next i

            For i As Integer = UBound(Grid) + 1 To 2 * UBound(Grid) + 1

                Vertex(i) = Grid(i - (UBound(Grid) + 1)).Position - Element_property.Thickness * 0.5 * oz

            Next i


            Field.Inertia = MIP.Inertia_6_node(Vertex, Field.Volume, Field.Centroid)

            Field.Density = Element_property.Material.Rho

            Field.Mass = Field.Volume * Field.Density

            ' Adding nonstructural mass

            Field.Mass += Face(0).area * Element_property.NSM

        End If

        If Element_property.Type = "PCOMP" Then

            Dim Volume(UBound(Element_property.Layer)) As Double

            Dim Mass(UBound(Element_property.Layer)) As Double

            Dim Centroid(UBound(Element_property.Layer)) As Vector

            Dim tmp_Inertia(,) As Double


            For i As Integer = 0 To UBound(Element_property.Layer)

                For j As Integer = 0 To UBound(Grid)

                    Vertex(j) = Grid(j).Position + Element_property.Layer(i).Thickness * 0.5 * oz

                Next j

                For j As Integer = UBound(Grid) + 1 To 2 * UBound(Grid) + 1

                    Vertex(j) = Grid(j - (UBound(Grid) + 1)).Position - Element_property.Layer(i).Thickness * 0.5 * oz

                Next j

                tmp_Inertia = MIP.Inertia_6_node(Vertex, Volume(i), Centroid(i))

                Field.Volume += Volume(i)

                Mass(i) = Volume(i) * Element_property.Layer(i).Material.Rho

                Field.Mass += Mass(i)

                For m As Integer = 0 To 2

                    For n As Integer = 0 To 2

                        Field.Inertia(m, n) = tmp_Inertia(m, n)

                    Next n

                Next m

            Next i

            For i As Integer = 0 To UBound(Element_property.Layer)

                Field.Centroid.Coord(0) += Centroid(i).Coord(0) * Mass(i) / Field.Mass

                Field.Centroid.Coord(1) += Centroid(i).Coord(1) * Mass(i) / Field.Mass

                Field.Centroid.Coord(2) += Centroid(i).Coord(2) * Mass(i) / Field.Mass

            Next i

        End If

    End Sub

End Class

<Serializable()> _
Public Class CTRIA6

    Inherits Finite_element

    ''' <summary>
    ''' Area of element
    ''' </summary>
    ''' <remarks></remarks>
    Public Area As Double

    ''' <summary>
    ''' THETA in Nastran formulation
    ''' </summary>
    ''' <remarks></remarks>
    Public Material_orientation_angle As Double

    ''' <summary>
    ''' MCoordinate_system in Nastran formulation
    ''' </summary>
    ''' <remarks></remarks>
    Public Coordinate_system As Object

    ''' <summary>
    ''' ZOFFS in Nastran formulation
    ''' </summary>
    ''' <remarks></remarks>
    Public Offset As Double

    Public T() As Double

    Public TFLAG As Long

    Public Overrides Sub Load()

        ReDim Grid(5)

        ReDim T(2)

        On Error Resume Next

        Number = Val(Fields_array(1))

        Element_property = Val(Fields_array(2))

        Grid(0) = Val(Fields_array(3))

        Grid(1) = Val(Fields_array(4))

        Grid(2) = Val(Fields_array(5))

        Grid(3) = Val(Fields_array(6))

        Grid(4) = Val(Fields_array(7))

        Grid(5) = Val(Fields_array(8))



        If InStr(Fields_array(11), ".") > 0 Then

            Material_orientation_angle = Val(Fields_array(11))

        Else

            Coordinate_system = Val(Fields_array(11))

        End If

        Offset = Val(Fields_array(12))

        T(0) = Val(Fields_array(13))

        T(1) = Val(Fields_array(14))

        T(2) = Val(Fields_array(15))

        TFLAG = Val(Fields_array(16))

        On Error GoTo 0

    End Sub
    Public Overrides Sub Prepare_Save()

        ReDim Fields_array(16)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(Number)

        Fields_array(2) = my_Str(Element_property)

        Fields_array(3) = my_Str(Grid(0))

        Fields_array(4) = my_Str(Grid(1))

        Fields_array(5) = my_Str(Grid(2))

        Fields_array(6) = my_Str(Grid(3))

        Fields_array(7) = my_Str(Grid(4))

        Fields_array(8) = my_Str(Grid(5))



        If Material_orientation_angle = Double.MinValue Then

            Fields_array(11) = my_Str(Coordinate_system)

        Else

            Fields_array(11) = my_Str(Material_orientation_angle)

        End If

        Fields_array(12) = my_Str(Offset)

        Fields_array(13) = my_Str(T(0))

        Fields_array(14) = my_Str(T(1))

        Fields_array(15) = my_Str(T(2))

        Fields_array(16) = my_Str(TFLAG)

        On Error GoTo 0

    End Sub

    Public Sub New(Optional ByRef String_in_file As String = "")

        MyBase.New(String_in_file)



    End Sub

    Public Sub New(ByRef Fields_array() As String)

        Me.Fields_array = Fields_array

        Coordinate_system = Long.MinValue

        Material_orientation_angle = Double.MinValue

        Load()

    End Sub

    ''' <summary>
    ''' This sub calculates element mass, volume, centroid and moment of inertia in element's grids coordinate system.
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Calc_volume_mass_inertia()

        Dim MIP As New MIP

        Dim Vertex(2 * UBound(Grid) + 1) As Vector

        Dim oz As Vector = Face(0).oz


        If Element_property.Type = "PSHELL" Then

            For i As Integer = 0 To UBound(Grid)

                Vertex(i) = Grid(i).Position + Element_property.Thickness * 0.5 * oz

            Next i

            For i As Integer = UBound(Grid) + 1 To 2 * UBound(Grid) + 1

                Vertex(i) = Grid(i - (UBound(Grid) + 1)).Position - Element_property.Thickness * 0.5 * oz

            Next i


            Field.Inertia = MIP.Inertia_6_node(Vertex, Field.Volume, Field.Centroid)

            Field.Density = Element_property.Material.Rho

            Field.Mass = Field.Volume * Field.Density

            ' Adding nonstructural mass

            Field.Mass += Face(0).area * Element_property.NSM

        End If

        If Element_property.Type = "PCOMP" Then

            Dim Volume(UBound(Element_property.Layer)) As Double

            Dim Mass(UBound(Element_property.Layer)) As Double

            Dim Centroid(UBound(Element_property.Layer)) As Vector

            Dim tmp_Inertia(,) As Double


            For i As Integer = 0 To UBound(Element_property.Layer)

                For j As Integer = 0 To UBound(Grid)

                    Vertex(j) = Grid(j).Position + Element_property.Layer(i).Thickness * 0.5 * oz

                Next j

                For j As Integer = UBound(Grid) + 1 To 2 * UBound(Grid) + 1

                    Vertex(j) = Grid(j - (UBound(Grid) + 1)).Position - Element_property.Layer(i).Thickness * 0.5 * oz

                Next j

                tmp_Inertia = MIP.Inertia_6_node(Vertex, Volume(i), Centroid(i))

                Field.Volume += Volume(i)

                Mass(i) = Volume(i) * Element_property.Layer(i).Material.Rho

                Field.Mass += Mass(i)

                For m As Integer = 0 To 2

                    For n As Integer = 0 To 2

                        Field.Inertia(m, n) = tmp_Inertia(m, n)

                    Next n

                Next m

            Next i

            For i As Integer = 0 To UBound(Element_property.Layer)

                Field.Centroid.Coord(0) += Centroid(i).Coord(0) * Mass(i) / Field.Mass

                Field.Centroid.Coord(1) += Centroid(i).Coord(1) * Mass(i) / Field.Mass

                Field.Centroid.Coord(2) += Centroid(i).Coord(2) * Mass(i) / Field.Mass

            Next i

        End If

    End Sub

End Class

<Serializable()> _
Public Class CQUAD4

    Inherits Finite_element

    ''' <summary>
    ''' Area of element
    ''' </summary>
    ''' <remarks></remarks>
    Public Area As Double

    ''' <summary>
    ''' THETA in Nastran formulation
    ''' </summary>
    ''' <remarks></remarks>
    Public Material_orientation_angle As Double

    ''' <summary>
    ''' MCoordinate_system in Nastran formulation
    ''' </summary>
    ''' <remarks></remarks>
    Public Coordinate_system As Object

    ''' <summary>
    ''' ZOFFS in Nastran formulation
    ''' </summary>
    ''' <remarks></remarks>
    Public Offset As Double

    Public TFLAG As Long

    Public T() As Double

    Public Overrides Sub Load()

        ReDim Grid(3)

        ReDim T(3)

        Number = Val(Fields_array(1))

        Element_property = Val(Fields_array(2))

        Grid(0) = Val(Fields_array(3))

        Grid(1) = Val(Fields_array(4))

        Grid(2) = Val(Fields_array(5))

        Grid(3) = Val(Fields_array(6))

        If InStr(Fields_array(7), ".") > 0 Then

            Material_orientation_angle = Val(Fields_array(7))

        Else

            Coordinate_system = Val(Fields_array(7))

        End If

        Offset = Val(Fields_array(8))


        If UBound(Fields_array) > 11 Then

            TFLAG = Val(Fields_array(12))

            T(0) = Val(Fields_array(13))

            T(1) = Val(Fields_array(14))

            T(2) = Val(Fields_array(15))

            T(3) = Val(Fields_array(16))

        End If

    End Sub
    Public Overrides Sub Prepare_Save()

        ReDim Fields_array(16)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(Number)

        Fields_array(2) = my_Str(Element_property)

        Fields_array(3) = my_Str(Grid(0))

        Fields_array(4) = my_Str(Grid(1))

        Fields_array(5) = my_Str(Grid(2))

        Fields_array(6) = my_Str(Grid(3))

        If Coordinate_system = Long.MinValue And Material_orientation_angle = Double.MinValue Then

            Fields_array(7) = ""

        Else

            If Material_orientation_angle <> Double.MinValue Then

                Fields_array(7) = my_Str(Material_orientation_angle)

            Else

                Fields_array(7) = my_Str(Coordinate_system)

            End If


        End If

        Fields_array(8) = my_Str(Offset)



        Fields_array(12) = my_Str(TFLAG)

        Fields_array(13) = my_Str(T(0))

        Fields_array(14) = my_Str(T(1))

        Fields_array(15) = my_Str(T(2))

        Fields_array(16) = my_Str(T(3))

        On Error GoTo 0

    End Sub

    Public Sub New(Optional ByRef String_in_file As String = "")

        MyBase.New(String_in_file)



    End Sub

    Public Sub New(ByRef Fields_array() As String)

        Me.Fields_array = Fields_array

        Coordinate_system = Long.MinValue

        Material_orientation_angle = Double.MinValue

        Load()

    End Sub

    ''' <summary>
    ''' This sub calculates element mass, volume, centroid and moment of inertia in element's grids coordinate system.
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Calc_volume_mass_inertia()

        Dim MIP As New MIP

        ReDim Field.Inertia(2, 2)

        Dim Vertex(2 * UBound(Grid) + 1) As Vector

        Dim oz As Vector = Face(0).oz


        If Element_property.Type = "PSHELL" Then

            For i As Integer = 0 To UBound(Grid)

                Vertex(i) = Grid(i).Position + Element_property.Thickness * 0.5 * oz

            Next i

            For i As Integer = UBound(Grid) + 1 To 2 * UBound(Grid) + 1

                Vertex(i) = Grid(i - (UBound(Grid) + 1)).Position - Element_property.Thickness * 0.5 * oz

            Next i


            Field.Inertia = MIP.Inertia_8_node(Vertex, Field.Volume, Field.Centroid)

            Field.Density = Element_property.Material.Rho

            Field.Mass = Field.Volume * Field.Density

            ' Adding nonstructural mass

            Field.Mass += Face(0).area * Element_property.NSM

        End If

        If Element_property.Type = "PCOMP" Then

            Dim Volume(UBound(Element_property.Layers)) As Double

            Dim Mass(UBound(Element_property.Layers)) As Double

            Dim Centroid(UBound(Element_property.Layers)) As Vector

            Dim tmp_Inertia(,) As Double


            For i As Integer = 0 To UBound(Element_property.Layers)

                For j As Integer = 0 To UBound(Grid)

                    Vertex(j) = Grid(j).Position + Element_property.Layers(i).Thickness * 0.5 * oz

                Next j

                For j As Integer = UBound(Grid) + 1 To 2 * UBound(Grid) + 1

                    Vertex(j) = Grid(j - (UBound(Grid) + 1)).Position - Element_property.Layers(i).Thickness * 0.5 * oz

                Next j

                tmp_Inertia = MIP.Inertia_8_node(Vertex, Volume(i), Centroid(i))

                Field.Volume += Volume(i)

                Mass(i) = Volume(i) * Element_property.Layers(i).Material.Rho

                Field.Mass += Mass(i)

                For m As Integer = 0 To 2

                    For n As Integer = 0 To 2

                        Field.Inertia(m, n) = tmp_Inertia(m, n)

                    Next n

                Next m

            Next i

            For i As Integer = 0 To UBound(Element_property.Layers)

                Field.Centroid.Coord(0) += Centroid(i).Coord(0) * Mass(i) / Field.Mass

                Field.Centroid.Coord(1) += Centroid(i).Coord(1) * Mass(i) / Field.Mass

                Field.Centroid.Coord(2) += Centroid(i).Coord(2) * Mass(i) / Field.Mass

            Next i

        End If

    End Sub

End Class

<Serializable()> _
Public Class CQUAD8

    Inherits Finite_element

    ''' <summary>
    ''' Area of element
    ''' </summary>
    ''' <remarks></remarks>
    Public Area As Double

    ''' <summary>
    ''' THETA in Nastran formulation
    ''' </summary>
    ''' <remarks></remarks>
    Public Material_orientation_angle As Double

    ''' <summary>
    ''' MCoordinate_system in Nastran formulation
    ''' </summary>
    ''' <remarks></remarks>
    Public Coordinate_system As Object

    ''' <summary>
    ''' ZOFFS in Nastran formulation
    ''' </summary>
    ''' <remarks></remarks>
    Public Offset As Double

    Public T() As Double

    Public TFLAG As Long

    Public Overrides Sub Load()

        ReDim Grid(7)

        ReDim T(3)

        On Error Resume Next

        Number = Val(Fields_array(1))

        Element_property = Val(Fields_array(2))

        Grid(0) = Val(Fields_array(3))

        Grid(1) = Val(Fields_array(4))

        Grid(2) = Val(Fields_array(5))

        Grid(3) = Val(Fields_array(6))

        Grid(4) = Val(Fields_array(7))

        Grid(5) = Val(Fields_array(8))



        Grid(6) = Val(Fields_array(11))

        Grid(7) = Val(Fields_array(12))

        T(0) = Val(Fields_array(13))

        T(1) = Val(Fields_array(14))

        T(2) = Val(Fields_array(15))

        T(3) = Val(Fields_array(16))

        If InStr(Fields_array(17), ".") > 0 Then

            Material_orientation_angle = Val(Fields_array(17))

        Else

            Coordinate_system = Val(Fields_array(17))

        End If

        Offset = Val(Fields_array(18))



        TFLAG = Val(Fields_array(21))

        On Error GoTo 0

    End Sub
    Public Overrides Sub Prepare_Save()

        ReDim Fields_array(16)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(Number)

        Fields_array(2) = my_Str(Element_property)

        Fields_array(3) = my_Str(Grid(0))

        Fields_array(4) = my_Str(Grid(1))

        Fields_array(5) = my_Str(Grid(2))

        Fields_array(6) = my_Str(Grid(3))

        Fields_array(7) = my_Str(Grid(4))

        Fields_array(8) = my_Str(Grid(5))



        Fields_array(11) = my_Str(Grid(6))

        Fields_array(12) = my_Str(Grid(7))

        Fields_array(13) = my_Str(T(0))

        Fields_array(14) = my_Str(T(1))

        Fields_array(15) = my_Str(T(2))

        Fields_array(16) = my_Str(T(3))


        If Material_orientation_angle = Double.MinValue Then

            Fields_array(17) = my_Str(Coordinate_system)

        Else

            Fields_array(17) = my_Str(Material_orientation_angle)

        End If

        Fields_array(18) = my_Str(Offset)

        Fields_array(21) = my_Str(TFLAG)

        On Error GoTo 0

    End Sub

    Public Sub New(Optional ByRef String_in_file As String = "")

        MyBase.New(String_in_file)



    End Sub

    Public Sub New(ByRef Fields_array() As String)

        Me.Fields_array = Fields_array

        Coordinate_system = Long.MinValue

        Material_orientation_angle = Double.MinValue

        Load()

    End Sub

    ''' <summary>
    ''' This sub calculates element mass, volume, centroid and moment of inertia in element's grids coordinate system.
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Calc_volume_mass_inertia()

        Dim MIP As New MIP

        Dim Vertex(2 * UBound(Grid) + 1) As Vector

        Dim oz As Vector = Face(0).oz


        If Element_property.Type = "PSHELL" Then

            For i As Integer = 0 To UBound(Grid)

                Vertex(i) = Grid(i).Position + Element_property.Thickness * 0.5 * oz

            Next i

            For i As Integer = UBound(Grid) + 1 To 2 * UBound(Grid) + 1

                Vertex(i) = Grid(i - (UBound(Grid) + 1)).Position - Element_property.Thickness * 0.5 * oz

            Next i


            Field.Inertia = MIP.Inertia_8_node(Vertex, Field.Volume, Field.Centroid)

            Field.Density = Element_property.Material.Rho

            Field.Mass = Field.Volume * Field.Density

            ' Adding nonstructural mass

            Field.Mass += Face(0).area * Element_property.NSM

        End If

        If Element_property.Type = "PCOMP" Then

            Dim Volume(UBound(Element_property.Layer)) As Double

            Dim Mass(UBound(Element_property.Layer)) As Double

            Dim Centroid(UBound(Element_property.Layer)) As Vector

            Dim tmp_Inertia(,) As Double


            For i As Integer = 0 To UBound(Element_property.Layer)

                For j As Integer = 0 To UBound(Grid)

                    Vertex(j) = Grid(j).Position + Element_property.Layer(i).Thickness * 0.5 * oz

                Next j

                For j As Integer = UBound(Grid) + 1 To 2 * UBound(Grid) + 1

                    Vertex(j) = Grid(j - (UBound(Grid) + 1)).Position - Element_property.Layer(i).Thickness * 0.5 * oz

                Next j

                tmp_Inertia = MIP.Inertia_8_node(Vertex, Volume(i), Centroid(i))

                Field.Volume += Volume(i)

                Mass(i) = Volume(i) * Element_property.Layer(i).Material.Rho

                Field.Mass += Mass(i)

                For m As Integer = 0 To 2

                    For n As Integer = 0 To 2

                        Field.Inertia(m, n) = tmp_Inertia(m, n)

                    Next n

                Next m

            Next i

            For i As Integer = 0 To UBound(Element_property.Layer)

                Field.Centroid.Coord(0) += Centroid(i).Coord(0) * Mass(i) / Field.Mass

                Field.Centroid.Coord(1) += Centroid(i).Coord(1) * Mass(i) / Field.Mass

                Field.Centroid.Coord(2) += Centroid(i).Coord(2) * Mass(i) / Field.Mass

            Next i

        End If

    End Sub

End Class

<Serializable()> _
Public Class CSHEAR

    Inherits Finite_element

    ''' <summary>
    ''' Area of element
    ''' </summary>
    ''' <remarks></remarks>
    Public Area As Double

    Public Overrides Sub Load()

        ReDim Grid(3)

        On Error Resume Next

        Number = Val(Fields_array(1))

        Element_property = Val(Fields_array(2))

        Grid(0) = Val(Fields_array(3))

        Grid(1) = Val(Fields_array(4))

        Grid(2) = Val(Fields_array(5))

        Grid(3) = Val(Fields_array(6))

        On Error GoTo 0

    End Sub
    Public Overrides Sub Prepare_Save()

        ReDim Fields_array(16)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(Number)

        Fields_array(2) = my_Str(Element_property)

        Fields_array(3) = my_Str(Grid(0))

        Fields_array(4) = my_Str(Grid(1))

        Fields_array(5) = my_Str(Grid(2))

        Fields_array(6) = my_Str(Grid(3))

        On Error GoTo 0

    End Sub

    Public Sub New(Optional ByRef String_in_file As String = "")

        MyBase.New(String_in_file)

    End Sub

    Public Sub New(ByRef Fields_array() As String)

        Me.Fields_array = Fields_array

        Load()

    End Sub

    ''' <summary>
    ''' This sub calculates element mass, volume, centroid and moment of inertia in element's grids coordinate system.
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Calc_volume_mass_inertia()

        Dim MIP As New MIP

        Dim Vertex(2 * UBound(Grid) + 1) As Vector

        Dim oz As Vector = Face(0).oz



        For i As Integer = 0 To UBound(Grid)

            Vertex(i) = Grid(i).Position + Element_property.Thickness * 0.5 * oz

        Next i

        For i As Integer = UBound(Grid) + 1 To 2 * UBound(Grid) + 1

            Vertex(i) = Grid(i).Position - Element_property.Thickness * 0.5 * oz

        Next i



        Field.Inertia = MIP.Inertia_8_node(Vertex, Field.Volume, Field.Centroid)

        Field.Density = Element_property.Material.Rho

        Field.Mass = Field.Volume * Field.Density

        ' Adding nonstructural mass

        Field.Mass += Face(0).area * Element_property.NSM

    End Sub

End Class

<Serializable()> _
Public Class CTETRA

    Inherits Finite_element


    Public Overrides Sub Load()

        ReDim Grid(9)

        On Error Resume Next

        Number = Val(Fields_array(1))

        Element_property = Val(Fields_array(2))

        Grid(0) = Val(Fields_array(3))

        Grid(1) = Val(Fields_array(4))

        Grid(2) = Val(Fields_array(5))

        Grid(3) = Val(Fields_array(6))

        Grid(4) = Val(Fields_array(7))

        Grid(5) = Val(Fields_array(8))



        Grid(6) = Val(Fields_array(11))

        Grid(7) = Val(Fields_array(12))

        Grid(8) = Val(Fields_array(13))

        Grid(9) = Val(Fields_array(14))

        On Error GoTo 0

    End Sub
    Public Overrides Sub Prepare_Save()

        ReDim Fields_array(14)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(Number)

        Fields_array(2) = my_Str(Element_property)

        Fields_array(3) = my_Str(Grid(0))

        Fields_array(4) = my_Str(Grid(1))

        Fields_array(5) = my_Str(Grid(2))

        Fields_array(6) = my_Str(Grid(3))

        Fields_array(7) = my_Str(Grid(4))

        Fields_array(8) = my_Str(Grid(5))



        Fields_array(11) = my_Str(Grid(6))

        Fields_array(12) = my_Str(Grid(7))

        Fields_array(13) = my_Str(Grid(8))

        Fields_array(14) = my_Str(Grid(9))

        On Error GoTo 0

    End Sub

    Public Sub New(Optional ByRef String_in_file As String = "")

        MyBase.New(String_in_file)

    End Sub

    Public Sub New(ByRef Fields_array() As String)

        Me.Fields_array = Fields_array

        Load()

    End Sub

    ''' <summary>
    ''' This sub calculates element mass, volume, centroid and moment of inertia in element's grids coordinate system.
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Calc_volume_mass_inertia()

        Dim MIP As New MIP

        Dim Vertex(UBound(Grid)) As Vector


        For i As Integer = 0 To UBound(Grid)

            Vertex(i) = Grid(i).Position

        Next i


        MIP.Inertia_4_node(Vertex, Field.Volume, Field.Centroid, Field.Inertia)

        Field.Density = Element_property.Material.Rho

        Field.Mass = Field.Volume * Field.Density

    End Sub
End Class

<Serializable()> _
Public Class CPENTA

    Inherits Finite_element

    Public Overrides Sub Load()

        ReDim Grid(14)

        On Error Resume Next

        Number = Val(Fields_array(1))

        Element_property = Val(Fields_array(2))

        Grid(0) = Val(Fields_array(3))

        Grid(1) = Val(Fields_array(4))

        Grid(2) = Val(Fields_array(5))

        Grid(3) = Val(Fields_array(6))

        Grid(4) = Val(Fields_array(7))

        Grid(5) = Val(Fields_array(8))



        Grid(6) = Val(Fields_array(11))

        Grid(7) = Val(Fields_array(12))

        Grid(8) = Val(Fields_array(13))

        Grid(9) = Val(Fields_array(14))

        Grid(10) = Val(Fields_array(15))

        Grid(11) = Val(Fields_array(16))

        Grid(12) = Val(Fields_array(17))

        Grid(13) = Val(Fields_array(18))


        Grid(14) = Val(Fields_array(21))

        On Error GoTo 0

    End Sub
    Public Overrides Sub Prepare_Save()

        ReDim Fields_array(21)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(Number)

        Fields_array(2) = my_Str(Element_property)

        Fields_array(3) = my_Str(Grid(0))

        Fields_array(4) = my_Str(Grid(1))

        Fields_array(5) = my_Str(Grid(2))

        Fields_array(6) = my_Str(Grid(3))

        Fields_array(7) = my_Str(Grid(4))

        Fields_array(8) = my_Str(Grid(5))



        Fields_array(11) = my_Str(Grid(6))

        Fields_array(12) = my_Str(Grid(7))

        Fields_array(13) = my_Str(Grid(8))

        Fields_array(14) = my_Str(Grid(9))

        Fields_array(15) = my_Str(Grid(10))

        Fields_array(16) = my_Str(Grid(11))

        Fields_array(17) = my_Str(Grid(12))

        Fields_array(18) = my_Str(Grid(13))


        Fields_array(21) = my_Str(Grid(14))

        On Error GoTo 0

    End Sub

    Public Sub New(Optional ByRef String_in_file As String = "")

        MyBase.New(String_in_file)

    End Sub

    Public Sub New(ByRef Fields_array() As String)

        Me.Fields_array = Fields_array

        Load()

    End Sub

    ''' <summary>
    ''' This sub calculates element mass, volume, centroid and moment of inertia in element's grids coordinate system.
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Calc_volume_mass_inertia()

        Dim MIP As New MIP

        Dim Vertex(UBound(Grid)) As Vector


        For i As Integer = 0 To UBound(Grid)

            Vertex(i) = Grid(i).Position

        Next i


        Field.Inertia = MIP.Inertia_6_node(Vertex, Field.Volume, Field.Centroid)

        Field.Density = Element_property.Material.Rho

        Field.Mass = Field.Volume * Field.Density

    End Sub
End Class

<Serializable()> _
Public Class CHEXA

    Inherits Finite_element

    Public Overrides Sub Load()

        ReDim Grid(19)

        On Error Resume Next

        Number = Val(Fields_array(1))

        Element_property = Val(Fields_array(2))

        Grid(0) = Val(Fields_array(3))

        Grid(1) = Val(Fields_array(4))

        Grid(2) = Val(Fields_array(5))

        Grid(3) = Val(Fields_array(6))

        Grid(4) = Val(Fields_array(7))

        Grid(5) = Val(Fields_array(8))



        Grid(6) = Val(Fields_array(11))

        Grid(7) = Val(Fields_array(12))

        Grid(8) = Val(Fields_array(13))

        Grid(9) = Val(Fields_array(14))

        Grid(10) = Val(Fields_array(15))

        Grid(11) = Val(Fields_array(16))

        Grid(12) = Val(Fields_array(17))

        Grid(13) = Val(Fields_array(18))


        Grid(14) = Val(Fields_array(21))

        Grid(15) = Val(Fields_array(22))

        Grid(16) = Val(Fields_array(23))

        Grid(17) = Val(Fields_array(24))

        Grid(18) = Val(Fields_array(25))

        Grid(19) = Val(Fields_array(26))

        On Error GoTo 0

    End Sub
    Public Overrides Sub Prepare_Save()

        ReDim Fields_array(26)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(Number)

        Fields_array(2) = my_Str(Element_property)

        Fields_array(3) = my_Str(Grid(0))

        Fields_array(4) = my_Str(Grid(1))

        Fields_array(5) = my_Str(Grid(2))

        Fields_array(6) = my_Str(Grid(3))

        Fields_array(7) = my_Str(Grid(4))

        Fields_array(8) = my_Str(Grid(5))



        Fields_array(11) = my_Str(Grid(6))

        Fields_array(12) = my_Str(Grid(7))

        Fields_array(13) = my_Str(Grid(8))

        Fields_array(14) = my_Str(Grid(9))

        Fields_array(15) = my_Str(Grid(10))

        Fields_array(16) = my_Str(Grid(11))

        Fields_array(17) = my_Str(Grid(12))

        Fields_array(18) = my_Str(Grid(13))


        Fields_array(21) = my_Str(Grid(14))

        Fields_array(22) = my_Str(Grid(15))

        Fields_array(23) = my_Str(Grid(16))

        Fields_array(24) = my_Str(Grid(17))

        Fields_array(25) = my_Str(Grid(18))

        Fields_array(26) = my_Str(Grid(19))

        On Error GoTo 0

    End Sub

    Public Sub New(Optional ByRef String_in_file As String = "")

        MyBase.New(String_in_file)

    End Sub

    Public Sub New(ByRef Fields_array() As String)

        Me.Fields_array = Fields_array

        Load()

    End Sub

    ''' <summary>
    ''' This sub calculates element mass, volume, centroid and moment of inertia in element's grids coordinate system.
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Calc_volume_mass_inertia()

        Dim MIP As New MIP

        Dim Vertex(UBound(Grid)) As Vector


        For i As Integer = 0 To UBound(Grid)

            Vertex(i) = Grid(i).Position

        Next i


        Field.Inertia = MIP.Inertia_8_node(Vertex, Field.Volume, Field.Centroid)

        Field.Density = Element_property.Material.Rho

        Field.Mass = Field.Volume * Field.Density

    End Sub
End Class

<Serializable()> _
Public MustInherit Class Loads

    Inherits Model_item

    ''' <summary>
    ''' SID in Nastran formulation
    ''' </summary>
    ''' <remarks></remarks>
    Public SID As Long

    ''' <summary>
    ''' G in Nastran formulation
    ''' </summary>
    ''' <remarks></remarks>
    Public Grid_ As Object

    ''' <summary>
    ''' Coordinate_system in Nastran formulation
    ''' </summary>
    ''' <remarks></remarks>
    Public Coordinate_system As Object

    ''' <summary>
    ''' F for forces, M for moments, A for accelerations in Nastran formulation
    ''' </summary>
    ''' <remarks></remarks>
    Public Scale_factor As Double

    Public Sub New(Optional ByRef String_in_file As String = "")

        MyBase.New(String_in_file)

    End Sub

    Public Sub New(ByRef Fields_array() As String)

        Me.Fields_array = Fields_array

        Load()

    End Sub
End Class

<Serializable()> _
Public Class FORCE

    Inherits Loads

    ''' <summary>
    ''' Components of force vector
    ''' </summary>
    ''' <remarks></remarks>
    Public N() As Double

    Public Overrides Sub Load()

        ReDim N(2)

        On Error Resume Next

        SID = Val(Fields_array(1))

        Grid_ = Val(Fields_array(2))

        If Fields_array(3) = "" Then

            Coordinate_system = -1

        Else

            Coordinate_system = Val(Fields_array(3))

        End If

        Scale_factor = Val(Fields_array(4))

        N(0) = Val(Fields_array(5))

        N(1) = Val(Fields_array(6))

        N(2) = Val(Fields_array(7))

        On Error GoTo 0

    End Sub
    Public Overrides Sub Prepare_Save()

        ReDim Fields_array(7)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(SID)

        Fields_array(2) = my_Str(Grid_)

        If Coordinate_system > -1 Then

            Fields_array(3) = my_Str(Coordinate_system)

        End If

        Fields_array(4) = my_Str(Scale_factor)

        Fields_array(5) = my_Str(N(0))

        Fields_array(6) = my_Str(N(1))

        Fields_array(7) = my_Str(N(2))

        On Error GoTo 0

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
Public Class FORCE1

    Inherits Loads

    Public Grid() As Object

    Public Overrides Sub Load()

        ReDim Grid(1)

        On Error Resume Next

        SID = Val(Fields_array(1))

        Grid_ = Val(Fields_array(2))

        Scale_factor = Val(Fields_array(3))

        Grid(0) = Val(Fields_array(4))

        Grid(1) = Val(Fields_array(5))

        On Error GoTo 0

    End Sub
    Public Overrides Sub Prepare_Save()

        ReDim Fields_array(5)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(SID)

        Fields_array(2) = my_Str(Grid_)

        Fields_array(3) = my_Str(Scale_factor)

        Fields_array(4) = my_Str(Grid(0))

        Fields_array(5) = my_Str(Grid(1))

        On Error GoTo 0

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
Public Class FORCE2

    Inherits Loads

    Public Grid() As Object

    Public Overrides Sub Load()

        ReDim Grid(3)

        On Error Resume Next

        SID = Val(Fields_array(1))

        Grid_ = Val(Fields_array(2))

        Scale_factor = Val(Fields_array(3))

        Grid(0) = Val(Fields_array(4))

        Grid(1) = Val(Fields_array(5))

        Grid(2) = Val(Fields_array(6))

        Grid(3) = Val(Fields_array(7))

        On Error GoTo 0

    End Sub
    Public Overrides Sub Prepare_Save()

        ReDim Fields_array(7)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(SID)

        Fields_array(2) = my_Str(Grid_)

        Fields_array(3) = my_Str(Scale_factor)

        Fields_array(4) = my_Str(Grid(0))

        Fields_array(5) = my_Str(Grid(1))

        Fields_array(6) = my_Str(Grid(2))

        Fields_array(7) = my_Str(Grid(3))

        On Error GoTo 0

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
Public Class MOMENT

    Inherits Loads

    ''' <summary>
    ''' Components of MOMENT vector
    ''' </summary>
    ''' <remarks></remarks>
    Public N() As Double

    Public Overrides Sub Load()

        ReDim N(2)

        On Error Resume Next

        SID = Val(Fields_array(1))

        Grid_ = Val(Fields_array(2))

        If Fields_array(3) = "" Then

            Coordinate_system = -1

        Else

            Coordinate_system = Val(Fields_array(3))

        End If

        Scale_factor = Val(Fields_array(4))

        N(0) = Val(Fields_array(5))

        N(1) = Val(Fields_array(6))

        N(2) = Val(Fields_array(7))

        On Error GoTo 0

    End Sub
    Public Overrides Sub Prepare_Save()

        ReDim Fields_array(7)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(SID)

        Fields_array(2) = my_Str(Grid_)

        If Coordinate_system > -1 Then

            Fields_array(3) = my_Str(Coordinate_system)

        End If

        Fields_array(4) = my_Str(Scale_factor)

        Fields_array(5) = my_Str(N(0))

        Fields_array(6) = my_Str(N(1))

        Fields_array(7) = my_Str(N(2))

        On Error GoTo 0

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
Public Class MOMENT1

    Inherits Loads

    Public Grid() As Object

    Public Overrides Sub Load()

        ReDim Grid(1)

        On Error Resume Next

        SID = Val(Fields_array(1))

        Grid_ = Val(Fields_array(2))

        Scale_factor = Val(Fields_array(3))

        Grid(0) = Val(Fields_array(4))

        Grid(1) = Val(Fields_array(5))

        On Error GoTo 0

    End Sub
    Public Overrides Sub Prepare_Save()

        ReDim Fields_array(5)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(SID)

        Fields_array(2) = my_Str(Grid_)

        Fields_array(3) = my_Str(Scale_factor)

        Fields_array(4) = my_Str(Grid(0))

        Fields_array(5) = my_Str(Grid(1))

        On Error GoTo 0

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
Public Class MOMENT2

    Inherits Loads

    Public Grid() As Object

    Public Overrides Sub Load()

        ReDim Grid(3)

        On Error Resume Next

        SID = Val(Fields_array(1))

        Grid_ = Val(Fields_array(2))

        Scale_factor = Val(Fields_array(3))

        Grid(0) = Val(Fields_array(4))

        Grid(1) = Val(Fields_array(5))

        Grid(2) = Val(Fields_array(6))

        Grid(3) = Val(Fields_array(7))

        On Error GoTo 0

    End Sub
    Public Overrides Sub Prepare_Save()

        ReDim Fields_array(7)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(SID)

        Fields_array(2) = my_Str(Grid_)

        Fields_array(3) = my_Str(Scale_factor)

        Fields_array(4) = my_Str(Grid(0))

        Fields_array(5) = my_Str(Grid(1))

        Fields_array(6) = my_Str(Grid(2))

        Fields_array(7) = my_Str(Grid(3))

        On Error GoTo 0

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
Public Class PLOAD

    Inherits Model_item

    Public SID As Long

    Public Grid() As Object

    ''' <summary>
    ''' P in NASTRAN formulation
    ''' </summary>
    ''' <remarks></remarks>
    Public Pressure As Double

    Public Overrides Sub Load()

        ReDim Grid(3)

        On Error Resume Next

        SID = Val(Fields_array(1))

        Pressure = Val(Fields_array(2))

        Grid(0) = Val(Fields_array(3))

        Grid(1) = Val(Fields_array(4))

        Grid(2) = Val(Fields_array(5))

        Grid(3) = Val(Fields_array(6))

        On Error GoTo 0

    End Sub
    Public Overrides Sub Prepare_Save()

        ReDim Fields_array(6)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(SID)

        Fields_array(2) = my_Str(Pressure)

        Fields_array(3) = my_Str(Grid(0))

        Fields_array(4) = my_Str(Grid(1))

        Fields_array(5) = my_Str(Grid(2))

        Fields_array(6) = my_Str(Grid(3))

        On Error GoTo 0

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
Public Class PLOAD1

    Inherits Model_item

    ''' <summary>
    ''' SID in NASTRAN formulation
    ''' </summary>
    ''' <remarks></remarks>
    Public SID As Long

    ''' <summary>
    ''' EID in NASTRAN formulation
    ''' </summary>
    ''' <remarks></remarks>
    Public Element_number As Object

    ''' <summary>
    ''' TYPE in NASTRAN formulation
    ''' </summary>
    ''' <remarks></remarks>
    Public Load_type As String

    ''' <summary>
    ''' SCALE in NASTRAN formulation
    ''' </summary>
    ''' <remarks></remarks>
    Public Load_scale As String

    Public X(), P() As Double


    Public Overrides Sub Load()

        ReDim X(1), P(1)

        On Error Resume Next

        SID = Val(Fields_array(1))

        Element_number = Val(Fields_array(2))

        Load_type = Val(Fields_array(3))

        Load_scale = Val(Fields_array(4))

        X(0) = Val(Fields_array(5))

        P(0) = Val(Fields_array(6))

        X(1) = Val(Fields_array(7))

        P(1) = Val(Fields_array(8))

        On Error GoTo 0

    End Sub
    Public Overrides Sub Prepare_Save()

        ReDim Fields_array(8)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(SID)

        Fields_array(2) = my_Str(Element_number)

        Fields_array(3) = my_Str(Load_type)

        Fields_array(4) = my_Str(Load_scale)

        Fields_array(5) = my_Str(X(0))

        Fields_array(6) = my_Str(P(0))

        Fields_array(7) = my_Str(X(1))

        Fields_array(8) = my_Str(P(1))

        On Error GoTo 0

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
Public Class PLOAD2

    Inherits Model_item

    Public SID As Long

    ''' <summary>
    ''' P in NASTRAN formulation
    ''' </summary>
    ''' <remarks></remarks>
    Public Pressure As Double

    ''' <summary>
    ''' EID1, EID2 and etc in NASTRAN formulation
    ''' </summary>
    ''' <remarks></remarks>
    Public Element_number() As Object

    ''' <summary>
    ''' THRU_Flag is used, when if is needs to determine Element_numbers range
    ''' If Ubound(Element_number) = 1 and THRU_Flag= True then output saves range,
    ''' else it saves only two additional elements
    ''' </summary>
    ''' <remarks></remarks>
    Public THRU_Flag As Boolean

    Public Overrides Sub Load()

        ReDim Element_number(0)

        On Error Resume Next

        SID = Val(Fields_array(1))

        Pressure = Val(Fields_array(2))

        If Fields_array(4) = "THRU" Then

            THRU_Flag = True

            ReDim Element_number(1)

            Element_number(0) = Val(Fields_array(3))

            Element_number(1) = Val(Fields_array(5))

            On Error GoTo 0

            Exit Sub

        Else

            ReDim Preserve Element_number(5)

            Element_number(0) = Val(Fields_array(3))

            Element_number(1) = Val(Fields_array(4))

            Element_number(2) = Val(Fields_array(5))

            Element_number(3) = Val(Fields_array(6))

            Element_number(4) = Val(Fields_array(7))

            Element_number(5) = Val(Fields_array(8))

        End If

        On Error GoTo 0

    End Sub
    Public Overrides Sub Prepare_Save()

        ReDim Fields_array(8)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(SID)

        Fields_array(2) = my_Str(Pressure)

        Fields_array(3) = my_Str(Element_number(0))

        If UBound(Element_number) = 1 And THRU_Flag = True Then

            Fields_array(4) = "THRU"

            Fields_array(5) = my_Str(Element_number(1))

        Else

            Fields_array(4) = my_Str(Element_number(1))

            Fields_array(5) = my_Str(Element_number(2))

            Fields_array(6) = my_Str(Element_number(3))

            Fields_array(7) = my_Str(Element_number(4))

            Fields_array(8) = my_Str(Element_number(5))

        End If

        On Error GoTo 0

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
Public Class PLOAD4

    Inherits Model_item

    Public SID As Long

    ''' <summary>
    ''' EID, EID1 and EID2 in NASTRAN formulation
    ''' </summary>
    ''' <remarks></remarks>
    Public Element_number() As Object

    ''' <summary>
    ''' P in NASTRAN formulation
    ''' </summary>
    ''' <remarks></remarks>
    Public Pressure(3) As Double

    ''' <summary>
    ''' "G1" and "G3 or G4" in NASTRAN formulation
    ''' </summary>
    ''' <remarks></remarks>
    Public Grid() As Object

    ''' <summary>
    ''' Coordinate_system in Nastran formulation
    ''' </summary>
    ''' <remarks></remarks>
    Public Coordinate_system As Object

    ''' <summary>
    ''' N1,N2 and N3 in NASTRAN formulation
    ''' </summary>
    ''' <remarks></remarks>
    Public N() As Long

    Public SORL As String

    Public LDIR As String

    ''' <summary>
    ''' THRU_Flag is used, when if is needs to determine Element_numbers range
    ''' If Ubound(Element_number) = 1 and THRU_Flag= True then output saves range,
    ''' else it saves only two additional elements
    ''' </summary>
    ''' <remarks></remarks>
    Public THRU_Flag As Boolean



    Public Overrides Sub Load()

        ReDim Pressure(3)

        ReDim N(2)

        On Error Resume Next

        SID = Val(Fields_array(1))

        If Fields_array(7) = "THRU" Then

            THRU_Flag = True

            ReDim Element_number(1)

            Element_number(0) = Val(Fields_array(2))

            Element_number(1) = Val(Fields_array(8))

        Else

            ReDim Element_number(0)

            ReDim Grid(1)

            Element_number(0) = Val(Fields_array(2))

            Grid(0) = Val(Fields_array(7))

            Grid(1) = Val(Fields_array(8))

        End If



        Pressure(0) = Val(Fields_array(3))

        Pressure(1) = Val(Fields_array(4))

        Pressure(2) = Val(Fields_array(5))

        Pressure(3) = Val(Fields_array(6))


        Coordinate_system = Val(Fields_array(11))

        N(0) = Val(Fields_array(12))

        N(1) = Val(Fields_array(13))

        N(2) = Val(Fields_array(14))

        SORL = Fields_array(15)

        LDIR = Fields_array(16)

        On Error GoTo 0

    End Sub
    Public Overrides Sub Prepare_Save()

        ReDim Fields_array(16)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(SID)

        Fields_array(2) = my_Str(Element_number(0))

        If UBound(Element_number) = 1 And THRU_Flag = True Then

            Fields_array(7) = "THRU"

            Fields_array(8) = my_Str(Element_number(1))

        Else

            Fields_array(7) = my_Str(Grid(0))

            Fields_array(8) = my_Str(Grid(1))

        End If


        Fields_array(3) = my_Str(Pressure(0))

        Fields_array(4) = my_Str(Pressure(1))

        Fields_array(5) = my_Str(Pressure(2))

        Fields_array(6) = my_Str(Pressure(3))


        Fields_array(11) = my_Str(Coordinate_system)

        Fields_array(12) = my_Str(N(0))

        Fields_array(13) = my_Str(N(1))

        Fields_array(14) = my_Str(N(2))

        Fields_array(15) = SORL

        Fields_array(16) = LDIR

        On Error GoTo 0

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
Public Class GRAV

    Inherits Loads

    ''' <summary>
    ''' Components of force vector
    ''' </summary>
    ''' <remarks></remarks>
    Public N() As Double

    Public MB As Long

    Public Overrides Sub Load()

        ReDim N(2)

        On Error Resume Next

        SID = Val(Fields_array(1))

        Coordinate_system = Val(Fields_array(2))

        Scale_factor = Val(Fields_array(3))

        N(0) = Val(Fields_array(4))

        N(1) = Val(Fields_array(5))

        N(2) = Val(Fields_array(6))

        MB = Val(Fields_array(7))

        On Error GoTo 0

    End Sub
    Public Overrides Sub Prepare_Save()

        ReDim Fields_array(7)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(SID)

        Fields_array(2) = my_Str(Coordinate_system)

        Fields_array(3) = my_Str(Scale_factor)

        Fields_array(4) = my_Str(N(0))

        Fields_array(5) = my_Str(N(1))

        Fields_array(6) = my_Str(N(2))

        Fields_array(7) = my_Str(MB)

        On Error GoTo 0

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
Public Class RFORCE

    Inherits Loads

    ''' <summary>
    ''' Rectangular components of rotation vector
    ''' </summary>
    ''' <remarks></remarks>
    Public R() As Double


    Public METHOD As Long

    ''' <summary>
    ''' Scale factor for acceleration
    ''' </summary>
    ''' <remarks></remarks>
    Public RACC As Double

    Public MB As Long

    Public Overrides Sub Load()

        ReDim R(2)

        On Error Resume Next

        SID = Val(Fields_array(1))

        Grid_ = Val(Fields_array(2))

        Coordinate_system = Val(Fields_array(3))

        Scale_factor = Val(Fields_array(4))

        R(0) = Val(Fields_array(5))

        R(1) = Val(Fields_array(6))

        R(2) = Val(Fields_array(7))

        METHOD = Val(Fields_array(8))


        RACC = Val(Fields_array(11))

        MB = Val(Fields_array(11))

        On Error GoTo 0

    End Sub
    Public Overrides Sub Prepare_Save()

        ReDim Fields_array(7)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(SID)

        Fields_array(2) = my_Str(Grid_)

        Fields_array(3) = my_Str(Coordinate_system)

        Fields_array(4) = my_Str(Scale_factor)

        Fields_array(5) = my_Str(R(0))

        Fields_array(6) = my_Str(R(1))

        Fields_array(7) = my_Str(R(2))

        Fields_array(8) = my_Str(METHOD)


        Fields_array(11) = my_Str(RACC)

        Fields_array(12) = my_Str(MB)

        On Error GoTo 0

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
Public Class LOAD

    Inherits Model_item

    ''' <summary>
    ''' SID in NASTRAN formulation
    ''' </summary>
    ''' <remarks></remarks>
    Public SID As Long

    ''' <summary>
    ''' S in NASTRAN formulation
    ''' </summary>
    ''' <remarks></remarks>
    Public Overall_scale_factor As Double

    ''' <summary>
    ''' S1,S2 and etc in NASTRAN formulation
    ''' </summary>
    ''' <remarks></remarks>
    Public Scale_factor() As Double

    ''' <summary>
    ''' L1,L2 and etc in NASTRAN formulation
    ''' </summary>
    ''' <remarks></remarks>
    Public Load_set_id() As Long


    Public Overrides Sub Load()

        Dim i_Load_set As Long = -1

        ReDim Scale_factor(0), Load_set_id(0)

        Dim tmp_String As String

        On Error Resume Next

        SID = Val(Fields_array(1))

        Overall_scale_factor = Val(Fields_array(2))

        If Fields_array(3) <> "" Then

            i_Load_set += 1

            ReDim Preserve Scale_factor(i_Load_set), Load_set_id(i_Load_set)

            Scale_factor(i_Load_set) = Val(Fields_array(3))

            Load_set_id(i_Load_set) = Val(Fields_array(4))

        End If

        If Fields_array(5) <> "" Then

            i_Load_set += 1

            ReDim Preserve Scale_factor(i_Load_set), Load_set_id(i_Load_set)

            Scale_factor(i_Load_set) = Val(Fields_array(5))

            Load_set_id(i_Load_set) = Val(Fields_array(6))

        End If


        If Fields_array(7) <> "" Then

            i_Load_set += 1

            ReDim Preserve Scale_factor(i_Load_set), Load_set_id(i_Load_set)

            Scale_factor(i_Load_set) = Val(Fields_array(7))

            Load_set_id(i_Load_set) = Val(Fields_array(8))

        End If

        For i As Integer = 21 To UBound(Fields_array) Step 10

            For j As Integer = i To i + 6 Step 2

                tmp_String = Fields_array(j)

                If tmp_String <> "" Then

                    i_Load_set += 1

                    ReDim Preserve Scale_factor(i_Load_set), Load_set_id(i_Load_set)

                    Scale_factor(i_Load_set) = Val(Fields_array(j))

                    Load_set_id(i_Load_set) = Val(Fields_array(j + 1))

                Else

                End If

            Next j

        Next i

        On Error GoTo 0

    End Sub

    Public Overrides Sub Prepare_Save()

        Dim N_lines_for_load_sets As Integer = 0

        Dim i_Load_set As Integer = 0

        If UBound(Scale_factor) = UBound(Scale_factor) \ 4 Then

            N_lines_for_load_sets = UBound(Scale_factor) \ 4

        Else

            N_lines_for_load_sets = UBound(Scale_factor) \ 4 + 1

        End If

        ReDim Fields_array(10 + N_lines_for_load_sets * 10)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(SID)

        Fields_array(2) = my_Str(Overall_scale_factor)

        Fields_array(3) = my_Str(Scale_factor(0))

        Fields_array(4) = my_Str(Load_set_id(0))

        Fields_array(5) = my_Str(Scale_factor(1))

        Fields_array(6) = my_Str(Load_set_id(1))

        Fields_array(7) = my_Str(Scale_factor(2))

        Fields_array(8) = my_Str(Load_set_id(2))

        i_Load_set = 3

        For i As Integer = 1 To N_lines_for_load_sets

            For j As Integer = 1 To 7 Step 2

                Fields_array(10 * i + j) = my_Str(Scale_factor(i_Load_set))

                Fields_array(10 * i + j + 1) = my_Str(Load_set_id(i_Load_set))

            Next j

        Next i

        On Error GoTo 0

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
Public Class ACCEL

    Inherits Model_item

    Public SID As Long

    Public Coordinate_system As Object

    Public N() As Double

    Public DIR As String

    Public Location(), Value() As Double

    Public Overrides Sub Load()

        ReDim N(2)

        On Error Resume Next

        SID = Val(Fields_array(1))

        Coordinate_system = Val(Fields_array(2))

        N(0) = Val(Fields_array(3))

        N(1) = Val(Fields_array(4))

        N(2) = Val(Fields_array(5))

        DIR = Fields_array(5)

        Dim i_Loc As Long = 0

        For i As Long = 10 To UBound(Fields_array) Step 10

            For j As Long = 1 To 8 Step 2

                If Fields_array(i + j) <> "" Then

                    ReDim Preserve Location(i_Loc), Value(i_Loc)

                    Location(i_Loc) = Val(Fields_array(i + j))

                    Value(i_Loc) = Val(Fields_array(i + j + 1))

                End If


            Next j

        Next i

        On Error GoTo 0

    End Sub

    Public Overrides Sub Prepare_Save()

        Dim N_lines_for_Locs As Integer = 0

        Dim i_Loc As Integer = 0

        If UBound(Location) = UBound(Location) \ 4 Then

            N_lines_for_Locs = UBound(Location) \ 4

        Else

            N_lines_for_Locs = UBound(Location) \ 4 + 1

        End If

        ReDim Fields_array(10 + N_lines_for_Locs * 10)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(SID)

        Fields_array(2) = my_Str(Coordinate_system)

        Fields_array(3) = my_Str(N(0))

        Fields_array(4) = my_Str(N(1))

        Fields_array(5) = my_Str(N(2))

        Fields_array(6) = DIR


        For i As Integer = 1 To N_lines_for_Locs

            For j As Integer = 1 To 7 Step 2

                Fields_array(10 * i + j) = my_Str(Location(i_Loc))

                Fields_array(10 * i + j + 1) = my_Str(Value(i_Loc))

                i_Loc += 1

            Next j

        Next i

        On Error GoTo 0

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
Public Class ACCEL1

    Inherits Model_item

    Public SID As Long

    Public Coordinate_system As Object

    ''' <summary>
    ''' Acceleration scale factor
    ''' </summary>
    ''' <remarks></remarks>
    Public Scale_factor As Double

    Public N() As Double

    Public DIR As String

    Public Grid() As Long

    Public Overrides Sub Load()

        ReDim N(2)

        On Error Resume Next

        SID = Val(Fields_array(1))

        Coordinate_system = Val(Fields_array(2))

        Scale_factor = Val(Fields_array(3))

        N(0) = Val(Fields_array(4))

        N(1) = Val(Fields_array(5))

        N(2) = Val(Fields_array(6))

        ReDim Grid(0)

        For i As Integer = 1 To UBound(Fields_array) \ 10

            For j As Integer = 10 * i + 1 To 10 * i + 8

                If Fields_array(j) <> "" Then

                    Grid(UBound(Grid)) = Val(Fields_array(j))

                    ReDim Preserve Grid(UBound(Grid) + 1)

                End If

                If Fields_array(j) = "THRU" Then Exit Sub

            Next j

        Next i

        ReDim Preserve Grid(UBound(Grid) - 1)

        On Error GoTo 0

    End Sub

    Public Overrides Sub Prepare_Save()

        Dim N_lines_for_Grid As Integer

        N_lines_for_Grid = (UBound(Grid)) / 8

        ReDim Fields_array(10 + N_lines_for_Grid * 10)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(SID)

        Fields_array(2) = my_Str(Coordinate_system)

        Fields_array(3) = my_Str(Scale_factor)

        Fields_array(4) = my_Str(N(0))

        Fields_array(5) = my_Str(N(1))

        Fields_array(6) = my_Str(N(2))

        Dim i_Grid As Long = 0

        Dim tmp_Grid_ As Long

        For i As Integer = 1 To N_lines_for_Grid

            For j As Integer = 1 To 8

                tmp_Grid_ = Grid(i_Grid)

                If tmp_Grid_ > 0 Then

                    Fields_array(10 * i + j) = my_Str(Grid(i_Grid))

                    i_Grid += 1

                End If

            Next j

        Next i

        On Error GoTo 0

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
Public Class TEMP

    Inherits Model_item

    Public SID As Long

    Public Grid() As Object

    Public Temperature() As Double

    Public Overrides Sub Load()

        On Error Resume Next

        SID = Val(Fields_array(1))

        If Fields_array(2) <> "" Then

            ReDim Preserve Grid(0), Temperature(0)

            Grid(UBound(Grid)) = Val(Fields_array(2))

            Temperature(UBound(Temperature)) = Val(Fields_array(3))

        End If

        If Fields_array(4) <> "" Then

            ReDim Preserve Grid(1), Temperature(1)

            Grid(UBound(Grid)) = Val(Fields_array(4))

            Temperature(UBound(Temperature)) = Val(Fields_array(5))

        End If

        If Fields_array(6) <> "" Then

            ReDim Preserve Grid(2), Temperature(2)

            Grid(UBound(Grid)) = Val(Fields_array(6))

            Temperature(UBound(Temperature)) = Val(Fields_array(7))

        End If

        On Error GoTo 0



    End Sub

    Public Overrides Sub Prepare_Save()

        ReDim Fields_array(7)

        Dim Grid_ As Long

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(SID)

        Grid_ = Grid(0)

        If Grid_ > 0 Then

            Fields_array(2) = my_Str(Grid(0))

            Fields_array(3) = my_Str(Temperature(0))

        End If

        Grid_ = Grid(1)

        If Grid_ > 0 Then

            Fields_array(4) = my_Str(Grid(1))

            Fields_array(5) = my_Str(Temperature(1))

        End If

        Grid_ = Grid(2)

        If Grid_ > 0 Then

            Fields_array(6) = my_Str(Grid(2))

            Fields_array(7) = my_Str(Temperature(2))

        End If

        On Error GoTo 0

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
Public Class TEMPBC

    Inherits Model_item

    Public SID As Long

    Public BC_Type As String

    Public Temperature() As Double

    Public Grid() As Object

    Public INC As Long

    Public Overrides Sub Load()

        On Error Resume Next

        SID = Val(Fields_array(1))

        BC_Type = Fields_array(2)

        If Fields_array(5) <> "THRU" And Fields_array(7) <> "BY" Then

            ReDim Grid(1), Temperature(0)

            Temperature(0) = Val(Fields_array(3))

            Grid(0) = Val(Fields_array(4))

            Temperature(1) = Val(Fields_array(6))

            INC = Val(Fields_array(8))

            On Error GoTo 0

            Exit Sub
        End If

        If Fields_array(3) <> "" Then

            ReDim Preserve Grid(0), Temperature(0)

            Temperature(UBound(Temperature)) = Val(Fields_array(3))

            Grid(UBound(Grid)) = Val(Fields_array(4))

        End If

        If Fields_array(5) <> "" Then

            ReDim Preserve Grid(1), Temperature(1)

            Temperature(UBound(Temperature)) = Val(Fields_array(5))

            Grid(UBound(Grid)) = Val(Fields_array(6))

        End If

        If Fields_array(7) <> "" Then

            ReDim Preserve Grid(2), Temperature(2)

            Temperature(UBound(Temperature)) = Val(Fields_array(7))

            Grid(UBound(Grid)) = Val(Fields_array(8))

        End If

        On Error GoTo 0



    End Sub

    Public Overrides Sub Prepare_Save()

        ReDim Fields_array(8)

        Dim Grid_ As Long

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(SID)

        Fields_array(2) = BC_Type

        If INC <> 0 Then

            Fields_array(3) = my_Str(Temperature(0))

            Fields_array(4) = my_Str(Grid(0))

            Fields_array(5) = "THRU"

            Fields_array(6) = my_Str(Grid(1))

            Fields_array(7) = "BY"

            Fields_array(8) = my_Str(INC)

            On Error GoTo 0

            Exit Sub
        End If

        Grid_ = Grid(0)

        If Grid_ > 0 Then

            Fields_array(3) = my_Str(Temperature(0))

            Fields_array(4) = my_Str(Grid(0))

        End If

        Grid_ = Grid(1)

        If Grid_ > 0 Then

            Fields_array(5) = my_Str(Temperature(1))

            Fields_array(6) = my_Str(Grid(1))

        End If

        Grid_ = Grid(2)

        If Grid_ > 0 Then

            Fields_array(7) = my_Str(Temperature(2))

            Fields_array(8) = my_Str(Grid(2))

        End If

        On Error GoTo 0

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
Public Class TEMPD

    Inherits Model_item

    Public SID() As Long

    Public Temperature() As Double

    Public Overrides Sub Load()

        On Error Resume Next

        ReDim Preserve SID(0), Temperature(0)

        For i As Integer = 1 To 7 Step 2

            If Fields_array(i) <> "" Then

                SID(UBound(SID)) = Val(Fields_array(i))

                Temperature(UBound(Temperature)) = Val(Fields_array(i + 1))

                ReDim Preserve SID(UBound(SID) + 1), Temperature(UBound(Temperature) + 1)

            End If

        Next i

        ReDim Preserve SID(UBound(SID) - 1), Temperature(UBound(Temperature) - 1)

        On Error GoTo 0

    End Sub

    Public Overrides Sub Prepare_Save()

        ReDim Fields_array(8)

        Dim Load_set_number As Long

        On Error Resume Next

        Fields_array(0) = Type

        For i As Integer = 0 To UBound(SID)

            Load_set_number = SID(i)

            If Load_set_number > 0 Then

                Fields_array(i * 2 + 1) = my_Str(SID(i))

                Fields_array(i * 2 + 2) = my_Str(Temperature(i))

            End If

        Next i

        On Error GoTo 0

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
Public Class TEMPF

    Inherits Model_item

    Public SID As Long

    Public Element_number() As Object

    Public FTEMP As Long

    Public FTABID As Long

    Public Overrides Sub Load()

        On Error Resume Next

        ReDim Preserve Element_number(0)

        SID = Val(Fields_array(1))

        Element_number(0) = Val(Fields_array(2))

        FTEMP = Val(Fields_array(3))

        FTABID = Val(Fields_array(4))

        If Fields_array(22) = "THRU" Then

            ReDim Preserve Element_number(2)

            Element_number(1) = Val(Fields_array(21))

            Element_number(2) = Val(Fields_array(23))

            On Error GoTo 0

            Exit Sub

        End If

        Dim i_EID As Integer = 0

        For i As Integer = 1 To UBound(Fields_array) \ 10

            For j As Integer = 10 * i + 1 To 10 * i + 8

                If Fields_array(j) <> "" Then

                    i_EID += 1

                    ReDim Preserve Element_number(i_EID)

                    Element_number(i_EID) = Val(Fields_array(j))

                End If

            Next j

        Next i

        On Error GoTo 0

    End Sub

    Public Overrides Sub Prepare_Save()

        Dim N_lines_for_EIDs As Integer

        N_lines_for_EIDs = UBound(Element_number) / 8

        ReDim Fields_array(10 + N_lines_for_EIDs * 10)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(SID)

        Fields_array(2) = my_Str(Element_number(0))

        Fields_array(3) = my_Str(FTEMP)

        Fields_array(4) = my_Str(FTABID)

        If UBound(Element_number) = 2 Then

            Fields_array(11) = my_Str(Element_number(1))

            Fields_array(12) = "THRU"

            Fields_array(13) = my_Str(Element_number(2))

            On Error GoTo 0

            Exit Sub

        End If

        Dim tmp_Element_number As Long

        Dim i_EID As Integer = 1

        For i As Integer = 1 To N_lines_for_EIDs

            For j As Integer = 1 To 8

                tmp_Element_number = Element_number(i_EID).Number

                If tmp_Element_number > 0 Then

                    Fields_array(10 * i + j) = my_Str(Element_number(i_EID))

                    i_EID += 1

                End If

            Next j

        Next i

        On Error GoTo 0

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
Public Class TEMPP1

    Inherits Model_item

    Public SID As Long

    Public Element_number() As Object

    Public TBAR As Long

    Public TPRIME As Long

    Public Temperature(1) As Double

    Public Overrides Sub Load()

        Dim i_EID As Long

        ReDim Temperature(1)

        On Error Resume Next

        ReDim Preserve Element_number(0)

        SID = Val(Fields_array(1))

        Element_number(0) = Val(Fields_array(2))

        TBAR = Val(Fields_array(3))

        TPRIME = Val(Fields_array(4))

        Temperature(0) = Val(Fields_array(5))

        Temperature(1) = Val(Fields_array(6))

        If Fields_array(12) = "THRU" Then

            Dim Elements_to_add, Number_start, Number_finish As Long

            Number_start = Val(Fields_array(11))

            Number_finish = Val(Fields_array(13))

            Elements_to_add = Number_finish - Number_start + 1

            i_EID = UBound(Element_number) + 1

            ReDim Element_number(UBound(Element_number) + Elements_to_add)

            For i As Long = Number_start To Number_finish

                Element_number(i_EID) = i

                i_EID += 1

            Next i

            If Fields_array(15) = "THRU" Then

                Number_start = Val(Fields_array(14))

                Number_finish = Val(Fields_array(16))

                Elements_to_add = Number_finish - Number_start + 1

                i_EID = UBound(Element_number) + 1

                ReDim Element_number(UBound(Element_number) + Elements_to_add)

                For i As Long = Number_start To Number_finish

                    Element_number(i_EID) = i

                    i_EID += 1

                Next i


            End If

            On Error GoTo 0

            Exit Sub

        End If

        i_EID = UBound(Element_number) + 1

        For i As Integer = 10 To UBound(Fields_array) Step 10

            For j As Integer = 1 To 8

                Element_number(i_EID) = Val(Fields_array(i + j))

                i_EID += 1

                ReDim Preserve Element_number(i_EID)

            Next j

        Next i

        On Error GoTo 0

    End Sub

    Public Overrides Sub Prepare_Save()

        Dim N_lines_for_EIDs As Integer

        N_lines_for_EIDs = UBound(Element_number) / 8

        ReDim Fields_array(10 + N_lines_for_EIDs * 10)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(SID)

        Fields_array(2) = my_Str(Element_number(0))

        Fields_array(3) = my_Str(TBAR)

        Fields_array(4) = my_Str(TPRIME)

        Fields_array(5) = my_Str(Temperature(0))

        Fields_array(6) = my_Str(Temperature(1))

        Dim tmp_Element_number As Long

        Dim i_EID As Integer = 1

        For i As Integer = 1 To N_lines_for_EIDs

            For j As Integer = 1 To 8

                tmp_Element_number = Element_number(i_EID).Number

                If tmp_Element_number > 0 Then

                    Fields_array(10 * i + j) = my_Str(Element_number(i_EID))

                    i_EID += 1

                End If

            Next j

        Next i

        On Error GoTo 0

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
Public Class TEMPRB

    Inherits Model_item

    Public SID As Long

    Public Element_number() As Object

    Public TA As Double

    Public TB As Double

    Public TP1A As Double

    Public TP1B As Double

    Public TP2A As Double

    Public TP2B As Double


    Public TCA As Double

    Public TDA As Double

    Public TEA As Double

    Public TFA As Double

    Public TCB As Double

    Public TDB As Double

    Public TEB As Double

    Public TFB As Double

    Public Overrides Sub Load()

        Dim i_EID As Long

        On Error Resume Next

        ReDim Preserve Element_number(0)

        SID = Val(Fields_array(1))

        Element_number(0) = Val(Fields_array(2))

        TA = Val(Fields_array(3))

        TB = Val(Fields_array(4))

        TP1A = Val(Fields_array(5))

        TP1B = Val(Fields_array(6))

        TP2A = Val(Fields_array(7))

        TP2B = Val(Fields_array(8))


        TCA = Val(Fields_array(11))

        TDA = Val(Fields_array(12))

        TEA = Val(Fields_array(13))

        TFA = Val(Fields_array(14))

        TCB = Val(Fields_array(15))

        TDB = Val(Fields_array(16))

        TEB = Val(Fields_array(17))

        TFB = Val(Fields_array(18))

        If Fields_array(22) = "THRU" Then

            Dim Elements_to_add, Number_start, Number_finish As Long

            Number_start = Val(Fields_array(21))

            Number_finish = Val(Fields_array(23))

            Elements_to_add = Number_finish - Number_start + 1

            i_EID = UBound(Element_number) + 1

            ReDim Element_number(UBound(Element_number) + Elements_to_add)

            For i As Long = Number_start To Number_finish

                Element_number(i_EID) = i

                i_EID += 1

            Next i

            If Fields_array(25) = "THRU" Then

                Number_start = Val(Fields_array(24))

                Number_finish = Val(Fields_array(26))

                Elements_to_add = Number_finish - Number_start + 1

                i_EID = UBound(Element_number) + 1

                ReDim Element_number(UBound(Element_number) + Elements_to_add)

                For i As Long = Number_start To Number_finish

                    Element_number(i_EID) = i

                    i_EID += 1

                Next i


            End If

            On Error GoTo 0

            Exit Sub

        End If

        i_EID = UBound(Element_number) + 1

        For i As Integer = 20 To UBound(Fields_array) Step 10

            For j As Integer = 1 To 8

                Element_number(i_EID) = Val(Fields_array(i + j))

                i_EID += 1

                ReDim Preserve Element_number(i_EID)

            Next j

        Next i



        On Error GoTo 0

    End Sub

    Public Overrides Sub Prepare_Save()

        Dim N_lines_for_EIDs As Integer

        N_lines_for_EIDs = UBound(Element_number) / 8

        ReDim Fields_array(10 + N_lines_for_EIDs * 10)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(SID)

        Fields_array(2) = my_Str(Element_number(0))

        Fields_array(3) = my_Str(TA)

        Fields_array(4) = my_Str(TB)

        Fields_array(5) = my_Str(TP1A)

        Fields_array(6) = my_Str(TP1B)

        Fields_array(7) = my_Str(TP2A)

        Fields_array(8) = my_Str(TP2B)


        Fields_array(11) = my_Str(TCA)

        Fields_array(12) = my_Str(TDA)

        Fields_array(13) = my_Str(TEA)

        Fields_array(14) = my_Str(TFA)

        Fields_array(15) = my_Str(TCB)

        Fields_array(16) = my_Str(TDB)

        Fields_array(17) = my_Str(TEB)

        Fields_array(18) = my_Str(TFB)

        Dim tmp_Element_number As Long

        Dim i_EID As Integer = 1

        For i As Integer = 1 To N_lines_for_EIDs

            For j As Integer = 1 To 8

                tmp_Element_number = Element_number(i_EID).Number

                If tmp_Element_number > 0 Then

                    Fields_array(10 * i + j) = my_Str(Element_number(i_EID))

                    i_EID += 1

                End If

            Next j

        Next i

        On Error GoTo 0

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
Public Class SPC

    Inherits Model_item

    Public SID As Long

    Public Grid() As Object

    ''' <summary>
    ''' Defines constrained degrees of freedom, i.e. 123456 - fully constrained grid
    ''' </summary>
    ''' <remarks></remarks>
    Public Component() As Long

    Public Enforced_motion() As Long


    Public Overrides Sub Load()

        On Error Resume Next

        ReDim Preserve Grid(0), Component(0), Enforced_motion(0)

        SID = Val(Fields_array(1))

        Grid(0) = Val(Fields_array(2))

        Component(0) = Val(Fields_array(3))

        Enforced_motion(0) = Val(Fields_array(4))

        If Fields_array(5) <> "" Then

            ReDim Preserve Grid(1), Component(1), Enforced_motion(1)

            Grid(1) = Val(Fields_array(5))

            Component(1) = Val(Fields_array(6))

            Enforced_motion(1) = Val(Fields_array(7))

        End If

        On Error GoTo 0

    End Sub

    Public Overrides Sub Prepare_Save()

        ReDim Fields_array(7)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(SID)

        Fields_array(2) = my_Str(Grid(0))

        Fields_array(3) = my_Str(Component(0))

        Fields_array(4) = my_Str(Enforced_motion(0))

        If UBound(Grid) = 1 Then

            Fields_array(5) = my_Str(Grid(1))

            Fields_array(6) = my_Str(Component(1))

            Fields_array(7) = my_Str(Enforced_motion(1))

        End If

        On Error GoTo 0

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
Public Class SPC1

    Inherits Model_item

    Public SID As Long

    Public Component As Long

    Public Grid() As Object

    Public Overrides Sub Load()

        On Error Resume Next

        ReDim Preserve Grid(0)

        SID = Val(Fields_array(1))

        Component = Val(Fields_array(2))

        If Fields_array(4) = "THRU" Then

            Dim Elements_to_add, Number_start, Number_finish, i_Grid As Long

            Number_start = Val(Fields_array(3))

            Number_finish = Val(Fields_array(5))

            Elements_to_add = Number_finish - Number_start

            i_Grid = UBound(Grid)

            ReDim Grid(UBound(Grid) + Elements_to_add)

            For i As Long = Number_start To Number_finish

                Grid(i_Grid) = i

                i_Grid += 1

            Next i

            On Error GoTo 0

            Exit Sub

        End If

        For i As Long = 3 To 8

            If Fields_array(i) <> "" Then

                Grid(UBound(Grid)) = Val(Fields_array(i))

                ReDim Preserve Grid(UBound(Grid) + 1)

            End If

        Next i

        ReDim Preserve Grid(UBound(Grid) - 1)

        For i As Integer = 1 To UBound(Fields_array) \ 10

            For j As Integer = 10 * i + 1 To 10 * i + 8

                If Fields_array(j) <> "" Then

                    ReDim Preserve Grid(UBound(Grid) + 1)

                    Grid(UBound(Grid)) = Val(Fields_array(j))

                End If

            Next j

        Next i

        On Error GoTo 0

    End Sub

    Public Overrides Sub Prepare_Save()

        Dim N_lines_for_Grid As Integer

        N_lines_for_Grid = (UBound(Grid) - 6) / 8

        ReDim Fields_array(10 + N_lines_for_Grid * 10)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(SID)

        Fields_array(2) = my_Str(Component)

        Dim i_Grid As Long = 0

        Dim tmp_Grid_number As Long

        For i_Grid = 0 To 5

            tmp_Grid_number = Grid(i_Grid).Number

            If tmp_Grid_number > 0 Then

                Fields_array(i_Grid + 3) = my_Str(tmp_Grid_number)

            Else

                On Error GoTo 0

                Exit Sub

            End If

        Next i_Grid

        For i As Integer = 1 To N_lines_for_Grid

            For j As Integer = 1 To 8

                tmp_Grid_number = Grid(i_Grid).Number

                If tmp_Grid_number > 0 Then

                    Fields_array(10 * i + j) = my_Str(Grid(i_Grid))

                    i_Grid += 1

                End If

            Next j

        Next i

        On Error GoTo 0

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
Public Class SPCADD

    Inherits Model_item

    Public SID As Long

    Public SID_for_adding() As Long

    Public Overrides Sub Load()

        On Error Resume Next

        ReDim Preserve SID_for_adding(0)

        SID = Val(Fields_array(1))

        For i As Long = 2 To 8

            If Fields_array(i) <> "" Then

                SID_for_adding(UBound(SID_for_adding)) = Val(Fields_array(i))

                ReDim Preserve SID_for_adding(UBound(SID_for_adding) + 1)

            End If

        Next i

        ReDim Preserve SID_for_adding(UBound(SID_for_adding) - 1)

        For i As Integer = 1 To UBound(Fields_array) \ 10

            For j As Integer = 10 * i + 1 To 10 * i + 8

                If Fields_array(j) <> "" Then

                    ReDim Preserve SID_for_adding(UBound(SID_for_adding) + 1)

                    SID_for_adding(UBound(SID_for_adding)) = Val(Fields_array(j))

                End If

            Next j

        Next i

        On Error GoTo 0

    End Sub

    Public Overrides Sub Prepare_Save()

        Dim N_lines_for_SIDs As Integer

        N_lines_for_SIDs = (UBound(SID_for_adding) - 7) / 8

        ReDim Fields_array(10 + N_lines_for_SIDs * 10)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(SID)

        Dim i_SID As Long = 0

        Dim tmp_SID_number As Long

        For i_SID = 0 To 6

            tmp_SID_number = SID_for_adding(i_SID)

            If tmp_SID_number > 0 Then

                Fields_array(i_SID + 2) = my_Str(tmp_SID_number)

            Else

                On Error GoTo 0

                Exit Sub

            End If

        Next i_SID

        For i As Integer = 1 To N_lines_for_SIDs

            For j As Integer = 1 To 8

                tmp_SID_number = SID_for_adding(i_SID)

                If tmp_SID_number > 0 Then

                    Fields_array(10 * i + j) = my_Str(SID_for_adding(i_SID))

                    i_SID += 1

                End If

            Next j

        Next i

        On Error GoTo 0

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
Public Class SPCD

    Inherits Model_item

    Public SID As Long

    Public Grid() As Object

    ''' <summary>
    ''' Defines degrees of freedom with enforced displacement
    ''' </summary>
    ''' <remarks></remarks>
    Public Component() As Long

    Public Enforced_motion() As Long


    Public Overrides Sub Load()

        On Error Resume Next

        ReDim Preserve Grid(0), Component(0), Enforced_motion(0)

        SID = Val(Fields_array(1))

        Grid(0) = Val(Fields_array(2))

        Component(0) = Val(Fields_array(3))

        Enforced_motion(0) = Val(Fields_array(4))

        If Fields_array(5) <> "" Then

            ReDim Preserve Grid(1), Component(1), Enforced_motion(1)

            Grid(1) = Val(Fields_array(5))

            Component(1) = Val(Fields_array(6))

            Enforced_motion(1) = Val(Fields_array(7))

        End If

        On Error GoTo 0

    End Sub

    Public Overrides Sub Prepare_Save()

        ReDim Fields_array(7)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(SID)

        Fields_array(2) = my_Str(Grid(0))

        Fields_array(3) = my_Str(Component(0))

        Fields_array(4) = my_Str(Enforced_motion(0))

        If UBound(Grid) = 1 Then

            Fields_array(5) = my_Str(Grid(1))

            Fields_array(6) = my_Str(Component(1))

            Fields_array(7) = my_Str(Enforced_motion(1))

        End If

        On Error GoTo 0

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
Public Class SPCOFF

    Inherits Model_item

    Public Grid() As Object

    ''' <summary>
    ''' Defines degrees of freedom with enforced displacement
    ''' </summary>
    ''' <remarks></remarks>
    Public Component() As Long

    Public Overrides Sub Load()

        On Error Resume Next

        ReDim Preserve Grid(0), Component(0)

        Grid(0) = Val(Fields_array(1))

        Component(0) = Val(Fields_array(2))

        For i As Long = 3 To 7 Step 2

            If Fields_array(i) <> "" Then

                ReDim Preserve Grid(UBound(Grid) + 1), Component(UBound(Component) + 1)

                Grid(UBound(Grid)) = Val(Fields_array(i))

                Component(UBound(Component)) = Val(Fields_array(i + 1))

            End If

        Next i

        On Error GoTo 0

    End Sub

    Public Overrides Sub Prepare_Save()

        ReDim Fields_array(8)

        On Error Resume Next

        Fields_array(0) = Type

        Dim tmp_Grid_number As Long

        For i As Long = 0 To UBound(Grid)

            tmp_Grid_number = Grid(i).Number

            If tmp_Grid_number > 0 Then

                Fields_array(i * 2 + 1) = my_Str(tmp_Grid_number)

                Fields_array(i * 2 + 2) = my_Str(Component(i))

            End If


        Next i

        On Error GoTo 0

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
Public Class SPCOFF1

    Inherits Model_item

    ''' <summary>
    ''' Defines degrees of freedom with enforced displacement
    ''' </summary>
    ''' <remarks></remarks>
    Public Component As Long

    Public Grid() As Object

    Public Overrides Sub Load()

        On Error Resume Next

        ReDim Preserve Grid(0)

        Component = Val(Fields_array(1))

        For i As Long = 2 To 8

            If Fields_array(i) <> "" Then

                Grid(UBound(Grid)) = Val(Fields_array(i))

                ReDim Preserve Grid(UBound(Grid) + 1)

            End If

        Next i

        ReDim Preserve Grid(UBound(Grid) - 1)

        For i As Long = 10 To UBound(Fields_array) Step 10

            For j As Long = 0 To 8

                If Fields_array(i + j + 1) <> "" Then

                    ReDim Preserve Grid(UBound(Grid) + 1)

                    Grid(UBound(Grid)) = Val(Fields_array(i + j + 1))

                End If

            Next j

        Next i

        On Error GoTo 0

    End Sub

    Public Overrides Sub Prepare_Save()

        Dim N_lines_for_Grid As Integer = 0

        Dim i_Grid As Integer = 0

        N_lines_for_Grid = (UBound(Grid) - 7) \ 8 + 1

        ReDim Fields_array(10 + N_lines_for_Grid * 10)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(Component)

        Dim tmp_Grid_number As Long

        For i_Grid = 0 To 6

            tmp_Grid_number = Grid(i_Grid).Number

            If tmp_Grid_number > 0 Then

                Fields_array(3 + i_Grid) = my_Str(tmp_Grid_number)

            End If

        Next i_Grid


        For i As Long = 1 To N_lines_for_Grid

            For j As Long = 0 To 8

                tmp_Grid_number = Grid(i_Grid).Number

                If tmp_Grid_number > 0 Then

                    Fields_array(10 * i + j + 1) = my_Str(tmp_Grid_number)

                    i_Grid += 1

                End If

            Next j

        Next i

        On Error GoTo 0

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
Public Class MPC

    Inherits Model_item

    Public SID As Long

    Public Grid() As Object

    ''' <summary>
    ''' Defines constrained degrees of freedom, i.e. 123456 - fully constrained grid
    ''' </summary>
    ''' <remarks></remarks>
    Public Component() As Long

    Public Coefficient() As Long


    Public Overrides Sub Load()

        On Error Resume Next

        ReDim Preserve Grid(0), Component(0), Coefficient(0)

        SID = Val(Fields_array(1))

        Grid(0) = Val(Fields_array(2))

        Component(0) = Val(Fields_array(3))

        Coefficient(0) = Val(Fields_array(4))

        If Fields_array(5) <> "" Then

            ReDim Preserve Grid(1), Component(1), Coefficient(1)

            Grid(1) = Val(Fields_array(5))

            Component(1) = Val(Fields_array(6))

            Coefficient(1) = Val(Fields_array(7))

        End If

        For i As Long = 1 To UBound(Fields_array) Step 10

            For j As Long = 2 To 5 Step 3

                If Fields_array(10 * i + j) = "" Then

                    ReDim Preserve Grid(UBound(Grid) + 1), Component(UBound(Component) + 1), Coefficient(UBound(Coefficient) + 1)

                    Grid(UBound(Grid)) = Val(Fields_array(10 * i + j))

                    Component(UBound(Component)) = Val(Fields_array(10 * i + j + 1))

                    Coefficient(UBound(Coefficient)) = Val(Fields_array(10 * i + j + 2))

                End If

            Next j

        Next i

        On Error GoTo 0

    End Sub

    Public Overrides Sub Prepare_Save()

        Dim N_lines_for_Grid As Long

        N_lines_for_Grid = (UBound(Grid) - 2) / 2

        ReDim Fields_array(10 + 10 * N_lines_for_Grid)

        On Error Resume Next

        Fields_array(0) = Type

        Dim tmp_Grid_number As Long

        For i As Long = 0 To 1

            tmp_Grid_number = Grid(i).Number

            If tmp_Grid_number > 0 Then

                Fields_array(i * 3 + 2) = my_Str(tmp_Grid_number)

                Fields_array(i * 3 + 3) = my_Str(Component(i))

                Fields_array(i * 3 + 4) = my_Str(Coefficient(i))

            End If

        Next i

        Dim i_Grid As Long = 2

        For i As Long = 1 To N_lines_for_Grid

            For j As Long = 2 To 5 Step 3

                tmp_Grid_number = Grid(i_Grid).Number

                If tmp_Grid_number > 0 Then

                    Fields_array(i * 10 + j) = my_Str(tmp_Grid_number)

                    Fields_array(i * 10 + j + 1) = my_Str(Component(i_Grid))

                    Fields_array(i * 10 + j + 2) = my_Str(Coefficient(i_Grid))

                    i_Grid += 1

                End If

            Next j

        Next i

        On Error GoTo 0

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
Public Class MPCADD

    Inherits Model_item

    Public SID As Long

    Public SID_for_adding() As Long

    Public Overrides Sub Load()

        On Error Resume Next

        ReDim Preserve SID_for_adding(0)

        SID = Val(Fields_array(1))

        For i As Long = 2 To 8

            If Fields_array(i) <> "" Then

                SID_for_adding(UBound(SID_for_adding)) = Val(Fields_array(i))

                ReDim Preserve SID_for_adding(UBound(SID_for_adding) + 1)

            End If

        Next i

        ReDim Preserve SID_for_adding(UBound(SID_for_adding) - 1)

        For i As Integer = 1 To UBound(Fields_array) \ 10

            For j As Integer = 10 * i + 1 To 10 * i + 8

                If Fields_array(j) <> "" Then

                    ReDim Preserve SID_for_adding(UBound(SID_for_adding) + 1)

                    SID_for_adding(UBound(SID_for_adding)) = Val(Fields_array(j))

                End If

            Next j

        Next i

        On Error GoTo 0

    End Sub

    Public Overrides Sub Prepare_Save()

        Dim N_lines_for_SIDs As Integer

        N_lines_for_SIDs = (UBound(SID_for_adding) - 7) / 8

        ReDim Fields_array(10 + N_lines_for_SIDs * 10)

        On Error Resume Next

        Fields_array(0) = Type

        Fields_array(1) = my_Str(SID)

        Dim i_SID As Long = 0

        Dim tmp_SID_number As Long

        For i_SID = 0 To 6

            tmp_SID_number = SID_for_adding(i_SID)

            If tmp_SID_number > 0 Then

                Fields_array(i_SID + 2) = my_Str(tmp_SID_number)

            Else

                On Error GoTo 0

                Exit Sub

            End If

        Next i_SID

        For i As Integer = 1 To N_lines_for_SIDs

            For j As Integer = 1 To 8

                tmp_SID_number = SID_for_adding(i_SID)

                If tmp_SID_number > 0 Then

                    Fields_array(10 * i + j) = my_Str(SID_for_adding(i_SID))

                    i_SID += 1

                End If

            Next j

        Next i

        On Error GoTo 0

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
Public Class Model

    Inherits Linear_algebra


    Public File_name As String

    Public File_path As String




    Public NASTRAN_Statement() As Object

    Public File_Management_Statements() As Object

    Public Executive_Control_Statements() As Object

    Public Case_control_Commands() As Object

    Public Grid() As Object

    Public Coordinate_system() As Object

    Public Material() As Object

    Public Element_property() As Object

    Public Element() As Object

    Public Load() As Object

    Public Boundary_condition() As Object


    ' Количество элементов в соотвестующих массивах 
    Public i_BC As Integer = -1

    Public i_CS As Integer = -1

    Public i_E As Integer = -1

    Public i_EP As Integer = -1

    Public i_G As Integer = -1

    Public i_L As Integer = -1

    Public i_M As Integer = -1

    ''' <summary>
    ''' The string, which represent model in file.
    ''' </summary>
    ''' <remarks></remarks>
    Public String_in_file As String

    ' Грани для различных целей

    Public Face() As Object

    Public Shining_face() As Face

    Public Polygon() As Visualisation.Polygon
    ' Термооптические переменные

    'инициализируем массивы для храния объектов

    Public Dependency() As DEPEND

    Public FTOP() As FTOP

    Public MEDIUM() As Object

    Public RSOURCE() As RSOURCE1

    Public HSOURCE() As HSOURCE1

    Public T_INIT() As T_INIT

    Public T_SET() As T_SET


    Dim i_D As Long = -1

    Dim i_F As Long = -1

    Dim i_MEDIUM As Long = -1

    Dim i_RSOURCE As Long = -1

    Dim i_HSOURCE As Long = -1

    Dim i_T_INIT As Long = -1

    Dim i_T_SET As Long = -1

    ''' <summary>
    ''' Uniform initial temperature field over all of model
    ''' </summary>
    ''' <remarks></remarks>
    Public T_ALL As Double

    Public Global_medium As Object

    ''' <summary>
    ''' The type of solver, used for problem solution
    ''' </summary>
    ''' <remarks></remarks>
    Public Solver_type As String

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
    ''' If dependency is used in radiation part of the solution, then Is_thermooptics_constant = false
    ''' </summary>
    ''' <remarks></remarks>
    Public Is_thermooptics_constant As Boolean


    ''' <summary>
    ''' If dependency is used in conductive part of the solution, then Is_thermophysics_constant = false
    ''' </summary>
    ''' <remarks></remarks>
    Public Is_thermophysics_constant As Boolean

    ''' <summary>
    ''' If dependency is used for boundary conditions definitions, then Is_boundary_conditions_constant = false
    ''' </summary>
    ''' <remarks></remarks>
    Public Is_boundary_conditions_constant As Boolean


    ''' <summary>
    ''' Is radiation heat transfer in model? Is any unzero elemet in radiation heat transfer matrix?
    ''' </summary>
    ''' <remarks></remarks>
    Public Is_radiation_heat_transfer As Boolean

    ''' <summary>
    ''' Is radiation heat sources are fixed, not movable? 
    ''' </summary>
    ''' <remarks></remarks>
    Public Is_radiation_heat_sources_are_fixed As Boolean


    '''' <summary>
    '''' Global maximum error, that can occur in any thermal-optical calculations
    '''' </summary>
    '''' <remarks></remarks>
    'Public TOP_Epsilon As Double


    ''' <summary>
    ''' Number of photons, emitted from face at time of one solution step
    ''' </summary>
    ''' <remarks></remarks>
    Public N_photon As Integer

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
    ''' The time step for saving data 
    ''' </summary>
    ''' <remarks></remarks>
    Public Save_step_time As Double



    Public State As String

    Public Maximum_length As Double

    Public Minimum_length As Double

    Public Center As Vector


    Public Enabled As Boolean


    Public Event Model_state_changed(ByRef Model As Model)

    Public Event Model_neigbours_changed(ByRef Model As Model)



    Public Model_neighbours_finding_state As Double

    'Функция timeGetTime() измеряет промежуток времени с момента запуска Windows
    Public Declare Function timeGetTime Lib "winmm.dll" () As Long


    ' Функция извлекает из длинной строки кусок, начинающийся с заглавной буквы и кончающийся CHR(13) перед заглавной буквой
    Function Extract_property_string(ByRef Long_string As String, ByVal Position As Long) As String

        Dim old_Position As Integer = 0

        Dim Property_string As String

        Extract_property_string = ""

        Do
            old_Position = Position

            Position = InStr(Position, Long_string, Chr(13)) + 1

            If Position = 1 Then Exit Do

            Property_string = Mid(Long_string, old_Position, Position - old_Position)

            If Asc(Mid(Property_string, 1, 1)) >= 65 And Asc(Mid(Property_string, 1, 1)) <= 90 Or (Mid(Property_string, 1, 1) = "$") Then

                Exit Do

            Else

                Extract_property_string = Extract_property_string & Property_string

                '    If Not (Mid(Property_string, 1, 1) = "$") Then

                '        Extract_property_string = Extract_property_string & Property_string

                '    End If

            End If

            Position = Position + 1

        Loop

    End Function

    '    Public Sub New_1(Optional ByVal File_Path_File_name As String = "")

    '        If File_Path_File_name <> "" Then

    '            Dim Whole_dat As String

    '            Dim Point_position, Slash_position As Integer

    '            'определяем положение последней точки
    '            Point_position = InStrRev(File_Path_File_name, ".")
    '            'определяем положение последего слэша
    '            Slash_position = InStrRev(File_Path_File_name, "\")
    '            ' имя файла
    '            File_name = Mid(File_Path_File_name, Slash_position + 1, Point_position - (Slash_position + 1))
    '            ' путь к файлу
    '            File_path = Mid(File_Path_File_name, 1, Slash_position)

    '            Const ForReading = 1, ForWriting = 2, ForAppending = 3
    '            Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
    '            Dim fs, F, s, ts
    '            fs = CreateObject("Scripting.FileSystemObject")

    '            If fs.fileExists(File_Path_File_name) Then

    '                F = fs.GetFile(File_Path_File_name)

    '            Else

    '                MsgBox("No such file " & File_Path_File_name)

    '                Exit Sub

    '            End If

    '            ts = F.OpenAsTextStream(ForReading, TristateFalse)

    '            ' считываем один входной файл для Nastran'a. Учитываем, что его расширение может быть как dat, так и bdf.
    '            Whole_dat = ts.ReadAll

    '            ' теперь в цикле проходим весь dat файл и считываем данные

    '            Dim Position As Long = 1

    '            Dim old_Position As Long = 1

    '            Dim Property_string As String

    '            Dim tmp_Array As String()

    '            Dim First_letter, Two_first_letters, Three_first_letters, Four_first_letters As String

    '            Dim Normalized_type As String

    '            Do
    '                old_Position = Position

    '                Position = InStr(Position, Whole_dat, Chr(13)) + 1

    '                If Position = 1 Then Exit Do

    '                Property_string = Mid(Whole_dat, old_Position, Position - old_Position)

    '                Position = Position + 1

    '                tmp_Array = Split(Property_string)

    '                Property_string = Property_string & Extract_property_string(Whole_dat, Position)


    '                First_letter = Mid(tmp_Array(0), 1, 1)

    '                Two_first_letters = Mid(tmp_Array(0), 1, 2)

    '                Three_first_letters = Mid(tmp_Array(0), 1, 3)

    '                Four_first_letters = Mid(tmp_Array(0), 1, 4)

    '                Normalized_type = Replace(tmp_Array(0), "*", "")

    '                If First_letter = "A" Then

    '                    Select Case Normalized_type

    '                        Case "ACCEL"

    '                            Dim Item As New ACCEL(Property_string)

    '                            Load.Add(Item)

    '                        Case "ACCEL1"

    '                            Dim Item As New ACCEL1(Property_string)

    '                            Load.Add(Item)


    '                    End Select


    '                    GoTo Skip_checks

    '                End If

    '                If First_letter = "C" Then


    '                    If Four_first_letters = "CORD" Then

    '                        Select Case Normalized_type

    '                            Case "CORD1C"

    '                                Dim Item As New CORD1C(Property_string)

    '                                Me.Coordinate_system.Add(Item, Item.Number)

    '                            Case "CORD2C"

    '                                Dim Item As New CORD2C(Property_string)

    '                                Me.Coordinate_system.Add(Item, Item.Number)

    '                            Case "CORD1R"

    '                                Dim Item As New CORD1R(Property_string)

    '                                Me.Coordinate_system.Add(Item, Item.Number)

    '                            Case "CORD2R"

    '                                Dim Item As New CORD2R(Property_string)

    '                                Me.Coordinate_system.Add(Item, Item.Number)

    '                            Case "CORD1S"

    '                                Dim Item As New CORD1S(Property_string)

    '                                Me.Coordinate_system.Add(Item, Item.Number)

    '                            Case "CORD2S"

    '                                Dim Item As New CORD2S(Property_string)

    '                                Me.Coordinate_system.Add(Item, Item.Number)

    '                        End Select


    '                        GoTo Skip_checks

    '                    End If

    '                    If Two_first_letters = "CB" Then

    '                        Select Case Normalized_type

    '                            Case "CBAR"

    '                                Dim Item As New CBAR(Property_string)

    '                                Me.Element.Add(Item, Item.Number)

    '                            Case "CBEAM"

    '                                Dim Item As New CBEAM(Property_string)

    '                                Me.Element.Add(Item, Item.Number)


    '                        End Select

    '                        GoTo Skip_checks

    '                    End If


    '                    Select Case Normalized_type

    '                        Case "CHEXA"

    '                            Dim Item As New CHEXA(Property_string)

    '                            Me.Element.Add(Item, Item.Number)

    '                        Case "CPENTA"

    '                            Dim Item As New CPENTA(Property_string)

    '                            Me.Element.Add(Item, Item.Number)

    '                        Case "CQUAD4"

    '                            Dim Item As New CQUAD4(Property_string)

    '                            Me.Element.Add(Item, Item.Number)

    '                        Case "CQUAD8"

    '                            Dim Item As New CQUAD8(Property_string)

    '                            Me.Element.Add(Item, Item.Number)

    '                        Case "CROD"

    '                            Dim Item As New CROD(Property_string)

    '                            Me.Element.Add(Item, Item.Number)

    '                        Case "CSHEAR"

    '                            Dim Item As New CSHEAR(Property_string)

    '                            Me.Element.Add(Item, Item.Number)


    '                    End Select

    '                    If Two_first_letters = "CT" Then

    '                        Select Case Normalized_type

    '                            Case "CTETRA"

    '                                Dim Item As New CTETRA(Property_string)

    '                                Me.Element.Add(Item, Item.Number)

    '                            Case "CTRIA3"

    '                                Dim Item As New CTRIA3(Property_string)

    '                                Me.Element.Add(Item, Item.Number)

    '                            Case "CTRIA6"

    '                                Dim Item As New CTRIA6(Property_string)

    '                                Me.Element.Add(Item, Item.Number)


    '                        End Select

    '                        GoTo Skip_checks

    '                    End If

    '                    GoTo Skip_checks

    '                End If



    '                If First_letter = "F" Then

    '                    Select Case Normalized_type

    '                        Case "FORCE"

    '                            Dim Item As New FORCE(Property_string)

    '                            Me.Load.Add(Item)

    '                        Case "FORCE1"

    '                            Dim Item As New FORCE1(Property_string)

    '                            Me.Load.Add(Item)

    '                        Case "FORCE2"

    '                            Dim Item As New FORCE2(Property_string)

    '                            Me.Load.Add(Item)


    '                    End Select

    '                    GoTo Skip_checks

    '                End If

    '                If First_letter = "G" Then

    '                    Select Case Normalized_type

    '                        Case "GRID"

    '                            Dim Item As New Grid(Property_string)

    '                            Me.Grid.Add(Item, Item.Number)

    '                        Case "GRAV"

    '                            Dim Item As New GRAV(Property_string)

    '                            Me.Load.Add(Item, Item.SID)



    '                    End Select

    '                    GoTo Skip_checks


    '                End If

    '                If First_letter = "L" Then

    '                    Select Case Normalized_type

    '                        Case "LOAD"

    '                            Dim Item As New LOAD(Property_string)

    '                            Me.Load.Add(Item)


    '                    End Select

    '                    GoTo Skip_checks


    '                End If

    '                If First_letter = "M" Then

    '                    Select Case Normalized_type

    '                        Case "MAT1"

    '                            Dim Item As New MAT1(Property_string)

    '                            Me.Material.Add(Item, Item.Number)

    '                        Case "MAT2"

    '                            Dim Item As New MAT2(Property_string)

    '                            Me.Material.Add(Item, Item.Number)

    '                        Case "MAT8"

    '                            Dim Item As New MAT8(Property_string)

    '                            Me.Material.Add(Item, Item.Number)

    '                        Case "MAT9"

    '                            Dim Item As New MAT9(Property_string)

    '                            Me.Material.Add(Item, Item.Number)

    '                        Case "MOMENT"

    '                            Dim Item As New MOMENT(Property_string)

    '                            Me.Load.Add(Item)

    '                        Case "MOMENT1"

    '                            Dim Item As New MOMENT1(Property_string)

    '                            Me.Load.Add(Item)

    '                        Case "MOMENT2"

    '                            Dim Item As New MOMENT2(Property_string)

    '                            Me.Load.Add(Item)

    '                        Case "MPC"

    '                            Dim Item As New MPC(Property_string)

    '                            Me.Boundary_condition.Add(Item)

    '                        Case "MPCADD"

    '                            Dim Item As New MPCADD(Property_string)

    '                            Me.Boundary_condition.Add(Item)


    '                    End Select

    '                    GoTo Skip_checks


    '                End If

    '                If First_letter = "P" Then

    '                    Select Case Normalized_type

    '                        Case "PBAR"

    '                            Dim Item As New PBAR(Property_string)

    '                            Me.Element_property.Add(Item, Item.Number)

    '                        Case "PBEAM"

    '                            Dim Item As New PBEAM(Property_string)

    '                            Me.Element_property.Add(Item, Item.Number)

    '                        Case "PCOMP"

    '                            Dim Item As New PCOMP(Property_string)

    '                            Me.Element_property.Add(Item, Item.Number)


    '                        Case "PROD"

    '                            Dim Item As New PROD(Property_string)

    '                            Me.Element_property.Add(Item, Item.Number)

    '                    End Select


    '                    If Four_first_letters = "PLOA" Then

    '                        Select Case Normalized_type


    '                            Case "PLOAD"

    '                                Dim Item As New PLOAD(Property_string)

    '                                Me.Load.Add(Item)

    '                            Case "PLOAD1"

    '                                Dim Item As New PLOAD1(Property_string)

    '                                Me.Load.Add(Item)

    '                            Case "PLOAD2"

    '                                Dim Item As New PLOAD2(Property_string)

    '                                Me.Load.Add(Item)

    '                            Case "PLOAD4"

    '                                Dim Item As New PLOAD4(Property_string)

    '                                Me.Load.Add(Item)

    '                        End Select

    '                        GoTo Skip_checks

    '                    End If

    '                    If Two_first_letters = "PS" Then

    '                        Select Case Normalized_type

    '                            Case "PSHEAR"

    '                                Dim Item As New PSHEAR(Property_string)

    '                                Me.Element_property.Add(Item, Item.Number)

    '                            Case "PSHELL"

    '                                Dim Item As New PSHELL(Property_string)

    '                                Me.Element_property.Add(Item, Item.Number)

    '                            Case "PSOLID"

    '                                Dim Item As New PSOLID(Property_string)

    '                                Me.Element_property.Add(Item, Item.Number)

    '                        End Select

    '                        GoTo Skip_checks

    '                    End If

    '                    GoTo Skip_checks


    '                End If

    '                If First_letter = "R" Then

    '                    Select Case Normalized_type

    '                        Case "RFORCE"

    '                            Dim Item As New RFORCE(Property_string)

    '                            Me.Load.Add(Item)


    '                    End Select

    '                    GoTo Skip_checks


    '                End If

    '                If First_letter = "S" Then

    '                    Select Case Normalized_type

    '                        Case "SPC"

    '                            Dim Item As New SPC(Property_string)

    '                            Me.Boundary_condition.Add(Item)

    '                        Case "SPC1"

    '                            Dim Item As New SPC1(Property_string)

    '                            Me.Boundary_condition.Add(Item)

    '                        Case "SPCADD"

    '                            Dim Item As New SPCADD(Property_string)

    '                            Me.Boundary_condition.Add(Item)

    '                        Case "SPCD"

    '                            Dim Item As New SPCD(Property_string)

    '                            Me.Boundary_condition.Add(Item)

    '                        Case "SPCOFF"

    '                            Dim Item As New SPCOFF(Property_string)

    '                            Me.Boundary_condition.Add(Item)

    '                        Case "SPCOFF1"

    '                            Dim Item As New SPCOFF1(Property_string)

    '                            Me.Boundary_condition.Add(Item)

    '                    End Select

    '                    GoTo Skip_checks


    '                End If


    '                If First_letter = "T" Then

    '                    Select Case Normalized_type

    '                        Case "TEMP"

    '                            Dim Item As New TEMP(Property_string)

    '                            Me.Load.Add(Item)

    '                        Case "TEMPBC"

    '                            Dim Item As New TEMPBC(Property_string)

    '                            Me.Boundary_condition.Add(Item)

    '                        Case "TEMPD"

    '                            Dim Item As New TEMPD(Property_string)

    '                            Me.Load.Add(Item)

    '                        Case "TEMPF"

    '                            Dim Item As New TEMPF(Property_string)

    '                            Me.Load.Add(Item)

    '                        Case "TEMPP1"

    '                            Dim Item As New TEMPP1(Property_string)

    '                            Me.Load.Add(Item)

    '                        Case "TEMPRB"

    '                            Dim Item As New TEMPRB(Property_string)

    '                            Me.Load.Add(Item)

    '                    End Select

    '                    GoTo Skip_checks

    '                End If

    '                If tmp_Array(0) = "ENDDATA" Then

    '                    Exit Do
    '                End If

    'Skip_checks:

    '            Loop

    '            Create_links()

    '            Create_faces()

    '            '' Процедура считывает и заполняет массивы термо-оптических свойств проекта
    '            '' Принимается, что файл с данными находится по тому же пути, что и файлы проекта, но имеет расширение *.top (Thermal and optical properties)
    '            'Call Read_and_fill_thermal_and_optical_properties(One_project)

    '            '' процедура рассчитывает положение характерных точек для систем координат всех элементов типа CQUAD4
    '            'Call fill_CQUAD4_coord_systems(One_project)

    '            '' процедура рассчитывает положение характерных точек для систем координат всех элементов типа CTRIA3
    '            'Call fill_CTRIA3_coord_systems(One_project)

    '            '' процедура заполняет матрицы перехода из различных систем координат в глобальную и наоборот
    '            'Call fill_Coord_sys_matrixes(One_project)

    '            '' процедура заполняет "упругие" и "допускаемые" поля в массиве композиционных материалов проекта
    '            'Call fill_PCOMP_Special_fields(One_project)

    '            '' процедура добавляет последним элементом массива систем координат глобальную систему координат
    '            'Call Add_coord_sys_0(One_project)

    '            '' Среда, заполняющая пространство в проекте имеет номер 0 в массиве MTOP1
    '            'Call fill_Medium(One_project)

    '            '' процедура переводит все узлы проекта в глобальную систему координат
    '            'Call Convert_Grid_to_Global_coordinates(One_project)

    '            '' Процедура заполняет массив граней проекта
    '            'Call fill_Project_faces(One_project)

    '            '' Процедура вычисляет объемы всех элементов модели (у которых есть объем)
    '            'Call fill_Volumes(One_project)

    '            '' в самый хвост процедуры
    '            'On Error GoTo 0

    '            'Exit Sub

    '        End If

    '    End Sub

    ' Процедура заполняет массив граней проекта
    Sub Create_faces()

        ' процедура заполняет в массиве граней проекта массивы узлов, на которых построены грани
        Fill_Faces_GRID()

        '' процедура заполняет максимальные допустимые ошибки граней тем же значением, что установлено для проекта в целом
        'Fill_Faces_TOP_Epsilon()

        '' процедура вычисляет площади граней
        Calc_Faces_Area()

        Calculate_Face_Center_and_coord_axes()

        Detect_neighbours_2()

        ' дорассчитываем центры и системы координат для кастированных граней
        Calculate_2_Grid_faces_Center_and_coord_axes()

        ' формирует коллекцию граней, которые могут излучать
        Extract_shining_faces()

        '' процедура рассчитывает компоненты цветов видимых граней
        Fill_Shining_faces_colours()




        '' процедура заполняет оптические свойства граней
        'Call fill_Faces_optic(One_project)


    End Sub

    ' процедура заполняет в массиве граней проекта массивы узлов, на которых построены грани
    ' кроме того, процедура устанавливает взаимно-однозначное соотвествие между элементами-предками и гранями-потомками
    Sub Fill_Faces_GRID()

        Dim i_Face As Long = -1

        ReDim Face(Max_items_default_2)

        For i As Integer = 0 To UBound(Element)

            If Element(i).Type = "CHEXA" Then

                ReDim Element(i).Face(5)


                i_Face += 1

                Face(i_Face) = New Face

                ReDim Face(i_Face).Grid(3)

                Face(i_Face).Grid(0) = Element(i).Grid(3)

                Face(i_Face).Grid(1) = Element(i).Grid(2)

                Face(i_Face).Grid(2) = Element(i).Grid(1)

                Face(i_Face).Grid(3) = Element(i).Grid(0)

                Face(i_Face).Element = Element(i)

                Element(i).Face(0) = Face(i_Face)


                i_Face += 1

                Face(i_Face) = New Face

                ReDim Face(i_Face).Grid(3)

                Face(i_Face).Grid(0) = Element(i).Grid(4)

                Face(i_Face).Grid(1) = Element(i).Grid(5)

                Face(i_Face).Grid(2) = Element(i).Grid(6)

                Face(i_Face).Grid(3) = Element(i).Grid(7)

                Face(i_Face).Element = Element(i)

                Element(i).Face(1) = Face(i_Face)


                i_Face += 1

                Face(i_Face) = New Face

                ReDim Face(i_Face).Grid(3)

                Face(i_Face).Grid(0) = Element(i).Grid(0)

                Face(i_Face).Grid(1) = Element(i).Grid(1)

                Face(i_Face).Grid(2) = Element(i).Grid(5)

                Face(i_Face).Grid(3) = Element(i).Grid(4)

                Face(i_Face).Element = Element(i)

                Element(i).Face(2) = Face(i_Face)



                i_Face += 1

                Face(i_Face) = New Face

                ReDim Face(i_Face).Grid(3)

                Face(i_Face).Grid(0) = Element(i).Grid(3)

                Face(i_Face).Grid(1) = Element(i).Grid(7)

                Face(i_Face).Grid(2) = Element(i).Grid(6)

                Face(i_Face).Grid(3) = Element(i).Grid(2)

                Face(i_Face).Element = Element(i)

                Element(i).Face(3) = Face(i_Face)



                i_Face += 1

                Face(i_Face) = New Face

                ReDim Face(i_Face).Grid(3)

                Face(i_Face).Grid(0) = Element(i).Grid(0)

                Face(i_Face).Grid(1) = Element(i).Grid(4)

                Face(i_Face).Grid(2) = Element(i).Grid(7)

                Face(i_Face).Grid(3) = Element(i).Grid(3)

                Face(i_Face).Element = Element(i)

                Element(i).Face(4) = Face(i_Face)



                i_Face += 1

                Face(i_Face) = New Face

                ReDim Face(i_Face).Grid(3)

                Face(i_Face).Grid(0) = Element(i).Grid(1)

                Face(i_Face).Grid(1) = Element(i).Grid(2)

                Face(i_Face).Grid(2) = Element(i).Grid(6)

                Face(i_Face).Grid(3) = Element(i).Grid(5)

                Face(i_Face).Element = Element(i)

                Element(i).Face(5) = Face(i_Face)


                GoTo skip_checks

            End If

            If Element(i).Type = "CPENTA" Then

                ReDim Element(i).Face(4)


                i_Face += 1

                Face(i_Face) = New Face

                ReDim Face(i_Face).Grid(2)

                Face(i_Face).Grid(0) = Element(i).Grid(2)

                Face(i_Face).Grid(1) = Element(i).Grid(1)

                Face(i_Face).Grid(2) = Element(i).Grid(0)

                Face(i_Face).Element = Element(i)

                Element(i).Face(0) = Face(i_Face)


                i_Face += 1

                Face(i_Face) = New Face

                ReDim Face(i_Face).Grid(2)

                Face(i_Face).Grid(0) = Element(i).Grid(3)

                Face(i_Face).Grid(1) = Element(i).Grid(4)

                Face(i_Face).Grid(2) = Element(i).Grid(5)

                Face(i_Face).Element = Element(i)

                Element(i).Face(1) = Face(i_Face)


                i_Face += 1

                Face(i_Face) = New Face

                ReDim Face(i_Face).Grid(3)

                Face(i_Face).Grid(0) = Element(i).Grid(1)

                Face(i_Face).Grid(1) = Element(i).Grid(2)

                Face(i_Face).Grid(2) = Element(i).Grid(5)

                Face(i_Face).Grid(3) = Element(i).Grid(4)

                Face(i_Face).Element = Element(i)

                Element(i).Face(2) = Face(i_Face)


                i_Face += 1

                Face(i_Face) = New Face

                ReDim Face(i_Face).Grid(3)

                Face(i_Face).Grid(0) = Element(i).Grid(0)

                Face(i_Face).Grid(1) = Element(i).Grid(1)

                Face(i_Face).Grid(2) = Element(i).Grid(4)

                Face(i_Face).Grid(3) = Element(i).Grid(3)

                Face(i_Face).Element = Element(i)

                Element(i).Face(3) = Face(i_Face)


                i_Face += 1

                Face(i_Face) = New Face

                ReDim Face(i_Face).Grid(3)

                Face(i_Face).Grid(0) = Element(i).Grid(0)

                Face(i_Face).Grid(1) = Element(i).Grid(3)

                Face(i_Face).Grid(2) = Element(i).Grid(5)

                Face(i_Face).Grid(3) = Element(i).Grid(2)

                Face(i_Face).Element = Element(i)

                Element(i).Face(4) = Face(i_Face)

                GoTo skip_checks

            End If

            If Element(i).Type = "CTETRA" Then

                ReDim Element(i).Face(3)


                i_Face += 1

                Face(i_Face) = New Face

                ReDim Face(i_Face).Grid(2)

                Face(i_Face).Grid(0) = Element(i).Grid(0)

                Face(i_Face).Grid(1) = Element(i).Grid(1)

                Face(i_Face).Grid(2) = Element(i).Grid(3)

                Face(i_Face).Element = Element(i)

                Element(i).Face(0) = Face(i_Face)



                i_Face += 1

                Face(i_Face) = New Face

                ReDim Face(i_Face).Grid(2)

                Face(i_Face).Grid(0) = Element(i).Grid(1)

                Face(i_Face).Grid(1) = Element(i).Grid(2)

                Face(i_Face).Grid(2) = Element(i).Grid(3)

                Face(i_Face).Element = Element(i)

                Element(i).Face(1) = Face(i_Face)



                i_Face += 1

                Face(i_Face) = New Face

                ReDim Face(i_Face).Grid(2)

                Face(i_Face).Grid(0) = Element(i).Grid(2)

                Face(i_Face).Grid(1) = Element(i).Grid(0)

                Face(i_Face).Grid(2) = Element(i).Grid(3)

                Face(i_Face).Element = Element(i)

                Element(i).Face(2) = Face(i_Face)


                i_Face += 1

                Face(i_Face) = New Face

                ReDim Face(i_Face).Grid(2)

                Face(i_Face).Grid(0) = Element(i).Grid(0)

                Face(i_Face).Grid(1) = Element(i).Grid(1)

                Face(i_Face).Grid(2) = Element(i).Grid(2)

                Face(i_Face).Element = Element(i)

                Element(i).Face(3) = Face(i_Face)


                GoTo skip_checks

            End If

            If Element(i).Type = "CQUAD4" Or Element(i).Type = "CQUAD8" Then

                ReDim Element(i).Face(5)


                i_Face += 1

                Face(i_Face) = New Face

                ReDim Face(i_Face).Grid(3)

                Face(i_Face).Grid(0) = Element(i).Grid(0)

                Face(i_Face).Grid(1) = Element(i).Grid(1)

                Face(i_Face).Grid(2) = Element(i).Grid(2)

                Face(i_Face).Grid(3) = Element(i).Grid(3)

                Face(i_Face).Element = Element(i)

                Element(i).Face(0) = Face(i_Face)


                i_Face += 1

                Face(i_Face) = New Face

                ReDim Face(i_Face).Grid(3)

                Face(i_Face).Grid(0) = Element(i).Grid(3)

                Face(i_Face).Grid(1) = Element(i).Grid(2)

                Face(i_Face).Grid(2) = Element(i).Grid(1)

                Face(i_Face).Grid(3) = Element(i).Grid(0)

                Face(i_Face).Element = Element(i)

                Element(i).Face(1) = Face(i_Face)

                ' Additional 2-grid faces

                i_Face += 1

                Face(i_Face) = New Face

                ReDim Face(i_Face).Grid(1)

                Face(i_Face).Grid(0) = Element(i).Grid(0)

                Face(i_Face).Grid(1) = Element(i).Grid(1)

                Face(i_Face).Element = Element(i)

                Element(i).Face(2) = Face(i_Face)



                i_Face += 1

                Face(i_Face) = New Face

                ReDim Face(i_Face).Grid(1)

                Face(i_Face).Grid(0) = Element(i).Grid(1)

                Face(i_Face).Grid(1) = Element(i).Grid(2)

                Face(i_Face).Element = Element(i)

                Element(i).Face(3) = Face(i_Face)


                i_Face += 1

                Face(i_Face) = New Face

                ReDim Face(i_Face).Grid(1)

                Face(i_Face).Grid(0) = Element(i).Grid(2)

                Face(i_Face).Grid(1) = Element(i).Grid(3)

                Face(i_Face).Element = Element(i)

                Element(i).Face(4) = Face(i_Face)


                i_Face += 1

                Face(i_Face) = New Face

                ReDim Face(i_Face).Grid(1)

                Face(i_Face).Grid(0) = Element(i).Grid(3)

                Face(i_Face).Grid(1) = Element(i).Grid(0)

                Face(i_Face).Element = Element(i)

                Element(i).Face(5) = Face(i_Face)


                GoTo skip_checks

            End If


            If Element(i).Type = "CTRIA3" Or Element(i).Type = "CTRIA6" Then

                ReDim Element(i).Face(4)


                i_Face += 1

                Face(i_Face) = New Face

                ReDim Face(i_Face).Grid(2)

                Face(i_Face).Grid(0) = Element(i).Grid(0)

                Face(i_Face).Grid(1) = Element(i).Grid(1)

                Face(i_Face).Grid(2) = Element(i).Grid(2)

                Face(i_Face).Element = Element(i)

                Element(i).Face(0) = Face(i_Face)


                i_Face += 1

                Face(i_Face) = New Face

                ReDim Face(i_Face).Grid(2)

                Face(i_Face).Grid(0) = Element(i).Grid(2)

                Face(i_Face).Grid(1) = Element(i).Grid(1)

                Face(i_Face).Grid(2) = Element(i).Grid(0)

                Face(i_Face).Element = Element(i)

                Element(i).Face(1) = Face(i_Face)


                'Additional 2-grid faces


                i_Face += 1

                Face(i_Face) = New Face

                ReDim Face(i_Face).Grid(1)

                Face(i_Face).Grid(0) = Element(i).Grid(0)

                Face(i_Face).Grid(1) = Element(i).Grid(1)

                Face(i_Face).Element = Element(i)

                Element(i).Face(2) = Face(i_Face)



                i_Face += 1

                Face(i_Face) = New Face

                ReDim Face(i_Face).Grid(1)

                Face(i_Face).Grid(0) = Element(i).Grid(1)

                Face(i_Face).Grid(1) = Element(i).Grid(2)

                Face(i_Face).Element = Element(i)

                Element(i).Face(3) = Face(i_Face)


                i_Face += 1

                Face(i_Face) = New Face

                ReDim Face(i_Face).Grid(1)

                Face(i_Face).Grid(0) = Element(i).Grid(2)

                Face(i_Face).Grid(1) = Element(i).Grid(0)

                Face(i_Face).Element = Element(i)

                Element(i).Face(4) = Face(i_Face)


                GoTo skip_checks

            End If

skip_checks:

        Next i



        ReDim Preserve Face(i_Face)



        For i As Integer = 0 To UBound(Grid)

            ReDim Grid(i).Face(Max_items_default)

            Grid(i).N_Face = -1

        Next i

        For i As Integer = 0 To UBound(Face)

            For j As Integer = 0 To UBound(Face(i).Grid)

                Face(i).Grid(j).N_face += 1

                Face(i).Grid(j).Face(Face(i).Grid(j).N_face) = Face(i)

            Next j

        Next i

        For i As Integer = 0 To UBound(Grid)

            ReDim Preserve Grid(i).Face(Grid(i).N_Face)

        Next i

    End Sub

    '' процедура заполняет максимальные допустимые ошибки граней тем же значением, что установлено для проекта в целом
    'Sub Fill_Faces_TOP_Epsilon()

    '    For i As Integer = 0 To UBound(Face)

    '        Face(i).TOP_Epsilon = Me.TOP_Epsilon

    '    Next i

    'End Sub

    ' процедура вычисляет площади граней
    Sub Calc_Faces_Area()

        Dim Vertex(2) As Vector

        For i As Integer = 0 To UBound(Face)

            ' Если грань двухугольник (боковая сторона пластины)

            If UBound(Face(i).Grid) = 1 Then

                Vertex(0) = Face(i).Grid(1).Position - Face(i).Grid(0).Position

                Face(i).Area = L_vector(Vertex(0)) * Face(i).Element.Element_property.Thickness

            End If

            ' Если грань треугольник

            If UBound(Face(i).Grid) = 2 Then

                For j As Integer = 0 To 2

                    Vertex(j) = Face(i).Grid(j).Position

                Next j

                Face(i).Area = Triangle_Area(Vertex)

            End If

            ' Если грань четырехугольник

            If UBound(Face(i).Grid) = 3 Then

                For j As Integer = 0 To 2

                    Vertex(j) = Face(i).Grid(j).Position

                Next j

                Face(i).Area = Triangle_Area(Vertex)


                Vertex(0) = Face(i).Grid(2).Position

                Vertex(1) = Face(i).Grid(3).Position

                Vertex(2) = Face(i).Grid(0).Position

                Face(i).Area = Face(i).Area + Triangle_Area(Vertex)

            End If

        Next i

    End Sub
    ' Функция вычисляет площадь треугольника, заданного вершинами
    Function Triangle_Area(ByVal Vertex() As Vector) As Double

        ' полупериметр
        Dim p As Double = 0

        ' площадь
        Dim Area As Double = 0

        ' стороны
        Dim Edge(0 To 2) As Vector

        ' длины сторон
        Dim L_edge(0 To 2) As Double

        ' Рассчитываем длины сторон

        Edge(0) = Vertex(1) - Vertex(0)

        Edge(1) = Vertex(2) - Vertex(1)

        Edge(2) = Vertex(0) - Vertex(2)

        For i As Integer = 0 To 2

            L_edge(i) = L_vector(Edge(i))

        Next i

        For i As Integer = 0 To 2

            p = p + L_edge(i)

        Next i

        p = p / 2

        Area = p

        For i As Integer = 0 To 2

            Area = Area * (p - L_edge(i))

        Next i

        Triangle_Area = Math.Sqrt(Area)

    End Function

    '    Sub Detect_neighbours()

    '        Me.State = "Finding neighbors"

    '        RaiseEvent Model_state_changed(Me)

    '        Dim Grid_on_face_1(), Grid_on_face_2() As Object

    '        Dim tmp_Value As Integer

    '        Dim Coincedence As Integer

    '        Dim Cos_zz As Double

    '        ' Точность выполнения операции угол между векторами
    '        Dim Epsilon As Double = 0.001

    '        On Error Resume Next

    '        For i As Integer = 0 To UBound(Element)

    '            ReDim Element(i).Neighbor(6)

    '            Element(i).N_Neighbor = -1

    '        Next i

    '        For i As Integer = 0 To UBound(Element)

    '            Me.Model_neighbours_finding_state = i / UBound(Element)

    '            RaiseEvent Model_neigbours_changed(Me)

    '            tmp_Value = -1

    '            tmp_Value = UBound(Element(i).Face)

    '            If tmp_Value > -1 Then

    '                For j As Integer = 0 To UBound(Element(i).Face)

    '                    Grid_on_face_1 = Element(i).Face(j).Grid

    '                    For k As Integer = 0 To UBound(Grid_on_face_1)

    '                        For l As Integer = 0 To UBound(Grid_on_face_1(k).Face)

    '                            If Not (Grid_on_face_1(k).Face(l).Can_Radiate) Then

    '                                GoTo Skip_face

    '                            End If


    '                            If Grid_on_face_1(k).Face(l).Element.Number = Element(i).Number Then

    '                                GoTo Skip_face

    '                            End If

    '                            'If Grid_on_face_1(k).Face(l).Grid(0) = Grid_on_face_1(0) _
    '                            'And Grid_on_face_1(k).Face(l).Grid(1) = Grid_on_face_1(1) _
    '                            'And Grid_on_face_1(k).Face(l).Grid(2) = Grid_on_face_1(2) Then

    '                            '    GoTo Skip_face

    '                            'End If

    '                            Cos_zz = -1

    '                            Cos_zz = Cos_angle_between_vectors(Element(i).Face(j).oz, Grid_on_face_1(k).Face(l).oz)

    '                            If Math.Abs(Cos_zz + 1) > Epsilon Then

    '                                GoTo Skip_face

    '                            End If

    '                            Grid_on_face_2 = Grid_on_face_1(k).Face(l).Grid

    '                            If UBound(Grid_on_face_2) <> UBound(Grid_on_face_1) Then

    '                                GoTo Skip_face

    '                            End If


    '                            Coincedence = -1

    '                            For i_1 As Integer = 0 To UBound(Grid_on_face_1)

    '                                For i_2 As Integer = 0 To UBound(Grid_on_face_2)

    '                                    If Grid_on_face_1(i_1).Number = Grid_on_face_2(i_2).Number Then

    '                                        Coincedence += 1

    '                                    End If

    '                                Next i_2

    '                            Next i_1


    '                            If Coincedence = UBound(Grid_on_face_1) Then

    '                                Element(i).N_Neighbor += 1

    '                                Element(i).Neighbor(Element(i).N_Neighbor) = New Neighbour

    '                                Element(i).Neighbor(Element(i).N_Neighbor).Element = Grid_on_face_1(k).Face(l).Element

    '                                Element(i).Neighbor(Element(i).N_Neighbor).Face = Element(i).Face(j)

    '                                Element(i).Face(j).Can_Radiate = False


    '                                Grid_on_face_1(k).Face(l).Element.N_Neighbor += 1

    '                                Grid_on_face_1(k).Face(l).Element.Neighbor(Grid_on_face_1(k).Face(l).Element.N_Neighbor) = New Neighbour

    '                                Grid_on_face_1(k).Face(l).Element.Neighbor(Grid_on_face_1(k).Face(l).Element.N_Neighbor).Element = Element(i)

    '                                Grid_on_face_1(k).Face(l).Element.Neighbor(Grid_on_face_1(k).Face(l).Element.N_Neighbor).Face = Grid_on_face_1(k).Face(l)

    '                                Grid_on_face_1(k).Face(l).Can_Radiate = False



    '                                If Element(i).Face(j).Area < Grid_on_face_1(k).Face(l).Area Then

    '                                    Element(i).Neighbor(Element(i).N_Neighbor).Area = Element(i).Face(j).Area

    '                                    Grid_on_face_1(k).Face(l).Element.Neighbor(Grid_on_face_1(k).Face(l).Element.N_Neighbor).Area = Element(i).Face(j).Area

    '                                Else

    '                                    Element(i).Neighbor(Element(i).N_Neighbor).Area = _
    '                                    Grid_on_face_1(k).Face(l).Element.Neighbor(Grid_on_face_1(k).Face(l).Element.N_Neighbor).Face.Area.Area

    '                                    Grid_on_face_1(k).Face(l).Element.Neighbor(Grid_on_face_1(k).Face(l).Element.N_Neighbor).Area = _
    '                                    Grid_on_face_1(k).Face(l).Element.Neighbor(Grid_on_face_1(k).Face(l).Element.N_Neighbor).Face.Area

    '                                End If


    '                            End If

    'Skip_face:
    '                        Next l

    '                    Next k

    '                Next j

    '            End If

    '        Next i


    '        For i As Integer = 0 To UBound(Element)

    '            ReDim Preserve Element(i).Neighbor(Element(i).N_Neighbor)

    '        Next i

    '        On Error GoTo 0
    '    End Sub

    Sub Detect_neighbours_2()

        Me.State = "Finding neighbors"

        RaiseEvent Model_state_changed(Me)

        Dim Grid_on_face_1(), Grid_on_face_2() As Object

        Dim tmp_Value As Integer

        Dim Coincedence As Integer

        Dim Cos_zz As Double

        ' Точность выполнения операции угол между векторами
        Dim Epsilon As Double = 0.001

        On Error Resume Next

        For i As Integer = 0 To UBound(Element)

            ReDim Element(i).Neighbor(99)

            Element(i).N_Neighbor = -1

        Next i

        For i As Integer = 0 To UBound(Grid)

            tmp_Value = -1

            tmp_Value = UBound(Grid(i).Face)

            If tmp_Value <> -1 Then


                For j As Integer = 0 To UBound(Grid(i).Face)

                    If Not (Grid(i).Face(j).Can_Radiate) Then

                        GoTo skip_face_1

                    End If

                    Grid_on_face_1 = Grid(i).Face(j).Grid

                    For k As Integer = 0 To UBound(Grid(i).Face)

                        If k = j Then

                            GoTo skip_face_2

                        End If

                        If Not (Grid(i).Face(k).Can_Radiate) Then

                            GoTo skip_face_2

                        End If

                        If Grid(i).Face(j).Element.Number = Grid(i).Face(k).Element.Number Then

                            GoTo skip_face_2

                        End If

                        If UBound(Grid(i).Face(j).Grid) <> UBound(Grid(i).Face(k).Grid) Then

                            GoTo skip_face_2

                        End If


                        Cos_zz = -1

                        Cos_zz = Cos_angle_between_vectors(Grid(i).Face(j).oz, Grid(i).Face(k).oz)

                        If Math.Abs(Cos_zz + 1) > Epsilon Then

                            GoTo skip_face_2

                        End If


                        Grid_on_face_2 = Grid(i).Face(k).Grid




                        Coincedence = -1

                        For i_1 As Integer = 0 To UBound(Grid_on_face_1)

                            For i_2 As Integer = 0 To UBound(Grid_on_face_2)

                                If Grid_on_face_1(i_1).Number = Grid_on_face_2(i_2).Number Then

                                    Coincedence += 1

                                End If

                            Next i_2

                        Next i_1


                        If Coincedence = UBound(Grid_on_face_1) Then

                            Grid(i).Face(j).Element.N_Neighbor += 1

                            Grid(i).Face(j).Element.Neighbor(Grid(i).Face(j).Element.N_Neighbor) = New Neighbour

                            Grid(i).Face(j).Element.Neighbor(Grid(i).Face(j).Element.N_Neighbor).Element = Grid(i).Face(k).Element

                            Grid(i).Face(j).Element.Neighbor(Grid(i).Face(j).Element.N_Neighbor).Area_1 = Grid(i).Face(j).Area

                            Grid(i).Face(j).Element.Neighbor(Grid(i).Face(j).Element.N_Neighbor).Area_2 = Grid(i).Face(k).Area

                            Grid(i).Face(j).Element.Neighbor(Grid(i).Face(j).Element.N_Neighbor).Face = Grid(i).Face(j)

                            Grid(i).Face(j).Can_Radiate = False

                            Grid(i).Face(j).Neighbour_face = Grid(i).Face(k)


                            Grid(i).Face(k).Element.N_Neighbor += 1

                            Grid(i).Face(k).Element.Neighbor(Grid(i).Face(k).Element.N_Neighbor) = New Neighbour

                            Grid(i).Face(k).Element.Neighbor(Grid(i).Face(k).Element.N_Neighbor).Element = Grid(i).Face(j).Element

                            Grid(i).Face(k).Element.Neighbor(Grid(i).Face(k).Element.N_Neighbor).Area_1 = Grid(i).Face(k).Area

                            Grid(i).Face(k).Element.Neighbor(Grid(i).Face(k).Element.N_Neighbor).Area_2 = Grid(i).Face(j).Area

                            Grid(i).Face(k).Element.Neighbor(Grid(i).Face(k).Element.N_Neighbor).Face = Grid(i).Face(k)

                            'Grid(i).Face(k).Can_Radiate = False

                            Grid(i).Face(k).Neighbour_face = Grid(i).Face(j)



                            'If Grid(i).Face(j).Area < Grid(i).Face(k).Area Then

                            '    Grid(i).Face(j).Element.Neighbor(Grid(i).Face(j).Element.N_Neighbor).Area = Grid(i).Face(j).Area

                            '    Grid(i).Face(k).Element.Neighbor(Grid(i).Face(k).Element.N_Neighbor).Area = Grid(i).Face(j).Area

                            'Else

                            '    Grid(i).Face(j).Element.Neighbor(Grid(i).Face(j).Element.N_Neighbor).Area = Grid(i).Face(k).Area

                            '    Grid(i).Face(k).Element.Neighbor(Grid(i).Face(k).Element.N_Neighbor).Area = Grid(i).Face(k).Area

                            'End If


                        End If
Skip_face_2:


                        Application.DoEvents()

                    Next k
Skip_face_1:

                    Application.DoEvents()

                Next j

            End If

        Next i

        For i As Integer = 0 To UBound(Face)

            If UBound(Face(i).Grid) = 1 Then

                Face(i).Can_Radiate = False

            End If

        Next i


        For i As Integer = 0 To UBound(Element)

            ReDim Preserve Element(i).Neighbor(Element(i).N_Neighbor)

        Next i

        On Error GoTo 0

    End Sub


    ' процедура рассчитывает компоненты цветов видимых граней
    Sub Fill_Shining_faces_colours()

        For i As Integer = 0 To UBound(Shining_face)

            Shining_face(i).Detect_face_number(Shining_face(i).Code, i + 1)

        Next i

    End Sub
    ''' <summary>
    ''' This sub fill Children fields in all of model items where possible
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Create_links()

        'Dim Grid_number, Element_property_string, Element_number, Material_number, Coordinate_system_number As String

        Dim tmp_Value As Long = -1

        Me.State = "Linking model"

        RaiseEvent Model_state_changed(Me)

        On Error Resume Next

        'Dim Boundary_condition_collection As Collection = Convert_array_to_collection(Boundary_condition)

        Dim Coordinate_system_collection As Collection = Convert_array_to_collection(Coordinate_system)

        Dim Element_collection As Collection = Convert_array_to_collection(Element)

        Dim Element_property_collection As Collection = Convert_array_to_collection(Element_property)

        Dim Grid_collection As Collection = Convert_array_to_collection(Grid)

        'Dim Load_collection As Collection = Convert_array_to_collection(Load)

        Dim Material_collection As Collection = Convert_array_to_collection(Material)





        ' Отмечаем родителей элементов
        For i As Integer = 0 To UBound(Element)

            ' отмечаемся в узлах
            For j As Long = 0 To UBound(Element(i).Grid)

                If Element(i).Grid(j) > 0 Then

                    Create_one_link(Element(i).Grid(j), Grid_collection)

                Else

                    ReDim Preserve Element(i).Grid(j - 1)

                    Exit For

                End If

            Next j

            ' отмечаемся в свойствах элемента

            Create_one_link(Element(i).Element_property, Element_property_collection)

        Next i


        For i As Integer = 0 To UBound(Grid)

            ReDim Grid(i).Element(Max_items_default)

            Grid(i).N_Element = -1

        Next i

        Dim tmp_Type As String

        For i As Integer = 0 To UBound(Element)

            For j As Integer = 0 To UBound(Element(i).Grid)

                tmp_Type = ""

                tmp_Type = Element(i).Grid(j).Type

                If tmp_Type = "GRID" Then

                    Element(i).Grid(j).N_Element += 1

                    Element(i).Grid(j).Element(Element(i).Grid(j).N_Element) = Element(i)

                Else

                    'ReDim Preserve Element(i).Grid(j - 1)

                    Exit For

                End If

            Next j

        Next i

        For i As Integer = 0 To UBound(Grid)

            ReDim Preserve Grid(i).Element(Grid(i).N_Element)

        Next i



        '' Отмечаем родителей свойств элементов

        For i As Integer = 0 To UBound(Element_property)

            Select Case Element_property(i).Type

                Case "PSHELL"

                    If Element_property(i).Material_for_bending > 0 Then


                        Create_one_link(Element_property(i).Material_for_bending, Material_collection)

                    End If

                    If Element_property(i).Material_for_shear > 0 Then

                        Create_one_link(Element_property(i).Material_for_shear, Material_collection)

                    End If

                    If Element_property(i).Material_for_membrane_bending_coupling > 0 Then

                        Create_one_link(Element_property(i).Material_for_membrane_bending_coupling, Material_collection)

                    End If


                Case "PCOMP"

                    For j As Integer = 0 To UBound(Element_property(i).Layers)

                        Create_one_link(Element_property(i).Layers(j).Material, Material_collection)

                    Next j

                Case "PSOLID"

                    If Element_property(i).Coordinate_system > 0 Then

                        Create_one_link(Element_property(i).Coordinate_system, Coordinate_system_collection)

                    End If


            End Select


            Create_one_link(Element_property(i).Material, Material_collection)

        Next i

        ' Отмечаем родителей материалов

        For i As Integer = 0 To UBound(Material)

            If Material(i).Coordinate_system > 0 Then

                Create_one_link(Material(i).Coordinate_system, Coordinate_system_collection)

            End If

        Next i

        ' Отмечаем родителей координатных систем

        For i As Integer = 0 To UBound(Coordinate_system)

            If Coordinate_system(i).Coordinate_system > 0 Then

                Create_one_link(Coordinate_system(i).Coordinate_system, Coordinate_system_collection)

            End If

            If Coordinate_system(i).point_A > 0 Then

                Create_one_link(Coordinate_system(i).point_A, Grid_collection)

                Create_one_link(Coordinate_system(i).point_B, Grid_collection)

                Create_one_link(Coordinate_system(i).point_C, Grid_collection)

            End If

        Next i

        'Отмечаем родителей узлов

        For i As Integer = 0 To UBound(Grid)

            If Grid(i).Coordinate_system > 0 Then

                Create_one_link(Grid(i).Coordinate_system, Coordinate_system_collection)

            End If

        Next i

        ' Отмечаем родителей нагрузок

        For i As Integer = 0 To UBound(Load)

            If Load(i).Coordinate_system > 0 Then

                Create_one_link(Load(i).Coordinate_system, Coordinate_system_collection)

            End If

            tmp_Value = -1

            tmp_Value = UBound(Load(i).Grid)

            If tmp_Value <> -1 Then

                ' отмечаемся в узлах
                For j As Long = 0 To UBound(Load(i).Grid)

                    If Load(i).Grid(j) > 0 Then

                        Create_one_link(Load(i).Grid(j), Grid_collection)

                    End If

                Next j

            End If

            tmp_Value = -1

            tmp_Value = Load(i).Grid_

            If tmp_Value <> -1 Then

                Create_one_link(Load(i).Grid_, Grid_collection)

            End If


            tmp_Value = -1

            tmp_Value = UBound(Load(i).Element_number)

            If tmp_Value <> -1 Then

                ' отмечаемся в элементах
                For j As Long = 0 To UBound(Load(i).Element_number)

                    Create_one_link(Load(i).Element_number(j), Element_collection)

                Next j

            End If

            tmp_Value = -1

            tmp_Value = Load(i).Element_number

            If tmp_Value <> -1 Then

                Create_one_link(Load(i).Element_number, Element_collection)

            End If


        Next i

        ' отмечаемся в граничных условиях

        For i As Integer = 0 To UBound(Boundary_condition)

            tmp_Value = -1

            tmp_Value = Boundary_condition(i).Coordinate_system

            If tmp_Value <> -1 Then

                Create_one_link(Boundary_condition(i), Coordinate_system_collection)

            End If

            ' отмечаемся в узлах
            For j As Long = 0 To UBound(Boundary_condition(i).Grid)

                If Boundary_condition(i).Grid(i) > 0 Then

                    Create_one_link(Boundary_condition(i).Grid(i), Grid_collection)

                End If

            Next j


        Next i

        On Error GoTo 0

    End Sub
    Function Convert_array_to_collection(ByRef Array() As Object) As Collection

        Convert_array_to_collection = New Collection

        For i As Long = 0 To UBound(Array)

            Convert_array_to_collection.Add(Array(i), Array(i).Number)

        Next i

    End Function

    'Function Convert_collection_to_array(ByRef Collection As Collection) As Object()

    '    Dim tmp_Array(Collection.Count - 1) As Object

    '    For i As Long = 1 To Collection.Count

    '        tmp_Array = Collection(i)

    '    Next i

    'End Function


    ''' <summary>
    ''' This sub links two model items (Elements, Grid, Element properties, Materials, Coordinate systems and etc.), one of them is a parent, another - child.
    ''' I.e., element CQUAD4 101 on Grid №№1001,1002, 1003, 1004 is a child for all of them,
    ''' and all of these Grid are a parents for it.
    ''' </summary>
    ''' <remarks></remarks>
    Sub Create_one_link(ByRef Reference As Object, ByRef Array() As Object)

        For i As Integer = 0 To UBound(Array)

            If Array(i).Number = Reference Then

                Reference = Array(i)

                Exit Sub

            End If

        Next i

    End Sub

    Sub Create_one_link(ByRef Reference As Object, ByRef Collection As Collection)

        Reference = Collection(Reference.ToString)

        'For i As Integer = 1 To Collection.Count

        '    If Array(i).Number = Reference Then

        '        Reference = Array(i)

        '        Exit Sub

        '    End If

        'Next i

    End Sub

    ' Процедура определяет центр массива точек и нормаль. Предполагается, что все точки находятся в одной плоскости
    ' Центр ищется как пересечение медиан первых двух углов, возвращается центр как вектор Center
    ' Нормаль ищется как векторное произведение векторов 1-2 и 2-3, и затем нормируется. Выдается в глобальной системе координат

    Sub Calculate_Face_Center_and_coord_axes()

        Dim Triangle As Boolean

        For i As Integer = 0 To UBound(Face)

            If UBound(Face(i).Grid) = 1 Then

                'Dim Element_grid_1, Element_grid_2, Diagonal, Side As New Vector

                'Element_grid_1 = Face(i).Element.Grid(0).Position

                'Element_grid_2 = Face(i).Element.Grid(2).Position

                'Diagonal = Element_grid_2 - Element_grid_1

                'Side = Face(i).Grid(1).Position - Face(i).Grid(0).Position

                'Face(i).Center = 0.5 * Side + Face(i).Grid(0).Position

                'Face(i).ox = Diagonal ^ Side

                'Face(i).ox = Norm_vector(Face(i).ox)

                'Face(i).oy = Norm_vector(Side)

                'Diagonal = Face(i).ox

                'Side = Face(i).oy

                'Face(i).oz = Diagonal ^ Side

                GoTo Skip_face

            End If

            If UBound(Face(i).Grid) = 2 Then Triangle = True

            If UBound(Face(i).Grid) = 3 Then Triangle = False

            Dim Grid_1, Grid_2, Grid_3, Grid_4 As Grid

            Grid_1 = Face(i).Grid(0)

            Grid_2 = Face(i).Grid(1)

            Grid_3 = Face(i).Grid(2)

            Grid_4 = New Grid

            If Not (Triangle) Then

                Grid_4 = Face(i).Grid(3)

            End If


            Dim x_1, y_1, z_1, x_2, y_2, z_2, x_3, y_3, z_3, x_4, y_4, z_4 As Double

            x_1 = Grid_1.Position.Coord(0)

            y_1 = Grid_1.Position.Coord(1)

            z_1 = Grid_1.Position.Coord(2)


            x_2 = Grid_2.Position.Coord(0)

            y_2 = Grid_2.Position.Coord(1)

            z_2 = Grid_2.Position.Coord(2)



            x_3 = Grid_3.Position.Coord(0)

            y_3 = Grid_3.Position.Coord(1)

            z_3 = Grid_3.Position.Coord(2)


            If Not (Triangle) Then

                x_4 = Grid_4.Position.Coord(0)

                y_4 = Grid_4.Position.Coord(1)

                z_4 = Grid_4.Position.Coord(2)


            End If

            ' угловые коэффиценты прямых-сторон

            Dim K_21_yx, K_21_zx, K_21_zy, K_32_yx, K_32_zx, K_32_zy, K_43_yx, _
            K_43_zx, K_43_zy, K_14_yx, K_14_zx, K_14_zy As Double

            K_21_yx = (y_2 - y_1) / (x_2 - x_1 + 0.0000000001)

            K_21_zx = (z_2 - z_1) / (x_2 - x_1 + 0.0000000001)

            K_21_zy = (z_2 - z_1) / (y_2 - y_1 + 0.0000000001)


            K_32_yx = (y_3 - y_2) / (x_3 - x_2 + 0.0000000001)

            K_32_zx = (z_3 - z_2) / (x_3 - x_2 + 0.0000000001)

            K_32_zy = (z_3 - z_2) / (y_3 - y_2 + 0.0000000001)

            If Not (Triangle) Then

                K_43_yx = (y_4 - y_3) / (x_4 - x_3 + 0.0000000001)

                K_43_zx = (z_4 - z_3) / (x_4 - x_3 + 0.0000000001)

                K_43_zy = (z_4 - z_3) / (y_4 - y_3 + 0.0000000001)


                K_14_yx = (y_1 - y_4) / (x_1 - x_4 + 0.0000000001)

                K_14_zx = (z_1 - z_4) / (x_1 - x_4 + 0.0000000001)

                K_14_zy = (z_1 - z_4) / (y_1 - y_4 + 0.0000000001)

            End If

            'координаты середин сторон

            Dim x_21_c, y_21_c, z_21_c, x_32_c, y_32_c, _
            z_32_c, x_43_c, y_43_c, z_43_c, x_14_c, y_14_c, z_14_c As Double

            x_21_c = (x_2 + x_1) / 2

            y_21_c = K_21_yx * (x_21_c - x_1 + 0.0000000001 / 2) + y_1

            z_21_c = K_21_zx * (x_21_c - x_1 + 0.0000000001 / 2) + z_1



            x_32_c = (x_3 + x_2) / 2

            y_32_c = K_32_yx * (x_32_c - x_2 + 0.0000000001 / 2) + y_2

            z_32_c = K_32_zx * (x_32_c - x_2 + 0.0000000001 / 2) + z_2


            If Not (Triangle) Then

                x_43_c = (x_4 + x_3) / 2

                y_43_c = K_43_yx * (x_43_c - x_3 + 0.0000000001 / 2) + y_3

                z_43_c = K_43_zx * (x_43_c - x_3 + 0.0000000001 / 2) + z_3


                x_14_c = (x_1 + x_4) / 2

                y_14_c = K_14_yx * (x_14_c - x_4 + 0.0000000001 / 2) + y_4

                z_14_c = K_14_zx * (x_14_c - x_4 + 0.0000000001 / 2) + z_4

            End If


            ' Определяем угловые коэффициенты прямых, которые при пересечении дают центр

            Dim K_1_yx, K_1_zx, K_1_zy, b_1_yx, b_1_zx, b_1_zy As Double

            Dim K_2_yx, K_2_zx, K_2_zy, b_2_yx, b_2_zx, b_2_zy As Double

            If Triangle Then

                K_1_yx = (y_3 - y_21_c + 0.0000000001) / (x_3 - x_21_c + 0.0000000001)

                K_1_zx = (z_3 - z_21_c + 0.0000000001) / (x_3 - x_21_c + 0.0000000001)

                K_1_zy = (z_3 - z_21_c + 0.0000000001) / (y_3 - y_21_c + 0.0000000001)

                b_1_yx = y_21_c - K_1_yx * x_21_c

                b_1_zx = z_21_c - K_1_zx * x_21_c

                b_1_zy = z_21_c - K_1_zy * y_21_c


                K_2_yx = (y_32_c - y_1 + 0.0000000001) / (x_32_c - x_1 + 0.0000000001)

                K_2_zx = (z_32_c - z_1 + 0.0000000001) / (x_32_c - x_1 + 0.0000000001)

                K_2_zy = (z_32_c - z_1 + 0.0000000001) / (y_32_c - y_1 + 0.0000000001)

                b_2_yx = y_1 - K_2_yx * x_1

                b_2_zx = z_1 - K_2_zx * x_1

                b_2_zy = z_1 - K_2_zy * y_1

            End If

            If Not (Triangle) Then

                K_1_yx = (y_43_c - y_21_c + 0.0000000001) / (x_43_c - x_21_c + 0.0000000001)

                K_1_zx = (z_43_c - z_21_c + 0.0000000001) / (x_43_c - x_21_c + 0.0000000001)

                K_1_zy = (z_43_c - z_21_c + 0.0000000001) / (y_43_c - y_21_c + 0.0000000001)

                b_1_yx = y_21_c - K_1_yx * (x_21_c + +0.0000000001 / 2)

                b_1_zx = z_21_c - K_1_zx * (x_21_c + 0.0000000001 / 2)

                b_1_zy = z_21_c - K_1_zy * (y_21_c + 0.0000000001 / 2)


                K_2_yx = (y_14_c - y_32_c + 0.0000000001) / (x_14_c - x_32_c + 0.0000000001)

                K_2_zx = (z_14_c - z_32_c + 0.0000000001) / (x_14_c - x_32_c + 0.0000000001)

                K_2_zy = (z_14_c - z_32_c + 0.0000000001) / (y_14_c - y_32_c + 0.0000000001)

                b_2_yx = y_32_c - K_2_yx * (x_32_c + 0.0000000001 / 2)

                b_2_zx = z_32_c - K_2_zx * (x_32_c + 0.0000000001 / 2)

                b_2_zy = z_32_c - K_2_zy * (y_32_c + 0.0000000001 / 2)


            End If

            ' координаты центра элемента

            Dim x_c, y_c, z_c As Double

            If K_1_yx - K_2_yx <> 0 Then

                x_c = (b_2_yx - b_1_yx) / (K_1_yx - K_2_yx)

                y_c = K_1_yx * x_c + b_1_yx

                z_c = K_1_zx * x_c + b_1_zx

            Else

                If K_1_zx - K_2_zx <> 0 Then

                    x_c = (b_2_zx - b_1_zx) / (K_1_zx - K_2_zx)

                    y_c = K_1_yx * x_c + b_1_yx

                    z_c = K_1_zx * x_c + b_1_zx

                End If

            End If



            ' запись координат в вектор

            Face(i).Center.Coord(0) = x_c

            Face(i).Center.Coord(1) = y_c

            Face(i).Center.Coord(2) = z_c

            ' расчет направляющих векторов системы координат элемента
            ' использованная система координат элемента отличается от принятой в NASTRAN
            ' центр находится на пересечении медиан сторон в случае треугольника.
            ' в случае четырехугольника центр располагается на пересечении прямых, соединяющих середины сторон
            ' ось z - перпендикулярна плоскости элемента
            ' ось x - параллельна стороне 1-2
            ' ось y  - нормированное векторное произведение (oZ x oX)


            Dim v_12 As Vector

            Dim v_23 As Vector

            v_12 = Grid_2.Position - Grid_1.Position

            v_23 = Grid_3.Position - Grid_2.Position

            ' ось Z
            Face(i).oz = v_12 ^ v_23

            Face(i).oz = Norm_vector(Face(i).oz)

            '' ось X
            'Face(i).ox = Norm_vector(v_12)

            '' ось Y
            'v_12 = Face(i).oz

            'v_23 = Face(i).ox

            'Face(i).oy = v_12 ^ v_23

            ' ось Y
            Face(i).oy = Norm_vector(v_12)

            ' ось X
            v_12 = Face(i).oz

            v_23 = Face(i).oy

            Face(i).ox = v_23 ^ v_12

            ' Рассчитываем матрицу перехода в систему координат грани и обратно

            Face(i).Calculate_transition_matrices()

            ' Рассчитываем координаты ыершин граней в системе координат грани

            Face(i).Calc_local_coordinates()

Skip_face:

        Next i

    End Sub


    Sub Calculate_2_Grid_faces_Center_and_coord_axes()

        For i As Integer = 0 To UBound(Face)

            If UBound(Face(i).Grid) = 1 Then

                Dim Element_grid_1, Element_grid_2, Diagonal, Side As New Vector

                Element_grid_1 = Face(i).Element.Grid(0).Position

                Element_grid_2 = Face(i).Element.Grid(2).Position

                Diagonal = Element_grid_2 - Element_grid_1

                Side = Face(i).Grid(1).Position - Face(i).Grid(0).Position

                Face(i).Center = 0.5 * Side + Face(i).Grid(0).Position

                Face(i).ox = Diagonal ^ Side

                Face(i).ox = Norm_vector(Face(i).ox)

                Face(i).oy = Norm_vector(Side)

                Diagonal = Face(i).ox

                Side = Face(i).oy

                Face(i).oz = Diagonal ^ Side

            End If

        Next i

    End Sub

    ''' <summary>
    ''' This function creates array of polygons, suitable for visualisation.
    ''' </summary>
    ''' <param name="Face">Incoming faces, basis for extraction</param>
    ''' <returns>Array of polygons, suitable for visualisation</returns>
    ''' <remarks></remarks>
    Public Function Extract_Polygon_from_Face(ByRef Face() As Face) As Visualisation.Polygon()

        ' Подготавливаем массив многоугольников для передачи 
        Dim Polygon(UBound(Face)) As Visualisation.Polygon

        Dim i_Polygon As Long = 0

        For i As Integer = 0 To UBound(Face)

            'If UBound(Face(i).Grid) > 1 Then

            Polygon(i_Polygon) = New Visualisation.Polygon(UBound(Face(i).Grid) + 1)

            Polygon(i_Polygon).Red = Face(i).Code.R

            Polygon(i_Polygon).Green = Face(i).Code.G

            Polygon(i_Polygon).Blue = Face(i).Code.B

            Polygon(i_Polygon).Text = Face(i).Code.Text

            For j As Integer = 0 To UBound(Face(i).Grid)

                Polygon(i_Polygon).Vertex(j) = New Visualisation.Vertex

                Polygon(i_Polygon).Vertex(j).x = Face(i).Grid(j).Position.Coord(0)

                Polygon(i_Polygon).Vertex(j).y = Face(i).Grid(j).Position.Coord(1)

                Polygon(i_Polygon).Vertex(j).z = Face(i).Grid(j).Position.Coord(2)

            Next j

            i_Polygon += 1

            'End If

        Next i

        'ReDim Preserve Polygon(i_Polygon - 1)

        Return Polygon

    End Function


    ''' <summary>
    ''' This function creates array of polygons, suitable for visualisation.
    ''' </summary>
    ''' <returns>Array of polygons, suitable for visualisation</returns>
    ''' <remarks></remarks>
    Public Function Extract_Polygon_colors_from_Face(ByRef Polygon() As Visualisation.Polygon) As Visualisation.Polygon()


        For i As Integer = 0 To UBound(Polygon)

            Polygon(i).Red = Shining_face(i).Code.R

            Polygon(i).Green = Shining_face(i).Code.G

            Polygon(i).Blue = Shining_face(i).Code.B

            Polygon(i).Text = Shining_face(i).Code.Text

        Next i

        Return Polygon

    End Function

    Sub Extract_shining_faces()

        '' Подготавливаем массив многоугольников для передачи 
        'ReDim Polygon(UBound(Face))

        'Dim i_Polygon As Long = 0

        'For i As Integer = 0 To UBound(Face)

        '    If UBound(Face(i).Grid) > 1 Then

        '        Polygon(i_Polygon) = New Visualisation.Polygon(UBound(Face(i).Grid) + 1)

        '        Dim F_c As New Face_code

        '        Face(i).Detect_face_number(F_c, i)

        '        Polygon(i_Polygon).Red = F_c.R

        '        Polygon(i_Polygon).Green = F_c.G

        '        Polygon(i_Polygon).Blue = F_c.B

        '        For j As Integer = 0 To UBound(Face(i).Grid)

        '            Polygon(i_Polygon).Vertex(j) = New Visualisation.Vertex

        '            Polygon(i_Polygon).Vertex(j).x = Face(i).Grid(j).Position.Coord(0)

        '            Polygon(i_Polygon).Vertex(j).y = Face(i).Grid(j).Position.Coord(1)

        '            Polygon(i_Polygon).Vertex(j).z = Face(i).Grid(j).Position.Coord(2)

        '        Next j

        '        i_Polygon += 1

        '    End If

        'Next i

        'ReDim Preserve Polygon(i_Polygon - 1)


        ReDim Shining_face(Max_items_default_2)

        Dim i_Shining_face As Integer = -1

        For i As Integer = 0 To UBound(Face)

            If Face(i).Can_Radiate Then

                If UBound(Face(i).Grid) > 1 Then

                    i_Shining_face += 1

                    Shining_face(i_Shining_face) = Face(i)

                    Shining_face(i_Shining_face).Number = i_Shining_face

                End If

            End If

        Next i

        ReDim Preserve Shining_face(i_Shining_face)


        'Polygon = Extract_Polygon_from_Face(Shining_face)


        ' теперь для тех излучающих граней, от которых излучение уходит в пустоту, выставляем признак Enabled = False
        ' Признак видимости - если нормаль к грани и нормаль какой-либой другой излучающей грани, расположенной ПЕРЕД
        ' испытуемой гранью образуют тупой угол (>90 градусов), то тогда считаем, 
        ' что испытуемая грань излучает не в пустоту

        ' сначала все грани считаются излучающими в пустоту

        For i As Integer = 0 To UBound(Shining_face)

            Shining_face(i).Enabled = True

        Next i


        'For i As Integer = 0 To UBound(Shining_face)

        '    For j As Integer = 0 To UBound(Shining_face)

        '        If Cos_angle_between_vectors(Shining_face(i).oz, Me.Check_zero_vector(Shining_face(j).Center - Shining_face(i).Center)) > 0.001 Then

        '            If Cos_angle_between_vectors(Shining_face(i).oz, Shining_face(j).oz) < 0 Then

        '                Shining_face(i).Enabled = True

        '                Exit For

        '            End If

        '        End If

        '    Next j


        'Next i

    End Sub


    Public Sub New(Optional ByRef File_Path_File_name As String = "")

        Loading(File_Path_File_name)

        Loading_TOP(File_Path_File_name)

    End Sub

    Public Sub Save(ByRef Output_filename As String)

        Me.String_in_file = ""

        For i As Integer = 0 To UBound(Boundary_condition)

            Boundary_condition(i).Save()

            Me.String_in_file &= Boundary_condition(i).String_in_file

        Next i

        For i As Integer = 0 To UBound(Coordinate_system)

            Coordinate_system(i).Save()

            Me.String_in_file &= Coordinate_system(i).String_in_file

        Next i

        For i As Integer = 0 To UBound(Element)

            Element(i).Save()

            Me.String_in_file &= Element(i).String_in_file

        Next i

        For i As Integer = 0 To UBound(Element_property)

            Element_property(i).Save()

            Me.String_in_file &= Element_property(i).String_in_file

        Next i

        For i As Integer = 0 To UBound(Grid)

            Grid(i).Save()

            Me.String_in_file &= Grid(i).String_in_file

        Next i

        For i As Integer = 0 To UBound(Load)

            Load(i).Save()

            Me.String_in_file &= Load(i).String_in_file

        Next i

        For i As Integer = 0 To UBound(Material)

            Material(i).Save()

            Me.String_in_file &= Material(i).String_in_file

        Next i

        Dim writer As System.IO.FileStream = New System.IO.FileStream(Output_filename, IO.FileMode.Create, IO.FileAccess.Write)

        Dim w As New System.IO.StreamWriter(writer)

        w.WriteLine(Me.String_in_file)

        w.Close()

        writer.Close()

        Me.String_in_file = ""


    End Sub


    Public Sub Loading(ByVal File_Path_File_name)

        Me.State = "File reading"

        RaiseEvent Model_state_changed(Me)

        If File_Path_File_name <> "" Then

            Dim Whole_dat As String

            Dim File_length As Long

            Dim Point_position, Slash_position As Integer

            'определяем положение последней точки
            Point_position = InStrRev(File_Path_File_name, ".")
            'определяем положение последего слэша
            Slash_position = InStrRev(File_Path_File_name, "\")
            ' имя файла
            File_name = Mid(File_Path_File_name, Slash_position + 1, Point_position - (Slash_position + 1))
            ' путь к файлу
            File_path = Mid(File_Path_File_name, 1, Slash_position)

            'Const ForReading = 1, ForWriting = 2, ForAppending = 3
            'Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
            Dim fs, F, ts
            fs = CreateObject("Scripting.FileSystemObject")

            If fs.fileExists(File_Path_File_name) Then

                F = fs.GetFile(File_Path_File_name)

            Else

                MsgBox("No such file " & File_Path_File_name)

                Exit Sub

            End If

            ts = F.OpenAsTextStream(1, 0)

            ' считываем один входной файл для Nastran'a. Учитываем, что его расширение может быть как dat, так и bdf.
            Whole_dat = ts.ReadAll

            File_length = Len(Whole_dat)

            ' индицицуем начало процесса считывания

            RaiseEvent Model_state_changed(Me)

            'инициализируем массивы для храния объектов

            ReDim Me.Boundary_condition(Max_items_default_2)

            ReDim Me.Coordinate_system(Max_items_default_2)

            ReDim Me.Element(Max_items_default_2)

            ReDim Me.Element_property(Max_items_default_2)

            ReDim Me.Grid(Max_items_default_2)

            ReDim Me.Load(Max_items_default_2)

            ReDim Me.Material(Max_items_default_2)




            Dim Position As Integer = 1

            Dim old_Position As Integer = 1

            'Dim Property_string As String


            Dim Whole_dat_array() As String

            Dim t As New ACCEL

            Whole_dat_array = t.Convert_string_in_file_to_array_3(Whole_dat)

            Whole_dat = ""

            ' BOUNDARY CONDITIONS

            ' MPC

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "MPC" Or Whole_dat_array(i) = "MPC*" Then

                    i_BC += 1

                    Boundary_condition(i_BC) = New MPC(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i


            ' MPCADD

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "MPCADD" Or Whole_dat_array(i) = "MPCADD*" Then

                    i_BC += 1

                    Boundary_condition(i_BC) = New MPCADD(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i


            ' SPC

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "SPC" Or Whole_dat_array(i) = "SPC*" Then

                    i_BC += 1

                    Boundary_condition(i_BC) = New SPC(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i


            ' SPC1

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "SPC1" Or Whole_dat_array(i) = "SPC1*" Then

                    i_BC += 1

                    Boundary_condition(i_BC) = New SPC1(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i

            ' SPCADD

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "SPCADD" Or Whole_dat_array(i) = "SPCADD*" Then

                    i_BC += 1

                    Boundary_condition(i_BC) = New SPCADD(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i



            ' SPCD

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "SPCD" Or Whole_dat_array(i) = "SPCD*" Then

                    i_BC += 1

                    Boundary_condition(i_BC) = New SPCD(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i

            ' SPCOFF

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "SPCOFF" Or Whole_dat_array(i) = "SPCOFF*" Then

                    i_BC += 1

                    Boundary_condition(i_BC) = New SPCOFF(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i


            ' SPCOFF1

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "SPCOFF1" Or Whole_dat_array(i) = "SPCOFF1*" Then

                    i_BC += 1

                    Boundary_condition(i_BC) = New SPCOFF1(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i

            ' COORDINATE SYSTEMS

            ' CORD1C

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "CORD1C" Or Whole_dat_array(i) = "CORD1C*" Then

                    i_CS += 1

                    Coordinate_system(i_CS) = New CORD1C(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i


            ' CORD2C

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "CORD2C" Or Whole_dat_array(i) = "CORD2C*" Then

                    i_CS += 1

                    Coordinate_system(i_CS) = New CORD2C(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i

            ' CORD1R

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "CORD1R" Or Whole_dat_array(i) = "CORD1R*" Then

                    i_CS += 1

                    Coordinate_system(i_CS) = New CORD1R(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i

            ' CORD2R

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "CORD2R" Or Whole_dat_array(i) = "CORD2R*" Then

                    i_CS += 1

                    Coordinate_system(i_CS) = New CORD2R(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i


            ' CORD1S

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "CORD1S" Or Whole_dat_array(i) = "CORD1S*" Then

                    i_CS += 1

                    Coordinate_system(i_CS) = New CORD1S(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i

            ' CORD2S

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "CORD2S" Or Whole_dat_array(i) = "CORD2S*" Then

                    i_CS += 1

                    Coordinate_system(i_CS) = New CORD2S(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i

            ' ELEMENTS


            ' CBAR

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "CBAR" Or Whole_dat_array(i) = "CBAR*" Then

                    i_E += 1

                    Element(i_E) = New CBAR(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i


            ' CBEAM

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "CBEAM" Or Whole_dat_array(i) = "CBEAM*" Then

                    i_E += 1

                    Element(i_E) = New CBEAM(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i


            ' CHEXA

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "CHEXA" Or Whole_dat_array(i) = "CHEXA*" Then

                    i_E += 1

                    Element(i_E) = New CHEXA(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i


            ' CPENTA

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "CPENTA" Or Whole_dat_array(i) = "CPENTA*" Then

                    i_E += 1

                    Element(i_E) = New CPENTA(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i


            ' CQUAD4

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "CQUAD4" Or Whole_dat_array(i) = "CQUAD4*" Then

                    i_E += 1

                    Element(i_E) = New CQUAD4(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i



            ' CQUAD8

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "CQUAD8" Or Whole_dat_array(i) = "CQUAD8*" Then

                    i_E += 1

                    Element(i_E) = New CQUAD8(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i


            ' CROD

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "CROD" Or Whole_dat_array(i) = "CROD*" Then

                    i_E += 1

                    Element(i_E) = New CROD(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i


            ' CSHEAR

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "CSHEAR" Or Whole_dat_array(i) = "CSHEAR*" Then

                    i_E += 1

                    Element(i_E) = New CSHEAR(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i

            ' CTETRA

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "CTETRA" Or Whole_dat_array(i) = "CTETRA*" Then

                    i_E += 1

                    Element(i_E) = New CTETRA(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i

            ' CTRIA3

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "CTRIA3" Or Whole_dat_array(i) = "CTRIA3*" Then

                    i_E += 1

                    Element(i_E) = New CTRIA3(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i


            ' CTRIA6

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "CTRIA6" Or Whole_dat_array(i) = "CTRIA6*" Then

                    i_E += 1

                    Element(i_E) = New CTRIA6(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i


            ' CBAR

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "CBAR" Or Whole_dat_array(i) = "CBAR*" Then

                    i_E += 1

                    Element(i_E) = New CBAR(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i


            ' ELEMENT PROPERTIES


            ' PBAR

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "PBAR" Or Whole_dat_array(i) = "PBAR*" Then

                    i_EP += 1

                    Element_property(i_EP) = New PBAR(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i



            ' PBEAM

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "PBEAM" Or Whole_dat_array(i) = "PBEAM*" Then

                    i_EP += 1

                    Element_property(i_EP) = New PBEAM(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i


            ' PCOMP

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "PCOMP" Or Whole_dat_array(i) = "PCOMP*" Then

                    i_EP += 1

                    Element_property(i_EP) = New PCOMP(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i

            ' PROD

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "PROD" Or Whole_dat_array(i) = "PROD*" Then

                    i_EP += 1

                    Element_property(i_EP) = New PROD(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i

            ' PSHEAR

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "PSHEAR" Or Whole_dat_array(i) = "PSHEAR*" Then

                    i_EP += 1

                    Element_property(i_EP) = New PSHEAR(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i


            ' PSHELL

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "PSHELL" Or Whole_dat_array(i) = "PSHELL*" Then

                    i_EP += 1

                    Element_property(i_EP) = New PSHELL(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i



            ' PSOLID

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "PSOLID" Or Whole_dat_array(i) = "PSOLID*" Then

                    i_EP += 1

                    Element_property(i_EP) = New PSOLID(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i




            ' GRID

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "GRID" Or Whole_dat_array(i) = "GRID*" Then

                    i_G += 1

                    Grid(i_G) = New Grid(Copy_array_8(Whole_dat_array, i, 8))

                End If

                Application.DoEvents() : Next i





            ' LOADS

            ' ACCEL

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "ACCEL" Or Whole_dat_array(i) = "ACCEL*" Then

                    i_L += 1

                    Load(i_L) = New ACCEL(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i

            ' ACCEL1

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "ACCEL1" Or Whole_dat_array(i) = "ACCEL1*" Then

                    i_L += 1

                    Load(i_L) = New ACCEL1(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i


            ' FORCE

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "FORCE" Or Whole_dat_array(i) = "FORCE*" Then

                    i_L += 1

                    Load(i_L) = New FORCE(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i

            ' FORCE1

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "FORCE1" Or Whole_dat_array(i) = "FORCE1*" Then

                    i_L += 1

                    Load(i_L) = New FORCE1(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i

            ' FORCE2

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "FORCE2" Or Whole_dat_array(i) = "FORCE2*" Then

                    i_L += 1

                    Load(i_L) = New FORCE2(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i

            ' GRAV

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "GRAV" Or Whole_dat_array(i) = "GRAV*" Then

                    i_L += 1

                    Load(i_L) = New GRAV(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i

            ' LOAD

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "LOAD" Or Whole_dat_array(i) = "LOAD*" Then

                    i_L += 1

                    Load(i_L) = New LOAD(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i

            ' MOMENT

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "MOMENT" Or Whole_dat_array(i) = "MOMENT*" Then

                    i_L += 1

                    Load(i_L) = New MOMENT(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i

            ' MOMENT1

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "MOMENT1" Or Whole_dat_array(i) = "MOMENT1*" Then

                    i_L += 1

                    Load(i_L) = New MOMENT1(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i

            ' MOMENT2

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "MOMENT2" Or Whole_dat_array(i) = "MOMENT2*" Then

                    i_L += 1

                    Load(i_L) = New MOMENT2(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i


            ' PLOAD

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "PLOAD" Or Whole_dat_array(i) = "PLOAD*" Then

                    i_L += 1

                    Load(i_L) = New PLOAD(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i


            ' PLOAD1

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "PLOAD1" Or Whole_dat_array(i) = "PLOAD1*" Then

                    i_L += 1

                    Load(i_L) = New PLOAD1(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i

            ' PLOAD2

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "PLOAD2" Or Whole_dat_array(i) = "PLOAD2*" Then

                    i_L += 1

                    Load(i_L) = New PLOAD2(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i


            ' PLOAD4

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "PLOAD4" Or Whole_dat_array(i) = "PLOAD4*" Then

                    i_L += 1

                    Load(i_L) = New PLOAD4(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i

            ' RFORCE

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "RFORCE" Or Whole_dat_array(i) = "RFORCE*" Then

                    i_L += 1

                    Load(i_L) = New RFORCE(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i

            ' TEMP

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "TEMP" Or Whole_dat_array(i) = "TEMP*" Then

                    i_L += 1

                    Load(i_L) = New TEMP(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i


            ' TEMPBC

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "TEMPBC" Or Whole_dat_array(i) = "TEMPBC*" Then

                    i_L += 1

                    Load(i_L) = New TEMPBC(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i


            ' TEMPD

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "TEMPD" Or Whole_dat_array(i) = "TEMPD*" Then

                    i_L += 1

                    Load(i_L) = New TEMPD(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i


            ' TEMPF

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "TEMPF" Or Whole_dat_array(i) = "TEMPF*" Then

                    i_L += 1

                    Load(i_L) = New TEMPF(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i



            ' TEMPP1

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "TEMPP1" Or Whole_dat_array(i) = "TEMPP1*" Then

                    i_L += 1

                    Load(i_L) = New TEMPP1(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i


            ' TEMPRB

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "TEMPRB" Or Whole_dat_array(i) = "TEMPRB*" Then

                    i_L += 1

                    Load(i_L) = New TEMPRB(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i

            ' MATERIALS

            ' MAT1

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "MAT1" Or Whole_dat_array(i) = "MAT1*" Then

                    i_M += 1

                    Material(i_M) = New MAT1(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i


            ' MAT2

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "MAT2" Or Whole_dat_array(i) = "MAT2*" Then

                    i_M += 1

                    Material(i_M) = New MAT2(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i

            ' MAT4

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "MAT4" Or Whole_dat_array(i) = "MAT4*" Then

                    i_M += 1

                    Material(i_M) = New MAT4(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i

            ' MAT8

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "MAT8" Or Whole_dat_array(i) = "MAT8*" Then

                    i_M += 1

                    Material(i_M) = New MAT8(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i

            ' MAT9

            For i As Long = 0 To UBound(Whole_dat_array) Step 1

                If Whole_dat_array(i) = "MAT9" Or Whole_dat_array(i) = "MAT9*" Then

                    i_M += 1

                    Material(i_M) = New MAT9(Copy_array_8(Whole_dat_array, i, Max_items_default))

                End If

                Application.DoEvents() : Next i




            'удаляем лишние элементы

            ReDim Preserve Boundary_condition(i_BC)

            ReDim Preserve Coordinate_system(i_CS)

            ReDim Preserve Element(i_E)

            ReDim Preserve Element_property(i_EP)

            ReDim Preserve Grid(i_G)

            ReDim Preserve Load(i_L)

            ReDim Preserve Material(i_M)


            'Boundary_condition = Clear_array(Boundary_condition)

            'Coordinate_system = Clear_array(Coordinate_system)

            'Element = Clear_array(Element)

            'Element_property = Clear_array(Element_property)

            'Grid = Clear_array(Grid)

            'Load = Clear_array(Load)

            'Material = Clear_array(Material)

            RaiseEvent Model_state_changed(Me)

            Create_links()

            Create_faces()

            Calc_mass_volumes_centroids_inertia_tensors()

            Find_Model_Center()

            Find_model_size()

            '' Процедура считывает и заполняет массивы термо-оптических свойств проекта
            '' Принимается, что файл с данными находится по тому же пути, что и файлы проекта, но имеет расширение *.top (Thermal and optical properties)
            'Call Read_and_fill_thermal_and_optical_properties(One_project)


            '' процедура рассчитывает положение характерных точек для систем координат всех элементов типа CQUAD4
            'Call fill_CQUAD4_coord_systems(One_project)

            '' процедура рассчитывает положение характерных точек для систем координат всех элементов типа CTRIA3
            'Call fill_CTRIA3_coord_systems(One_project)

            '' процедура заполняет матрицы перехода из различных систем координат в глобальную и наоборот
            'Call fill_Coord_sys_matrixes(One_project)

            '' процедура заполняет "упругие" и "допускаемые" поля в массиве композиционных материалов проекта
            'Call fill_PCOMP_Special_fields(One_project)

            '' процедура добавляет последним элементом массива систем координат глобальную систему координат
            'Call Add_coord_sys_0(One_project)

            '' Среда, заполняющая пространство в проекте имеет номер 0 в массиве MTOP1
            'Call fill_Medium(One_project)

            '' процедура переводит все узлы проекта в глобальную систему координат
            'Call Convert_Grid_to_Global_coordinates(One_project)

            '' Процедура вычисляет объемы всех элементов модели (у которых есть объем)
            'Call fill_Volumes(One_project)

            '' в самый хвост процедуры
            'On Error GoTo 0

            'Exit Sub

        End If

    End Sub



    Function Copy_array_8(ByRef In_array() As String, ByVal i_Start As Long, ByVal Length As Long) As String()

        Dim tmp_Array() As String

        Dim First_character_code As Integer

        ReDim tmp_Array(Length)

        tmp_Array(0) = In_array(i_Start)

        If i_Start + Length > UBound(In_array) Then

            Length = UBound(In_array) - i_Start

        End If

        For i As Long = i_Start + 1 To i_Start + Length

            First_character_code = Asc(Mid(In_array(i) & " ", 1, 1))

            If (First_character_code >= 65 And First_character_code <= 90) Then

                ReDim Preserve tmp_Array(i - i_Start - 1)

                Exit For

            End If

            tmp_Array(i - i_Start) = In_array(i)

        Next i

        Copy_array_8 = tmp_Array

    End Function

    Function Copy_array_16(ByRef In_array() As String, ByVal i_Start As Long, ByVal Length As Long) As String()

        Dim tmp_Array() As String

        ReDim tmp_Array(Length)

        For i As Long = i_Start To i_Start + Length

            tmp_Array(i - i_Start) = In_array(i)

        Next i

        Copy_array_16 = tmp_Array

    End Function

    Function Extract_field_array(ByRef In_array() As String, ByVal i_Start As Long) As String()

        Dim tmp_Array() As String

        Dim First_character_code, i As Long

        ReDim tmp_Array(500)

        tmp_Array(0) = In_array(i_Start)

        For i = i_Start + 1 To UBound(In_array)

            If In_array(i) <> "" Then

                First_character_code = Asc(Mid(In_array(i), 1, 1))

                If First_character_code >= 65 And First_character_code <= 90 Then

                    Exit For

                Else

                    tmp_Array(i - i_Start) = In_array(i)

                End If

            End If

        Next i

        ReDim Preserve tmp_Array(i - i_Start)

        Extract_field_array = tmp_Array

    End Function





    '''' <summary>
    '''' Function removes empty element from model item element array
    '''' </summary>
    '''' <param name="Array_for_cleaning"></param>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public Function Clear_array(ByRef Array_for_cleaning() As Object) As Object()

    '    On Error Resume Next

    '    Dim tmp_Array() As Object

    '    Dim tmp_i As Integer

    '    Dim tmp_Str As String

    '    ReDim tmp_Array(UBound(Array_for_cleaning))

    '    tmp_i = -1

    '    tmp_i = UBound(Array_for_cleaning)

    '    If tmp_i = -1 Then Return Nothing

    '    tmp_i = -1

    '    For i As Integer = 0 To UBound(Array_for_cleaning)

    '        tmp_Str = ""

    '        tmp_Str = Array_for_cleaning(i).String_in_file

    '        If tmp_Str <> "" Then

    '            tmp_i += 1

    '            tmp_Array(tmp_i) = Array_for_cleaning(i)

    '        End If

    '    Next i

    '    Clear_array = tmp_Array

    '    On Error GoTo 0

    'End Function


    ' Функция находит наибольшее растояние от центра модели и наименьшее расстояние между узлами в модели.
    Public Sub Find_model_size()

        Dim L As New Vector

        Maximum_length = -Double.MaxValue + 1

        For i As Long = 0 To UBound(Grid)

            L = Grid(i).Position - Center

            If L_vector(L) > Maximum_length Then

                Maximum_length = L_vector(L)

            End If

        Next i

        Maximum_length *= 2


        'Minimum_length = Double.MaxValue

        'Maximum_length = -Double.MaxValue

        'Dim tmp_L As Double

        'For i As Long = 0 To UBound(Element)

        '    tmp_L = L_vector(Element(i).Field.Centroid - Center)

        '    If tmp_L > Maximum_length Then

        '        Maximum_length = tmp_L

        '    End If

        'Next i

        Minimum_length = 0.001 * Maximum_length



        'Minimum_length = Double.MaxValue

        'Dim tmp_L As Double

        'For i As Long = 0 To UBound(Grid)

        '    For j As Long = 0 To UBound(Grid)

        '        tmp_L = L_vector(Grid(i).Position - Grid(j).Position)

        '        If tmp_L < Minimum_length Then

        '            If tmp_L > Epsilon Then

        '                Minimum_length = tmp_L

        '            End If

        '        End If

        '        If tmp_L > Maximum_length Then

        '            Maximum_length = tmp_L

        '        End If

        '    Next j

        '    Application.DoEvents()

        'Next i

    End Sub

    Public Sub Find_Model_Center()

        Dim i, N_Grids As Long

        Center = New Vector

        N_Grids = UBound(Grid) + 1

        For i = 0 To UBound(Grid)

            Center = Grid(i).Position + Center

        Next i

        Center.Coord(0) = Center.Coord(0) / N_Grids

        Center.Coord(1) = Center.Coord(1) / N_Grids

        Center.Coord(2) = Center.Coord(2) / N_Grids

    End Sub

    Public Sub Calc_mass_volumes_centroids_inertia_tensors()

        For i As Integer = 0 To UBound(Element)

            Element(i).Calc_volume_mass_inertia()

            Application.DoEvents()

        Next i

    End Sub

    Public Sub Loading_TOP(ByVal File_Path_File_name)

        Me.State = "File reading"

        If File_Path_File_name <> "" Then

            Dim Whole_TOP As String

            Dim File_length As Long

            Dim Point_position, Slash_position As Integer

            'определяем положение последней точки
            Point_position = InStrRev(File_Path_File_name, ".")
            'определяем положение последего слэша
            Slash_position = InStrRev(File_Path_File_name, "\")
            ' имя файла
            File_name = Mid(File_Path_File_name, Slash_position + 1, Point_position - (Slash_position + 1))
            ' путь к файлу
            File_path = Mid(File_Path_File_name, 1, Slash_position)

            File_Path_File_name = File_path & File_name & ".top"

            'Const ForReading = 1, ForWriting = 2, ForAppending = 3
            'Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
            Dim fs, F, ts
            fs = CreateObject("Scripting.FileSystemObject")

            If fs.fileExists(File_Path_File_name) Then

                F = fs.GetFile(File_Path_File_name)

            Else

                MsgBox("No such file " & File_Path_File_name)

                Exit Sub

            End If

            ts = F.OpenAsTextStream(1, 0)

            ' считываем один входной файл для Nastran'a. Учитываем, что его расширение может быть как dat, так и bdf.
            Whole_TOP = ts.ReadAll

            File_length = Len(Whole_TOP)

            ' индицицуем начало процесса считывания

            RaiseEvent Model_state_changed(Me)


            'инициализируем массивы для храния объектов

            ReDim Dependency(Max_items_default_2)

            ReDim FTOP(Max_items_default_2)

            ReDim MEDIUM(Max_items_default_2)

            ReDim RSOURCE(Max_items_default_2)

            ReDim HSOURCE(Max_items_default_2)

            ReDim T_INIT(Max_items_default_2)

            ReDim T_SET(Max_items_default_2)



            Dim Position As Integer = 1

            Dim old_Position As Integer = 1


            Dim Whole_TOP_array() As String

            Dim tmp As New ACCEL

            Whole_TOP_array = tmp.Convert_string_in_file_to_array_3(Whole_TOP)

            Whole_TOP = ""

            ' DEPENDENCIES

            For i As Long = 0 To UBound(Whole_TOP_array) Step 1

                If Whole_TOP_array(i) = "DEPEND" Or Whole_TOP_array(i) = "DEPEND*" Then

                    i_D += 1

                    Dependency(i_D) = New DEPEND(Copy_array_8(Whole_TOP_array, i, Max_items_default_2))

                End If

            Next i

            ' Face's thermal-optical properties, FTOP

            For i As Long = 0 To UBound(Whole_TOP_array) Step 1

                If Whole_TOP_array(i) = "FTOP" Or Whole_TOP_array(i) = "FTOP*" Then

                    i_F += 1

                    FTOP(i_F) = New FTOP(Copy_array_8(Whole_TOP_array, i, Max_items_default_2))

                End If

            Next i

            ' Element's thermal-optical properties, MEDIUM

            For i As Long = 0 To UBound(Whole_TOP_array) Step 1

                If Whole_TOP_array(i) = "MEDIUM1" Or Whole_TOP_array(i) = "MEDIUM1*" Then

                    i_MEDIUM += 1

                    MEDIUM(i_MEDIUM) = New MEDIUM1(Copy_array_8(Whole_TOP_array, i, Max_items_default_2))

                End If

            Next i

            'For i As Long = 0 To UBound(Whole_TOP_array) Step 1

            '    If Whole_TOP_array(i) = "MEDIUM9" Or Whole_TOP_array(i) = "MEDIUM9*" Then

            '        i_MEDIUM += 1

            '        MEDIUM(i_MEDIUM) = New MEDIUM9(Copy_array_8(Whole_TOP_array, i, Max_items_default_2))

            '    End If

            'Next i

            ' Radiation sources, RSOURCE1

            For i As Long = 0 To UBound(Whole_TOP_array) Step 1

                If Whole_TOP_array(i) = "RSOURCE1" Or Whole_TOP_array(i) = "RSOURCE1*" Then

                    i_RSOURCE += 1

                    RSOURCE(i_RSOURCE) = New RSOURCE1(Copy_array_8(Whole_TOP_array, i, Max_items_default_2))

                End If

            Next i

            ' Heat sources, HSOURCE1

            For i As Long = 0 To UBound(Whole_TOP_array) Step 1

                If Whole_TOP_array(i) = "HSOURCE1" Or Whole_TOP_array(i) = "HSOURCE1*" Then

                    i_HSOURCE += 1

                    HSOURCE(i_HSOURCE) = New HSOURCE1(Copy_array_8(Whole_TOP_array, i, Max_items_default_2))

                End If

            Next i

            ' Initial temperature field

            For i As Long = 0 To UBound(Whole_TOP_array) Step 1

                If Whole_TOP_array(i) = "T_INIT" Or Whole_TOP_array(i) = "T_INIT*" Then

                    i_T_INIT += 1

                    T_INIT(i_T_INIT) = New T_INIT(Copy_array_8(Whole_TOP_array, i, Max_items_default_2))

                End If

            Next i

            For i As Long = 0 To UBound(Whole_TOP_array) Step 1

                If Whole_TOP_array(i) = "T_ALL" Then

                    T_ALL = Val(Whole_TOP_array(i + 2))

                    Exit For

                End If

            Next i

            ' Temperature definiion during the calculation sequence

            For i As Long = 0 To UBound(Whole_TOP_array) Step 1

                If Whole_TOP_array(i) = "T_SET" Or Whole_TOP_array(i) = "T_SET*" Then

                    i_T_SET += 1

                    T_SET(i_T_SET) = New T_SET(Copy_array_8(Whole_TOP_array, i, Max_items_default_2))

                End If

            Next i

            ' Global medium

            For i As Long = 0 To UBound(Whole_TOP_array) Step 1

                If Whole_TOP_array(i) = "G_MEDIUM" Then

                    Global_medium = Val(Whole_TOP_array(i + 1))

                    Exit For

                End If

            Next i

            ' Numerical integration parameters

            'I_PARAM	I_TYPE	t_START	t_LENGTH	t_STEP	EPS/N_PH		


            For i As Long = 0 To UBound(Whole_TOP_array) Step 1

                If Whole_TOP_array(i) = "I_PARAM" Then

                    Solver_type = Whole_TOP_array(i + 1)

                    Start_time = Val(Whole_TOP_array(i + 2))

                    Lenght_time = Val(Whole_TOP_array(i + 3))

                    Step_time = Val(Whole_TOP_array(i + 4))

                    N_photon = Val(Whole_TOP_array(i + 5))

                    K_T_min = Val(Whole_TOP_array(i + 6))

                    K_T_max = Val(Whole_TOP_array(i + 7))

                    If i + 12 <= UBound(Whole_TOP_array) Then

                        If Whole_TOP_array(i + 10) = "" Then

                            Save_step_time = Val(Whole_TOP_array(i + 11))

                        End If

                    End If

                End If

            Next i


            'удаляем лишние элементы

            ReDim Preserve Dependency(i_D)

            ReDim Preserve FTOP(i_F)

            ReDim Preserve MEDIUM(i_MEDIUM)

            ReDim Preserve RSOURCE(i_RSOURCE)

            ReDim Preserve HSOURCE(i_HSOURCE)

            ReDim Preserve T_INIT(i_T_INIT)

            ReDim Preserve T_SET(i_T_SET)

            'заполянем поле N_photon источников излучения 


            For i As Integer = 0 To UBound(RSOURCE)

                If RSOURCE(i).N_photon = 0 Then

                    RSOURCE(i).N_photon = Me.N_photon

                End If

                'If RSOURCE(i).Epsilon = 0 Or RSOURCE(i).N_photon = 0 Then

                '    RSOURCE(i).Epsilon = Me.TOP_Epsilon

                '    RSOURCE(i).N_photon = Me.N_photon

                'End If

            Next i

            Create_TOP_links()

        End If

    End Sub

    ''' <summary>
    ''' This sub creates links in all of thermal-optical model part items 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Create_TOP_links()

        'Dim Grid_number, Element_property_string, Element_number, Material_number, Coordinate_system_number As String

        Dim tmp_Value As Long = -1

        Dim Element_collection As Collection = Convert_array_to_collection(Element)

        Dim Dependency_collection As Collection = Convert_array_to_collection(Dependency)

        Dim FTOP_collection As Collection = Convert_array_to_collection(FTOP)

        Dim MEDIUM_collection As Collection = Convert_array_to_collection(MEDIUM)

        Dim Material_collection As Collection = Convert_array_to_collection(Material)


        Is_thermooptics_constant = True

        Is_thermophysics_constant = True

        Is_boundary_conditions_constant = True

        Is_radiation_heat_sources_are_fixed = True


        ' Отмечаем термооптические свойства элементов
        For i As Integer = 0 To UBound(Element)

            ' отмечаемся в свойствах элемента

            Element(i).MEDIUM = MEDIUM_collection(Element(i).Element_property.Number.ToString)

            ' Сохраняем предковые свойства

            Element(i).MEDIUM.Parent = Element(i).Element_property

            ' Сохраняем предковый материал

            Element(i).MEDIUM.Parental_material = Element(i).Element_property.Material

            ' Отмечаем свойства граней

            For j As Integer = 0 To UBound(Element(i).Face)

                Element(i).Face(j).FTOP = FTOP_collection(Element(i).MEDIUM.FTOP(j).ToString)

            Next j

        Next i

        ' Отмечаем термооптические свойства граней и зависимости для тепловых и оптических коэффициентов в термооптических свойствах сред
        ' Учитываем, что при использовании материала MAT4 свойства берутся из него
        For i As Integer = 0 To UBound(MEDIUM)

            For j As Integer = 0 To UBound(MEDIUM(i).FTOP)

                MEDIUM(i).FTOP(j) = FTOP_collection(MEDIUM(i).FTOP(j).ToString)

                For l As Integer = 0 To UBound(MEDIUM(i).n)

                    ' если у нас стоит целое число, то это, несомненно, номер зависимости
                    If MEDIUM(i).n(l).GetType.ToString = "System.Int64" Then

                        MEDIUM(i).n(l) = Dependency_collection(MEDIUM(i).n(l).ToString)

                        Is_thermooptics_constant = False

                    End If

                    ' если у нас стоит целое число, то это, несомненно, номер зависимости
                    If MEDIUM(i).k(l).GetType.ToString = "System.Int64" Then

                        MEDIUM(i).k(l) = Dependency_collection(MEDIUM(i).k(l).ToString)

                        Is_thermooptics_constant = False

                    End If


                    ' если у нас стоит целое число, то это, несомненно, номер зависимости
                    If MEDIUM(i).lambda(l).GetType.ToString = "System.Int64" Then

                        MEDIUM(i).lambda(l) = Dependency_collection(MEDIUM(i).lambda(l).ToString)

                        Is_thermophysics_constant = False

                    End If

                Next l



                ' если у нас стоит целое число, то это, несомненно, номер зависимости
                If MEDIUM(i).c.GetType.ToString = "System.Int64" Then

                    MEDIUM(i).c = Dependency_collection(MEDIUM(i).c.ToString)

                    Is_thermophysics_constant = False

                End If

                If Not (MEDIUM(i).Parental_material Is Nothing) Then

                    If MEDIUM(i).Parental_material.Type = "MAT4" Then

                        ReDim MEDIUM(i).lambda(0)

                        MEDIUM(i).lambda(0) = MEDIUM(i).Parental_material.lambda

                        MEDIUM(i).c = MEDIUM(i).Parental_material.c

                    End If

                End If


            Next j

        Next i

        ' Заполняем свойства граней или просто числами или зависимостями
        For i As Integer = 0 To UBound(FTOP)

            ' если у нас стоит целое число, то это, несомненно, номер зависимости
            If FTOP(i).R_dif.GetType.ToString = "System.Int64" Then

                FTOP(i).R_dif = Dependency_collection(FTOP(i).R_dif.ToString)

                Is_thermooptics_constant = False

            End If

            ' если у нас стоит целое число, то это, несомненно, номер зависимости
            If FTOP(i).R_mir.GetType.ToString = "System.Int64" Then

                FTOP(i).R_mir = Dependency_collection(FTOP(i).R_mir.ToString)

                Is_thermooptics_constant = False

            End If

            ' если у нас стоит целое число, то это, несомненно, номер зависимости
            If FTOP(i).A.GetType.ToString = "System.Int64" Then

                FTOP(i).A = Dependency_collection(FTOP(i).A.ToString)

                Is_thermooptics_constant = False

            End If

            ' если у нас стоит целое число, то это, несомненно, номер зависимости
            If FTOP(i).D.GetType.ToString = "System.Int64" Then

                FTOP(i).D = Dependency_collection(FTOP(i).D.ToString)

                Is_thermooptics_constant = False

            End If

            ' если у нас стоит целое число, то это, несомненно, номер зависимости
            If FTOP(i).Eps.GetType.ToString = "System.Int64" Then

                If FTOP(i).Eps > 0 Then

                    FTOP(i).Eps = Dependency_collection(FTOP(i).Eps.ToString)

                    Is_thermooptics_constant = False

                End If



            End If

            ' если у нас стоит целое число, то это, несомненно, номер зависимости
            If FTOP(i).CDF.GetType.ToString = "System.Int64" Then

                If FTOP(i).CDF > 0 Then

                    FTOP(i).CDF = Dependency_collection(FTOP(i).CDF.ToString)

                    Is_thermooptics_constant = False

                Else

                    FTOP(i).CDF = Nothing

                End If

            End If


            ' если у нас стоит целое число, то это, несомненно, номер зависимости
            If FTOP(i).T_eff.GetType.ToString = "System.Int64" And FTOP(i).T_eff > 0 Then

                FTOP(i).T_eff = Dependency_collection(FTOP(i).T_eff.ToString)

                Is_thermooptics_constant = False


            End If

            If FTOP(i).T_eff = 0 Then

                FTOP(i).T_eff = Nothing

            End If

            If FTOP(i).Eps.GetType.ToString <> "THORIUM.DEPEND" Then

                If FTOP(i).Eps = 0 Then

                    FTOP(i).Eps = FTOP(i).A

                End If

            End If
        Next i


        ' Заполняем свойства источников излучения или просто числами или зависимостями
        For i As Integer = 0 To UBound(RSOURCE)

            ' если у нас стоит целое число, то это, несомненно, номер зависимости
            If RSOURCE(i).X.GetType.ToString = "System.Int64" Then

                Is_radiation_heat_sources_are_fixed = False

                RSOURCE(i).POS = Dependency_collection(RSOURCE(i).X.ToString)

                RSOURCE(i).X = Nothing

                RSOURCE(i).Y = Nothing

                RSOURCE(i).Z = Nothing

            Else

                RSOURCE(i).POS = Nothing

                RSOURCE(i).Current_position.Coord(0) = RSOURCE(i).X

                RSOURCE(i).Current_position.Coord(1) = RSOURCE(i).Y

                RSOURCE(i).Current_position.Coord(2) = RSOURCE(i).Z

            End If

            ' если у нас стоит целое число, то это, несомненно, номер зависимости
            If RSOURCE(i).Q.GetType.ToString = "System.Int64" Then

                RSOURCE(i).Q = Dependency_collection(RSOURCE(i).Q.ToString)

                Is_boundary_conditions_constant = False

            End If

            ' если у нас стоит целое число, то это, несомненно, номер зависимости
            If RSOURCE(i).T.GetType.ToString = "System.Int64" Then

                RSOURCE(i).T = Dependency_collection(RSOURCE(i).T.ToString)

                Is_boundary_conditions_constant = False

            Else

                RSOURCE(i).CDF = 0.0

            End If


            ' если у нас стоит целое число, то это, несомненно, номер зависимости
            If RSOURCE(i).CDF.GetType.ToString = "System.Int64" Then

                RSOURCE(i).CDF = Dependency_collection(RSOURCE(i).CDF.ToString)

                Is_boundary_conditions_constant = False

            Else

                RSOURCE(i).CDF = Nothing

            End If

        Next i


        ' Заполняем свойства источников излучения или просто числами или зависимостями
        For i As Integer = 0 To UBound(HSOURCE)

            ' если у нас стоит целое число, то это, несомненно, номер зависимости
            If HSOURCE(i).Q.GetType.ToString = "System.Int64" Then

                HSOURCE(i).Q = Dependency_collection(HSOURCE(i).Q.ToString)

                Is_boundary_conditions_constant = False

            End If

            ' связываем элементы и источники тепла

            For j As Integer = 0 To UBound(HSOURCE(i).Element)

                HSOURCE(i).Element(j) = Element_collection(HSOURCE(i).Element(j).ToString)

            Next j

        Next i

        ' Заполняем массив элементов в начальном распределении температур элементами
        For i As Integer = 0 To UBound(T_INIT)

            ' связываем элементы и начальное распределение температур

            For j As Integer = 0 To UBound(T_INIT(i).Element)

                T_INIT(i).Element(j) = Element_collection(T_INIT(i).Element(j).ToString)

            Next j

        Next i

        ' Заполняем массив элементов в заданном распределении температур 
        For i As Integer = 0 To UBound(T_SET)

            ' связываем элементы и начальное распределение температур

            For j As Integer = 0 To UBound(T_SET(i).Element)

                T_SET(i).Element(j) = Element_collection(T_SET(i).Element(j).ToString)

                ' проверяем, если у нас стоит целое число, то это, несомненно, номер зависимости
                ' температуры от времени

                ' если у нас стоит целое число, то это, несомненно, номер зависимости
                If T_SET(i).Temp(j).GetType.ToString = "System.Int64" Then

                    T_SET(i).Temp(j) = Dependency_collection(T_SET(i).Temp(j).ToString)

                    Is_boundary_conditions_constant = False

                End If

            Next j

        Next i

        ' Заполянем элемент "Глобальная среда"

        Global_medium = MEDIUM_collection(Global_medium.ToString)

        ' Заполянем поля глобальной среды для каждой грани

        For i As Integer = 0 To UBound(Face)

            Face(i).Global_medium = Global_medium

        Next i

        ' Заполянем поля глобальной среды для каждого источника излучения

        For i As Integer = 0 To UBound(RSOURCE)

            RSOURCE(i).Emitter_Medium = Global_medium

        Next i

    End Sub

End Class



