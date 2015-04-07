Public Class Matrix_inverser

    ' function of calculating of b matrix
    Public Shared Function Inverse(ByVal A(,) As Double) As Double(,)

        Dim M As Long = UBound(A, 1)                ' Upper boundary by first index
        Dim N As Long = UBound(A, 2)                ' Upper boundary by second index
        Dim MinSize As Long = Math.Min(M - 1, N - 1)  ' Long of permutation matrix
        Dim Pivots(MinSize) As Long                   ' permutation matrix
        Dim Result As Boolean


        ' LU-expansion of A matrix
        Pivots = RMatrixLU(A, M, N)

        ' Inverse of A matrix
        Result = RMatrixLUInverse(A, Pivots, N)


        ' Returning of inverse matrix
        Return A

    End Function

    Public Structure Complex
        Dim X As Double
        Dim Y As Double
    End Structure

    Public Const MachineEpsilon As Double = 0.0000000000000005
    Public Const MaxRealNumber As Double = 1.0E+300
    Public Const MinRealNumber As Double = 1.0E-300

    Private Const BigNumber As Double = 1.0E+70
    Private Const SmallNumber As Double = 1.0E-70
    Private Const PiNumber As Double = 3.14159265358979
    Public Function MaxReal(ByVal M1 As Double, ByVal M2 As Double) As Double
        If M1 > M2 Then
            MaxReal = M1
        Else
            MaxReal = M2
        End If
    End Function

    Public Function MinReal(ByVal M1 As Double, ByVal M2 As Double) As Double
        If M1 < M2 Then
            MinReal = M1
        Else
            MinReal = M2
        End If
    End Function

    Public Shared Function MaxInt(ByVal M1 As Integer, ByVal M2 As Integer) As Integer
        If M1 > M2 Then
            MaxInt = M1
        Else
            MaxInt = M2
        End If
    End Function

    Public Shared Function MinInt(ByVal M1 As Integer, ByVal M2 As Integer) As Integer
        If M1 < M2 Then
            MinInt = M1
        Else
            MinInt = M2
        End If
    End Function

    Public Function ArcSin(ByVal X As Double) As Double
        Dim T As Double
        T = System.Math.Sqrt(1 - X * X)
        If T < SmallNumber Then
            ArcSin = System.Math.Atan(BigNumber * System.Math.Sign(X))
        Else
            ArcSin = System.Math.Atan(X / T)
        End If
    End Function

    Public Function ArcCos(ByVal X As Double) As Double
        Dim T As Double
        T = System.Math.Sqrt(1 - X * X)
        If T < SmallNumber Then
            ArcCos = System.Math.Atan(BigNumber * System.Math.Sign(-X)) + 2 * System.Math.Atan(1)
        Else
            ArcCos = System.Math.Atan(-X / T) + 2 * System.Math.Atan(1)
        End If
    End Function

    Public Function SinH(ByVal X As Double) As Double
        SinH = (System.Math.Exp(X) - System.Math.Exp(-X)) / 2
    End Function

    Public Function CosH(ByVal X As Double) As Double
        CosH = (System.Math.Exp(X) + System.Math.Exp(-X)) / 2
    End Function

    Public Function TanH(ByVal X As Double) As Double
        TanH = (System.Math.Exp(X) - System.Math.Exp(-X)) / (System.Math.Exp(X) + System.Math.Exp(-X))
    End Function

    Public Function Pi() As Double
        Pi = PiNumber
    End Function

    Public Function Power(ByVal Base As Double, ByVal Exponent As Double) As Double
        Power = Base ^ Exponent
    End Function

    Public Function Square(ByVal X As Double) As Double
        Square = X * X
    End Function

    Public Function Log10(ByVal X As Double) As Double
        Log10 = System.Math.Log(X) / System.Math.Log(10)
    End Function

    Public Function Ceil(ByVal X As Double) As Double
        Ceil = -Int(-X)
    End Function

    Public Function RandomInteger(ByVal X As Integer) As Integer
        RandomInteger = Int(Rnd() * X)
    End Function

    Public Function Atn2(ByVal Y As Double, ByVal X As Double) As Double
        If SmallNumber * System.Math.Abs(Y) < System.Math.Abs(X) Then
            If X < 0 Then
                If Y = 0 Then
                    Atn2 = Pi()
                Else
                    Atn2 = System.Math.Atan(Y / X) + Pi() * System.Math.Sign(Y)
                End If
            Else
                Atn2 = System.Math.Atan(Y / X)
            End If
        Else
            Atn2 = System.Math.Sign(Y) * Pi() / 2
        End If
    End Function

    Public Function C_Complex(ByVal X As Double) As Complex
        Dim Result As Complex

        Result.X = X
        Result.Y = 0

        'UPGRADE_WARNING: Couldn't resolve default property of object C_Complex. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        C_Complex = Result
    End Function


    Public Function AbsComplex(ByRef Z As Complex) As Double
        Dim Result As Double
        Dim W As Double
        Dim XABS As Double
        Dim YABS As Double
        Dim V As Double

        XABS = System.Math.Abs(Z.X)
        YABS = System.Math.Abs(Z.Y)
        W = MaxReal(XABS, YABS)
        V = MinReal(XABS, YABS)
        If V = 0 Then
            Result = W
        Else
            Result = W * System.Math.Sqrt(1 + Square(V / W))
        End If

        AbsComplex = Result
    End Function


    Public Function C_Opposite(ByRef Z As Complex) As Complex
        Dim Result As Complex

        Result.X = -Z.X
        Result.Y = -Z.Y

        'UPGRADE_WARNING: Couldn't resolve default property of object C_Opposite. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        C_Opposite = Result
    End Function


    Public Function Conj(ByRef Z As Complex) As Complex
        Dim Result As Complex

        Result.X = Z.X
        Result.Y = -Z.Y

        'UPGRADE_WARNING: Couldn't resolve default property of object Conj. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        Conj = Result
    End Function


    Public Function CSqr(ByRef Z As Complex) As Complex
        Dim Result As Complex

        Result.X = Square(Z.X) - Square(Z.Y)
        Result.Y = 2 * Z.X * Z.Y

        'UPGRADE_WARNING: Couldn't resolve default property of object CSqr. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        CSqr = Result
    End Function


    Public Function C_Add(ByRef Z1 As Complex, ByRef Z2 As Complex) As Complex
        Dim Result As Complex

        Result.X = Z1.X + Z2.X
        Result.Y = Z1.Y + Z2.Y

        'UPGRADE_WARNING: Couldn't resolve default property of object C_Add. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        C_Add = Result
    End Function


    Public Function C_Mul(ByRef Z1 As Complex, ByRef Z2 As Complex) As Complex
        Dim Result As Complex

        Result.X = Z1.X * Z2.X - Z1.Y * Z2.Y
        Result.Y = Z1.X * Z2.Y + Z1.Y * Z2.X

        'UPGRADE_WARNING: Couldn't resolve default property of object C_Mul. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        C_Mul = Result
    End Function


    Public Function C_AddR(ByRef Z1 As Complex, ByVal R As Double) As Complex
        Dim Result As Complex

        Result.X = Z1.X + R
        Result.Y = Z1.Y

        'UPGRADE_WARNING: Couldn't resolve default property of object C_AddR. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        C_AddR = Result
    End Function


    Public Function C_MulR(ByRef Z1 As Complex, ByVal R As Double) As Complex
        Dim Result As Complex

        Result.X = Z1.X * R
        Result.Y = Z1.Y * R

        'UPGRADE_WARNING: Couldn't resolve default property of object C_MulR. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        C_MulR = Result
    End Function


    Public Function C_Sub(ByRef Z1 As Complex, ByRef Z2 As Complex) As Complex
        Dim Result As Complex

        Result.X = Z1.X - Z2.X
        Result.Y = Z1.Y - Z2.Y

        'UPGRADE_WARNING: Couldn't resolve default property of object C_Sub. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        C_Sub = Result
    End Function


    Public Function C_SubR(ByRef Z1 As Complex, ByVal R As Double) As Complex
        Dim Result As Complex

        Result.X = Z1.X - R
        Result.Y = Z1.Y

        'UPGRADE_WARNING: Couldn't resolve default property of object C_SubR. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        C_SubR = Result
    End Function


    Public Function C_RSub(ByVal R As Double, ByRef Z1 As Complex) As Complex
        Dim Result As Complex

        Result.X = R - Z1.X
        Result.Y = -Z1.Y

        'UPGRADE_WARNING: Couldn't resolve default property of object C_RSub. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        C_RSub = Result
    End Function


    Public Function C_Div(ByRef Z1 As Complex, ByRef Z2 As Complex) As Complex
        Dim Result As Complex
        Dim A As Double
        Dim B As Double
        Dim C As Double
        Dim D As Double
        Dim E As Double
        Dim F As Double

        A = Z1.X
        B = Z1.Y
        C = Z2.X
        D = Z2.Y
        If System.Math.Abs(D) < System.Math.Abs(C) Then
            E = D / C
            F = C + D * E
            Result.X = (A + B * E) / F
            Result.Y = (B - A * E) / F
        Else
            E = C / D
            F = D + C * E
            Result.X = (B + A * E) / F
            Result.Y = (-A + B * E) / F
        End If

        'UPGRADE_WARNING: Couldn't resolve default property of object C_Div. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        C_Div = Result
    End Function


    Public Function C_DivR(ByRef Z1 As Complex, ByVal R As Double) As Complex
        Dim Result As Complex

        Result.X = Z1.X / R
        Result.Y = Z1.Y / R

        'UPGRADE_WARNING: Couldn't resolve default property of object C_DivR. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        C_DivR = Result
    End Function


    Public Function C_RDiv(ByVal R As Double, ByRef Z2 As Complex) As Complex
        Dim Result As Complex
        Dim A As Double
        Dim C As Double
        Dim D As Double
        Dim E As Double
        Dim F As Double

        A = R
        C = Z2.X
        D = Z2.Y
        If System.Math.Abs(D) < System.Math.Abs(C) Then
            E = D / C
            F = C + D * E
            Result.X = A / F
            Result.Y = -(A * E / F)
        Else
            E = C / D
            F = D + C * E
            Result.X = A * E / F
            Result.Y = -(A / F)
        End If

        'UPGRADE_WARNING: Couldn't resolve default property of object C_RDiv. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
        C_RDiv = Result
    End Function


    Public Function C_Equal(ByRef Z1 As Complex, ByRef Z2 As Complex) As Boolean
        Dim Result As Boolean

        Result = Z1.X = Z2.X And Z1.Y = Z2.Y

        C_Equal = Result
    End Function


    Public Function C_NotEqual(ByRef Z1 As Complex, ByRef Z2 As Complex) As Boolean
        Dim Result As Boolean

        Result = Z1.X <> Z2.X Or Z1.Y <> Z2.Y

        C_NotEqual = Result
    End Function

    Public Function C_EqualR(ByRef Z1 As Complex, ByVal R As Double) As Boolean
        Dim Result As Boolean

        Result = Z1.X = R And Z1.Y = 0

        C_EqualR = Result
    End Function


    Public Function C_NotEqualR(ByRef Z1 As Complex, ByVal R As Double) As Boolean
        Dim Result As Boolean

        Result = Z1.X <> R Or Z1.Y <> 0

        C_NotEqualR = Result
    End Function

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Copyright (c) 1992-2007 The University of Tennessee. All rights reserved.
    '
    'Contributors:
    '    * Sergey Bochkanov (ALGLIB project). Translation from FORTRAN to
    '      pseudocode.
    '
    'See subroutines comments for additional copyrights.
    '
    'Redistribution and use in source and binary forms, with or without
    'modification, are permitted provided that the following conditions are
    'met:
    '
    '- Redistributions of source code must retain the above copyright
    '  notice, this list of conditions and the following disclaimer.
    '
    '- Redistributions in binary form must reproduce the above copyright
    '  notice, this list of conditions and the following disclaimer listed
    '  in this license in the documentation and/or other materials
    '  provided with the distribution.
    '
    '- Neither the name of the copyright holders nor the names of its
    '  contributors may be used to endorse or promote products derived from
    '  this software without specific prior written permission.
    '
    'THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
    '"AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
    'LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
    'A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT
    'OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
    'SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
    'LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
    'DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
    'THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
    '(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
    'OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Global constants
    Private Const LUNB As Long = 8.0#


    'Routines
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'LU-разложение матрицы общего вида размера M x N
    '
    'Подпрограмма вычисляет LU-разложение прямоугольной матрицы общего  вида  с
    'частичным выбором ведущего элемента (с перестановками строк).
    '
    'Входные параметры:
    '    A       -   матрица A. Нумерация элементов: [0..M-1, 0..N-1]
    '    M       -   число строк в матрице A
    '    N       -   число столбцов в матрице A
    '
    'Выходные параметры:
    '    A       -   матрицы L и U в компактной форме (см. ниже).
    '                Нумерация элементов: [0..M-1, 0..N-1]
    '    Pivots  -   матрица перестановок в компактной форме (см. ниже).
    '                Нумерация элементов: [0..Min(M-1,N-1)]
    '
    'Матрица A представляется, как A = P * L * U, где P - матрица перестановок,
    'матрица L - нижнетреугольная (или нижнетрапецоидальная, если M>N) матрица,
    'U - верхнетреугольная (или верхнетрапецоидальная, если M<N) матрица.
    '
    'Рассмотрим разложение более подробно на примере при M=4, N=3:
    '
    '                   (  1          )    ( U11 U12 U13  )
    'A = P1 * P2 * P3 * ( L21  1      )  * (     U22 U23  )
    '                   ( L31 L32  1  )    (         U33  )
    '                   ( L41 L42 L43 )
    '
    'Здесь матрица L  имеет  размер  M  x  Min(M,N),  матрица  U  имеет  размер
    'Min(M,N) x N, матрица  P(i)  получается  путем  перестановки  в  единичной
    'матрице размером M x M строк с номерами I и Pivots[I]
    '
    'Результатом работы алгоритма являются массив Pivots  и  следующая матрица,
    'замещающая  матрицу  A,  и  сохраняющая  в компактной форме матрицы L и U
    '(пример приведен для M=4, N=3):
    '
    ' ( U11 U12 U13 )
    ' ( L21 U22 U23 )
    ' ( L31 L32 U33 )
    ' ( L41 L42 L43 )
    '
    'Как видно, единичная диагональ матрицы L  не  сохраняется.
    'Если N>M, то соответственно меняются размеры матриц и расположение
    'элементов.
    '
    '  -- LAPACK routine (version 3.0) --
    '     Univ. of Tennessee, Univ. of California Berkeley, NAG Ltd.,
    '     Courant Institute, Argonne National Lab, and Rice University
    '     June 30, 1992
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Shared Function RMatrixLU(ByRef A(,) As Double, _
             ByVal M As Long, _
             ByVal N As Long)
        Dim Pivots() As Long
        Dim B(,) As Double
        Dim T() As Double
        Dim BP() As Long
        Dim MinMN As Long
        Dim I As Long
        Dim IP As Long
        Dim J As Long
        Dim J1 As Long
        Dim J2 As Long
        Dim CB As Long
        Dim NB As Long
        Dim V As Double
        Dim i_ As Long
        Dim i1_ As Long

        NB = LUNB

        '
        ' Decide what to use - blocked or unblocked code
        '
        If N < 1.0# Or MinInt(M, N) <= NB Or NB = 0 Then

            '
            ' Unblocked code
            '
            Pivots = RMatrixLU2(A, M, N)
        Else

            '
            ' Blocked code.
            ' First, prepare temporary matrix and indices
            '
            ReDim B(M - 1.0, NB - 1.0)
            ReDim T(0 To N - 1.0#)
            ReDim Pivots(0 To MinInt(M, N) - 1.0#)
            MinMN = MinInt(M, N)
            J1 = 0
            J2 = MinInt(MinMN, NB) - 1.0#

            '
            ' Main cycle
            '
            Do While J1 < MinMN
                CB = J2 - J1 + 1.0#

                '
                ' LU factorization of diagonal and subdiagonal blocks:
                ' 1. Copy columns J1..J2 of A to B
                ' 2. LU(B)
                ' 3. Copy result back to A
                ' 4. Copy pivots, apply pivots
                '
                For I = J1 To M - 1.0# Step 1
                    i1_ = (J1) - (0)
                    For i_ = 0 To CB - 1.0# Step 1
                        B(I - J1, i_) = A(I, i_ + i1_)
                    Next i_
                Next I
                BP = RMatrixLU2(B, M - J1, CB)
                For I = J1 To M - 1.0# Step 1
                    i1_ = (0) - (J1)
                    For i_ = J1 To J2 Step 1
                        A(I, i_) = B(I - J1, i_ + i1_)
                    Next i_
                Next I
                For I = 0 To CB - 1.0# Step 1
                    IP = BP(I)
                    Pivots(J1 + I) = J1 + IP
                    If BP(I) <> I Then
                        If J1 <> 0 Then

                            '
                            ' Interchange columns 0:J1-1
                            '
                            For i_ = 0 To J1 - 1.0# Step 1
                                T(i_) = A(J1 + I, i_)
                            Next i_
                            For i_ = 0 To J1 - 1.0# Step 1
                                A(J1 + I, i_) = A(J1 + IP, i_)
                            Next i_
                            For i_ = 0 To J1 - 1.0# Step 1
                                A(J1 + IP, i_) = T(i_)
                            Next i_
                        End If
                        If J2 < N - 1.0# Then

                            '
                            ' Interchange the rest of the matrix, if needed
                            '
                            For i_ = J2 + 1.0# To N - 1.0# Step 1
                                T(i_) = A(J1 + I, i_)
                            Next i_
                            For i_ = J2 + 1.0# To N - 1.0# Step 1
                                A(J1 + I, i_) = A(J1 + IP, i_)
                            Next i_
                            For i_ = J2 + 1.0# To N - 1.0# Step 1
                                A(J1 + IP, i_) = T(i_)
                            Next i_
                        End If
                    End If
                Next I

                '
                ' Compute block row of U
                '
                If J2 < N - 1.0# Then
                    For I = J1 + 1.0# To J2 Step 1
                        For J = J1 To I - 1.0# Step 1
                            V = A(I, J)
                            For i_ = J2 + 1.0# To N - 1.0# Step 1
                                A(I, i_) = A(I, i_) - V * A(J, i_)
                            Next i_
                        Next J
                    Next I
                End If

                '
                ' Update trailing submatrix
                '
                If J2 < N - 1.0# Then
                    For I = J2 + 1.0# To M - 1.0# Step 1
                        For J = J1 To J2 Step 1
                            V = A(I, J)
                            For i_ = J2 + 1.0# To N - 1.0# Step 1
                                A(I, i_) = A(I, i_) - V * A(J, i_)
                            Next i_
                        Next J
                    Next I
                End If

                '
                ' Next step
                '
                J1 = J2 + 1.0#
                J2 = MinInt(MinMN, J1 + NB) - 1.0#
            Loop
        End If

        Return Pivots

    End Function


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Obsolete 1-based subroutine. Left for backward compatibility.
    'See RMatrixLU for 0-based replacement.
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function LUDecomposition(ByRef A(,) As Double, _
             ByVal M As Long, _
             ByVal N As Long)
        Dim Pivots() As Long
        Dim I As Long
        Dim J As Long
        Dim JP As Long
        Dim T1() As Double
        Dim s As Double
        Dim i_ As Long

        ReDim Pivots(0 To MinInt(M, N))
        ReDim T1(0 To MaxInt(M, N))

        '
        ' Quick return if possible
        '
        If M = 0 Or N = 0 Then
            Return Nothing
        End If
        For J = 1.0# To MinInt(M, N) Step 1

            '
            ' Find pivot and test for singularity.
            '
            JP = J
            For I = J + 1.0# To M Step 1
                If Math.Abs(A(I, J)) > Math.Abs(A(JP, J)) Then
                    JP = I
                End If
            Next I
            Pivots(J) = JP
            If A(JP, J) <> 0 Then

                '
                'Apply the interchange to rows
                '
                If JP <> J Then
                    For i_ = 1.0# To N Step 1
                        T1(i_) = A(J, i_)
                    Next i_
                    For i_ = 1.0# To N Step 1
                        A(J, i_) = A(JP, i_)
                    Next i_
                    For i_ = 1.0# To N Step 1
                        A(JP, i_) = T1(i_)
                    Next i_
                End If

                '
                'Compute elements J+1:M of J-th column.
                '
                If J < M Then

                    '
                    ' CALL DSCAL( M-J, ONE / A( J, J ), A( J+1, J ), 1 )
                    '
                    JP = J + 1.0#
                    s = 1.0# / A(J, J)
                    For i_ = JP To M Step 1
                        A(i_, J) = s * A(i_, J)
                    Next i_
                End If
            End If
            If J < MinInt(M, N) Then

                '
                'Update trailing submatrix.
                'CALL DGER( M-J, N-J, -ONE, A( J+1, J ), 1, A( J, J+1 ), LDA,A( J+1, J+1 ), LDA )
                '
                JP = J + 1.0#
                For I = J + 1.0# To M Step 1
                    s = A(I, J)
                    For i_ = JP To N Step 1
                        A(I, i_) = A(I, i_) - s * A(J, i_)
                    Next i_
                Next I
            End If
        Next J

        Return Pivots
    End Function


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Obsolete 1-based subroutine. Left for backward compatibility.
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Sub LUDecompositionUnpacked(ByRef A_(,) As Double, _
             ByVal M As Long, _
             ByVal N As Long, _
             ByRef L(,) As Double, _
             ByRef U(,) As Double, _
             ByRef Pivots() As Long)
        Dim A(,) As Double
        Dim I As Long
        Dim J As Long
        Dim MinMN As Long
        A = A_

        If M = 0 Or N = 0 Then
            Exit Sub
        End If
        MinMN = MinInt(M, N)
        ReDim L(0 To M, 0 To MinMN)
        ReDim U(0 To MinMN, 0 To N)
        Pivots = LUDecomposition(A, M, N)
        For I = 1.0# To M Step 1
            For J = 1.0# To MinMN Step 1
                If J > I Then
                    L(I, J) = 0
                End If
                If J = I Then
                    L(I, J) = 1.0#
                End If
                If J < I Then
                    L(I, J) = A(I, J)
                End If
            Next J
        Next I
        For I = 1.0# To MinMN Step 1
            For J = 1.0# To N Step 1
                If J < I Then
                    U(I, J) = 0
                End If
                If J >= I Then
                    U(I, J) = A(I, J)
                End If
            Next J
        Next I
    End Sub


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Level 2 BLAS version of RMatrixLU
    '
    '  -- LAPACK routine (version 3.0) --
    '     Univ. of Tennessee, Univ. of California Berkeley, NAG Ltd.,
    '     Courant Institute, Argonne National Lab, and Rice University
    '     June 30, 1992
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Shared Function RMatrixLU2(ByRef A(,) As Double, _
             ByVal M As Long, _
             ByVal N As Long)
        Dim Pivots() As Long
        Dim I As Long
        Dim J As Long
        Dim JP As Long
        Dim T1() As Double
        Dim s As Double
        Dim i_ As Long

        ReDim Pivots(0 To MinInt(M - 1.0#, N - 1.0#))
        ReDim T1(0 To MaxInt(M - 1.0#, N - 1.0#))

        '
        ' Quick return if possible
        '
        If M = 0 Or N = 0 Then
            Return Nothing
        End If
        For J = 0 To MinInt(M - 1.0#, N - 1.0#) Step 1

            '
            ' Find pivot and test for singularity.
            '
            JP = J
            For I = J + 1.0# To M - 1.0# Step 1
                If Math.Abs(A(I, J)) > Math.Abs(A(JP, J)) Then
                    JP = I
                End If
            Next I
            Pivots(J) = JP
            If A(JP, J) <> 0 Then

                '
                'Apply the interchange to rows
                '
                If JP <> J Then
                    For i_ = 0 To N - 1.0# Step 1
                        T1(i_) = A(J, i_)
                    Next i_
                    For i_ = 0 To N - 1.0# Step 1
                        A(J, i_) = A(JP, i_)
                    Next i_
                    For i_ = 0 To N - 1.0# Step 1
                        A(JP, i_) = T1(i_)
                    Next i_
                End If

                '
                'Compute elements J+1:M of J-th column.
                '
                If J < M Then
                    JP = J + 1.0#
                    s = 1.0# / A(J, J)
                    For i_ = JP To M - 1.0# Step 1
                        A(i_, J) = s * A(i_, J)
                    Next i_
                End If
            End If
            If J < MinInt(M, N) - 1.0# Then

                '
                'Update trailing submatrix.
                '
                JP = J + 1.0#
                For I = J + 1.0# To M - 1.0# Step 1
                    s = A(I, J)
                    For i_ = JP To N - 1.0# Step 1
                        A(I, i_) = A(I, i_) - s * A(J, i_)
                    Next i_
                Next I
            End If
        Next J

        Return Nothing

    End Function

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Copyright (c) 1992-2007 The University of Tennessee. All rights reserved.
    '
    'Contributors:
    '    * Sergey Bochkanov (ALGLIB project). Translation from FORTRAN to
    '      pseudocode.
    '
    'See subroutines comments for additional copyrights.
    '
    'Redistribution and use in source and binary forms, with or without
    'modification, are permitted provided that the following conditions are
    'met:
    '
    '- Redistributions of source code must retain the above copyright
    '  notice, this list of conditions and the following disclaimer.
    '
    '- Redistributions in binary form must reproduce the above copyright
    '  notice, this list of conditions and the following disclaimer listed
    '  in this license in the documentation and/or other materials
    '  provided with the distribution.
    '
    '- Neither the name of the copyright holders nor the names of its
    '  contributors may be used to endorse or promote products derived from
    '  this software without specific prior written permission.
    '
    'THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
    '"AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
    'LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
    'A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT
    'OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
    'SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
    'LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
    'DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
    'THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
    '(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
    'OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Routines
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Обращение матрицы, заданной LU-разложением
    '
    'Входные параметры:
    '    A       -   LU-разложение  матрицы   (результат   работы  подпрограммы
    '                RMatrixLU).
    '    Pivots  -   таблица перестановок,  произведенных в ходе LU-разложения.
    '                (результат работы подпрограммы RMatrixLU).
    '    N       -   размерность матрицы
    '
    'Выходные параметры:
    '    A       -   матрица, обратная к исходной. Массив с нумерацией
    '                элементов [0..N-1, 0..N-1]
    '
    'Результат:
    '    True,  если исходная матрица невырожденная.
    '    False, если исходная матрица вырожденная.
    '
    '  -- LAPACK routine (version 3.0) --
    '     Univ. of Tennessee, Univ. of California Berkeley, NAG Ltd.,
    '     Courant Institute, Argonne National Lab, and Rice University
    '     February 29, 1992
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Shared Function RMatrixLUInverse(ByRef A(,) As Double, _
             ByRef Pivots() As Long, _
             ByVal N As Long) As Boolean
        Dim Result As Boolean
        Dim WORK() As Double
        Dim I As Long
        'Dim IWS As Long
        Dim J As Long
        'Dim JB As Long
        'Dim JJ As Long
        Dim JP As Long
        Dim V As Double
        Dim i_ As Long

        Result = True

        '
        ' Quick return if possible
        '
        If N = 0.0# Then
            RMatrixLUInverse = Result
            Exit Function
        End If
        ReDim WORK(0 To N - 1.0#)

        '
        ' Form inv(U)
        '
        If Not RMatrixTRInverse(A, N, True, False) Then
            Result = False
            RMatrixLUInverse = Result
            Exit Function
        End If

        '
        ' Solve the equation inv(A)*L = inv(U) for inv(A).
        '
        For J = N - 1.0# To 0.0# Step -1

            '
            ' Copy current column of L to WORK and replace with zeros.
            '
            For I = J + 1.0# To N - 1.0# Step 1
                WORK(I) = A(I, J)
                A(I, J) = 0.0#
            Next I

            '
            ' Compute current column of inv(A).
            '
            If J < N - 1.0# Then
                For I = 0.0# To N - 1.0# Step 1
                    V = 0.0
                    For i_ = J + 1.0# To N - 1.0# Step 1
                        V = V + A(I, i_) * WORK(i_)
                    Next i_
                    A(I, J) = A(I, J) - V
                Next I
            End If
        Next J

        '
        ' Apply column interchanges.
        '
        For J = N - 2.0# To 0.0# Step -1
            JP = Pivots(J)
            If JP <> J Then
                For i_ = 0.0# To N - 1.0# Step 1
                    WORK(i_) = A(i_, J)
                Next i_
                For i_ = 0.0# To N - 1.0# Step 1
                    A(i_, J) = A(i_, JP)
                Next i_
                For i_ = 0.0# To N - 1.0# Step 1
                    A(i_, JP) = WORK(i_)
                Next i_
            End If
        Next J

        RMatrixLUInverse = Result
    End Function


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Обращение матрицы общего вида
    '
    'Входные параметры:
    '    A   -   матрица. массив с нумерацией элементов [0..N-1, 0..N-1]
    '    N   -   размерность матрицы A
    '
    'Выходные параметры:
    '    A   -   матрица, обратная к исходной. Массив с нумерацией элементов
    '            [0..N-1, 0..N-1]
    '
    'Результат:
    '    True,  если исходная матрица невырожденная.
    '    False, если исходная матрица вырожденная.
    '
    '  -- ALGLIB --
    '     Copyright 2005 by Bochkanov Sergey
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function RMatrixInverse(ByRef A(,) As Double, ByVal N As Long) As Boolean
        Dim Result As Boolean
        Dim Pivots() As Long

        Pivots = RMatrixLU(A, N, N)
        Result = RMatrixLUInverse(A, Pivots, N)

        RMatrixInverse = Result
    End Function


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Obsolete 1-based subroutine.
    '
    'See RMatrixLUInverse for 0-based replacement.
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function InverseLU(ByRef A(,) As Double, _
             ByRef Pivots() As Long, _
             ByVal N As Long) As Boolean
        Dim Result As Boolean
        Dim WORK() As Double
        Dim I As Long
        'Dim IWS As Long
        Dim J As Long
        'Dim JB As Long
        'Dim JJ As Long
        Dim JP As Long
        Dim JP1 As Long
        Dim V As Double
        Dim i_ As Long

        Result = True

        '
        ' Quick return if possible
        '
        If N = 0.0# Then
            InverseLU = Result
            Exit Function
        End If
        ReDim WORK(N)

        '
        ' Form inv(U)
        '
        If Not InvTriangular(A, N, True, False) Then
            Result = False
            InverseLU = Result
            Exit Function
        End If

        '
        ' Solve the equation inv(A)*L = inv(U) for inv(A).
        '
        For J = N To 1.0# Step -1

            '
            ' Copy current column of L to WORK and replace with zeros.
            '
            For I = J + 1.0# To N Step 1
                WORK(I) = A(I, J)
                A(I, J) = 0.0#
            Next I

            '
            ' Compute current column of inv(A).
            '
            If J < N Then
                JP1 = J + 1.0#
                For I = 1.0# To N Step 1
                    V = 0.0
                    For i_ = JP1 To N Step 1
                        V = V + A(I, i_) * WORK(i_)
                    Next i_
                    A(I, J) = A(I, J) - V
                Next I
            End If
        Next J

        '
        ' Apply column interchanges.
        '
        For J = N - 1.0# To 1.0# Step -1
            JP = Pivots(J)
            If JP <> J Then
                For i_ = 1.0# To N Step 1
                    WORK(i_) = A(i_, J)
                Next i_
                For i_ = 1.0# To N Step 1
                    A(i_, J) = A(i_, JP)
                Next i_
                For i_ = 1.0# To N Step 1
                    A(i_, JP) = WORK(i_)
                Next i_
            End If
        Next J

        InverseLU = Result
    End Function


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Obsolete 1-based subroutine.
    '
    'See RMatrixInverse for 0-based replacement.
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function Inverse(ByRef A(,) As Double, ByVal N As Long) As Boolean
        Dim Result As Boolean
        Dim Pivots() As Long

        Pivots = LUDecomposition(A, N, N)
        Result = InverseLU(A, Pivots, N)

        Inverse = Result
    End Function


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Copyright (c) 1992-2007 The University of Tennessee.  All rights reserved.
    '
    'Contributors:
    '    * Sergey Bochkanov (ALGLIB project). Translation from FORTRAN to
    '      pseudocode.
    '
    'See subroutines comments for additional copyrights.
    '
    'Redistribution and use in source and binary forms, with or without
    'modification, are permitted provided that the following conditions are
    'met:
    '
    '- Redistributions of source code must retain the above copyright
    '  notice, this list of conditions and the following disclaimer.
    '
    '- Redistributions in binary form must reproduce the above copyright
    '  notice, this list of conditions and the following disclaimer listed
    '  in this license in the documentation and/or other materials
    '  provided with the distribution.
    '
    '- Neither the name of the copyright holders nor the names of its
    '  contributors may be used to endorse or promote products derived from
    '  this software without specific prior written permission.
    '
    'THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
    '"AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
    'LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
    'A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT
    'OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
    'SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
    'LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
    'DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
    'THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
    '(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
    'OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Routines
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Обращение треугольной матрицы
    '
    'Подпрограмма обращает следующие типы матриц:
    '    * верхнетреугольные
    '    * верхнетреугольные с единичной диагональю
    '    * нижнетреугольные
    '    * нижнетреугольные с единичной диагональю
    '
    'В случае, если матрица верхне(нижне)треугольная, то  матрица,  обратная  к
    'ней, тоже верхне(нижне)треугольная, и после  завершения  работы  алгоритма
    'обратная матрица замещает переданную. При этом элементы расположенные ниже
    '(выше) диагонали не меняются в ходе работы алгоритма.
    '
    'Если матрица с единичной  диагональю, то обратная к  ней  матрица  тоже  с
    'единичной  диагональю.  В  алгоритм  передаются  только    внедиагональные
    'элементы. При этом в результате работы алгоритма диагональные элементы  не
    'меняются.
    '
    'Входные параметры:
    '    A           -   матрица. Массив с нумерацией элементов [0..N-1,0..N-1]
    '    N           -   размер матрицы
    '    IsUpper     -   True, если матрица верхнетреугольная
    '    IsUnitTriangular-   True, если матрица с единичной диагональю.
    '
    'Выходные параметры:
    '    A           -   матрица, обратная к входной, если задача не вырождена.
    '
    'Результат:
    '    True, если матрица не вырождена
    '    False, если матрица вырождена
    '
    '  -- LAPACK routine (version 3.0) --
    '     Univ. of Tennessee, Univ. of California Berkeley, NAG Ltd.,
    '     Courant Institute, Argonne National Lab, and Rice University
    '     February 29, 1992
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Shared Function RMatrixTRInverse(ByRef A(,) As Double, _
             ByVal N As Long, _
             ByVal IsUpper As Boolean, _
             ByVal IsUnitTriangular As Boolean) As Boolean
        Dim Result As Boolean
        Dim NOUNIT As Boolean
        Dim I As Long
        Dim J As Long
        Dim V As Double
        Dim AJJ As Double
        Dim T() As Double
        Dim i_ As Long

        Result = True
        ReDim T(0 To N - 1.0#)

        '
        ' Test the input parameters.
        '
        NOUNIT = Not IsUnitTriangular
        If IsUpper Then

            '
            ' Compute inverse of upper triangular matrix.
            '
            For J = 0.0# To N - 1.0# Step 1
                If NOUNIT Then
                    If A(J, J) = 0.0# Then
                        Result = False
                        RMatrixTRInverse = Result
                        Exit Function
                    End If
                    A(J, J) = 1.0# / A(J, J)
                    AJJ = -A(J, J)
                Else
                    AJJ = -1.0#
                End If

                '
                ' Compute elements 1:j-1 of j-th column.
                '
                If J > 0.0# Then
                    For i_ = 0.0# To J - 1.0# Step 1
                        T(i_) = A(i_, J)
                    Next i_
                    For I = 0.0# To J - 1.0# Step 1
                        If I < J - 1.0# Then
                            V = 0.0
                            For i_ = I + 1.0# To J - 1.0# Step 1
                                V = V + A(I, i_) * T(i_)
                            Next i_
                        Else
                            V = 0.0#
                        End If
                        If NOUNIT Then
                            A(I, J) = V + A(I, I) * T(I)
                        Else
                            A(I, J) = V + T(I)
                        End If
                    Next I
                    For i_ = 0.0# To J - 1.0# Step 1
                        A(i_, J) = AJJ * A(i_, J)
                    Next i_
                End If
            Next J
        Else

            '
            ' Compute inverse of lower triangular matrix.
            '
            For J = N - 1.0# To 0.0# Step -1
                If NOUNIT Then
                    If A(J, J) = 0.0# Then
                        Result = False
                        RMatrixTRInverse = Result
                        Exit Function
                    End If
                    A(J, J) = 1.0# / A(J, J)
                    AJJ = -A(J, J)
                Else
                    AJJ = -1.0#
                End If
                If J < N - 1.0# Then

                    '
                    ' Compute elements j+1:n of j-th column.
                    '
                    For i_ = J + 1.0# To N - 1.0# Step 1
                        T(i_) = A(i_, J)
                    Next i_
                    For I = J + 1.0# To N - 1.0# Step 1
                        If I > J + 1.0# Then
                            V = 0.0
                            For i_ = J + 1.0# To I - 1.0# Step 1
                                V = V + A(I, i_) * T(i_)
                            Next i_
                        Else
                            V = 0.0#
                        End If
                        If NOUNIT Then
                            A(I, J) = V + A(I, I) * T(I)
                        Else
                            A(I, J) = V + T(I)
                        End If
                    Next I
                    For i_ = J + 1.0# To N - 1.0# Step 1
                        A(i_, J) = AJJ * A(i_, J)
                    Next i_
                End If
            Next J
        End If

        RMatrixTRInverse = Result
    End Function


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Obsolete 1-based subroutine.
    'See RMatrixTRInverse for 0-based replacement.
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function InvTriangular(ByRef A(,) As Double, _
             ByVal N As Long, _
             ByVal IsUpper As Boolean, _
             ByVal IsUnitTriangular As Boolean) As Boolean
        Dim Result As Boolean
        Dim NOUNIT As Boolean
        Dim I As Long
        Dim J As Long
        Dim NMJ As Long
        Dim JM1 As Long
        Dim JP1 As Long
        Dim V As Double
        Dim AJJ As Double
        Dim T() As Double
        Dim i_ As Long

        Result = True
        ReDim T(N)

        '
        ' Test the input parameters.
        '
        NOUNIT = Not IsUnitTriangular
        If IsUpper Then

            '
            ' Compute inverse of upper triangular matrix.
            '
            For J = 1.0# To N Step 1
                If NOUNIT Then
                    If A(J, J) = 0.0# Then
                        Result = False
                        InvTriangular = Result
                        Exit Function
                    End If
                    A(J, J) = 1.0# / A(J, J)
                    AJJ = -A(J, J)
                Else
                    AJJ = -1.0#
                End If

                '
                ' Compute elements 1:j-1 of j-th column.
                '
                If J > 1.0# Then
                    JM1 = J - 1.0#
                    For i_ = 1.0# To JM1 Step 1
                        T(i_) = A(i_, J)
                    Next i_
                    For I = 1.0# To J - 1.0# Step 1
                        If I < J - 1.0# Then
                            V = 0.0
                            For i_ = I + 1.0# To JM1 Step 1
                                V = V + A(I, i_) * T(i_)
                            Next i_
                        Else
                            V = 0.0#
                        End If
                        If NOUNIT Then
                            A(I, J) = V + A(I, I) * T(I)
                        Else
                            A(I, J) = V + T(I)
                        End If
                    Next I
                    For i_ = 1.0# To JM1 Step 1
                        A(i_, J) = AJJ * A(i_, J)
                    Next i_
                End If
            Next J
        Else

            '
            ' Compute inverse of lower triangular matrix.
            '
            For J = N To 1.0# Step -1
                If NOUNIT Then
                    If A(J, J) = 0.0# Then
                        Result = False
                        InvTriangular = Result
                        Exit Function
                    End If
                    A(J, J) = 1.0# / A(J, J)
                    AJJ = -A(J, J)
                Else
                    AJJ = -1.0#
                End If
                If J < N Then

                    '
                    ' Compute elements j+1:n of j-th column.
                    '
                    NMJ = N - J
                    JP1 = J + 1.0#
                    For i_ = JP1 To N Step 1
                        T(i_) = A(i_, J)
                    Next i_
                    For I = J + 1.0# To N Step 1
                        If I > J + 1.0# Then
                            V = 0.0
                            For i_ = JP1 To I - 1.0# Step 1
                                V = V + A(I, i_) * T(i_)
                            Next i_
                        Else
                            V = 0.0#
                        End If
                        If NOUNIT Then
                            A(I, J) = V + A(I, I) * T(I)
                        Else
                            A(I, J) = V + T(I)
                        End If
                    Next I
                    For i_ = JP1 To N Step 1
                        A(i_, J) = AJJ * A(i_, J)
                    Next i_
                End If
            Next J
        End If

        InvTriangular = Result
    End Function

End Class
