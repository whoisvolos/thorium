''' <summary>
''' Mass-inertia properties calculation class
''' </summary>
''' <remarks></remarks>
''' 
<Serializable()> _
Public Class MIP

    Inherits Tetrahedron

    ''' <summary>
    ''' This sub return tensor of inertia for 2-noded element. i.e. CROD, CBAR or CBEAM in element nodes coordinate system, 
    ''' and its volume and centroid.
    ''' </summary>
    ''' <param name="Vertex"></param>
    ''' Array of element nodes coordinates. It must be in one coordinate system 
    ''' <remarks></remarks>
    ''' 
    Function Inertia_2_node(ByRef Vertex() As Vector, ByRef Area As Double, ByRef Volume As Double, ByRef Centroid As Vector) As Double(,)

        Dim Inertia_tensor(,) As Double

        ReDim Inertia_tensor(2, 2)

        Dim ox, oy, oz, Glob_x As New Vector

        Dim Vertex_1(7) As Vector



        Dim R As Double

        R = Math.Sqrt(Area / 2)

        Glob_x.Coord(0) = 1

        Glob_x.Coord(1) = 0

        Glob_x.Coord(2) = 0

        ox = Vertex(1) - Vertex(0)

        If Not (ox Like Glob_x) Then

            oz = ox ^ Glob_x

        Else

            Dim Glob_y = New Vector

            Glob_y.Coord(0) = 0

            Glob_y.Coord(1) = 1

            Glob_y.Coord(2) = 0

            oz = Glob_x ^ Glob_y

        End If


        oz = Norm_vector(oz)

        oz = R * oz


        oy = ox ^ oz

        oy = Norm_vector(oy)

        oy = R * oy


        Vertex_1(0) = Vertex(0) + oz

        Vertex_1(1) = Vertex(0) + oy

        Vertex_1(2) = Vertex(0) - oz

        Vertex_1(3) = Vertex(0) - oy


        Vertex_1(4) = Vertex(1) + oz

        Vertex_1(5) = Vertex(1) + oy

        Vertex_1(6) = Vertex(1) - oz

        Vertex_1(7) = Vertex(1) - oy

        Inertia_tensor = Inertia_8_node(Vertex_1, Volume, Centroid)

        Return Inertia_tensor

    End Function

    ''' <summary>
    ''' This sub return tensor of inertia for 4-noded element. i.e. CTETRA in element nodes coordinate system, 
    ''' and its volume and centroid.
    ''' </summary>
    ''' <param name="Vertex"></param>
    ''' Array of element nodes coordinates. It must be in one coordinate system 
    ''' <remarks></remarks>
    Sub Inertia_4_node(ByRef Vertex() As Vector, ByRef Volume As Double, ByRef Centroid As Vector, ByRef Inertia_tensor(,) As Double)

        Dim Tetra As New Tetrahedron(Vertex)

        ReDim Inertia_tensor(2, 2)

        Volume = Tetra.Volume

        Centroid = Tetra.Centroid

        Inertia_tensor = Tetra.Inertia_tensor


    End Sub


    ''' <summary>
    ''' This sub return tensor of inertia for 6-noded element. i.e. CPENTA or CTRIA in element nodes coordinate system, 
    ''' and its volume and centroid.
    ''' </summary>
    ''' <param name="Vertex"></param>
    ''' Array of element nodes coordinates. It must be in one coordinate system 
    ''' <remarks></remarks>
    Function Inertia_6_node(ByRef Vertex() As Vector, ByRef Volume As Double, ByRef Centroid As Vector) As Double(,)

        Dim Tetra(2) As Tetrahedron

        Dim tmp_Vertex(3) As Vector

        Dim Inertia_tensor(,) As Double

        ReDim Inertia_tensor(2, 2)


        tmp_Vertex(0) = Vertex(0)

        tmp_Vertex(1) = Vertex(1)

        tmp_Vertex(2) = Vertex(2)

        tmp_Vertex(3) = Vertex(3)

        Tetra(0) = New Tetrahedron(tmp_Vertex)

        tmp_Vertex(0) = Vertex(1)

        tmp_Vertex(1) = Vertex(2)

        tmp_Vertex(2) = Vertex(3)

        tmp_Vertex(3) = Vertex(5)

        Tetra(1) = New Tetrahedron(tmp_Vertex)

        tmp_Vertex(0) = Vertex(1)

        tmp_Vertex(1) = Vertex(3)

        tmp_Vertex(2) = Vertex(4)

        tmp_Vertex(3) = Vertex(5)

        Tetra(2) = New Tetrahedron(tmp_Vertex)


        Volume = 0

        For i As Integer = 0 To 2

            Volume += Tetra(i).Volume

        Next i


        Centroid = New Vector

        For i As Integer = 0 To 2

            Centroid += Tetra(i).Volume * Tetra(i).Centroid

        Next i

        Centroid = 1 / Volume * Centroid


        ReDim Inertia_tensor(2, 2)

        For i As Integer = 0 To 2

            For m As Integer = 0 To 2

                For n As Integer = 0 To 2

                    Inertia_tensor(m, n) += Tetra(i).Inertia_tensor(m, n)

                Next n

            Next m

        Next i

        Return Inertia_tensor

    End Function


    ''' <summary>
    ''' This function return tensor of inertia for 8-noded element. i.e. CHEXA or CQUAD in element nodes coordinate system, 
    ''' and its volume and centroid.
    ''' </summary>
    ''' <param name="Vertex"></param>
    ''' Array of element nodes coordinates. It must be in one coordinate system 
    ''' <remarks></remarks>
    Function Inertia_8_node(ByRef Vertex() As Vector, ByRef Volume As Double, ByRef Centroid As Vector) As Double(,)

        Dim Tetra(5) As Tetrahedron

        Dim tmp_Vertex(3) As Vector

        Dim Inertia_tensor(,) As Double

        ReDim Inertia_tensor(2, 2)


        tmp_Vertex(0) = Vertex(0)

        tmp_Vertex(1) = Vertex(2)

        tmp_Vertex(2) = Vertex(7)

        tmp_Vertex(3) = Vertex(3)

        Tetra(0) = New Tetrahedron(tmp_Vertex)

        tmp_Vertex(0) = Vertex(0)

        tmp_Vertex(1) = Vertex(5)

        tmp_Vertex(2) = Vertex(2)

        tmp_Vertex(3) = Vertex(1)

        Tetra(1) = New Tetrahedron(tmp_Vertex)

        tmp_Vertex(0) = Vertex(0)

        tmp_Vertex(1) = Vertex(4)

        tmp_Vertex(2) = Vertex(7)

        tmp_Vertex(3) = Vertex(2)

        Tetra(2) = New Tetrahedron(tmp_Vertex)


        tmp_Vertex(0) = Vertex(0)

        tmp_Vertex(1) = Vertex(5)

        tmp_Vertex(2) = Vertex(4)

        tmp_Vertex(3) = Vertex(2)

        Tetra(3) = New Tetrahedron(tmp_Vertex)

        tmp_Vertex(0) = Vertex(4)

        tmp_Vertex(1) = Vertex(6)

        tmp_Vertex(2) = Vertex(7)

        tmp_Vertex(3) = Vertex(2)

        Tetra(4) = New Tetrahedron(tmp_Vertex)

        tmp_Vertex(0) = Vertex(4)

        tmp_Vertex(1) = Vertex(5)

        tmp_Vertex(2) = Vertex(6)

        tmp_Vertex(3) = Vertex(2)

        Tetra(5) = New Tetrahedron(tmp_Vertex)


        Volume = 0

        For i As Integer = 0 To 5

            Volume += Tetra(i).Volume

        Next i


        Centroid = New Vector

        For i As Integer = 0 To 5

            Centroid += Tetra(i).Volume * Tetra(i).Centroid

        Next i

        Centroid = 1 / Volume * Centroid


        ReDim Inertia_tensor(2, 2)

        For i As Integer = 0 To 5

            For m As Integer = 0 To 2

                For n As Integer = 0 To 2

                    Inertia_tensor(m, n) += Tetra(i).Inertia_tensor(m, n)

                Next n

            Next m

        Next i

        Return Inertia_tensor

    End Function



End Class

<Serializable()> _
Public Class Tetrahedron

    Inherits Linear_algebra

    ''' <summary>
    ''' Array of Tetrahedron vertexes
    ''' </summary>
    ''' <remarks></remarks>
    ''' 
    Public Vertex() As Vector

    Public Volume As Double

    Public Centroid As New Vector

    Public Inertia_tensor(,) As Double

    Public Sub New(Optional ByRef Vertex() As Vector = Nothing)

        If Vertex Is Nothing Then Exit Sub

        ' Preliminary definitions

        Dim x_1, x_2, x_3, x_4 As Double

        Dim y_1, y_2, y_3, y_4 As Double

        Dim z_1, z_2, z_3, z_4 As Double


        x_1 = Vertex(0).Coord(0)

        y_1 = Vertex(0).Coord(1)

        z_1 = Vertex(0).Coord(2)



        x_2 = Vertex(1).Coord(0)

        y_2 = Vertex(1).Coord(1)

        z_2 = Vertex(1).Coord(2)



        x_3 = Vertex(2).Coord(0)

        y_3 = Vertex(2).Coord(1)

        z_3 = Vertex(2).Coord(2)



        x_4 = Vertex(3).Coord(0)

        y_4 = Vertex(3).Coord(1)

        z_4 = Vertex(3).Coord(2)





        ' Volume calculation

        ' Determinant of Jacobian

        Volume = Math.Abs(((x_2 - x_1) * ((y_3 - y_1) * (z_4 - z_1) - (y_4 - y_1) * (z_3 - z_1)) _
        - (x_3 - x_1) * ((y_2 - y_1) * (z_4 - z_1) - (y_4 - y_1) * (z_2 - z_1)) _
        + (x_4 - x_1) * ((y_2 - y_1) * (z_3 - z_1) - (y_3 - y_1) * (z_2 - z_1))) / 6)


        ' Coordinate center calculation

        For i As Integer = 0 To 2

            For j As Integer = 0 To UBound(Vertex)

                Centroid.Coord(i) += Vertex(j).Coord(i)

            Next j

            Centroid.Coord(i) *= 0.25

        Next i

        Dim a, b, c, a_, b_, c_ As Double

        a = Volume / 10 * (y_1 ^ 2 + y_1 * y_2 + y_2 ^ 2 + y_1 * y_3 + y_2 * y_3 + y_3 ^ 2 + y_1 * y_4 + y_2 * y_4 + y_3 * y_4 + y_4 ^ 2 + _
        z_1 ^ 2 + z_1 * z_2 + z_2 ^ 2 + z_1 * z_3 + z_2 * z_3 + z_3 ^ 2 + z_1 * z_4 + z_2 * z_4 + z_3 * z_4 + z_4 ^ 2)


        b = Volume / 10 * (x_1 ^ 2 + x_1 * x_2 + x_2 ^ 2 + x_1 * x_3 + x_2 * x_3 + x_3 ^ 2 + x_1 * x_4 + x_2 * x_4 + x_3 * x_4 + x_4 ^ 2 + _
        z_1 ^ 2 + z_1 * z_2 + z_2 ^ 2 + z_1 * z_3 + z_2 * z_3 + z_3 ^ 2 + z_1 * z_4 + z_2 * z_4 + z_3 * z_4 + z_4 ^ 2)

        c = Volume / 10 * (x_1 ^ 2 + x_1 * x_2 + x_2 ^ 2 + x_1 * x_3 + x_2 * x_3 + x_3 ^ 2 + x_1 * x_4 + x_2 * x_4 + x_3 * x_4 + x_4 ^ 2 + _
        y_1 ^ 2 + y_1 * y_2 + y_2 ^ 2 + y_1 * y_3 + y_2 * y_3 + y_3 ^ 2 + y_1 * y_4 + y_2 * y_4 + y_3 * y_4 + y_4 ^ 2)



        a_ = Volume / 20 * (2 * y_1 * z_1 + y_2 * z_1 + y_3 * z_1 + y_4 * z_1 + y_1 * z_2 + 2 * y_2 * z_2 + y_3 * z_2 + _
        y_4 * z_2 + y_1 * z_3 + y_2 * z_3 + 2 * y_3 * z_3 + y_4 * z_3 + y_1 * z_4 + y_2 * z_4 + y_3 * z_4 + 2 * y_4 * z_4)

        b_ = Volume / 20 * (2 * x_1 * z_1 + x_2 * z_1 + x_3 * z_1 + x_4 * z_1 + x_1 * z_2 + 2 * x_2 * z_2 + x_3 * z_2 + _
        x_4 * z_2 + x_1 * z_3 + x_2 * z_3 + 2 * x_3 * z_3 + x_4 * z_3 + x_1 * z_4 + x_2 * z_4 + x_3 * z_4 + 2 * x_4 * z_4)

        c_ = Volume / 20 * (2 * x_1 * y_1 + x_2 * y_1 + x_3 * y_1 + x_4 * y_1 + x_1 * y_2 + 2 * x_2 * y_2 + x_3 * y_2 + _
        x_4 * y_2 + x_1 * y_3 + x_2 * y_3 + 2 * x_3 * y_3 + x_4 * y_3 + x_1 * y_4 + x_2 * y_4 + x_3 * y_4 + 2 * x_4 * y_4)


        ' Translation to inertia at the mass center

        'a = a - Volume * (Centroid.Coord(1) * Centroid.Coord(1) + Centroid.Coord(2) * Centroid.Coord(2))

        'b = b - Volume * (Centroid.Coord(0) * Centroid.Coord(0) + Centroid.Coord(2) * Centroid.Coord(2))

        'c = c - Volume * (Centroid.Coord(0) * Centroid.Coord(0) + Centroid.Coord(1) * Centroid.Coord(1))



        'a_ = a_ - Volume * (Centroid.Coord(1) * Centroid.Coord(2))

        'b_ = b_ - Volume * (Centroid.Coord(0) * Centroid.Coord(1))

        'c_ = c_ - Volume * (Centroid.Coord(0) * Centroid.Coord(2))




        ReDim Inertia_tensor(2, 2)


        Inertia_tensor(0, 0) = a

        Inertia_tensor(0, 1) = -b_

        Inertia_tensor(0, 2) = -c_

        Inertia_tensor(1, 0) = -b_

        Inertia_tensor(1, 1) = b

        Inertia_tensor(1, 2) = -a_

        Inertia_tensor(2, 0) = -c_

        Inertia_tensor(2, 1) = -a_

        Inertia_tensor(2, 2) = c


    End Sub


End Class

Public Class Test_Tetrahedron

    Inherits Linear_algebra

    Public Sub New()

        Dim Volume As Double

        Dim Centroid As Vector = New Vector

        Dim Inertia(,) As Double


        Dim Vertex_1(5) As Vector

        Vertex_1(0) = New Vector

        Vertex_1(1) = New Vector

        Vertex_1(2) = New Vector

        Vertex_1(3) = New Vector

        Vertex_1(4) = New Vector

        Vertex_1(5) = New Vector


        Vertex_1(0).Coord(0) = 0

        Vertex_1(0).Coord(1) = 0

        Vertex_1(0).Coord(2) = 0


        Vertex_1(1).Coord(0) = 1

        Vertex_1(1).Coord(1) = 0

        Vertex_1(1).Coord(2) = 0


        Vertex_1(2).Coord(0) = 0

        Vertex_1(2).Coord(1) = 1

        Vertex_1(2).Coord(2) = 0


        Vertex_1(3).Coord(0) = 0

        Vertex_1(3).Coord(1) = 0

        Vertex_1(3).Coord(2) = 1


        Vertex_1(4).Coord(0) = 1

        Vertex_1(4).Coord(1) = 0

        Vertex_1(4).Coord(2) = 1


        Vertex_1(5).Coord(0) = 0

        Vertex_1(5).Coord(1) = 1

        Vertex_1(5).Coord(2) = 1

        Dim MIP As New MIP

        Inertia = MIP.Inertia_6_node(Vertex_1, Volume, Centroid)


        Dim Vertex_2(7) As Vector


        Vertex_2(0) = New Vector

        Vertex_2(1) = New Vector

        Vertex_2(2) = New Vector

        Vertex_2(3) = New Vector

        Vertex_2(4) = New Vector

        Vertex_2(5) = New Vector

        Vertex_2(6) = New Vector

        Vertex_2(7) = New Vector



        Vertex_2(0).Coord(0) = 0

        Vertex_2(0).Coord(1) = 0

        Vertex_2(0).Coord(2) = 0


        Vertex_2(1).Coord(0) = 1

        Vertex_2(1).Coord(1) = 0

        Vertex_2(1).Coord(2) = 0


        Vertex_2(2).Coord(0) = 1

        Vertex_2(2).Coord(1) = 1

        Vertex_2(2).Coord(2) = 0


        Vertex_2(3).Coord(0) = 0

        Vertex_2(3).Coord(1) = 1

        Vertex_2(3).Coord(2) = 0



        Vertex_2(4).Coord(0) = 0

        Vertex_2(4).Coord(1) = 0

        Vertex_2(4).Coord(2) = 1


        Vertex_2(5).Coord(0) = 1

        Vertex_2(5).Coord(1) = 0

        Vertex_2(5).Coord(2) = 1


        Vertex_2(6).Coord(0) = 1

        Vertex_2(6).Coord(1) = 1

        Vertex_2(6).Coord(2) = 1


        Vertex_2(7).Coord(0) = 0

        Vertex_2(7).Coord(1) = 1

        Vertex_2(7).Coord(2) = 1

        Inertia = MIP.Inertia_8_node(Vertex_2, Volume, Centroid)


        Dim Vertex_3(1) As Vector

        Dim Area As Double = 1


        Vertex_3(0) = New Vector

        Vertex_3(1) = New Vector


        Vertex_3(0).Coord(0) = 0

        Vertex_3(0).Coord(1) = 0

        Vertex_3(0).Coord(2) = 0


        Vertex_3(1).Coord(0) = 1

        Vertex_3(1).Coord(1) = 1

        Vertex_3(1).Coord(2) = 1


        Inertia = MIP.Inertia_2_node(Vertex_3, Area, Volume, Centroid)


    End Sub
End Class