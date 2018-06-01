Public Class Upper
    Inherits StandardCabinet

    Public Overridable Function Gable(h, d, a)
        Dim Q As Integer = 2 * a
        Dim GableX As Double = (h * 10) - 1
        Dim GableY As Double = (d * 10)
        Dim GableZ As Double = 16
        Dim Material As String = BoxMaterial(0)
        Dim Edge(2) As String
        Dim EdgeCode(2) As String

        If (GableX > GableY) Then
            Edge(0) = "E2S"
            Edge(1) = "E1L"
        Else
            Edge(0) = "E2L"
            Edge(1) = "E1S"
        End If

        EdgeCode(0) = BoxMaterial(1)
        EdgeCode(1) = PMaterial(2)

        Form1.DataGridView1.Rows.Add(Q, "Gable", GableY, GableX, GableZ, Material, Edge(0), EdgeCode(0), Edge(1), EdgeCode(1))
        Form1.DataGridView1.Rows.Add("")
        Return True
    End Function

    Public Overridable Function Top(w, d, a)
        Dim Q As Integer = 1 * a
        Dim X As Double = (w * 10) - 32 - 1
        Dim Y As Double = (d * 10)
        Dim Z As Double = 16
        Dim Material As String = BoxMaterial(0)
        Dim Edge As String = "" = ""
        Dim EdgeCode As String = BoxMaterial(1)

        If (X > Y) Then
            Edge = "E1L"
        Else
            Edge = "E1S"
        End If

        Form1.DataGridView1.Rows.Add(Q, "Top", Y, X, Z, Material, Edge, EdgeCode)
        Return True
    End Function

    Public Overridable Function TopHinge(w, h, d, a)
        Dim Q As Integer = 1 * a
        Dim X As Double = (w * 10) - 32 - 1
        Dim Y As Double = (d * 10)
        Dim Z As Double = 16
        Dim Material As String = BoxMaterial(0)
        Dim Edge As String
        Dim EdgeCode As String = BoxMaterial(1)

        If (X > Y) Then
            Edge = "E1L"
        Else
            Edge = "E1S"
        End If

        Form1.DataGridView1.Rows.Add(Q, "Top Hinge", Y, X, Z, Material, Edge, EdgeCode)
        Return True
    End Function

    Public Overridable Function Bottom(w, d, a)
        Dim Q As Integer = 1 * a
        Dim BottomX As Double = (w * 10) - 32 - 1
        Dim BottomY As Double = (d * 10)
        Dim BottomZ As Double = 16
        Dim Material As String = BoxMaterial(0)
        Dim Edge As String = "" = ""
        Dim EdgeCode As String = PMaterial(2)

        If (BottomX > BottomY) Then
            Edge = "E1L"
        Else
            Edge = "E1S"
        End If

        Form1.DataGridView1.Rows.Add(Q, "Bottom", BottomY, BottomX, BottomZ, Material, Edge, EdgeCode)
        Return True
    End Function

    Public Overridable Function TopBtm(w, d, a)
        Dim Q As Integer = 2 * a
        Dim X As Double = (w * 10) - (16 * 2) - 1
        Dim Y As Double = (d * 10)
        Dim Z As Double = 16
        Dim Material As String = BoxMaterial(0)
        Dim Edge As String = "" = ""
        Dim EdgeCode As String = ""

        If (X > Y) Then
            Edge = "E1L"
        Else
            Edge = "E1S"
        End If

        If (Form1.groupBox.Text = "Group 1") Then
            EdgeCode = BoxMaterial(1)
        ElseIf (Form1.groupBox.Text = "Group 2") Then
            EdgeCode = PMaterial(2)
        End If

        Form1.DataGridView1.Rows.Add(Q, "TopBtm", Y, X, Z, Material, Edge, EdgeCode)
        Return True
    End Function

    Public Overridable Function AdjSh(w, h, d, a)
        Dim Q As Integer
        Dim X As Double = (w * 10) - (16 * 2) - 3
        Dim Y As Double
        Dim Z As Double = 16
        Dim Material As String = BoxMaterial(0)
        Dim Edge As String = "" = ""
        Dim EdgeCode As String = BoxMaterial(1)

        Dim Size = d
        Select Case d
            Case 0 To 48
                Y = (d * 10) - 30
            Case 48 To 61
                Y = 450
            Case Else
                Y = (d * 10) - 30
        End Select

        If (X > Y) Then
            Edge = "E1L"
        Else
            Edge = "E1S"
        End If

        Dim GableHeight As Double = (h * 10) - 1
        Select Case GableHeight
            Case 0 To 497
                Return False
            Case 498 To 698
                Q = 1 * a
                Form1.DataGridView1.Rows.Add(Q, "Adj Sh.", Y, X, Z, Material, Edge, EdgeCode)
            Case 699 To 897
                Q = 2 * a
                Form1.DataGridView1.Rows.Add(Q, "Adj Sh.", Y, X, Z, Material, Edge, EdgeCode)
            Case 898 To 1439
                Q = 3 * a
                Form1.DataGridView1.Rows.Add(Q, "Adj Sh.", Y, X, Z, Material, Edge, EdgeCode)
            Case Else
                Q = 4 * a
                Form1.DataGridView1.Rows.Add(Q, "Adj Sh.", Y, X, Z, Material, Edge, EdgeCode)
        End Select

        Return True
    End Function

    Public Overridable Function Back(w, h, a)
        Dim Q As Integer = 1 * a
        Dim X As Double = (h * 10) - 23
        Dim Y As Double = (w * 10) - 23
        Dim Z As Double = 3
        Dim Material As String = BoxMaterial(0)

        Form1.DataGridView1.Rows.Add(Q, "Back", Y, X, Z, Material)
        Return True
    End Function

End Class

Public Class UpperCornerDiagonal
    Inherits Upper

    Public Overrides Function Gable(h, d, a)
        Dim Q As Integer = 2 * a
        Dim GableX As Double = (h * 10) - 1
        Dim GableY As Double = (d * 10)
        Dim GableZ As Double = 19
        Dim Material As String = BoxMaterial(0)
        Dim Edge(2) As String
        Dim EdgeCode(2) As String

        If (GableX > GableY) Then
            Edge(0) = "E2S"
            Edge(1) = "E1L"
        Else
            Edge(0) = "E2L"
            Edge(1) = "E1S"
        End If

        EdgeCode(0) = BoxMaterial(1)
        EdgeCode(1) = PMaterial(2)

        Form1.DataGridView1.Rows.Add(Q, "Gable", GableY, GableX, GableZ, Material, Edge(0), EdgeCode(0), Edge(1), EdgeCode(1))
        Form1.DataGridView1.Rows.Add("")
        Gable = ""
    End Function

    Public Overrides Function TopBtm(w, d, a)
        Dim Q As Integer = 2 * a
        Dim TopBtmX As Double = ((w * 10) - 19) + 1
        Dim TopBtmY As Double = ((w * 10) - 19) + 1
        Dim TopBtmZ As Double = 16
        Dim Material As String = BoxMaterial(0)
        Dim Edge As String = "E1M"
        Dim EdgeCode As String = ""

        If (Form1.groupBox.Text = "Group 1") Then
            EdgeCode = BoxMaterial(1)
        ElseIf (Form1.groupBox.Text = "Group 2") Then
            EdgeCode = PMaterial(2)
        End If

        Form1.DataGridView1.Rows.Add(Q, "TopBtm", TopBtmY, TopBtmX, TopBtmZ, Material, Edge, EdgeCode)
        TopBtm = ""
    End Function

    Public Overrides Function AdjSh(w, h, d, a)
        Dim Q As Integer
        Dim X As Double = (((w * 10) - 19) - 22) - 4
        Dim Y As Double = (((w * 10) - 19) - 22) - 4
        Dim Z As Double = 16
        Dim Material As String = BoxMaterial(0)
        Dim Edge As String = "E1M"
        Dim EdgeCode As String = BoxMaterial(1)

        Dim GableHeight As Double = (h * 10) - 1
        Select Case GableHeight
            Case 0 To 497
                Return False
            Case 498 To 698
                Q = 1 * a
                Form1.DataGridView1.Rows.Add(Q, "Adj Sh.", Y, X, Z, Material, Edge, EdgeCode)
            Case 699 To 897
                Q = 2 * a
                Form1.DataGridView1.Rows.Add(Q, "Adj Sh.", Y, X, Z, Material, Edge, EdgeCode)
            Case 898 To 1439
                Q = 3 * a
                Form1.DataGridView1.Rows.Add(Q, "Adj Sh.", Y, X, Z, Material, Edge, EdgeCode)
            Case Else
                Q = 4 * a
                Form1.DataGridView1.Rows.Add(Q, "Adj Sh.", Y, X, Z, Material, Edge, EdgeCode)
        End Select

        Return True
    End Function

    Public Overridable Function BackStrap(h, a)
        Dim Q As Integer = 1 * a
        Dim X As Double = (h * 10) - 32
        Dim Y As Double = 96
        Dim Z As Double = 16
        Dim Material As String = BoxMaterial(0)
        Form1.DataGridView1.Rows.Add(Q, "BackStrap", Y, X, Z, Material)
        Return True
    End Function

    Public Overrides Function Back(w, h, a)
        Dim Q As Integer = 2 * a
        Dim X As Double = ((h * 10) - (16 * 2)) + 10 - 1
        Dim Y As Double = (w * 10) - 80
        Dim Z As Double = 16
        Dim Material As String = BoxMaterial(0)
        Form1.DataGridView1.Rows.Add(Q, "Back", Y, X, Z, Material)
        Return True
    End Function

End Class

Public Class UpperCornerDiagonalMI
    Inherits UpperMI

    Public Overrides Function Gable(h, d, a)
        Dim Q As Integer = 2 * a
        Dim X As Double = (h * 10) - 1
        Dim Y As Double = (d * 10)
        Dim Z As Double = 19
        Dim Material As String
        Dim Edge(2) As String
        Dim EdgeCode As String

        If (Form1.speciesBox.Text = "PVC") Then
            Material = PMaterial(0)
            EdgeCode = PMaterial(2)
        Else
            Material = VMaterial(0)
            EdgeCode = VMaterial(3)
        End If

        If (X > Y) Then
            Edge(0) = "E2S"
            Edge(1) = "E1L"
        Else
            Edge(0) = "E2L"
            Edge(1) = "E1S"
        End If

        Form1.DataGridView1.Rows.Add(Q, "Gable", Y, X, Z, Material, Edge(0), EdgeCode, Edge(1), EdgeCode)
        Form1.DataGridView1.Rows.Add("")
        Return True
    End Function

    Public Overrides Function TopBtm(w, d, a)
        Dim Q As Integer = 2 * a
        Dim X As Double = ((w * 10) - 19) + 1
        Dim Y As Double = ((w * 10) - 19) + 1
        Dim Z As Double = 16
        Dim Material As String
        Dim Edge As String = "E1M"
        Dim EdgeCode As String

        If (Form1.speciesBox.Text = "PVC") Then
            Material = PMaterial(0)
            EdgeCode = PMaterial(2)
        ElseIf (Form1.speciesBox.Text = "MDF") Then
            Material = VMaterial(1)
            EdgeCode = VMaterial(3)
        Else
            Material = VMaterial(0)
            EdgeCode = VMaterial(3)
        End If

        Form1.DataGridView1.Rows.Add(Q, "TopBtm", Y, X, Z, Material, Edge, EdgeCode)
        Return True
    End Function

    Public Overrides Function AdjSh(w, h, d, a)
        Dim Q As Integer
        Dim X As Double = (((w * 10) - 19) - 35) - 4
        Dim Y As Double = (((w * 10) - 19) - 35) - 4
        Dim Z As Double = 16
        Dim Material As String
        Dim Edge As String = "E1M"
        Dim EdgeCode As String

        If (Form1.speciesBox.Text = "PVC") Then
            Material = PMaterial(0)
            EdgeCode = PMaterial(2)
        Else
            Material = VMaterial(0)
            EdgeCode = VMaterial(3)
        End If

        Dim GableHeight As Double = (h * 10) - 1
        Select Case GableHeight
            Case 0 To 497
                Return False
            Case 498 To 698
                Q = 1 * a
                Form1.DataGridView1.Rows.Add(Q, "Adj Sh.", Y, X, Z, Material, Edge, EdgeCode)
            Case 699 To 897
                Q = 2 * a
                Form1.DataGridView1.Rows.Add(Q, "Adj Sh.", Y, X, Z, Material, Edge, EdgeCode)
            Case 898 To 1439
                Q = 3 * a
                Form1.DataGridView1.Rows.Add(Q, "Adj Sh.", Y, X, Z, Material, Edge, EdgeCode)
            Case Else
                Q = 4 * a
                Form1.DataGridView1.Rows.Add(Q, "Adj Sh.", Y, X, Z, Material, Edge, EdgeCode)
        End Select

        Return True
    End Function

    Public Overridable Function SmallBack(w, h, a)
        Dim Q As Integer = 1 * a
        Dim X As Double = (((h * 10) - (16 * 2)) + 10) - 1
        Dim Y As Double = (((w * 10) - 19) - 35) + 10
        Dim Z As Double = 16
        Dim Material As String

        If (Form1.speciesBox.Text = "PVC") Then
            Material = PMaterial(0)
        Else
            Material = VMaterial(0)
        End If

        Form1.DataGridView1.Rows.Add(Q, "S-Back", Y, X, Z, Material)
        Return True
    End Function

    Public Overridable Function LargeBack(w, h, a)
        Dim Q As Integer = 1 * a
        Dim X As Double = (((h * 10) - (16 * 2)) + 10) - 1
        Dim Y As Double = ((((w * 10) - 19) - 35) + 16) + 5
        Dim Z As Double = 16
        Dim Material As String

        If (Form1.speciesBox.Text = "PVC") Then
            Material = PMaterial(0)
        Else
            Material = VMaterial(0)
        End If

        Form1.DataGridView1.Rows.Add(Q, "L-Back", Y, X, Z, Material)
        Return True
    End Function

End Class

Public Class UpperMI
    Inherits Upper

    Public Overrides Function Gable(h, d, a)
        Dim Q As Integer = 2 * a
        Dim X As Double = (h * 10) - 1
        Dim Y As Double = (d * 10)
        Dim Z As Double = 16
        Dim Material As String
        Dim Edge(2) As String
        Dim EdgeCode As String
        If (Form1.speciesBox.Text = "PVC") Then
            Material = PMaterial(0)
            EdgeCode = PMaterial(2)
        Else
            Material = VMaterial(0)
            EdgeCode = VMaterial(3)
        End If
        If (X > Y) Then
            Edge(0) = "E2S"
            Edge(1) = "E1L"
        Else
            Edge(0) = "E2L"
            Edge(1) = "E1S"
        End If
        Form1.DataGridView1.Rows.Add(Q, "Gable", Y, X, Z, Material, Edge(0), EdgeCode, Edge(1), EdgeCode)
        Form1.DataGridView1.Rows.Add("")
        Return True
    End Function

    Public Overrides Function TopBtm(w, d, a)
        Dim Q As Integer = 2 * a
        Dim X As Double = (w * 10) - 32
        Dim Y As Double = (d * 10)
        Dim Z As Double = 16
        Dim Material As String
        Dim Edge As String = ""
        Dim EdgeCode As String
        If (Form1.speciesBox.Text = "PVC") Then
            Material = PMaterial(0)
            EdgeCode = PMaterial(2)
        ElseIf (Form1.speciesBox.Text = "MDF") Then
            Material = VMaterial(1)
            EdgeCode = VMaterial(3)
        Else
            Material = VMaterial(0)
            EdgeCode = VMaterial(3)
        End If
        If (X > Y) Then
            Edge = "E1L"
        Else
            Edge = "E1S"
        End If
        Form1.DataGridView1.Rows.Add(Q, "TopBtm", Y, X, Z, Material, Edge, EdgeCode)
        Return True
    End Function

    Public Overrides Function AdjSh(w, h, d, a)
        Dim Q As Integer
        Dim X As Double = (w * 10) - 32 - 3
        Dim Y As Double
        Select Case d
            Case 0 To 48
                Y = (d * 10) - 30
            Case 48 To 61
                Y = 450
            Case Else
                Y = (d * 10) - 30
        End Select
        Dim Z As Double = 16
        Dim Material As String
        Dim Edge As String = ""
        Dim EdgeCode As String
        If (Form1.speciesBox.Text = "PVC") Then
            Material = PMaterial(0)
            EdgeCode = PMaterial(2)
        Else
            Material = VMaterial(0)
            EdgeCode = VMaterial(3)
        End If
        If (X > Y) Then
            Edge = "E1L"
        Else
            Edge = "E1S"
        End If

        Dim GableHeight As Double = (h * 10) - 1
        Select Case GableHeight
            Case 0 To 497
                Return False
            Case 498 To 698
                Q = 1 * a
                Form1.DataGridView1.Rows.Add(Q, "Adj Sh.", Y, X, Z, Material, Edge, EdgeCode)
            Case 699 To 897
                Q = 2 * a
                Form1.DataGridView1.Rows.Add(Q, "Adj Sh.", Y, X, Z, Material, Edge, EdgeCode)
            Case 898 To 1439
                Q = 3 * a
                Form1.DataGridView1.Rows.Add(Q, "Adj Sh.", Y, X, Z, Material, Edge, EdgeCode)
            Case Else
                Q = 4 * a
                Form1.DataGridView1.Rows.Add(Q, "Adj Sh.", Y, X, Z, Material, Edge, EdgeCode)
        End Select

        Return True
    End Function

    Public Overrides Function Back(w, h, a)
        Dim BackX As Double = (h * 10) - 23
        Dim BackY As Double = (w * 10) - 23
        Dim BackZ As Double = 16
        Dim Material As String = BoxMaterial(0)
        If (Form1.speciesBox.Text = "PVC") Then
            Material = PMaterial(0)
        Else
            Material = VMaterial(0)
        End If
        Form1.DataGridView1.Rows.Add(1, "Back", BackY, BackX, BackZ, Material)
        Return True
    End Function

End Class

Public Class UpperOpen
    Inherits UpperMI

End Class

Public Class OpenAppliance
    Inherits UpperOpen

    Public Overrides Function Header(w, h, d, a)
        Dim Code As String = w & "-" & h & "-" & d
        Form1.DataGridView1.Rows.Add("UA", "Qty", a, "", "", "", "")
        Form1.DataGridView1.Rows.Add(Code, "", "", "", "", "", "")
        Form1.DataGridView1.Rows.Add("Qty", "Part", "H", "W", "D", "Mat", "Edge", "Code", "Edge", "Code", "Notes")
        Header = ""
    End Function

    Public Overrides Function Gable(h, d, a)
        Dim Q As Integer = 2 * a
        Dim X As Double = (h * 10) - 1
        Dim Y As Double = (d * 10)
        Dim Z As Double = 19
        Dim Material As String
        Dim Edge(2) As String
        Dim EdgeCode As String
        If (Form1.speciesBox.Text = "PVC") Then
            Material = PMaterial(0)
            EdgeCode = PMaterial(2)
        Else
            Material = VMaterial(0)
            EdgeCode = VMaterial(3)
        End If
        If (X > Y) Then
            Edge(0) = "E2S"
            Edge(1) = "E1L"
        Else
            Edge(0) = "E2L"
            Edge(1) = "E1S"
        End If
        Form1.DataGridView1.Rows.Add(Q, "Gable", Y, X, Z, Material, Edge(0), EdgeCode, Edge(1), EdgeCode)
        Form1.DataGridView1.Rows.Add("")
        Return True
    End Function

    Public Overrides Function TopBtm(w, d, a)
        Dim Q As Integer = 2 * a
        Dim X As Double = (w * 10) - 38
        Dim Y As Double = (d * 10)
        Dim Z As Double = 16
        Dim Material As String
        Dim Edge As String = ""
        Dim EdgeCode As String
        If (Form1.speciesBox.Text = "PVC") Then
            Material = PMaterial(0)
            EdgeCode = PMaterial(2)
        ElseIf (Form1.speciesBox.Text = "MDF") Then
            Material = VMaterial(1)
            EdgeCode = VMaterial(3)
        Else
            Material = VMaterial(0)
            EdgeCode = VMaterial(3)
        End If
        If (X > Y) Then
            Edge = "E1L"
        Else
            Edge = "E1S"
        End If
        Form1.DataGridView1.Rows.Add(Q, "TopBtm", Y, X, Z, Material, Edge, EdgeCode)
        Return True
    End Function

End Class

Public Class UpperMicrowave
    Inherits UpperMI

    Public Overridable Function MicrowaveShelf(w, a)
        Dim Q As Integer = 1 * a
        Dim X As Double = (w * 10) - (16 * 2) - 3
        Dim Y As Double = 460
        Dim Z As Double = 16
        Dim Material As String
        Dim Edge(2) As String
        Dim EdgeCode As String

        If (Form1.speciesBox.Text = "PVC") Then
            Material = PMaterial(0)
            EdgeCode = PMaterial(2)
        Else
            Material = VMaterial(0)
            EdgeCode = VMaterial(3)
        End If

        If (X > Y) Then
            Edge(0) = "E2S"
            Edge(1) = "E1L"
        Else
            Edge(0) = "E2L"
            Edge(1) = "E1S"
        End If

        Form1.DataGridView1.Rows.Add(Q, "Micro Sh.", Y, X, Z, Material, Edge(0), EdgeCode, Edge(1), EdgeCode)

        Return True
    End Function

End Class

Public Class UpperTrayDivider
    Inherits Upper

    Public Overridable Function Divider(h, d, a)
        Dim Q As Integer = 1 * a
        Dim X As Double = (h * 10) - (16 * 2) - 1
        Dim Y As Double = (d * 10) - 22
        Dim Z As Double = 16
        Dim Material As String = BoxMaterial(0)
        Dim Edge As String
        Dim EdgeCode As String = BoxMaterial(1)

        If (X > Y) Then
            Edge = "E1L"
        Else
            Edge = "E1S"
        End If

        Form1.DataGridView1.Rows.Add(Q, "Divider", Y, X, Z, Material, Edge, EdgeCode)
        Return True
    End Function

End Class