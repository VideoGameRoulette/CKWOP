Public Class Canopy
    Inherits StandardCabinet

    Public Overridable Function Gable(h, d, a)
        Dim Q As Double = 2 * a
        Dim X As Double = (h * 10) - 1
        Dim Y As Double = d * 10
        Dim Z As Double = 16

        Dim Material As String = BoxMaterial(0)
        Dim Edge(2) As String
        Dim EdgeCode As String = ""

        If (X > Y) Then
            Edge(0) = "E2S"
            Edge(1) = "E1L"
        Else
            Edge(0) = "E2L"
            Edge(1) = "E1S"
        End If

        Select Case Form1.speciesBox.Text.ToUpper
            Case "PVC"
                EdgeCode = PMaterial(2)
            Case Else
                EdgeCode = VMaterial(3)
        End Select

        Form1.DataGridView1.Rows.Add(Q, "Gable", Y, X, Z, Material, Edge(0), EdgeCode, Edge(1), EdgeCode)
        Gable = ""
    End Function

    Public Overridable Function Top(w, d, a)
        Dim X As Double = ((w * 10) - (16 * 2)) - 1
        Dim Y As Double = d * 10
        Dim Z As Double = 16
        Dim Q As Double = 1 * a
        Dim Material As String = BoxMaterial(0)
        Dim Edge As String = ""
        Dim EdgeCode As String = ""

        If (X > Y) Then
            Edge = "E1L"
        Else
            Edge = "E1S"
        End If

        Select Case Form1.speciesBox.Text.ToUpper
            Case "PVC"
                EdgeCode = PMaterial(2)
            Case Else
                EdgeCode = VMaterial(3)
        End Select

        Form1.DataGridView1.Rows.Add(Q, "Top", Y, X, Z, Material, Edge, EdgeCode)
        Top = ""
    End Function

    Public Overridable Function Front(w, h, a)
        Dim Q As Double = 1 * a
        Dim X As Double
        Dim Y As Double
        Dim Z As Double = 19
        Dim Material As String = ""
        Dim Edge As String = ""
        Dim EdgeCode As String = ""

        Select Case Form1.speciesBox.Text
            Case "MDF"
                X = h * 10
                Y = w * 10
                Material = VMaterial(0)
            Case Else
                X = (h * 10) - 1
                Y = (w * 10) - 1
                Material = VMaterial(0)
                Edge = "E4S"
                EdgeCode = VMaterial(3)
        End Select

        Form1.DataGridView1.Rows.Add(Q, "Front", Y, X, Z, Material, Edge, EdgeCode)
        Front = ""
    End Function

    Public Overridable Function FanShelf(w, d2, a)
        Dim Q As Double = 1 * a
        Dim X As Double = w * 10
        Dim Y As Double = d2 * 10
        Dim Z As Double = 19
        Dim Material As String = "Ply"
        Form1.DataGridView1.Rows.Add(Q, "Fan Sh.", Y, X, Z, Material)
        FanShelf = ""
    End Function

    Public Overridable Function SmallFront(w, h2, a)
        Dim Q As Double = 1 * a
        Dim X As Double = w * 10
        Dim Y As Double = h2 * 10
        Dim Z As Double = 19
        Dim Material As String = ""

        Select Case Form1.speciesBox.Text.ToUpper
            Case "PVC"
                Material = PMaterial(0)
            Case Else
                Material = VMaterial(0)
        End Select

        Form1.DataGridView1.Rows.Add(Q, "SmallFront", Y, X, Z, Material)
        SmallFront = ""
    End Function

    Public Overridable Function SmallTop(w, d1, d2, a)
        Dim Q As Double = 1 * a
        Dim X As Double = w * 10
        Dim Y As Double = (d2 - d1) * 10
        Dim Z As Double = 19
        Dim Material As String = ""

        Select Case Form1.speciesBox.Text.ToUpper
            Case "PVC"
                Material = PMaterial(0)
            Case Else
                Material = VMaterial(0)
        End Select

        Form1.DataGridView1.Rows.Add(Q, "SmallTop", Y, X, Z, Material)
        SmallTop = ""
    End Function

    Public Overridable Function InsertPanel(w, h, h2, a)
        Dim Q As Integer
        Dim Q2 As Integer
        Dim X As Double
        Dim X2 As Double
        Dim Y As Double
        Dim Y2 As Double
        Dim Z As Double = 6
        Dim Material As String = ""

        Select Case Form1.CabCode.Text.ToUpper
            Case "RHAW-S"
                X = (h * 10) - (h2 * 10) - 128
                Y = (w * 10) - 120
                Q = 1 * a
            Case "RHAW-D"
                X = (h * 10) - (h2 * 10) - 128
                Y = ((w * 10) - 180) / 2
                Q = 2 * a
            Case "RHAW-T"
                X = (h * 10) - (h2 * 10) - 128
                Y = ((w * 10) - 240) / 2
                Q = 1 * a
                X2 = (h * 10) - (h2 * 10) - 128
                Y2 = (((w * 10) - 240) / 2) / 2
                Q2 = 2 * a
        End Select

        Select Case Form1.speciesBox.Text.ToUpper
            Case "PVC"
                Material = PMaterial(0)
            Case Else
                Material = VMaterial(0)
        End Select

        Form1.DataGridView1.Rows.Add(Q, "IPanel", Y, X, Z, Material)
        If (Form1.CabCode.Text.ToUpper = "RHAW-T") Then
            Form1.DataGridView1.Rows.Add(Q2, "IPanel2", Y2, X2, Z, Material)
        End If
        InsertPanel = ""
    End Function

    Public Overridable Function Back(w, h, d, a)
        Dim X As Double
        If (d < 40) Then
            X = (h * 10) - 130
        Else
            X = (h * 10) - 16 + 5
        End If
        Dim Y As Double = (w * 10) - 23
        Dim Z As Double = 16
        Dim Q As Double = 1 * a
        Dim Material As String = BoxMaterial(0)
        Dim Edge As String = ""
        Dim EdgeCode As String = ""

        If (X > Y) Then
            Edge = "E1L"
        Else
            Edge = "E1S"
        End If

        Select Case Form1.speciesBox.Text.ToUpper
            Case "PVC"
                EdgeCode = PMaterial(2)
            Case Else
                EdgeCode = VMaterial(3)
        End Select

        Form1.DataGridView1.Rows.Add(Q, "Back", Y, X, Z, Material, Edge, EdgeCode)
        Back = ""
    End Function

End Class

Public Class CanopyHood
    Inherits Canopy

    Public Overrides Function Back(w, h, d, a)
        Dim Q As Double = 1 * a
        Dim X As Double

        If (d < 40) Then
            X = (h * 10) - Form1.overrideFSL.Text - 19 - 16 + 5
        Else
            X = (h * 10) - 11
        End If

        Dim Y As Double = (w * 10) - 23
        Dim Z As Double = 16
        Dim Material As String = BoxMaterial(0)
        Dim Edge As String
        Dim EdgeCode As String

        If (X > Y) Then
            Edge = "E1L"
        Else
            Edge = "E1S"
        End If

        Select Case Form1.speciesBox.Text.ToUpper
            Case "PVC"
                EdgeCode = PMaterial(2)
            Case Else
                EdgeCode = VMaterial(3)
        End Select

        Form1.DataGridView1.Rows.Add(Q, "Back", Y, X, Z, Material, Edge, EdgeCode)
        Back = ""
    End Function

End Class