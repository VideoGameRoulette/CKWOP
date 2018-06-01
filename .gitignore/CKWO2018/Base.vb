Public Class Base
    Inherits StandardCabinet

    Public Overrides Function Header(w, h, d, a)
        Dim Code As String = w & "-" & h & "-" & d
        Form1.DataGridView1.Rows.Add("B", "Qty", a, "", "", "", "")
        Form1.DataGridView1.Rows.Add(Code, "", "", "", "", "", "")
        Form1.DataGridView1.Rows.Add("Qty", "Part", "H", "W", "D", "Mat", "Edge", "Code", "Edge", "Code", "Notes")
        Header = ""
    End Function

    Public Overridable Function Back(w, h, a)
        Dim Q As Integer = 1 * a
        Dim X As Double
        Dim Y As Double = (w * 10) - 23
        Dim Z As Double = 3
        Dim Material As String = BoxMaterial(0)

        If (Form1.topStrapY.Checked = True) Then
            X = (h * 10) - 120 - 26
        Else
            X = (h * 10) - 120 - 23
        End If

        Form1.DataGridView1.Rows.Add(Q, "Back", Y, X, Z, Material)
        Back = ""
    End Function

    Public Overridable Function TopBtm(w, d, a)
        Dim Q As Integer = 2 * a
        Dim X As Double = (w * 10) - 32 - 1
        Dim Y As Double = 300
        Dim Z As Double = 16
        Dim Material As String = BoxMaterial(0)
        Dim Edge As String = ""
        Dim EdgeCode As String = PMaterial(2)

        If (X > Y) Then
            Edge = "E1L"
        Else
            Edge = "E1S"
        End If

        Form1.DataGridView1.Rows.Add(Q, "TopBtm", Y, X, Z, Material, Edge, EdgeCode)
        TopBtm = ""
    End Function

    Public Overridable Function TopStrap(w, d, a)
        Dim Q As Integer
        Dim X As Double = (w * 10) - 32 - 1
        Dim Y As Double
        Dim Z As Double = 16
        Dim Material As String = BoxMaterial(0)
        Dim Edge As String = ""
        Dim EdgeCode As String = BoxMaterial(1)

        If (Form1.topStrapY.Checked = True) Then
            Q = 2 * a
            Y = 115
        Else
            Q = 1 * a
            Y = 300
        End If

        If (X > Y) Then
            Edge = "E1L"
        Else
            Edge = "E1S"
        End If

        Form1.DataGridView1.Rows.Add(Q, "Top Strap", Y, X, Z, Material, Edge, EdgeCode)
        TopStrap = ""
    End Function

    Public Overridable Function Top(w, d, a)
        Dim Q As Integer = 1 * a
        Dim X As Double = (w * 10) - 32 - 1
        Dim Y As Double = (d * 10)
        Dim Z As Double = 16
        Dim Material As String = BoxMaterial(0)
        Dim Edge As String
        Dim EdgeCode As String

        If (X > Y) Then
            Edge = "E1L"
        Else
            Edge = "E1S"
        End If

        If (Form1.groupBox.Text = "Group 1") Then
            EdgeCode = BoxMaterial(1)
        Else
            EdgeCode = PMaterial(2)
        End If

        Form1.DataGridView1.Rows.Add(Q, "Top", Y, X, Z, Material, Edge, EdgeCode)
        Top = ""
    End Function

    Public Overridable Function Bottom(w, d, a)
        Dim Q As Integer = 1 * a
        Dim X As Double = (w * 10) - 32 - 1
        Dim Y As Double = (d * 10)
        Dim Z As Double = 16
        Dim Material As String = BoxMaterial(0)
        Dim Edge As String = ""
        Dim EdgeCode As String = BoxMaterial(1)

        If (X > Y) Then
            Edge = "E1L"
        Else
            Edge = "E1S"
        End If

        Form1.DataGridView1.Rows.Add(Q, "Bottom", Y, X, Z, Material, Edge, EdgeCode)
        Bottom = ""
    End Function

    Public Overridable Function Strap(w, d, a)
        Dim Q As Integer = 1 * a
        Dim X As Double = (w * 10) - 32 - 1
        Dim Y As Double = 60
        Dim Z As Double = 16
        Dim Material As String = BoxMaterial(0)
        Dim Edge As String
        Dim EdgeCode As String = PMaterial(2)

        If (X > Y) Then
            Edge = "E1L"
        Else
            Edge = "E1S"
        End If

        Form1.DataGridView1.Rows.Add(Q, "Strap", Y, X, Z, Material, Edge, EdgeCode)
        Strap = ""
    End Function

    Public Overridable Function AdjSh(w, d, a)
        Dim Q As Integer = 1 * a
        Dim X As Double = (w * 10) - 32 - 5
        Dim Y As Double
        Dim Size = d
        Select Case d
            Case 0 To 48
                Y = (d * 10) - 30
            Case 48 To 61
                Y = 450
            Case Else
                Y = (d * 10) - 30
        End Select
        Dim Z As Double = 16
        Dim Material As String = BoxMaterial(0)
        Dim Edge As String = ""
        Dim EdgeCode As String = BoxMaterial(1)

        If (X > Y) Then
            Edge = "E1L"
        Else
            Edge = "E1S"
        End If

        Form1.DataGridView1.Rows.Add(Q, "Adj Sh.", Y, X, Z, Material, Edge, EdgeCode)
        AdjSh = ""
    End Function

    Public Overridable Function Divider(d, a)
        Dim FrontH As Double = 157
        Dim Q As Integer = 1 * a
        Dim X As Double = FrontH - 18
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
        Divider = ""
    End Function

    Public Overridable Function Gable(h, d, a)
        Dim Q As Integer = 2 * a
        Dim X As Double = (h * 10)
        Dim Y As Double = (d * 10)
        Dim Z As Double = 16
        Dim Material As String = BoxMaterial(0)
        Dim Edge As String = ""
        Dim EdgeCode As String = PMaterial(2)

        If (X > Y) Then
            Edge = "E1L"
        Else
            Edge = "E1S"
        End If

        Form1.DataGridView1.Rows.Add(Q, "Gable", Y, X, Z, Material, Edge, EdgeCode)
        Form1.DataGridView1.Rows.Add("")
        Gable = ""
    End Function

    Public Overridable Function SinkStrap(w, d, a)
        Dim Q As Integer = 1 * a
        Dim X As Double = (w * 10) - 32 - 1
        Dim Y As Double = 60
        Dim Z As Double = 16
        Dim Material As String = BoxMaterial(0)

        Form1.DataGridView1.Rows.Add(Q, "Sink Strap", Y, X, Z, Material)
        SinkStrap = ""
    End Function

End Class

Public Class BaseMI
    Inherits Base

    Public Overrides Function Header(w, h, d, a)
        Dim Code As String = w & "-" & h & "-" & d
        Form1.DataGridView1.Rows.Add("BMI", "Qty", a, "", "", "", "")
        Form1.DataGridView1.Rows.Add(Code, "", "", "", "", "", "")
        Form1.DataGridView1.Rows.Add("Qty", "Part", "H", "W", "D", "Mat", "Edge", "Code", "Edge", "Code", "Notes")
        Header = ""
    End Function

    Public Overrides Function Back(w, h, a)
        Dim Q As Integer = 1 * a
        Dim X As Double
        Dim Y As Double = (w * 10) - 23
        Dim Z As Double = 16
        Dim Material As String

        If (Form1.topStrapY.Checked = True) Then
            X = (h * 10) - 120 - 26
        Else
            X = (h * 10) - 120 - 23
        End If

        If (Form1.speciesBox.Text = "PVC") Then
            Material = PMaterial(0)
        Else
            Material = VMaterial(0)
        End If

        Form1.DataGridView1.Rows.Add(Q, "Back", Y, X, Z, Material)
        Back = ""
    End Function

    Public Overrides Function TopBtm(w, d, a)
        Dim Q As Integer = 2 * a
        Dim X As Double = (w * 10) - 32 - 1
        Dim Y As Double = 300
        Dim Z As Double = 16
        Dim Material As String
        Dim Edge As String
        Dim EdgeCode As String

        If (X > Y) Then
            Edge = "E1L"
        Else
            Edge = "E1S"
        End If

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
        TopBtm = ""
    End Function

    Public Overrides Function TopStrap(w, d, a)
        Dim Q As Integer
        Dim X As Double = (w * 10) - 32 - 1
        Dim Y As Double
        Dim Z As Double = 16
        Dim Material As String
        Dim Edge As String
        Dim EdgeCode As String

        If (Form1.topStrapY.Checked = True) Then
            Q = 2 * a
            Y = 115
        Else
            Q = 1 * a
            Y = 300
        End If

        If (X > Y) Then
            Edge = "E1L"
        Else
            Edge = "E1S"
        End If

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

        Form1.DataGridView1.Rows.Add(Q, "Top Strap", Y, X, Z, Material, Edge, EdgeCode)
        TopStrap = ""
    End Function

    Public Overrides Function Bottom(w, d, a)
        Dim Q As Integer = 1 * a
        Dim X As Double = (w * 10) - 32 - 1
        Dim Y As Double = (d * 10)
        Dim Z As Double = 16
        Dim Material As String
        Dim Edge As String
        Dim EdgeCode As String

        If (X > Y) Then
            Edge = "E1L"
        Else
            Edge = "E1S"
        End If

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

        Form1.DataGridView1.Rows.Add(Q, "Bottom", Y, X, Z, Material, Edge, EdgeCode)
        Bottom = ""
    End Function

    Public Overrides Function AdjSh(w, d, a)
        Dim Q As Integer = 1 * a
        Dim X As Double = (w * 10) - 32 - 5
        Dim Y As Double
        Dim Z As Double = 16
        Dim Material As String
        Dim Edge As String
        Dim EdgeCode As String

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

        If (Form1.speciesBox.Text = "PVC") Then
            Material = PMaterial(0)
            EdgeCode = PMaterial(2)
        Else
            Material = VMaterial(0)
            EdgeCode = VMaterial(3)
        End If

        Form1.DataGridView1.Rows.Add(Q, "Adj Sh.", Y, X, Z, Material, Edge, EdgeCode)
        AdjSh = ""
    End Function

    Public Overrides Function Gable(h, d, a)
        Dim Q As Integer = 2 * a
        Dim X As Double = (h * 10)
        Dim Y As Double = (d * 10)
        Dim Z As Double = 16
        Dim Material As String
        Dim Edge As String
        Dim EdgeCode As String

        If (X > Y) Then
            Edge = "E1L"
        Else
            Edge = "E1S"
        End If

        If (Form1.speciesBox.Text = "PVC") Then
            Material = PMaterial(0)
            EdgeCode = PMaterial(2)
        Else
            Material = VMaterial(0)
            EdgeCode = VMaterial(3)
        End If

        Form1.DataGridView1.Rows.Add(Q, "Gable", Y, X, Z, Material, Edge, EdgeCode)
        Form1.DataGridView1.Rows.Add("")
        Gable = ""
    End Function

End Class

Public Class BaseDrawer
    Inherits Base

    Public Overrides Function Header(w, h, d, a)
        Dim Code As String = w & "-" & h & "-" & d
        Dim CodeExt As String = Form1.CabCode.Text & "-" & Form1.Hardware
        Form1.DataGridView1.Rows.Add(CodeExt, "Qty", a, "", "", "", "")
        Form1.DataGridView1.Rows.Add(Code, "", "", "", "", "", "")
        Form1.DataGridView1.Rows.Add("Qty", "Part", "H", "W", "D", "Mat", "Edge", "Code", "Edge", "Code", "Notes")
        Header = ""
    End Function

End Class

Public Class BaseRanged
    Inherits Base

    Public Overrides Function Header(w, h, d, a)
        Dim Code As String = w & "-" & h & "-" & d
        Form1.DataGridView1.Rows.Add(Form1.CabCode.Text, "Qty", a, "", "", "", "")
        Form1.DataGridView1.Rows.Add(Code, "", "", "", "", "", "")
        Form1.DataGridView1.Rows.Add("Qty", "Part", "H", "W", "D", "Mat", "Edge", "Code", "Edge", "Code", "Notes")
        Header = ""
    End Function

    Public Overrides Function Top(w, d, a)
        Dim Q As Integer = 1 * a
        Dim X As Double = (w * 10) - 32 - 1
        Dim Y As Double = (d * 10)
        Dim Z As Double = 19
        Dim Material As String = "Ply"
        Dim Edge As String = ""
        Dim EdgeCode As String = PMaterial(2)

        If (X > Y) Then
            Edge = "E1L"
        Else
            Edge = "E1S"
        End If

        Form1.DataGridView1.Rows.Add(1, "Top", Y, X, Z, Material, Edge, EdgeCode)
        Top = ""
    End Function

    Public Overrides Function Back(w, h, a)
        Dim Q As Integer = 1 * a
        Dim X As Double = ((h * 10) - (120 + 16 + 19)) + 10
        Dim Y As Double = (w * 10) - 23
        Dim Z As Double = 3
        Dim Material As String = BoxMaterial(0)

        Form1.DataGridView1.Rows.Add(Q, "Back", Y, X, Z, Material)
        Back = ""
    End Function

End Class

Public Class BaseRangedDrawer
    Inherits BaseRanged

    Public Overrides Function Header(w, h, d, a)
        Dim Code As String = w & "-" & h & "-" & d
        Dim CodeExt As String = Form1.CabCode.Text & "-" & Form1.Hardware
        Form1.DataGridView1.Rows.Add(CodeExt, "Qty", a, "", "", "", "")
        Form1.DataGridView1.Rows.Add(Code, "", "", "", "", "", "")
        Form1.DataGridView1.Rows.Add("Qty", "Part", "H", "W", "D", "Mat", "Edge", "Code", "Edge", "Code", "Notes")
        Header = ""
    End Function

End Class

Public Class BaseSink
    Inherits Base

    Public Overrides Function Header(w, h, d, a)
        Dim Code As String = w & "-" & h & "-" & d
        Form1.DataGridView1.Rows.Add("BS", "Qty", a, "", "", "", "")
        Form1.DataGridView1.Rows.Add(Code, "", "", "", "", "", "")
        Form1.DataGridView1.Rows.Add("Qty", "Part", "H", "W", "D", "Mat", "Edge", "Code", "Edge", "Code", "Notes")
        Header = ""
    End Function

End Class

Public Class BaseWastePullOut
    Inherits BaseDrawer

    Public Overrides Function Header(w, h, d, a)
        Dim Code As String = w & "-" & h & "-" & d
        Dim Cab As String = Form1.CabCode.Text & "-" & Form1.POHardware
        Form1.DataGridView1.Rows.Add(Cab, "Qty", a, "", "", "", "")
        Form1.DataGridView1.Rows.Add(Code, "", "", "", "", "", "")
        Form1.DataGridView1.Rows.Add("Qty", "Part", "H", "W", "D", "Mat", "Edge", "Code", "Edge", "Code", "Notes")
        Header = ""
    End Function

    Public Overridable Function Header2(w, h, d, a)
        Form1.DataGridView1.Rows.Add("Dr. Box", "Qty", a, "", "", "", "")
        Form1.DataGridView1.Rows.Add("Qty", "Part", "H", "W", "D", "Mat", "Edge", "Code", "Edge", "Code", "Notes")
        Return True
    End Function


    Public Overridable Function DBottomM(a)
        Dim Q As Integer = 1 * a
        Dim X As Double = 527
        Dim Y As Double = 359
        Dim Z As Double = 16
        Dim Material As String = BoxMaterial(0)

        Form1.DataGridView1.Rows.Add(Q, "Bottom", Y, X, Z, Material)
        Return True
    End Function

    Public Overridable Function DGableM(a)
        Dim Q As Integer = 2 * a
        Dim X As Double = 543
        Dim Y As Double = 270
        Dim Z As Double = 16
        Dim Material As String = BoxMaterial(0)
        Dim Edge As String = ""
        Dim EdgeCode As String = BoxMaterial(1)

        If (X > Y) Then
            Edge = "E1L"
        Else
            Edge = "E1S"
        End If

        Form1.DataGridView1.Rows.Add(Q, "Gable", Y, X, Z, Material, Edge, EdgeCode)
        Return True
    End Function

    Public Overridable Function DBackM(a)
        Dim Q As Integer = 1 * a
        Dim X As Double = 359
        Dim Y As Double = 270
        Dim Z As Double = 16
        Dim Material As String = BoxMaterial(0)
        Dim Edge As String = ""
        Dim EdgeCode As String = BoxMaterial(1)

        If (X > Y) Then
            Edge = "E1L"
        Else
            Edge = "E1S"
        End If

        Form1.DataGridView1.Rows.Add(Q, "Back", Y, X, Z, Material, Edge, EdgeCode)
        Return True
    End Function

    Public Overridable Function DPanelM(a)
        Dim Q As Integer = 1 * a
        Dim X As Double = 680
        Dim Y As Double = 391
        Dim Z As Double = 16
        Dim Material As String = BoxMaterial(0)
        Dim Edge As String = "E4S"
        Dim EdgeCode As String = BoxMaterial(1)

        Form1.DataGridView1.Rows.Add(Q, "Panel", Y, X, Z, Material, Edge, EdgeCode)
        Form1.DataGridView1.Rows.Add("")
        Return True
    End Function

End Class