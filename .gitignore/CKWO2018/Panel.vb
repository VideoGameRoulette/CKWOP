Public Class OvenPanel
    Inherits StandardCabinet

    Public Overrides Function Header(w, h, d, a)
        Dim Code As String = w & "-" & h & "-" & d
        Form1.DataGridView1.Rows.Add(Form1.CabCode.Text, "Qty", 1, "", "", "", "")
        Form1.DataGridView1.Rows.Add(Code, "", "", "", "", "", "")
        Form1.DataGridView1.Rows.Add("Qty", "Part", "H", "W", "D", "Mat", "Edge", "Code", "Edge", "Code", "Notes")
        Header = ""
    End Function

    Public Overridable Function OVP(w, h, d, a)

        Dim X As Double
        Dim Y As Double = w
        Dim Z As Double = d
        Dim Material As String
        Dim Edge As String = ""
        Dim EdgeCode As String = ""
        Dim Note As String = ""

        If (Form1.speciesBox.Text = "PVC") Then
            Material = PMaterial(0)
            EdgeCode = PMaterial(2)
        Else
            Material = VMaterial(0)
            EdgeCode = VMaterial(3)
        End If

        If (Form1.groupBox.Text = "Group 1") Then
            If (Form1.speciesBox.Text = "MDF") Then
                Select Case Form1.doorStyleBox.Text
                    Case "Aston", "Bristol", "Nottingham", "Paris", "Preston", "Siena", "Sonoma"
                        X = ""
                        Y = ""
                        Edge = ""
                        EdgeCode = ""
                        Note = "Ordered Joe G."
                    Case Else
                        X = h
                        Y = w
                        Edge = ""
                        EdgeCode = ""
                        Note = ""
                End Select
            ElseIf (Form1.speciesBox.Text = "PVC") Then
                X = ""
                Y = ""
                Edge = ""
                EdgeCode = ""
                Note = "Ordered Joe G."
            Else
                X = h - 1
                Y = w - 1
                Edge = "E4S"
                EdgeCode = ""
                Note = ""
            End If
        ElseIf (Form1.groupBox.Text = "Group 2") Then
            If (Form1.speciesBox.Text = "MDF") Then
                Select Case Form1.doorStyleBox.Text
                    Case "Aston", "Bristol", "Nottingham", "Paris", "Preston", "Siena", "Sonoma"
                        X = ""
                        Y = ""
                        Edge = ""
                        EdgeCode = ""
                        Note = "Ordered Joe G."
                    Case Else
                        X = h
                        Y = w
                        Edge = ""
                        EdgeCode = ""
                        Note = ""
                End Select
            ElseIf (Form1.speciesBox.Text = "PVC") Then
                X = ""
                Y = ""
                Edge = ""
                EdgeCode = ""
                Note = "Ordered Joe G."
            Else
                X = h
                Y = w
                Edge = ""
                EdgeCode = ""
                Note = ""
            End If
        End If

        Form1.DataGridView1.Rows.Add(2, "OVP1", Y, X, Z, Material, Edge, EdgeCode, "", "", Note)
        Form1.DataGridView1.Rows.Add("")
        OVP = ""
    End Function

End Class