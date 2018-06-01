Public Class Laundry
    Inherits Upper

    Public Overrides Function Header(w, h, d, a)
        Dim Code As String = w & "-" & h & "-" & d
        Form1.DataGridView1.Rows.Add(Form1.CabCode.Text, "Qty", a, "", "", "", "")
        Form1.DataGridView1.Rows.Add(Code, "", "", "", "", "", "")
        Form1.DataGridView1.Rows.Add("Qty", "Part", "H", "W", "D", "Mat", "Edge", "Code", "Edge", "Code", "Notes")
        Header = ""
    End Function

End Class
