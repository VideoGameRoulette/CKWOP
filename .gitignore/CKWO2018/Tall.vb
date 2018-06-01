Public Class TallUpper
    Inherits Upper
    Public Overrides Function Header(w, h, d, a)
        Dim Code As String = w & "-" & h & "-" & d
        Form1.DataGridView1.Rows.Add("U", "Qty", 1, "", "", "", "")
        Form1.DataGridView1.Rows.Add(Code, "", "", "", "", "", "")
        Form1.DataGridView1.Rows.Add("Qty", "Part", "H", "W", "D", "Mat", "Edge", "Code", "Edge", "Code", "Notes")
        Header = ""
    End Function
End Class

Public Class TallBase
    Inherits Base
    Public Overrides Function Header(w, h, d, a)
        Dim Code As String = w & "-" & h & "-" & d
        Form1.DataGridView1.Rows.Add("B", "Qty", 1, "", "", "", "")
        Form1.DataGridView1.Rows.Add(Code, "", "", "", "", "", "")
        Form1.DataGridView1.Rows.Add("Qty", "Part", "H", "W", "D", "Mat", "Edge", "Code", "Edge", "Code", "Notes")
        Header = ""
    End Function
End Class

Public Class TallBaseDrawer
    Inherits BaseDrawer
    Public Overrides Function Header(w, h, d, a)
        Dim Code As String = w & "-" & h & "-" & d
        Form1.DataGridView1.Rows.Add("B" & Form1.CabCode3.Text, "Qty", 1, "", "", "", "")
        Form1.DataGridView1.Rows.Add(Code, "", "", "", "", "", "")
        Form1.DataGridView1.Rows.Add("Qty", "Part", "H", "W", "D", "Mat", "Edge", "Code", "Edge", "Code", "Notes")
        Header = ""
    End Function
End Class

Public Class TallUpperMI
    Inherits UpperMI
    Public Overrides Function Header(w, h, d, a)
        Dim Code As String = w & "-" & h & "-" & d
        Form1.DataGridView1.Rows.Add("UMI", "Qty", 1, "", "", "", "")
        Form1.DataGridView1.Rows.Add(Code, "", "", "", "", "", "")
        Form1.DataGridView1.Rows.Add("Qty", "Part", "H", "W", "D", "Mat", "Edge", "Code", "Edge", "Code", "Notes")
        Header = ""
    End Function
End Class

Public Class TallBaseMI
    Inherits BaseMI
    Public Overrides Function Header(w, h, d, a)
        Dim Code As String = w & "-" & h & "-" & d
        Form1.DataGridView1.Rows.Add("BMI", "Qty", 1, "", "", "", "")
        Form1.DataGridView1.Rows.Add(Code, "", "", "", "", "", "")
        Form1.DataGridView1.Rows.Add("Qty", "Part", "H", "W", "D", "Mat", "Edge", "Code", "Edge", "Code", "Notes")
        Header = ""
    End Function
End Class