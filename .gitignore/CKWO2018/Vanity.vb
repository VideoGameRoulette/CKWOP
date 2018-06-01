Public Class Vanity
    Inherits Base

    Public Overrides Function Header(w, h, d, a)
        Dim Code As String = w & "-" & h & "-" & d
        Form1.DataGridView1.Rows.Add("V", "Qty", 1, "", "", "", "")
        Form1.DataGridView1.Rows.Add(Code, "", "", "", "", "", "")
        Form1.DataGridView1.Rows.Add("Qty", "Part", "H", "W", "D", "Mat", "Edge", "Code", "Edge", "Code", "Notes")
        Header = ""
    End Function

End Class

Public Class VanitySink
    Inherits BaseSink

    Public Overrides Function Header(w, h, d, a)
        Dim Code As String = w & "-" & h & "-" & d
        Form1.DataGridView1.Rows.Add("VS", "Qty", 1, "", "", "", "")
        Form1.DataGridView1.Rows.Add(Code, "", "", "", "", "", "")
        Form1.DataGridView1.Rows.Add("Qty", "Part", "H", "W", "D", "Mat", "Edge", "Code", "Edge", "Code", "Notes")
        Header = ""
    End Function

End Class

Public Class VanityMI
    Inherits BaseMI

    Public Overrides Function Header(w, h, d, a)
        Dim Code As String = w & "-" & h & "-" & d
        Form1.DataGridView1.Rows.Add("VMI", "Qty", 1, "", "", "", "")
        Form1.DataGridView1.Rows.Add(Code, "", "", "", "", "", "")
        Form1.DataGridView1.Rows.Add("Qty", "Part", "H", "W", "D", "Mat", "Edge", "Code", "Edge", "Code", "Notes")
        Header = ""
    End Function

End Class

Public Class VanityElevated
    Inherits Vanity

    Public Overrides Function Header(w, h, d, a)
        Dim Code As String = w & "-" & h & "-" & d
        Form1.DataGridView1.Rows.Add("VE", "Qty", 1, "", "", "", "")
        Form1.DataGridView1.Rows.Add(Code, "", "", "", "", "", "")
        Form1.DataGridView1.Rows.Add("Qty", "Part", "H", "W", "D", "Mat", "Edge", "Code", "Edge", "Code", "Notes")
        Header = ""
    End Function

    Public Overrides Function Back(w, h, a)
        Dim X As Double = (h * 10) - 120 - 23
        Dim Y As Double = (w * 10) - 23
        Dim Z As Double = 3
        Form1.DataGridView1.Rows.Add(1, "Back", Y, X, Z, BoxMaterial(0))
        Back = ""
    End Function

End Class

Public Class VanityElevatedSink
    Inherits VanitySink

    Public Overrides Function Header(w, h, d, a)
        Dim Code As String = w & "-" & h & "-" & d
        Form1.DataGridView1.Rows.Add("VES", "Qty", 1, "", "", "", "")
        Form1.DataGridView1.Rows.Add(Code, "", "", "", "", "", "")
        Form1.DataGridView1.Rows.Add("Qty", "Part", "H", "W", "D", "Mat", "Edge", "Code", "Edge", "Code", "Notes")
        Header = ""
    End Function

End Class