Public Class StandardCabinet

    '#############################
    '# White and Hard Rock Maple #
    '#############################
    '
    Public Overridable Function Header(w, h, d, a)
        Dim Code As String = w & "-" & h & "-" & d
        Form1.DataGridView1.Rows.Add(Form1.CabCode.Text, "Qty", a, "", "", "", "")
        Form1.DataGridView1.Rows.Add(Code, "", "", "", "", "", "")
        Form1.DataGridView1.Rows.Add("Qty", "Part", "H", "W", "D", "Mat", "Edge", "Code", "Edge", "Code", "Notes")
        Header = ""
    End Function

    Public Overridable Function BoxMaterial()
        Dim boxMat(2) As String
        Select Case Form1.materialBox.SelectedIndex
            Case 0
                boxMat(0) = "W. Mel"
                boxMat(1) = "White"
            Case 1
                boxMat(0) = "HRM"
                boxMat(1) = "6116"
        End Select
        BoxMaterial = boxMat
    End Function

    '#####################
    '# VENEER MATERIAL 1 #
    '#####################
    '
    Public Overridable Function VMaterial()
        Dim vMat(4) As String

        Select Case Form1.speciesBox.Text
            Case "Cherry"
                vMat(0) = "Cherry"
                vMat(1) = "None"
                vMat(2) = "Solid Cherry"
                vMat(3) = "CH/V"
            Case "Maple"
                vMat(0) = "Maple"
                vMat(1) = "None"
                vMat(2) = "Solid Maple"
                vMat(3) = "M/V"
            Case "MDF"
                vMat(0) = "MDF"
                vMat(1) = "Maple"
                vMat(2) = "Solid Poplar"
                vMat(3) = "M/V"
            Case "Oak"
                vMat(0) = "Oak"
                vMat(1) = "None"
                vMat(2) = "Solid Oak"
                vMat(3) = "O/V"
            Case "Pine"
                vMat(0) = "Pine"
                vMat(1) = "None"
                vMat(2) = "Solid Pine"
                vMat(3) = "P/V"
            Case "Walnut"
                vMat(0) = "Walnut"
                vMat(1) = "None"
                vMat(2) = "Solid Walnut"
                vMat(3) = "W/V"
        End Select

        VMaterial = vMat
    End Function

    '#######################
    '# VENEER PVC MATERIAL #
    '#######################
    '
    Public Overridable Function PMaterial()
        Dim pMat(4) As String
        Select Case Form1.doorFinishBox.Text
            Case "Antique White PVC"
                pMat(0) = "Ant.W.Mel"
                pMat(1) = "Ant.W.PVC"
                pMat(2) = "1438"
                pMat(3) = "1438"
            Case "Bleached Maple PVC"
                pMat(0) = "Bl.M.Mel"
                pMat(1) = "Bl.M.PVC"
                pMat(2) = "9530"
                pMat(3) = "9530"
            Case "Charcoal Melamine"
                pMat(0) = "Char.Mel"
                pMat(1) = "Char.Mel"
                pMat(2) = "1315"
                pMat(3) = "1315"
            Case "Chocolate Maple PVC"
                pMat(0) = "Choc.Mel"
                pMat(1) = "Choc.PVC"
                pMat(2) = "6121"
                pMat(3) = "6121"
            Case "Honey Apple PVC"
                pMat(0) = "H.A.Mel"
                pMat(1) = "H.A.PVC"
                pMat(2) = "7216"
                pMat(3) = "7216"
            Case "Italian Walnut PVC"
                pMat(0) = "Itl.W.Mel"
                pMat(1) = "Itl.W.PVC"
                pMat(2) = "5233"
                pMat(3) = "5233"
            Case "Java Glow PVC"
                pMat(0) = "Java G.Mel"
                pMat(1) = "Java G.PVC"
                pMat(2) = "9513"
                pMat(3) = "9513"
            Case "Majestic Walnut PVC"
                pMat(0) = "Maj.W.Mel"
                pMat(1) = "Maj.W.PVC"
                pMat(2) = "5476"
                pMat(3) = "5476"
            Case "Mystic PVC"
                pMat(0) = "Mys.Mel"
                pMat(1) = "Mys.PVC"
                pMat(2) = "6464"
                pMat(3) = "6464"
            Case "Natural Maple PVC"
                pMat(0) = "Nat.M.Mel"
                pMat(1) = "Nat.M.PVC"
                pMat(2) = "6116"
                pMat(3) = "6116"
            Case "Pink Maple PVC"
                pMat(0) = "Pink.M.Mel"
                pMat(1) = "Pink.M.PVC"
                pMat(2) = "8567"
                pMat(3) = "8567"
            Case "Portland Cherry PVC"
                pMat(0) = "Por.Ch.Mel"
                pMat(1) = "Por.Ch.PVC"
                pMat(2) = "9774"
                pMat(3) = "9774"
            Case "Red Apple PVC"
                pMat(0) = "Red.A.Mel"
                pMat(1) = "Red.A.PVC"
                pMat(2) = "2668"
                pMat(3) = "2668"
            Case "Silken Maple PVC"
                pMat(0) = "Sil.M.Mel"
                pMat(1) = "Sil.M.PVC"
                pMat(2) = "5557"
                pMat(3) = "5557"
            Case "Silver PVC"
                pMat(0) = "Silv.Mel"
                pMat(1) = "Silv.PVC"
                pMat(2) = "151"
                pMat(3) = "151"
            Case "Vanilla Stix PVC"
                pMat(0) = "Van.St.Mel"
                pMat(1) = "Van.St.PVC"
                pMat(2) = "5942"
                pMat(3) = "5942"
            Case "White Ash PVC"
                pMat(0) = "W.Ash Mel"
                pMat(1) = "W.Ash PVC"
                pMat(2) = "White"
                pMat(3) = "White"
            Case "White Crystal PVC"
                pMat(0) = "555 White"
                pMat(1) = "W.Cr.PVC"
                pMat(2) = "White"
                pMat(3) = "White"
            Case "Black Semigloss"
                pMat(2) = "Black"
                pMat(3) = "Black"
            Case "Grey Light"
                pMat(2) = "105"
                pMat(3) = "M/V"
            Case "OC-26"
                pMat(2) = "112"
                pMat(3) = "M/V"
            Case "2111-60"
                pMat(2) = "190"
                pMat(3) = "M/V"
            Case "Caramel Maple", "Caramel Oak"
                pMat(2) = "205"
            Case "Fruitwood Oak"
                pMat(2) = "386"
            Case "Tulip Maple"
                pMat(2) = "1037"
            Case "Mango Maple"
                pMat(2) = "1117"
            Case "Charcoal Maple", "Charcoal Oak", "Chocolate Maple", "Chocolate Oak", "Ebony Maple", "Ebony Oak"
                pMat(2) = "1130"
            Case "Grey Stone Maple", "Grey Stone Oak"
                pMat(2) = "1156"
            Case "2124-30", "CSP-110"
                pMat(2) = "1315"
                pMat(3) = "M/V"
            Case "Antique White", "CC-490", "CREAM MATTE", "DOVE MATTE", "DOVE WHITE", "OC-20", "OC-46"
                pMat(2) = "1438"
                pMat(3) = "M/V"
            Case "Tulip Oak"
                pMat(2) = "2567"
            Case "Pecan Maple", "Pecan Oak"
                pMat(2) = "2668"
            Case "Ginger Maple", "Ginger Oak", "Mocha Maple", "Mocha Oak"
                pMat(2) = "5038"
            Case "Graphite Maple", "Graphite Oak", "Slate Maple", "Slate Oak"
                pMat(2) = "5741"
            Case "Natural Maple", "Nutmeg Maple"
                pMat(2) = "6116"
            Case "Butternut Maple", "Butternut Oak"
                pMat(2) = "6117"
            Case "Espresso Maple", "Espresso Oak"
                pMat(2) = "6119"
            Case "Aspen Maple", "Aspen Oak"
                pMat(2) = "6120"
            Case "Khaki Maple", "Khaki Oak"
                pMat(2) = "6120"
            Case "Brown Cherry Cherry", "Cocoa Maple", "Cocoa Oak", "Walnut Maple"
                pMat(2) = "6121"
            Case "Almond Matte", "Mushroom"
                pMat(2) = "7002"
                pMat(3) = "M/V"
            Case "2126-60", "2134-60", "Fog Grey"
                pMat(2) = "7006"
                pMat(3) = "M/V"
            Case "AF-95"
                pMat(2) = "7043"
                pMat(3) = "M/V"
            Case "Grey Dark"
                pMat(2) = "7071"
                pMat(3) = "M/V"
            Case "Caramel Pine"
                pMat(2) = "7806"
            Case "Natural Oak", "Nutmeg Oak"
                pMat(2) = "8001"
            Case "Peach Oak"
                pMat(2) = "8050"
            Case "Chestnut Cherry", "Noce Walnut", "Wild Walnut Maple"
                pMat(2) = "8103"
            Case "Antique Brown Maple", "Antique Brown Oak", "Cinnamon Cherry"
                pMat(2) = "8114"
            Case "Medium Brown Maple", "Medium Brown Oak"
                pMat(2) = "8336"
            Case "Harvest Maple"
                pMat(2) = "8597"
            Case "White Wash Oak"
                pMat(2) = "8670"
            Case "Olive Maple", "Olive Oak"
                pMat(2) = "8693"
            Case "Peach Maple", "White Wash Maple"
                pMat(2) = "8699"
            Case "Cognac Cherry"
                pMat(2) = "8725"
            Case "Natural Cherry"
                pMat(2) = "8876"
            Case "Harvest Oak"
                pMat(2) = "8883"
            Case "Burgundy Maple", "Burgundy Oak", "Rasberry Maple", "Rasberry Oak", "Red Cherry Cherry"
                pMat(2) = "9774"
            Case "Cafe Maple"
                pMat(2) = "20416"
            Case "Toffee Maple", "Toffee Oak"
                pMat(2) = "40439"
            Case "Pearl Semi-Gloss", "White Matte MDF"
                pMat(2) = "White"
                pMat(3) = "M/V"
            Case "CC-30", "CC-40", "OC-17", "OC-25", "OC-30", "OC-65"
                pMat(2) = "White"
                pMat(3) = "White"
            Case "Escarpment", "CC-518"
                pMat(2) = "7213"
                pMat(3) = "M/V"
        End Select
        PMaterial = pMat
    End Function

End Class

Public Class PremierCabinet
    Inherits StandardCabinet
End Class

Public Class ClassicCabinet
    Inherits StandardCabinet
End Class

Public Class EuroPremierCabinet
    Inherits PremierCabinet
End Class

Public Class EuroClassicCabinet
    Inherits ClassicCabinet
End Class