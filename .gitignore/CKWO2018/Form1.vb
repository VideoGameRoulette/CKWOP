Imports System.IO
Imports System.Drawing.Printing


Public Class Form1

    '#####################
    '# PRIVATE VARIABLES #
    '#####################
    '

    '####################
    '# PUBLIC VARIABLES #
    '####################
    '
    Public wBox As Double
    Public hBox As Double
    Public dBox As Double
    Public hBox2 As Double
    Public dBox2 As Double
    Public hBox3 As Double
    Public dBox3 As Double
    Public aBox As Integer
    Public Hardware As String = ""
    Public POHardware As String
    Public headCheck As Boolean = False
    Public Department As String = ""
    Public AppPath As String = System.Windows.Forms.Application.StartupPath
    Public WOPath As String = AppPath & "\WorkOrders\"
    Public InfoPath As String = AppPath & "\Info\"
    Public RoomPath As String = WOPath
    Public PrgmPath As String = AppPath & "\Programs\"
    Public TempPath As String = AppPath & "\Templates\"
    'Public PrgmPath As String = "M:\CortinaPrograms\Programs\"
    'Public TempPath As String = "M:\CortinaPrograms\Templates\"

    '#######################
    '# ON APPLICATION LOAD #
    '#######################
    '
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        Timer1.Start()
        Lines.LineCount()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim totalHeight As Double

        Try
            hBox = heightBox.Text
            hBox2 = heightBox2.Text
            hBox3 = heightBox3.Text
            totalHeight = hBox + hBox2 + hBox3
        Catch ex As Exception
        End Try

        If (hardwareBox.Text = "Multitech") Then
            Hardware = "M"
        ElseIf (hardwareBox.Text = "Innotech") Then
            Hardware = "I"
        End If

        If (CabCode2.Text = "517173100") Then
            POHardware = "517173100"
        ElseIf (CabCode2.Text = "517174100") Then
            POHardware = "517174100"
        ElseIf (CabCode2.Text = "361412100") Then
            POHardware = "361412100"
        ElseIf (CabCode2.Text = "4614100") Then
            POHardware = "4614100"
        ElseIf (CabCode2.Text = "461460100") Then
            POHardware = "461460100"
        ElseIf (CabCode2.Text = "6405-30") Then
            POHardware = "6405-30-" & Hardware
        End If

        Select Case CabCode.Text
            Case "B1D", "B2D", "B3D", "B4D", "BR2D"
                lblOutput.Text = CabCode.Text & "-" & Hardware & "-" & widthBox.Text & "-" & heightBox.Text & "-" & depthBox.Text
                lblOutput2.Text = CabCode.Text & "-" & Hardware & "-TEMPLATE"
                lblOutput3.Text = CabCode.Text & "-" & Hardware & "-" & widthBox.Text & "-" & heightBox.Text & "-" & depthBox.Text & "-TEMPLATE"
            Case "BSPO", "BWPO"
                lblOutput.Text = CabCode.Text & "-" & POHardware & "-" & widthBox.Text & "-" & heightBox.Text & "-" & depthBox.Text
                lblOutput2.Text = CabCode.Text & "-" & POHardware & "-TEMPLATE"
                lblOutput3.Text = CabCode.Text & "-" & POHardware & "-" & widthBox.Text & "-" & heightBox.Text & "-" & depthBox.Text & "-TEMPLATE"
            Case "TIA"
                Select Case CabCode3.Text
                    Case "1D", "2D", "3D", "4D"
                        lblOutput.Text = CabCode.Text & "-" & CabCode2.Text & "-" & CabCode3.Text & "-" & Hardware & "-" & widthBox.Text & "-" & totalHeight & "-" & depthBox.Text
                        lblOutput2.Text = CabCode.Text & "-" & CabCode2.Text & "-" & CabCode3.Text & "-" & Hardware & "-TEMPLATE"
                        lblOutput3.Text = CabCode.Text & "-" & CabCode2.Text & "-" & CabCode3.Text & "-" & Hardware & "-" & widthBox.Text & "-" & totalHeight & "-" & depthBox.Text & "-TEMPLATE"
                    Case Else
                        lblOutput.Text = CabCode.Text & "-" & CabCode2.Text & "-" & CabCode3.Text & "-" & widthBox.Text & "-" & totalHeight & "-" & depthBox.Text
                        lblOutput2.Text = CabCode.Text & "-" & CabCode2.Text & "-" & CabCode3.Text & "-TEMPLATE"
                        lblOutput3.Text = CabCode.Text & "-" & CabCode2.Text & "-" & CabCode3.Text & "-" & widthBox.Text & "-" & totalHeight & "-" & depthBox.Text & "-TEMPLATE"
                End Select
            Case Else
                lblOutput.Text = CabCode.Text & "-" & widthBox.Text & "-" & heightBox.Text & "-" & depthBox.Text
                lblOutput2.Text = CabCode.Text & "-TEMPLATE"
                lblOutput3.Text = CabCode.Text & "-" & widthBox.Text & "-" & heightBox.Text & "-" & depthBox.Text & "-TEMPLATE"
        End Select

    End Sub

    '###############################################################################################################
    '# FUNCTIONS ###################################################################################################
    '###############################################################################################################
    '
    '######################
    '# SEND DATA FUNCTION #
    '######################
    '
    Public Function SendData(ByVal a As String)
        Department = a
        WOPath = AppPath & "\WorkOrders\" & WONum.Text & "\" & Department
        RoomPath = WOPath & roomBox.Text & "\"
        Try
            If Not Directory.Exists(RoomPath) Then
                Directory.CreateDirectory(RoomPath)
            End If
        Catch error8 As Exception
        End Try
        wBox = widthBox.Text
        hBox = heightBox.Text
        dBox = depthBox.Text
        hBox2 = heightBox2.Text
        dBox2 = depthBox2.Text
        hBox3 = heightBox3.Text
        dBox3 = depthBox3.Text
        aBox = amountBox.Text

        If (headCheck = False) Then
            DataGridView1.Rows.Add("", "", "", "", "", "", "", "", "", Label2.Text, WONum.Text)
            DataGridView1.Rows.Add(roomBox.Text, "", "", "", "", "", "", "", "", "", LineCode.Text)
            headCheck = True
        End If

        Select Case CabCode.Text.ToUpper
            Case "B"
                Base()
            Case "B1D", "B2D", "B3D", "B4D"
                BaseDrawer()
            Case "BBD"
                BaseBottomDrawer()
            Case "BMI"
                BaseMI()
            Case "BR", "BR2D"
                BaseRanged()
            Case "BS"
                BaseSink()
            Case "BSPO"
                BaseSpicePullOut()
            Case "BTD"
                BaseTopDrawer()
            Case "BTDD"
                BaseTopDoubleDrawer()
            Case "BWPO"
                BaseWastePullOut()
            Case "HOOD"
                CanopyHood()
            Case "OVP1", "OVP2", "OVP3"
                OV()
            Case "RHAW-S", "RHAW-D", "RHAW-T"
                Canopy()
            Case "T"
                Tall()
            Case "TIA"
                TallIntegratedAppliance()
            Case "TMI"
                TallMI()
            Case "TU"
                TallUtility()
            Case "U"
                Upper()
            Case "UCD"
                UpperCornerDiagonal()
            Case "UCDMI"
                UpperCornerDiagonalMI()
            Case "UM"
                UpperMicrowave()
            Case "UMI"
                UpperMI()
            Case "UTD"
                UpperTrayDivider()
            Case "V"
                Vanity()
            Case "VE"
                VanityElevated()
            Case "VES"
                VanityElevatedSink()
            Case "VMI"
                VanityMI()
            Case "VS"
                VanitySink()
            Case Else
                Return False
        End Select
        'Functions.CopyDirectory()
        'Functions.ClearFields()
        Return True
    End Function

    '#######################
    '# SET FIELDS FUNCTION #
    '#######################
    '
    Public Function SetField()

        CabCode2.Items.Clear()

        heightBox2.Visible = False
        depthBox2.Visible = False
        heightBox3.Visible = False
        depthBox3.Visible = False

        CabCode2.Visible = False
        CabCode3.Visible = False

        OverrideBase.Visible = False
        lblVEBackZ.Visible = False
        veBackZ.Visible = False

        OverrideHood.Visible = False

        Select Case CabCode.Text
            Case "B", "B1D", "B2D", "B3D", "B4D", "BTD"
                OverrideBase.Visible = True
            Case "BSPO"
                CabCode2.Items.Clear()
                CabCode2.Items.AddRange(IO.File.ReadAllLines(InfoPath & "BSPOHardware.txt"))
                CabCode2.Visible = True
            Case "BWPO"
                CabCode2.Items.Clear()
                CabCode2.Items.AddRange(IO.File.ReadAllLines(InfoPath & "BWPOHardware.txt"))
                CabCode2.Visible = True
            Case "TIA"
                CabCode2.Items.Clear()
                CabCode2.Items.AddRange(IO.File.ReadAllLines(InfoPath & "TIAOptions.txt"))
                heightBox2.Visible = True
                depthBox2.Visible = True
                heightBox3.Visible = True
                depthBox3.Visible = True
                CabCode2.Visible = True
                CabCode3.Visible = True
            Case "RHAW-S", "RHAW-D", "RHAW-T"
                CabCode2.Items.Clear()
                heightBox2.Visible = True
                depthBox2.Visible = True
            Case "HOOD"
                CabCode2.Items.Clear()
                heightBox2.Visible = True
                depthBox2.Visible = True
                OverrideHood.Visible = True
            Case "VE", "VES"
                CabCode2.Items.Clear()
                OverrideBase.Visible = True
                lblVEBackZ.Visible = True
                veBackZ.Visible = True
            Case Else
        End Select
        SetField = ""
    End Function

    '##########################
    '# HEADER OUTPUT FUNCTION #
    '##########################
    '
    Public Function Header()
        Dim Size As String = wBox & "-" & (hBox + hBox2 + hBox3) & "-" & dBox
        DataGridView1.Rows.Add(CabCode.Text, "Qty", aBox, "", "", "", "")
        DataGridView1.Rows.Add(Size, "", "", "", "", "", "")
        Return True
    End Function

    '########################
    '# BASE OUTPUT FUNCTION #
    '########################
    '
    Public Function Base()

        Dim Parts As New Base
        Parts.Header(wBox, hBox, dBox, aBox)
        Parts.Back(wBox, hBox, aBox)
        Parts.TopStrap(wBox, dBox, aBox)
        Parts.Bottom(wBox, dBox, aBox)
        Parts.AdjSh(wBox, dBox, aBox)
        Parts.Gable(hBox, dBox, aBox)

        Return True
    End Function

    '###############################
    '# BASE DRAWER OUTPUT FUNCTION #
    '###############################
    '
    Public Function BaseDrawer()

        Dim Parts As New BaseDrawer
        Parts.Header(wBox, hBox, dBox, aBox)
        Parts.Back(wBox, hBox, aBox)
        Parts.TopStrap(wBox, dBox, aBox)
        Parts.Bottom(wBox, dBox, aBox)
        Parts.Gable(hBox, dBox, aBox)

        Return True
    End Function

    '###############################
    '# BASE RANGED OUTPUT FUNCTION #
    '###############################
    '
    Public Function BaseRanged()

        Dim Parts As New BaseRanged
        Parts.Header(wBox, hBox, dBox, aBox)
        Parts.Back(wBox, hBox, aBox)
        Parts.Top(wBox, dBox, aBox)
        Parts.Bottom(wBox, dBox, aBox)
        Parts.Gable(hBox, dBox, aBox)

        Return True
    End Function

    '######################################
    '# BASE RANGED DRAWER OUTPUT FUNCTION #
    '######################################
    '
    Public Function BaseRangedDrawer()

        Dim Parts As New BaseRangedDrawer
        Parts.Header(wBox, hBox, dBox, aBox)
        Parts.Back(wBox, hBox, aBox)
        Parts.Top(wBox, dBox, aBox)
        Parts.Bottom(wBox, dBox, aBox)
        Parts.Gable(hBox, dBox, aBox)

        Return True
    End Function

    '#############################
    '# BASE SINK OUTPUT FUNCTION #
    '#############################
    '
    Public Function BaseSink()

        Dim Parts As New BaseSink
        Parts.Header(wBox, hBox, dBox, aBox)
        Parts.Back(wBox, hBox, aBox)
        Parts.SinkStrap(wBox, dBox, aBox)
        Parts.TopStrap(wBox, dBox, aBox)
        Parts.Bottom(wBox, dBox, aBox)
        Parts.AdjSh(wBox, dBox, aBox)
        Parts.Gable(hBox, dBox, aBox)

        Return True
    End Function

    '#######################################
    '# BASE SPICE PULL OUT OUTPUT FUNCTION #
    '#######################################
    '
    Public Function BaseSpicePullOut()

        Dim Parts As New BaseDrawer
        Parts.Header(wBox, hBox, dBox, aBox)
        'Parts.Back(wBox, hBox, aBox)
        'Parts.TopStrap(wBox, dBox, aBox)
        'Parts.Bottom(wBox, dBox, aBox)
        Parts.Gable(hBox, dBox, aBox)

        Return True
    End Function

    '#######################################
    '# BASE WASTE PULL OUT OUTPUT FUNCTION #
    '#######################################
    '
    Public Function BaseWastePullOut()

        Dim Parts As New BaseWastePullOut
        If (materialBox.Text = "White Melamine Box") Then
            Parts.Header(wBox, hBox, dBox, aBox)
            Parts.Gable(hBox, dBox, aBox)
        Else
            Parts.Header(wBox, hBox, dBox, aBox)
            Parts.Back(wBox, hBox, aBox)
            Parts.TopStrap(wBox, dBox, aBox)
            Parts.Bottom(wBox, dBox, aBox)
            Parts.Gable(hBox, dBox, aBox)
        End If

        If (CabCode2.Text = "6405-30") Then
            Parts.Header2(wBox, hBox, dBox, aBox)
            Parts.DBottomM(aBox)
            Parts.DGableM(aBox)
            Parts.DBackM(aBox)
            Parts.DPanelM(aBox)
        End If

        Return True
    End Function

    '###################################
    '# BASE TOP DRAWER OUTPUT FUNCTION #
    '###################################
    '
    Public Function BaseTopDrawer()

        Dim Parts As New BaseDrawer
        Parts.Header(wBox, hBox, dBox, aBox)
        Parts.Back(wBox, hBox, aBox)
        Parts.TopStrap(wBox, dBox, aBox)
        Parts.Strap(wBox, dBox, aBox)
        Parts.Bottom(wBox, dBox, aBox)
        Parts.Gable(hBox, dBox, aBox)

        Return True
    End Function

    '######################################
    '# BASE BOTTOM DRAWER OUTPUT FUNCTION #
    '######################################
    '
    Public Function BaseBottomDrawer()

        Dim Parts As New BaseDrawer
        Parts.Header(wBox, hBox, dBox, aBox)
        Parts.Back(wBox, hBox, aBox)
        Parts.TopStrap(wBox, dBox, aBox)
        Parts.Strap(wBox, dBox, aBox)
        Parts.Bottom(wBox, dBox, aBox)
        Parts.Gable(hBox, dBox, aBox)

        Return True
    End Function

    '##########################################
    '# BASE TOP DOUBLE DRAWER OUTPUT FUNCTION #
    '##########################################
    '
    Public Function BaseTopDoubleDrawer()

        Dim Parts As New BaseDrawer
        Parts.Header(wBox, hBox, dBox, aBox)
        Parts.Back(wBox, hBox, aBox)
        Parts.Top(wBox, dBox, aBox)
        Parts.Bottom(wBox, dBox, aBox)
        Parts.Strap(wBox, dBox, aBox)
        Parts.Divider(dBox, aBox)
        Parts.Gable(hBox, dBox, aBox)

        Return True
    End Function

    '##########################################
    '# BASE MATCHING INTERIOR OUTPUT FUNCTION #
    '##########################################
    '
    Public Function BaseMI()

        Dim Parts As New BaseMI
        Parts.Header(wBox, hBox, dBox, aBox)
        Parts.Back(wBox, hBox, aBox)
        Parts.TopBtm(wBox, dBox, aBox)
        Parts.AdjSh(wBox, dBox, aBox)
        Parts.Gable(hBox, dBox, aBox)

        Return True
    End Function

    '##########################
    '# CANOPY OUTPUT FUNCTION #
    '##########################
    '
    Public Function Canopy()
        Dim Parts As New Canopy
        Parts.Header(wBox, hBox, dBox, aBox)
        Parts.Back(wBox, hBox, dBox, aBox)
        Parts.Top(wBox, dBox, aBox)
        Parts.Gable(hBox, dBox, aBox)
        Parts.Front(wBox, hBox, aBox)
        Parts.FanShelf(wBox, dBox2, aBox)
        Parts.SmallFront(wBox, hBox2, aBox)
        Parts.SmallTop(wBox, dBox, dBox2, aBox)
        Parts.InsertPanel(wBox, hBox, hBox2, aBox)
        Return True
    End Function

    '###############################
    '# CANOPY HOOD OUTPUT FUNCTION #
    '###############################
    '
    Public Function CanopyHood()
        Dim Parts As New CanopyHood
        Parts.Header(wBox, hBox, dBox, aBox)
        Parts.Back(wBox, hBox, dBox, aBox)
        Parts.Top(wBox, dBox, aBox)
        Parts.Gable(hBox, dBox, aBox)
        Parts.Front(wBox, hBox, aBox)
        Parts.FanShelf(wBox, dBox2, aBox)
        Return True
    End Function

    '###############################
    '# CANOPY HOOD OUTPUT FUNCTION #
    '###############################
    '
    Public Function Laundry()

        Dim Parts As New Laundry
        Parts.Header(wBox, hBox, dBox, aBox)
        Parts.Back(wBox, hBox, aBox)
        Parts.TopBtm(wBox, dBox, aBox)
        Parts.AdjSh(wBox, hBox, dBox, aBox)
        Parts.Gable(hBox, dBox, aBox)

        Return True
    End Function

    '##############################
    '# OVER PANEL OUTPUT FUNCTION #
    '##############################
    '
    Public Function OV()

        Dim Parts As New OvenPanel
        Parts.Header(wBox, hBox, dBox, aBox)
        Parts.OVP(wBox, hBox, dBox, aBox)
        Return True

    End Function

    '########################
    '# TALL OUTPUT FUNCTION #
    '########################
    '
    Public Function Tall()

        Dim BParts As New Base
        BParts.Header(wBox, hBox, dBox, aBox)
        BParts.Back(wBox, hBox, aBox)
        BParts.TopBtm(wBox, dBox, aBox)
        BParts.AdjSh(wBox, dBox, aBox)
        BParts.Gable(hBox, dBox, aBox)

        Return True
    End Function

    '##########################################
    '# TALL MATCHING INTERIOR OUTPUT FUNCTION #
    '##########################################
    '
    Public Function TallMI()

        Dim Parts As New BaseMI
        Parts.Header(wBox, hBox, dBox, aBox)
        Parts.Back(wBox, hBox, aBox)
        Parts.TopBtm(wBox, dBox, aBox)
        Parts.AdjSh(wBox, dBox, aBox)
        Parts.Gable(hBox, dBox, aBox)

        Return True
    End Function

    '##########################################
    '# TALL UTILITY (2 UNITS) OUTPUT FUNCTION #
    '##########################################
    '
    Public Function TallUtility()
        Header()
        Dim PartsA As New TallUpper
        PartsA.Header(wBox, hBox, dBox, aBox)
        PartsA.Back(wBox, hBox, aBox)
        PartsA.Top(wBox, dBox, aBox)
        PartsA.Bottom(wBox, dBox, aBox)
        PartsA.AdjSh(wBox, hBox, dBox, aBox)
        PartsA.Gable(hBox, dBox, aBox)

        Dim PartsB As New TallBase
        PartsB.Header(wBox, hBox2, dBox2, aBox)
        PartsB.Back(wBox, hBox2, aBox)
        PartsB.Top(wBox, dBox2, aBox)
        PartsB.Bottom(wBox, dBox2, aBox)
        PartsB.AdjSh(wBox, dBox2, aBox)
        PartsB.Gable(hBox2, dBox2, aBox)

        Return True
    End Function

    '###########################################################
    '# TALL UTILITY MATCHING INTERIOR (2 UNIT) OUTPUT FUNCTION #
    '###########################################################
    '
    Public Function TallUtilityMI()
        Dim PartsA As New UpperMI
        PartsA.Header(wBox, hBox, dBox, aBox)
        PartsA.Back(wBox, hBox, aBox)
        PartsA.TopBtm(wBox, dBox, aBox)
        PartsA.AdjSh(wBox, hBox, dBox, aBox)
        PartsA.Gable(hBox, dBox, aBox)

        Dim Parts As New TallBaseMI
        Parts.Header(wBox, hBox2, dBox2, aBox)
        Parts.Back(wBox, hBox2, aBox)
        Parts.TopBtm(wBox, dBox2, aBox)
        Parts.AdjSh(wBox, dBox2, aBox)
        Parts.Gable(hBox2, dBox2, aBox)

        Return True
    End Function

    '#######################################################
    '# TALL INTEGRATED APPLIANCE (3 UNITS) OUTPUT FUNCTION #
    '#######################################################
    '
    Public Function TallIntegratedAppliance()
        Header()
        Dim PartsA As New TallUpper
        PartsA.Header(wBox, hBox, dBox, aBox)
        PartsA.Back(wBox, hBox, aBox)
        PartsA.Top(wBox, dBox, aBox)
        PartsA.Bottom(wBox, dBox, aBox)
        PartsA.AdjSh(wBox, hBox, dBox, aBox)
        PartsA.Gable(hBox, dBox, aBox)

        Dim PartsB As New OpenAppliance
        PartsB.Header(wBox, hBox2, dBox2, aBox)
        PartsB.Back(wBox, hBox2, aBox)
        PartsB.TopBtm(wBox, dBox2, aBox)
        PartsB.Gable(hBox2, dBox2, aBox)

        Select Case CabCode3.Text
            Case "B1D", "B2D", "B3D", "B4D"
                Dim PartsC As New BaseDrawer
                PartsC.Header(wBox, hBox3, dBox3, aBox)
                PartsC.Back(wBox, hBox3, aBox)
                PartsC.TopStrap(wBox, dBox3, aBox)
                PartsC.Bottom(wBox, dBox3, aBox)
                PartsC.AdjSh(wBox, dBox3, aBox)
                PartsC.Gable(hBox3, dBox3, aBox)
            Case Else
                Dim PartsC As New Base
                PartsC.Header(wBox, hBox3, dBox3, aBox)
                PartsC.Back(wBox, hBox3, aBox)
                PartsC.TopStrap(wBox, dBox3, aBox)
                PartsC.Bottom(wBox, dBox3, aBox)
                PartsC.AdjSh(wBox, dBox3, aBox)
                PartsC.Gable(hBox3, dBox3, aBox)
        End Select

        Return True
    End Function

    '#########################
    '# UPPER OUTPUT FUNCTION #
    '#########################
    '
    Public Function Upper()

        Dim Parts As New Upper
        Parts.Header(wBox, hBox, dBox, aBox)
        Parts.Back(wBox, hBox, aBox)
        Parts.TopBtm(wBox, dBox, aBox)
        Parts.AdjSh(wBox, hBox, dBox, aBox)
        Parts.Gable(hBox, dBox, aBox)

        Return True
    End Function

    '#########################################
    '# UPPER CORNER DIAGONAL OUTPUT FUNCTION #
    '#########################################
    '
    Public Function UpperCornerDiagonal()

        Dim Parts As New UpperCornerDiagonal
        Parts.Header(wBox, hBox, dBox, aBox)
        Parts.Back(wBox, hBox, aBox)
        Parts.BackStrap(hBox, aBox)
        Parts.TopBtm(wBox, dBox, aBox)
        Parts.AdjSh(wBox, hBox, dBox, aBox)
        Parts.Gable(hBox, dBox, aBox)

        Return True
    End Function

    '###########################################################
    '# UPPER CORNER DIAGONAL MATCHING INTERIOR OUTPUT FUNCTION #
    '###########################################################
    '
    Public Function UpperCornerDiagonalMI()

        Dim Parts As New UpperCornerDiagonalMI
        Parts.Header(wBox, hBox, dBox, aBox)
        Parts.LargeBack(wBox, hBox, aBox)
        Parts.SmallBack(wBox, hBox, aBox)
        Parts.TopBtm(wBox, dBox, aBox)
        Parts.AdjSh(wBox, hBox, dBox, aBox)
        Parts.Gable(hBox, dBox, aBox)

        Return True
    End Function

    '###########################
    '# UPPERMI OUTPUT FUNCTION #
    '###########################
    '
    Public Function UpperMI()

        Dim Parts As New UpperMI
        Parts.Header(wBox, hBox, dBox, aBox)
        Parts.Back(wBox, hBox, aBox)
        Parts.TopBtm(wBox, dBox, aBox)
        Parts.AdjSh(wBox, hBox, dBox, aBox)
        Parts.Gable(hBox, dBox, aBox)

        Return True
    End Function

    '###########################
    '# UPPER MICROWAVE OUTPUT FUNCTION #
    '###########################
    '
    Public Function UpperMicrowave()
        Dim topHeight As Double
        Dim bottomHeight As Double

        Select Case hBox
            Case "100"
                topHeight = 55
                bottomHeight = 45
            Case Else
                topHeight = hBox
                bottomHeight = hBox2
        End Select

        Header()

        Dim PartsA As New Upper
        PartsA.Header(wBox, topHeight, dBox, aBox)
        PartsA.Back(wBox, topHeight, aBox)
        PartsA.TopBtm(wBox, dBox, aBox)
        PartsA.AdjSh(wBox, topHeight, dBox, aBox)
        PartsA.Gable(topHeight, dBox, aBox)

        Dim PartsB As New UpperMicrowave
        PartsB.Header(wBox, bottomHeight, dBox, aBox)
        PartsB.Back(wBox, bottomHeight, aBox)
        PartsB.TopBtm(wBox, dBox, aBox)
        PartsB.MicrowaveShelf(wBox, aBox)
        PartsB.Gable(bottomHeight, dBox, aBox)

        Return True
    End Function

    '#######################################
    '# UPPER TRAY DIVIDERS OUTPUT FUNCTION #
    '#######################################
    '
    Public Function UpperTrayDivider()

        Dim Parts As New UpperTrayDivider
        Parts.Header(wBox, hBox, dBox, aBox)
        Parts.Back(wBox, hBox, aBox)
        Parts.TopBtm(wBox, dBox, aBox)
        Parts.Divider(hBox, dBox, aBox)
        Parts.Gable(hBox, dBox, aBox)

        Return True
    End Function

    '##########################
    '# VANITY OUTPUT FUNCTION #
    '##########################
    '
    Public Function Vanity()

        Dim Parts As New Vanity
        Parts.Header(wBox, hBox, dBox, aBox)
        Parts.Back(wBox, hBox, aBox)
        Parts.TopStrap(wBox, dBox, aBox)
        Parts.Bottom(wBox, dBox, aBox)
        Parts.Gable(hBox, dBox, aBox)

        Return True
    End Function

    '###################################
    '# VANITY ELEVATED OUTPUT FUNCTION #
    '###################################
    '
    Public Function VanityElevated()

        Dim Parts As New VanityElevated
        Parts.Header(wBox, hBox, dBox, aBox)
        Parts.Back(wBox, hBox, aBox)
        Parts.TopStrap(wBox, dBox, aBox)
        Parts.Bottom(wBox, dBox, aBox)
        Parts.Gable(hBox, dBox, aBox)

        Return True
    End Function

    '###############################
    '# VANITY SINK OUTPUT FUNCTION #
    '###############################
    '
    Public Function VanitySink()

        Dim Parts As New VanitySink
        Parts.Header(wBox, hBox, dBox, aBox)
        Parts.Back(wBox, hBox, aBox)
        Parts.TopStrap(wBox, dBox, aBox)
        Parts.Bottom(wBox, dBox, aBox)
        Parts.Gable(hBox, dBox, aBox)

        Return True
    End Function

    '########################################
    '# VANITY ELEVATED SINK OUTPUT FUNCTION #
    '########################################
    '
    Public Function VanityElevatedSink()

        Dim Parts As New VanityElevatedSink
        Parts.Header(wBox, hBox, dBox, aBox)
        Parts.Back(wBox, hBox, aBox)
        Parts.TopStrap(wBox, dBox, aBox)
        Parts.Bottom(wBox, dBox, aBox)
        Parts.Gable(hBox, dBox, aBox)

        Return True
    End Function

    '############################
    '# VANITY MATCHING INTERIOR #
    '############################
    '
    Public Function VanityMI()

        Dim Parts As New VanityMI
        Parts.Header(wBox, hBox, dBox, aBox)
        Parts.Back(wBox, hBox, aBox)
        Parts.TopBtm(wBox, dBox, aBox)
        Parts.Gable(hBox, dBox, aBox)

        Return True
    End Function

    '###############################################################################################################
    '# SINGLE CLICK FUNCTIONS ######################################################################################
    '###############################################################################################################
    '
    '######################
    '# FILE MENU (NEW WO) #
    '######################
    '
    Private Sub NewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewToolStripMenuItem.Click
        headCheck = False
        Functions.NewWorkOrder()
    End Sub

    '#######################
    '# FILE MENU (OPEN WO) #
    '#######################
    '
    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        WOPath = AppPath & "\WorkOrders\" & WONum.Text & "\"
        Dim FLE As String = WOPath & "CUSTOM_CUTLIST.xml"
        Dim EXL As String = "C:\Program Files (x86)\Microsoft Office\Office12\EXCEL.EXE" ' PATH OF/ EXCEL.EXE IN YOUR MICROSOFT OFFICE
        Dim EXL16 As String = "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"
        Try
            Shell(Chr(34) & EXL & Chr(34) & " " & Chr(34) & FLE & Chr(34), vbNormalFocus) ' OPEN XML WITH EXCEL
        Catch error4 As Exception
            Shell(Chr(34) & EXL16 & Chr(34) & " " & Chr(34) & FLE & Chr(34), vbNormalFocus) ' OPEN XML WITH EXCEL
        End Try
    End Sub

    '#######################
    '# FILE MENU (SAVE WO) #
    '#######################
    '
    Private Sub SaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem.Click
        WOPath = AppPath & "\WorkOrders\" & WONum.Text & "\"
        Try
            Dim DTB = New Data.DataTable, RWS As Integer, CLS As Integer
            Dim DRW As DataRow
            Dim DST As New DataSet
            Dim FLE As String = WOPath & "CUSTOM_CUTLIST.xml" ' PATH AND FILE NAME WHERE THE XML WIL BE CREATED (EXEMPLE: C:\REPS\XML.xml)
            Dim EXL As String = "C:\Program Files (x86)\Microsoft Office\Office12\EXCEL.EXE" ' PATH OF/ EXCEL.EXE IN YOUR MICROSOFT OFFICE
            Dim EXL16 As String = "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"
            System.IO.File.Delete(FLE)
            For CLS = 0 To DataGridView1.ColumnCount - 1 ' COLUMNS OF DTB
                DTB.Columns.Add(DataGridView1.Columns(CLS).Name.ToString)
            Next
            For RWS = 0 To DataGridView1.Rows.Count - 1 ' FILL DTB WITH DATAGRIDVIEW
                DRW = DTB.NewRow
                For CLS = 0 To DataGridView1.ColumnCount - 1
                    Try
                        DRW(DTB.Columns(CLS).ColumnName.ToString) = DataGridView1.Rows(RWS).Cells(CLS).Value.ToString
                    Catch error3 As Exception
                    End Try
                Next
                DTB.Rows.Add(DRW)
            Next
            DTB.AcceptChanges()
            DST.Tables.Add(DTB)
            DTB.WriteXml(FLE)
        Catch error5 As Exception
        End Try
    End Sub

    '#########################
    '# FILE MENU (CLOSE APP) #
    '#########################
    '
    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    '###########################
    '# EDIT MENU (UNDO ACTION) #
    '###########################
    '
    Private Sub UndoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UndoToolStripMenuItem.Click

    End Sub

    '###########################
    '# EDIT MENU (REDO ACTION) #
    '###########################
    '
    Private Sub RedoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RedoToolStripMenuItem.Click

    End Sub

    '##########################
    '# EDIT MENU (CUT ACTION) #
    '##########################
    '
    Private Sub CutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CutToolStripMenuItem.Click
        CopyToClipboard()
        For counter As Integer = 0 To DataGridView1.SelectedCells.Count - 1
            DataGridView1.SelectedCells(counter).Value = String.Empty
        Next
    End Sub

    '###########################
    '# EDIT MENU (COPY ACTION) #
    '###########################
    '
    Private Sub CopyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyToolStripMenuItem.Click
        CopyToClipboard()
    End Sub

    '############################
    '# EDIT MENU (PASTE ACTION) #
    '############################
    '
    Private Sub PasteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PasteToolStripMenuItem.Click
        PasteClipboardValue()
    End Sub

    '#############################
    '# EDIT MENU (DELETE ACTION) #
    '#############################
    '
    Private Sub DeleteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteToolStripMenuItem.Click
        Try
            For Each row As DataGridViewRow In DataGridView1.SelectedRows
                DataGridView1.Rows.Remove(row)
            Next
        Catch ex As Exception
        End Try
    End Sub

    '#################################
    '# EDIT MENU (SELECT ALL ACTION) #
    '#################################
    '
    Private Sub SelectAllToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SelectAllToolStripMenuItem.Click
        DataGridView1.SelectAll()
    End Sub

    '################################
    '# EDIT MENU (CLEAR ALL ACTION) #
    '################################
    '
    Private Sub ClearAllToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearAllToolStripMenuItem.Click
        DataGridView1.Rows.Clear()
    End Sub

    '############################
    '# DEPARTMENT MENU (CUSTOM) #
    '############################
    '
    Private Sub CustomToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CustomDept.Click
        CustomDept.Checked = True
        ProductionDept.Checked = False
    End Sub

    '################################
    '# DEPARTMENT MENU (PRODUCTION) #
    '################################
    '
    Private Sub ProductionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProductionDept.Click
        ProductionDept.Checked = True
        CustomDept.Checked = False
    End Sub

    '###########################
    '# CLEAR PROPERTIES BUTTON #
    '###########################
    '
    Private Sub btnClearRow_Click(sender As Object, e As EventArgs) Handles btnClearRow.Click
        roomBox.Text = ""
        doorStyleBox.Text = ""
        doorFinishBox.Text = ""
        speciesBox.Text = ""
        groupBox.Text = ""
        materialBox.Text = ""
        hardwareBox.Text = ""
    End Sub

    Private Sub LineCode_TextChanged(sender As Object, e As EventArgs) Handles LineCode.TextChanged

        Select Case LineCode.Text
            Case "Standard"
                materialBox.Items.Clear()
                materialBox.Items.AddRange(IO.File.ReadAllLines(AppPath & "\BoxMaterial1.txt"))
                hardwareBox.Items.Clear()
                hardwareBox.Items.AddRange(IO.File.ReadAllLines(AppPath & "\StandardHardware.txt"))
                materialBox.Text = "White Melamine Box"
                hardwareBox.Text = "Multitech"
            Case "Premier"
                materialBox.Items.Clear()
                materialBox.Items.AddRange(IO.File.ReadAllLines(AppPath & "\BoxMaterial2.txt"))
                hardwareBox.Items.Clear()
                hardwareBox.Items.AddRange(IO.File.ReadAllLines(AppPath & "\PremierHardware.txt"))
                materialBox.Text = ""
                hardwareBox.Text = ""
            Case "Classic"
                materialBox.Items.Clear()
                materialBox.Items.AddRange(IO.File.ReadAllLines(AppPath & "\BoxMaterial2.txt"))
                hardwareBox.Items.Clear()
                hardwareBox.Items.AddRange(IO.File.ReadAllLines(AppPath & "\ClassicHardware.txt"))
                materialBox.Text = ""
                hardwareBox.Text = ""
            Case "Euro Premier"
                materialBox.Items.Clear()
                materialBox.Items.AddRange(IO.File.ReadAllLines(AppPath & "\BoxMaterial3.txt"))
                hardwareBox.Items.Clear()
                hardwareBox.Items.AddRange(IO.File.ReadAllLines(AppPath & "\PremierHardware.txt"))
                materialBox.Text = ""
                hardwareBox.Text = ""
            Case "Euro Classic"
                materialBox.Items.Clear()
                materialBox.Items.AddRange(IO.File.ReadAllLines(AppPath & "\BoxMaterial2.txt"))
                hardwareBox.Items.Clear()
                hardwareBox.Items.AddRange(IO.File.ReadAllLines(AppPath & "\ClassicHardware.txt"))
                materialBox.Text = ""
                hardwareBox.Text = ""
            Case "Enter Line Type"
                materialBox.Items.Clear()
                hardwareBox.Items.Clear()
                materialBox.Text = ""
                hardwareBox.Text = ""
            Case Else
                Return
        End Select

    End Sub

    Private Sub doorFinishBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles doorFinishBox.SelectedIndexChanged
        Select Case doorFinishBox.Text
            Case "Almond Matte", "Antique White", "Black Semigloss", "Burgundy Maple", "Chocolate Maple", "Cocoa Maple", "Cream Matte", "Dove Matte", "Dove White", "Escarpment", "Fog Grey", "Grey Dark", "Grey Light", "Mushroom", "Pearl Semigloss", "White Matte", "2111-60", "2124-30", "2126-60", "2134-40", "2134-60", "AF-95", "CC-30", "CC-40", "CC-490", "CC-518", "CSP-75", "CSP-90", "CSP-100", "CSP-110", "HC-168", "HC-169", "OC-17", "OC-20", "OC-25", "OC-26", "OC-30", "OC-46", "OC-65", "OC-117"
                speciesBox.Text = "MDF"

            Case "Antique Brown Maple", "Aspen Maple", "Butternut Maple", "Cafe Maple", "Caramel Maple", "Charcoal Maple", "Ebony Maple", "Espresso Maple", "Ginger Maple", "Graphite Maple", "Grey Stone Maple", "Harvest Maple", "Khaki Maple", "Mango Maple", "Medium Brown Maple", "Mocha Maple", "Natural Maple", "Nutmeg Maple", "Olive Maple", "Peach Maple", "Pecan Maple", "Rasberry Maple", "Slate Maple", "Toffee Maple", "Tulip Maple", "Walnut Maple", "Wild Walnut Maple", "White Wash Maple"
                speciesBox.Text = "Maple"

            Case "Antique Brown Oak", "Aspen Oak", "Burgundy Oak", "Butternut Oak", "Caramel Oak", "Charcoal Oak", "Chocolate Oak", "Cocoa Oak", "Ebony Oak", "Espresso Oak", "Fruitwood Oak", "Ginger Oak", "Graphite Oak", "Grey Stone Oak", "Harvest Oak", "Khaki Oak", "Medium Brown Oak", "Mocha Oak", "Natural Oak", "Nutmeg Oak", "Olive Oak", "Peach Oak", "Pecan Oak", "Rasberry Oak", "Slate Oak", "Toffee Oak", "Tulip Oak", "White Wash Oak"
                speciesBox.Text = "Oak"

            Case "Antique White PVC", "Bleached Maple PVC", "Chocolate Maple PVC", "Honey Apple PVC", "Italian Walnut PVC", "Java Glow PVC", "Majestic Walnut PVC", "Mystic PVC", "Natural Maple PVC", "Pink Maple PVC", "Portland Cherry PVC", "Red Apple PVC", "Silken Maple PVC", "Silver PVC", "Vanilla Stix PVC", "White Ash PVC", "White Crystal PVC"
                speciesBox.Text = "PVC"

            Case "Brown Cherry Cherry", "Chestnut Cherry", "Cinnamon Cherry", "Cognac Cherry", "Natural Cherry", "Red Cherry Cherry"
                speciesBox.Text = "Cherry"

            Case "Caramel Pine"
                speciesBox.Text = "Pine"

            Case "Charcoal Melamine"
                speciesBox.Text = "Melamine"

            Case "Noce Walnut"
                speciesBox.Text = "Walnut"

            Case Else
                Return
        End Select
    End Sub

    '###############################################
    '# CABINET CARCUS CONSTRUCTION OVERRIDE FIELDS #
    '###############################################
    '
    Private Sub CabCode_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CabCode.SelectedIndexChanged
        SetField()
    End Sub

    '###############################################################################################################
    '# SELECTED INDEX CHANGED FUNCTIONS ############################################################################
    '###############################################################################################################
    '

    '###############################################################################################################
    '# GOT FOCUS FUNCTIONS #########################################################################################
    '###############################################################################################################
    '

    Private Sub LineCode_GotFocus(sender As Object, e As EventArgs) Handles LineCode.GotFocus
        If (LineCode.Text = "Enter Line Type") Then
            LineCode.ForeColor = SystemColors.Window
            LineCode.Text = ""
        End If
    End Sub

    Private Sub WONum_GotFocus(sender As Object, e As EventArgs) Handles WONum.GotFocus
        If (WONum.Text = "0000") Then
            WONum.ForeColor = SystemColors.Window
            WONum.Clear()
        End If
    End Sub

    Private Sub CabCode_GotFocus(sender As Object, e As EventArgs) Handles CabCode.GotFocus
        If (CabCode.Text = "None") Then
            CabCode.ForeColor = SystemColors.Window
            CabCode.Text = ""
        End If
    End Sub

    Private Sub CabCode3_GotFocus(sender As Object, e As EventArgs) Handles CabCode3.GotFocus
        If (CabCode3.Text = "None") Then
            CabCode3.ForeColor = SystemColors.Window
            CabCode3.Text = ""
        End If
    End Sub

    Private Sub widthBox_GotFocus(sender As Object, e As EventArgs) Handles widthBox.GotFocus
        If (widthBox.Text = "0") Then
            widthBox.ForeColor = SystemColors.Window
            widthBox.Clear()
        End If
    End Sub

    Private Sub heightBox_GotFocus(sender As Object, e As EventArgs) Handles heightBox.GotFocus
        If (heightBox.Text = "0") Then
            heightBox.ForeColor = SystemColors.Window
            heightBox.Clear()
        End If
    End Sub

    Private Sub heightBox2_GotFocus(sender As Object, e As EventArgs) Handles heightBox2.GotFocus
        If (heightBox2.Text = "0") Then
            heightBox2.ForeColor = SystemColors.Window
            heightBox2.Clear()
        End If
    End Sub

    Private Sub heightBox3_GotFocus(sender As Object, e As EventArgs) Handles heightBox3.GotFocus
        If (heightBox3.Text = "0") Then
            heightBox3.ForeColor = SystemColors.Window
            heightBox3.Clear()
        End If
    End Sub

    Private Sub depthBox_GotFocus(sender As Object, e As EventArgs) Handles depthBox.GotFocus
        If (depthBox.Text = "0") Then
            depthBox.ForeColor = SystemColors.Window
            depthBox.Clear()
        End If
    End Sub

    Private Sub depthBox2_GotFocus(sender As Object, e As EventArgs) Handles depthBox2.GotFocus
        If (depthBox2.Text = "0") Then
            depthBox2.ForeColor = SystemColors.Window
            depthBox2.Clear()
        End If
    End Sub

    Private Sub depthBox3_GotFocus(sender As Object, e As EventArgs) Handles depthBox3.GotFocus
        If (depthBox3.Text = "0") Then
            depthBox3.ForeColor = SystemColors.Window
            depthBox3.Clear()
        End If
    End Sub

    Private Sub amountBox_GotFocus(sender As Object, e As EventArgs) Handles amountBox.GotFocus
        If (depthBox2.Text = "0") Then
            depthBox2.ForeColor = SystemColors.Window
            depthBox2.Clear()
        End If
    End Sub

    '###############################################################################################################
    '# LOST FOCUS FUNCTIONS                                                                                        #
    '###############################################################################################################
    '
    Private Sub LineCode_LostFocus(sender As Object, e As EventArgs) Handles LineCode.LostFocus
        If (LineCode.Text = "") Then
            LineCode.ForeColor = SystemColors.InactiveCaption
            LineCode.Text = "Enter Line Type"
        End If
    End Sub

    Private Sub WONum_LostFocus(sender As Object, e As EventArgs) Handles WONum.LostFocus
        If (WONum.Text = "") Then
            WONum.ForeColor = SystemColors.InactiveCaption
            WONum.Text = "0000"
        End If
    End Sub

    Private Sub CabCode_LostFocus(sender As Object, e As EventArgs) Handles CabCode.LostFocus
        If (CabCode.Text = "") Then
            CabCode.ForeColor = SystemColors.InactiveCaption
            CabCode.Text = "None"
        End If
    End Sub

    Private Sub CabCode3_LostFocus(sender As Object, e As EventArgs) Handles CabCode3.LostFocus
        If (CabCode3.Text = "") Then
            CabCode3.ForeColor = SystemColors.InactiveCaption
            CabCode3.Text = "None"
        End If
    End Sub

    Private Sub widthBox_LostFocus(sender As Object, e As EventArgs) Handles widthBox.LostFocus
        If (widthBox.Text = "") Then
            widthBox.ForeColor = SystemColors.InactiveCaption
            widthBox.Text = "0"
        End If
    End Sub

    Private Sub heightBox_LostFocus(sender As Object, e As EventArgs) Handles heightBox.LostFocus
        If (heightBox.Text = "") Then
            heightBox.ForeColor = SystemColors.InactiveCaption
            heightBox.Text = "0"
        End If
    End Sub

    Private Sub heightBox2_LostFocus(sender As Object, e As EventArgs) Handles heightBox2.LostFocus
        If (heightBox2.Text = "") Then
            heightBox2.ForeColor = SystemColors.InactiveCaption
            heightBox2.Text = "0"
        End If
    End Sub

    Private Sub heightBox3_LostFocus(sender As Object, e As EventArgs) Handles heightBox3.LostFocus
        If (heightBox3.Text = "") Then
            heightBox3.ForeColor = SystemColors.InactiveCaption
            heightBox3.Text = "0"
        End If
    End Sub

    Private Sub depthBox_LostFocus(sender As Object, e As EventArgs) Handles depthBox.LostFocus
        If (depthBox.Text = "") Then
            depthBox.ForeColor = SystemColors.InactiveCaption
            depthBox.Text = "0"
        End If
    End Sub

    Private Sub depthBox2_LostFocus(sender As Object, e As EventArgs) Handles depthBox2.LostFocus
        If (depthBox2.Text = "") Then
            depthBox2.ForeColor = SystemColors.InactiveCaption
            depthBox2.Text = "0"
        End If
    End Sub

    Private Sub depthBox3_LostFocus(sender As Object, e As EventArgs) Handles depthBox3.LostFocus
        If (depthBox3.Text = "") Then
            depthBox3.ForeColor = SystemColors.InactiveCaption
            depthBox3.Text = "0"
        End If
    End Sub

    Private Sub amountBox_LostFocus(sender As Object, e As EventArgs) Handles amountBox.LostFocus
        If (depthBox2.Text = "") Then
            depthBox2.ForeColor = SystemColors.InactiveCaption
            depthBox2.Text = "0"
        End If
    End Sub

    '###############################################################################################################
    '# KEY PRESSED CHECK FUNCTION                                                                                  #
    '###############################################################################################################
    '
    Private Sub CalcCheck_KeyPress(sender As Object, e As KeyPressEventArgs) Handles widthBox.KeyPress, heightBox.KeyPress, heightBox2.KeyPress, heightBox3.KeyPress, depthBox.KeyPress, depthBox2.KeyPress, depthBox3.KeyPress
        e.Handled = Not (Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8 Or ((e.KeyChar = ".") And (sender.Text.IndexOf(".") = -1)))
    End Sub

    Private Sub NumberCheck_KeyPress(sender As Object, e As KeyPressEventArgs) Handles WONum.KeyPress
        e.Handled = Not (Char.IsDigit(e.KeyChar) Or Asc(e.KeyChar) = 8)
    End Sub

    '###############################################################################################################
    '# DEVELOPMENT FUNCTIONS                                                                                       #
    '###############################################################################################################
    '
    Private Sub btnDebug_Click(sender As Object, e As EventArgs) Handles btnDebug.Click
        LineCode.Text = "Premier"
        CabCode.Text = "BWPO"
        widthBox.Text = 45
        heightBox.Text = 88.3
        depthBox.Text = 59
        amountBox.Text = 1
        roomBox.Text = "Kitchen"
        doorStyleBox.Text = "Antica"
        doorFinishBox.Text = "Antique Brown Maple"
        groupBox.Text = "Group 1"
        materialBox.Text = "White Melamine Box"
        hardwareBox.Text = "Multitech"
    End Sub

    '################################
    '# SEND PARTS TO CUTLIST BUTTON #
    '################################
    '
    Private Sub btnSend_Click(sender As Object, e As EventArgs) Handles btnSend.Click
        If (CustomDept.Checked = True) Then
            SendData("Custom\")
        Else
            SendData("Production\")
        End If
    End Sub

    Private Sub CopyToClipboard()
        Dim dataObj As DataObject = DataGridView1.GetClipboardContent
        If Not IsNothing(dataObj) Then
            Clipboard.SetDataObject(dataObj)
        End If
    End Sub

    Private Sub PasteClipboardValue()
        If DataGridView1.SelectedCells.Count = 0 Then
            MessageBox.Show("No Cell selected", "Paste")
            Exit Sub
        End If

        Dim StartingCell As DataGridViewCell = GetStartingCell(DataGridView1)
        Dim rowCount = DataGridView1.SelectedCells.OfType(Of DataGridViewCell)().Select(Function(x) x.RowIndex).Distinct().Count()
        Dim cbvalue As Dictionary(Of Integer, Dictionary(Of Integer, String)) = ClipboardValues(Clipboard.GetText)
        Dim repeat As Integer = 0
        If rowCount > cbvalue.Keys.Count Then
            If rowCount Mod cbvalue.Keys.Count <> 0 Then
                MessageBox.Show("Selected destination doesn't match")
                Exit Sub
            Else
                repeat = CInt(rowCount / cbvalue.Keys.Count)
            End If
        End If

        Dim irowindex = StartingCell.RowIndex
        For x As Integer = 1 To repeat
            For Each rowkey As Integer In cbvalue.Keys
                Dim icolindex As Integer = StartingCell.ColumnIndex
                For Each cellkey As Integer In cbvalue(rowkey).Keys
                    If icolindex <= DataGridView1.Columns.Count - 1 And irowindex <= DataGridView1.Rows.Count - 1 Then
                        Dim cell As DataGridViewCell = DataGridView1(icolindex, irowindex)
                        cell.Value = cbvalue(rowkey)(cellkey)
                    End If
                    icolindex += 1
                Next
                irowindex += 1
            Next
        Next

    End Sub

    Private Function GetStartingCell(dgView As DataGridView) As DataGridViewCell
        If dgView.SelectedCells.Count = 0 Then Return Nothing

        Dim rowIndex As Integer = dgView.Rows.Count - 1
        Dim ColIndex As Integer = dgView.Columns.Count - 1


        For Each dgvcell As DataGridViewCell In dgView.SelectedCells

            If dgvcell.RowIndex < rowIndex Then rowIndex = dgvcell.RowIndex
            If dgvcell.ColumnIndex < ColIndex Then ColIndex = dgvcell.ColumnIndex
        Next

        Return dgView(ColIndex, rowIndex)
    End Function

    Private Function ClipboardValues(clipboardvalue As String) As Dictionary(Of Integer, Dictionary(Of Integer, String))
        Dim lines() As String = clipboardvalue.Split(CChar(Environment.NewLine))
        Dim copyValues As Dictionary(Of Integer, Dictionary(Of Integer, String)) = New Dictionary(Of Integer, Dictionary(Of Integer, String))
        For i As Integer = 0 To lines.Length - 1
            copyValues.Item(i) = New Dictionary(Of Integer, String)
            Dim linecontent() As String = lines(i).Split(ChrW(Keys.Tab))
            If linecontent.Length = 0 Then
                copyValues(i)(0) = String.Empty
            Else
                For j As Integer = 0 To linecontent.Length - 1
                    copyValues(i)(j) = linecontent(j)
                Next
            End If
        Next
        Return copyValues
    End Function

    '###############################################################################################################
    '#  ############################################################################################################
    '###############################################################################################################
    '
    '######################
    '#  #
    '######################
    '
End Class

Public Class Lines
    Public Shared Function LineCount()
        Dim BaseLines As Integer = 326
        Dim CabinetLines As Integer = 291
        Dim CanopyLines As Integer = 242
        Dim Form1Lines As Integer = 1154
        Dim FunctionsLines As Integer = 117
        Dim PanelLines As Integer = 95
        Dim TallLines As Integer = 54
        Dim UpperLines As Integer = 633
        Dim VanityLines As Integer = 64
        Form1.lblLineCount.Text = BaseLines + CabinetLines + CanopyLines + Form1Lines + FunctionsLines + PanelLines + TallLines + UpperLines + VanityLines
        LineCount = BaseLines + CabinetLines + CanopyLines + Form1Lines + FunctionsLines + PanelLines + TallLines + UpperLines + VanityLines
    End Function
End Class