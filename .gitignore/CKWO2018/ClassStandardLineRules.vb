'######################
'# IMPORT DIRECTORIES #
'######################
'
'####################################
'# PUBLIC STANDARD LINE RULES CLASS #
'####################################
'
Public Class ClassStandardLineRules
    '##########################
    '# GLOBAL CLASS VARIABLES #
    '##########################
    '

    '##################################
    '# STANDARD LINE COMMAND FUNCTION #
    '##################################
    ' To Use Type: ClassStandardLineRules.StandardLineCmdU((LINE)GROUPBOX(NUMBER).TEXT, (LINE)BOXSTYLEBOX(NUMBER).TEXT, (LINE)SPECIESBOX(NUMBER).TEXT, (LINE)DOORFINISHBOX(NUMBER).TEXT)
    Public Shared Sub StandardLineCmdU(ByVal a As String, ByVal b As String, ByVal c As String, ByVal d As String, ByVal e As String, ByVal DGV As DataGridView)

        '#######################
        '# SUB CLASS VARIABLES #
        '#######################
        '
        Dim BackGrove As String

        '#############################################
        '# ROOM ARGUMENTS FOR STANDARD LINE CABINETS #
        '#############################################
        '
        Dim CKWOPLANNERSGroupBox = a
        Dim CKWOPLANNERSBoxStyleBox = b
        Dim CKWOPLANNERSSpeciesBox = c
        Dim CKWOPLANNERSDoorFinishBox = d
        Dim CKWOPLANNERSDoorStyleBox = e
        Dim DCHeight As Integer
        Dim Hardware As String = ""
        If (CKWOPLANNER.CutListNum = 1) Then
            Hardware = CKWOPLANNER.HardwareType
        End If
        If (CKWOPLANNER.CutListNum = 2) Then
            Hardware = CKWOPLANNER.HardwareType2
        End If
        If (CKWOPLANNER.CutListNum = 3) Then
            Hardware = CKWOPLANNER.HardwareType3
        End If
        If (CKWOPLANNER.CutListNum = 4) Then
            Hardware = CKWOPLANNER.HardwareType4
        End If
        If (CKWOPLANNER.CutListNum = 5) Then
            Hardware = CKWOPLANNER.HardwareType5
        End If
        If (CKWOPLANNER.CutListNum = 6) Then
            Hardware = CKWOPLANNER.HardwareType6
        End If
        If (CKWOPLANNER.CutListNum = 7) Then
            Hardware = CKWOPLANNER.HardwareType7
        End If
        If (CKWOPLANNER.CutListNum = 8) Then
            Hardware = CKWOPLANNER.HardwareType8
        End If
        If (CKWOPLANNER.CutListNum = 9) Then
            Hardware = CKWOPLANNER.HardwareType9
        End If

        '##########################
        '# CABINET CODE VARIABLES #
        '##########################
        '
        Dim CabCode As String = ""
        Dim HB1 As Double = CutlistForm.HeightBox1.Text
        Dim HB2 As Double = CutlistForm.HeightBox2.Text
        Dim HB3 As Double = CutlistForm.HeightBox3.Text
        Dim TotalCabSize As Double = HB1 + HB2 + HB3
        Dim TotalCabSizeb As Double = HB1 + HB2
        Dim CabSize1 As String = CutlistForm.WidthBox1.Text & "-" & CutlistForm.HeightBox1.Text & "-" & CutlistForm.DepthBox1.Text
        Dim CabSize1b As String = CutlistForm.WidthBox1.Text & "-" & TotalCabSizeb & "-" & CutlistForm.DepthBox1.Text
        Dim CabSize1c As String = CutlistForm.WidthBox1.Text & "-" & TotalCabSize & "-" & CutlistForm.DepthBox1.Text
        Dim CabSize2 As String = CutlistForm.WidthBox1.Text & "-" & CutlistForm.HeightBox2.Text & "-" & CutlistForm.DepthBox1.Text
        Dim CabSize2b As String = CutlistForm.WidthBox1.Text & "-" & CutlistForm.HeightBox2.Text & "-" & CutlistForm.DepthBox2.Text
        Dim CabSize3 As String = CutlistForm.WidthBox1.Text & "-" & CutlistForm.HeightBox3.Text & "-" & CutlistForm.DepthBox1.Text
        Dim CabSize3b As String = CutlistForm.WidthBox1.Text & "-" & CutlistForm.HeightBox3.Text & "-" & CutlistForm.DepthBox3.Text

        '#####################################################################################
        '# GET CABINET INFORMATION FROM USER INPUT [CAB TYPE, WIDTH, HEIGHT, DEPTH, QUANITY] #
        '#####################################################################################
        '
        Dim VarCabCode As String = CutlistForm.CabCodeBox1.Text 'GETS CABINET CODE FROM "CabCodeBox1" TEXT AND STORES AS VARIABLE "VarCabCode"
        Dim VarWidthI As Double = CutlistForm.WidthBox1.Text  'GETS CABINET WIDTH FROM "WidthBox1" TEXT AND STORES AS VARIABLE "VarWidthI"
        Dim VarHeightI As Double = CutlistForm.HeightBox1.Text  'GETS CABINET HEIGHT FROM "HeightBox1" TEXT AND STORES AS VARIABLE "VarHeightI"
        Dim VarDepthI As Double = CutlistForm.DepthBox1.Text  'GETS CABINET DEPTH FROM "DepthBox1" TEXT AND STORES AS VARIABLE "VarDepthI"
        Dim VarAmountI As Integer = CutlistForm.AmountBox1.Text 'GETS CABINET QUANTITY FROM "AmountBox1" TEXT AND STORES AS VARIABLE "VarAmountI"
        Dim VarHeightI2 As Double = CutlistForm.HeightBox2.Text  'GETS CABINET HEIGHT FROM "HeightBox1" TEXT AND STORES AS VARIABLE "VarHeightI"
        Dim VarDepthI2 As Double = CutlistForm.DepthBox2.Text  'GETS CABINET DEPTH FROM "DepthBox1" TEXT AND STORES AS VARIABLE "VarDepthI"

        '#######################
        '# EDGE BAND VARIABLES #
        '#######################
        '
        Dim VeneerEdgeCode As String = "" 'DECLARES VENEER EDGE CODE AS A STRING VARIABLE
        Dim PVCEdgeCode As String = "" 'DECLARES PVC EDGE CODE AS A STRING VARIABLE
        Dim BMEdgeCode As String = "" 'DECLARES BOX MATERIAL EDGE CODE AS A STRING VARIABLE

        '#########################
        '# UPPER GABLE VARIABLES #
        '#########################
        '
        Dim UGableX As Double 'DECLARES UPPER GABLE X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UGableY As Double 'DECLARES UPPER GABLE Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UGableZ As Double 'DECLARES UPPER GABLE Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UGableQ As Integer 'DECLARES UPPER GABLE QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim UGableM As String = "" 'DECLARES UPPER GABLE MATERIAL AS STRING VARIABLE
        Dim UGEdgeSeq As String = "" 'DECLARES UPPER GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim UGEdgeSeq2 As String = "" 'DECLARES UPPER GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim UGEdgeCode As String = "" 'DECLARES UPPER GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim UGEdgeCode2 As String = "" 'DECLARES UPPER GABLE EDGE SEQUENCE AS A STRING VARIABLE

        '#######################
        '# UPPER TOB VARIABLES #
        '#######################
        '
        Dim UTopX As Double 'DECLARES UPPER TOP AND BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UTopY As Double 'DECLARES UPPER TOP AND BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UTopZ As Double 'DECLARES UPPER TOP AND BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UTopQ As Integer 'DECLARES UPPER TOP AND BOTTOM QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim UTopM As String = "" 'DECLARES UPPER TOP AND BOTTOM MATERIAL AS STRING VARIABLE
        Dim UTEdgeSeq As String = "" 'DECLARES UPPER TOP LIGHT EDGE SEQUENCE AS A STRING VARIABLE
        Dim UTEdgeCode As String = "" 'DECLARES UPPER TOP LIGHT EDGE SEQUENCE AS A STRING VARIABLE

        '#############################
        '# UPPER TOP HINGE VARIABLES #
        '#############################
        '
        Dim UTopHingeX As Double 'DECLARES UPPER TOP AND BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UTopHingeY As Double 'DECLARES UPPER TOP AND BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UTopHingeZ As Double 'DECLARES UPPER TOP AND BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UTopHingeQ As Integer 'DECLARES UPPER TOP AND BOTTOM QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim UTopHingeM As String = "" 'DECLARES UPPER TOP AND BOTTOM MATERIAL AS STRING VARIABLE
        Dim UTHEdgeSeq As String = "" 'DECLARES UPPER TOP LIGHT EDGE SEQUENCE AS A STRING VARIABLE
        Dim UTHEdgeCode As String = "" 'DECLARES UPPER TOP LIGHT EDGE SEQUENCE AS A STRING VARIABLE

        '#######################
        '# UPPER TOB VARIABLES #
        '#######################
        '
        Dim UFanSHX As Double 'DECLARES UPPER TOP AND BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UFanSHY As Double 'DECLARES UPPER TOP AND BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UFanSHZ As Double 'DECLARES UPPER TOP AND BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UFanSHQ As Integer 'DECLARES UPPER TOP AND BOTTOM QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim UFanSHM As String = "" 'DECLARES UPPER TOP AND BOTTOM MATERIAL AS STRING VARIABLE

        '################################
        '# UPPER TOB LIGHT VARIABLES #
        '################################
        '
        Dim UTopLightX As Double 'DECLARES UPPER TOP LIGHT X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UTopLightY As Double 'DECLARES UPPER TOP LIGHT Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UTopLightZ As Double 'DECLARES UPPER TOP LIGHT Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UTopLightQ As Integer 'DECLARES UPPER TOP LIGHT QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim UTopLightM As String = "" 'DECLARES UPPER TOP LIGHT MATERIAL AS STRING VARIABLE
        Dim UTLEdgeSeq As String = "" 'DECLARES UPPER TOP LIGHT EDGE SEQUENCE AS A STRING VARIABLE
        Dim UTLEdgeCode As String = "" 'DECLARES UPPER TOP LIGHT EDGE SEQUENCE AS A STRING VARIABLE

        '##########################
        '# UPPER BOTTOM VARIABLES #
        '##########################
        '
        Dim UBtmX As Double 'DECLARES UPPER BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UBtmY As Double 'DECLARES UPPER BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UBtmZ As Double 'DECLARES UPPER BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UBtmQ As Integer 'DECLARES UPPER BOTTOM QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim UBtmM As String = "" 'DECLARES UPPER BOTTOM MATERIAL AS STRING VARIABLE
        Dim UBEdgeSeq As String = "" 'DECLARES UPPER BOTTOM EDGE SEQUENCE AS A STRING VARIABLE
        Dim UBEdgeCode As String = "" 'DECLARES UPPER BOTTOM EDGE SEQUENCE AS A STRING VARIABLE

        '################################
        '# UPPER TOB + BOTTOM VARIABLES #
        '################################
        '
        Dim UTopBtmX As Double 'DECLARES UPPER TOP AND BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UTopBtmY As Double 'DECLARES UPPER TOP AND BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UTopBtmZ As Double 'DECLARES UPPER TOP AND BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UTopBtmQ As Integer 'DECLARES UPPER TOP AND BOTTOM QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim UTopBtmM As String = "" 'DECLARES UPPER TOP AND BOTTOM MATERIAL AS STRING VARIABLE
        Dim UTBEdgeSeq As String = "" 'DECLARES UPPER TOP AND BOTTOM EDGE SEQUENCE AS A STRING VARIABLE
        Dim UTBEdgeCode As String = "" 'DECLARES UPPER TOP AND BOTTOM EDGE SEQUENCE AS A STRING VARIABLE

        '################################
        '# UPPER TOB + BOTTOM VARIABLES #
        '################################
        '
        Dim UFDividerX As Double 'DECLARES UPPER TOP AND BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UFDividerY As Double 'DECLARES UPPER TOP AND BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UFDividerZ As Double 'DECLARES UPPER TOP AND BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UFDividerQ As Integer 'DECLARES UPPER TOP AND BOTTOM QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim UFDividerM As String = "" 'DECLARES UPPER TOP AND BOTTOM MATERIAL AS STRING VARIABLE
        Dim UTFDEdgeSeq As String = "" 'DECLARES UPPER TOP AND BOTTOM EDGE SEQUENCE AS A STRING VARIABLE
        Dim UTFDEdgeCode As String = "" 'DECLARES UPPER TOP AND BOTTOM EDGE SEQUENCE AS A STRING VARIABLE

        '################################
        '# UPPER TOB + BOTTOM VARIABLES #
        '################################
        '
        Dim UDividerX As Double 'DECLARES UPPER TOP AND BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UDividerY As Double 'DECLARES UPPER TOP AND BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UDividerZ As Double 'DECLARES UPPER TOP AND BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UDividerQ As Integer 'DECLARES UPPER TOP AND BOTTOM QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim UDividerM As String = "" 'DECLARES UPPER TOP AND BOTTOM MATERIAL AS STRING VARIABLE
        Dim UTDEdgeSeq As String = "" 'DECLARES UPPER TOP AND BOTTOM EDGE SEQUENCE AS A STRING VARIABLE
        Dim UTDEdgeCode As String = "" 'DECLARES UPPER TOP AND BOTTOM EDGE SEQUENCE AS A STRING VARIABLE

        '######################################
        '# UPPER ADJUSTABLE SHELVES VARIABLES #
        '######################################
        '
        Dim UAdjShelfX As Double 'DECLARES UPPER ADJUSTABLE SHELF X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UAdjShelfY As Double 'DECLARES UPPER ADJUSTABLE SHELF Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UAdjShelfZ As Double 'DECLARES UPPER ADJUSTABLE SHELF Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UAdjShelfQ As Integer 'DECLARES UPPER ADJUSTABLE SHELF QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim UAdjShelfM As String = "" 'DECLARES UPPER ADJUSTABLE SHELF MATERIAL AS STRING VARIABLE
        Dim UASEdgeSeq As String = "" 'DECLARES UPPER ADJUSTABLE SHELF EDGE SEQUENCE AS A STRING VARIABLE
        Dim UASEdgeCode As String = "" 'DECLARES UPPER ADJUSTABLE SHELF EDGE SEQUENCE AS A STRING VARIABLE

        '###########################
        '# UPPER BACKING VARIABLES #
        '###########################
        '
        Dim UBackX As Double 'DECLARES UPPER BACK X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UBackY As Double 'DECLARES UPPER BACK Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UBackZ As Double 'DECLARES UPPER BACK Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UBackQ As Integer 'DECLARES UPPER BACK QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim UBackM As String = "" 'DECLARES UPPER BACK MATERIAL AS STRING VARIABLE

        '###################################
        '# UPPER GABLE RULES AND EQUATIONS #
        '###################################
        '
        UGableX = (VarHeightI * 10) - 1
        UGableY = VarDepthI * 10
        UGableZ = 16
        UGableQ = 2 * VarAmountI
        UGableM = ""

        '#############################################
        '# UPPER TOP + FAN SHELF RULES AND EQUATIONS #
        '#############################################
        '
        If (CutlistForm.UpperCabCodeBox.Text = "UA") Then
            UTopX = ((VarWidthI * 10) - (UGableZ * 2))
            UTopY = VarDepthI * 10
            UTopZ = 16
            UTopQ = 1 * VarAmountI
            UTopM = ""

            UFanSHX = ((VarWidthI * 10) - (UGableZ * 2))
            UFanSHY = (VarDepthI * 10) - 17
            UFanSHZ = 19
            UFanSHQ = 2 * VarAmountI
            UFanSHM = "PLY"
        End If

        '########################################
        '# UPPER TOP HINDGE RULES AND EQUATIONS #
        '########################################
        '
        If (CutlistForm.UpperCabCodeBox.Text = "UL" Or CutlistForm.UpperCabCodeBox.Text = "UDL" Or CutlistForm.UpperCabCodeBox.Text = "UUDL" Or CutlistForm.UpperCabCodeBox.Text = "USDL") Then
            UTopHingeX = ((VarWidthI * 10) - (UGableZ * 2)) - 1
            UTopHingeY = VarDepthI * 10
            UTopHingeZ = 16
            UTopHingeQ = 1 * VarAmountI
            UTopHingeM = ""
        End If

        '#######################################
        '# UPPER TOP LIGHT RULES AND EQUATIONS #
        '#######################################
        '
        UTopLightX = ((VarWidthI * 10) - (UGableZ * 2)) - 1
        UTopLightY = VarDepthI * 10
        UTopLightZ = 16
        UTopLightQ = 1 * VarAmountI
        UTopLightM = ""

        '####################################
        '# UPPER BOTTOM RULES AND EQUATIONS #
        '####################################
        '
        UBtmX = ((VarWidthI * 10) - (UGableZ * 2)) - 1
        UBtmY = VarDepthI * 10
        UBtmZ = 16
        UBtmQ = 1 * VarAmountI
        UBtmM = ""

        '##########################################
        '# UPPER TOP + BOTTOM RULES AND EQUATIONS #
        '##########################################
        '
        UTopBtmX = ((VarWidthI * 10) - (UGableZ * 2)) - 1
        UTopBtmY = VarDepthI * 10
        UTopBtmZ = 16
        UTopBtmQ = 2 * VarAmountI
        UTopBtmM = ""

        '################################################
        '# UPPER TRAY FIXED DIVIDER RULES AND EQUATIONS #
        '################################################
        '
        UFDividerX = ((VarHeightI * 10) - (UTopBtmZ * 2))
        If (VarDepthI <= 63) Then
            UFDividerY = 450
        Else
            UFDividerY = 530
        End If
        UFDividerZ = 16
        UFDividerQ = 1 * VarAmountI
        UFDividerM = ""

        '##########################################
        '# UPPER TRAY DIVIDER RULES AND EQUATIONS #
        '##########################################
        '
        UDividerX = ((VarHeightI * 10) - (UTopBtmZ * 2)) - 2
        If (VarDepthI <= 63) Then
            UDividerY = 450
        Else
            UDividerY = 530
        End If
        UDividerZ = 16
        If (CutlistForm.DividerCountBox.Text = "") Then
            UDividerQ = 0
        Else
            UDividerQ = CutlistForm.DividerCountBox.Text * VarAmountI
        End If
        UDividerM = ""

        '##############################################
        '# UPPER ADJUSTABLE SHELF RULES AND EQUATIONS #
        '##############################################
        '
        Dim UCase1 = CutlistForm.UpperCabCodeBox.Text
        Select Case UCase1
            Case "UTDH"
                UAdjShelfX = (UTopBtmX / 2) - (UFDividerZ / 2) - 4
            Case Else
                UAdjShelfX = ((VarWidthI * 10) - (UGableZ * 2)) - 5
        End Select

        '##################################################
        '# GET SIZE OF SHELVES ACCORDING TO CABINET DEPTH #
        '##################################################
        '
        Dim UCase2 = CutlistForm.UpperCabCodeBox.Text
        Select Case UCase2
            Case "USHF", "USHS", "UUHF", "UUHS"
                UAdjShelfY = (VarDepthI * 10) - 44
            Case "USHK", "UUHK"
                UAdjShelfY = (VarDepthI * 10) - 30
            Case "USHL", "UUHL"
                UAdjShelfY = (VarDepthI * 10) - 78
            Case Else
                If VarDepthI <= 48 Then UAdjShelfY = (VarDepthI * 10) - 30
                If VarDepthI > 48 Then UAdjShelfY = 450
                If VarDepthI > 62 Then UAdjShelfY = 530
        End Select

        UAdjShelfZ = 16
        UAdjShelfM = ""

        '####################################################
        '# GET QUANITY OF SHELVES ACCORDING TO GABLE HEIGHT #
        '####################################################
        '
        If (UGableX <= 497) Then UAdjShelfQ = 0
        If (UGableX >= 498) Then UAdjShelfQ = 1 * VarAmountI
        If (UGableX >= 699) Then UAdjShelfQ = 2 * VarAmountI
        If (UGableX >= 898) Then UAdjShelfQ = 3 * VarAmountI
        If (UGableX >= 1440) Then UAdjShelfQ = 4 * VarAmountI
        CutlistForm.PubASQuantity = UAdjShelfQ

        '#####################################
        '# UPPER BACKING RULES AND EQUATIONS #
        '#####################################
        '
        If (CutlistForm.UpperCabCodeBox.Text = "UA") Then
            UBackY = (VarWidthI * 10) - 22
            UBackX = (VarHeightI * 10) - 12
            UBackZ = 3
            UBackQ = 1 * VarAmountI
            UBackM = ""
        Else
            UBackY = (VarWidthI * 10) - 23
            UBackX = (VarHeightI * 10) - 23
            UBackZ = 3
            UBackQ = 1 * VarAmountI
            UBackM = ""
        End If

        '###########################################
        '# UPPER MATCHING INTERIOR GABLE VARIABLES #
        '###########################################
        '
        Dim UMIGableX As Double 'DECLARES UPPER MATCHING INTERIOR GABLE X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMIGableY As Double 'DECLARES UPPER MATCHING INTERIOR GABLE Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMIGableZ As Double 'DECLARES UPPER MATCHING INTERIOR GABLE Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMIGableQ As Integer 'DECLARES UPPER MATCHING INTERIOR GABLE QUANITY AMOUNT AS INTEGER VARIABLE
        Dim UMIGableM As String = "" 'DECLARES UPPER MATCHING INTERIOR GABLE MATERIAL AS STRING VARIABLE
        Dim UMIGEdgeSeq As String = "" 'DECLARES UPPER MATCHING INTERIOR GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim UMIGEdgeSeq2 As String = "" 'DECLARES UPPER MATCHING INTERIOR GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim UMIGEdgeCode As String = "" 'DECLARES UPPER MATCHING INTERIOR GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim UMIGEdgeCode2 As String = "" 'DECLARES UPPER MATCHING INTERIOR GABLE EDGE SEQUENCE AS A STRING VARIABLE

        '###############################################
        '# UPPER MATCHING INTERIOR TOP LIGHT VARIABLES #
        '###############################################
        '
        Dim UMITopLightX As Double 'DECLARES UPPER MATCHING INTERIOR TOP AND BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMITopLightY As Double 'DECLARES UPPER MATCHING INTERIOR TOP AND BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMITopLightZ As Double 'DECLARES UPPER MATCHING INTERIOR TOP AND BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMITopLightQ As Integer 'DECLARES UPPER MATCHING INTERIOR TOP AND BOTTOM QUANITY AMOUNT AS INTEGER VARIABLE
        Dim UMITopLightM As String = "" 'DECLARES UPPER MATCHING INTERIOR TOP AND BOTTOM MATERIAL AS STRING VARIABLE
        Dim UMITLEdgeSeq As String = "" 'DECLARES UPPER MATCHING INTERIOR TOP LIGHT EDGE SEQUENCE AS A STRING VARIABLE
        Dim UMITLEdgeCode As String = "" 'DECLARES UPPER MATCHING INTERIOR TOP LIGHT EDGE SEQUENCE AS A STRING VARIABLE

        '############################################
        '# UPPER MATCHING INTERIOR BOTTOM VARIABLES #
        '############################################
        '
        Dim UMIBtmX As Double 'DECLARES UPPER MATCHING INTERIOR TOP AND BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMIBtmY As Double 'DECLARES UPPER MATCHING INTERIOR TOP AND BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMIBtmZ As Double 'DECLARES UPPER MATCHING INTERIOR TOP AND BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMIBtmQ As Integer 'DECLARES UPPER MATCHING INTERIOR TOP AND BOTTOM QUANITY AMOUNT AS INTEGER VARIABLE
        Dim UMIBtmM As String = "" 'DECLARES UPPER MATCHING INTERIOR TOP AND BOTTOM MATERIAL AS STRING VARIABLE
        Dim UMIBEdgeSeq As String = "" 'DECLARES UPPER MATCHING INTERIOR BOTTOM EDGE SEQUENCE AS A STRING VARIABLE
        Dim UMIBEdgeCode As String = "" 'DECLARES UPPER MATCHING INTERIOR BOTTOM EDGE SEQUENCE AS A STRING VARIABLE

        '##################################################
        '# UPPER MATCHING INTERIOR TOP + BOTTOM VARIABLES #
        '##################################################
        '
        Dim UMITopBtmX As Double 'DECLARES UPPER MATCHING INTERIOR TOP AND BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMITopBtmY As Double 'DECLARES UPPER MATCHING INTERIOR TOP AND BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMITopBtmZ As Double 'DECLARES UPPER MATCHING INTERIOR TOP AND BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMITopBtmQ As Integer 'DECLARES UPPER MATCHING INTERIOR TOP AND BOTTOM QUANITY AMOUNT AS INTEGER VARIABLE
        Dim UMITopBtmM As String = "" 'DECLARES UPPER MATCHING INTERIOR TOP AND BOTTOM MATERIAL AS STRING VARIABLE
        Dim UMITBEdgeSeq As String = "" 'DECLARES UPPER MATCHING INTERIOR TOP AND BOTTOM EDGE SEQUENCE AS A STRING VARIABLE
        Dim UMITBEdgeCode As String = "" 'DECLARES UPPER MATCHING INTERIOR TOP AND BOTTOM EDGE SEQUENCE AS A STRING VARIABLE

        '########################################################
        '# UPPER MATCHING INTERIOR ADJUSTABLE SHELVES VARIABLES #
        '########################################################
        '
        Dim UMIAdjShelfX As Double 'DECLARES UPPER MATCHING INTERIOR ADJUSTABLE SHELF X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMIAdjShelfY As Double 'DECLARES UPPER MATCHING INTERIOR ADJUSTABLE SHELF Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMIAdjShelfZ As Double 'DECLARES UPPER MATCHING INTERIOR ADJUSTABLE SHELF Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMIAdjShelfQ As Integer 'DECLARES UPPER MATCHING INTERIOR ADJUSTABLE SHELF QUANITY AMOUNT AS INTEGER VARIABLE
        Dim UMIAdjShelfM As String = "" 'DECLARES UPPER MATCHING INTERIOR ADJUSTABLE SHELF MATERIAL AS STRING VARIABLE
        Dim UMIASEdgeSeq As String = "" 'DECLARES UPPER MATCHING INTERIOR ADJUSTABLE SHELF EDGE SEQUENCE AS A STRING VARIABLE
        Dim UMIASEdgeCode As String = "" 'DECLARES UPPER MATCHING INTERIOR ADJUSTABLE SHELF EDGE SEQUENCE AS A STRING VARIABLE

        '#############################################
        '# UPPER MATCHING INTERIOR BACKING VARIABLES #
        '#############################################
        '
        Dim UMIBackX As Double 'DECLARES UPPER MATCHING INTERIOR BACK X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMIBackY As Double 'DECLARES UPPER MATCHING INTERIOR BACK Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMIBackZ As Double 'DECLARES UPPER MATCHING INTERIOR BACK Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMIBackQ As Integer 'DECLARES UPPER MATCHING INTERIOR BACK QUANITY AMOUNT AS INTEGER VARIABLE
        Dim UMIBackM As String = "" 'DECLARES UPPER MATCHING INTERIOR BACK MATERIAL AS STRING VARIABLE

        '###############################################
        '# UPPER MATCHING INTERIOR FAN SHELF VARIABLES #
        '###############################################
        '
        Dim UMIFanSHX As Double 'DECLARES UPPER MATCHING INTERIOR FAN SHELF X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMIFanSHY As Double 'DECLARES UPPER MATCHING INTERIOR FAN SHELF Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMIFanSHZ As Double 'DECLARES UPPER MATCHING INTERIOR FAN SHELF Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMIFanSHQ As Integer 'DECLARES UPPER MATCHING INTERIOR FAN SHELF Q AMOUNT AS INTEGER VARIABLE
        Dim UMIFanSHM As String = "" 'DECLARES UPPER MATCHING INTERIOR FAN SHELF MATERIAL AS STRING VARIABLE

        '####################################################
        '# UPPER MATCING INTERIOR GABLE RULES AND EQUATIONS #
        '####################################################
        '
        UMIGableX = (VarHeightI * 10) - 1
        UMIGableY = VarDepthI * 10
        UMIGableZ = 16
        UMIGableQ = 2 * VarAmountI
        UMIGableM = ""

        '########################################################
        '# UPPER MATCING INTERIOR TOP LIGHT RULES AND EQUATIONS #
        '########################################################
        '
        UMITopLightX = ((VarWidthI * 10) - (UMIGableZ * 2))
        UMITopLightY = VarDepthI * 10
        UMITopLightZ = 16
        UMITopLightQ = 1 * VarAmountI
        UMITopLightM = ""

        '#####################################################
        '# UPPER MATCING INTERIOR BOTTOM RULES AND EQUATIONS #
        '#####################################################
        '
        UMIBtmX = ((VarWidthI * 10) - (UMIGableZ * 2))
        UMIBtmY = VarDepthI * 10
        UMIBtmZ = 16
        UMIBtmQ = 1 * VarAmountI
        UMIBtmM = ""

        '###########################################################
        '# UPPER MATCING INTERIOR TOP + BOTTOM RULES AND EQUATIONS #
        '###########################################################
        '
        UMITopBtmX = ((VarWidthI * 10) - (UMIGableZ * 2))
        UMITopBtmY = VarDepthI * 10
        UMITopBtmZ = 16
        UMITopBtmQ = 2 * VarAmountI
        UMITopBtmM = ""

        '###############################################################
        '# UPPER MATCING INTERIOR ADJUSTABLE SHELF RULES AND EQUATIONS #
        '###############################################################
        '
        UMIAdjShelfX = ((VarWidthI * 10) - (UMIGableZ * 2)) - 3
        UMIAdjShelfY = (VarDepthI * 10) - 30
        UMIAdjShelfZ = 16
        UMIAdjShelfM = ""

        '####################################################
        '# GET QUANITY OF SHELVES ACCORDING TO GABLE HEIGHT #
        '####################################################
        '
        If (UMIGableX <= 497) Then
            UMIAdjShelfQ = 0
        End If
        If (UMIGableX >= 498) Then
            UMIAdjShelfQ = 1 * VarAmountI
        End If
        If (UMIGableX >= 699) Then
            UMIAdjShelfQ = 2 * VarAmountI
        End If
        If (UMIGableX >= 898) Then
            UMIAdjShelfQ = 3 * VarAmountI
        End If
        If (UMIGableX >= 1440) Then
            UMIAdjShelfQ = 4 * VarAmountI
        End If
        CutlistForm.PubASQuantity = UMIAdjShelfQ

        '######################################################
        '# UPPER MATCING INTERIOR BACKING RULES AND EQUATIONS #
        '######################################################
        '
        UMIBackY = (VarWidthI * 10) - 23
        UMIBackX = (VarHeightI * 10) - 23
        UMIBackZ = 16
        UMIBackQ = 1 * VarAmountI
        UMIBackM = ""

        '########################################################
        '# UPPER MATCING INTERIOR FAN SHELF RULES AND EQUATIONS #
        '########################################################
        '
        UMIFanSHX = ((VarWidthI * 10) - (UGableZ * 2))
        UMIFanSHY = (VarDepthI * 10) - 17
        UMIFanSHZ = 19
        UMIFanSHQ = 2 * VarAmountI
        UMIFanSHM = "PLY"

        '###################################
        '# UPPER MICROWAVE GABLE VARIABLES #
        '###################################
        '
        Dim UMGableX As Double 'DECLARES UPPER GABLE X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMGableY As Double 'DECLARES UPPER GABLE Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMGableZ As Double 'DECLARES UPPER GABLE Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMGableQ As Integer 'DECLARES UPPER GABLE QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim UMGableM As String = "" 'DECLARES UPPER GABLE MATERIAL AS STRING VARIABLE
        Dim UMGEdgeSeq As String = "" 'DECLARES UPPER MICROWAVE GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim UMGEdgeSeq2 As String = "" 'DECLARES UPPER MICROWAVE GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim UMGEdgeCode As String = "" 'DECLARES UPPER MICROWAVE GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim UMGEdgeCode2 As String = "" 'DECLARES UPPER MICROWAVE GABLE EDGE SEQUENCE AS A STRING VARIABLE

        '##########################################
        '# UPPER MICROWAVE TOB + BOTTOM VARIABLES #
        '##########################################
        '
        Dim UMTopX As Double 'DECLARES UPPER TOP X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMTopY As Double 'DECLARES UPPER TOP Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMTopZ As Double 'DECLARES UPPER TOP Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMTopQ As Integer 'DECLARES UPPER TOP QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim UMTopM As String = "" 'DECLARES UPPER TOP MATERIAL AS STRING VARIABLE
        Dim UMTEdgeSeq As String = "" 'DECLARES UPPER MICROWAVE TOP EDGE SEQUENCE AS A STRING VARIABLE
        Dim UMTEdgeCode As String = "" 'DECLARES UPPER MICROWAVE TOP EDGE SEQUENCE AS A STRING VARIABLE

        '##########################################
        '# UPPER MICROWAVE TOB + BOTTOM VARIABLES #
        '##########################################
        '
        Dim UMBtmX As Double 'DECLARES UPPER BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMBtmY As Double 'DECLARES UPPER BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMBtmZ As Double 'DECLARES UPPER BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMBtmQ As Integer 'DECLARES UPPER BOTTOM QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim UMBtmM As String = "" 'DECLARES UPPER BOTTOM MATERIAL AS STRING VARIABLE
        Dim UMBEdgeSeq As String = "" 'DECLARES UPPER MICROWAVE BOTTOM EDGE SEQUENCE AS A STRING VARIABLE
        Dim UMBEdgeCode As String = "" 'DECLARES UPPER MICROWAVE BOTTOM EDGE SEQUENCE AS A STRING VARIABLE

        '################################################
        '# UPPER MICROWAVE ADJUSTABLE SHELVES VARIABLES #
        '################################################
        '
        Dim UMAdjShelfX As Double 'DECLARES UPPER ADJUSTABLE SHELF X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMAdjShelfY As Double 'DECLARES UPPER ADJUSTABLE SHELF Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMAdjShelfZ As Double 'DECLARES UPPER ADJUSTABLE SHELF Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMAdjShelfQ As Integer 'DECLARES UPPER ADJUSTABLE SHELF QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim UMAdjShelfM As String = "" 'DECLARES UPPER ADJUSTABLE SHELF MATERIAL AS STRING VARIABLE
        Dim UMASEdgeSeq As String = "" 'DECLARES UPPER MICROWAVE ADJUSTABLE SHELF EDGE SEQUENCE AS A STRING VARIABLE
        Dim UMASEdgeCode As String = "" 'DECLARES UPPER MICROWAVE ADJUSTABLE SHELF EDGE SEQUENCE AS A STRING VARIABLE

        '#####################################
        '# UPPER MICROWAVE BACKING VARIABLES #
        '#####################################
        '
        Dim UMBackX As Double 'DECLARES UPPER BACK X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMBackY As Double 'DECLARES UPPER BACK Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMBackZ As Double 'DECLARES UPPER BACK Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMBackQ As Integer 'DECLARES UPPER BACK QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim UMBackM As String = "" 'DECLARES UPPER BACK MATERIAL AS STRING VARIABLE

        '#####################################################
        '# UPPER MICROWAVE MATCHING INTERIOR GABLE VARIABLES #
        '#####################################################
        '
        Dim UMMIGableX As Double 'DECLARES UPPER GABLE X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMMIGableY As Double 'DECLARES UPPER GABLE Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMMIGableZ As Double 'DECLARES UPPER GABLE Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMMIGableQ As Integer 'DECLARES UPPER GABLE QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim UMMIGableM As String = "" 'DECLARES UPPER GABLE MATERIAL AS STRING VARIABLE
        Dim UMMIGEdgeSeq As String = "" 'DECLARES UPPER MICROWAVE MATCHING INTERIOR GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim UMMIGEdgeSeq2 As String = "" 'DECLARES UPPER MICROWAVE MATCHING INTERIOR GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim UMMIGEdgeCode As String = "" 'DECLARES UPPER MICROWAVE MATCHING INTERIOR GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim UMMIGEdgeCode2 As String = "" 'DECLARES UPPER MICROWAVE MATCHING INTERIOR GABLE EDGE SEQUENCE AS A STRING VARIABLE

        '############################################################
        '# UPPER MICROWAVE MATCHING INTERIOR TOB + BOTTOM VARIABLES #
        '############################################################
        '
        Dim UMMITopBtmX As Double 'DECLARES UPPER TOP AND BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMMITopBtmY As Double 'DECLARES UPPER TOP AND BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMMITopBtmZ As Double 'DECLARES UPPER TOP AND BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMMITopBtmQ As Integer 'DECLARES UPPER TOP AND BOTTOM QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim UMMITopBtmM As String = "" 'DECLARES UPPER TOP AND BOTTOM MATERIAL AS STRING VARIABLE
        Dim UMMITBEdgeSeq As String = "" 'DECLARES UPPER MICROWAVE MATCHING INTERIOR TOP AND BOTTOM EDGE SEQUENCE AS A STRING VARIABLE
        Dim UMMITBEdgeCode As String = "" 'DECLARES UPPER MICROWAVE MATCHING INTERIOR TOP AND BOTTOM EDGE SEQUENCE AS A STRING VARIABLE

        '##################################################################
        '# UPPER MICROWAVE MATCHING INTERIOR ADJUSTABLE SHELVES VARIABLES #
        '##################################################################
        '
        Dim UMMIMicroSHX As Double 'DECLARES UPPER ADJUSTABLE SHELF X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMMIMicroSHY As Double 'DECLARES UPPER ADJUSTABLE SHELF Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMMIMicroSHZ As Double 'DECLARES UPPER ADJUSTABLE SHELF Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMMIMicroSHQ As Integer 'DECLARES UPPER ADJUSTABLE SHELF QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim UMMIMicroSHM As String = "" 'DECLARES UPPER ADJUSTABLE SHELF MATERIAL AS STRING VARIABLE
        Dim UMMIMSEdgeSeq As String = "" 'DECLARES UPPER MICROWAVE MATCHING INTERIOR ADJUSTABLE SHELF EDGE SEQUENCE AS A STRING VARIABLE
        Dim UMMIMSEdgeCode As String = "" 'DECLARES UPPER MICROWAVE MATCHING INTERIOR ADJUSTABLE SHELF EDGE SEQUENCE AS A STRING VARIABLE
        Dim UMMIMSEdgeSeq2 As String = "" 'DECLARES UPPER MICROWAVE MATCHING INTERIOR ADJUSTABLE SHELF EDGE SEQUENCE AS A STRING VARIABLE
        Dim UMMIMSEdgeCode2 As String = "" 'DECLARES UPPER MICROWAVE MATCHING INTERIOR ADJUSTABLE SHELF EDGE SEQUENCE AS A STRING VARIABLE

        '#######################################################
        '# UPPER MICROWAVE MATCHING INTERIOR BACKING VARIABLES #
        '#######################################################
        '
        Dim UMMIBackX As Double 'DECLARES UPPER BACK X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMMIBackY As Double 'DECLARES UPPER BACK Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMMIBackZ As Double 'DECLARES UPPER BACK Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UMMIBackQ As Integer 'DECLARES UPPER BACK QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim UMMIBackM As String = "" 'DECLARES UPPER BACK MATERIAL AS STRING VARIABLE

        '#######################################
        '# UPPER MICROWAVE RULES AND EQUATIONS #
        '#######################################
        '
        If (CutlistForm.CabCodeBox1.Text = "UPPER MICROWAVE") Then
            If (VarHeightI = 76 Or VarHeightI = 90 Or VarHeightI = 100) Then
                Dim Total = VarHeightI
                Select Case Total
                    Case 76
                        '#############################################
                        '# UPPER MICROWAVE GABLE RULES AND EQUATIONS #
                        '#############################################
                        '
                        UMGableX = 310 - 1
                        UMGableY = VarDepthI * 10
                        UMGableZ = 16
                        UMGableQ = 2 * VarAmountI
                        UMGableM = ""

                        '#############################################
                        '# UPPER MICROWAVE GABLE RULES AND EQUATIONS #
                        '#############################################
                        '
                        UMMIGableX = 450 - 1
                        UMMIGableY = VarDepthI * 10
                        UMMIGableZ = 16
                        UMMIGableQ = 2 * VarAmountI
                        UMMIGableM = ""

                    Case 90
                        '#############################################
                        '# UPPER MICROWAVE GABLE RULES AND EQUATIONS #
                        '#############################################
                        '
                        UMGableX = 440 - 1
                        UMGableY = VarDepthI * 10
                        UMGableZ = 16
                        UMGableQ = 2 * VarAmountI
                        UMGableM = ""

                        '#############################################
                        '# UPPER MICROWAVE GABLE RULES AND EQUATIONS #
                        '#############################################
                        '
                        UMMIGableX = 460 - 1
                        UMMIGableY = VarDepthI * 10
                        UMMIGableZ = 16
                        UMMIGableQ = 2 * VarAmountI
                        UMMIGableM = ""

                    Case 100
                        '#############################################
                        '# UPPER MICROWAVE GABLE RULES AND EQUATIONS #
                        '#############################################
                        '
                        UMGableX = 550 - 1
                        UMGableY = VarDepthI * 10
                        UMGableZ = 16
                        UMGableQ = 2 * VarAmountI
                        UMGableM = ""

                        '#############################################
                        '# UPPER MICROWAVE GABLE RULES AND EQUATIONS #
                        '#############################################
                        '
                        UMMIGableX = 450 - 1
                        UMMIGableY = VarDepthI * 10
                        UMMIGableZ = 16
                        UMMIGableQ = 2 * VarAmountI
                        UMMIGableM = ""
                    Case Else
                End Select

                '####################################################
                '# UPPER MICROWAVE TOP + BOTTOM RULES AND EQUATIONS #
                '####################################################
                '
                UMTopX = ((VarWidthI * 10) - (UMGableZ * 2)) - 1
                UMTopY = VarDepthI * 10
                UMTopZ = 16
                UMTopQ = 1 * VarAmountI
                UMTopM = ""

                '####################################################
                '# UPPER MICROWAVE TOP + BOTTOM RULES AND EQUATIONS #
                '####################################################
                '
                UMBtmX = ((VarWidthI * 10) - (UMGableZ * 2)) - 1
                UMBtmY = VarDepthI * 10
                UMBtmZ = 16
                UMBtmQ = 1 * VarAmountI
                UMBtmM = ""

                '########################################################
                '# UPPER MICROWAVE ADJUSTABLE SHELF RULES AND EQUATIONS #
                '########################################################
                '
                UMAdjShelfX = ((VarWidthI * 10) - (UMGableZ * 2)) - 5

                '##################################################
                '# GET SIZE OF SHELVES ACCORDING TO CABINET DEPTH #
                '##################################################
                '
                If VarDepthI <= 48 Then UMAdjShelfY = (VarDepthI * 10) - 30

                UMAdjShelfZ = 16
                UMAdjShelfM = ""

                '####################################################
                '# GET QUANITY OF SHELVES ACCORDING TO GABLE HEIGHT #
                '####################################################
                '
                If (UMGableX <= 497) Then UMAdjShelfQ = 0
                If (UMGableX >= 498) Then UMAdjShelfQ = 1 * VarAmountI
                CutlistForm.PubASQuantity = UMAdjShelfQ

                '###############################################
                '# UPPER MICROWAVE BACKING RULES AND EQUATIONS #
                '###############################################
                '
                UMBackY = (VarWidthI * 10) - 23
                UMBackX = UMGableX + 1 - 23
                UMBackZ = 3
                UMBackQ = 1 * VarAmountI
                UMBackM = ""

                '######################################################################
                '# UPPER MICROWAVE MATCHING INTERIOR TOP + BOTTOM RULES AND EQUATIONS #
                '######################################################################
                '
                UMMITopBtmX = ((VarWidthI * 10) - (UMMIGableZ * 2))
                UMMITopBtmY = VarDepthI * 10
                UMMITopBtmZ = 16
                UMMITopBtmQ = 2 * VarAmountI
                UMMITopBtmM = ""

                '###############################################################
                '# UPPER MICROWAVE MATCHING INTERIOR SHELF RULES AND EQUATIONS #
                '###############################################################
                '
                UMMIMicroSHX = ((VarWidthI * 10) - (UMMIGableZ * 2)) - 3

                '#####################
                '# GET SIZE OF SHELF #
                '#####################
                '
                UMMIMicroSHY = 460
                UMMIMicroSHZ = 16
                UMMIMicroSHM = ""

                '####################################################
                '# GET QUANITY OF SHELVES ACCORDING TO GABLE HEIGHT #
                '####################################################
                '
                UMMIMicroSHQ = 1 * VarAmountI

                '#################################################################
                '# UPPER MICROWAVE MATCHING INTERIOR BACKING RULES AND EQUATIONS #
                '#################################################################
                '
                UMMIBackY = (VarWidthI * 10) - 23
                UMMIBackX = UMMIGableX + 1 - 23
                UMMIBackZ = 16
                UMMIBackQ = 1 * VarAmountI
                UMMIBackM = ""
            Else
                MsgBox("ERROR: This is a not a standard size Upper Microwave! Click Special Check Box.")
                Exit Sub
            End If
        End If

        '#########################################
        '# UPPER CORNER DIAGONAL GABLE VARIABLES #
        '#########################################
        '
        Dim UCDGableX As Double 'DECLARES UPPER CORNER DIAGONAL GABLE X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UCDGableY As Double 'DECLARES UPPER CORNER DIAGONAL GABLE Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UCDGableZ As Double 'DECLARES UPPER CORNER DIAGONAL GABLE Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UCDGableQ As Integer 'DECLARES UPPER CORNER DIAGONAL GABLE QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim UCDGableM As String = "" 'DECLARES UPPER CORNER DIAGONAL GABLE MATERIAL AS STRING VARIABLE
        Dim UCDGEdgeSeq As String = "" 'DECLARES UPPER CORNER DIAGONAL GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim UCDGEdgeSeq2 As String = "" 'DECLARES UPPER CORNER DIAGONAL GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim UCDGEdgeCode As String = "" 'DECLARES UPPER CORNER DIAGONAL GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim UCDGEdgeCode2 As String = "" 'DECLARES UPPER CORNER DIAGONAL GABLE EDGE SEQUENCE AS A STRING VARIABLE

        '################################################
        '# UPPER CORNER DIAGONAL TOP + BOTTOM VARIABLES #
        '################################################
        '
        Dim UCDTopBtmX As Double 'DECLARES UPPER CORNER DIAGONAL TOP AND BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UCDTopBtmY As Double 'DECLARES UPPER CORNER DIAGONAL TOP AND BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UCDTopBtmZ As Double 'DECLARES UPPER CORNER DIAGONAL TOP AND BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UCDTopBtmQ As Integer 'DECLARES UPPER CORNER DIAGONAL TOP AND BOTTOM QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim UCDTopBtmM As String = "" 'DECLARES UPPER CORNER DIAGONAL TOP AND BOTTOM MATERIAL AS STRING VARIABLE
        Dim UCDTBEdgeSeq As String = "" 'DECLARES UPPER CORNER DIAGONAL TOP AND BOTTOM EDGE SEQUENCE AS A STRING VARIABLE
        Dim UCDTBEdgeCode As String = "" 'DECLARES UPPER CORNER DIAGONAL TOP AND BOTTOM EDGE SEQUENCE AS A STRING VARIABLE

        '######################################################
        '# UPPER CORNER DIAGONAL ADJUSTABLE SHELVES VARIABLES #
        '######################################################
        '
        Dim UCDAdjShelfX As Double 'DECLARES UPPER CORNER DIAGONAL ADJUSTABLE SHELF X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UCDAdjShelfY As Double 'DECLARES UPPER CORNER DIAGONAL ADJUSTABLE SHELF Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UCDAdjShelfZ As Double 'DECLARES UPPER CORNER DIAGONAL ADJUSTABLE SHELF Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UCDAdjShelfQ As Integer 'DECLARES UPPER CORNER DIAGONAL ADJUSTABLE SHELF QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim UCDAdjShelfM As String = "" 'DECLARES UPPER CORNER DIAGONAL ADJUSTABLE SHELF MATERIAL AS STRING VARIABLE
        Dim UCDASEdgeSeq As String = "" 'DECLARES UPPER CORNER DIAGONAL ADJUSTABLE SHELF EDGE SEQUENCE AS A STRING VARIABLE
        Dim UCDASEdgeCode As String = "" 'DECLARES UPPER CORNER DIAGONAL ADJUSTABLE SHELF EDGE SEQUENCE AS A STRING VARIABLE

        '##############################################
        '# UPPER CORNER DIAGONAL BACK STRAP VARIABLES #
        '##############################################
        '
        Dim UCDBackStrapX As Double 'DECLARES UPPER CORNER DIAGONAL BACK STRAP X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UCDBackStrapY As Double 'DECLARES UPPER CORNER DIAGONAL BACK STRAP Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UCDBackStrapZ As Double 'DECLARES UPPER CORNER DIAGONAL BACK STRAP Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UCDBackStrapQ As Integer 'DECLARES UPPER CORNER DIAGONAL BACK STRAP QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim UCDBackStrapM As String = "" 'DECLARES UPPER CORNER DIAGONAL BACK STRAP MATERIAL AS STRING VARIABLE

        '###########################################
        '# UPPER CORNER DIAGONAL BACKING VARIABLES #
        '###########################################
        '
        Dim UCDBackX As Double 'DECLARES UPPER CORNER DIAGONAL BACK X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UCDBackY As Double 'DECLARES UPPER CORNER DIAGONAL BACK Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UCDBackZ As Double 'DECLARES UPPER CORNER DIAGONAL BACK Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UCDBackQ As Integer 'DECLARES UPPER CORNER DIAGONAL BACK QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim UCDBackM As String = "" 'DECLARES UPPER CORNER DIAGONAL BACK MATERIAL AS STRING VARIABLE

        '###################################################
        '# UPPER CORNER DIAGONAL GABLE RULES AND EQUATIONS #
        '###################################################
        '
        UCDGableX = (VarHeightI * 10) - 1
        UCDGableY = VarDepthI * 10
        UCDGableZ = 19
        UCDGableQ = 2 * VarAmountI
        UCDGableM = ""

        '##########################################################
        '# UPPER CORNER DIAGONAL TOP + BOTTOM RULES AND EQUATIONS #
        '##########################################################
        '
        UCDTopBtmX = (VarWidthI * 10) - UCDGableZ + 1
        UCDTopBtmY = (VarWidthI * 10) - UCDGableZ + 1
        UCDTopBtmZ = 16
        UCDTopBtmQ = 2 * VarAmountI
        UCDTopBtmM = ""

        '##############################################################
        '# UPPER CORNER DIAGONAL ADJUSTABLE SHELF RULES AND EQUATIONS #
        '##############################################################
        '
        UCDAdjShelfX = (((VarWidthI * 10) - UCDGableZ) - 22) - 4
        UCDAdjShelfY = (((VarWidthI * 10) - UCDGableZ) - 22) - 4
        UCDAdjShelfZ = 16
        UCDAdjShelfM = ""

        '####################################################
        '# GET QUANITY OF SHELVES ACCORDING TO GABLE HEIGHT #
        '####################################################
        '
        If (UCDGableX <= 497) Then
            UCDAdjShelfQ = 0
        End If
        If (UCDGableX >= 498) Then
            UCDAdjShelfQ = 1 * VarAmountI
        End If
        If (UCDGableX >= 699) Then
            UCDAdjShelfQ = 2 * VarAmountI
        End If
        If (UCDGableX >= 898) Then
            UCDAdjShelfQ = 3 * VarAmountI
        End If
        If (UCDGableX >= 1440) Then
            UCDAdjShelfQ = 4 * VarAmountI
        End If
        CutlistForm.PubASQuantity = UCDAdjShelfQ

        '##########################################################
        '# UPPER CORNER DIAGONAL LONG BACKING RULES AND EQUATIONS #
        '##########################################################
        '
        UCDBackStrapX = (VarHeightI * 10) - 32
        UCDBackStrapY = 96
        UCDBackStrapZ = 16
        UCDBackStrapQ = 1 * VarAmountI
        UCDBackStrapM = ""

        '##########################################################
        '# UPPER CORNER DIAGONAL LONG BACKING RULES AND EQUATIONS #
        '##########################################################
        '
        UCDBackX = ((VarHeightI * 10) - (UCDTopBtmZ * 2)) + 10 - 1
        UCDBackY = (VarWidthI * 10) - 80
        UCDBackZ = 3
        UCDBackQ = 2 * VarAmountI
        UCDBackM = ""


        '###########################################################
        '# UPPER CORNER DIAGONAL MATCHING INTERIOR GABLE VARIABLES #
        '###########################################################
        '
        Dim UCDMIGableX As Double 'DECLARES UPPER CORNER DIAGONAL MATCHING INTERIOR GABLE X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UCDMIGableY As Double 'DECLARES UPPER CORNER DIAGONAL MATCHING INTERIOR GABLE Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UCDMIGableZ As Double 'DECLARES UPPER CORNER DIAGONAL MATCHING INTERIOR GABLE Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UCDMIGableQ As Integer 'DECLARES UPPER CORNER DIAGONAL MATCHING INTERIOR BACK QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim UCDMIGableM As String = "" 'DECLARES UPPER CORNER DIAGONAL MATCHING INTERIOR BACK MATERIAL AS STRING VARIABLE
        Dim UCDMIGEdgeSeq As String = "" 'DECLARES UPPER CORNER DIAGONAL MATCHING INTERIOR GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim UCDMIGEdgeSeq2 As String = "" 'DECLARES UPPER CORNER DIAGONAL MATCHING INTERIOR GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim UCDMIGEdgeCode As String = "" 'DECLARES UPPER CORNER DIAGONAL MATCHING INTERIOR GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim UCDMIGEdgeCode2 As String = "" 'DECLARES UPPER CORNER DIAGONAL MATCHING INTERIOR GABLE EDGE SEQUENCE AS A STRING VARIABLE

        '##################################################################
        '# UPPER CORNER DIAGONAL MATCHING INTERIOR TOP + BOTTOM VARIABLES #
        '##################################################################
        '
        Dim UCDMITopBtmX As Double 'DECLARES UPPER CORNER DIAGONAL MATCHING INTERIOR TOP AND BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UCDMITopBtmY As Double 'DECLARES UPPER CORNER DIAGONAL MATCHING INTERIOR TOP AND BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UCDMITopBtmZ As Double 'DECLARES UPPER CORNER DIAGONAL MATCHING INTERIOR TOP AND BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UCDMITopBtmQ As Integer 'DECLARES UPPER CORNER DIAGONAL MATCHING INTERIOR TOP AND BOTTOM QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim UCDMITopBtmM As String = "" 'DECLARES UPPER CORNER DIAGONAL MATCHING INTERIOR TOP AND BOTTOM MATERIAL AS STRING VARIABLE
        Dim UCDMITBEdgeSeq As String = "" 'DECLARES UPPER CORNER DIAGONAL MATCHING INTERIOR TOP AND BOTTOM EDGE SEQUENCE AS A STRING VARIABLE
        Dim UCDMITBEdgeCode As String = "" 'DECLARES UPPER CORNER DIAGONAL MATCHING INTERIOR TOP AND BOTTOM EDGE SEQUENCE AS A STRING VARIABLE

        '######################################################################
        '# UPPER CORNER DIAGONAL MATCHING INTERIOR ADJUSTABLE SHELF VARIABLES #
        '######################################################################
        '
        Dim UCDMIAdjShelfX As Double 'DECLARES UPPER CORNER DIAGONAL MATCHING INTERIOR ADJUSTABLE SHELF X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UCDMIAdjShelfY As Double 'DECLARES UPPER CORNER DIAGONAL MATCHING INTERIOR ADJUSTABLE SHELF Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UCDMIAdjShelfZ As Double 'DECLARES UPPER CORNER DIAGONAL MATCHING INTERIOR ADJUSTABLE SHELF Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UCDMIAdjShelfQ As Integer 'DECLARES UPPER CORNER DIAGONAL MATCHING INTERIOR ADJUSTABLE SHELF QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim UCDMIAdjShelfM As String = "" 'DECLARES UPPER CORNER DIAGONAL MATCHING INTERIOR ADJUSTABLE SHELF MATERIAL AS STRING VARIABLE
        Dim UCDMIASEdgeSeq As String = "" 'DECLARES UPPER CORNER DIAGONAL MATCHING INTERIOR ADJUSTABLE SHELF EDGE SEQUENCE AS A STRING VARIABLE
        Dim UCDMIASEdgeCode As String = "" 'DECLARES UPPER CORNER DIAGONAL MATCHING INTERIOR ADJUSTABLE SHELF EDGE SEQUENCE AS A STRING VARIABLE

        '##################################################################
        '# UPPER CORNER DIAGONAL MATCHING INTERIOR LONG BACKING VARIABLES #
        '##################################################################
        '
        Dim UCDMILBackX As Double 'DECLARES UPPER CORNER DIAGONAL MATCHING INTERIOR LONG BACK X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UCDMILBackY As Double 'DECLARES UPPER CORNER DIAGONAL MATCHING INTERIOR LONG BACK Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UCDMILBackZ As Double 'DECLARES UPPER CORNER DIAGONAL MATCHING INTERIOR LONG BACK Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UCDMILBackQ As Integer 'DECLARES UPPER CORNER DIAGONAL MATCHING INTERIOR LONG BACK QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim UCDMILBackM As String = "" 'DECLARES UPPER CORNER DIAGONAL MATCHING INTERIOR LONG BACK MATERIAL AS STRING VARIABLE

        '###################################################################
        '# UPPER CORNER DIAGONAL MATCHING INTERIOR SHORT BACKING VARIABLES #
        '###################################################################
        '
        Dim UCDMISBackX As Double 'DECLARES UPPER CORNER DIAGONAL MATCHING INTERIOR SHORT BACK X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UCDMISBackY As Double 'DECLARES UPPER CORNER DIAGONAL MATCHING INTERIOR SHORT BACK Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UCDMISBackZ As Double 'DECLARES UPPER CORNER DIAGONAL MATCHING INTERIOR SHORT BACK Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UCDMISBackQ As Integer 'DECLARES UPPER CORNER DIAGONAL MATCHING INTERIOR SHORT BACK QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim UCDMISBackM As String = "" 'DECLARES UPPER CORNER DIAGONAL MATCHING INTERIOR SHORT BACK MATERIAL AS STRING VARIABLE

        '#####################################################################
        '# UPPER CORNER DIAGONAL MATCHING INTERIOR GABLE RULES AND EQUATIONS #
        '#####################################################################
        '
        UCDMIGableX = VarHeightI * 10 - 1
        UCDMIGableY = VarDepthI * 10
        UCDMIGableZ = 19
        UCDMIGableQ = 2 * VarAmountI
        UCDMIGableM = ""

        '############################################################################
        '# UPPER CORNER DIAGONAL MATCHING INTERIOR TOP + BOTTOM RULES AND EQUATIONS #
        '############################################################################
        '
        UCDMITopBtmX = (VarWidthI * 10) - UCDMIGableZ + 1
        UCDMITopBtmY = (VarWidthI * 10) - UCDMIGableZ + 1
        UCDMITopBtmZ = 16
        UCDMITopBtmQ = 2 * VarAmountI
        UCDMITopBtmM = ""

        '################################################################################
        '# UPPER CORNER DIAGONAL MATCHING INTERIOR ADJUSTABLE SHELF RULES AND EQUATIONS #
        '################################################################################
        '
        UCDMIAdjShelfX = (((VarWidthI * 10) - UCDMIGableZ) - 35) - 4
        UCDMIAdjShelfY = (((VarWidthI * 10) - UCDMIGableZ) - 35) - 4
        UCDMIAdjShelfZ = 16
        UCDMIAdjShelfM = ""

        '####################################################
        '# GET QUANITY OF SHELVES ACCORDING TO GABLE HEIGHT #
        '####################################################
        '
        If (UCDMIGableX <= 497) Then
            UCDMIAdjShelfQ = 0
        End If
        If (UCDMIGableX >= 498) Then
            UCDMIAdjShelfQ = 1 * VarAmountI
        End If
        If (UCDMIGableX >= 699) Then
            UCDMIAdjShelfQ = 2 * VarAmountI
        End If
        If (UCDMIGableX >= 898) Then
            UCDMIAdjShelfQ = 3 * VarAmountI
        End If
        If (UCDMIGableX >= 1440) Then
            UCDMIAdjShelfQ = 4 * VarAmountI
        End If
        CutlistForm.PubASQuantity = UCDMIAdjShelfQ

        '############################################################################
        '# UPPER CORNER DIAGONAL MATCHING INTERIOR LONG BACKING RULES AND EQUATIONS #
        '############################################################################
        '
        UCDMILBackX = (((VarHeightI * 10) - (UCDMITopBtmZ * 2)) + 10) - 1
        UCDMILBackY = ((((VarWidthI * 10) - UCDMIGableZ) - 35) + UCDMITopBtmZ) + 5
        UCDMILBackZ = 16
        UCDMILBackQ = 1 * VarAmountI
        UCDMILBackM = ""

        '#############################################################################
        '# UPPER CORNER DIAGONAL MATCHING INTERIOR SHORT BACKING RULES AND EQUATIONS #
        '#############################################################################
        '
        UCDMISBackX = (((VarHeightI * 10) - (UCDMITopBtmZ * 2)) + 10) - 1
        UCDMISBackY = (((VarWidthI * 10) - UCDMIGableZ) - 35) + 10
        UCDMISBackZ = 16
        UCDMISBackQ = 1 * VarAmountI
        UCDMISBackM = ""

        '#####################################
        '# UPPER END SHELF CABINET VARIABLES #
        '#####################################
        '
        '###########################################
        '# UPPER END SHELF UPPER L-GABLE VARIABLES #
        '###########################################
        '
        Dim UESLGableX As Double 'DECLARES UPPER END SHELF GABLE X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UESLGableY As Double 'DECLARES UPPER END SHELF GABLE Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UESLGableZ As Double 'DECLARES UPPER END SHELF GABLE Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UESLGableQ As Integer 'DECLARES UPPER END SHELF GABLE QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim UESLGableM As String = "" 'DECLARES UPPER END SHELF GABLE MATERIAL AS STRING VARIABLE
        Dim UESLGEdgeSeq As String = "" 'DECLARES UPPER END SHELF UPPER GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim UESLGEdgeSeq2 As String = "" 'DECLARES UPPER END SHELF UPPER GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim UESLGEdgeCode As String = "" 'DECLARES UPPER END SHELF UPPER GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim UESLGEdgeCode2 As String = "" 'DECLARES UPPER END SHELF UPPER GABLE EDGE SEQUENCE AS A STRING VARIABLE

        '###########################################
        '# UPPER END SHELF UPPER S-GABLE VARIABLES #
        '###########################################
        '
        Dim UESSGableX As Double 'DECLARES UPPER END SHELF GABLE X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UESSGableY As Double 'DECLARES UPPER END SHELF GABLE Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UESSGableZ As Double 'DECLARES UPPER END SHELF GABLE Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UESSGableQ As Integer 'DECLARES UPPER END SHELF GABLE QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim UESSGableM As String = "" 'DECLARES UPPER END SHELF GABLE MATERIAL AS STRING VARIABLE
        Dim UESSGEdgeSeq As String = "" 'DECLARES UPPER END SHELF UPPER GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim UESSGEdgeSeq2 As String = "" 'DECLARES UPPER END SHELF UPPER GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim UESSGEdgeCode As String = "" 'DECLARES UPPER END SHELF UPPER GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim UESSGEdgeCode2 As String = "" 'DECLARES UPPER END SHELF UPPER GABLE EDGE SEQUENCE AS A STRING VARIABLE

        '#######################################
        '# UPPER END SHELF UPPER TOP VARIABLES #
        '#######################################
        '
        Dim UESTopX As Double 'DECLARES UPPER END SHELF TOP + BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UESTopY As Double 'DECLARES UPPER END SHELF TOP + BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UESTopZ As Double 'DECLARES UPPER END SHELF TOP + BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UESTopQ As Integer 'DECLARES UPPER END SHELF TOP + BOTTOM QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim UESTopM As String = "" 'DECLARES UPPER END SHELF TOP + BOTTOM MATERIAL AS STRING VARIABLE
        Dim UESTEdgeSeq As String = "" 'DECLARES UPPER END SHELF UPPER TOP EDGE SEQUENCE AS A STRING VARIABLE
        Dim UESTEdgeCode As String = "" 'DECLARES UPPER END SHELF UPPER TOP EDGE SEQUENCE AS A STRING VARIABLE
        Dim UESBEdgeSeq As String = "" 'DECLARES UPPER END SHELF UPPER BOTTOM EDGE SEQUENCE AS A STRING VARIABLE
        Dim UESBEdgeCode As String = "" 'DECLARES UPPER END SHELF UPPER BOTTOM EDGE SEQUENCE AS A STRING VARIABLE

        '####################################################
        '# UPPER END SHELF UPPER ADJUSTABLE SHELF VARIABLES #
        '####################################################
        '
        Dim UESFixedShelfX As Double 'DECLARES UPPER END SHELF ADJUSTABLE SHELF X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UESFixedShelfY As Double 'DECLARES UPPER END SHELF ADJUSTABLE SHELF Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UESFixedShelfZ As Double 'DECLARES UPPER END SHELF ADJUSTABLE SHELF Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UESFixedShelfQ As Integer 'DECLARES UPPER END SHELF ADJUSTABLE SHELF QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim UESFixedShelfM As String = "" 'DECLARES UPPER END SHELF ADJUSTABLE SHELF MATERIAL AS STRING VARIABLE
        Dim UESFSHEdgeSeq As String = "" 'DECLARES UPPER END SHELF UPPER ADJUSTABLE SHELF EDGE SEQUENCE AS A STRING VARIABLE
        Dim UESFSHEdgeCode As String = "" 'DECLARES UPPER END SHELF UPPER ADJUSTABLE SHELF EDGE SEQUENCE AS A STRING VARIABLE

        '################################################
        '# STANDARD UPPER END SHELF RULES AND EQUATIONS #
        '################################################
        '
        '#####################################################
        '# UPPER END SHELF UPPER L-GABLE RULES AND EQUATIONS #
        '#####################################################
        '
        UESLGableX = (VarHeightI * 10) - 1
        UESLGableY = VarWidthI * 10
        UESLGableZ = CutlistForm.GableThickness
        UESLGableQ = 1 * VarAmountI
        UESLGableM = ""

        '#####################################################
        '# UPPER END SHELF UPPER S-GABLE RULES AND EQUATIONS #
        '#####################################################
        '
        UESSGableX = (VarHeightI * 10) - 1
        UESSGableY = (VarDepthI * 10) - UESLGableZ
        UESSGableZ = CutlistForm.GableThickness
        UESSGableQ = 1 * VarAmountI
        UESSGableM = ""

        '#################################################
        '# UPPER END SHELF UPPER TOP RULES AND EQUATIONS #
        '#################################################
        '
        UESTopX = UESSGableY - 1
        UESTopY = (UESLGableY - UESSGableZ) - 1
        UESTopZ = CutlistForm.GableThickness
        UESTopQ = 1 * VarAmountI
        UESTopM = ""

        '##############################################################
        '# UPPER END SHELF UPPER ADJUSTABLE SHELF RULES AND EQUATIONS #
        '##############################################################
        '
        UESFixedShelfX = UESSGableY - 1

        '########################################################
        '# GET SIZE OF SHELVES ACCORDING TO UPPER CABINET DEPTH #
        '########################################################
        '
        UESFixedShelfY = (UESLGableY - UESSGableZ) - 1
        UESFixedShelfZ = CutlistForm.GableThickness
        UESFixedShelfM = ""

        '##########################################################
        '# GET QUANITY OF SHELVES ACCORDING TO UPPER GABLE HEIGHT #
        '##########################################################
        '
        If (UESLGableX <= 497) Then UESFixedShelfQ = 1
        If (UESLGableX >= 498) Then UESFixedShelfQ = 2 * VarAmountI
        If (UESLGableX >= 699) Then UESFixedShelfQ = 3 * VarAmountI
        If (UESLGableX >= 899) Then UESFixedShelfQ = 4 * VarAmountI
        If (UESLGableX >= 1479) Then UESFixedShelfQ = 5 * VarAmountI
        UESFixedShelfM = ""

        '########################
        '# BASE GABLE VARIABLES #
        '########################
        '
        Dim BGableX As Double 'DECLARES BASE GABLE X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BGableY As Double 'DECLARES BASE GABLE Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BGableZ As Double 'DECLARES BASE GABLE Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BGableQ As Integer 'DECLARES BASE GABLE QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim BGableM As String = "" 'DECLARES BASE GABLE MATERIAL AS STRING VARIABLE
        Dim BGEdgeSeq As String = "" 'DECLARES BASE GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim BGEdgeSeq2 As String = "" 'DECLARES BASE GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim BGEdgeCode As String = "" 'DECLARES BASE GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim BGEdgeCode2 As String = "" 'DECLARES BASE GABLE EDGE SEQUENCE AS A STRING VARIABLE

        '########################
        '# BEAD GABLE VARIABLES #
        '########################
        '
        Dim BEADSGableX As Double 'DECLARES BASE GABLE X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BEADSGableY As Double 'DECLARES BASE GABLE Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BEADSGableZ As Double 'DECLARES BASE GABLE Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BEADSGableQ As Integer 'DECLARES BASE GABLE QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim BEADLGableQ As Integer 'DECLARES BASE GABLE QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim BEADSGableM As String = "" 'DECLARES BASE GABLE MATERIAL AS STRING VARIABLE
        Dim BEADSGEdgeSeq As String = "" 'DECLARES BASE GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim BEADSGEdgeSeq2 As String = "" 'DECLARES BASE GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim BEADSGEdgeCode As String = "" 'DECLARES BASE GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim BEADSGEdgeCode2 As String = "" 'DECLARES BASE GABLE EDGE SEQUENCE AS A STRING VARIABLE

        '############################
        '# BASE TOB STRAP VARIABLES #
        '############################
        '
        Dim BTopStrapX As Double 'DECLARES BASE TOP STRAP X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BTopStrapY As Double 'DECLARES BASE TOP STRAP Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BTopStrapZ As Double 'DECLARES BASE TOP STRAP Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BTopStrapQ As Integer 'DECLARES BASE TOP STRAP QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim BTopStrapM As String = "" 'DECLARES BASE TOP STRAP MATERIAL AS STRING VARIABLE
        Dim BTSEdgeSeq As String = "" 'DECLARES BASE TOP STRAP EDGE SEQUENCE AS A STRING VARIABLE
        Dim BTSEdgeCode As String = "" 'DECLARES BASE TOP STRAP EDGE SEQUENCE AS A STRING VARIABLE

        '######################
        '# BASE TOP VARIABLES #
        '######################
        '
        Dim BTopX As Double 'DECLARES BASE TOP X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BTopY As Double 'DECLARES BASE TOP Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BTopZ As Double 'DECLARES BASE TOP Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BTopQ As Integer 'DECLARES BASE TOP QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim BTopM As String = "" 'DECLARES BASE TOP MATERIAL AS STRING VARIABLE
        Dim BTEdgeSeq As String = "" 'DECLARES BASE TOP EDGE SEQUENCE AS A STRING VARIABLE
        Dim BTEdgeCode As String = "" 'DECLARES BASE TOP EDGE SEQUENCE AS A STRING VARIABLE
        Dim BRTEdgeSeq As String = "" 'DECLARES BASE RANGE TOP EDGE SEQUENCE AS A STRING VARIABLE
        Dim BRTEdgeCode As String = "" 'DECLARES BASE RANGE TOP EDGE SEQUENCE AS A STRING VARIABLE

        '#########################
        '# BASE BOTTOM VARIABLES #
        '#########################
        '
        Dim BBtmX As Double 'DECLARES BASE BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BBtmY As Double 'DECLARES BASE BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BBtmZ As Double 'DECLARES BASE BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BBtmQ As Integer 'DECLARES BASE BOTTOM QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim BBtmM As String = "" 'DECLARES BASE BOTTOM MATERIAL AS STRING VARIABLE
        Dim BBEdgeSeq As String = "" 'DECLARES BASE BOTTOM EDGE SEQUENCE AS A STRING VARIABLE
        Dim BBEdgeCode As String = "" 'DECLARES BASE BOTTOM EDGE SEQUENCE AS A STRING VARIABLE

        '###################################
        '# BASE FULL FIXED SHELF VARIABLES #
        '###################################
        '
        Dim BFFShelfX As Double 'DECLARES BASE FULL FIXED SHELF X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BFFShelfY As Double 'DECLARES BASE FULL FIXED SHELF Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BFFShelfZ As Double 'DECLARES BASE FULL FIXED SHELF Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BFFShelfQ As Integer 'DECLARES BASE FULL FIXED SHELF QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim BFFShelfM As String = "" 'DECLARES BASE FULL FIXED SHELF MATERIAL AS STRING VARIABLE
        Dim BFFSHEdgeSeq As String = "" 'DECLARES FULL FIXED SHELF ADJUSTABLE SHELF EDGE SEQUENCE AS A STRING VARIABLE
        Dim BFFSHEdgeCode As String = "" 'DECLARES FULL FIXED SHELF ADJUSTABLE SHELF EDGE SEQUENCE AS A STRING VARIABLE

        '########################
        '# BASE STRAP VARIABLES #
        '########################
        '
        Dim BStrapX As Double 'DECLARES BASE STRAP X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BStrapY As Double 'DECLARES BASE STRAP Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BStrapZ As Double 'DECLARES BASE STRAP Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BStrapQ As Integer 'DECLARES BASE STRAP QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim BStrapM As String = "" 'DECLARES BASE STRAP MATERIAL AS STRING VARIABLE
        Dim BSEdgeSeq As String = "" 'DECLARES STRAP ADJUSTABLE SHELF EDGE SEQUENCE AS A STRING VARIABLE
        Dim BSEdgeCode As String = "" 'DECLARES STRAP ADJUSTABLE SHELF EDGE SEQUENCE AS A STRING VARIABLE

        '##########################
        '# BASE DIVIDER VARIABLES #
        '##########################
        '
        Dim BDividerX As Double 'DECLARES BASE DIVIDER X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BDividerY As Double 'DECLARES BASE DIVIDER Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BDividerZ As Double 'DECLARES BASE DIVIDER Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BDividerQ As Integer 'DECLARES BASE DIVIDER QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim BDividerM As String = "" 'DECLARES BASE DIVIDER MATERIAL AS STRING VARIABLE
        Dim BDEdgeSeq As String = "" 'DECLARES DIVIDER ADJUSTABLE SHELF EDGE SEQUENCE AS A STRING VARIABLE
        Dim BDEdgeCode As String = "" 'DECLARES DIVIDER ADJUSTABLE SHELF EDGE SEQUENCE AS A STRING VARIABLE

        '#####################################
        '# BASE ADJUSTABLE SHELVES VARIABLES #
        '#####################################
        '
        Dim BAdjShelfX As Double 'DECLARES BASE ADJUSTABLE SHELF X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BAdjShelfY As Double 'DECLARES BASE ADJUSTABLE SHELF Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BAdjShelfZ As Double 'DECLARES BASE ADJUSTABLE SHELF Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BAdjShelfQ As Integer 'DECLARES BASE ADJUSTABLE SHELF QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim BAdjShelfM As String = "" 'DECLARES BASE ADJUSTABLE SHELF MATERIAL AS STRING VARIABLE
        Dim BASEdgeSeq As String = "" 'DECLARES BASE ADJUSTABLE SHELF EDGE SEQUENCE AS A STRING VARIABLE
        Dim BASEdgeCode As String = "" 'DECLARES BASE ADJUSTABLE SHELF EDGE SEQUENCE AS A STRING VARIABLE

        '#####################################
        '# BEAD ADJUSTABLE SHELVES VARIABLES #
        '#####################################
        '
        Dim BEADAdjShelfX As Double 'DECLARES BASE ADJUSTABLE SHELF X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BEADAdjShelfY As Double 'DECLARES BASE ADJUSTABLE SHELF Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BEADAdjShelfZ As Double 'DECLARES BASE ADJUSTABLE SHELF Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BEADAdjShelfQ As Integer 'DECLARES BASE ADJUSTABLE SHELF QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim BEADAdjShelfM As String = "" 'DECLARES BASE ADJUSTABLE SHELF MATERIAL AS STRING VARIABLE
        Dim BEADASEdgeSeq As String = "" 'DECLARES BASE ADJUSTABLE SHELF EDGE SEQUENCE AS A STRING VARIABLE
        Dim BEADASEdgeCode As String = "" 'DECLARES BASE ADJUSTABLE SHELF EDGE SEQUENCE AS A STRING VARIABLE

        '##########################
        '# BASE BACKING VARIABLES #
        '##########################
        '
        Dim BBackX As Double 'DECLARES BASE BACK X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BBackY As Double 'DECLARES BASE BACK Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BBackZ As Double 'DECLARES BASE BACK Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BBackQ As Integer 'DECLARES BASE BACK QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim BBackM As String = "" 'DECLARES BASE BACK MATERIAL AS STRING VARIABLE
        Dim BBackEdgeSeq As String = "" 'DECLARES BASE BACK EDGE SEQUENCE AS A STRING VARIABLE
        Dim BBackEdgeCode As String = "" 'DECLARES BASE BACK EDGE SEQUENCE AS A STRING VARIABLE

        '##################################
        '# BASE GABLE RULES AND EQUATIONS #
        '##################################
        '
        BGableX = VarHeightI * 10
        BGableY = VarDepthI * 10
        BGableZ = 16
        BGableQ = 2 * VarAmountI
        BGableM = ""

        '##################################
        '# BEAD S-GABLE + L-GABLE RULES AND EQUATIONS #
        '##################################
        '
        BEADSGableX = VarHeightI * 10
        BEADSGableY = 32.3 * 10
        BEADSGableZ = 16
        BEADSGableQ = 1 * VarAmountI
        BEADLGableQ = 1 * VarAmountI

        '######################################
        '# BASE TOP STRAP RULES AND EQUATIONS #
        '######################################
        '
        BTopStrapX = ((VarWidthI * 10) - (BGableZ * 2)) - 1
        BTopStrapY = 115
        BTopStrapZ = 16
        BTopStrapQ = 2 * VarAmountI
        BTopStrapM = ""

        '################################
        '# BASE TOP RULES AND EQUATIONS #
        '################################
        '
        BTopX = ((VarWidthI * 10) - (BGableZ * 2)) - 1
        BTopY = VarDepthI * 10
        If (CutlistForm.BaseCabCodeBox.Text = "BR" Or CutlistForm.BaseCabCodeBox.Text = "BR2D" Or CutlistForm.BaseCabCodeBox.Text = "BSFM" Or CutlistForm.BaseCabCodeBox.Text = "BSFM2D") Then
            BTopZ = 19
        Else
            BTopZ = 16
        End If
        BTopQ = 1 * VarAmountI
        BTopM = ""

        '###################################
        '# BASE BOTTOM RULES AND EQUATIONS #
        '###################################
        '
        BBtmX = ((VarWidthI * 10) - (BGableZ * 2)) - 1
        BBtmY = VarDepthI * 10
        BBtmZ = 16
        BBtmQ = 1 * VarAmountI
        BBtmM = ""

        '##################################
        '# BASE STRAP RULES AND EQUATIONS #
        '##################################
        '
        BStrapX = ((VarWidthI * 10) - (BGableZ * 2)) - 1
        BStrapY = 60
        BStrapZ = 16
        Dim SSCase = CutlistForm.BaseCabCodeBox.Text
        Select Case SSCase
            Case "BSFF", "BTD", "BTDD"
                If (CutlistForm.WidthBox1.Text >= 60) Then
                    BStrapQ = 2 * VarAmountI
                Else
                    BStrapQ = 1 * VarAmountI
                End If
            Case Else
                BStrapQ = 1 * VarAmountI
        End Select

        BStrapM = ""

        '####################################
        '# BASE DIVIDER RULES AND EQUATIONS #
        '####################################
        '
        BDividerX = 139
        BDividerY = (VarDepthI * 10) - 22
        BDividerZ = 16
        BDividerQ = 2 * VarAmountI
        BDividerM = ""

        '#############################################
        '# BASE FULL FIXED SHELF RULES AND EQUATIONS #
        '#############################################
        '
        BFFShelfX = ((VarWidthI * 10) - (BGableZ * 2)) - 1
        BFFShelfY = (VarDepthI * 10) - 22
        BFFShelfZ = 16
        BFFShelfQ = 1 * VarAmountI
        BFFShelfM = ""

        '#############################################
        '# BASE ADJUSTABLE SHELF RULES AND EQUATIONS #
        '#############################################
        '

        BAdjShelfX = ((VarWidthI * 10) - (BGableZ * 2)) - 5

        If (VarDepthI <= 48) Then
            BAdjShelfY = (VarDepthI * 10) - 30
        Else
            BAdjShelfY = 530
        End If

        BAdjShelfZ = 16
        BAdjShelfQ = 1 * VarAmountI
        BAdjShelfM = ""
        CutlistForm.PubASQuantity = BAdjShelfQ

        '#############################################
        '# BEAD ADJUSTABLE SHELF RULES AND EQUATIONS #
        '#############################################
        '

        BEADAdjShelfX = ((VarWidthI * 10) - (BGableZ * 2)) - 5
        BEADAdjShelfY = 530
        BEADAdjShelfZ = 16
        BEADAdjShelfQ = 1 * VarAmountI
        CutlistForm.PubASQuantity = BEADAdjShelfQ

        '####################################
        '# BASE BACKING RULES AND EQUATIONS #
        '####################################
        '
        Dim BaseBackCase = CutlistForm.BaseCabCodeBox.Text
        Select Case BaseBackCase
            Case "BSFM"
                BBackX = (VarHeightI2 * 10) - 23
                BBackY = (VarWidthI * 10) - 23
                BBackZ = 3
                BBackQ = 1 * VarAmountI
                BBackM = ""
            Case "BR", "BR2D"
                BBackX = (((BGableX - 120) - BTopZ) - BBtmZ) + 10
                BBackY = (VarWidthI * 10) - 23
                BBackZ = 3
                BBackQ = 1 * VarAmountI
                BBackM = ""
            Case Else
                BBackX = ((VarHeightI * 10) - 120) - 26
                BBackY = (VarWidthI * 10) - 23
                BBackZ = 3
                BBackQ = 1 * VarAmountI
                BBackM = ""
        End Select

        '##############################################################
        '# BASE PENINSULA OPEN TOP, BOTTOM, AND FIXED SHELF VARIABLES #
        '##############################################################
        '
        Dim BPOPTopBtmFSHX As Double '
        Dim BPOPTopBtmFSHY As Double '
        Dim BPOPTopBtmFSHZ As Double '
        Dim BPOPTopBtmFSHQ As Double '
        Dim BPOPTopBtmFSHM As String = "" '
        Dim BPOPTBEdgeSeq As String = ""
        Dim BPOPTBEdgeCode As String = ""

        '#########################################
        '# BASE PENINSULA OPEN S-GABLE VARIABLES #
        '#########################################
        '
        Dim BPOPSGableX As Double '
        Dim BPOPSGableY As Double '
        Dim BPOPSGableZ As Double '
        Dim BPOPSGableQ As Double '
        Dim BPOPSGableM As String = "" '
        Dim BPOPSGEdgeSeq As String = ""
        Dim BPOPSGEdgeCode As String = ""

        '#########################################
        '# BASE PENINSULA OPEN L-GABLE VARIABLES #
        '#########################################
        '
        Dim BPOPLGableX As Double '
        Dim BPOPLGableY As Double '
        Dim BPOPLGableZ As Double '
        Dim BPOPLGableQ As Double '
        Dim BPOPLGableM As String = "" '
        Dim BPOPLGEdgeSeq As String = ""
        Dim BPOPLGEdgeCode As String = ""

        '########################################################################
        '# BASE PENINSULA OPEN TOP, BOTTOM, AND FIXED SHELF RULES AND EQUATIONS #
        '########################################################################
        '
        BPOPTopBtmFSHX = 391
        BPOPTopBtmFSHY = 364
        BPOPTopBtmFSHZ = 16
        BPOPTopBtmFSHQ = 3 * VarAmountI
        BPOPTopBtmFSHM = ""

        '#############################################################
        '# BASE PENINSULA OPEN S-GABLE + L-GABLE RULES AND EQUATIONS #
        '#############################################################
        '
        Dim BPOPCase = CutlistForm.OrientationBox.Text
        Select Case BPOPCase
            Case "LEFT"
                '###################################################
                '# BASE PENINSULA OPEN S-GABLE RULES AND EQUATIONS #
                '###################################################
                '
                BPOPSGableX = 883
                BPOPSGableY = 364
                BPOPSGableZ = 16
                BPOPSGableQ = 1
                BPOPSGableM = ""

                '###################################################
                '# BASE PENINSULA OPEN L-GABLE RULES AND EQUATIONS #
                '###################################################
                '
                BPOPLGableX = 883
                BPOPLGableY = 600
                BPOPLGableZ = 16
                BPOPLGableQ = 1
                BPOPLGableM = ""

            Case "RIGHT"
                '###################################################
                '# BASE PENINSULA OPEN S-GABLE RULES AND EQUATIONS #
                '###################################################
                '
                BPOPSGableX = 883
                BPOPSGableY = 364
                BPOPSGableZ = 16
                BPOPSGableQ = 1
                BPOPSGableM = ""

                '###################################################
                '# BASE PENINSULA OPEN L-GABLE RULES AND EQUATIONS #
                '###################################################
                '
                BPOPLGableX = 883
                BPOPLGableY = 600
                BPOPLGableZ = 16
                BPOPLGableQ = 1
                BPOPLGableM = ""

        End Select

        '##########################################
        '# BASE MATCHING INTERIOR GABLE VARIABLES #
        '##########################################
        '
        Dim BMIGableX As Double 'DECLARES BASE MATCHING INTERIOR GABLE X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMIGableY As Double 'DECLARES BASE MATCHING INTERIOR GABLE Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMIGableZ As Double 'DECLARES BASE MATCHING INTERIOR GABLE Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMIGableQ As Integer 'DECLARES BASE MATCHING INTERIOR GABLE QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim BMIGableM As String = "" 'DECLARES BASE MATCHING INTERIOR GABLE MATERIAL AS STRING VARIABLE
        Dim BMIGEdgeSeq As String = "" 'DECLARES BASE MATCHING INTERIOR GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim BMIGEdgeSeq2 As String = "" 'DECLARES BASE MATCHING INTERIOR GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim BMIGEdgeCode As String = "" 'DECLARES BASE MATCHING INTERIOR GABLE EDGE CODE AS A STRING VARIABLE
        Dim BMIGEdgeCode2 As String = "" 'DECLARES BASE MATCHING INTERIOR GABLE EDGE CODE AS A STRING VARIABLE

        '##############################################
        '# BASE MATCHING INTERIOR TOP STRAP VARIABLES #
        '##############################################
        '
        Dim BMITopStrapX As Double 'DECLARES BASE MATCHING INTERIOR TOP STRAP X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMITopStrapY As Double 'DECLARES BASE MATCHING INTERIOR TOP STRAP Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMITopStrapZ As Double 'DECLARES BASE MATCHING INTERIOR TOP STRAP Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMITopStrapQ As Integer 'DECLARES BASE MATCHING INTERIOR TOP STRAP QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim BMITopStrapM As String = "" 'DECLARES BASE MATCHING INTERIOR TOP STRAP MATERIAL AS STRING VARIABLE
        Dim BMITSEdgeSeq As String = "" 'DECLARES BASE MATCHING INTERIOR TOP STRAP EDGE SEQUENCE AS A STRING VARIABLE
        Dim BMITSEdgeCode As String = "" 'DECLARES BASE MATCHING INTERIOR TOP STRAP EDGE CODE AS A STRING VARIABLE

        '##########################################
        '# BASE MATCHING INTERIOR STRAP VARIABLES #
        '##########################################
        '
        Dim BMIStrapX As Double 'DECLARES BASE MATCHING INTERIOR TOP STRAP X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMIStrapY As Double 'DECLARES BASE MATCHING INTERIOR TOP STRAP Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMIStrapZ As Double 'DECLARES BASE MATCHING INTERIOR TOP STRAP Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMIStrapQ As Integer 'DECLARES BASE MATCHING INTERIOR TOP STRAP QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim BMIStrapM As String = "" 'DECLARES BASE MATCHING INTERIOR TOP STRAP MATERIAL AS STRING VARIABLE
        Dim BMISEdgeSeq As String = "" 'DECLARES BASE MATCHING INTERIOR STRAP EDGE SEQUENCE AS A STRING VARIABLE
        Dim BMISEdgeCode As String = "" 'DECLARES BASE MATCHING INTERIOR STRAP EDGE CODE AS A STRING VARIABLE

        '########################################
        '# BASE MATCHING INTERIOR TOP VARIABLES #
        '########################################
        '
        Dim BMITopX As Double 'DECLARES BASE MATCHING INTERIOR TOP X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMITopY As Double 'DECLARES BASE MATCHING INTERIOR TOP Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMITopZ As Double 'DECLARES BASE MATCHING INTERIOR TOP Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMITopQ As Integer 'DECLARES BASE MATCHING INTERIOR TOP QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim BMITopM As String = "" 'DECLARES BASE MATCHING INTERIOR TOP MATERIAL AS STRING VARIABLE
        Dim BMITEdgeSeq As String = "" 'DECLARES BASE MATCHING INTERIOR TOP EDGE SEQUENCE AS A STRING VARIABLE
        Dim BMITEdgeCode As String = "" 'DECLARES BASE MATCHING INTERIOR TOP EDGE CODE AS A STRING VARIABLE

        '############################################
        '# BASE MATCHING INTERIOR DIVIDER VARIABLES #
        '############################################
        '
        Dim BMIDividerX As Double 'DECLARES BASE MATCHING INTERIOR DIVIDER X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMIDividerY As Double 'DECLARES BASE MATCHING INTERIOR DIVIDER Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMIDividerZ As Double 'DECLARES BASE MATCHING INTERIOR DIVIDER Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMIDividerQ As Integer 'DECLARES BASE MATCHING INTERIOR DIVIDER QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim BMIDividerM As String = "" 'DECLARES BASE MATCHING INTERIOR DIVIDER MATERIAL AS STRING VARIABLE
        Dim BMIDEdgeSeq As String = "" 'DECLARES BASE MATCHING INTERIOR DIVIDER EDGE SEQUENCE AS A STRING VARIABLE
        Dim BMIDEdgeCode As String = "" 'DECLARES BASE MATCHING INTERIOR DIVIDER EDGE CODE AS A STRING VARIABLE

        '#################################################
        '# BASE MATCHING INTERIOR TOP + BOTTOM VARIABLES #
        '#################################################
        '
        Dim BMIBtmX As Double 'DECLARES BASE MATCHING INTERIOR BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMIBtmY As Double 'DECLARES BASE MATCHING INTERIOR BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMIBtmZ As Double 'DECLARES BASE MATCHING INTERIOR BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMIBtmQ As Integer 'DECLARES BASE MATCHING INTERIOR BOTTOM QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim BMIBtmM As String = "" 'DECLARES BASE MATCHING INTERIOR BOTTOM MATERIAL AS STRING VARIABLE
        Dim BMIBEdgeSeq As String = "" 'DECLARES BASE MATCHING INTERIOR BOTTOM EDGE SEQUENCE AS A STRING VARIABLE
        Dim BMIBEdgeCode As String = "" 'DECLARES BASE MATCHING INTERIOR BOTTOM EDGE CODE AS A STRING VARIABLE

        '#####################################################
        '# BASE MATCHING INTERIOR ADJUSTABLE SHELF VARIABLES #
        '#####################################################
        '
        Dim BMIAdjShelfX As Double 'DECLARES BASE MATCHING INTERIOR ADJUSTABLE SHELF X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMIAdjShelfY As Double 'DECLARES BASE MATCHING INTERIOR ADJUSTABLE SHELF Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMIAdjShelfZ As Double 'DECLARES BASE MATCHING INTERIOR ADJUSTABLE SHELF Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMIAdjShelfQ As Integer 'DECLARES BASE MATCHING INTERIOR ADJUSTABLE SHELF QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim BMIAdjShelfM As String = "" 'DECLARES BASE MATCHING INTERIOR ADJUSTABLE SHELF MATERIAL AS STRING VARIABLE
        Dim BMIASEdgeSeq As String = "" 'DECLARES BASE MATCHING INTERIOR ADJUSTABLE SHELF EDGE SEQUENCE AS A STRING VARIABLE
        Dim BMIASEdgeCode As String = "" 'DECLARES BASE MATCHING INTERIOR ADJUSTABLE SHELF EDGE CODE AS A STRING VARIABLE

        '############################################
        '# BASE MATCHING INTERIOR BACKING VARIABLES #
        '############################################
        '
        Dim BMIBackX As Double 'DECLARES BASE MATCHING INTERIOR BACK X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMIBackY As Double 'DECLARES BASE MATCHING INTERIOR BACK Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMIBackZ As Double 'DECLARES BASE MATCHING INTERIOR BACK Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMIBackQ As Integer 'DECLARES BASE MATCHING INTERIOR BACK QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim BMIBackM As String = "" 'DECLARES BASE MATCHING INTERIOR BACK MATERIAL AS STRING VARIABLE

        '###################################################
        '# BASE MATCING INTERIOR GABLE RULES AND EQUATIONS #
        '###################################################
        '
        BMIGableX = VarHeightI * 10
        BMIGableY = VarDepthI * 10
        BMIGableZ = 16
        BMIGableQ = 2 * VarAmountI
        BMIGableM = ""

        '####################################################
        '# BASE MATCING INTERIOR BOTTOM RULES AND EQUATIONS #
        '####################################################
        '
        BMITopStrapX = ((VarWidthI * 10) - (BMIGableZ * 2))
        BMITopStrapY = 300
        BMITopStrapZ = 16
        BMITopStrapQ = 1 * VarAmountI
        BMITopStrapM = ""

        '#################################################
        '# BASE MATCING INTERIOR TOP RULES AND EQUATIONS #
        '#################################################
        '
        BMITopX = ((VarWidthI * 10) - (BMIGableZ * 2))
        BMITopY = VarDepthI * 10
        BMITopZ = 16
        BMITopQ = 1 * VarAmountI
        BMITopM = ""

        '###################################################
        '# BASE MATCING INTERIOR STRAP RULES AND EQUATIONS #
        '###################################################
        '
        BMIStrapX = ((VarWidthI * 10) - (BMIGableZ * 2))
        BMIStrapY = 60
        BMIStrapZ = 16
        BMIStrapQ = 1 * VarAmountI
        BMIStrapM = ""

        '#####################################################
        '# BASE MATCING INTERIOR DIVIDER RULES AND EQUATIONS #
        '#####################################################
        '
        BMIDividerX = 139
        BMIDividerY = (VarDepthI * 10) - 22
        BMIDividerZ = 16
        BMIDividerQ = 2 * VarAmountI
        BMIDividerM = ""

        '####################################################
        '# BASE MATCING INTERIOR BOTTOM RULES AND EQUATIONS #
        '####################################################
        '
        BMIBtmX = ((VarWidthI * 10) - (BMIGableZ * 2))
        BMIBtmY = VarDepthI * 10
        BMIBtmZ = 16
        BMIBtmQ = 1 * VarAmountI
        BMIBtmM = ""

        '##############################################################
        '# BASE MATCING INTERIOR ADJUSTABLE SHELF RULES AND EQUATIONS #
        '##############################################################
        '
        BMIAdjShelfX = ((VarWidthI * 10) - (BMIGableZ * 2)) - 3
        BMIAdjShelfY = (VarDepthI * 10) - 30
        BMIAdjShelfZ = 16
        BMIAdjShelfQ = 1
        BMIAdjShelfM = ""
        CutlistForm.PubASQuantity = BMIAdjShelfQ

        '#####################################################
        '# BASE MATCING INTERIOR BACKING RULES AND EQUATIONS #
        '#####################################################
        '
        BMIBackY = (VarWidthI * 10) - 23
        BMIBackX = ((VarHeightI * 10) - 120) - 23
        BMIBackZ = 16
        BMIBackQ = 1 * VarAmountI
        BMIBackM = ""

        '######################################
        '# BASE MICROWAVE OPEN SHELF 1 DRAWER #
        '######################################
        '
        '######################################################
        '# BASE MICROWAVE OPEN SHELF 1 DRAWER GABLE VARIABLES #
        '######################################################
        '
        Dim BMO1DUGableX As Double 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER GABLE X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMO1DUGableY As Double 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER GABLE Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMO1DUGableZ As Double 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER GABLE Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMO1DUGableQ As Integer 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER GABLE QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim BMO1DUGableM As String = "" 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER GABLE MATERIAL AS STRING VARIABLE
        Dim BMO1DUGEdgeSeq As String = "" 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim BMO1DUGEdgeSeq2 As String = "" 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim BMO1DUGEdgeCode As String = "" 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER GABLE EDGE CODE AS A STRING VARIABLE
        Dim BMO1DUGEdgeCode2 As String = "" 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER GABLE EDGE CODE AS A STRING VARIABLE
        '
        Dim BMO1DBGableX As Double 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER GABLE X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMO1DBGableY As Double 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER GABLE Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMO1DBGableZ As Double 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER GABLE Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMO1DBGableQ As Integer 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER GABLE QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim BMO1DBGableM As String = "" 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER GABLE MATERIAL AS STRING VARIABLE
        Dim BMO1DBGEdgeSeq As String = "" 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim BMO1DBGEdgeCode As String = "" 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER GABLE EDGE SEQUENCE AS A STRING VARIABLE

        '####################################################
        '# BASE MICROWAVE OPEN SHELF 1 DRAWER TOP VARIABLES #
        '####################################################
        '
        Dim BMO1DUTopX As Double 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER TOP X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMO1DUTopY As Double 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER TOP Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMO1DUTopZ As Double 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER TOP Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMO1DUTopQ As Integer 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER TOP QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim BMO1DUTopM As String = "" 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER TOP MATERIAL AS STRING VARIABLE
        Dim BMO1DUTEdgeSeq As String = "" 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER TOP EDGE SEQUENCE AS A STRING VARIABLE
        Dim BMO1DUTEdgeCode As String = "" 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER TOP EDGE CODE AS A STRING VARIABLE

        '##########################################################
        '# BASE MICROWAVE OPEN SHELF 1 DRAWER TOP STRAP VARIABLES #
        '##########################################################
        '
        Dim BMO1DBTopStrapX As Double 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER TOP X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMO1DBTopStrapY As Double 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER TOP Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMO1DBTopStrapZ As Double 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER TOP Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMO1DBTopStrapQ As Integer 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER TOP QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim BMO1DBTopStrapM As String = "" 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER TOP MATERIAL AS STRING VARIABLE
        Dim BMO1DBTSEdgeSeq As String = "" 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER TOP STRAP EDGE SEQUENCE AS A STRING VARIABLE
        Dim BMO1DBTSEdgeCode As String = "" 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER TOP STRAP EDGE CODE AS A STRING VARIABLE

        '#######################################################
        '# BASE MICROWAVE OPEN SHELF 1 DRAWER BOTTOM VARIABLES #
        '#######################################################
        '
        Dim BMO1DUBtmX As Double 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMO1DUBtmY As Double 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMO1DUBtmZ As Double 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMO1DUBtmQ As Integer 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER BOTTOM QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim BMO1DUBtmM As String = "" 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER BOTTOM MATERIAL AS STRING VARIABLE
        Dim BMO1DUBEdgeSeq As String = "" 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER BOTTOM EDGE SEQUENCE AS A STRING VARIABLE
        Dim BMO1DUBEdgeCode As String = "" 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER BOTTOM EDGE CODE AS A STRING VARIABLE
        '
        Dim BMO1DBBtmX As Double 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMO1DBBtmY As Double 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMO1DBBtmZ As Double 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMO1DBBtmQ As Integer 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER BOTTOM QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim BMO1DBBtmM As String = "" 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER BOTTOM MATERIAL AS STRING VARIABLE
        Dim BMO1DBBEdgeSeq As String = "" 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER BOTTOM EDGE SEQUENCE AS A STRING VARIABLE
        Dim BMO1DBBEdgeCode As String = "" 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER BOTTOM EDGE CODE AS A STRING VARIABLE

        '#####################################################
        '# BASE MICROWAVE OPEN SHELF 1 DRAWER BACK VARIABLES #
        '#####################################################
        '
        Dim BMO1DUBackX As Double 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER BACK X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMO1DUBackY As Double 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER BACK Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMO1DUBackZ As Double 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER BACK Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMO1DUBackQ As Integer 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER BACK QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim BMO1DUBackM As String = "" 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER BACK MATERIAL AS STRING VARIABLE
        '
        Dim BMO1DBBackX As Double 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER BACK X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMO1DBBackY As Double 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER BACK Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMO1DBBackZ As Double 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER BACK Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BMO1DBBackQ As Integer 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER BACK QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim BMO1DBBackM As String = "" 'DECLARES BASE MICROWAVE OPEN SHELF 1 DRAWER BACK MATERIAL AS STRING VARIABLE

        '################################################################
        '# BASE MICROWAVE OPEN SHELF 1 DRAWER GABLE RULES AND EQUATIONS #
        '################################################################
        '
        BMO1DUGableX = (VarHeightI * 10) - 1
        BMO1DUGableY = VarDepthI * 10
        BMO1DUGableZ = 16
        BMO1DUGableQ = 2 * VarAmountI
        BMO1DUGableM = ""

        '##############################################################
        '# BASE MICROWAVE OPEN SHELF 1 DRAWER TOP RULES AND EQUATIONS #
        '##############################################################
        '
        BMO1DUTopX = ((VarWidthI * 10) - (UMIGableZ * 2))
        BMO1DUTopY = VarDepthI * 10
        BMO1DUTopZ = 16
        BMO1DUTopQ = 1 * VarAmountI
        BMO1DUTopM = ""

        '#################################################################
        '# BASE MICROWAVE OPEN SHELF 1 DRAWER BOTTOM RULES AND EQUATIONS #
        '#################################################################
        '
        BMO1DUBtmX = ((VarWidthI * 10) - (BMO1DUGableZ * 2))
        BMO1DUBtmY = VarDepthI * 10
        BMO1DUBtmZ = 16
        BMO1DUBtmQ = 1 * VarAmountI
        BMO1DUBtmM = ""

        '##################################################################
        '# BASE MICROWAVE OPEN SHELF 1 DRAWER BACKING RULES AND EQUATIONS #
        '##################################################################
        '
        BMO1DUBackY = (VarWidthI * 10) - 23
        BMO1DUBackX = (VarHeightI * 10) - 23
        BMO1DUBackZ = 16
        BMO1DUBackQ = 1 * VarAmountI
        BMO1DUBackM = ""

        '################################################################
        '# BASE MICROWAVE OPEN SHELF 1 DRAWER GABLE RULES AND EQUATIONS #
        '################################################################
        '
        BMO1DBGableX = VarHeightI * 10
        BMO1DBGableY = VarDepthI * 10
        BMO1DBGableZ = 16
        BMO1DBGableQ = 2 * VarAmountI
        BMO1DBGableM = ""

        '####################################################################
        '# BASE MICROWAVE OPEN SHELF 1 DRAWER TOP STRAP RULES AND EQUATIONS #
        '####################################################################
        '
        BMO1DBTopStrapX = ((VarWidthI * 10) - (BMO1DBGableZ * 2))
        BMO1DBTopStrapY = 300
        BMO1DBTopStrapZ = 16
        BMO1DBTopStrapQ = 1 * VarAmountI
        BMO1DBTopStrapM = ""

        '#################################################################
        '# BASE MICROWAVE OPEN SHELF 1 DRAWER BOTTOM RULES AND EQUATIONS #
        '#################################################################
        '
        BMO1DBBtmX = ((VarWidthI * 10) - (BMO1DBGableZ * 2))
        BMO1DBBtmY = VarDepthI * 10
        BMO1DBBtmZ = 16
        BMO1DBBtmQ = 1 * VarAmountI
        BMO1DBBtmM = ""

        '##################################################################
        '# BASE MICROWAVE OPEN SHELF 1 DRAWER BACKING RULES AND EQUATIONS #
        '##################################################################
        '
        BMO1DBBackY = (VarWidthI * 10) - 23
        BMO1DBBackX = ((VarHeightI * 10) - 120) - 23
        BMO1DBBackZ = 16
        BMO1DBBackQ = 1 * VarAmountI
        BMO1DBBackM = ""

        '#########################################
        '# TALL UTILITY 1 UNIT CABINET VARIABLES #
        '#########################################
        '
        '#######################################
        '# TALL UTILITY 1 UNIT GABLE VARIABLES #
        '#######################################
        '
        Dim TU1UGableX As Double 'DECLARES TALL UTILITY 1 UNIT GABLE X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU1UGableY As Double 'DECLARES TALL UTILITY 1 UNIT GABLE Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU1UGableZ As Double 'DECLARES TALL UTILITY 1 UNIT GABLE Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU1UGableQ As Integer 'DECLARES TALL UTILITY 1 UNIT GABLE QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim TU1UGableM As String = "" 'DECLARES TALL UTILITY 1 UNIT GABLE MATERIAL AS STRING VARIABLE
        Dim TU1UGEdgeSeq As String = "" 'DECLARES TALL UTILITY 1 UNIT GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim TU1UGEdgeCode As String = "" 'DECLARES TALL UTILITY 1 UNIT GABLE EDGE SEQUENCE AS A STRING VARIABLE

        '##############################################
        '# TALL UTILITY 1 UNIT TOP + BOTTOM VARIABLES #
        '##############################################
        '
        Dim TU1UTopX As Double 'DECLARES TALL UTILITY 1 UNIT TOP X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU1UTopY As Double 'DECLARES TALL UTILITY 1 UNIT TOP Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU1UTopZ As Double 'DECLARES TALL UTILITY 1 UNIT TOP Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU1UTopQ As Integer 'DECLARES TALL UTILITY 1 UNIT TOP QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim TU1UTopM As String = "" 'DECLARES TALL UTILITY 1 UNIT TOP STRAP MATERIAL AS STRING VARIABLE
        Dim TU1UTEdgeSeq As String = "" 'DECLARES TALL UTILITY 1 UNIT TOP EDGE SEQUENCE AS A STRING VARIABLE
        Dim TU1UTEdgeCode As String = "" 'DECLARES TALL UTILITY 1 UNIT TOP EDGE SEQUENCE AS A STRING VARIABLE

        '##############################################
        '# TALL UTILITY 1 UNIT TOP + BOTTOM VARIABLES #
        '##############################################
        '
        Dim TU1UBotX As Double 'DECLARES TALL UTILITY 1 UNIT BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU1UBotY As Double 'DECLARES TALL UTILITY 1 UNIT BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU1UBotZ As Double 'DECLARES TALL UTILITY 1 UNIT BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU1UBotQ As Integer 'DECLARES TALL UTILITY 1 UNIT BOTTOM QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim TU1UBotM As String = "" 'DECLARES TALL UTILITY 1 UNIT BOTTOM MATERIAL AS STRING VARIABLE
        Dim TU1UBEdgeSeq As String = "" 'DECLARES TALL UTILITY 1 UNIT BOTTOM EDGE SEQUENCE AS A STRING VARIABLE
        Dim TU1UBEdgeCode As String = "" 'DECLARES TALL UTILITY 1 UNIT BOTTOM EDGE SEQUENCE AS A STRING VARIABLE

        '##################################################
        '# TALL UTILITY 1 UNIT FULL FIXED SHELF VARIABLES #
        '##################################################
        '
        Dim TU1UFFSX As Double 'DECLARES TALL UTILITY 1 UNIT FULL FIXED SHELF X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU1UFFSY As Double 'DECLARES TALL UTILITY 1 UNIT FULL FIXED SHELF  Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU1UFFSZ As Double 'DECLARES TALL UTILITY 1 UNIT FULL FIXED SHELF  Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU1UFFSQ As Integer 'DECLARES TALL UTILITY 1 UNIT FULL FIXED SHELF  QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim TU1UFFSM As String = "" 'DECLARES TALL UTILITY 1 UNIT FULL FIXED SHELF  MATERIAL AS STRING VARIABLE
        Dim TU1UFFSEdgeSeq As String = "" 'DECLARES TALL UTILITY 1 UNIT FULL FIXED SHELF EDGE SEQUENCE AS A STRING VARIABLE
        Dim TU1UFFSEdgeCode As String = "" 'DECLARES TALL UTILITY 1 UNIT FULL FIXED SHELF EDGE SEQUENCE AS A STRING VARIABLE

        '#######################################
        '# TALL UTILITY 1 UNIT STRAP VARIABLES #
        '#######################################
        '
        Dim TU1UStrapX As Double 'DECLARES TALL UTILITY 1 UNIT STRAP X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU1UStrapY As Double 'DECLARES TALL UTILITY 1 UNIT STRAP Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU1UStrapZ As Double 'DECLARES TALL UTILITY 1 UNIT STRAP Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU1UStrapQ As Integer 'DECLARES TALL UTILITY 1 UNIT STRAP QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim TU1UStrapM As String = "" 'DECLARES TALL UTILITY 1 UNIT STRAP MATERIAL AS STRING VARIABLE
        Dim TU1USEdgeSeq As String = "" 'DECLARES TALL UTILITY 1 UNIT STRAP EDGE SEQUENCE AS A STRING VARIABLE
        Dim TU1USEdgeCode As String = "" 'DECLARES TALL UTILITY 1 UNIT STRAP EDGE SEQUENCE AS A STRING VARIABLE

        '#####################################################
        '# TALL UTILITY 1 UNIT SHALLOW FIXED SHELF VARIABLES #
        '#####################################################
        '
        Dim TU1USFSX As Double 'DECLARES TALL UTILITY 1 UNIT SHALLOW FIXED SHELF X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU1USFSY As Double 'DECLARES TALL UTILITY 1 UNIT SHALLOW FIXED SHELF Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU1USFSZ As Double 'DECLARES TALL UTILITY 1 UNIT SHALLOW FIXED SHELF Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU1USFSQ As Integer 'DECLARES TALL UTILITY 1 UNIT SHALLOW FIXED SHELF QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim TU1USFSM As String = "" 'DECLARES TALL UTILITY 1 UNIT SHALLOW FIXED SHELF MATERIAL AS STRING VARIABLE
        Dim TU1USFSEdgeSeq As String = "" 'DECLARES TALL UTILITY 1 UNIT SHALLOW FIXED SHELF EDGE SEQUENCE AS A STRING VARIABLE
        Dim TU1USFSEdgeCode As String = "" 'DECLARES TALL UTILITY 1 UNIT SHALLOW FIXED SHELF EDGE SEQUENCE AS A STRING VARIABLE

        '##################################################
        '# TALL UTILITY 1 UNIT ADJUSTABLE SHELF VARIABLES #
        '##################################################
        '
        Dim TU1UAdjShelfX As Double 'DECLARES TALL UTILITY 1 UNIT ADJUSTABLE SHELF X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU1UAdjShelfY As Double 'DECLARES TALL UTILITY 1 UNIT ADJUSTABLE SHELF Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU1UAdjShelfZ As Double 'DECLARES TALL UTILITY 1 UNIT ADJUSTABLE SHELF Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU1UAdjShelfQ As Integer 'DECLARES TALL UTILITY 1 UNIT ADJUSTABLE SHELF QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim TU1UAdjShelfM As String = "" 'DECLARES TALL UTILITY 1 UNIT ADJUSTABLE SHELF MATERIAL AS STRING VARIABLE
        Dim TU1UASEdgeSeq As String = "" 'DECLARES TALL UTILITY 1 UNIT ADJUSTABLE SHELF EDGE SEQUENCE AS A STRING VARIABLE
        Dim TU1UASEdgeCode As String = "" 'DECLARES TALL UTILITY 1 UNIT ADJUSTABLE SHELF EDGE SEQUENCE AS A STRING VARIABLE

        '#########################################
        '# TALL UTILITY 1 UNIT BACKING VARIABLES #
        '#########################################
        '
        Dim TU1UBackX As Double 'DECLARES TALL UTILITY 1 UNIT BACK X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU1UBackY As Double 'DECLARES TALL UTILITY 1 UNIT BACK Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU1UBackZ As Double 'DECLARES TALL UTILITY 1 UNIT BACK Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU1UBackQ As Integer 'DECLARES TALL UTILITY 1 UNIT BACK QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim TU1UBackM As String = "" 'DECLARES TALL UTILITY 1 UNIT BACK MATERIAL AS STRING VARIABLE

        '#################################################
        '# TALL UTILITY 1 UNIT GABLE RULES AND EQUATIONS #
        '#################################################
        '
        TU1UGableX = VarHeightI * 10
        TU1UGableY = VarDepthI * 10
        TU1UGableZ = 16
        TU1UGableQ = 2 * VarAmountI
        TU1UGableM = ""

        '###############################################
        '# TALL UTILITY 1 UNIT TOP RULES AND EQUATIONS #
        '###############################################
        '
        TU1UTopX = ((VarWidthI * 10) - (TU1UGableZ * 2)) - 1
        TU1UTopY = VarDepthI * 10
        TU1UTopZ = 16
        TU1UTopQ = 1 * VarAmountI
        TU1UTopM = ""

        '##################################################
        '# TALL UTILITY 1 UNIT BOTTOM RULES AND EQUATIONS #
        '##################################################
        '
        TU1UBotX = ((VarWidthI * 10) - (TU1UGableZ * 2)) - 1
        TU1UBotY = VarDepthI * 10
        TU1UBotZ = 16
        TU1UBotQ = 1 * VarAmountI
        TU1UBotM = ""

        '############################################################
        '# TALL UTILITY 1 UNIT ADJUSTABLE SHELF RULES AND EQUATIONS #
        '############################################################
        '
        TU1UAdjShelfX = ((VarWidthI * 10) - (TU1UGableZ * 2)) - 5

        '##################################################
        '# GET SIZE OF SHELVES ACCORDING TO CABINET DEPTH #
        '##################################################
        '
        If VarDepthI <= 48 Then TU1UAdjShelfY = (VarDepthI * 10) - 30
        If VarDepthI > 48 Then TU1UAdjShelfY = 450
        If VarDepthI > 62 Then TU1UAdjShelfY = 530
        TU1UAdjShelfZ = 16

        '######################################################
        '# GET QUANITY OF SHELVES ACCORDING TO CABINET HEIGHT #
        '######################################################
        '
        If (VarHeightI = 209) Then TU1UAdjShelfQ = 4 * VarAmountI
        If (VarHeightI = 214) Then TU1UAdjShelfQ = 3 * VarAmountI
        If (VarHeightI = 220) Then TU1UAdjShelfQ = 4 * VarAmountI
        If (VarHeightI = 232) Then TU1UAdjShelfQ = 4 * VarAmountI
        If (VarHeightI = 242) Then TU1UAdjShelfQ = 5 * VarAmountI
        TU1UAdjShelfM = ""
        CutlistForm.PubASQuantity = TU1UAdjShelfQ

        '############################################################
        '# TALL UTILITY 1 UNIT FULL FIXED SHELF RULES AND EQUATIONS #
        '############################################################
        '
        TU1UFFSX = ((VarWidthI * 10) - (TU1UGableZ * 2)) - 1
        TU1UFFSY = (VarDepthI * 10) - 22
        TU1UFFSZ = 16
        TU1UFFSQ = 1 * VarAmountI
        TU1UFFSM = ""

        '#################################################
        '# TALL UTILITY 1 UNIT STRAP RULES AND EQUATIONS #
        '#################################################
        '
        TU1UStrapX = ((VarWidthI * 10) - (TU1UGableZ * 2)) - 1
        TU1UStrapY = 60
        TU1UStrapZ = 16
        TU1UStrapQ = 1 * VarAmountI
        TU1UStrapM = ""

        '###############################################################
        '# TALL UTILITY 1 UNIT SHALLOW FIXED SHELF RULES AND EQUATIONS #
        '###############################################################
        '
        TU1USFSX = ((VarWidthI * 10) - (TU1UGableZ * 2)) - 1
        If VarDepthI <= 48 Then TU1USFSY = (VarDepthI * 10) - 30
        If VarDepthI > 48 Then TU1USFSY = 450
        If VarDepthI > 62 Then TU1USFSY = 530
        TU1USFSZ = 16
        TU1USFSQ = 1 * VarAmountI
        TU1USFSM = ""

        '################################################
        '# TALL UTILITY 1 UNIT BACK RULES AND EQUATIONS #
        '################################################
        '
        TU1UBackX = ((VarHeightI * 10) - 120) - 23
        TU1UBackY = (VarWidthI * 10) - 23
        TU1UBackZ = 3
        TU1UBackQ = 1 * VarAmountI
        TU1UBackM = ""

        '#############################################
        '# TALL UTILITY 2 UNIT UPPER GABLE VARIABLES #
        '#############################################
        '
        Dim TU2UUGableX As Double 'DECLARES TALL UTILITY 2 UNIT GABLE X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UUGableY As Double 'DECLARES TALL UTILITY 2 UNIT GABLE Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UUGableZ As Double 'DECLARES TALL UTILITY 2 UNIT GABLE Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UUGableQ As Integer 'DECLARES TALL UTILITY 2 UNIT GABLE QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim TU2UUGableM As String = "" 'DECLARES TALL UTILITY 2 UNIT GABLE MATERIAL AS STRING VARIABLE
        Dim TU2UUGEdgeSeq As String = "" 'DECLARES TALL UTILITY 2 UNIT UPPER GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim TU2UUGEdgeSeq2 As String = "" 'DECLARES TALL UTILITY 2 UNIT UPPER GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim TU2UUGEdgeCode As String = "" 'DECLARES TALL UTILITY 2 UNIT UPPER GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim TU2UUGEdgeCode2 As String = "" 'DECLARES TALL UTILITY 2 UNIT UPPER GABLE EDGE SEQUENCE AS A STRING VARIABLE

        '####################################################
        '# TALL UTILITY 2 UNIT UPPER TOP + BOTTOM VARIABLES #
        '####################################################
        '
        Dim TU2UUTopBtmX As Double 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UUTopBtmY As Double 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UUTopBtmZ As Double 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UUTopBtmQ As Integer 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim TU2UUTopBtmM As String = "" 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM MATERIAL AS STRING VARIABLE
        Dim TU2UUTBEdgeSeq As String = "" 'DECLARES TALL UTILITY 2 UNIT UPPER TOP AND BOTTOM EDGE SEQUENCE AS A STRING VARIABLE
        Dim TU2UUTBEdgeCode As String = "" 'DECLARES TALL UTILITY 2 UNIT UPPER TOP AND BOTTOM EDGE SEQUENCE AS A STRING VARIABLE

        '###########################################
        '# TALL UTILITY 2 UNIT UPPER TOP VARIABLES #
        '###########################################
        '
        Dim TU2UUTopX As Double 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UUTopY As Double 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UUTopZ As Double 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UUTopQ As Integer 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim TU2UUTopM As String = "" 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM MATERIAL AS STRING VARIABLE
        Dim TU2UUTEdgeSeq As String = "" 'DECLARES TALL UTILITY 2 UNIT UPPER TOP EDGE SEQUENCE AS A STRING VARIABLE
        Dim TU2UUTEdgeCode As String = "" 'DECLARES TALL UTILITY 2 UNIT UPPER TOP EDGE SEQUENCE AS A STRING VARIABLE

        '##############################################
        '# TALL UTILITY 2 UNIT UPPER BOTTOM VARIABLES #
        '##############################################
        '
        Dim TU2UUBotX As Double 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UUBotY As Double 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UUBotZ As Double 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UUBotQ As Integer 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim TU2UUBotM As String = "" 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM MATERIAL AS STRING VARIABLE
        Dim TU2UUBEdgeSeq As String = "" 'DECLARES TALL UTILITY 2 UNIT UPPER BOTTOM EDGE SEQUENCE AS A STRING VARIABLE
        Dim TU2UUBEdgeCode As String = "" 'DECLARES TALL UTILITY 2 UNIT UPPER BOTTOM EDGE SEQUENCE AS A STRING VARIABLE

        '########################################################
        '# TALL UTILITY 2 UNIT UPPER ADJUSTABLE SHELF VARIABLES #
        '########################################################
        '
        Dim TU2UUAdjShelfX As Double 'DECLARES TALL UTILITY 2 UNIT ADJUSTABLE SHELF X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UUAdjShelfY As Double 'DECLARES TALL UTILITY 2 UNIT ADJUSTABLE SHELF Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UUAdjShelfZ As Double 'DECLARES TALL UTILITY 2 UNIT ADJUSTABLE SHELF Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UUAdjShelfQ As Integer 'DECLARES TALL UTILITY 2 UNIT ADJUSTABLE SHELF QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim TU2UUAdjShelfM As String = "" 'DECLARES TALL UTILITY 2 UNIT ADJUSTABLE SHELF MATERIAL AS STRING VARIABLE        
        Dim TU2UUASEdgeSeq As String = "" 'DECLARES TALL UTILITY 2 UNIT UPPER ADJUSTABLE SHELF EDGE SEQUENCE AS A STRING VARIABLE
        Dim TU2UUASEdgeCode As String = "" 'DECLARES TALL UTILITY 2 UNIT UPPER ADJUSTABLE SHELF EDGE SEQUENCE AS A STRING VARIABLE

        '###############################################
        '# TALL UTILITY 2 UNIT UPPER BACKING VARIABLES #
        '###############################################
        '
        Dim TU2UUBackX As Double 'DECLARES TALL UTILITY 2 UNIT BACK X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UUBackY As Double 'DECLARES TALL UTILITY 2 UNIT BACK Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UUBackZ As Double 'DECLARES TALL UTILITY 2 UNIT BACK Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UUBackQ As Integer 'DECLARES TALL UTILITY 2 UNIT BACK QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim TU2UUBackM As String = "" 'DECLARES TALL UTILITY 2 UNIT BACK MATERIAL AS STRING VARIABLE

        '#############################################
        '# TALL UTILITY 2 UNIT UPPER GABLE VARIABLES #
        '#############################################
        '
        Dim TU2UUMIGableX As Double 'DECLARES TALL UTILITY 2 UNIT GABLE X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UUMIGableY As Double 'DECLARES TALL UTILITY 2 UNIT GABLE Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UUMIGableZ As Double 'DECLARES TALL UTILITY 2 UNIT GABLE Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UUMIGableQ As Integer 'DECLARES TALL UTILITY 2 UNIT GABLE QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim TU2UUMIGableM As String = "" 'DECLARES TALL UTILITY 2 UNIT GABLE MATERIAL AS STRING VARIABLE
        Dim TU2UUMIGEdgeSeq As String = "" 'DECLARES TALL UTILITY 2 UNIT UPPER GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim TU2UUMIGEdgeSeq2 As String = "" 'DECLARES TALL UTILITY 2 UNIT UPPER GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim TU2UUMIGEdgeCode As String = "" 'DECLARES TALL UTILITY 2 UNIT UPPER GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim TU2UUMIGEdgeCode2 As String = "" 'DECLARES TALL UTILITY 2 UNIT UPPER GABLE EDGE SEQUENCE AS A STRING VARIABLE

        '####################################################
        '# TALL UTILITY 2 UNIT UPPER TOP + BOTTOM VARIABLES #
        '####################################################
        '
        Dim TU2UUMITopBtmX As Double 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UUMITopBtmY As Double 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UUMITopBtmZ As Double 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UUMITopBtmQ As Integer 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim TU2UUMITopBtmM As String = "" 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM MATERIAL AS STRING VARIABLE
        Dim TU2UUMITBEdgeSeq As String = "" 'DECLARES TALL UTILITY 2 UNIT UPPER TOP AND BOTTOM EDGE SEQUENCE AS A STRING VARIABLE
        Dim TU2UUMITBEdgeCode As String = "" 'DECLARES TALL UTILITY 2 UNIT UPPER TOP AND BOTTOM EDGE SEQUENCE AS A STRING VARIABLE

        '###########################################
        '# TALL UTILITY 2 UNIT UPPER TOP VARIABLES #
        '###########################################
        '
        Dim TU2UUMITopX As Double 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UUMITopY As Double 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UUMITopZ As Double 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UUMITopQ As Integer 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim TU2UUMITopM As String = "" 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM MATERIAL AS STRING VARIABLE
        Dim TU2UUMITEdgeSeq As String = "" 'DECLARES TALL UTILITY 2 UNIT UPPER TOP EDGE SEQUENCE AS A STRING VARIABLE
        Dim TU2UUMITEdgeCode As String = "" 'DECLARES TALL UTILITY 2 UNIT UPPER TOP EDGE SEQUENCE AS A STRING VARIABLE

        '##############################################
        '# TALL UTILITY 2 UNIT UPPER BOTTOM VARIABLES #
        '##############################################
        '
        Dim TU2UUMIBotX As Double 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UUMIBotY As Double 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UUMIBotZ As Double 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UUMIBotQ As Integer 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim TU2UUMIBotM As String = "" 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM MATERIAL AS STRING VARIABLE
        Dim TU2UUMIBEdgeSeq As String = "" 'DECLARES TALL UTILITY 2 UNIT UPPER BOTTOM EDGE SEQUENCE AS A STRING VARIABLE
        Dim TU2UUMIBEdgeCode As String = "" 'DECLARES TALL UTILITY 2 UNIT UPPER BOTTOM EDGE SEQUENCE AS A STRING VARIABLE

        '########################################################
        '# TALL UTILITY 2 UNIT UPPER ADJUSTABLE SHELF VARIABLES #
        '########################################################
        '
        Dim TU2UUMIAdjShelfX As Double 'DECLARES TALL UTILITY 2 UNIT ADJUSTABLE SHELF X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UUMIAdjShelfY As Double 'DECLARES TALL UTILITY 2 UNIT ADJUSTABLE SHELF Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UUMIAdjShelfZ As Double 'DECLARES TALL UTILITY 2 UNIT ADJUSTABLE SHELF Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UUMIAdjShelfQ As Integer 'DECLARES TALL UTILITY 2 UNIT ADJUSTABLE SHELF QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim TU2UUMIAdjShelfM As String = "" 'DECLARES TALL UTILITY 2 UNIT ADJUSTABLE SHELF MATERIAL AS STRING VARIABLE
        Dim TU2UUMIASEdgeSeq As String = "" 'DECLARES TALL UTILITY 2 UNIT UPPER ADJUSTABLE SHELF EDGE SEQUENCE AS A STRING VARIABLE
        Dim TU2UUMIASEdgeCode As String = "" 'DECLARES TALL UTILITY 2 UNIT UPPER ADJUSTABLE SHELF EDGE SEQUENCE AS A STRING VARIABLE

        '###############################################
        '# TALL UTILITY 2 UNIT UPPER BACKING VARIABLES #
        '###############################################
        '
        Dim TU2UUMIBackX As Double 'DECLARES TALL UTILITY 2 UNIT BACK X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UUMIBackY As Double 'DECLARES TALL UTILITY 2 UNIT BACK Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UUMIBackZ As Double 'DECLARES TALL UTILITY 2 UNIT BACK Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UUMIBackQ As Integer 'DECLARES TALL UTILITY 2 UNIT BACK QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim TU2UUMIBackM As String = "" 'DECLARES TALL UTILITY 2 UNIT BACK MATERIAL AS STRING VARIABLE

        '############################################
        '# TALL UTILITY 2 UNIT BASE GABLE VARIABLES #
        '############################################
        '
        Dim TU2UBGableX As Double 'DECLARES TALL UTILITY 2 UNIT GABLE X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UBGableY As Double 'DECLARES TALL UTILITY 2 UNIT GABLE Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UBGableZ As Double 'DECLARES TALL UTILITY 2 UNIT GABLE Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UBGableQ As Integer 'DECLARES TALL UTILITY 2 UNIT GABLE QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim TU2UBGableM As String = "" 'DECLARES TALL UTILITY 2 UNIT GABLE MATERIAL AS STRING VARIABLE
        Dim TU2UBGEdgeSeq As String = "" 'DECLARES TALL UTILITY 2 UNIT TALL GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim TU2UBGEdgeCode As String = "" 'DECLARES TALL UTILITY 2 UNIT TALL GABLE EDGE SEQUENCE AS A STRING VARIABLE

        '###################################################
        '# TALL UTILITY 2 UNIT BASE TOP + BOTTOM VARIABLES #
        '###################################################
        '
        Dim TU2UBTopBtmX As Double 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UBTopBtmY As Double 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UBTopBtmZ As Double 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UBTopBtmQ As Integer 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim TU2UBTopBtmM As String = "" 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM MATERIAL AS STRING VARIABLE
        Dim TU2UBTBEdgeSeq As String = "" 'DECLARES TALL UTILITY 2 UNIT TALL TOP AND BOTTOM EDGE SEQUENCE AS A STRING VARIABLE
        Dim TU2UBTBEdgeCode As String = "" 'DECLARES TALL UTILITY 2 UNIT TALL TOP AND BOTTOM EDGE SEQUENCE AS A STRING VARIABLE

        '##########################################
        '# TALL UTILITY 2 UNIT BASE TOP VARIABLES #
        '##########################################
        '
        Dim TU2UBTopX As Double 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UBTopY As Double 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UBTopZ As Double 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UBTopQ As Integer 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim TU2UBTopM As String = "" 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM MATERIAL AS STRING VARIABLE
        Dim TU2UBTEdgeSeq As String = "" 'DECLARES TALL UTILITY 2 UNIT TALL TOP EDGE SEQUENCE AS A STRING VARIABLE
        Dim TU2UBTEdgeCode As String = "" 'DECLARES TALL UTILITY 2 UNIT TALL TOP EDGE SEQUENCE AS A STRING VARIABLE

        '#############################################
        '# TALL UTILITY 2 UNIT BASE BOTTOM VARIABLES #
        '#############################################
        '
        Dim TU2UBBotX As Double 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UBBotY As Double 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UBBotZ As Double 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UBBotQ As Integer 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim TU2UBBotM As String = "" 'DECLARES TALL UTILITY 2 UNIT TOP + BOTTOM MATERIAL AS STRING VARIABLE
        Dim TU2UBBEdgeSeq As String = "" 'DECLARES TALL UTILITY 2 UNIT TALL BOTTOM EDGE SEQUENCE AS A STRING VARIABLE
        Dim TU2UBBEdgeCode As String = "" 'DECLARES TALL UTILITY 2 UNIT TALL BOTTOM EDGE SEQUENCE AS A STRING VARIABLE

        '##########################################################
        '# TALL UTILITY 2 UNIT BASE SHALLOW FIXED SHELF VARIABLES #
        '##########################################################
        '
        Dim TU2UBSFSX As Double 'DECLARES TALL UTILITY 2 UNIT SHALLOW FIXED SHELFX MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UBSFSY As Double 'DECLARES TALL UTILITY 2 UNIT SHALLOW FIXED SHELFY MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UBSFSZ As Double 'DECLARES TALL UTILITY 2 UNIT SHALLOW FIXED SHELFZ MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UBSFSQ As Integer 'DECLARES TALL UTILITY 2 UNIT SHALLOW FIXED SHELFQUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim TU2UBSFSM As String = "" 'DECLARES TALL UTILITY 2 UNIT SHALLOW FIXED SHELFMATERIAL AS STRING VARIABLE
        Dim TU2UBSFSEdgeSeq As String = "" 'DECLARES TALL UTILITY 2 UNIT TALL SHALLOW FIXED SHELF EDGE SEQUENCE AS A STRING VARIABLE
        Dim TU2UBSFSEdgeCode As String = "" 'DECLARES TALL UTILITY 2 UNIT TALL SHALLOW FIXED SHELF EDGE SEQUENCE AS A STRING VARIABLE

        '#######################################################
        '# TALL UTILITY 2 UNIT BASE ADJUSTABLE SHELF VARIABLES #
        '#######################################################
        '
        Dim TU2UBAdjShelfX As Double 'DECLARES TALL UTILITY 2 UNIT ADJUSTABLE SHELF X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UBAdjShelfY As Double 'DECLARES TALL UTILITY 2 UNIT ADJUSTABLE SHELF Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UBAdjShelfZ As Double 'DECLARES TALL UTILITY 2 UNIT ADJUSTABLE SHELF Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UBAdjShelfQ As Integer 'DECLARES TALL UTILITY 2 UNIT ADJUSTABLE SHELF QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim TU2UBAdjShelfM As String = "" 'DECLARES TALL UTILITY 2 UNIT ADJUSTABLE SHELF MATERIAL AS STRING VARIABLE
        Dim TU2UBASEdgeSeq As String = "" 'DECLARES TALL UTILITY 2 UNIT TALL ADJUSTABLE SHELF EDGE SEQUENCE AS A STRING VARIABLE
        Dim TU2UBASEdgeCode As String = "" 'DECLARES TALL UTILITY 2 UNIT TALL ADJUSTABLE SHELF EDGE SEQUENCE AS A STRING VARIABLE

        '##############################################
        '# TALL UTILITY 2 UNIT BASE BACKING VARIABLES #
        '##############################################
        '
        Dim TU2UBBackX As Double 'DECLARES TALL UTILITY 2 UNIT BACK X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UBBackY As Double 'DECLARES TALL UTILITY 2 UNIT BACK Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UBBackZ As Double 'DECLARES TALL UTILITY 2 UNIT BACK Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim TU2UBBackQ As Integer 'DECLARES TALL UTILITY 2 UNIT BACK QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim TU2UBBackM As String = "" 'DECLARES TALL UTILITY 2 UNIT BACK MATERIAL AS STRING VARIABLE

        '#######################################################
        '# TALL UTILITY 2 UNIT UPPER GABLE RULES AND EQUATIONS #
        '#######################################################
        '
        Dim TUCase = CutlistForm.TUCabBox.Text
        Select Case TUCase
            Case "TU"
                If (VarHeightI = 232) Then TU2UUGableX = 600 - 1
                If (VarHeightI = 242) Then TU2UUGableX = 760 - 1
                TU2UUGableY = VarDepthI * 10
                TU2UUGableZ = 16
                TU2UUGableQ = 2 * VarAmountI
                TU2UUGableM = ""
            Case "TURS"
                If (VarHeightI = 232) Then TU2UUGableX = (VarHeightI * 10) - 883 - 1
                If (VarHeightI = 242) Then TU2UUGableX = (VarHeightI * 10) - 883 - 1
                TU2UUGableY = VarDepthI * 10
                TU2UUGableZ = 16
                TU2UUGableQ = 2 * VarAmountI
                TU2UUGableM = ""
        End Select

        '###############
        '# GROUP RULES #
        '###############
        '
        If (CKWOPLANNERSGroupBox = "GROUP1") Then
            TU2UUTopX = ((VarWidthI * 10) - (TU2UUGableZ * 2)) - 1
            TU2UUTopY = VarDepthI * 10
            TU2UUTopZ = 16
            TU2UUTopQ = 1 * VarAmountI
            TU2UUTopM = ""
            TU2UUBotX = ((VarWidthI * 10) - (TU2UUGableZ * 2)) - 1
            TU2UUBotY = VarDepthI * 10
            TU2UUBotZ = 16
            TU2UUBotQ = 1 * VarAmountI
            TU2UUBotM = ""
        Else
            TU2UUTopBtmX = ((VarWidthI * 10) - (TU2UUGableZ * 2)) - 1
            TU2UUTopBtmY = VarDepthI * 10
            TU2UUTopBtmZ = 16
            TU2UUTopBtmQ = 2 * VarAmountI
            TU2UUTopBtmM = ""
        End If

        '##################################################################
        '# TALL UTILITY 2 UNIT UPPER ADJUSTABLE SHELF RULES AND EQUATIONS #
        '##################################################################
        '
        TU2UUAdjShelfX = ((VarWidthI * 10) - (TU1UGableZ * 2)) - 5
        '########################################################
        '# GET SIZE OF SHELVES ACCORDING TO UPPER CABINET DEPTH #
        '########################################################
        '
        If VarDepthI <= 48 Then TU2UUAdjShelfY = (VarDepthI * 10) - 30
        If VarDepthI > 48 Then TU2UUAdjShelfY = 450
        If VarDepthI > 62 Then TU2UUAdjShelfY = 530
        TU2UUAdjShelfZ = 16

        '##########################################################
        '# GET QUANITY OF SHELVES ACCORDING TO UPPER GABLE HEIGHT #
        '##########################################################
        '
        If (TU2UUGableX <= 497) Then TU2UUAdjShelfQ = 0
        If (TU2UUGableX >= 498) Then TU2UUAdjShelfQ = 1 * VarAmountI
        If (TU2UUGableX >= 699) Then TU2UUAdjShelfQ = 2 * VarAmountI
        If (TU2UUGableX >= 898) Then TU2UUAdjShelfQ = 3 * VarAmountI
        If (TU2UUGableX >= 1440) Then TU2UUAdjShelfQ = 4 * VarAmountI
        TU2UUAdjShelfM = ""
        CutlistForm.PubASQuantity = TU2UUAdjShelfQ

        '######################################################
        '# TALL UTILITY 2 UNIT UPPER BACK RULES AND EQUATIONS #
        '######################################################
        '
        If (VarHeightI = 232) Then TU2UUBackX = TU2UUGableX - 22
        If (VarHeightI = 242) Then TU2UUBackX = TU2UUGableX - 22
        TU2UUBackY = (VarWidthI * 10) - 23
        TU2UUBackZ = 3
        TU2UUBackQ = 1 * VarAmountI
        TU2UUBackM = ""

        '#######################################################
        '# TALL UTILITY 2 UNIT UPPER GABLE RULES AND EQUATIONS #
        '#######################################################
        '
        Dim TUCaseMI = CutlistForm.TUCabBox.Text
        Select Case TUCaseMI
            Case "TU"
                If (VarHeightI = 232) Then TU2UUMIGableX = 600 - 1
                If (VarHeightI = 242) Then TU2UUMIGableX = 760 - 1
                TU2UUMIGableY = VarDepthI * 10
                TU2UUMIGableZ = 16
                TU2UUMIGableQ = 2 * VarAmountI
                TU2UUMIGableM = ""
            Case "TURS"
                If (VarHeightI = 232) Then TU2UUMIGableX = (VarHeightI * 10) - 883 - 1
                If (VarHeightI = 242) Then TU2UUMIGableX = (VarHeightI * 10) - 883 - 1
                TU2UUMIGableY = VarDepthI * 10
                TU2UUMIGableZ = 16
                TU2UUMIGableQ = 2 * VarAmountI
                TU2UUMIGableM = ""
        End Select

        '###############
        '# GROUP RULES #
        '###############
        '
        If (CKWOPLANNERSGroupBox = "GROUP1") Then
            TU2UUMITopX = ((VarWidthI * 10) - (TU2UUMIGableZ * 2))
            TU2UUMITopY = VarDepthI * 10
            TU2UUMITopZ = 16
            TU2UUMITopQ = 1 * VarAmountI
            TU2UUMITopM = ""
            TU2UUMIBotX = ((VarWidthI * 10) - (TU2UUMIGableZ * 2))
            TU2UUMIBotY = VarDepthI * 10
            TU2UUMIBotZ = 16
            TU2UUMIBotQ = 1 * VarAmountI
            TU2UUMIBotM = ""
        Else
            TU2UUMITopBtmX = ((VarWidthI * 10) - (TU2UUMIGableZ * 2))
            TU2UUMITopBtmY = VarDepthI * 10
            TU2UUMITopBtmZ = 16
            TU2UUMITopBtmQ = 2 * VarAmountI
            TU2UUMITopBtmM = ""
        End If

        '##################################################################
        '# TALL UTILITY 2 UNIT UPPER ADJUSTABLE SHELF RULES AND EQUATIONS #
        '##################################################################
        '
        TU2UUMIAdjShelfX = ((VarWidthI * 10) - (TU1UGableZ * 2)) - 3
        '########################################################
        '# GET SIZE OF SHELVES ACCORDING TO UPPER CABINET DEPTH #
        '########################################################
        '
        If VarDepthI <= 48 Then TU2UUMIAdjShelfY = (VarDepthI * 10) - 30
        If VarDepthI > 48 Then TU2UUMIAdjShelfY = 450
        If VarDepthI > 62 Then TU2UUMIAdjShelfY = 530
        TU2UUMIAdjShelfZ = 16

        '##########################################################
        '# GET QUANITY OF SHELVES ACCORDING TO UPPER GABLE HEIGHT #
        '##########################################################
        '
        If (TU2UUMIGableX <= 497) Then TU2UUMIAdjShelfQ = 0
        If (TU2UUMIGableX >= 498) Then TU2UUMIAdjShelfQ = 1 * VarAmountI
        If (TU2UUMIGableX >= 699) Then TU2UUMIAdjShelfQ = 2 * VarAmountI
        If (TU2UUMIGableX >= 898) Then TU2UUMIAdjShelfQ = 3 * VarAmountI
        If (TU2UUMIGableX >= 1440) Then TU2UUMIAdjShelfQ = 4 * VarAmountI
        TU2UUMIAdjShelfM = ""
        CutlistForm.PubASQuantity = TU2UUMIAdjShelfQ

        '######################################################
        '# TALL UTILITY 2 UNIT UPPER BACK RULES AND EQUATIONS #
        '######################################################
        '
        If (VarHeightI = 232) Then TU2UUMIBackX = TU2UUMIGableX - 22
        If (VarHeightI = 242) Then TU2UUMIBackX = TU2UUMIGableX - 22
        TU2UUMIBackY = (VarWidthI * 10) - 23
        TU2UUMIBackZ = 16
        TU2UUMIBackQ = 1 * VarAmountI
        TU2UUMIBackM = ""

        '######################################################
        '# TALL UTILITY 2 UNIT TALL PARTS RULES AND EQUATIONS #
        '######################################################
        '
        '######################################################
        '# TALL UTILITY 2 UNIT TALL GABLE RULES AND EQUATIONS #
        '######################################################
        '
        Dim TUCase2 = CutlistForm.TUCabBox.Text
        Select Case TUCase2
            Case "TU"
                If (VarHeightI = 232) Then TU2UBGableX = (VarHeightI * 10) - 600
                If (VarHeightI = 242) Then TU2UBGableX = (VarHeightI * 10) - 760
                TU2UBGableY = VarDepthI * 10
                TU2UBGableZ = 16
                TU2UBGableQ = 2 * VarAmountI
                TU2UBGableM = ""
            Case "TURS"
                TU2UBGableX = 883
                TU2UBGableY = VarDepthI * 10
                TU2UBGableZ = 16
                TU2UBGableQ = 2 * VarAmountI
                TU2UBGableM = ""
        End Select

        '#############################################################
        '# TALL UTILITY 2 UNIT TALL TOP + BOTTOM RULES AND EQUATIONS #
        '#############################################################
        '
        Dim TUCase3 = CutlistForm.TUCabBox.Text
        Select Case TUCase3
            Case "TU"
                If (CKWOPLANNERSGroupBox = "GROUP1") Then
                    TU2UBTopX = ((VarWidthI * 10) - (TU2UBGableZ * 2)) - 1
                    TU2UBTopY = VarDepthI * 10
                    TU2UBTopZ = 16
                    TU2UBTopQ = 1 * VarAmountI
                    TU2UBTopM = ""
                    TU2UBBotX = ((VarWidthI * 10) - (TU2UBGableZ * 2)) - 1
                    TU2UBBotY = VarDepthI * 10
                    TU2UBBotZ = 16
                    TU2UBBotQ = 1 * VarAmountI
                    TU2UBBotM = ""
                Else
                    TU2UBTopBtmX = ((VarWidthI * 10) - (TU2UBGableZ * 2)) - 1
                    TU2UBTopBtmY = VarDepthI * 10
                    TU2UBTopBtmZ = 16
                    TU2UBTopBtmQ = 2 * VarAmountI
                    TU2UBTopBtmM = ""
                End If
            Case "TURS"
                TU2UBTopX = ((VarWidthI * 10) - (TU2UBGableZ * 2)) - 1
                TU2UBTopY = 115
                TU2UBTopZ = 16
                TU2UBTopQ = 1 * VarAmountI
                TU2UBTopM = ""
                TU2UBBotX = ((VarWidthI * 10) - (TU2UBGableZ * 2)) - 1
                TU2UBBotY = VarDepthI * 10
                TU2UBBotZ = 16
                TU2UBBotQ = 1 * VarAmountI
                TU2UBBotM = ""
        End Select

        '#################################################################
        '# TALL UTILITY 2 UNIT TALL ADJUSTABLE SHELF RULES AND EQUATIONS #
        '#################################################################
        '
        TU2UBAdjShelfX = ((VarWidthI * 10) - (TU2UBGableZ * 2)) - 5

        '#######################################################
        '# GET SIZE OF SHELVES ACCORDING TO BASE CABINET DEPTH #
        '#######################################################
        '
        If VarDepthI <= 48 Then TU2UBAdjShelfY = (VarDepthI * 10) - 30
        If VarDepthI > 48 Then TU2UBAdjShelfY = 450
        If VarDepthI > 62 Then TU2UBAdjShelfY = 530
        TU2UBAdjShelfZ = 16

        '#########################################################
        '# GET QUANITY OF SHELVES ACCORDING TO BASE GABLE HEIGHT #
        '#########################################################
        '
        Dim TUCase4 = CutlistForm.TUCabBox.Text
        Select Case TUCase4
            Case "TU"
                TU2UBAdjShelfQ = 3 * VarAmountI
            Case "TURS"
                TU2UBAdjShelfQ = 1 * VarAmountI
        End Select
        TU2UBAdjShelfM = ""
        CutlistForm.PubASQuantity = TU2UUAdjShelfQ

        '####################################################################
        '# TALL UTILITY 2 UNIT BASE SHALLOW FIXED SHELF RULES AND EQUATIONS #
        '####################################################################
        '
        TU2UBSFSX = ((VarWidthI * 10) - (TU2UBGableZ * 2)) - 1

        '#######################################################
        '# GET SIZE OF SHELVES ACCORDING TO BASE CABINET DEPTH #
        '#######################################################
        '
        If VarDepthI <= 48 Then TU2UBSFSY = (VarDepthI * 10) - 30
        If VarDepthI > 48 Then TU2UBSFSY = 450
        If VarDepthI > 62 Then TU2UBSFSY = 530
        TU2UBSFSZ = 16
        TU2UBSFSQ = 1 * VarAmountI
        TU2UBSFSM = ""

        '#####################################################
        '# TALL UTILITY 2 UNIT TALL BACK RULES AND EQUATIONS #
        '#####################################################
        '
        TU2UBBackX = (TU2UBGableX - 120) - 23
        If (VarDepthI * 10 <= 30) Then
            TU2UBBackY = (VarWidthI * 10) - 23
        Else
            TU2UBBackY = (VarWidthI * 10) - 26
        End If
        TU2UBBackZ = 3
        TU2UBBackQ = 1 * VarAmountI
        TU2UBBackM = ""

        '###################################
        '# UPPER WINE RACK GABLE VARIABLES #
        '###################################
        '
        Dim UWRGableX As Double 'DECLARES UPPER WINE RACK GABLE X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UWRGableY As Double 'DECLARES UPPER WINE RACK GABLE Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UWRGableZ As Double 'DECLARES UPPER WINE RACK GABLE Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UWRGableQ As Integer 'DECLARES UPPER WINE RACK GABLE QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim UWRGableM As String = "" 'DECLARES UPPER WINE RACK GABLE MATERIAL AS STRING VARIABLE
        Dim UWRGEdgeSeq As String = ""
        Dim UWRGEdgeCode As String = ""
        Dim UWRGEdgeSeq2 As String = ""
        Dim UWRGEdgeCode2 As String = ""

        '##########################################
        '# UPPER WINE RACK TOP + BOTTOM VARIABLES #
        '##########################################
        '
        Dim UWRTopBtmX As Double 'DECLARES UPPER WINE RACK TOP + BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UWRTopBtmY As Double 'DECLARES UPPER WINE RACK TOP + BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UWRTopBtmZ As Double 'DECLARES UPPER WINE RACK TOP + BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UWRTopBtmQ As Integer 'DECLARES UPPER WINE RACK TOP + BOTTOM QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim UWRTopBtmM As String = "" 'DECLARES UPPER WINE RACK TOP + BOTTOM MATERIAL AS STRING VARIABLE
        Dim UWRTBEdgeSeq As String = ""
        Dim UWRTBEdgeCode As String = ""

        '##################################
        '# UPPER WINE RACK BACK VARIABLES #
        '##################################
        '
        Dim UWRBackX As Double 'DECLARES UPPER WINE RACK BACK X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UWRBackY As Double 'DECLARES UPPER WINE RACK BACK Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UWRBackZ As Double 'DECLARES UPPER WINE RACK BACK Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UWRBackQ As Integer 'DECLARES UPPER WINE RACK BACK QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim UWRBackM As String = "" 'DECLARES UPPER WINE RACK BACK MATERIAL AS STRING VARIABLE

        '##############################################
        '# UPPER WINE RACK FULL FIXED SHELF VARIABLES #
        '##############################################
        '
        Dim UWRFFSHX As Double 'DECLARES UPPER WINE RACK FULL FIXED SHELF X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UWRFFSHY As Double 'DECLARES UPPER WINE RACK FULL FIXED SHELF Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UWRFFSHZ As Double 'DECLARES UPPER WINE RACK FULL FIXED SHELF Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UWRFFSHQ As Integer 'DECLARES UPPER WINE RACK FULL FIXED SHELF QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim UWRFFSHM As String = "" 'DECLARES UPPER WINE RACK FULL FIXED SHELF MATERIAL AS STRING VARIABLE
        Dim UWRFFSEdgeSeq As String = ""
        Dim UWRFFSEdgeCode As String = ""

        '#########################################
        '# UPPER WINE RACK FIXED SHELF VARIABLES #
        '#########################################
        '
        Dim UWRFShelfX As Double 'DECLARES UPPER WINE RACK FULL FIXED SHELF X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UWRFShelfY As Double 'DECLARES UPPER WINE RACK FULL FIXED SHELF Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UWRFShelfZ As Double 'DECLARES UPPER WINE RACK FULL FIXED SHELF Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UWRFShelfQ As Integer 'DECLARES UPPER WINE RACK FULL FIXED SHELF QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim UWRFShelfM As String = "" 'DECLARES UPPER WINE RACK FULL FIXED SHELF MATERIAL AS STRING VARIABLE
        Dim UWRFSEdgeSeq As String = ""
        Dim UWRFSEdgeCode As String = ""

        '#####################################
        '# UPPER WINE RACK DIVIDER VARIABLES #
        '#####################################
        '
        Dim UWRDivX As Double 'DECLARES UPPER WINE RACK FULL FIXED SHELF X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UWRDivY As Double 'DECLARES UPPER WINE RACK FULL FIXED SHELF Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UWRDivZ As Double 'DECLARES UPPER WINE RACK FULL FIXED SHELF Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim UWRDivQ As Integer 'DECLARES UPPER WINE RACK FULL FIXED SHELF QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim UWRDivM As String = "" 'DECLARES UPPER WINE RACK FULL FIXED SHELF MATERIAL AS STRING VARIABLE
        Dim UWRDEdgeSeq As String = ""
        Dim UWRDEdgeCode As String = ""
        Dim UWRDEdgeSeq2 As String = ""
        Dim UWRDEdgeCode2 As String = ""

        '##################################
        '# BASE WINE RACK GABLE VARIABLES #
        '##################################
        '
        Dim BWRGableX As Double 'DECLARES BASE WINE RACK GABLE X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BWRGableY As Double 'DECLARES BASE WINE RACK GABLE Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BWRGableZ As Double 'DECLARES BASE WINE RACK GABLE Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BWRGableQ As Integer 'DECLARES BASE WINE RACK GABLE QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim BWRGableM As String = "" 'DECLARES BASE WINE RACK GABLE MATERIAL AS STRING VARIABLE
        Dim BWRGEdgeSeq As String = ""
        Dim BWRGEdgeCode As String = ""

        '#########################################
        '# BASE WINE RACK TOP + BOTTOM VARIABLES #
        '#########################################
        '
        Dim BWRTopBtmX As Double 'DECLARES BASE WINE RACK TOP + BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BWRTopBtmY As Double 'DECLARES BASE WINE RACK TOP + BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BWRTopBtmZ As Double 'DECLARES BASE WINE RACK TOP + BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BWRTopBtmQ As Integer 'DECLARES BASE WINE RACK TOP + BOTTOM QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim BWRTopBtmM As String = "" 'DECLARES BASE WINE RACK TOP + BOTTOM MATERIAL AS STRING VARIABLE
        Dim BWRTBEdgeSeq As String = ""
        Dim BWRTBEdgeCode As String = ""

        '#################################
        '# BASE WINE RACK BACK VARIABLES #
        '#################################
        '
        Dim BWRBackX As Double 'DECLARES BASE WINE RACK BACK X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BWRBackY As Double 'DECLARES BASE WINE RACK BACK Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BWRBackZ As Double 'DECLARES BASE WINE RACK BACK Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BWRBackQ As Integer 'DECLARES BASE WINE RACK BACK QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim BWRBackM As String = "" 'DECLARES BASE WINE RACK BACK MATERIAL AS STRING VARIABLE

        '#############################################
        '# BASE WINE RACK FULL FIXED SHELF VARIABLES #
        '#############################################
        '
        Dim BWRFFSHX As Double 'DECLARES BASE WINE RACK FULL FIXED SHELF X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BWRFFSHY As Double 'DECLARES BASE WINE RACK FULL FIXED SHELF Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BWRFFSHZ As Double 'DECLARES BASE WINE RACK FULL FIXED SHELF Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BWRFFSHQ As Integer 'DECLARES BASE WINE RACK FULL FIXED SHELF QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim BWRFFSHM As String = "" 'DECLARES BASE WINE RACK FULL FIXED SHELF MATERIAL AS STRING VARIABLE
        Dim BWRFFSEdgeSeq As String = ""
        Dim BWRFFSEdgeCode As String = ""

        '####################################
        '# BASE WINE RACK VALANCE VARIABLES #
        '####################################
        '
        Dim BWRValanceX As Double 'DECLARES BASE WINE RACK VALANCE X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BWRValanceY As Double 'DECLARES BASE WINE RACK VALANCE Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BWRValanceZ As Double 'DECLARES BASE WINE RACK VALANCE Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim BWRValanceQ As Integer 'DECLARES BASE WINE RACK VALANCE QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim BWRValanceM As String = "" 'DECLARES BASE WINE RACK VALANCE MATERIAL AS STRING VARIABLE
        Dim BWRVN As String = "" 'DECLARES BASE WINE RACK VALANCE NOTES AS STRING VARIABLE


        '################################################
        '# UPPER AND BASE WINE RACK RULES AND EQUATIONS #
        '################################################
        '
        'UPPER WINE RACK
        UWRBackX = (VarHeightI * 10) - 23
        UWRBackY = (VarWidthI * 10) - 23
        UWRBackZ = 16
        UWRBackQ = 1 * VarAmountI
        UWRBackM = ""

        UWRGableX = (VarHeightI * 10) - 1
        UWRGableY = (VarDepthI * 10)
        UWRGableZ = 16
        UWRGableQ = 2 * VarAmountI
        UWRGableM = ""

        UWRTopBtmX = (VarWidthI * 10) - (UWRGableZ * 2)
        UWRTopBtmY = (VarDepthI * 10)
        UWRTopBtmZ = 16
        UWRTopBtmQ = 2 * VarAmountI
        UWRTopBtmM = ""

        UWRFFSHX = (VarWidthI * 10) - (UWRGableZ * 2)
        UWRFFSHY = (VarDepthI * 10) - 22
        UWRFFSHZ = 16
        UWRFFSHQ = 1 * VarAmountI
        UWRFFSHM = ""

        UWRFShelfX = (VarWidthI * 10) - (UWRGableZ * 2)
        UWRFShelfY = (VarDepthI * 10) - 22
        UWRFShelfZ = 16
        UWRFShelfQ = 6 * VarAmountI
        UWRFShelfM = ""

        UWRDivX = (VarWidthI * 10) - (UWRGableZ * 2)
        UWRDivY = (VarDepthI * 10) - 22
        UWRDivZ = 16
        UWRDivQ = 14 * VarAmountI
        UWRDivM = ""

        'BASE WINE RACK
        BWRBackX = (VarHeightI * 10) - 23 - 120
        BWRBackY = (VarWidthI * 10) - 23
        BWRBackZ = 16
        BWRBackQ = 1 * VarAmountI
        BWRBackM = ""

        BWRTopBtmX = (VarWidthI * 10) - (UWRGableZ * 2)
        BWRTopBtmY = (VarDepthI * 10)
        BWRTopBtmZ = 16
        BWRTopBtmQ = 2 * VarAmountI
        BWRTopBtmM = ""

        BWRGableX = (VarHeightI * 10)
        BWRGableY = (VarDepthI * 10)
        BWRGableZ = 16
        BWRGableQ = 2 * VarAmountI
        BWRGableM = ""

        BWRFFSHX = (VarWidthI * 10) - (UWRGableZ * 2)
        BWRFFSHY = (VarDepthI * 10) - 22
        BWRFFSHZ = 16
        BWRFFSHQ = 1 * VarAmountI
        BWRFFSHM = ""

        BWRValanceX = (VarWidthI * 10)
        BWRValanceY = 100
        BWRValanceZ = 19
        BWRValanceQ = 1 * VarAmountI
        BWRValanceM = "" 'SOLID WOOD MATERIALS! NEW FUNCTIONS REQUIRED.

        '##################################
        '# UNIVERSAL OVEN PANEL VARIABLES #
        '##################################
        '
        Dim OV1X As Double 'DECLARES OVEN PANEL BACK X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim OV1Y As Double 'DECLARES OVEN PANEL BACK Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim OV1Z As Double 'DECLARES OVEN PANEL BACK Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim OV1Q As Integer 'DECLARES OVEN PANEL QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim OV1M As String = "" 'DECLARES OVEN PANEL MATERIAL AS STRING VARIABLE
        Dim OVPLEdgeSeq As String = "" 'DECLARES OVEN PANEL EDGE SEQUENCE AS STRING VARIABLE
        Dim OVPLEdgeCode As String = "" 'DECLARES OVEN PANEL EDGE SEQUENCE AS STRING VARIABLE

        '##################################
        '# OVEN PANEL RULES AND EQUATIONS #
        '##################################
        '
        Dim OV1N As String = ""
        If (CKWOPLANNERSGroupBox = "GROUP1") Then
            Dim SCase = CKWOPLANNERSSpeciesBox
            Select Case SCase
                Case "MAPLE", "OAK", "CHERRY", "PINE", "WALNUT"
                    OV1Y = (VarWidthI) - 1
                    OV1X = (VarHeightI) - 1
                    OV1Z = VarDepthI
                    OV1Q = 1 * VarAmountI
                    OV1M = ""
                    OV1N = ""
                Case "MDF"
                    OV1Y = VarWidthI
                    OV1X = VarHeightI
                    OV1Z = VarDepthI
                    OV1Q = 1 * VarAmountI
                    OV1M = ""
                    Dim DCase = CKWOPLANNERSDoorStyleBox
                    Select Case DCase
                        Case "NOTINGHAM", "SONOMA", "SIENA"
                            OV1N = "Ordered Joe G."
                        Case Else
                            OV1N = ""
                    End Select
                Case "PVC"
                    OV1Y = VarWidthI
                    OV1X = VarHeightI
                    OV1Z = VarDepthI
                    OV1Q = 1 * VarAmountI
                    OV1M = ""
                    OV1N = "Ordered Joe G."
            End Select
        Else
            OV1Y = VarWidthI
            OV1X = VarHeightI
            OV1Z = VarDepthI
            OV1Q = 1 * VarAmountI
            OV1M = ""
            OV1N = ""
            If (CKWOPLANNERSSpeciesBox = "PVC") Then
                OV1N = "Ordered Joe G."
            End If
        End If

        '##############################
        '# HUTCH UNIT GABLE VARIABLES #
        '##############################
        '
        Dim HGableX As Double 'DECLARES HUTCH UNIT GABLE X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HGableY As Double 'DECLARES HUTCH UNIT GABLE Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HGableZ As Double 'DECLARES HUTCH UNIT GABLE Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HGableQ As Integer 'DECLARES HUTCH UNIT GABLE QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim HGableM As String = "" 'DECLARES HUTCH UNIT GABLE MATERIAL AS STRING VARIABLE
        Dim HGEdgeSeq As String = "" 'DECLARES HUTCH UNIT GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim HGEdgeSeq2 As String = "" 'DECLARES HUTCH UNIT GABLE EDGE CODE AS A STRING VARIABLE
        Dim HGEdgeCode As String = "" 'DECLARES HUTCH UNIT GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim HGEdgeCode2 As String = "" 'DECLARES HUTCH UNIT GABLE EDGE CODE AS A STRING VARIABLE

        '###############################################
        '# HUTCH UNIT TOP + FULL FIXED SHELF VARIABLES #
        '###############################################
        '
        Dim HTopFFShelfX As Double 'DECLARES HUTCH UNIT FULL FIXED SHELF X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HTopFFShelfY As Double 'DECLARES HUTCH UNIT FULL FIXED SHELF Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HTopFFShelfZ As Double 'DECLARES HUTCH UNIT FULL FIXED SHELF Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HTopFFShelfQ As Integer 'DECLARES HUTCH UNIT FULL FIXED SHELF QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim HTopFFShelfM As String = "" 'DECLARES HUTCH UNIT FULL FIXED SHELF MATERIAL AS STRING VARIABLE
        Dim HTFFSEdgeSeq As String = "" 'DECLARES HUTCH UNIT TOP FULL FIXED SHELF EDGE SEQUENCE AS A STRING VARIABLE
        Dim HTFFSEdgeCode As String = "" 'DECLARES HUTCH UNIT TOP FULL FIXED SHELF EDGE CODE AS A STRING VARIABLE

        '##################################################
        '# HUTCH UNIT BOTTOM + FULL FIXED SHELF VARIABLES #
        '##################################################
        '
        Dim HBtmFFShelfX As Double 'DECLARES HUTCH UNIT BOTTOM + FULL FIXED SHELF X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HBtmFFShelfY As Double 'DECLARES HUTCH UNIT BOTTOM + FULL FIXED SHELF Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HBtmFFShelfZ As Double 'DECLARES HUTCH UNIT BOTTOM + FULL FIXED SHELF Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HBtmFFShelfQ As Integer 'DECLARES HUTCH UNIT BOTTOM + FULL FIXED SHELF QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim HBtmFFShelfM As String = "" 'DECLARES HUTCH UNIT BOTTOM + FULL FIXED SHELF MATERIAL AS STRING VARIABLE
        Dim HBFFSEdgeSeq As String = "" 'DECLARES HUTCH UNIT BOTTOM FULL FIXED SHELF EDGE SEQUENCE AS A STRING VARIABLE
        Dim HBFFSEdgeCode As String = "" 'DECLARES HUTCH UNIT BOTTOM FULL FIXED SHELF EDGE CODE AS A STRING VARIABLE

        '###############################
        '# HUTCH UNIT BOTTOM VARIABLES #
        '###############################
        '
        Dim HTopX As Double 'DECLARES HUTCH UNIT BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HTopY As Double 'DECLARES HUTCH UNIT BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HTopZ As Double 'DECLARES HUTCH UNIT BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HTopQ As Integer 'DECLARES HUTCH UNIT BOTTOM QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim HTopM As String = "" 'DECLARES HUTCH UNIT BOTTOM MATERIAL AS STRING VARIABLE
        Dim HTEdgeSeq As String = "" 'DECLARES HUTCH UNIT TOP EDGE SEQUENCE AS A STRING VARIABLE
        Dim HTEdgeCode As String = "" 'DECLARES HUTCH UNIT TOP EDGE CODE AS A STRING VARIABLE

        '###############################
        '# HUTCH UNIT BOTTOM VARIABLES #
        '###############################
        '
        Dim HBtmX As Double 'DECLARES HUTCH UNIT BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HBtmY As Double 'DECLARES HUTCH UNIT BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HBtmZ As Double 'DECLARES HUTCH UNIT BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HBtmQ As Integer 'DECLARES HUTCH UNIT BOTTOM QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim HBtmM As String = "" 'DECLARES HUTCH UNIT BOTTOM MATERIAL AS STRING VARIABLE
        Dim HBEdgeSeq As String = "" 'DECLARES HUTCH UNIT BOTTOM EDGE SEQUENCE AS A STRING VARIABLE
        Dim HBEdgeCode As String = "" 'DECLARES HUTCH UNIT BOTTOM EDGE CODE AS A STRING VARIABLE

        '################################
        '# HUTCH UNIT DIVIDER VARIABLES #
        '################################
        '
        Dim HDividerX As Double 'DECLARES HUTCH UNIT DIVIDER X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HDividerY As Double 'DECLARES HUTCH UNIT DIVIDER Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HDividerZ As Double 'DECLARES HUTCH UNIT DIVIDER Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HDividerQ As Integer 'DECLARES HUTCH UNIT DIVIDER QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim HDividerM As String = "" 'DECLARES HUTCH UNIT DIVIDER MATERIAL AS STRING VARIABLE
        Dim HDEdgeSeq As String = "" 'DECLARES HUTCH UNIT DIVIDER EDGE SEQUENCE AS A STRING VARIABLE
        Dim HDEdgeCode As String = "" 'DECLARES HUTCH UNIT DIVIDER EDGE CODE AS A STRING VARIABLE

        '##############################
        '# HUTCH UNIT STRAP VARIABLES #
        '##############################
        '
        Dim HStrapX As Double 'DECLARES HUTCH UNIT STRAP X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HStrapY As Double 'DECLARES HUTCH UNIT STRAP Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HStrapZ As Double 'DECLARES HUTCH UNIT STRAP Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HStrapQ As Integer 'DECLARES HUTCH UNIT STRAP QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim HStrapM As String = "" 'DECLARES HUTCH UNIT STRAP MATERIAL AS STRING VARIABLE
        Dim HSEdgeSeq As String = "" 'DECLARES HUTCH UNIT STRAP EDGE SEQUENCE AS A STRING VARIABLE
        Dim HSEdgeCode As String = "" 'DECLARES HUTCH UNIT STRAP EDGE CODE AS A STRING VARIABLE

        '#########################################
        '# HUTCH UNIT ADJUSTABLE SHELF VARIABLES #
        '#########################################
        '
        Dim HAdjShelfX As Double 'DECLARES HUTCH UNIT ADJUSTABLE SHELF X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HAdjShelfY As Double 'DECLARES HUTCH UNIT ADJUSTABLE SHELF Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HAdjShelfZ As Double 'DECLARES HUTCH UNIT ADJUSTABLE SHELF Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HAdjShelfQ As Integer 'DECLARES HUTCH UNIT ADJUSTABLE SHELF QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim HAdjShelfM As String = "" 'DECLARES HUTCH UNIT ADJUSTABLE SHELF MATERIAL AS STRING VARIABLE
        Dim HASEdgeSeq As String = "" 'DECLARES HUTCH UNIT ADJUSTABLE SHELF EDGE SEQUENCE AS A STRING VARIABLE
        Dim HASEdgeCode As String = "" 'DECLARES HUTCH UNIT ADJUSTABLE SHELF EDGE CODE AS A STRING VARIABLE

        '################################
        '# HUTCH UNIT BACKING VARIABLES #
        '################################
        '
        Dim HBackX As Double 'DECLARES HUTCH UNIT BACK X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HBackY As Double 'DECLARES HUTCH UNIT BACK Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HBackZ As Double 'DECLARES HUTCH UNIT BACK Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HBackQ As Integer 'DECLARES HUTCH UNIT BACK QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim HBackM As String = "" 'DECLARES HUTCH UNIT BACK MATERIAL AS STRING VARIABLE

        '###################################
        '# HUTCH GABLE RULES AND EQUATIONS #
        '###################################
        '
        HGableX = (VarHeightI * 10) - 1
        HGableY = VarDepthI * 10
        HGableZ = 16
        HGableQ = 2 * VarAmountI
        HGableM = ""

        '####################################################
        '# HUTCH TOP + FULL FIXED SHELF RULES AND EQUATIONS #
        '####################################################
        '
        HTopFFShelfX = ((VarWidthI * 10) - (HGableZ * 2)) - 1
        HTopFFShelfY = VarDepthI * 10
        HTopFFShelfZ = 16
        HTopFFShelfQ = 1 * VarAmountI
        HTopFFShelfM = ""

        '#######################################################
        '# HUTCH BOTTOM + FULL FIXED SHELF RULES AND EQUATIONS #
        '#######################################################
        '
        HBtmFFShelfX = ((VarWidthI * 10) - (HGableZ * 2)) - 1
        HBtmFFShelfY = VarDepthI * 10
        HBtmFFShelfZ = 16
        HBtmFFShelfQ = 1 * VarAmountI
        HBtmFFShelfM = ""

        '#################################
        '# HUTCH TOP RULES AND EQUATIONS #
        '#################################
        '
        HTopX = ((VarWidthI * 10) - (HGableZ * 2)) - 1
        HTopY = VarDepthI * 10
        HTopZ = 16
        HTopQ = 1 * VarAmountI
        HTopM = ""

        '####################################
        '# HUTCH BOTTOM RULES AND EQUATIONS #
        '####################################
        '
        HBtmX = ((VarWidthI * 10) - (HGableZ * 2)) - 1
        HBtmY = VarDepthI * 10
        HBtmZ = 16
        HBtmQ = 1 * VarAmountI
        HBtmM = ""

        '###################################
        '# HUTCH STRAP RULES AND EQUATIONS #
        '###################################
        '
        HStrapX = ((VarWidthI * 10) - (HGableZ * 2)) - 1
        HStrapY = 60
        HStrapZ = 16
        HStrapQ = 1 * VarAmountI
        HStrapM = ""

        '#####################################
        '# HUTCH DIVIDER RULES AND EQUATIONS #
        '#####################################
        '
        If (CutlistForm.CabCodeBox1.Text = "HUTCH DRAWER" Or CutlistForm.CabCodeBox1.Text = "HUTCH DOUBLE DRAWER") Then
            HDividerX = (160 - (HBtmFFShelfZ * 2))
            HDividerY = VarDepthI * 10
            HDividerZ = 16
            HDividerQ = 2 * VarAmountI
            HDividerM = ""
            DCHeight = 160
        End If

        If (CutlistForm.CabCodeBox1.Text = "HUTCH DRAWER STACK" Or CutlistForm.CabCodeBox1.Text = "HUTCH DOUBLE DRAWER STACK") Then
            HDividerX = (320 - (HBtmFFShelfZ * 2)) 'DOUBLE CHECK WITH DAD
            HDividerY = VarDepthI * 10
            HDividerZ = 16
            HDividerQ = 2 * VarAmountI
            HDividerM = ""
            DCHeight = 320
        End If

        '##############################################
        '# HUTCH ADJUSTABLE SHELF RULES AND EQUATIONS #
        '##############################################
        '
        HAdjShelfX = ((VarWidthI * 10) - (HGableZ * 2)) - 5

        '##################################################
        '# GET SIZE OF SHELVES ACCORDING TO CABINET DEPTH #
        '##################################################
        '
        If VarDepthI <= 48 Then HAdjShelfY = (VarDepthI * 10) - 30
        If VarDepthI > 48 Then HAdjShelfY = 450
        If VarDepthI > 62 Then HAdjShelfY = 530

        HAdjShelfZ = 16
        HAdjShelfM = ""

        '####################################################
        '# GET QUANITY OF SHELVES ACCORDING TO GABLE HEIGHT #
        '####################################################
        '
        If ((HGableX - DCHeight) <= 497) Then HAdjShelfQ = 0
        If ((HGableX - DCHeight) >= 498) Then HAdjShelfQ = 1 * VarAmountI
        If ((HGableX - DCHeight) >= 699) Then HAdjShelfQ = 2 * VarAmountI
        If ((HGableX - DCHeight) >= 898) Then HAdjShelfQ = 3 * VarAmountI
        If ((HGableX - DCHeight) >= 1440) Then HAdjShelfQ = 4 * VarAmountI
        CutlistForm.PubASQuantity = HAdjShelfQ

        '#####################################
        '# HUTCH BACKING RULES AND EQUATIONS #
        '#####################################
        '
        HBackY = (VarWidthI * 10) - 23
        HBackX = ((VarHeightI * 10) - DCHeight) - 23
        HBackZ = 3
        HBackQ = 1 * VarAmountI
        HBackM = ""

        '################################################
        '# HUTCH MATCHING INTERIOR UNIT GABLE VARIABLES #
        '################################################
        '
        Dim HMIGableX As Double 'DECLARES HMI UNIT GABLE X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HMIGableY As Double 'DECLARES HMI UNIT GABLE Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HMIGableZ As Double 'DECLARES HMI UNIT GABLE Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HMIGableQ As Integer 'DECLARES HMI UNIT GABLE QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim HMIGableM As String = "" 'DECLARES HMI UNIT GABLE MATERIAL AS STRING VARIABLE
        Dim HMIGEdgeSeq As String = "" 'DECLARES HMI UNIT GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim HMIGEdgeSeq2 As String = "" 'DECLARES HMI UNIT GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim HMIGEdgeCode As String = "" 'DECLARES HMI UNIT GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim HMIGEdgeCode2 As String = "" 'DECLARES HMI UNIT GABLE EDGE SEQUENCE AS A STRING VARIABLE

        '#################################################################
        '# HUTCH MATCHING INTERIOR UNIT TOP + FULL FIXED SHELF VARIABLES #
        '#################################################################
        '
        Dim HMITopFFShelfX As Double 'DECLARES HMI UNIT FULL FIXED SHELF X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HMITopFFShelfY As Double 'DECLARES HMI UNIT FULL FIXED SHELF Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HMITopFFShelfZ As Double 'DECLARES HMI UNIT FULL FIXED SHELF Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HMITopFFShelfQ As Integer 'DECLARES HMI UNIT FULL FIXED SHELF QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim HMITopFFShelfM As String = "" 'DECLARES HMI UNIT FULL FIXED SHELF MATERIAL AS STRING VARIABLE
        Dim HMITFFSEdgeSeq As String = "" 'DECLARES HMI UNIT TOP FULL FIXED SHELF EDGE SEQUENCE AS A STRING VARIABLE
        Dim HMITFFSEdgeCode As String = "" 'DECLARES HMI UNIT TOP FULL FIXED SHELF EDGE SEQUENCE AS A STRING VARIABLE

        '####################################################################
        '# HUTCH MATCHING INTERIOR UNIT BOTTOM + FULL FIXED SHELF VARIABLES #
        '####################################################################
        '
        Dim HMIBtmFFShelfX As Double 'DECLARES HMI UNIT BOTTOM + FULL FIXED SHELF X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HMIBtmFFShelfY As Double 'DECLARES HMI UNIT BOTTOM + FULL FIXED SHELF Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HMIBtmFFShelfZ As Double 'DECLARES HMI UNIT BOTTOM + FULL FIXED SHELF Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HMIBtmFFShelfQ As Integer 'DECLARES HMI UNIT BOTTOM + FULL FIXED SHELF QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim HMIBtmFFShelfM As String = "" 'DECLARES HMI UNIT BOTTOM + FULL FIXED SHELF MATERIAL AS STRING VARIABLE
        Dim HMIBFFSEdgeSeq As String = "" 'DECLARES HMI UNIT BOTTOM FULL FIXED SHELF EDGE SEQUENCE AS A STRING VARIABLE
        Dim HMIBFFSEdgeCode As String = "" 'DECLARES HMI UNIT BOTTOM FULL FIXED SHELF EDGE SEQUENCE AS A STRING VARIABLE

        '##############################################
        '# HUTCH MATCHING INTERIOR UNIT TOP VARIABLES #
        '##############################################
        '
        Dim HMITopX As Double 'DECLARES HMI UNIT TOP X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HMITopY As Double 'DECLARES HMI UNIT TOP Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HMITopZ As Double 'DECLARES HMI UNIT TOP Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HMITopQ As Integer 'DECLARES HMI UNIT TOP QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim HMITopM As String = "" 'DECLARES HMI UNIT TOP MATERIAL AS STRING VARIABLE
        Dim HMITEdgeSeq As String = "" 'DECLARES HMI UNIT TOP EDGE SEQUENCE AS A STRING VARIABLE
        Dim HMITEdgeCode As String = "" 'DECLARES HMI UNIT TOP EDGE SEQUENCE AS A STRING VARIABLE

        '#################################################
        '# HUTCH MATCHING INTERIOR UNIT BOTTOM VARIABLES #
        '#################################################
        '
        Dim HMIBtmX As Double 'DECLARES HMI UNIT BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HMIBtmY As Double 'DECLARES HMI UNIT BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HMIBtmZ As Double 'DECLARES HMI UNIT BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HMIBtmQ As Integer 'DECLARES HMI UNIT BOTTOM QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim HMIBtmM As String = "" 'DECLARES HMI UNIT BOTTOM MATERIAL AS STRING VARIABLE
        Dim HMIBEdgeSeq As String = "" 'DECLARES HMI UNIT BOTTOM EDGE SEQUENCE AS A STRING VARIABLE
        Dim HMIBEdgeCode As String = "" 'DECLARES HMI UNIT BOTTOM EDGE SEQUENCE AS A STRING VARIABLE

        '##################################################
        '# HUTCH MATCHING INTERIOR UNIT DIVIDER VARIABLES #
        '##################################################
        '
        Dim HMIDividerX As Double 'DECLARES HMI UNIT DIVIDER X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HMIDividerY As Double 'DECLARES HMI UNIT DIVIDER Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HMIDividerZ As Double 'DECLARES HMI UNIT DIVIDER Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HMIDividerQ As Integer 'DECLARES HMI UNIT DIVIDER QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim HMIDividerM As String = "" 'DECLARES HMI UNIT DIVIDER MATERIAL AS STRING VARIABLE
        Dim HMIDEdgeSeq As String = "" 'DECLARES HMI UNIT DIVIDER EDGE SEQUENCE AS A STRING VARIABLE
        Dim HMIDEdgeCode As String = "" 'DECLARES HMI UNIT DIVIDER EDGE SEQUENCE AS A STRING VARIABLE

        '################################################
        '# HUTCH MATCHING INTERIOR UNIT STRAP VARIABLES #
        '################################################
        '
        Dim HMIStrapX As Double 'DECLARES HMI UNIT STRAP X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HMIStrapY As Double 'DECLARES HMI UNIT STRAP Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HMIStrapZ As Double 'DECLARES HMI UNIT STRAP Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HMIStrapQ As Integer 'DECLARES HMI UNIT STRAP QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim HMIStrapM As String = "" 'DECLARES HMI UNIT STRAP MATERIAL AS STRING VARIABLE
        Dim HMISEdgeSeq As String = "" 'DECLARES HMI UNIT STRAP EDGE SEQUENCE AS A STRING VARIABLE
        Dim HMISEdgeCode As String = "" 'DECLARES HMI UNIT STRAP EDGE SEQUENCE AS A STRING VARIABLE

        '###########################################################
        '# HUTCH MATCHING INTERIOR UNIT ADJUSTABLE SHELF VARIABLES #
        '###########################################################
        '
        Dim HMIAdjShelfX As Double 'DECLARES HMI UNIT ADJUSTABLE SHELF X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HMIAdjShelfY As Double 'DECLARES HMI UNIT ADJUSTABLE SHELF Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HMIAdjShelfZ As Double 'DECLARES HMI UNIT ADJUSTABLE SHELF Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HMIAdjShelfQ As Integer 'DECLARES HMI UNIT ADJUSTABLE SHELF QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim HMIAdjShelfM As String = "" 'DECLARES HMI UNIT ADJUSTABLE SHELF MATERIAL AS STRING VARIABLE
        Dim HMIASEdgeSeq As String = "" 'DECLARES HMI UNIT ADJUSTABLE SHELF EDGE SEQUENCE AS A STRING VARIABLE
        Dim HMIASEdgeCode As String = "" 'DECLARES HMI UNIT ADJUSTABLE SHELF EDGE SEQUENCE AS A STRING VARIABLE

        '##################################################
        '# HUTCH MATCHING INTERIOR UNIT BACKING VARIABLES #
        '##################################################
        '
        Dim HMIBackX As Double 'DECLARES HMI UNIT BACK X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HMIBackY As Double 'DECLARES HMI UNIT BACK Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HMIBackZ As Double 'DECLARES HMI UNIT BACK Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim HMIBackQ As Integer 'DECLARES HMI UNIT BACK QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim HMIBackM As String = "" 'DECLARES HMI UNIT BACK MATERIAL AS STRING VARIABLE

        '#####################################################
        '# HUTCH MATCHING INTERIOR GABLE RULES AND EQUATIONS #
        '#####################################################
        '
        HMIGableX = (VarHeightI * 10) - 1
        HMIGableY = VarDepthI * 10
        HMIGableZ = 16
        HMIGableQ = 2 * VarAmountI
        HMIGableM = ""

        '######################################################################
        '# HUTCH MATCHING INTERIOR TOP + FULL FIXED SHELF RULES AND EQUATIONS #
        '######################################################################
        '
        HMITopFFShelfX = ((VarWidthI * 10) - (HMIGableZ * 2))
        HMITopFFShelfY = VarDepthI * 10
        HMITopFFShelfZ = 16
        If (CutlistForm.CabCodeBox1.Text = "HUTCH DOUBLE DRAWER") Then
            HMITopFFShelfQ = 2 * VarAmountI
        Else
            HMITopFFShelfQ = 1 * VarAmountI
        End If

        HMITopFFShelfM = ""

        '#########################################################################
        '# HUTCH MATCHING INTERIOR BOTTOM + FULL FIXED SHELF RULES AND EQUATIONS #
        '#########################################################################
        '
        HMIBtmFFShelfX = ((VarWidthI * 10) - (HMIGableZ * 2))
        HMIBtmFFShelfY = VarDepthI * 10
        HMIBtmFFShelfZ = 16
        If (CutlistForm.CabCodeBox1.Text = "HUTCH DOUBLE DRAWER") Then
            HMIBtmFFShelfQ = 2 * VarAmountI
        Else
            HMIBtmFFShelfQ = 1 * VarAmountI
        End If

        HMIBtmFFShelfM = ""

        '###################################################
        '# HUTCH MATCHING INTERIOR TOP RULES AND EQUATIONS #
        '###################################################
        '
        HMITopX = ((VarWidthI * 10) - (HMIGableZ * 2))
        HMITopY = VarDepthI * 10
        HMITopZ = 16
        HMITopQ = 1 * VarAmountI
        HMITopM = ""

        '######################################################
        '# HUTCH MATCHING INTERIOR BOTTOM RULES AND EQUATIONS #
        '######################################################
        '
        HMIBtmX = ((VarWidthI * 10) - (HMIGableZ * 2))
        HMIBtmY = VarDepthI * 10
        HMIBtmZ = 16
        HMIBtmQ = 1 * VarAmountI
        HMIBtmM = ""

        '#####################################################
        '# HUTCH MATCHING INTERIOR STRAP RULES AND EQUATIONS #
        '#####################################################
        '
        HMIStrapX = ((VarWidthI * 10) - (HMIGableZ * 2))
        HMIStrapY = 60
        HMIStrapZ = 16
        HMIStrapQ = 1 * VarAmountI
        HMIStrapM = ""

        '#######################################################
        '# HUTCH MATCHING INTERIOR DIVIDER RULES AND EQUATIONS #
        '#######################################################
        '
        If (CutlistForm.CabCodeBox1.Text = "HUTCH DRAWER" Or CutlistForm.CabCodeBox1.Text = "HUTCH DOUBLE DRAWER") Then
            DCHeight = 160
            HMIDividerX = (DCHeight - (HMIBtmFFShelfZ * 2))
            HMIDividerY = VarDepthI * 10
            HMIDividerZ = 16
            HMIDividerQ = 2 * VarAmountI
            HMIDividerM = ""
            CutlistForm.PubDCHeight = DCHeight
        End If

        If (CutlistForm.CabCodeBox1.Text = "HUTCH DRAWER STACK" Or CutlistForm.CabCodeBox1.Text = "HUTCH DOUBLE DRAWER STACK") Then
            DCHeight = 320
            HMIDividerX = (DCHeight - (HMIBtmFFShelfZ * 2)) 'DOUBLE CHECK WITH DAD
            HMIDividerY = VarDepthI * 10
            HMIDividerZ = 16
            HMIDividerQ = 2 * VarAmountI
            HMIDividerM = ""
            CutlistForm.PubDCHeight = DCHeight
        End If

        '################################################################
        '# HUTCH MATCHING INTERIOR ADJUSTABLE SHELF RULES AND EQUATIONS #
        '################################################################
        '
        HMIAdjShelfX = ((VarWidthI * 10) - (HMIGableZ * 2)) - 3

        '##################################################
        '# GET SIZE OF SHELVES ACCORDING TO CABINET DEPTH #
        '##################################################
        '
        If VarDepthI <= 48 Then HMIAdjShelfY = (VarDepthI * 10) - 30
        If VarDepthI > 48 Then HMIAdjShelfY = 450
        If VarDepthI > 62 Then HMIAdjShelfY = 530

        HMIAdjShelfZ = 16
        HMIAdjShelfM = ""

        '####################################################
        '# GET QUANITY OF SHELVES ACCORDING TO GABLE HEIGHT #
        '####################################################
        '
        If ((HMIGableX - DCHeight) <= 497) Then HMIAdjShelfQ = 0
        If ((HMIGableX - DCHeight) >= 498) Then HMIAdjShelfQ = 1 * VarAmountI
        If ((HMIGableX - DCHeight) >= 699) Then HMIAdjShelfQ = 2 * VarAmountI
        If ((HMIGableX - DCHeight) >= 898) Then HMIAdjShelfQ = 3 * VarAmountI
        If ((HMIGableX - DCHeight) >= 1440) Then HMIAdjShelfQ = 4 * VarAmountI
        CutlistForm.PubASQuantity = HMIAdjShelfQ

        '#######################################################
        '# HUTCH MATCHING INTERIOR BACKING RULES AND EQUATIONS #
        '#######################################################
        '
        HMIBackY = (VarWidthI * 10) - 23
        HMIBackX = ((VarHeightI * 10) - DCHeight) - 23
        HMIBackZ = 16
        HMIBackQ = 1 * VarAmountI
        HMIBackM = ""

        '##########################
        '# VANITY GABLE VARIABLES #
        '##########################
        '
        Dim VGableX As Double 'DECLARES VANITY GABLE X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VGableY As Double 'DECLARES VANITY GABLE Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VGableZ As Double 'DECLARES VANITY GABLE Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VGableQ As Integer 'DECLARES VANITY GABLE QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim VGableM As String = "" 'DECLARES VANITY GABLE MATERIAL AS STRING VARIABLE
        Dim VGEdgeSeq As String = "" 'DECLARES VANITY GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim VGEdgeCode As String = "" 'DECLARES VANITY GABLE EDGE SEQUENCE AS A STRING VARIABLE

        '##############################
        '# VANITY TOP STRAP VARIABLES #
        '##############################
        '
        Dim VTopStrapX As Double 'DECLARES VANITY BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VTopStrapY As Double 'DECLARES VANITY BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VTopStrapZ As Double 'DECLARES VANITY BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VTopStrapQ As Integer 'DECLARES VANITY BOTTOM QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim VTopStrapM As String = "" 'DECLARES VANITY BOTTOM MATERIAL AS STRING VARIABLE
        Dim VTSEdgeSeq As String = "" 'DECLARES VANITY TOP STRAP EDGE SEQUENCE AS A STRING VARIABLE
        Dim VTSEdgeCode As String = "" 'DECLARES VANITY TOP STRAP EDGE SEQUENCE AS A STRING VARIABLE

        '############################
        '# VANITY DIVIDER VARIABLES #
        '############################
        '
        Dim VDividerX As Double 'DECLARES VANITY BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VDividerY As Double 'DECLARES VANITY BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VDividerZ As Double 'DECLARES VANITY BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VDividerQ As Integer 'DECLARES VANITY BOTTOM QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim VDividerM As String = "" 'DECLARES VANITY BOTTOM MATERIAL AS STRING VARIABLE
        Dim VDEdgeSeq As String = "" 'DECLARES VANITY DIVIDER EDGE SEQUENCE AS A STRING VARIABLE
        Dim VDEdgeCode As String = "" 'DECLARES VANITY DIVIDER EDGE SEQUENCE AS A STRING VARIABLE

        '##########################
        '# VANITY STRAP VARIABLES #
        '##########################
        '
        Dim VStrapX As Double 'DECLARES VANITY STRAP X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VStrapY As Double 'DECLARES VANITY STRAP Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VStrapZ As Double 'DECLARES VANITY STRAP Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VStrapQ As Integer 'DECLARES VANITY STRAP QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim VStrapM As String = "" 'DECLARES VANITY STRAP MATERIAL AS STRING VARIABLE
        Dim VSEdgeSeq As String = "" 'DECLARES VANITY STRAP EDGE SEQUENCE AS A STRING VARIABLE
        Dim VSEdgeCode As String = "" 'DECLARES VANITY STRAP EDGE SEQUENCE AS A STRING VARIABLE

        '#####################################
        '# VANITY FULL FIXED SHELF VARIABLES #
        '#####################################
        '
        Dim VFFShelfX As Double 'DECLARES VANITY FULL FIXED SHELF X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VFFShelfY As Double 'DECLARES VANITY FULL FIXED SHELF Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VFFShelfZ As Double 'DECLARES VANITY FULL FIXED SHELF Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VFFShelfQ As Integer 'DECLARES VANITY FULL FIXED SHELF QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim VFFShelfM As String = "" 'DECLARES VANITY FULL FIXED SHELF MATERIAL AS STRING VARIABLE
        Dim VFFSHEdgeSeq As String = "" 'DECLARES VANITY FULL FIXED SHELF EDGE SEQUENCE AS A STRING VARIABLE
        Dim VFFSHEdgeCode As String = "" 'DECLARES VANITY FULL FIXED SHELF EDGE SEQUENCE AS A STRING VARIABLE

        '########################
        '# VANITY TOP VARIABLES #
        '########################
        '
        Dim VTopX As Double 'DECLARES VANITY TOP X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VTopY As Double 'DECLARES VANITY TOP Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VTopZ As Double 'DECLARES VANITY TOP Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VTopQ As Integer 'DECLARES VANITY TOP QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim VTopM As String = "" 'DECLARES VANITY TOP MATERIAL AS STRING VARIABLE
        Dim VTEdgeSeq As String = "" 'DECLARES VANITY TOP EDGE SEQUENCE AS A STRING VARIABLE
        Dim VTEdgeCode As String = "" 'DECLARES VANITY TOP EDGE SEQUENCE AS A STRING VARIABLE

        '###########################
        '# VANITY BOTTOM VARIABLES #
        '###########################
        '
        Dim VBtmX As Double 'DECLARES VANITY BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VBtmY As Double 'DECLARES VANITY BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VBtmZ As Double 'DECLARES VANITY BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VBtmQ As Integer 'DECLARES VANITY BOTTOM QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim VBtmM As String = "" 'DECLARES VANITY BOTTOM MATERIAL AS STRING VARIABLE
        Dim VBEdgeSeq As String = "" 'DECLARES VANITY BOTTOM EDGE SEQUENCE AS A STRING VARIABLE
        Dim VBEdgeCode As String = "" 'DECLARES VANITY BOTTOM EDGE SEQUENCE AS A STRING VARIABLE

        '############################
        '# VANITY BACKING VARIABLES #
        '############################
        '
        Dim VBackX As Double 'DECLARES VANITY BACK X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VBackY As Double 'DECLARES VANITY BACK Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VBackZ As Double 'DECLARES VANITY BACK Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VBackQ As Integer 'DECLARES VANITY BACK QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim VBackM As String = "" 'DECLARES VANITY BACK MATERIAL AS STRING VARIABLE

        '####################################
        '# VANITY GABLE RULES AND EQUATIONS #
        '####################################
        '
        VGableX = (VarHeightI * 10)
        VGableY = VarDepthI * 10
        VGableZ = 16
        VGableQ = 2 * VarAmountI
        VGableM = ""

        '########################################
        '# VANITY TOP STRAP RULES AND EQUATIONS #
        '########################################
        '
        VTopStrapX = ((VarWidthI * 10) - (VGableZ * 2)) - 1
        VTopStrapY = 115
        VTopStrapZ = 16
        VTopStrapQ = 1 * VarAmountI
        VTopStrapM = ""

        '########################################
        '# VANITY TOP STRAP RULES AND EQUATIONS #
        '########################################
        '
        VDividerX = 139
        VDividerY = (VarDepthI * 10) - 23
        VDividerZ = 16
        VDividerQ = 2 * VarAmountI
        VDividerM = ""

        '########################################
        '# VANITY TOP STRAP RULES AND EQUATIONS #
        '########################################
        '
        VStrapX = ((VarWidthI * 10) - (VGableZ * 2)) - 1
        VStrapY = 60
        VStrapZ = 16
        VStrapQ = 1 * VarAmountI
        VStrapM = ""

        '##################################
        '# VANITY TOP RULES AND EQUATIONS #
        '##################################
        '
        VTopX = ((VarWidthI * 10) - (VGableZ * 2)) - 1
        VTopY = VarDepthI * 10
        VTopZ = 16
        VTopQ = 1 * VarAmountI
        VTopM = ""

        '#####################################
        '# VANITY BOTTOM RULES AND EQUATIONS #
        '#####################################
        '
        VBtmX = ((VarWidthI * 10) - (VGableZ * 2)) - 1
        VBtmY = VarDepthI * 10
        VBtmZ = 16
        VBtmQ = 1 * VarAmountI
        VBtmM = ""

        '###############################################
        '# VANITY FULL FIXED SHELF RULES AND EQUATIONS #
        '###############################################
        '
        VFFShelfX = ((VarWidthI * 10) - (VGableZ * 2)) - 1
        VFFShelfY = (VarDepthI * 10) - 22
        VFFShelfZ = 16
        VFFShelfQ = 1 * VarAmountI
        VFFShelfM = ""

        '######################################
        '# VANITY BACKING RULES AND EQUATIONS #
        '######################################
        '
        VBackY = (VarWidthI * 10) - 23
        VBackX = ((VarHeightI * 10) - 120) - 23
        VBackZ = 3
        VBackQ = 1 * VarAmountI
        VBackM = ""

        '############################################
        '# VANITY MATCHING INTERIOR GABLE VARIABLES #
        '############################################
        '
        Dim VMIGableX As Double 'DECLARES VANITY MATCHING INTERIOR GABLE X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VMIGableY As Double 'DECLARES VANITY MATCHING INTERIOR GABLE Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VMIGableZ As Double 'DECLARES VANITY MATCHING INTERIOR GABLE Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VMIGableQ As Integer 'DECLARES VANITY MATCHING INTERIOR GABLE QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim VMIGableM As String = "" 'DECLARES VANITY MATCHING INTERIOR GABLE MATERIAL AS STRING VARIABLE
        Dim VMIGEdgeSeq As String = "" 'DECLARES VANITY MATCHING INTERIOR GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim VMIGEdgeCode As String = "" 'DECLARES VANITY MATCHING INTERIOR GABLE EDGE SEQUENCE AS A STRING VARIABLE

        '################################################
        '# VANITY MATCHING INTERIOR TOP STRAP VARIABLES #
        '################################################
        '
        Dim VMITopStrapX As Double 'DECLARES VANITY MATCHING INTERIOR BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VMITopStrapY As Double 'DECLARES VANITY MATCHING INTERIOR BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VMITopStrapZ As Double 'DECLARES VANITY MATCHING INTERIOR BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VMITopStrapQ As Integer 'DECLARES VANITY MATCHING INTERIOR BOTTOM QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim VMITopStrapM As String = "" 'DECLARES VANITY MATCHING INTERIOR BOTTOM MATERIAL AS STRING VARIABLE
        Dim VMITSEdgeSeq As String = "" 'DECLARES VANITY MATCHING INTERIOR TOP STRAP EDGE SEQUENCE AS A STRING VARIABLE
        Dim VMITSEdgeCode As String = "" 'DECLARES VANITY MATCHING INTERIOR TOP STRAP EDGE SEQUENCE AS A STRING VARIABLE

        '##############################################
        '# VANITY MATCHING INTERIOR DIVIDER VARIABLES #
        '##############################################
        '
        Dim VMIDividerX As Double 'DECLARES VANITY MATCHING INTERIOR BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VMIDividerY As Double 'DECLARES VANITY MATCHING INTERIOR BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VMIDividerZ As Double 'DECLARES VANITY MATCHING INTERIOR BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VMIDividerQ As Integer 'DECLARES VANITY MATCHING INTERIOR BOTTOM QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim VMIDividerM As String = "" 'DECLARES VANITY MATCHING INTERIOR BOTTOM MATERIAL AS STRING VARIABLE
        Dim VMIDEdgeSeq As String = "" 'DECLARES VANITY MATCHING INTERIOR DIVIDER EDGE SEQUENCE AS A STRING VARIABLE
        Dim VMIDEdgeCode As String = "" 'DECLARES VANITY MATCHING INTERIOR DIVIDER EDGE SEQUENCE AS A STRING VARIABLE

        '############################################
        '# VANITY MATCHING INTERIOR STRAP VARIABLES #
        '############################################
        '
        Dim VMIStrapX As Double 'DECLARES VANITY MATCHING INTERIOR STRAP X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VMIStrapY As Double 'DECLARES VANITY MATCHING INTERIOR STRAP Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VMIStrapZ As Double 'DECLARES VANITY MATCHING INTERIOR STRAP Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VMIStrapQ As Integer 'DECLARES VANITY MATCHING INTERIOR STRAP QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim VMIStrapM As String = "" 'DECLARES VANITY MATCHING INTERIOR STRAP MATERIAL AS STRING VARIABLE
        Dim VMISEdgeSeq As String = "" 'DECLARES VANITY MATCHING INTERIOR STRAP EDGE SEQUENCE AS A STRING VARIABLE
        Dim VMISEdgeCode As String = "" 'DECLARES VANITY MATCHING INTERIOR STRAP EDGE SEQUENCE AS A STRING VARIABLE

        '#######################################################
        '# VANITY MATCHING INTERIOR FULL FIXED SHELF VARIABLES #
        '#######################################################
        '
        Dim VMIFFShelfX As Double 'DECLARES VANITY MATCHING INTERIOR FULL FIXED SHELF X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VMIFFShelfY As Double 'DECLARES VANITY MATCHING INTERIOR FULL FIXED SHELF Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VMIFFShelfZ As Double 'DECLARES VANITY MATCHING INTERIOR FULL FIXED SHELF Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VMIFFShelfQ As Integer 'DECLARES VANITY MATCHING INTERIOR FULL FIXED SHELF QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim VMIFFShelfM As String = "" 'DECLARES VANITY MATCHING INTERIOR FULL FIXED SHELF MATERIAL AS STRING VARIABLE
        Dim VMIFFSHEdgeSeq As String = "" 'DECLARES VANITY MATCHING INTERIOR FULL FIXED SHELF EDGE SEQUENCE AS A STRING VARIABLE
        Dim VMIFFSHEdgeCode As String = "" 'DECLARES VANITY MATCHING INTERIOR FULL FIXED SHELF EDGE SEQUENCE AS A STRING VARIABLE

        '##########################################
        '# VANITY MATCHING INTERIOR TOP VARIABLES #
        '##########################################
        '
        Dim VMITopX As Double 'DECLARES VANITY MATCHING INTERIOR TOP X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VMITopY As Double 'DECLARES VANITY MATCHING INTERIOR TOP Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VMITopZ As Double 'DECLARES VANITY MATCHING INTERIOR TOP Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VMITopQ As Integer 'DECLARES VANITY MATCHING INTERIOR TOP QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim VMITopM As String = "" 'DECLARES VANITY MATCHING INTERIOR TOP MATERIAL AS STRING VARIABLE
        Dim VMITEdgeSeq As String = "" 'DECLARES VANITY MATCHING INTERIOR TOP EDGE SEQUENCE AS A STRING VARIABLE
        Dim VMITEdgeCode As String = "" 'DECLARES VANITY MATCHING INTERIOR TOP EDGE SEQUENCE AS A STRING VARIABLE

        '#############################################
        '# VANITY MATCHING INTERIOR BOTTOM VARIABLES #
        '#############################################
        '
        Dim VMIBtmX As Double 'DECLARES VANITY MATCHING INTERIOR BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VMIBtmY As Double 'DECLARES VANITY MATCHING INTERIOR BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VMIBtmZ As Double 'DECLARES VANITY MATCHING INTERIOR BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VMIBtmQ As Integer 'DECLARES VANITY MATCHING INTERIOR BOTTOM QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim VMIBtmM As String = "" 'DECLARES VANITY MATCHING INTERIOR BOTTOM MATERIAL AS STRING VARIABLE
        Dim VMIBEdgeSeq As String = "" 'DECLARES VANITY MATCHING INTERIOR BOTTOM EDGE SEQUENCE AS A STRING VARIABLE
        Dim VMIBEdgeCode As String = "" 'DECLARES VANITY MATCHING INTERIOR BOTTOM EDGE SEQUENCE AS A STRING VARIABLE

        '##############################################
        '# VANITY MATCHING INTERIOR BACKING VARIABLES #
        '##############################################
        '
        Dim VMIBackX As Double 'DECLARES VANITY MATCHING INTERIOR BACK X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VMIBackY As Double 'DECLARES VANITY MATCHING INTERIOR BACK Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VMIBackZ As Double 'DECLARES VANITY MATCHING INTERIOR BACK Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VMIBackQ As Integer 'DECLARES VANITY MATCHING INTERIOR BACK QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim VMIBackM As String = "" 'DECLARES VANITY MATCHING INTERIOR BACK MATERIAL AS STRING VARIABLE

        '######################################################
        '# VANITY MATCHING INTERIOR GABLE RULES AND EQUATIONS #
        '######################################################
        '
        VMIGableX = (VarHeightI * 10)
        VMIGableY = VarDepthI * 10
        VMIGableZ = 16
        VMIGableQ = 2 * VarAmountI
        VMIGableM = ""

        '##########################################################
        '# VANITY MATCHING INTERIOR TOP STRAP RULES AND EQUATIONS #
        '##########################################################
        '
        VMITopStrapX = ((VarWidthI * 10) - (VMIGableZ * 2)) - 1
        VMITopStrapY = 115
        VMITopStrapZ = 16
        VMITopStrapQ = 1 * VarAmountI
        VMITopStrapM = ""

        '##########################################################
        '# VANITY MATCHING INTERIOR TOP STRAP RULES AND EQUATIONS #
        '##########################################################
        '
        VMIDividerX = 139
        VMIDividerY = (VarDepthI * 10) - 23
        VMIDividerZ = 16
        VMIDividerQ = 2 * VarAmountI
        VMIDividerM = ""

        '##########################################################
        '# VANITY MATCHING INTERIOR TOP STRAP RULES AND EQUATIONS #
        '##########################################################
        '
        VMIStrapX = ((VarWidthI * 10) - (VMIGableZ * 2)) - 1
        VMIStrapY = 60
        VMIStrapZ = 16
        VMIStrapQ = 1 * VarAmountI
        VMIStrapM = ""

        '####################################################
        '# VANITY MATCHING INTERIOR TOP RULES AND EQUATIONS #
        '####################################################
        '
        VMITopX = ((VarWidthI * 10) - (VMIGableZ * 2)) - 1
        VMITopY = VarDepthI * 10
        VMITopZ = 16
        VMITopQ = 1 * VarAmountI
        VMITopM = ""

        '#######################################################
        '# VANITY MATCHING INTERIOR BOTTOM RULES AND EQUATIONS #
        '#######################################################
        '
        VMIBtmX = ((VarWidthI * 10) - (VMIGableZ * 2)) - 1
        VMIBtmY = VarDepthI * 10
        VMIBtmZ = 16
        VMIBtmQ = 1 * VarAmountI
        VMIBtmM = ""

        '#################################################################
        '# VANITY MATCHING INTERIOR FULL FIXED SHELF RULES AND EQUATIONS #
        '#################################################################
        '
        VMIFFShelfX = ((VarWidthI * 10) - (VMIGableZ * 2)) - 1
        VMIFFShelfY = (VarDepthI * 10) - 22
        VMIFFShelfZ = 16
        VMIFFShelfQ = 1 * VarAmountI
        VMIFFShelfM = ""

        '########################################################
        '# VANITY MATCHING INTERIOR BACKING RULES AND EQUATIONS #
        '########################################################
        '
        VMIBackY = (VarWidthI * 10) - 23
        VMIBackX = ((VarHeightI * 10) - 120) - 23
        VMIBackZ = 16
        VMIBackQ = 1 * VarAmountI
        VMIBackM = ""

        '###################################
        '# VANITY ELEVATED GABLE VARIABLES #
        '###################################
        '
        Dim VEGableX As Double 'DECLARES VANITY ELEVATED GABLE X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VEGableY As Double 'DECLARES VANITY ELEVATED GABLE Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VEGableZ As Double 'DECLARES VANITY ELEVATED GABLE Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VEGableQ As Integer 'DECLARES VANITY ELEVATED GABLE QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim VEGableM As String = "" 'DECLARES VANITY ELEVATED GABLE MATERIAL AS STRING VARIABLE
        Dim VEGEdgeSeq As String = "" 'DECLARES VANITY ELEVATED GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim VEGEdgeCode As String = "" 'DECLARES VANITY ELEVATED GABLE EDGE SEQUENCE AS A STRING VARIABLE

        '#######################################
        '# VANITY ELEVATED TOP STRAP VARIABLES #
        '#######################################
        '
        Dim VETopStrapX As Double 'DECLARES VANITY ELEVATED BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VETopStrapY As Double 'DECLARES VANITY ELEVATED BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VETopStrapZ As Double 'DECLARES VANITY ELEVATED BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VETopStrapQ As Integer 'DECLARES VANITY ELEVATED BOTTOM QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim VETopStrapM As String = "" 'DECLARES VANITY ELEVATED BOTTOM MATERIAL AS STRING VARIABLE
        Dim VETSEdgeSeq As String = "" 'DECLARES VANITY ELEVATED TOP STRAP EDGE SEQUENCE AS A STRING VARIABLE
        Dim VETSEdgeCode As String = "" 'DECLARES VANITY ELEVATED TOP STRAP EDGE SEQUENCE AS A STRING VARIABLE

        '#####################################
        '# VANITY ELEVATED DIVIDER VARIABLES #
        '#####################################
        '
        Dim VEDividerX As Double 'DECLARES VANITY ELEVATED BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VEDividerY As Double 'DECLARES VANITY ELEVATED BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VEDividerZ As Double 'DECLARES VANITY ELEVATED BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VEDividerQ As Integer 'DECLARES VANITY ELEVATED BOTTOM QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim VEDividerM As String = "" 'DECLARES VANITY ELEVATED BOTTOM MATERIAL AS STRING VARIABLE
        Dim VEDEdgeSeq As String = "" 'DECLARES VANITY ELEVATED DIVIDER EDGE SEQUENCE AS A STRING VARIABLE
        Dim VEDEdgeCode As String = "" 'DECLARES VANITY ELEVATED DIVIDER EDGE SEQUENCE AS A STRING VARIABLE

        '###################################
        '# VANITY ELEVATED STRAP VARIABLES #
        '###################################
        '
        Dim VEStrapX As Double 'DECLARES VANITY ELEVATED STRAP X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VEStrapY As Double 'DECLARES VANITY ELEVATED STRAP Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VEStrapZ As Double 'DECLARES VANITY ELEVATED STRAP Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VEStrapQ As Integer 'DECLARES VANITY ELEVATED STRAP QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim VEStrapM As String = "" 'DECLARES VANITY ELEVATED STRAP MATERIAL AS STRING VARIABLE
        Dim VESEdgeSeq As String = "" 'DECLARES VANITY ELEVATED STRAP EDGE SEQUENCE AS A STRING VARIABLE
        Dim VESEdgeCode As String = "" 'DECLARES VANITY ELEVATED STRAP EDGE SEQUENCE AS A STRING VARIABLE

        '##############################################
        '# VANITY ELEVATED FULL FIXED SHELF VARIABLES #
        '##############################################
        '
        Dim VEFFShelfX As Double 'DECLARES VANITY ELEVATED FULL FIXED SHELF X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VEFFShelfY As Double 'DECLARES VANITY ELEVATED FULL FIXED SHELF Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VEFFShelfZ As Double 'DECLARES VANITY ELEVATED FULL FIXED SHELF Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VEFFShelfQ As Integer 'DECLARES VANITY ELEVATED FULL FIXED SHELF QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim VEFFShelfM As String = "" 'DECLARES VANITY ELEVATED FULL FIXED SHELF MATERIAL AS STRING VARIABLE
        Dim VEFFSHEdgeSeq As String = "" 'DECLARES VANITY ELEVATED FULL FIXED SHELF EDGE SEQUENCE AS A STRING VARIABLE
        Dim VEFFSHEdgeCode As String = "" 'DECLARES VANITY ELEVATED FULL FIXED SHELF EDGE SEQUENCE AS A STRING VARIABLE

        '#################################
        '# VANITY ELEVATED TOP VARIABLES #
        '#################################
        '
        Dim VETopX As Double 'DECLARES VANITY ELEVATED TOP X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VETopY As Double 'DECLARES VANITY ELEVATED TOP Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VETopZ As Double 'DECLARES VANITY ELEVATED TOP Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VETopQ As Integer 'DECLARES VANITY ELEVATED TOP QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim VETopM As String = "" 'DECLARES VANITY ELEVATED TOP MATERIAL AS STRING VARIABLE
        Dim VETEdgeSeq As String = "" 'DECLARES VANITY ELEVATED TOP EDGE SEQUENCE AS A STRING VARIABLE
        Dim VETEdgeCode As String = "" 'DECLARES VANITY ELEVATED TOP EDGE SEQUENCE AS A STRING VARIABLE

        '####################################
        '# VANITY ELEVATED BOTTOM VARIABLES #
        '####################################
        '
        Dim VEBtmX As Double 'DECLARES VANITY ELEVATED BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VEBtmY As Double 'DECLARES VANITY ELEVATED BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VEBtmZ As Double 'DECLARES VANITY ELEVATED BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VEBtmQ As Integer 'DECLARES VANITY ELEVATED BOTTOM QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim VEBtmM As String = "" 'DECLARES VANITY ELEVATED BOTTOM MATERIAL AS STRING VARIABLE
        Dim VEBEdgeSeq As String = "" 'DECLARES VANITY ELEVATED BOTTOM EDGE SEQUENCE AS A STRING VARIABLE
        Dim VEBEdgeCode As String = "" 'DECLARES VANITY ELEVATED BOTTOM EDGE SEQUENCE AS A STRING VARIABLE

        '#####################################
        '# VANITY ELEVATED BACKING VARIABLES #
        '#####################################
        '
        Dim VEBackX As Double 'DECLARES VANITY ELEVATED BACK X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VEBackY As Double 'DECLARES VANITY ELEVATED BACK Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VEBackZ As Double 'DECLARES VANITY ELEVATED BACK Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VEBackQ As Integer 'DECLARES VANITY ELEVATED BACK QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim VEBackM As String = "" 'DECLARES VANITY ELEVATED BACK MATERIAL AS STRING VARIABLE

        '#############################################
        '# VANITY ELEVATED GABLE RULES AND EQUATIONS #
        '#############################################
        '
        VEGableX = (VarHeightI * 10)
        VEGableY = VarDepthI * 10
        VEGableZ = 16
        VEGableQ = 2 * VarAmountI
        VEGableM = ""

        '#################################################
        '# VANITY ELEVATED TOP STRAP RULES AND EQUATIONS #
        '#################################################
        '
        VETopStrapX = ((VarWidthI * 10) - (VEGableZ * 2)) - 1
        VETopStrapY = 300
        VETopStrapZ = 16
        VETopStrapQ = 1 * VarAmountI
        VETopStrapM = ""

        '#################################################
        '# VANITY ELEVATED TOP STRAP RULES AND EQUATIONS #
        '#################################################
        '
        VEDividerX = 139
        VEDividerY = (VarDepthI * 10) - 23
        VEDividerZ = 16
        VEDividerQ = 2 * VarAmountI
        VEDividerM = ""

        '#################################################
        '# VANITY ELEVATED TOP STRAP RULES AND EQUATIONS #
        '#################################################
        '
        VEStrapX = ((VarWidthI * 10) - (VEGableZ * 2)) - 1
        VEStrapY = 60
        VEStrapZ = 16
        VEStrapQ = 1 * VarAmountI
        VEStrapM = ""

        '###########################################
        '# VANITY ELEVATED TOP RULES AND EQUATIONS #
        '###########################################
        '
        VETopX = ((VarWidthI * 10) - (VEGableZ * 2)) - 1
        VETopY = VarDepthI * 10
        VETopZ = 16
        VETopQ = 1 * VarAmountI
        VETopM = ""

        '##############################################
        '# VANITY ELEVATED BOTTOM RULES AND EQUATIONS #
        '##############################################
        '
        VEBtmX = ((VarWidthI * 10) - (VEGableZ * 2)) - 1
        VEBtmY = VarDepthI * 10
        VEBtmZ = 16
        VEBtmQ = 1 * VarAmountI
        VEBtmM = ""

        '########################################################
        '# VANITY ELEVATED FULL FIXED SHELF RULES AND EQUATIONS #
        '########################################################
        '
        VEFFShelfX = ((VarWidthI * 10) - (VEGableZ * 2)) - 1
        VEFFShelfY = (VarDepthI * 10) - 22
        VEFFShelfZ = 16
        VEFFShelfQ = 1 * VarAmountI
        VEFFShelfM = ""

        '###############################################
        '# VANITY ELEVATED BACKING RULES AND EQUATIONS #
        '###############################################
        '
        VEBackY = (VarWidthI * 10) - 23
        VEBackX = (VarHeightI * 10) - 23
        If (CutlistForm.BackGrooveBox.Text = "3mm") Then
            VEBackZ = 3
        Else
            VEBackZ = 16
        End If
        VEBackQ = 1 * VarAmountI
        VEBackM = ""

        '#####################################################
        '# VANITY ELEVATED MATCHING INTERIOR GABLE VARIABLES #
        '#####################################################
        '
        Dim VEMIGableX As Double 'DECLARES VANITY ELEVATED MATCHING INTERIOR GABLE X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VEMIGableY As Double 'DECLARES VANITY ELEVATED MATCHING INTERIOR GABLE Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VEMIGableZ As Double 'DECLARES VANITY ELEVATED MATCHING INTERIOR GABLE Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VEMIGableQ As Integer 'DECLARES VANITY ELEVATED MATCHING INTERIOR GABLE QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim VEMIGableM As String = "" 'DECLARES VANITY ELEVATED MATCHING INTERIOR GABLE MATERIAL AS STRING VARIABLE
        Dim VEMIGEdgeSeq As String = "" 'DECLARES VANITY ELEVATED MATCHING INTERIOR GABLE EDGE SEQUENCE AS A STRING VARIABLE
        Dim VEMIGEdgeCode As String = "" 'DECLARES VANITY ELEVATED MATCHING INTERIOR GABLE EDGE SEQUENCE AS A STRING VARIABLE

        '#########################################################
        '# VANITY ELEVATED MATCHING INTERIOR TOP STRAP VARIABLES #
        '#########################################################
        '
        Dim VEMITopStrapX As Double 'DECLARES VANITY ELEVATED MATCHING INTERIOR BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VEMITopStrapY As Double 'DECLARES VANITY ELEVATED MATCHING INTERIOR BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VEMITopStrapZ As Double 'DECLARES VANITY ELEVATED MATCHING INTERIOR BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VEMITopStrapQ As Integer 'DECLARES VANITY ELEVATED MATCHING INTERIOR BOTTOM QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim VEMITopStrapM As String = "" 'DECLARES VANITY ELEVATED MATCHING INTERIOR BOTTOM MATERIAL AS STRING VARIABLE
        Dim VEMITSEdgeSeq As String = "" 'DECLARES VANITY ELEVATED MATCHING INTERIOR TOP STRAP EDGE SEQUENCE AS A STRING VARIABLE
        Dim VEMITSEdgeCode As String = "" 'DECLARES VANITY ELEVATED MATCHING INTERIOR TOP STRAP EDGE SEQUENCE AS A STRING VARIABLE

        '#######################################################
        '# VANITY ELEVATED MATCHING INTERIOR DIVIDER VARIABLES #
        '#######################################################
        '
        Dim VEMIDividerX As Double 'DECLARES VANITY ELEVATED MATCHING INTERIOR BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VEMIDividerY As Double 'DECLARES VANITY ELEVATED MATCHING INTERIOR BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VEMIDividerZ As Double 'DECLARES VANITY ELEVATED MATCHING INTERIOR BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VEMIDividerQ As Integer 'DECLARES VANITY ELEVATED MATCHING INTERIOR BOTTOM QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim VEMIDividerM As String = "" 'DECLARES VANITY ELEVATED MATCHING INTERIOR BOTTOM MATERIAL AS STRING VARIABLE
        Dim VEMIDEdgeSeq As String = "" 'DECLARES VANITY ELEVATED MATCHING INTERIOR DIVIDER EDGE SEQUENCE AS A STRING VARIABLE
        Dim VEMIDEdgeCode As String = "" 'DECLARES VANITY ELEVATED MATCHING INTERIOR DIVIDER EDGE SEQUENCE AS A STRING VARIABLE

        '#####################################################
        '# VANITY ELEVATED MATCHING INTERIOR STRAP VARIABLES #
        '#####################################################
        '
        Dim VEMIStrapX As Double 'DECLARES VANITY ELEVATED MATCHING INTERIOR STRAP X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VEMIStrapY As Double 'DECLARES VANITY ELEVATED MATCHING INTERIOR STRAP Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VEMIStrapZ As Double 'DECLARES VANITY ELEVATED MATCHING INTERIOR STRAP Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VEMIStrapQ As Integer 'DECLARES VANITY ELEVATED MATCHING INTERIOR STRAP QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim VEMIStrapM As String = "" 'DECLARES VANITY ELEVATED MATCHING INTERIOR STRAP MATERIAL AS STRING VARIABLE
        Dim VEMISEdgeSeq As String = "" 'DECLARES VANITY ELEVATED MATCHING INTERIOR STRAP EDGE SEQUENCE AS A STRING VARIABLE
        Dim VEMISEdgeCode As String = "" 'DECLARES VANITY ELEVATED MATCHING INTERIOR STRAP EDGE SEQUENCE AS A STRING VARIABLE

        '################################################################
        '# VANITY ELEVATED MATCHING INTERIOR FULL FIXED SHELF VARIABLES #
        '################################################################
        '
        Dim VEMIFFShelfX As Double 'DECLARES VANITY ELEVATED MATCHING INTERIOR FULL FIXED SHELF X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VEMIFFShelfY As Double 'DECLARES VANITY ELEVATED MATCHING INTERIOR FULL FIXED SHELF Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VEMIFFShelfZ As Double 'DECLARES VANITY ELEVATED MATCHING INTERIOR FULL FIXED SHELF Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VEMIFFShelfQ As Integer 'DECLARES VANITY ELEVATED MATCHING INTERIOR FULL FIXED SHELF QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim VEMIFFShelfM As String = "" 'DECLARES VANITY ELEVATED MATCHING INTERIOR FULL FIXED SHELF MATERIAL AS STRING VARIABLE
        Dim VEMIFFSHEdgeSeq As String = "" 'DECLARES VANITY ELEVATED MATCHING INTERIOR FULL FIXED SHELF EDGE SEQUENCE AS A STRING VARIABLE
        Dim VEMIFFSHEdgeCode As String = "" 'DECLARES VANITY ELEVATED MATCHING INTERIOR FULL FIXED SHELF EDGE SEQUENCE AS A STRING VARIABLE

        '###################################################
        '# VANITY ELEVATED MATCHING INTERIOR TOP VARIABLES #
        '###################################################
        '
        Dim VEMITopX As Double 'DECLARES VANITY ELEVATED MATCHING INTERIOR TOP X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VEMITopY As Double 'DECLARES VANITY ELEVATED MATCHING INTERIOR TOP Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VEMITopZ As Double 'DECLARES VANITY ELEVATED MATCHING INTERIOR TOP Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VEMITopQ As Integer 'DECLARES VANITY ELEVATED MATCHING INTERIOR TOP QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim VEMITopM As String = "" 'DECLARES VANITY ELEVATED MATCHING INTERIOR TOP MATERIAL AS STRING VARIABLE
        Dim VEMITEdgeSeq As String = "" 'DECLARES VANITY ELEVATED MATCHING INTERIOR TOP EDGE SEQUENCE AS A STRING VARIABLE
        Dim VEMITEdgeCode As String = "" 'DECLARES VANITY ELEVATED MATCHING INTERIOR TOP EDGE SEQUENCE AS A STRING VARIABLE

        '######################################################
        '# VANITY ELEVATED MATCHING INTERIOR BOTTOM VARIABLES #
        '######################################################
        '
        Dim VEMIBtmX As Double 'DECLARES VANITY ELEVATED MATCHING INTERIOR BOTTOM X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VEMIBtmY As Double 'DECLARES VANITY ELEVATED MATCHING INTERIOR BOTTOM Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VEMIBtmZ As Double 'DECLARES VANITY ELEVATED MATCHING INTERIOR BOTTOM Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VEMIBtmQ As Integer 'DECLARES VANITY ELEVATED MATCHING INTERIOR BOTTOM QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim VEMIBtmM As String = "" 'DECLARES VANITY ELEVATED MATCHING INTERIOR BOTTOM MATERIAL AS STRING VARIABLE
        Dim VEMIBEdgeSeq As String = "" 'DECLARES VANITY ELEVATED MATCHING INTERIOR BOTTOM EDGE SEQUENCE AS A STRING VARIABLE
        Dim VEMIBEdgeCode As String = "" 'DECLARES VANITY ELEVATED MATCHING INTERIOR BOTTOM EDGE SEQUENCE AS A STRING VARIABLE

        '#######################################################
        '# VANITY ELEVATED MATCHING INTERIOR BACKING VARIABLES #
        '#######################################################
        '
        Dim VEMIBackX As Double 'DECLARES VANITY ELEVATED MATCHING INTERIOR BACK X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VEMIBackY As Double 'DECLARES VANITY ELEVATED MATCHING INTERIOR BACK Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VEMIBackZ As Double 'DECLARES VANITY ELEVATED MATCHING INTERIOR BACK Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim VEMIBackQ As Integer 'DECLARES VANITY ELEVATED MATCHING INTERIOR BACK QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim VEMIBackM As String = "" 'DECLARES VANITY ELEVATED MATCHING INTERIOR BACK MATERIAL AS STRING VARIABLE

        '###############################################################
        '# VANITY ELEVATED MATCHING INTERIOR GABLE RULES AND EQUATIONS #
        '###############################################################
        '
        VEMIGableX = (VarHeightI * 10)
        VEMIGableY = VarDepthI * 10
        VEMIGableZ = 16
        VEMIGableQ = 2 * VarAmountI
        VEMIGableM = ""

        '###################################################################
        '# VANITY ELEVATED MATCHING INTERIOR TOP STRAP RULES AND EQUATIONS #
        '###################################################################
        '
        VEMITopStrapX = ((VarWidthI * 10) - (VEMIGableZ * 2))
        VEMITopStrapY = 300
        VEMITopStrapZ = 16
        VEMITopStrapQ = 1 * VarAmountI
        VEMITopStrapM = ""

        '###################################################################
        '# VANITY ELEVATED MATCHING INTERIOR TOP STRAP RULES AND EQUATIONS #
        '###################################################################
        '
        VEMIDividerX = 139
        VEMIDividerY = (VarDepthI * 10) - 23
        VEMIDividerZ = 16
        VEMIDividerQ = 2 * VarAmountI
        VEMIDividerM = ""

        '###################################################################
        '# VANITY ELEVATED MATCHING INTERIOR TOP STRAP RULES AND EQUATIONS #
        '###################################################################
        '
        VEMIStrapX = ((VarWidthI * 10) - (VEMIGableZ * 2))
        VEMIStrapY = 60
        VEMIStrapZ = 16
        VEMIStrapQ = 1 * VarAmountI
        VEMIStrapM = ""

        '#############################################################
        '# VANITY ELEVATED MATCHING INTERIOR TOP RULES AND EQUATIONS #
        '#############################################################
        '
        VEMITopX = ((VarWidthI * 10) - (VEMIGableZ * 2))
        VEMITopY = VarDepthI * 10
        VEMITopZ = 16
        VEMITopQ = 1 * VarAmountI
        VEMITopM = ""

        '################################################################
        '# VANITY ELEVATED MATCHING INTERIOR BOTTOM RULES AND EQUATIONS #
        '################################################################
        '
        VEMIBtmX = ((VarWidthI * 10) - (VEMIGableZ * 2))
        VEMIBtmY = VarDepthI * 10
        VEMIBtmZ = 16
        VEMIBtmQ = 1 * VarAmountI
        VEMIBtmM = ""

        '##########################################################################
        '# VANITY ELEVATED MATCHING INTERIOR FULL FIXED SHELF RULES AND EQUATIONS #
        '##########################################################################
        '
        VEMIFFShelfX = ((VarWidthI * 10) - (VEMIGableZ * 2))
        VEMIFFShelfY = (VarDepthI * 10) - 22
        VEMIFFShelfZ = 16
        VEMIFFShelfQ = 1 * VarAmountI
        VEMIFFShelfM = ""

        '#################################################################
        '# VANITY ELEVATED MATCHING INTERIOR BACKING RULES AND EQUATIONS #
        '#################################################################
        '
        VEMIBackY = (VarWidthI * 10) - 23
        VEMIBackX = (VarHeightI * 10) - 23
        If (CutlistForm.BackGrooveBox.Text = "3mm") Then
            VEMIBackZ = 3
        Else
            VEMIBackZ = 16
        End If
        VEMIBackQ = 1 * VarAmountI
        VEMIBackM = ""

        '#########################
        '# CANOPY BACK VARIABLES #
        '#########################
        '
        Dim RHBackX As Integer
        Dim RHBackX2 As Integer
        Dim RHBackY As Integer
        Dim RHBackZ As Integer
        Dim RHBackQ As Integer
        Dim RHBackM As String = ""

        '########################
        '# CANOPY TOP VARIABLES #
        '########################
        '
        Dim RHTopX As Integer
        Dim RHTopY As Integer
        Dim RHTopZ As Integer
        Dim RHTopQ As Integer
        Dim RHTopM As String = ""
        Dim RHTEdgeSeq As String = ""
        Dim RHTEdgeCode As String = ""

        '##########################
        '# CANOPY GABLE VARIABLES #
        '##########################
        '
        Dim RHGableX As Integer
        Dim RHGableY As Integer
        Dim RHGableZ As Integer
        Dim RHGableQ As Integer
        Dim RHGableM As String = ""
        Dim RHGEdgeSeq As String = ""
        Dim RHGEdgeCode As String = ""
        Dim RHGEdgeSeq2 As String = ""
        Dim RHGEdgeCode2 As String = ""

        '##########################
        '# CANOPY FRONT VARIABLES #
        '##########################
        '
        Dim RHFrontX As Integer
        Dim RHFrontY As Integer
        Dim RHFrontZ As Integer
        Dim RHFrontQ As Integer
        Dim RHFrontM As String = ""
        Dim RHFEdgeSeq As String = ""
        Dim RHFEdgeCode As String = ""

        '##############################
        '# CANOPY FAN SHELF VARIABLES #
        '##############################
        '
        Dim RHFanSHX As Integer
        Dim RHFanSHY As Integer
        Dim RHFanSHZ As Integer
        Dim RHFanSHQ As Integer
        Dim RHFanSHM As String = ""

        '################################
        '# CANOPY SMALL FRONT VARIABLES #
        '################################
        '
        Dim RHSFrontX As Integer
        Dim RHSFrontY As Integer
        Dim RHSFrontZ As Integer
        Dim RHSFrontQ As Integer
        Dim RHSFrontM As String = ""

        '##############################
        '# CANOPY SMALL TOP VARIABLES #
        '##############################
        '
        Dim RHSTopX As Integer
        Dim RHSTopY As Integer
        Dim RHSTopZ As Integer
        Dim RHSTopQ As Integer
        Dim RHSTopM As String = ""
        Dim RHBEdgeSeq As String = ""
        Dim RHBEdgeCode As String = ""

        '#################################
        '# CANOPY INSERT PANEL VARIABLES #
        '#################################
        '
        Dim RHIPanelX As Integer
        Dim RHIPanelY As Integer
        Dim RHIPanelZ As Integer
        Dim RHIPanelQ As Integer
        Dim RHIPanelM As String = ""
        Dim RHIPanel2X As Integer
        Dim RHIPanel2Y As Integer
        Dim RHIPanel2Z As Integer
        Dim RHIPanel2Q As Integer

        '####################################
        '# CANOPY GABLE RULES AND EQUATIONS #
        '####################################
        '
        RHGableX = (VarHeightI * 10) - 1
        RHGableY = VarDepthI * 10
        RHGableZ = 16
        RHGableQ = 2 * VarAmountI
        RHGableM = ""

        '##################################
        '# CANOPY TOP RULES AND EQUATIONS #
        '##################################
        '
        RHTopX = ((VarWidthI * 10) - (RHGableZ * 2)) - 1
        RHTopY = VarDepthI * 10
        RHTopZ = 16
        RHTopQ = 1 * VarAmountI
        RHTopM = ""

        '####################################
        '# CANOPY FRONT RULES AND EQUATIONS #
        '####################################
        '
        Dim RHCase1 = CKWOPLANNERSSpeciesBox
        Select Case RHCase1
            Case "MAPLE", "OAK", "CHERRY", "PINE", "WALNUT"
                RHFrontX = (VarHeightI * 10) - 1
                RHFrontY = (VarWidthI * 10) - 1
                RHFrontZ = 19
                RHFrontQ = 1 * VarAmountI
                RHFrontM = ""
            Case "MDF"
                RHFrontX = VarHeightI * 10
                RHFrontY = VarWidthI * 10
                RHFrontZ = 19
                RHFrontQ = 1 * VarAmountI
                RHFrontM = ""
            Case Else
        End Select

        '########################################
        '# CANOPY FAN SHELF RULES AND EQUATIONS #
        '########################################
        '
        RHFanSHX = VarWidthI * 10
        RHFanSHY = CutlistForm.DepthBox2.Text * 10
        RHFanSHZ = 19
        RHFanSHQ = 1 * VarAmountI
        RHFanSHM = ""

        '##########################################
        '# CANOPY SMALL FRONT RULES AND EQUATIONS #
        '##########################################
        '
        RHSFrontX = VarWidthI * 10
        RHSFrontY = CutlistForm.HeightBox2.Text * 10
        RHSFrontZ = 19
        RHSFrontQ = 1 * VarAmountI
        RHSFrontM = ""

        '########################################
        '# CANOPY SMALL TOP RULES AND EQUATIONS #
        '########################################
        '
        RHSTopX = VarWidthI * 10
        RHSTopY = (CutlistForm.DepthBox2.Text - CutlistForm.DepthBox1.Text) * 10
        RHSTopZ = 19
        RHSTopQ = 1 * VarAmountI
        RHSTopM = ""

        '###########################################
        '# CANOPY INSERT PANEL RULES AND EQUATIONS #
        '###########################################
        '
        Dim RHCase = CutlistForm.CabCodeBox1.Text
        Select Case RHCase
            Case "CANOPY SINGLE"
                RHIPanelX = ((VarHeightI * 10) - RHSFrontY) - 128
                RHIPanelY = (VarWidthI * 10) - 120
                RHIPanelZ = 6
                RHIPanelQ = 1 * VarAmountI
                RHIPanelM = ""
            Case "CANOPY DOUBLE"
                RHIPanelX = ((VarHeightI * 10) - RHSFrontY) - 128
                RHIPanelY = ((VarWidthI * 10) - 180) / 2
                RHIPanelZ = 6
                RHIPanelQ = 2 * VarAmountI
                RHIPanelM = ""
            Case "CANOPY TRIPLE"
                RHIPanelX = ((VarHeightI * 10) - RHSFrontY) - 128
                RHIPanelY = (((VarWidthI * 10) - 240) / 2)
                RHIPanelZ = 6
                RHIPanelQ = 1 * VarAmountI
                RHIPanel2X = ((VarHeightI * 10) - RHSFrontY) - 128
                RHIPanel2Y = (((VarWidthI * 10) - 240) / 2) / 2
                RHIPanel2Z = 6
                RHIPanel2Q = 2 * VarAmountI
                RHIPanelM = ""
            Case Else
        End Select

        '###################################
        '# CANOPY BACK RULES AND EQUATIONS #
        '###################################
        '
        If (VarDepthI < 40) Then
            RHBackX = ((RHGableX - 100) - RHFanSHZ) - 10
            RHBackX2 = ((VarHeightI * 10) - RHTopZ) + 5
            RHBackY = (VarWidthI * 10) - 23
            RHBackZ = 16
            RHBackQ = 1 * VarAmountI
            RHBackM = ""
        Else
            RHBackX = RHGableX - 10
            RHBackY = (VarWidthI * 10) - 23
            RHBackZ = 16
            RHBackQ = 1 * VarAmountI
            RHBackM = ""
        End If

        '############################
        '# FLOATING SHELF VARIABLES #
        '############################
        '
        Dim FShelfX As Double 'DECLARES UPPER GABLE X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim FShelfY As Double 'DECLARES UPPER GABLE Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim FShelfZ As Double 'DECLARES UPPER GABLE Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim FShelfQ As Integer 'DECLARES UPPER GABLE QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim FShelfM As String = "" 'DECLARES UPPER GABLE MATERIAL AS STRING VARIABLE

        '######################################
        '# FLOATING SHELF RULES AND EQUATIONS #
        '######################################
        '
        FShelfX = (VarHeightI * 10) + 10
        FShelfY = (VarDepthI * 10) + 10
        FShelfZ = 16
        FShelfQ = 2 * VarAmountI
        FShelfM = ""

        '##########################
        '# WINDOW PANEL VARIABLES #
        '##########################
        '
        Dim WPanelX As Double 'DECLARES WINDOW PANEL X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim WPanelY As Double 'DECLARES WINDOW PANEL Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim WPanelZ As Double 'DECLARES WINDOW PANEL Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim WPanelQ As Integer 'DECLARES WINDOW PANEL QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim WPanelM As String = "" 'DECLARES WINDOW PANEL MATERIAL AS STRING VARIABLE

        '####################################
        '# WINDOW PANEL RULES AND EQUATIONS #
        '####################################
        '
        WPanelX = VarHeightI * 10
        WPanelY = VarWidthI * 10
        WPanelZ = 16
        WPanelQ = 1 * VarAmountI
        WPanelM = ""

        '################################
        '# DESK SUPPORT GABLE VARIABLES #
        '################################
        '
        Dim DSGableX As Double 'DECLARES WINDOW PANEL X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim DSGableY As Double 'DECLARES WINDOW PANEL Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim DSGableZ As Double 'DECLARES WINDOW PANEL Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim DSGableQ As Integer 'DECLARES WINDOW PANEL QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim DSGableM As String = "" 'DECLARES WINDOW PANEL MATERIAL AS STRING VARIABLE

        '##########################################
        '# DESK SUPPORT GABLE RULES AND EQUATIONS #
        '##########################################
        '
        DSGableX = VarHeightI * 10
        DSGableY = VarWidthI * 10
        DSGableZ = VarDepthI
        DSGableQ = 1 * VarAmountI
        DSGableM = ""

        '################################
        '# FANCY BRACKET VARIABLES #
        '################################
        '
        Dim FBracketX As Double 'DECLARES WINDOW PANEL X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim FBracketY As Double 'DECLARES WINDOW PANEL Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim FBracketZ As Double 'DECLARES WINDOW PANEL Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim FBracketQ As Integer 'DECLARES WINDOW PANEL QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim FBracketM As String = "" 'DECLARES WINDOW PANEL MATERIAL AS STRING VARIABLE

        '##########################################
        '# FANCY BRACKET RULES AND EQUATIONS #
        '##########################################
        '
        FBracketX = VarHeightI * 10
        FBracketY = VarWidthI * 10
        FBracketZ = VarDepthI
        FBracketQ = 1 * VarAmountI
        FBracketM = ""

        '#####################################
        '# FANCY/FURNITURE VALANCE VARIABLES #
        '#####################################
        '
        Dim FValanceX As Double 'DECLARES WINDOW PANEL X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim FValanceY As Double 'DECLARES WINDOW PANEL Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim FValanceZ As Double 'DECLARES WINDOW PANEL Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim FValanceQ As Integer 'DECLARES WINDOW PANEL QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim FValanceM As String = "" 'DECLARES WINDOW PANEL MATERIAL AS STRING VARIABLE

        '##########################################
        '# FANCY/FURNITURE VALANCE RULES AND EQUATIONS #
        '##########################################
        '
        FValanceX = VarHeightI * 10
        FValanceY = VarWidthI * 10
        FValanceZ = VarDepthI
        FValanceQ = 1 * VarAmountI
        FValanceM = ""

        '##################################################
        '# BASE INTEGRATED APPLIANCE, APPLIANCE VARIABLES #
        '##################################################
        '
        Dim BIAAGableX As Double 'DECLARES BASE INTEGRATED APPLIANCE, APPLIANCE GABLE X MEASUREMENT AS DOUBLE
        Dim BIAAGableY As Double 'DECLARES BASE INTEGRATED APPLIANCE, APPLIANCE GABLE Y MEASUREMENT AS DOUBLE
        Dim BIAAGableZ As Double 'DECLARES BASE INTEGRATED APPLIANCE, APPLIANCE GABLE Z MEASUREMENT AS DOUBLE
        Dim BIAAGableQ As Integer 'DECLARES BASE INTEGRATED APPLIANCE, APPLIANCE GABLE Q MEASUREMENT AS INTEGER
        Dim BIAAGableM As String = "" 'DECLARES BASE INTEGRATED APPLIANCE, APPLIANCE GABLE MATERIAL AS STRING
        Dim BIAAGEdgeSeq As String = "" 'DECLARES BASE INTEGRATED APPLIANCE, APPLIANCE GABLE EDGE SEQUENCE AS STRING
        Dim BIAAGEdgeCode As String = "" 'DECLARES BASE INTEGRATED APPLIANCE, APPLIANCE GABLE EDGE CODE AS STRING
        Dim BIAAGEdgeSeq2 As String = "" 'DECLARES BASE INTEGRATED APPLIANCE, APPLIANCE GABLE EDGE SEQUENCE 2 AS STRING
        Dim BIAAGEdgeCode2 As String = "" 'DECLARES BASE INTEGRATED APPLIANCE, APPLIANCE GABLE EDGE CODE 2 AS STRING

        Dim BIAATopStrapX As Double 'DECLARES BASE INTEGRATED APPLIANCE, APPLIANCE TOP STRAP X MEASUREMENT AS DOUBLE
        Dim BIAATopStrapY As Double 'DECLARES BASE INTEGRATED APPLIANCE, APPLIANCE TOP STRAP Y MEASUREMENT AS DOUBLE
        Dim BIAATopStrapZ As Double 'DECLARES BASE INTEGRATED APPLIANCE, APPLIANCE TOP STRAP Z MEASUREMENT AS DOUBLE
        Dim BIAATopStrapQ As Integer 'DECLARES BASE INTEGRATED APPLIANCE, APPLIANCE TOP STRAP Q MEASUREMENT AS INTEGER
        Dim BIAATopStrapM As String = "" 'DECLARES BASE INTEGRATED APPLIANCE, APPLIANCE TOP STRAP MATERIAL AS STRING
        Dim BIAATSEdgeSeq As String = "" 'DECLARES BASE INTEGRATED APPLIANCE, APPLIANCE TOP STRAP EDGE SEQUENCE AS STRING
        Dim BIAATSEdgeCode As String = "" 'DECLARES BASE INTEGRATED APPLIANCE, APPLIANCE TOP STRAP EDGE CODE AS STRING

        Dim BIAABottomX As Double 'DECLARES BASE INTEGRATED APPLIANCE, APPLIANCE BOTTOM X MEASUREMENT AS DOUBLE
        Dim BIAABottomY As Double 'DECLARES BASE INTEGRATED APPLIANCE, APPLIANCE BOTTOM Y MEASUREMENT AS DOUBLE
        Dim BIAABottomZ As Double 'DECLARES BASE INTEGRATED APPLIANCE, APPLIANCE BOTTOM Z MEASUREMENT AS DOUBLE
        Dim BIAABottomQ As Integer 'DECLARES BASE INTEGRATED APPLIANCE, APPLIANCE BOTTOM Q MEASUREMENT AS INTEGER
        Dim BIAABottomM As String = "" 'DECLARES BASE INTEGRATED APPLIANCE, APPLIANCE BOTTOM MATERIAL AS STRING
        Dim BIAABEdgeSeq As String = "" 'DECLARES BASE INTEGRATED APPLIANCE, APPLIANCE BOTTOM EDGE SEQUENCE AS STRING
        Dim BIAABEdgeCode As String = "" 'DECLARES BASE INTEGRATED APPLIANCE, APPLIANCE BOTTOM EDGE CODE AS STRING

        '#############################################
        '# BASE INTEGRATED APPLIANCE, BASE VARIABLES #
        '#############################################
        '
        Dim BIABGableX As Double 'DECLARES BASE INTEGRATED APPLIANCE, BASE GABLE X MEASUREMENT AS DOUBLE
        Dim BIABGableY As Double 'DECLARES BASE INTEGRATED APPLIANCE, BASE GABLE Y MEASUREMENT AS DOUBLE
        Dim BIABGableZ As Double 'DECLARES BASE INTEGRATED APPLIANCE, BASE GABLE Z MEASUREMENT AS DOUBLE
        Dim BIABGableQ As Integer 'DECLARES BASE INTEGRATED APPLIANCE, BASE GABLE Q MEASUREMENT AS INTEGER
        Dim BIABGableM As String = "" 'DECLARES BASE INTEGRATED APPLIANCE, BASE GABLE MATERIAL AS STRING
        Dim BIABGEdgeSeq As String = "" 'DECLARES BASE INTEGRATED APPLIANCE, BASE GABLE EDGE SEQUENCE AS STRING
        Dim BIABGEdgeCode As String = "" 'DECLARES BASE INTEGRATED APPLIANCE, BASE GABLE EDGE CODE AS STRING

        Dim BIABTopStrapX As Double 'DECLARES BASE INTEGRATED APPLIANCE, BASE TOP STRAP X MEASUREMENT AS DOUBLE
        Dim BIABTopStrapY As Double 'DECLARES BASE INTEGRATED APPLIANCE, BASE TOP STRAP Y MEASUREMENT AS DOUBLE
        Dim BIABTopStrapZ As Double 'DECLARES BASE INTEGRATED APPLIANCE, BASE TOP STRAP Z MEASUREMENT AS DOUBLE
        Dim BIABTopStrapQ As Integer 'DECLARES BASE INTEGRATED APPLIANCE, BASE TOP STRAP Q MEASUREMENT AS INTEGER
        Dim BIABTopStrapM As String = "" 'DECLARES BASE INTEGRATED APPLIANCE, BASE TOP STRAP MATERIAL AS STRING
        Dim BIABTSEdgeSeq As String = "" 'DECLARES BASE INTEGRATED APPLIANCE, BASE TOP STRAP EDGE SEQUENCE AS STRING
        Dim BIABTSEdgeCode As String = "" 'DECLARES BASE INTEGRATED APPLIANCE, BASE TOP STRAP EDGE CODE AS STRING

        Dim BIABBottomX As Double 'DECLARES BASE INTEGRATED APPLIANCE, BASE BOTTOM X MEASUREMENT AS DOUBLE
        Dim BIABBottomY As Double 'DECLARES BASE INTEGRATED APPLIANCE, BASE BOTTOM Y MEASUREMENT AS DOUBLE
        Dim BIABBottomZ As Double 'DECLARES BASE INTEGRATED APPLIANCE, BASE BOTTOM Z MEASUREMENT AS DOUBLE
        Dim BIABBottomQ As Integer 'DECLARES BASE INTEGRATED APPLIANCE, BASE BOTTOM Q MEASUREMENT AS INTEGER
        Dim BIABBottomM As String = "" 'DECLARES BASE INTEGRATED APPLIANCE, BASE BOTTOM MATERIAL AS STRING
        Dim BIABBEdgeSeq As String = "" 'DECLARES BASE INTEGRATED APPLIANCE, BASE BOTTOM EDGE SEQUENCE AS STRING
        Dim BIABBEdgeCode As String = "" 'DECLARES BASE INTEGRATED APPLIANCE, BASE BOTTOM EDGE CODE AS STRING

        Dim BIABBackX As Double 'DECLARES BASE INTEGRATED APPLIANCE, BASE BACK X MEASUREMENT AS DOUBLE
        Dim BIABBackY As Double 'DECLARES BASE INTEGRATED APPLIANCE, BASE BACK Y MEASUREMENT AS DOUBLE
        Dim BIABBackZ As Double 'DECLARES BASE INTEGRATED APPLIANCE, BASE BACK Z MEASUREMENT AS DOUBLE
        Dim BIABBackQ As Integer 'DECLARES BASE INTEGRATED APPLIANCE, BASE BACK Q MEASUREMENT AS INTEGER
        Dim BIABBackM As String = "" 'DECLARES BASE INTEGRATED APPLIANCE, BASE BACK MATERIAL AS STRING

        '##################################################################
        '# BASE INTEGRATED APPLIANCE, APPLIANCE GABLE RULES AND EQUATIONS #
        '##################################################################
        '
        BIAAGableX = (VarHeightI * 10) - 1
        BIAAGableY = VarDepthI * 10
        BIAAGableZ = 16
        BIAAGableQ = 2 * VarAmountI
        BIAAGableM = ""

        '######################################################################
        '# BASE INTEGRATED APPLIANCE, APPLIANCE TOP STRAP RULES AND EQUATIONS #
        '######################################################################
        '
        BIAATopStrapX = ((VarWidthI * 10) - (BIAAGableZ * 2))
        BIAATopStrapY = 115
        BIAATopStrapZ = 16
        BIAATopStrapQ = 2 * VarAmountI
        BIAATopStrapM = ""

        '###################################################################
        '# BASE INTEGRATED APPLIANCE, APPLIANCE BOTTOM RULES AND EQUATIONS #
        '###################################################################
        '
        BIAABottomX = ((VarWidthI * 10) - (BIAAGableZ * 2))
        BIAABottomY = VarDepthI * 10
        BIAABottomZ = 16
        BIAABottomQ = 1 * VarAmountI
        BIAABottomM = ""

        '#############################################################
        '# BASE INTEGRATED APPLIANCE, BASE GABLE RULES AND EQUATIONS #
        '#############################################################
        '
        BIABGableX = VarHeightI2 * 10
        BIABGableY = VarDepthI2 * 10
        BIABGableZ = 16
        BIABGableQ = 2 * VarAmountI
        BIABGableM = ""

        '#################################################################
        '# BASE INTEGRATED APPLIANCE, BASE TOP STRAP RULES AND EQUATIONS #
        '#################################################################
        '
        BIABTopStrapX = ((VarWidthI * 10) - (BIABGableZ * 2))
        BIABTopStrapY = 300
        BIABTopStrapZ = 16
        BIABTopStrapQ = 1 * VarAmountI
        BIABTopStrapM = ""

        '##############################################################
        '# BASE INTEGRATED APPLIANCE, BASE BOTTOM RULES AND EQUATIONS #
        '##############################################################
        '
        BIABBottomX = ((VarWidthI * 10) - (BIABGableZ * 2))
        BIABBottomY = VarDepthI2 * 10
        BIABBottomZ = 16
        BIABBottomQ = 1 * VarAmountI
        BIABBottomM = ""

        '##############################################################
        '# BASE INTEGRATED APPLIANCE, BASE BOTTOM RULES AND EQUATIONS #
        '##############################################################
        '
        BIABBackX = (VarWidthI * 10) - 23
        BIABBackY = ((VarHeightI2 * 10) - 120) - 23
        BIABBackZ = 3
        BIABBackQ = 1 * VarAmountI
        BIABBackM = ""

        '#########################
        '# SUPPORT BOX VARIABLES #
        '#########################
        '
        Dim SBGableX As Double 'DECLARES SUPPORT BOX GABLE X MEASUREMENT AS DOUBLE
        Dim SBGableY As Double 'DECLARES SUPPORT BOX GABLE Y MEASUREMENT AS DOUBLE
        Dim SBGableZ As Double 'DECLARES SUPPORT BOX GABLE Z MEASUREMENT AS DOUBLE
        Dim SBGableQ As Integer 'DECLARES SUPPORT BOX GABLE Q MEASUREMENT AS INTEGER
        Dim SBGableM As String = "" 'DECLARES SUPPORT BOX GABLE MATERIAL AS STRING

        Dim SBRailX As Double 'DECLARES SUPPORT BOX TOP STRAP X MEASUREMENT AS DOUBLE
        Dim SBRailY As Double 'DECLARES SUPPORT BOX TOP STRAP Y MEASUREMENT AS DOUBLE
        Dim SBRailZ As Double 'DECLARES SUPPORT BOX TOP STRAP Z MEASUREMENT AS DOUBLE
        Dim SBRailQ As Integer 'DECLARES SUPPORT BOX TOP STRAP Q MEASUREMENT AS INTEGER
        Dim SBRailM As String = "" 'DECLARES SUPPORT BOX TOP STRAP MATERIAL AS STRING

        '#########################################
        '# SUPPORT BOX GABLE RULES AND EQUATIONS #
        '#########################################
        '
        SBGableX = 883
        SBGableY = VarDepthI * 10
        SBGableZ = 16
        SBGableQ = 2 * VarAmountI
        SBGableM = "Scrap"

        '#############################################
        '# SUPPORT BOX TOP STRAP RULES AND EQUATIONS #
        '#############################################
        '
        SBRailX = ((VarWidthI * 10) - (SBGableZ * 2))
        SBRailY = 150
        SBRailZ = 16
        SBRailQ = 8 * VarAmountI
        SBRailM = "Scrap"

        '############################
        '# PLUMBING COVER VARIABLES #
        '############################
        '
        Dim PCPanelX As Double 'DECLARES SUPPORT BOX GABLE X MEASUREMENT AS DOUBLE
        Dim PCPanelY As Double 'DECLARES SUPPORT BOX GABLE Y MEASUREMENT AS DOUBLE
        Dim PCPanelZ As Double 'DECLARES SUPPORT BOX GABLE Z MEASUREMENT AS DOUBLE
        Dim PCPanelQ As Integer 'DECLARES SUPPORT BOX GABLE Q MEASUREMENT AS INTEGER
        Dim PCPanelM As String = "" 'DECLARES SUPPORT BOX GABLE MATERIAL AS STRING

        '#########################################
        '# SUPPORT BOX GABLE RULES AND EQUATIONS #
        '#########################################
        '
        PCPanelX = (VarWidthI + VarHeightI + 5) * 10
        PCPanelY = VarDepthI * 10
        PCPanelZ = 16
        PCPanelQ = 1 * VarAmountI
        PCPanelM = ""

        '##########################
        '# WINDOW BOX VARIABLES #
        '##########################
        '
        Dim WTPanelX As Double 'DECLARES WINDOW PANEL X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim WTPanelY As Double 'DECLARES WINDOW PANEL Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim WTPanelZ As Double 'DECLARES WINDOW PANEL Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim WTPanelQ As Integer 'DECLARES WINDOW PANEL QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim WTPanelM As String = "" 'DECLARES WINDOW PANEL MATERIAL AS STRING VARIABLE

        '##########################
        '# WINDOW FRONT VARIABLES #
        '##########################
        '
        Dim WFrontX As Double 'DECLARES WINDOW PANEL X MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim WFrontY As Double 'DECLARES WINDOW PANEL Y MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim WFrontZ As Double 'DECLARES WINDOW PANEL Z MEASUREMENT AS DOUBLE VARIABLE (ACCEPTS BOTH INTEGERS AND DECIMALS)
        Dim WFrontQ As Integer 'DECLARES WINDOW PANEL QUANITY MEASUREMENT AS INTEGER VARIABLE
        Dim WFrontM As String = "" 'DECLARES WINDOW PANEL MATERIAL AS STRING VARIABLE

        '####################################
        '# WINDOW PANEL RULES AND EQUATIONS #
        '####################################
        '
        WTPanelX = (VarWidthI * 10) - 20
        WTPanelY = VarDepthI * 10
        WTPanelZ = 16
        WTPanelQ = 1 * VarAmountI
        WTPanelM = ""

        '####################################
        '# WINDOW FRONT RULES AND EQUATIONS #
        '####################################
        '
        WFrontX = ((VarDepthI * 10) * 2) + (VarWidthI * 10) + 60
        WFrontY = VarHeightI * 10
        WFrontZ = 16
        WFrontQ = 1 * VarAmountI
        WFrontM = ""

        '########################################################
        '# STANDARD BOX MATERIAL AND EDGING RULES AND VARIABLES #
        '########################################################
        '
        '###########################
        '# BOX MATERIAL EDGE CODES #
        '###########################
        '
        Dim BM = CKWOPLANNERSBoxStyleBox
        Select Case BM
            Case "WHITE MELAMINE BOX"
                Dim EDGECODE As String = "White"
                Dim BOXMATERIAL As String = "W.Mel"
                BMEdgeCode = EDGECODE 'SETS UPPER BOX MATERIAL EDGE CODE TO WHITE
                'UPPER
                UBackM = BOXMATERIAL 'SETS UPPER BACK MATERIAL TO WHITE MELAMINE
                UGableM = BOXMATERIAL 'SETS UPPER GABLE MATERIAL TO WHITE MELAMINE
                UTopLightM = BOXMATERIAL 'SETS UPPER TOP AND BOTTOM MATERIAL TO WHITE MELAMINE
                UBtmM = BOXMATERIAL 'SETS UPPER TOP AND BOTTOM MATERIAL TO WHITE MELAMINE
                UTopBtmM = BOXMATERIAL 'SETS UPPER TOP AND BOTTOM MATERIAL TO WHITE MELAMINE
                UAdjShelfM = BOXMATERIAL 'SETS UPPER ADJUSTABLE SHELF MATERIAL TO WHITE MELAMINE
                UTopHingeM = BOXMATERIAL 'SETS UPPER TOP AND BOTTOM MATERIAL TO WHITE MELAMINE
                UFDividerM = BOXMATERIAL 'SETS UPPER FIXED DIVIDER MATERIAL TO WHITE MELAMINE
                UDividerM = BOXMATERIAL 'SETS UPPER DIVIDER MATERIAL TO WHITE MELAMINE
                'UPPER MICROWAVE
                UMBackM = BOXMATERIAL 'SETS UPPER MICROWAVE BACK MATERIAL TO WHITE MELAMINE
                UMGableM = BOXMATERIAL 'SETS UPPER MICROWAVE BACK MATERIAL TO WHITE MELAMINE
                UMTopM = BOXMATERIAL 'SETS UPPER MICROWAVE BACK MATERIAL TO WHITE MELAMINE
                UMBtmM = BOXMATERIAL 'SETS UPPER MICROWAVE BACK MATERIAL TO WHITE MELAMINE
                UMAdjShelfM = BOXMATERIAL 'SETS UPPER MICROWAVE BACK MATERIAL TO WHITE MELAMINE
                'UPPER CORNER DIAGONAL
                UCDBackM = BOXMATERIAL 'SETS UPPER CORNER DIAGONAL BACK MATERIAL TO WHITE MELAMINE
                UCDBackStrapM = BOXMATERIAL 'SETS UPPER CORNER DIAGONAL BACK STRAP MATERIAL TO WHITE MELAMINE
                UCDGableM = BOXMATERIAL 'SETS UPPER CORNER DIAGONAL GABLE MATERIAL TO WHITE MELAMINE
                UCDTopBtmM = BOXMATERIAL 'SETS UPPER CORNER DIAGONAL TOP AND BOTTOM MATERIAL TO WHITE MELAMINE
                UCDAdjShelfM = BOXMATERIAL 'SETS UPPER CORNER DIAGONAL ADJUSTABLE SHELF MATERIAL TO WHITE MELAMINE
                'BASE
                BBackM = BOXMATERIAL 'SETS BASE BACK MATERIAL TO WHITE MELAMINE
                BGableM = BOXMATERIAL 'SETS BASE GABLE MATERIAL TO WHITE MELAMINE
                BTopStrapM = BOXMATERIAL 'SETS BASE TOP STRAP MATERIAL TO WHITE MELAMINE
                BBtmM = BOXMATERIAL 'SETS BASE BOTTOM MATERIAL TO WHITE MELAMINE
                BAdjShelfM = BOXMATERIAL 'SETS BASE ADJUSTABLE SHELF MATERIAL TO WHITE MELAMINE
                BTopM = BOXMATERIAL 'SETS BASE TOP MATERIAL TO WHITE MELAMINE
                BStrapM = BOXMATERIAL 'SETS BASE STRAP MATERIAL TO WHITE MELAMINE
                BDividerM = BOXMATERIAL 'SETS BASE DIVIDER MATERIAL TO WHITE MELAMINE
                BFFShelfM = BOXMATERIAL 'SETS BASE FULL FIXED SHELF MATERIAL TO WHITE MELAMINE
                '# BASE MICROWAVE OPEN SHELF 1 DRAWER #
                BMO1DBBackM = BOXMATERIAL 'SETS BASE MICROWAVE BACK MATERIAL TO WHITE MELAMINE
                BMO1DBGableM = BOXMATERIAL 'SETS BASE MICROWAVE GABLE MATERIAL TO WHITE MELAMINE
                BMO1DBTopStrapM = BOXMATERIAL 'SETS BASE MICROWAVE TOP STRAP MATERIAL TO WHITE MELAMINE
                BMO1DBBtmM = BOXMATERIAL 'SETS BASE MICROWAVE BOTTOM MATERIAL TO WHITE MELAMINE
                'TALL UTILITY 1 UNIT
                TU1UBackM = BOXMATERIAL 'SETS TALL UTILITY BACK MATERIAL TO WHITE MELAMINE
                TU1UGableM = BOXMATERIAL 'SETS TALL UTILITY GABLE MATERIAL TO WHITE MELAMINE
                TU1UTopM = BOXMATERIAL 'SETS TALL UTILITY TOP + BOTTOM MATERIAL TO WHITE MELAMINE
                TU1UBotM = BOXMATERIAL 'SETS TALL UTILITY TOP + BOTTOM MATERIAL TO WHITE MELAMINE
                TU1USFSM = BOXMATERIAL 'SETS TALL UTILITY SHALLOW FIXED SHELF MATERIAL TO WHITE MELAMINE
                TU1UFFSM = BOXMATERIAL 'SETS TALL UTILITY FULL FIXED SHELF MATERIAL TO WHITE MELAMINE
                TU1UStrapM = BOXMATERIAL 'SETS TALL UTILITY STRAP MATERIAL TO WHITE MELAMINE
                TU1UAdjShelfM = BOXMATERIAL 'SETS TALL UTILITY ADJUSTABLE SHELF BACK MATERIAL TO WHITE MELAMINE
                'TALL UTILITY 2 UNITS UPPER
                TU2UUBackM = BOXMATERIAL 'SETS TALL UTILITY UPPER BACK MATERIAL TO WHITE MELAMINE
                TU2UUGableM = BOXMATERIAL 'SETS TALL UTILITY UPPER GABLE MATERIAL TO WHITE MELAMINE
                TU2UUTopBtmM = BOXMATERIAL 'SETS TALL UTILITY UPPER TOP + BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUTopM = BOXMATERIAL 'SETS TALL UTILITY UPPER TOP MATERIAL TO WHITE MELAMINE
                TU2UUBotM = BOXMATERIAL 'SETS TALL UTILITY UPPER BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUAdjShelfM = BOXMATERIAL 'SETS TALL UTILITY UPPER ADJUSTABLE SHELF BACK MATERIAL TO WHITE MELAMINE
                'TALL UTILITY 2 UNITS TALL
                TU2UBBackM = BOXMATERIAL 'SETS TALL UTILITY TALL BACK MATERIAL TO WHITE MELAMINE
                TU2UBGableM = BOXMATERIAL 'SETS TALL UTILITY TALL GABLE MATERIAL TO WHITE MELAMINE
                TU2UBTopBtmM = BOXMATERIAL 'SETS TALL UTILITY TALL TOP + BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UBTopM = BOXMATERIAL 'SETS TALL UTILITY TALL TOP MATERIAL TO WHITE MELAMINE
                TU2UBBotM = BOXMATERIAL 'SETS TALL UTILITY TALL BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UBAdjShelfM = BOXMATERIAL 'SETS TALL UTILITY TALL ADJUSTABLE SHELF BACK MATERIAL TO WHITE MELAMINE
                TU2UBSFSM = BOXMATERIAL 'SETS TALL UTILITY TALL SHALLOW FIXED SHELF BACK MATERIAL TO WHITE MELAMINE
                'HUTCH
                HBackM = BOXMATERIAL 'SETS HUTCH BACK MATERIAL TO WHITE MELAMINE
                HGableM = BOXMATERIAL 'SETS HUTCH GABLE MATERIAL TO WHITE MELAMINE
                HTopFFShelfM = BOXMATERIAL 'SETS HUTCH TOP FULL FIXED SHELF MATERIAL TO WHITE MELAMINE
                HBtmFFShelfM = BOXMATERIAL 'SETS HUTCH BOTTOM FULL FIXED SHELF MATERIAL TO WHITE MELAMINE
                HTopM = BOXMATERIAL 'SETS HUTCH TOP MATERIAL TO WHITE MELAMINE
                HBtmM = BOXMATERIAL 'SETS HUTCH BOTTOM MATERIAL TO WHITE MELAMINE
                HStrapM = BOXMATERIAL 'SETS HUTCH STRAP MATERIAL TO WHITE MELAMINE
                HAdjShelfM = BOXMATERIAL 'SETS HUTCH ADJUSTABLE SHELF MATERIAL TO WHITE MELAMINE
                HDividerM = BOXMATERIAL 'SETS HUTCH DIVIDER MATERIAL TO WHITE MELAMINE
                'VANITY
                VBackM = BOXMATERIAL 'SETS VANITY BACK MATERIAL TO WHITE MELAMINE
                VGableM = BOXMATERIAL 'SETS VANITY GABLE MATERIAL TO WHITE MELAMINE
                VTopStrapM = BOXMATERIAL 'SETS VANITY TOP STRAP MATERIAL TO WHITE MELAMINE
                VTopM = BOXMATERIAL 'SETS VANITY TOP MATERIAL TO WHITE MELAMINE
                VBtmM = BOXMATERIAL 'SETS VANITY BOTTOM MATERIAL TO WHITE MELAMINE
                VFFShelfM = BOXMATERIAL 'SETS VANITY FULL FIXED SHELF MATERIAL TO WHITE MELAMINE
                VStrapM = BOXMATERIAL 'SETS VANITY STRAP MATERIAL TO WHITE MELAMINE
                VDividerM = BOXMATERIAL 'SETS VANITY DIVIDER MATERIAL TO WHITE MELAMINE
                'VANITY ELEVATED
                VEBackM = BOXMATERIAL 'SETS VANITY ELEVATED BACK MATERIAL TO WHITE MELAMINE
                VEGableM = BOXMATERIAL 'SETS VANITY ELEVATED GABLE MATERIAL TO WHITE MELAMINE
                VETopStrapM = BOXMATERIAL 'SETS VANITY ELEVATED TOP STRAP MATERIAL TO WHITE MELAMINE
                VETopM = BOXMATERIAL 'SETS VANITY ELEVATED TOP MATERIAL TO WHITE MELAMINE
                VEBtmM = BOXMATERIAL 'SETS VANITY ELEVATED BOTTOM MATERIAL TO WHITE MELAMINE
                VEFFShelfM = BOXMATERIAL 'SETS VANITY ELEVATED FULL FIXED SHELF MATERIAL TO WHITE MELAMINE
                VEStrapM = BOXMATERIAL 'SETS VANITY ELEVATED STRAP MATERIAL TO WHITE MELAMINE
                VEDividerM = BOXMATERIAL 'SETS VANITY ELEVATED DIVIDER MATERIAL TO WHITE MELAMINE
                '# CANOPY #
                RHBackM = BOXMATERIAL 'SETS CANOPY BACK MATERIAL TO WHITE MELAMINE
                RHTopM = BOXMATERIAL 'SETS CANOPY TOP MATERIAL TO WHITE MELAMINE
                RHGableM = BOXMATERIAL 'SETS CANOPY GABLE MATERIAL TO WHITE MELAMINE
                '# BASE INTEGRATED APPLIANCE #
                BIAAGableM = BOXMATERIAL 'SETS BASE INTEGRATED APPLIANCE, APPLIANCE GABLE MATERIAL TO WHITE MELAMINE
                BIAATopStrapM = BOXMATERIAL 'SETS BASE INTEGRATED APPLIANCE, APPLIANCE TOP STRAP MATERIAL TO WHITE MELAMINE
                BIAABottomM = BOXMATERIAL 'SETS BASE INTEGRATED APPLIANCE, APPLIANCE BOTTOM MATERIAL TO WHITE MELAMINE
                BIABGableM = BOXMATERIAL 'SETS BASE INTEGRATED APPLIANCE, BASE GABLE MATERIAL TO WHITE MELAMINE
                BIABTopStrapM = BOXMATERIAL 'SETS BASE INTEGRATED APPLIANCE, BASE TOP STRAP MATERIAL TO WHITE MELAMINE
                BIABBottomM = BOXMATERIAL 'SETS BASE INTEGRATED APPLIANCE, BASE BOTTOM MATERIAL TO WHITE MELAMINE
                BIABBackM = BOXMATERIAL 'SETS BASE INTEGRATED APPLIANCE, BASE BACK MATERIAL TO WHITE MELAMINE

            Case "HARDROCK MAPLE BOX"
                Dim EDGECODE As String = "6116"
                Dim BOXMATERIAL As String = "HRM"
                BMEdgeCode = EDGECODE 'SETS UPPER BOX MATERIAL EDGE CODE TO 6116 (HARD ROCK MAPLE)
                'UPPER
                UBackM = BOXMATERIAL 'SETS UPPER BACK MATERIAL TO HARD ROCK MAPLE
                UGableM = BOXMATERIAL 'SETS UPPER GABLE MATERIAL TO HARD ROCK MAPLE
                UTopLightM = BOXMATERIAL 'SETS UPPER TOP AND BOTTOM MATERIAL TO HARD ROCK MAPLE
                UBtmM = BOXMATERIAL 'SETS UPPER TOP AND BOTTOM MATERIAL TO HARD ROCK MAPLE
                UTopBtmM = BOXMATERIAL 'SETS UPPER TOP AND BOTTOM MATERIAL TO HARD ROCK MAPLE
                UAdjShelfM = BOXMATERIAL 'SETS UPPER ADJUSTABLE SHELF MATERIAL TO HARD ROCK MAPLE
                UTopHingeM = BOXMATERIAL 'SETS UPPER TOP AND BOTTOM MATERIAL TO HARD ROCK MAPLE
                UFDividerM = BOXMATERIAL 'SETS UPPER FIXED DIVIDER MATERIAL TO HARD ROCK MAPLE
                UDividerM = BOXMATERIAL 'SETS UPPER DIVIDER MATERIAL TO WHITE HARD ROCK MAPLE
                'UPPER MICROWAVE
                UMBackM = BOXMATERIAL 'SETS UPPER MICROWAVE BACK MATERIAL TO HARD ROCK MAPLE
                UMGableM = BOXMATERIAL 'SETS UPPER MICROWAVE BACK MATERIAL TO HARD ROCK MAPLE
                UMTopM = BOXMATERIAL 'SETS UPPER MICROWAVE BACK MATERIAL TO HARD ROCK MAPLE
                UMBtmM = BOXMATERIAL 'SETS UPPER MICROWAVE BACK MATERIAL TO HARD ROCK MAPLE
                UMAdjShelfM = BOXMATERIAL 'SETS UPPER MICROWAVE BACK MATERIAL TO HARD ROCK MAPLE
                'UPPER CORNER DIAGONAL
                UCDBackM = BOXMATERIAL 'SETS UPPER CORNER DIAGONAL BACK MATERIAL TO HARD ROCK MAPLE
                UCDBackStrapM = BOXMATERIAL 'SETS UPPER CORNER DIAGONAL BACK STRAP MATERIAL TO HARD ROCK MAPLE
                UCDGableM = BOXMATERIAL 'SETS UPPER CORNER DIAGONAL GABLE MATERIAL TO HARD ROCK MAPLE
                UCDTopBtmM = BOXMATERIAL 'SETS UPPER CORNER DIAGONAL TOP AND BOTTOM MATERIAL TO HARD ROCK MAPLE
                UCDAdjShelfM = BOXMATERIAL 'SETS UPPER CORNER DIAGONAL ADJUSTABLE SHELF MATERIAL TO HARD ROCK MAPLE
                'BASE
                BBackM = BOXMATERIAL 'SETS BASE BACK MATERIAL TO HARD ROCK MAPLE
                BGableM = BOXMATERIAL 'SETS BASE GABLE MATERIAL TO HARD ROCK MAPLE
                BTopStrapM = BOXMATERIAL 'SETS BASE TOP STRAP MATERIAL TO HARD ROCK MAPLE
                BBtmM = BOXMATERIAL 'SETS BASE BOTTOM MATERIAL TO HARD ROCK MAPLE
                BAdjShelfM = BOXMATERIAL 'SETS BASE ADJUSTABLE SHELF MATERIAL TO HARD ROCK MAPLE
                '# BASE MICROWAVE OPEN SHELF 1 DRAWER #
                BMO1DBBackM = BOXMATERIAL
                BMO1DBGableM = BOXMATERIAL
                BMO1DBTopStrapM = BOXMATERIAL
                BMO1DBBtmM = BOXMATERIAL
                'TALL UTILITY 1 UNIT
                TU1UBackM = BOXMATERIAL 'SETS TALL UTILITY BACK MATERIAL TO HARD ROCK MAPLE
                TU1UGableM = BOXMATERIAL 'SETS TALL UTILITY GABLE MATERIAL TO HARD ROCK MAPLE
                TU1UTopM = BOXMATERIAL 'SETS TALL UTILITY TOP + BOTTOM MATERIAL TO HARD ROCK MAPLE
                TU1UBotM = BOXMATERIAL 'SETS TALL UTILITY TOP + BOTTOM MATERIAL TO HARD ROCK MAPLE
                TU1USFSM = BOXMATERIAL 'SETS TALL UTILITY SHALLOW FIXED SHELF MATERIAL TO HARD ROCK MAPLE
                TU1UFFSM = BOXMATERIAL 'SETS TALL UTILITY FULL FIXED SHELF MATERIAL TO HARD ROCK MAPLE
                TU1UStrapM = BOXMATERIAL 'SETS TALL UTILITY STRAP MATERIAL TO HARD ROCK MAPLE
                TU1UAdjShelfM = BOXMATERIAL 'SETS TALL UTILITY ADJUSTABLE SHELF BACK MATERIAL TO HARD ROCK MAPLE
                'TALL UTILITY 2 UNITS UPPER
                TU2UUBackM = BOXMATERIAL 'SETS TALL UTILITY UPPER BACK MATERIAL TO HARD ROCK MAPLE
                TU2UUGableM = BOXMATERIAL 'SETS TALL UTILITY UPPER GABLE MATERIAL TO HARD ROCK MAPLE
                TU2UUTopBtmM = BOXMATERIAL 'SETS TALL UTILITY UPPER TOP + BOTTOM MATERIAL TO HARD ROCK MAPLE
                TU2UUTopM = BOXMATERIAL 'SETS TALL UTILITY UPPER TOP MATERIAL TO HARD ROCK MAPLE
                TU2UUBotM = BOXMATERIAL 'SETS TALL UTILITY UPPER BOTTOM MATERIAL TO HARD ROCK MAPLE
                TU2UUAdjShelfM = BOXMATERIAL 'SETS TALL UTILITY UPPER ADJUSTABLE SHELF BACK MATERIAL TO HARD ROCK MAPLE
                'TALL UTILITY 2 UNITS TALL
                TU2UBBackM = BOXMATERIAL 'SETS TALL UTILITY TALL BACK MATERIAL TO HARD ROCK MAPLE
                TU2UBGableM = BOXMATERIAL 'SETS TALL UTILITY TALL GABLE MATERIAL TO HARD ROCK MAPLE
                TU2UBTopBtmM = BOXMATERIAL 'SETS TALL UTILITY TALL TOP + BOTTOM MATERIAL TO HARD ROCK MAPLE
                TU2UBTopM = BOXMATERIAL 'SETS TALL UTILITY TALL TOP MATERIAL TO HARD ROCK MAPLE
                TU2UBBotM = BOXMATERIAL 'SETS TALL UTILITY TALL BOTTOM MATERIAL TO HARD ROCK MAPLE
                TU2UBAdjShelfM = BOXMATERIAL 'SETS TALL UTILITY TALL ADJUSTABLE SHELF BACK MATERIAL TO HARD ROCK MAPLE
                TU2UBSFSM = BOXMATERIAL 'SETS TALL UTILITY TALL SHALLOW FIXED SHELF BACK MATERIAL TO HARD ROCK MAPLE
                'HUTCH
                HBackM = BOXMATERIAL 'SETS HUTCH BACK MATERIAL TO HARD ROCK MAPLE
                HGableM = BOXMATERIAL 'SETS HUTCH GABLE MATERIAL TO HARD ROCK MAPLE
                HTopFFShelfM = BOXMATERIAL 'SETS HUTCH TOP FULL FIXED SHELF MATERIAL TO HARD ROCK MAPLE
                HBtmFFShelfM = BOXMATERIAL 'SETS HUTCH BOTTOM FULL FIXED SHELF MATERIAL TO HARD ROCK MAPLE
                HTopM = BOXMATERIAL 'SETS HUTCH TOP MATERIAL TO HARD ROCK MAPLE
                HBtmM = BOXMATERIAL 'SETS HUTCH BOTTOM MATERIAL TO HARD ROCK MAPLE
                HStrapM = BOXMATERIAL 'SETS HUTCH TOP AND BOTTOM MATERIAL TO HARD ROCK MAPLE
                HAdjShelfM = BOXMATERIAL 'SETS HUTCH ADJUSTABLE SHELF MATERIAL TO HARD ROCK MAPLE
                HDividerM = BOXMATERIAL 'SETS HUTCH DIVIDER MATERIAL TO HARD ROCK MAPLE
                'VANITY
                VBackM = BOXMATERIAL 'SETS VANITY BACK MATERIAL TO HARD ROCK MAPLE
                VGableM = BOXMATERIAL 'SETS VANITY GABLE MATERIAL TO HARD ROCK MAPLE
                VTopStrapM = BOXMATERIAL 'SETS VANITY TOP AND BOTTOM MATERIAL TO HARD ROCK MAPLE
                VTopM = BOXMATERIAL 'SETS VANITY TOP AND BOTTOM MATERIAL TO HARD ROCK MAPLE
                VBtmM = BOXMATERIAL 'SETS VANITY TOP AND BOTTOM MATERIAL TO HARD ROCK MAPLE
                VFFShelfM = BOXMATERIAL 'SETS VANITY TOP AND BOTTOM MATERIAL TO HARD ROCK MAPLE
                VStrapM = BOXMATERIAL 'SETS VANITY STRAP MATERIAL TO HARD ROCK MAPLE
                VDividerM = BOXMATERIAL 'SETS VANITY DIVIDER MATERIAL TO HARD ROCK MAPLE
                'VANITY ELEVATED
                VEBackM = BOXMATERIAL 'SETS VANITY ELEVATED BACK MATERIAL TO HARD ROCK MAPLE
                VEGableM = BOXMATERIAL 'SETS VANITY ELEVATED GABLE MATERIAL TO HARD ROCK MAPLE
                VETopStrapM = BOXMATERIAL 'SETS VANITY ELEVATED TOP STRAP MATERIAL TO HARD ROCK MAPLE
                VETopM = BOXMATERIAL 'SETS VANITY ELEVATED TOP MATERIAL TO HARD ROCK MAPLE
                VEBtmM = BOXMATERIAL 'SETS VANITY ELEVATED BOTTOM MATERIAL TO HARD ROCK MAPLE
                VEFFShelfM = BOXMATERIAL 'SETS VANITY ELEVATED FULL FIXED SHELF MATERIAL TO HARD ROCK MAPLE
                VEStrapM = BOXMATERIAL 'SETS VANITY ELEVATED STRAP MATERIAL TO HARD ROCK MAPLE
                VEDividerM = BOXMATERIAL 'SETS VANITY ELEVATED DIVIDER MATERIAL TO HARD ROCK MAPLE
                '# CANOPY #
                RHBackM = BOXMATERIAL 'SETS CANOPY BACK MATERIAL TO HARD ROCK MAPLE
                RHTopM = BOXMATERIAL 'SETS CANOPY TOP MATERIAL TO HARD ROCK MAPLE
                RHGableM = BOXMATERIAL 'SETS CANOPY GABLE MATERIAL TO HARD ROCK MAPLE
                '# BASE INTEGRATED APPLIANCE #
                BIAAGableM = BOXMATERIAL 'SETS BASE INTEGRATED APPLIANCE, APPLIANCE GABLE MATERIAL TO HARD ROCK MAPLE
                BIAATopStrapM = BOXMATERIAL 'SETS BASE INTEGRATED APPLIANCE, APPLIANCE TOP STRAP MATERIAL TO HARD ROCK MAPLE
                BIAABottomM = BOXMATERIAL 'SETS BASE INTEGRATED APPLIANCE, APPLIANCE BOTTOM MATERIAL TO HARD ROCK MAPLE
                BIABGableM = BOXMATERIAL 'SETS BASE INTEGRATED APPLIANCE, BASE GABLE MATERIAL TO HARD ROCK MAPLE
                BIABTopStrapM = BOXMATERIAL 'SETS BASE INTEGRATED APPLIANCE, BASE TOP STRAP MATERIAL TO HARD ROCK MAPLE
                BIABBottomM = BOXMATERIAL 'SETS BASE INTEGRATED APPLIANCE, BASE BOTTOM MATERIAL TO HARD ROCK MAPLE
                BIABBackM = BOXMATERIAL 'SETS BASE INTEGRATED APPLIANCE, BASE BACK MATERIAL TO HARD ROCK MAPLE
            Case Else
        End Select

        '###################################
        '# STANDARD MATERIAL SPECIES RULES #
        '###################################
        '
        Dim Species = CKWOPLANNERSSpeciesBox
        Select Case Species
            Case "CHERRY"
                Dim VMATERIAL As String = "Cherry"
                Dim VSMATERIAL As String = "Solid Cherry"
                Dim VEDGECODE As String = "CH/V"
                'TALL UTILITY 2 UNITS UPPER
                TU2UUMIBackM = VMATERIAL 'SETS TALL UTILITY UPPER BACK MATERIAL TO WHITE MELAMINE
                TU2UUMIGableM = VMATERIAL 'SETS TALL UTILITY UPPER GABLE MATERIAL TO WHITE MELAMINE
                TU2UUMITopBtmM = VMATERIAL 'SETS TALL UTILITY UPPER TOP + BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMITopM = VMATERIAL 'SETS TALL UTILITY UPPER TOP MATERIAL TO WHITE MELAMINE
                TU2UUMIBotM = VMATERIAL 'SETS TALL UTILITY UPPER BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMIAdjShelfM = VMATERIAL 'SETS TALL UTILITY UPPER ADJUSTABLE SHELF BACK MATERIAL TO WHITE MELAMINE
                '# UPPER MATCHING INTERIOR #
                UMIBackM = VMATERIAL
                UMIGableM = VMATERIAL
                UMITopLightM = VMATERIAL
                UMIBtmM = VMATERIAL
                UMITopBtmM = VMATERIAL
                UMIAdjShelfM = VMATERIAL
                '# UPPER MICROWAVE MATCHING INTERIOR #
                UMMIBackM = VMATERIAL
                UMMIGableM = VMATERIAL
                UMMITopBtmM = VMATERIAL
                UMMIMicroSHM = VMATERIAL
                '# UPPER CORNER DIAGONAL MATCHING INTERIOR #
                UCDMILBackM = VMATERIAL
                UCDMISBackM = VMATERIAL
                UCDMIGableM = VMATERIAL
                UCDMITopBtmM = VMATERIAL
                UCDMIAdjShelfM = VMATERIAL
                '# UPPER WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                UWRBackM = VMATERIAL
                UWRGableM = VMATERIAL
                UWRTopBtmM = VMATERIAL
                UWRFFSHM = VMATERIAL
                UWRFShelfM = VMATERIAL
                UWRDivM = VMATERIAL
                '# UPPER END SHELF #
                UESLGableM = VMATERIAL
                UESSGableM = VMATERIAL
                UESTopM = VMATERIAL
                UESFixedShelfM = VMATERIAL
                '# BASE MATCHING INTERIOR #
                BMIBackM = VMATERIAL
                BMIGableM = VMATERIAL
                BMITopStrapM = VMATERIAL
                BMIBtmM = VMATERIAL
                BMIAdjShelfM = VMATERIAL
                BMITopM = VMATERIAL
                BMIStrapM = VMATERIAL
                BMIDividerM = VMATERIAL
                '# BASE PENINSULA OPEN #
                BPOPTopBtmFSHM = VMATERIAL
                BPOPSGableM = VMATERIAL
                BPOPLGableM = VMATERIAL
                '# VANITY MATCHING INTERIOR #
                VMIBackM = VMATERIAL
                VMIGableM = VMATERIAL
                VMITopStrapM = VMATERIAL
                VMITopM = VMATERIAL
                VMIBtmM = VMATERIAL
                VMIFFShelfM = VMATERIAL
                VMIStrapM = VMATERIAL
                VMIDividerM = VMATERIAL
                '# VANITY ELEVATED MATCHING INTERIOR #
                VEMIBackM = VMATERIAL
                VEMIGableM = VMATERIAL
                VEMITopStrapM = VMATERIAL
                VEMITopM = VMATERIAL
                VEMIBtmM = VMATERIAL
                VEMIFFShelfM = VMATERIAL
                VEMIStrapM = VMATERIAL
                VEMIDividerM = VMATERIAL
                '# BASE MICROWAVE OPEN SHELF 1 DRAWER #
                BMO1DUBackM = VMATERIAL
                BMO1DUGableM = VMATERIAL
                BMO1DUTopM = VMATERIAL
                BMO1DUBtmM = VMATERIAL
                '# BASE WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                BWRBackM = VMATERIAL
                BWRGableM = VMATERIAL
                BWRTopBtmM = VMATERIAL
                BWRFFSHM = VMATERIAL
                BWRValanceM = VSMATERIAL
                BWRVN = "Ordered Joe G."
                '# HUTCH MATCHING INTERIOR MATCHING INTERIOR #
                HMIBackM = VMATERIAL
                HMIGableM = VMATERIAL
                HMITopFFShelfM = VMATERIAL
                HMIBtmFFShelfM = VMATERIAL
                HMITopM = VMATERIAL
                HMIBtmM = VMATERIAL
                HMIStrapM = VMATERIAL
                HMIDividerM = VMATERIAL
                HMIAdjShelfM = VMATERIAL
                '# CANOPY #
                RHFanSHM = "Ply"
                RHFrontM = VMATERIAL
                RHSFrontM = VMATERIAL
                RHSTopM = VMATERIAL
                RHIPanelM = VMATERIAL
                '# OVEN PANEL MATERIAL #
                OV1M = VMATERIAL
                '# FLOATING SHELF MATERIAL #
                FShelfM = VMATERIAL
                '# WINDOW PANEL MATERIAL #
                WPanelM = VMATERIAL
                '# WINDOW BOX MATERIAL #
                WTPanelM = VMATERIAL
                WFrontM = VMATERIAL
                '# DESK SUPPORT GABLE MATERIAL #
                DSGableM = VMATERIAL
                '# FANCY BRACKET MATERIAL #
                FBracketM = VMATERIAL
                '# FANCY VALANCE MATERIAL #
                FValanceM = VMATERIAL
                '# PLUMBING COVER PANEL MATERIAL #
                PCPanelM = VMATERIAL
                '# EDGE CODES VENEER #
                VeneerEdgeCode = VEDGECODE

            Case "MAPLE"
                Dim VMATERIAL As String = "Maple"
                Dim VSMATERIAL As String = "Solid Maple"
                Dim VEDGECODE As String = "M/V"
                'TALL UTILITY 2 UNITS UPPER
                TU2UUMIBackM = VMATERIAL 'SETS TALL UTILITY UPPER BACK MATERIAL TO WHITE MELAMINE
                TU2UUMIGableM = VMATERIAL 'SETS TALL UTILITY UPPER GABLE MATERIAL TO WHITE MELAMINE
                TU2UUMITopBtmM = VMATERIAL 'SETS TALL UTILITY UPPER TOP + BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMITopM = VMATERIAL 'SETS TALL UTILITY UPPER TOP MATERIAL TO WHITE MELAMINE
                TU2UUMIBotM = VMATERIAL 'SETS TALL UTILITY UPPER BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMIAdjShelfM = VMATERIAL 'SETS TALL UTILITY UPPER ADJUSTABLE SHELF BACK MATERIAL TO WHITE MELAMINE
                '# UPPER MATCHING INTERIOR #
                UMIBackM = VMATERIAL
                UMIGableM = VMATERIAL
                UMITopLightM = VMATERIAL
                UMIBtmM = VMATERIAL
                UMITopBtmM = VMATERIAL
                UMIAdjShelfM = VMATERIAL
                '# UPPER MICROWAVE MATCHING INTERIOR #
                UMMIBackM = VMATERIAL
                UMMIGableM = VMATERIAL
                UMMITopBtmM = VMATERIAL
                UMMIMicroSHM = VMATERIAL
                '# UPPER CORNER DIAGONAL MATCHING INTERIOR #
                UCDMILBackM = VMATERIAL
                UCDMISBackM = VMATERIAL
                UCDMIGableM = VMATERIAL
                UCDMITopBtmM = VMATERIAL
                UCDMIAdjShelfM = VMATERIAL
                '# UPPER WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                UWRBackM = VMATERIAL
                UWRGableM = VMATERIAL
                UWRTopBtmM = VMATERIAL
                UWRFFSHM = VMATERIAL
                UWRFShelfM = VMATERIAL
                UWRDivM = VMATERIAL
                '# UPPER END SHELF #
                UESLGableM = VMATERIAL
                UESSGableM = VMATERIAL
                UESTopM = VMATERIAL
                UESFixedShelfM = VMATERIAL
                '# BASE MATCHING INTERIOR #
                BMIBackM = VMATERIAL
                BMIGableM = VMATERIAL
                BMITopStrapM = VMATERIAL
                BMIBtmM = VMATERIAL
                BMIAdjShelfM = VMATERIAL
                BMITopM = VMATERIAL
                BMIStrapM = VMATERIAL
                BMIDividerM = VMATERIAL
                '# BASE PENINSULA OPEN #
                BPOPTopBtmFSHM = VMATERIAL
                BPOPSGableM = VMATERIAL
                BPOPLGableM = VMATERIAL
                '# VANITY MATCHING INTERIOR #
                VMIBackM = VMATERIAL
                VMIGableM = VMATERIAL
                VMITopStrapM = VMATERIAL
                VMITopM = VMATERIAL
                VMIBtmM = VMATERIAL
                VMIFFShelfM = VMATERIAL
                VMIStrapM = VMATERIAL
                VMIDividerM = VMATERIAL
                '# VANITY ELEVATED MATCHING INTERIOR #
                VEMIBackM = VMATERIAL
                VEMIGableM = VMATERIAL
                VEMITopStrapM = VMATERIAL
                VEMITopM = VMATERIAL
                VEMIBtmM = VMATERIAL
                VEMIFFShelfM = VMATERIAL
                VEMIStrapM = VMATERIAL
                VEMIDividerM = VMATERIAL
                '# BASE MICROWAVE OPEN SHELF 1 DRAWER #
                BMO1DUBackM = VMATERIAL
                BMO1DUGableM = VMATERIAL
                BMO1DUTopM = VMATERIAL
                BMO1DUBtmM = VMATERIAL
                '# BASE WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                BWRBackM = VMATERIAL
                BWRGableM = VMATERIAL
                BWRTopBtmM = VMATERIAL
                BWRFFSHM = VMATERIAL
                BWRValanceM = VSMATERIAL
                BWRVN = "Ordered Joe G."
                '# HUTCH MATCHING INTERIOR MATCHING INTERIOR #
                HMIBackM = VMATERIAL
                HMIGableM = VMATERIAL
                HMITopFFShelfM = VMATERIAL
                HMIBtmFFShelfM = VMATERIAL
                HMITopM = VMATERIAL
                HMIBtmM = VMATERIAL
                HMIStrapM = VMATERIAL
                HMIDividerM = VMATERIAL
                HMIAdjShelfM = VMATERIAL
                '# CANOPY #
                RHFanSHM = "Ply"
                RHFrontM = VMATERIAL
                RHSFrontM = VMATERIAL
                RHSTopM = VMATERIAL
                RHIPanelM = VMATERIAL
                '# OVEN PANEL MATERIAL #
                OV1M = VMATERIAL
                '# FLOATING SHELF MATERIAL #
                FShelfM = VMATERIAL
                '# WINDOW PANEL MATERIAL #
                WPanelM = VMATERIAL
                '# WINDOW BOX MATERIAL #
                WTPanelM = VMATERIAL
                WFrontM = VMATERIAL
                '# DESK SUPPORT GABLE MATERIAL #
                DSGableM = VMATERIAL
                '# FANCY BRACKET MATERIAL #
                FBracketM = VMATERIAL
                '# FANCY VALANCE MATERIAL #
                FValanceM = VMATERIAL
                '# PLUMBING COVER PANEL MATERIAL #
                PCPanelM = VMATERIAL
                '# EDGE CODES VENEER #
                VeneerEdgeCode = VEDGECODE

            Case "MDF"
                Dim VMATERIAL As String = "MDF"
                Dim VMATERIAL2 As String = "Maple"
                Dim VSMATERIAL As String = "Solid Poplar"
                Dim VEDGECODE As String = "M/V"
                'TALL UTILITY 2 UNITS UPPER
                TU2UUMIBackM = VMATERIAL 'SETS TALL UTILITY UPPER BACK MATERIAL TO WHITE MELAMINE
                TU2UUMIGableM = VMATERIAL 'SETS TALL UTILITY UPPER GABLE MATERIAL TO WHITE MELAMINE
                TU2UUMITopBtmM = VMATERIAL2 'SETS TALL UTILITY UPPER TOP + BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMITopM = VMATERIAL2 'SETS TALL UTILITY UPPER TOP MATERIAL TO WHITE MELAMINE
                TU2UUMIBotM = VMATERIAL2 'SETS TALL UTILITY UPPER BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMIAdjShelfM = VMATERIAL 'SETS TALL UTILITY UPPER ADJUSTABLE SHELF BACK MATERIAL TO WHITE MELAMINE
                '# UPPER MATCHING INTERIOR #
                UMIBackM = VMATERIAL
                UMIGableM = VMATERIAL
                UMITopLightM = VMATERIAL2
                UMIBtmM = VMATERIAL2
                UMITopBtmM = VMATERIAL2
                UMIAdjShelfM = VMATERIAL
                '# UPPER MICROWAVE MATCHING INTERIOR #
                UMMIBackM = VMATERIAL
                UMMIGableM = VMATERIAL
                UMMITopBtmM = VMATERIAL2
                UMMIMicroSHM = VMATERIAL
                '# UPPER CORNER DIAGONAL MATCHING INTERIOR #
                UCDMILBackM = VMATERIAL
                UCDMISBackM = VMATERIAL
                UCDMIGableM = VMATERIAL
                UCDMITopBtmM = VMATERIAL2
                UCDMIAdjShelfM = VMATERIAL
                '# UPPER WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                UWRBackM = VMATERIAL
                UWRGableM = VMATERIAL
                UWRTopBtmM = VMATERIAL2
                UWRFFSHM = VMATERIAL2
                UWRFShelfM = VMATERIAL2
                UWRDivM = VMATERIAL
                '# UPPER END SHELF #
                UESLGableM = VMATERIAL
                UESSGableM = VMATERIAL
                UESTopM = VMATERIAL2
                UESFixedShelfM = VMATERIAL2
                '# BASE MATCHING INTERIOR #
                BMIBackM = VMATERIAL
                BMIGableM = VMATERIAL
                BMITopStrapM = VMATERIAL2
                BMIBtmM = VMATERIAL2
                BMIAdjShelfM = VMATERIAL
                BMITopM = VMATERIAL2
                BMIStrapM = VMATERIAL2
                BMIDividerM = VMATERIAL2
                '# BASE PENINSULA OPEN #
                BPOPTopBtmFSHM = VMATERIAL2
                BPOPSGableM = VMATERIAL
                BPOPLGableM = VMATERIAL
                '# VANITY MATCHING INTERIOR #
                VMIBackM = VMATERIAL
                VMIGableM = VMATERIAL
                VMITopStrapM = VMATERIAL2
                VMITopM = VMATERIAL2
                VMIBtmM = VMATERIAL2
                VMIFFShelfM = VMATERIAL2
                VMIStrapM = VMATERIAL2
                VMIDividerM = VMATERIAL2
                '# VANITY ELEVATED MATCHING INTERIOR #
                VEMIBackM = VMATERIAL
                VEMIGableM = VMATERIAL
                VEMITopStrapM = VMATERIAL2
                VEMITopM = VMATERIAL2
                VEMIBtmM = VMATERIAL2
                VEMIFFShelfM = VMATERIAL2
                VEMIStrapM = VMATERIAL2
                VEMIDividerM = VMATERIAL2
                '# BASE MICROWAVE OPEN SHELF 1 DRAWER #
                BMO1DUBackM = VMATERIAL
                BMO1DUGableM = VMATERIAL
                BMO1DUTopM = VMATERIAL2
                BMO1DUBtmM = VMATERIAL2
                '# BASE WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                BWRBackM = VMATERIAL
                BWRGableM = VMATERIAL
                BWRTopBtmM = VMATERIAL2
                BWRFFSHM = VMATERIAL2
                BWRValanceM = VSMATERIAL
                BWRVN = "Ordered Joe G."
                '# HUTCH MATCHING INTERIOR MATCHING INTERIOR #
                HMIBackM = VMATERIAL
                HMIGableM = VMATERIAL
                HMITopFFShelfM = VMATERIAL2
                HMIBtmFFShelfM = VMATERIAL2
                HMITopM = VMATERIAL2
                HMIBtmM = VMATERIAL2
                HMIStrapM = VMATERIAL2
                HMIDividerM = VMATERIAL2
                HMIAdjShelfM = VMATERIAL
                '# CANOPY #
                RHFanSHM = "Ply"
                RHFrontM = VMATERIAL
                RHSFrontM = VMATERIAL
                RHSTopM = VMATERIAL
                RHIPanelM = VMATERIAL
                '# OVEN PANEL MATERIAL #
                OV1M = VMATERIAL
                '# FLOATING SHELF MATERIAL #
                FShelfM = VMATERIAL
                '# WINDOW PANEL MATERIAL #
                WPanelM = VMATERIAL
                '# WINDOW BOX MATERIAL #
                WTPanelM = VMATERIAL
                WFrontM = VMATERIAL
                '# DESK SUPPORT GABLE MATERIAL #
                DSGableM = VMATERIAL
                '# FANCY BRACKET MATERIAL #
                FBracketM = VMATERIAL
                '# FANCY VALANCE MATERIAL #
                FValanceM = VMATERIAL
                '# PLUMBING COVER PANEL MATERIAL #
                PCPanelM = VMATERIAL
                '# EDGE CODES VENEER #
                VeneerEdgeCode = VEDGECODE

            Case "OAK"
                Dim VMATERIAL As String = "Oak"
                Dim VSMATERIAL As String = "Solid Oak"
                Dim VEDGECODE As String = "O/V"
                'TALL UTILITY 2 UNITS UPPER
                TU2UUMIBackM = VMATERIAL 'SETS TALL UTILITY UPPER BACK MATERIAL TO WHITE MELAMINE
                TU2UUMIGableM = VMATERIAL 'SETS TALL UTILITY UPPER GABLE MATERIAL TO WHITE MELAMINE
                TU2UUMITopBtmM = VMATERIAL 'SETS TALL UTILITY UPPER TOP + BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMITopM = VMATERIAL 'SETS TALL UTILITY UPPER TOP MATERIAL TO WHITE MELAMINE
                TU2UUMIBotM = VMATERIAL 'SETS TALL UTILITY UPPER BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMIAdjShelfM = VMATERIAL 'SETS TALL UTILITY UPPER ADJUSTABLE SHELF BACK MATERIAL TO WHITE MELAMINE
                '# UPPER MATCHING INTERIOR #
                UMIBackM = VMATERIAL
                UMIGableM = VMATERIAL
                UMITopLightM = VMATERIAL
                UMIBtmM = VMATERIAL
                UMITopBtmM = VMATERIAL
                UMIAdjShelfM = VMATERIAL
                '# UPPER MICROWAVE MATCHING INTERIOR #
                UMMIBackM = VMATERIAL
                UMMIGableM = VMATERIAL
                UMMITopBtmM = VMATERIAL
                UMMIMicroSHM = VMATERIAL
                '# UPPER CORNER DIAGONAL MATCHING INTERIOR #
                UCDMILBackM = VMATERIAL
                UCDMISBackM = VMATERIAL
                UCDMIGableM = VMATERIAL
                UCDMITopBtmM = VMATERIAL
                UCDMIAdjShelfM = VMATERIAL
                '# UPPER WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                UWRBackM = VMATERIAL
                UWRGableM = VMATERIAL
                UWRTopBtmM = VMATERIAL
                UWRFFSHM = VMATERIAL
                UWRFShelfM = VMATERIAL
                UWRDivM = VMATERIAL
                '# UPPER END SHELF #
                UESLGableM = VMATERIAL
                UESSGableM = VMATERIAL
                UESTopM = VMATERIAL
                UESFixedShelfM = VMATERIAL
                '# BASE MATCHING INTERIOR #
                BMIBackM = VMATERIAL
                BMIGableM = VMATERIAL
                BMITopStrapM = VMATERIAL
                BMIBtmM = VMATERIAL
                BMIAdjShelfM = VMATERIAL
                BMITopM = VMATERIAL
                BMIStrapM = VMATERIAL
                BMIDividerM = VMATERIAL
                '# BASE PENINSULA OPEN #
                BPOPTopBtmFSHM = VMATERIAL
                BPOPSGableM = VMATERIAL
                BPOPLGableM = VMATERIAL
                '# VANITY MATCHING INTERIOR #
                VMIBackM = VMATERIAL
                VMIGableM = VMATERIAL
                VMITopStrapM = VMATERIAL
                VMITopM = VMATERIAL
                VMIBtmM = VMATERIAL
                VMIFFShelfM = VMATERIAL
                VMIStrapM = VMATERIAL
                VMIDividerM = VMATERIAL
                '# VANITY ELEVATED MATCHING INTERIOR #
                VEMIBackM = VMATERIAL
                VEMIGableM = VMATERIAL
                VEMITopStrapM = VMATERIAL
                VEMITopM = VMATERIAL
                VEMIBtmM = VMATERIAL
                VEMIFFShelfM = VMATERIAL
                VEMIStrapM = VMATERIAL
                VEMIDividerM = VMATERIAL
                '# BASE MICROWAVE OPEN SHELF 1 DRAWER #
                BMO1DUBackM = VMATERIAL
                BMO1DUGableM = VMATERIAL
                BMO1DUTopM = VMATERIAL
                BMO1DUBtmM = VMATERIAL
                '# BASE WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                BWRBackM = VMATERIAL
                BWRGableM = VMATERIAL
                BWRTopBtmM = VMATERIAL
                BWRFFSHM = VMATERIAL
                BWRValanceM = VSMATERIAL
                BWRVN = "Ordered Joe G."
                '# HUTCH MATCHING INTERIOR MATCHING INTERIOR #
                HMIBackM = VMATERIAL
                HMIGableM = VMATERIAL
                HMITopFFShelfM = VMATERIAL
                HMIBtmFFShelfM = VMATERIAL
                HMITopM = VMATERIAL
                HMIBtmM = VMATERIAL
                HMIStrapM = VMATERIAL
                HMIDividerM = VMATERIAL
                HMIAdjShelfM = VMATERIAL
                '# CANOPY #
                RHFanSHM = "Ply"
                RHFrontM = VMATERIAL
                RHSFrontM = VMATERIAL
                RHSTopM = VMATERIAL
                RHIPanelM = VMATERIAL
                '# OVEN PANEL MATERIAL #
                OV1M = VMATERIAL
                '# FLOATING SHELF MATERIAL #
                FShelfM = VMATERIAL
                '# WINDOW PANEL MATERIAL #
                WPanelM = VMATERIAL
                '# WINDOW BOX MATERIAL #
                WTPanelM = VMATERIAL
                WFrontM = VMATERIAL
                '# DESK SUPPORT GABLE MATERIAL #
                DSGableM = VMATERIAL
                '# FANCY BRACKET MATERIAL #
                FBracketM = VMATERIAL
                '# FANCY VALANCE MATERIAL #
                FValanceM = VMATERIAL
                '# PLUMBING COVER PANEL MATERIAL #
                PCPanelM = VMATERIAL
                '# EDGE CODES VENEER #
                VeneerEdgeCode = VEDGECODE

            Case "PINE"
                Dim VMATERIAL As String = "Pine"
                Dim VSMATERIAL As String = "Solid Pine"
                Dim VEDGECODE As String = "P/V"
                'TALL UTILITY 2 UNITS UPPER
                TU2UUMIBackM = VMATERIAL 'SETS TALL UTILITY UPPER BACK MATERIAL TO WHITE MELAMINE
                TU2UUMIGableM = VMATERIAL 'SETS TALL UTILITY UPPER GABLE MATERIAL TO WHITE MELAMINE
                TU2UUMITopBtmM = VMATERIAL 'SETS TALL UTILITY UPPER TOP + BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMITopM = VMATERIAL 'SETS TALL UTILITY UPPER TOP MATERIAL TO WHITE MELAMINE
                TU2UUMIBotM = VMATERIAL 'SETS TALL UTILITY UPPER BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMIAdjShelfM = VMATERIAL 'SETS TALL UTILITY UPPER ADJUSTABLE SHELF BACK MATERIAL TO WHITE MELAMINE
                '# UPPER MATCHING INTERIOR #
                UMIBackM = VMATERIAL
                UMIGableM = VMATERIAL
                UMITopLightM = VMATERIAL
                UMIBtmM = VMATERIAL
                UMITopBtmM = VMATERIAL
                UMIAdjShelfM = VMATERIAL
                '# UPPER MICROWAVE MATCHING INTERIOR #
                UMMIBackM = VMATERIAL
                UMMIGableM = VMATERIAL
                UMMITopBtmM = VMATERIAL
                UMMIMicroSHM = VMATERIAL
                '# UPPER CORNER DIAGONAL MATCHING INTERIOR #
                UCDMILBackM = VMATERIAL
                UCDMISBackM = VMATERIAL
                UCDMIGableM = VMATERIAL
                UCDMITopBtmM = VMATERIAL
                UCDMIAdjShelfM = VMATERIAL
                '# UPPER WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                UWRBackM = VMATERIAL
                UWRGableM = VMATERIAL
                UWRTopBtmM = VMATERIAL
                UWRFFSHM = VMATERIAL
                UWRFShelfM = VMATERIAL
                UWRDivM = VMATERIAL
                '# UPPER END SHELF #
                UESLGableM = VMATERIAL
                UESSGableM = VMATERIAL
                UESTopM = VMATERIAL
                UESFixedShelfM = VMATERIAL
                '# BASE MATCHING INTERIOR #
                BMIBackM = VMATERIAL
                BMIGableM = VMATERIAL
                BMITopStrapM = VMATERIAL
                BMIBtmM = VMATERIAL
                BMIAdjShelfM = VMATERIAL
                BMITopM = VMATERIAL
                BMIStrapM = VMATERIAL
                BMIDividerM = VMATERIAL
                '# BASE PENINSULA OPEN #
                BPOPTopBtmFSHM = VMATERIAL
                BPOPSGableM = VMATERIAL
                BPOPLGableM = VMATERIAL
                '# VANITY MATCHING INTERIOR #
                VMIBackM = VMATERIAL
                VMIGableM = VMATERIAL
                VMITopStrapM = VMATERIAL
                VMITopM = VMATERIAL
                VMIBtmM = VMATERIAL
                VMIFFShelfM = VMATERIAL
                VMIStrapM = VMATERIAL
                VMIDividerM = VMATERIAL
                '# VANITY ELEVATED MATCHING INTERIOR #
                VEMIBackM = VMATERIAL
                VEMIGableM = VMATERIAL
                VEMITopStrapM = VMATERIAL
                VEMITopM = VMATERIAL
                VEMIBtmM = VMATERIAL
                VEMIFFShelfM = VMATERIAL
                VEMIStrapM = VMATERIAL
                VEMIDividerM = VMATERIAL
                '# BASE MICROWAVE OPEN SHELF 1 DRAWER #
                BMO1DUBackM = VMATERIAL
                BMO1DUGableM = VMATERIAL
                BMO1DUTopM = VMATERIAL
                BMO1DUBtmM = VMATERIAL
                '# BASE WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                BWRBackM = VMATERIAL
                BWRGableM = VMATERIAL
                BWRTopBtmM = VMATERIAL
                BWRFFSHM = VMATERIAL
                BWRValanceM = VSMATERIAL
                BWRVN = "Ordered Joe G."
                '# HUTCH MATCHING INTERIOR MATCHING INTERIOR #
                HMIBackM = VMATERIAL
                HMIGableM = VMATERIAL
                HMITopFFShelfM = VMATERIAL
                HMIBtmFFShelfM = VMATERIAL
                HMITopM = VMATERIAL
                HMIBtmM = VMATERIAL
                HMIStrapM = VMATERIAL
                HMIDividerM = VMATERIAL
                HMIAdjShelfM = VMATERIAL
                '# CANOPY #
                RHFanSHM = "Ply"
                RHFrontM = VMATERIAL
                RHSFrontM = VMATERIAL
                RHSTopM = VMATERIAL
                RHIPanelM = VMATERIAL
                '# OVEN PANEL MATERIAL #
                OV1M = VMATERIAL
                '# FLOATING SHELF MATERIAL #
                FShelfM = VMATERIAL
                '# WINDOW PANEL MATERIAL #
                WPanelM = VMATERIAL
                '# WINDOW BOX MATERIAL #
                WTPanelM = VMATERIAL
                WFrontM = VMATERIAL
                '# DESK SUPPORT GABLE MATERIAL #
                DSGableM = VMATERIAL
                '# FANCY BRACKET MATERIAL #
                FBracketM = VMATERIAL
                '# FANCY VALANCE MATERIAL #
                FValanceM = VMATERIAL
                '# PLUMBING COVER PANEL MATERIAL #
                PCPanelM = VMATERIAL
                '# EDGE CODES VENEER #
                VeneerEdgeCode = VEDGECODE

            Case "WALNUT"
                Dim VMATERIAL As String = "Walnut"
                Dim VSMATERIAL As String = "Solid Walnut"
                Dim VEDGECODE As String = "W/V"
                'TALL UTILITY 2 UNITS UPPER
                TU2UUMIBackM = VMATERIAL 'SETS TALL UTILITY UPPER BACK MATERIAL TO WHITE MELAMINE
                TU2UUMIGableM = VMATERIAL 'SETS TALL UTILITY UPPER GABLE MATERIAL TO WHITE MELAMINE
                TU2UUMITopBtmM = VMATERIAL 'SETS TALL UTILITY UPPER TOP + BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMITopM = VMATERIAL 'SETS TALL UTILITY UPPER TOP MATERIAL TO WHITE MELAMINE
                TU2UUMIBotM = VMATERIAL 'SETS TALL UTILITY UPPER BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMIAdjShelfM = VMATERIAL 'SETS TALL UTILITY UPPER ADJUSTABLE SHELF BACK MATERIAL TO WHITE MELAMINE
                '# UPPER MATCHING INTERIOR #
                UMIBackM = VMATERIAL
                UMIGableM = VMATERIAL
                UMITopLightM = VMATERIAL
                UMIBtmM = VMATERIAL
                UMITopBtmM = VMATERIAL
                UMIAdjShelfM = VMATERIAL
                '# UPPER MICROWAVE MATCHING INTERIOR #
                UMMIBackM = VMATERIAL
                UMMIGableM = VMATERIAL
                UMMITopBtmM = VMATERIAL
                UMMIMicroSHM = VMATERIAL
                '# UPPER CORNER DIAGONAL MATCHING INTERIOR #
                UCDMILBackM = VMATERIAL
                UCDMISBackM = VMATERIAL
                UCDMIGableM = VMATERIAL
                UCDMITopBtmM = VMATERIAL
                UCDMIAdjShelfM = VMATERIAL
                '# UPPER WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                UWRBackM = VMATERIAL
                UWRGableM = VMATERIAL
                UWRTopBtmM = VMATERIAL
                UWRFFSHM = VMATERIAL
                UWRFShelfM = VMATERIAL
                UWRDivM = VMATERIAL
                '# UPPER END SHELF #
                UESLGableM = VMATERIAL
                UESSGableM = VMATERIAL
                UESTopM = VMATERIAL
                UESFixedShelfM = VMATERIAL
                '# BASE MATCHING INTERIOR #
                BMIBackM = VMATERIAL
                BMIGableM = VMATERIAL
                BMITopStrapM = VMATERIAL
                BMIBtmM = VMATERIAL
                BMIAdjShelfM = VMATERIAL
                BMITopM = VMATERIAL
                BMIStrapM = VMATERIAL
                BMIDividerM = VMATERIAL
                '# BASE PENINSULA OPEN #
                BPOPTopBtmFSHM = VMATERIAL
                BPOPSGableM = VMATERIAL
                BPOPLGableM = VMATERIAL
                '# VANITY MATCHING INTERIOR #
                VMIBackM = VMATERIAL
                VMIGableM = VMATERIAL
                VMITopStrapM = VMATERIAL
                VMITopM = VMATERIAL
                VMIBtmM = VMATERIAL
                VMIFFShelfM = VMATERIAL
                VMIStrapM = VMATERIAL
                VMIDividerM = VMATERIAL
                '# VANITY ELEVATED MATCHING INTERIOR #
                VEMIBackM = VMATERIAL
                VEMIGableM = VMATERIAL
                VEMITopStrapM = VMATERIAL
                VEMITopM = VMATERIAL
                VEMIBtmM = VMATERIAL
                VEMIFFShelfM = VMATERIAL
                VEMIStrapM = VMATERIAL
                VEMIDividerM = VMATERIAL
                '# BASE MICROWAVE OPEN SHELF 1 DRAWER #
                BMO1DUBackM = VMATERIAL
                BMO1DUGableM = VMATERIAL
                BMO1DUTopM = VMATERIAL
                BMO1DUBtmM = VMATERIAL
                '# BASE WINE RACK LATTICE, BASE WINE RACK LATTICE HALF, BASE WINE RACK LATTICE HALF VANALACE #
                BWRBackM = VMATERIAL
                BWRGableM = VMATERIAL
                BWRTopBtmM = VMATERIAL
                BWRFFSHM = VMATERIAL
                BWRValanceM = VSMATERIAL
                BWRVN = "Ordered Joe G."
                '# HUTCH MATCHING INTERIOR MATCHING INTERIOR #
                HMIBackM = VMATERIAL
                HMIGableM = VMATERIAL
                HMITopFFShelfM = VMATERIAL
                HMIBtmFFShelfM = VMATERIAL
                HMITopM = VMATERIAL
                HMIBtmM = VMATERIAL
                HMIStrapM = VMATERIAL
                HMIDividerM = VMATERIAL
                HMIAdjShelfM = VMATERIAL
                '# CANOPY #
                RHFanSHM = "Ply"
                RHFrontM = VMATERIAL
                RHSFrontM = VMATERIAL
                RHSTopM = VMATERIAL
                RHIPanelM = VMATERIAL
                '# OVEN PANEL MATERIAL #
                OV1M = VMATERIAL
                '# FLOATING SHELF MATERIAL #
                FShelfM = VMATERIAL
                '# WINDOW PANEL MATERIAL #
                WPanelM = VMATERIAL
                '# WINDOW BOX MATERIAL #
                WTPanelM = VMATERIAL
                WFrontM = VMATERIAL
                '# DESK SUPPORT GABLE MATERIAL #
                DSGableM = VMATERIAL
                '# FANCY BRACKET MATERIAL #
                FBracketM = VMATERIAL
                '# FANCY VALANCE MATERIAL #
                FValanceM = VMATERIAL
                '# PLUMBING COVER PANEL MATERIAL #
                PCPanelM = VMATERIAL
                '# EDGE CODES VENEER #
                VeneerEdgeCode = VEDGECODE
            Case Else
        End Select

        '#######################################################
        '# STANDARD PVC MATERIAL SPECIES AND DOOR FINISH RULES #
        '#######################################################
        '
        Dim PFinish = CKWOPLANNERSDoorFinishBox
        Select Case PFinish
            Case "ANTIQUE WHITE PVC"
                Dim VMATERIAL As String = "Ant.W.Mel"
                Dim VSMATERIAL As String = "Ant.W.PVC"
                Dim VEDGECODE As String = "1438"
                Dim PEDGECODE As String = "1438"
                'TALL UTILITY 2 UNITS UPPER
                TU2UUMIBackM = VMATERIAL 'SETS TALL UTILITY UPPER BACK MATERIAL TO WHITE MELAMINE
                TU2UUMIGableM = VMATERIAL 'SETS TALL UTILITY UPPER GABLE MATERIAL TO WHITE MELAMINE
                TU2UUMITopBtmM = VMATERIAL 'SETS TALL UTILITY UPPER TOP + BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMITopM = VMATERIAL 'SETS TALL UTILITY UPPER TOP MATERIAL TO WHITE MELAMINE
                TU2UUMIBotM = VMATERIAL 'SETS TALL UTILITY UPPER BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMIAdjShelfM = VMATERIAL 'SETS TALL UTILITY UPPER ADJUSTABLE SHELF BACK MATERIAL TO WHITE MELAMINE
                '# UPPER MATCHING INTERIOR #
                UMIBackM = VMATERIAL
                UMIGableM = VMATERIAL
                UMITopLightM = VMATERIAL
                UMIBtmM = VMATERIAL
                UMITopBtmM = VMATERIAL
                UMIAdjShelfM = VMATERIAL
                '# UPPER MICROWAVE MATCHING INTERIOR #
                UMMIBackM = VMATERIAL
                UMMIGableM = VMATERIAL
                UMMITopBtmM = VMATERIAL
                UMMIMicroSHM = VMATERIAL
                '# UPPER CORNER DIAGONAL MATCHING INTERIOR #
                UCDMILBackM = VMATERIAL
                UCDMISBackM = VMATERIAL
                UCDMIGableM = VMATERIAL
                UCDMITopBtmM = VMATERIAL
                UCDMIAdjShelfM = VMATERIAL
                '# UPPER WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                UWRBackM = VMATERIAL
                UWRGableM = VMATERIAL
                UWRTopBtmM = VMATERIAL
                UWRFFSHM = VMATERIAL
                UWRFShelfM = VMATERIAL
                UWRDivM = VMATERIAL
                '# UPPER END SHELF #
                UESLGableM = VMATERIAL
                UESSGableM = VMATERIAL
                UESTopM = VMATERIAL
                UESFixedShelfM = VMATERIAL
                '# BASE MATCHING INTERIOR #
                BMIBackM = VMATERIAL
                BMIGableM = VMATERIAL
                BMITopStrapM = VMATERIAL
                BMIBtmM = VMATERIAL
                BMIAdjShelfM = VMATERIAL
                BMITopM = VMATERIAL
                BMIStrapM = VMATERIAL
                BMIDividerM = VMATERIAL
                '# BASE PENINSULA OPEN #
                BPOPTopBtmFSHM = VMATERIAL
                BPOPSGableM = VMATERIAL
                BPOPLGableM = VMATERIAL
                '# VANITY MATCHING INTERIOR #
                VMIBackM = VMATERIAL
                VMIGableM = VMATERIAL
                VMITopStrapM = VMATERIAL
                VMITopM = VMATERIAL
                VMIBtmM = VMATERIAL
                VMIFFShelfM = VMATERIAL
                VMIStrapM = VMATERIAL
                VMIDividerM = VMATERIAL
                '# VANITY ELEVATED MATCHING INTERIOR #
                VEMIBackM = VMATERIAL
                VEMIGableM = VMATERIAL
                VEMITopStrapM = VMATERIAL
                VEMITopM = VMATERIAL
                VEMIBtmM = VMATERIAL
                VEMIFFShelfM = VMATERIAL
                VEMIStrapM = VMATERIAL
                VEMIDividerM = VMATERIAL
                '# BASE MICROWAVE OPEN SHELF 1 DRAWER #
                BMO1DUBackM = VMATERIAL
                BMO1DUGableM = VMATERIAL
                BMO1DUTopM = VMATERIAL
                BMO1DUBtmM = VMATERIAL
                '# BASE WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                BWRBackM = VMATERIAL
                BWRGableM = VMATERIAL
                BWRTopBtmM = VMATERIAL
                BWRFFSHM = VMATERIAL
                BWRValanceM = VSMATERIAL
                BWRVN = "Ordered Joe G."
                '# HUTCH MATCHING INTERIOR #
                HMIBackM = VMATERIAL
                HMIGableM = VMATERIAL
                HMITopFFShelfM = VMATERIAL
                HMIBtmFFShelfM = VMATERIAL
                HMITopM = VMATERIAL
                HMIBtmM = VMATERIAL
                HMIStrapM = VMATERIAL
                HMIDividerM = VMATERIAL
                HMIAdjShelfM = VMATERIAL
                '# CANOPY #
                RHFanSHM = "Ply"
                RHFrontM = VMATERIAL
                RHSFrontM = VMATERIAL
                RHSTopM = VMATERIAL
                RHIPanelM = VMATERIAL
                '# OVEN PANEL MATERIAL #
                OV1M = VMATERIAL
                '# FLOATING SHELF MATERIAL #
                FShelfM = VMATERIAL
                '# WINDOW PANEL MATERIAL #
                WPanelM = VMATERIAL
                '# WINDOW BOX MATERIAL #
                WTPanelM = VMATERIAL
                WFrontM = VMATERIAL
                '# DESK SUPPORT GABLE MATERIAL #
                DSGableM = VMATERIAL
                '# FANCY BRACKET MATERIAL #
                FBracketM = VMATERIAL
                '# FANCY VALANCE MATERIAL #
                FValanceM = VMATERIAL
                '# PLUMBING COVER PANEL MATERIAL #
                PCPanelM = VMATERIAL
                '# EDGE CODES VENEER AND PVC #
                VeneerEdgeCode = VEDGECODE
                PVCEdgeCode = PEDGECODE

            Case "BLEACHED MAPLE PVC"
                Dim VMATERIAL As String = "Bl.M.Mel"
                Dim VSMATERIAL As String = "Bl.M.PVC"
                Dim VEDGECODE As String = "9530"
                Dim PEDGECODE As String = "9530"
                'TALL UTILITY 2 UNITS UPPER
                TU2UUMIBackM = VMATERIAL 'SETS TALL UTILITY UPPER BACK MATERIAL TO WHITE MELAMINE
                TU2UUMIGableM = VMATERIAL 'SETS TALL UTILITY UPPER GABLE MATERIAL TO WHITE MELAMINE
                TU2UUMITopBtmM = VMATERIAL 'SETS TALL UTILITY UPPER TOP + BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMITopM = VMATERIAL 'SETS TALL UTILITY UPPER TOP MATERIAL TO WHITE MELAMINE
                TU2UUMIBotM = VMATERIAL 'SETS TALL UTILITY UPPER BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMIAdjShelfM = VMATERIAL 'SETS TALL UTILITY UPPER ADJUSTABLE SHELF BACK MATERIAL TO WHITE MELAMINE
                '# UPPER MATCHING INTERIOR #
                UMIBackM = VMATERIAL
                UMIGableM = VMATERIAL
                UMITopLightM = VMATERIAL
                UMIBtmM = VMATERIAL
                UMITopBtmM = VMATERIAL
                UMIAdjShelfM = VMATERIAL
                '# UPPER MICROWAVE MATCHING INTERIOR #
                UMMIBackM = VMATERIAL
                UMMIGableM = VMATERIAL
                UMMITopBtmM = VMATERIAL
                UMMIMicroSHM = VMATERIAL
                '# UPPER CORNER DIAGONAL MATCHING INTERIOR #
                UCDMILBackM = VMATERIAL
                UCDMISBackM = VMATERIAL
                UCDMIGableM = VMATERIAL
                UCDMITopBtmM = VMATERIAL
                UCDMIAdjShelfM = VMATERIAL
                '# UPPER WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                UWRBackM = VMATERIAL
                UWRGableM = VMATERIAL
                UWRTopBtmM = VMATERIAL
                UWRFFSHM = VMATERIAL
                UWRFShelfM = VMATERIAL
                UWRDivM = VMATERIAL
                '# UPPER END SHELF #
                UESLGableM = VMATERIAL
                UESSGableM = VMATERIAL
                UESTopM = VMATERIAL
                UESFixedShelfM = VMATERIAL
                '# BASE MATCHING INTERIOR #
                BMIBackM = VMATERIAL
                BMIGableM = VMATERIAL
                BMITopStrapM = VMATERIAL
                BMIBtmM = VMATERIAL
                BMIAdjShelfM = VMATERIAL
                BMITopM = VMATERIAL
                BMIStrapM = VMATERIAL
                BMIDividerM = VMATERIAL
                '# BASE PENINSULA OPEN #
                BPOPTopBtmFSHM = VMATERIAL
                BPOPSGableM = VMATERIAL
                BPOPLGableM = VMATERIAL
                '# VANITY MATCHING INTERIOR #
                VMIBackM = VMATERIAL
                VMIGableM = VMATERIAL
                VMITopStrapM = VMATERIAL
                VMITopM = VMATERIAL
                VMIBtmM = VMATERIAL
                VMIFFShelfM = VMATERIAL
                VMIStrapM = VMATERIAL
                VMIDividerM = VMATERIAL
                '# VANITY ELEVATED MATCHING INTERIOR #
                VEMIBackM = VMATERIAL
                VEMIGableM = VMATERIAL
                VEMITopStrapM = VMATERIAL
                VEMITopM = VMATERIAL
                VEMIBtmM = VMATERIAL
                VEMIFFShelfM = VMATERIAL
                VEMIStrapM = VMATERIAL
                VEMIDividerM = VMATERIAL
                '# BASE MICROWAVE OPEN SHELF 1 DRAWER #
                BMO1DUBackM = VMATERIAL
                BMO1DUGableM = VMATERIAL
                BMO1DUTopM = VMATERIAL
                BMO1DUBtmM = VMATERIAL
                '# BASE WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                BWRBackM = VMATERIAL
                BWRGableM = VMATERIAL
                BWRTopBtmM = VMATERIAL
                BWRFFSHM = VMATERIAL
                BWRValanceM = VSMATERIAL
                BWRVN = "Ordered Joe G."
                '# HUTCH MATCHING INTERIOR #
                HMIBackM = VMATERIAL
                HMIGableM = VMATERIAL
                HMITopFFShelfM = VMATERIAL
                HMIBtmFFShelfM = VMATERIAL
                HMITopM = VMATERIAL
                HMIBtmM = VMATERIAL
                HMIStrapM = VMATERIAL
                HMIDividerM = VMATERIAL
                HMIAdjShelfM = VMATERIAL
                '# CANOPY #
                RHFanSHM = "Ply"
                RHFrontM = VMATERIAL
                RHSFrontM = VMATERIAL
                RHSTopM = VMATERIAL
                RHIPanelM = VMATERIAL
                '# OVEN PANEL MATERIAL #
                OV1M = VMATERIAL
                '# FLOATING SHELF MATERIAL #
                FShelfM = VMATERIAL
                '# WINDOW PANEL MATERIAL #
                WPanelM = VMATERIAL
                '# WINDOW BOX MATERIAL #
                WTPanelM = VMATERIAL
                WFrontM = VMATERIAL
                '# DESK SUPPORT GABLE MATERIAL #
                DSGableM = VMATERIAL
                '# FANCY BRACKET MATERIAL #
                FBracketM = VMATERIAL
                '# FANCY VALANCE MATERIAL #
                FValanceM = VMATERIAL
                '# PLUMBING COVER PANEL MATERIAL #
                PCPanelM = VMATERIAL
                '# EDGE CODES VENEER AND PVC #
                VeneerEdgeCode = VEDGECODE
                PVCEdgeCode = PEDGECODE

            Case "CHARCOAL MELAMINE"
                Dim VMATERIAL As String = "Char.Mel"
                Dim VSMATERIAL As String = "Char.Mel"
                Dim VEDGECODE As String = "1315"
                Dim PEDGECODE As String = "1315"
                'TALL UTILITY 2 UNITS UPPER
                TU2UUMIBackM = VMATERIAL 'SETS TALL UTILITY UPPER BACK MATERIAL TO WHITE MELAMINE
                TU2UUMIGableM = VMATERIAL 'SETS TALL UTILITY UPPER GABLE MATERIAL TO WHITE MELAMINE
                TU2UUMITopBtmM = VMATERIAL 'SETS TALL UTILITY UPPER TOP + BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMITopM = VMATERIAL 'SETS TALL UTILITY UPPER TOP MATERIAL TO WHITE MELAMINE
                TU2UUMIBotM = VMATERIAL 'SETS TALL UTILITY UPPER BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMIAdjShelfM = VMATERIAL 'SETS TALL UTILITY UPPER ADJUSTABLE SHELF BACK MATERIAL TO WHITE MELAMINE
                '# UPPER MATCHING INTERIOR #
                UMIBackM = VMATERIAL
                UMIGableM = VMATERIAL
                UMITopLightM = VMATERIAL
                UMIBtmM = VMATERIAL
                UMITopBtmM = VMATERIAL
                UMIAdjShelfM = VMATERIAL
                '# UPPER MICROWAVE MATCHING INTERIOR #
                UMMIBackM = VMATERIAL
                UMMIGableM = VMATERIAL
                UMMITopBtmM = VMATERIAL
                UMMIMicroSHM = VMATERIAL
                '# UPPER CORNER DIAGONAL MATCHING INTERIOR #
                UCDMILBackM = VMATERIAL
                UCDMISBackM = VMATERIAL
                UCDMIGableM = VMATERIAL
                UCDMITopBtmM = VMATERIAL
                UCDMIAdjShelfM = VMATERIAL
                '# UPPER WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                UWRBackM = VMATERIAL
                UWRGableM = VMATERIAL
                UWRTopBtmM = VMATERIAL
                UWRFFSHM = VMATERIAL
                UWRFShelfM = VMATERIAL
                UWRDivM = VMATERIAL
                '# UPPER END SHELF #
                UESLGableM = VMATERIAL
                UESSGableM = VMATERIAL
                UESTopM = VMATERIAL
                UESFixedShelfM = VMATERIAL
                '# BASE MATCHING INTERIOR #
                BMIBackM = VMATERIAL
                BMIGableM = VMATERIAL
                BMITopStrapM = VMATERIAL
                BMIBtmM = VMATERIAL
                BMIAdjShelfM = VMATERIAL
                BMITopM = VMATERIAL
                BMIStrapM = VMATERIAL
                BMIDividerM = VMATERIAL
                '# BASE PENINSULA OPEN #
                BPOPTopBtmFSHM = VMATERIAL
                BPOPSGableM = VMATERIAL
                BPOPLGableM = VMATERIAL
                '# VANITY MATCHING INTERIOR #
                VMIBackM = VMATERIAL
                VMIGableM = VMATERIAL
                VMITopStrapM = VMATERIAL
                VMITopM = VMATERIAL
                VMIBtmM = VMATERIAL
                VMIFFShelfM = VMATERIAL
                VMIStrapM = VMATERIAL
                VMIDividerM = VMATERIAL
                '# VANITY ELEVATED MATCHING INTERIOR #
                VEMIBackM = VMATERIAL
                VEMIGableM = VMATERIAL
                VEMITopStrapM = VMATERIAL
                VEMITopM = VMATERIAL
                VEMIBtmM = VMATERIAL
                VEMIFFShelfM = VMATERIAL
                VEMIStrapM = VMATERIAL
                VEMIDividerM = VMATERIAL
                '# BASE MICROWAVE OPEN SHELF 1 DRAWER #
                BMO1DUBackM = VMATERIAL
                BMO1DUGableM = VMATERIAL
                BMO1DUTopM = VMATERIAL
                BMO1DUBtmM = VMATERIAL
                '# BASE WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                BWRBackM = VMATERIAL
                BWRGableM = VMATERIAL
                BWRTopBtmM = VMATERIAL
                BWRFFSHM = VMATERIAL
                BWRValanceM = VSMATERIAL
                BWRVN = "Ordered Joe G."
                '# HUTCH MATCHING INTERIOR #
                HMIBackM = VMATERIAL
                HMIGableM = VMATERIAL
                HMITopFFShelfM = VMATERIAL
                HMIBtmFFShelfM = VMATERIAL
                HMITopM = VMATERIAL
                HMIBtmM = VMATERIAL
                HMIStrapM = VMATERIAL
                HMIDividerM = VMATERIAL
                HMIAdjShelfM = VMATERIAL
                '# CANOPY #
                RHFanSHM = "Ply"
                RHFrontM = VMATERIAL
                RHSFrontM = VMATERIAL
                RHSTopM = VMATERIAL
                RHIPanelM = VMATERIAL
                '# OVEN PANEL MATERIAL #
                OV1M = VMATERIAL
                '# FLOATING SHELF MATERIAL #
                FShelfM = VMATERIAL
                '# WINDOW PANEL MATERIAL #
                WPanelM = VMATERIAL
                '# WINDOW BOX MATERIAL #
                WTPanelM = VMATERIAL
                WFrontM = VMATERIAL
                '# DESK SUPPORT GABLE MATERIAL #
                DSGableM = VMATERIAL
                '# FANCY BRACKET MATERIAL #
                FBracketM = VMATERIAL
                '# FANCY VALANCE MATERIAL #
                FValanceM = VMATERIAL
                '# PLUMBING COVER PANEL MATERIAL #
                PCPanelM = VMATERIAL
                '# EDGE CODES VENEER AND PVC #
                VeneerEdgeCode = VEDGECODE
                PVCEdgeCode = PEDGECODE

            Case "CHOCOLATE MAPLE PVC"
                Dim VMATERIAL As String = "Choc.Mel"
                Dim VSMATERIAL As String = "Choc.PVC"
                Dim VEDGECODE As String = "6121"
                Dim PEDGECODE As String = "6121"
                'TALL UTILITY 2 UNITS UPPER
                TU2UUMIBackM = VMATERIAL 'SETS TALL UTILITY UPPER BACK MATERIAL TO WHITE MELAMINE
                TU2UUMIGableM = VMATERIAL 'SETS TALL UTILITY UPPER GABLE MATERIAL TO WHITE MELAMINE
                TU2UUMITopBtmM = VMATERIAL 'SETS TALL UTILITY UPPER TOP + BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMITopM = VMATERIAL 'SETS TALL UTILITY UPPER TOP MATERIAL TO WHITE MELAMINE
                TU2UUMIBotM = VMATERIAL 'SETS TALL UTILITY UPPER BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMIAdjShelfM = VMATERIAL 'SETS TALL UTILITY UPPER ADJUSTABLE SHELF BACK MATERIAL TO WHITE MELAMINE
                '# UPPER MATCHING INTERIOR #
                UMIBackM = VMATERIAL
                UMIGableM = VMATERIAL
                UMITopLightM = VMATERIAL
                UMIBtmM = VMATERIAL
                UMITopBtmM = VMATERIAL
                UMIAdjShelfM = VMATERIAL
                '# UPPER MICROWAVE MATCHING INTERIOR #
                UMMIBackM = VMATERIAL
                UMMIGableM = VMATERIAL
                UMMITopBtmM = VMATERIAL
                UMMIMicroSHM = VMATERIAL
                '# UPPER CORNER DIAGONAL MATCHING INTERIOR #
                UCDMILBackM = VMATERIAL
                UCDMISBackM = VMATERIAL
                UCDMIGableM = VMATERIAL
                UCDMITopBtmM = VMATERIAL
                UCDMIAdjShelfM = VMATERIAL
                '# UPPER WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                UWRBackM = VMATERIAL
                UWRGableM = VMATERIAL
                UWRTopBtmM = VMATERIAL
                UWRFFSHM = VMATERIAL
                UWRFShelfM = VMATERIAL
                UWRDivM = VMATERIAL
                '# UPPER END SHELF #
                UESLGableM = VMATERIAL
                UESSGableM = VMATERIAL
                UESTopM = VMATERIAL
                UESFixedShelfM = VMATERIAL
                '# BASE MATCHING INTERIOR #
                BMIBackM = VMATERIAL
                BMIGableM = VMATERIAL
                BMITopStrapM = VMATERIAL
                BMIBtmM = VMATERIAL
                BMIAdjShelfM = VMATERIAL
                BMITopM = VMATERIAL
                BMIStrapM = VMATERIAL
                BMIDividerM = VMATERIAL
                '# BASE PENINSULA OPEN #
                BPOPTopBtmFSHM = VMATERIAL
                BPOPSGableM = VMATERIAL
                BPOPLGableM = VMATERIAL
                '# VANITY MATCHING INTERIOR #
                VMIBackM = VMATERIAL
                VMIGableM = VMATERIAL
                VMITopStrapM = VMATERIAL
                VMITopM = VMATERIAL
                VMIBtmM = VMATERIAL
                VMIFFShelfM = VMATERIAL
                VMIStrapM = VMATERIAL
                VMIDividerM = VMATERIAL
                '# VANITY ELEVATED MATCHING INTERIOR #
                VEMIBackM = VMATERIAL
                VEMIGableM = VMATERIAL
                VEMITopStrapM = VMATERIAL
                VEMITopM = VMATERIAL
                VEMIBtmM = VMATERIAL
                VEMIFFShelfM = VMATERIAL
                VEMIStrapM = VMATERIAL
                VEMIDividerM = VMATERIAL
                '# BASE MICROWAVE OPEN SHELF 1 DRAWER #
                BMO1DUBackM = VMATERIAL
                BMO1DUGableM = VMATERIAL
                BMO1DUTopM = VMATERIAL
                BMO1DUBtmM = VMATERIAL
                '# BASE WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                BWRBackM = VMATERIAL
                BWRGableM = VMATERIAL
                BWRTopBtmM = VMATERIAL
                BWRFFSHM = VMATERIAL
                BWRValanceM = VSMATERIAL
                BWRVN = "Ordered Joe G."
                '# HUTCH MATCHING INTERIOR #
                HMIBackM = VMATERIAL
                HMIGableM = VMATERIAL
                HMITopFFShelfM = VMATERIAL
                HMIBtmFFShelfM = VMATERIAL
                HMITopM = VMATERIAL
                HMIBtmM = VMATERIAL
                HMIStrapM = VMATERIAL
                HMIDividerM = VMATERIAL
                HMIAdjShelfM = VMATERIAL
                '# CANOPY #
                RHFanSHM = "Ply"
                RHFrontM = VMATERIAL
                RHSFrontM = VMATERIAL
                RHSTopM = VMATERIAL
                RHIPanelM = VMATERIAL
                '# OVEN PANEL MATERIAL #
                OV1M = VMATERIAL
                '# FLOATING SHELF MATERIAL #
                FShelfM = VMATERIAL
                '# WINDOW PANEL MATERIAL #
                WPanelM = VMATERIAL
                '# WINDOW BOX MATERIAL #
                WTPanelM = VMATERIAL
                WFrontM = VMATERIAL
                '# DESK SUPPORT GABLE MATERIAL #
                DSGableM = VMATERIAL
                '# FANCY BRACKET MATERIAL #
                FBracketM = VMATERIAL
                '# FANCY VALANCE MATERIAL #
                FValanceM = VMATERIAL
                '# PLUMBING COVER PANEL MATERIAL #
                PCPanelM = VMATERIAL
                '# EDGE CODES VENEER AND PVC #
                VeneerEdgeCode = VEDGECODE
                PVCEdgeCode = PEDGECODE

            Case "HONEY APPLE PVC"
                Dim VMATERIAL As String = "H.A.Mel"
                Dim VSMATERIAL As String = "H.A.PVC"
                Dim VEDGECODE As String = "7216"
                Dim PEDGECODE As String = "7216"
                'TALL UTILITY 2 UNITS UPPER
                TU2UUMIBackM = VMATERIAL 'SETS TALL UTILITY UPPER BACK MATERIAL TO WHITE MELAMINE
                TU2UUMIGableM = VMATERIAL 'SETS TALL UTILITY UPPER GABLE MATERIAL TO WHITE MELAMINE
                TU2UUMITopBtmM = VMATERIAL 'SETS TALL UTILITY UPPER TOP + BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMITopM = VMATERIAL 'SETS TALL UTILITY UPPER TOP MATERIAL TO WHITE MELAMINE
                TU2UUMIBotM = VMATERIAL 'SETS TALL UTILITY UPPER BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMIAdjShelfM = VMATERIAL 'SETS TALL UTILITY UPPER ADJUSTABLE SHELF BACK MATERIAL TO WHITE MELAMINE
                '# UPPER MATCHING INTERIOR #
                UMIBackM = VMATERIAL
                UMIGableM = VMATERIAL
                UMITopLightM = VMATERIAL
                UMIBtmM = VMATERIAL
                UMITopBtmM = VMATERIAL
                UMIAdjShelfM = VMATERIAL
                '# UPPER MICROWAVE MATCHING INTERIOR #
                UMMIBackM = VMATERIAL
                UMMIGableM = VMATERIAL
                UMMITopBtmM = VMATERIAL
                UMMIMicroSHM = VMATERIAL
                '# UPPER CORNER DIAGONAL MATCHING INTERIOR #
                UCDMILBackM = VMATERIAL
                UCDMISBackM = VMATERIAL
                UCDMIGableM = VMATERIAL
                UCDMITopBtmM = VMATERIAL
                UCDMIAdjShelfM = VMATERIAL
                '# UPPER WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                UWRBackM = VMATERIAL
                UWRGableM = VMATERIAL
                UWRTopBtmM = VMATERIAL
                UWRFFSHM = VMATERIAL
                UWRFShelfM = VMATERIAL
                UWRDivM = VMATERIAL
                '# UPPER END SHELF #
                UESLGableM = VMATERIAL
                UESSGableM = VMATERIAL
                UESTopM = VMATERIAL
                UESFixedShelfM = VMATERIAL
                '# BASE MATCHING INTERIOR #
                BMIBackM = VMATERIAL
                BMIGableM = VMATERIAL
                BMITopStrapM = VMATERIAL
                BMIBtmM = VMATERIAL
                BMIAdjShelfM = VMATERIAL
                BMITopM = VMATERIAL
                BMIStrapM = VMATERIAL
                BMIDividerM = VMATERIAL
                '# BASE PENINSULA OPEN #
                BPOPTopBtmFSHM = VMATERIAL
                BPOPSGableM = VMATERIAL
                BPOPLGableM = VMATERIAL
                '# VANITY MATCHING INTERIOR #
                VMIBackM = VMATERIAL
                VMIGableM = VMATERIAL
                VMITopStrapM = VMATERIAL
                VMITopM = VMATERIAL
                VMIBtmM = VMATERIAL
                VMIFFShelfM = VMATERIAL
                VMIStrapM = VMATERIAL
                VMIDividerM = VMATERIAL
                '# VANITY ELEVATED MATCHING INTERIOR #
                VEMIBackM = VMATERIAL
                VEMIGableM = VMATERIAL
                VEMITopStrapM = VMATERIAL
                VEMITopM = VMATERIAL
                VEMIBtmM = VMATERIAL
                VEMIFFShelfM = VMATERIAL
                VEMIStrapM = VMATERIAL
                VEMIDividerM = VMATERIAL
                '# BASE MICROWAVE OPEN SHELF 1 DRAWER #
                BMO1DUBackM = VMATERIAL
                BMO1DUGableM = VMATERIAL
                BMO1DUTopM = VMATERIAL
                BMO1DUBtmM = VMATERIAL
                '# BASE WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                BWRBackM = VMATERIAL
                BWRGableM = VMATERIAL
                BWRTopBtmM = VMATERIAL
                BWRFFSHM = VMATERIAL
                BWRValanceM = VSMATERIAL
                BWRVN = "Ordered Joe G."
                '# HUTCH MATCHING INTERIOR #
                HMIBackM = VMATERIAL
                HMIGableM = VMATERIAL
                HMITopFFShelfM = VMATERIAL
                HMIBtmFFShelfM = VMATERIAL
                HMITopM = VMATERIAL
                HMIBtmM = VMATERIAL
                HMIStrapM = VMATERIAL
                HMIDividerM = VMATERIAL
                HMIAdjShelfM = VMATERIAL
                '# CANOPY #
                RHFanSHM = "Ply"
                RHFrontM = VMATERIAL
                RHSFrontM = VMATERIAL
                RHSTopM = VMATERIAL
                RHIPanelM = VMATERIAL
                '# OVEN PANEL MATERIAL #
                OV1M = VMATERIAL
                '# FLOATING SHELF MATERIAL #
                FShelfM = VMATERIAL
                '# WINDOW PANEL MATERIAL #
                WPanelM = VMATERIAL
                '# WINDOW BOX MATERIAL #
                WTPanelM = VMATERIAL
                WFrontM = VMATERIAL
                '# DESK SUPPORT GABLE MATERIAL #
                DSGableM = VMATERIAL
                '# FANCY BRACKET MATERIAL #
                FBracketM = VMATERIAL
                '# FANCY VALANCE MATERIAL #
                FValanceM = VMATERIAL
                '# PLUMBING COVER PANEL MATERIAL #
                PCPanelM = VMATERIAL
                '# EDGE CODES VENEER AND PVC #
                VeneerEdgeCode = VEDGECODE
                PVCEdgeCode = PEDGECODE

            Case "ITALIAN WALNUT PVC"
                Dim VMATERIAL As String = "Itl.W.Mel"
                Dim VSMATERIAL As String = "Itl.W.PVC"
                Dim VEDGECODE As String = "5233"
                Dim PEDGECODE As String = "5233"
                'TALL UTILITY 2 UNITS UPPER
                TU2UUMIBackM = VMATERIAL 'SETS TALL UTILITY UPPER BACK MATERIAL TO WHITE MELAMINE
                TU2UUMIGableM = VMATERIAL 'SETS TALL UTILITY UPPER GABLE MATERIAL TO WHITE MELAMINE
                TU2UUMITopBtmM = VMATERIAL 'SETS TALL UTILITY UPPER TOP + BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMITopM = VMATERIAL 'SETS TALL UTILITY UPPER TOP MATERIAL TO WHITE MELAMINE
                TU2UUMIBotM = VMATERIAL 'SETS TALL UTILITY UPPER BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMIAdjShelfM = VMATERIAL 'SETS TALL UTILITY UPPER ADJUSTABLE SHELF BACK MATERIAL TO WHITE MELAMINE
                '# UPPER MATCHING INTERIOR #
                UMIBackM = VMATERIAL
                UMIGableM = VMATERIAL
                UMITopLightM = VMATERIAL
                UMIBtmM = VMATERIAL
                UMITopBtmM = VMATERIAL
                UMIAdjShelfM = VMATERIAL
                '# UPPER MICROWAVE MATCHING INTERIOR #
                UMMIBackM = VMATERIAL
                UMMIGableM = VMATERIAL
                UMMITopBtmM = VMATERIAL
                UMMIMicroSHM = VMATERIAL
                '# UPPER CORNER DIAGONAL MATCHING INTERIOR #
                UCDMILBackM = VMATERIAL
                UCDMISBackM = VMATERIAL
                UCDMIGableM = VMATERIAL
                UCDMITopBtmM = VMATERIAL
                UCDMIAdjShelfM = VMATERIAL
                '# UPPER WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                UWRBackM = VMATERIAL
                UWRGableM = VMATERIAL
                UWRTopBtmM = VMATERIAL
                UWRFFSHM = VMATERIAL
                UWRFShelfM = VMATERIAL
                UWRDivM = VMATERIAL
                '# UPPER END SHELF #
                UESLGableM = VMATERIAL
                UESSGableM = VMATERIAL
                UESTopM = VMATERIAL
                UESFixedShelfM = VMATERIAL
                '# BASE MATCHING INTERIOR #
                BMIBackM = VMATERIAL
                BMIGableM = VMATERIAL
                BMITopStrapM = VMATERIAL
                BMIBtmM = VMATERIAL
                BMIAdjShelfM = VMATERIAL
                BMITopM = VMATERIAL
                BMIStrapM = VMATERIAL
                BMIDividerM = VMATERIAL
                '# BASE PENINSULA OPEN #
                BPOPTopBtmFSHM = VMATERIAL
                BPOPSGableM = VMATERIAL
                BPOPLGableM = VMATERIAL
                '# VANITY MATCHING INTERIOR #
                VMIBackM = VMATERIAL
                VMIGableM = VMATERIAL
                VMITopStrapM = VMATERIAL
                VMITopM = VMATERIAL
                VMIBtmM = VMATERIAL
                VMIFFShelfM = VMATERIAL
                VMIStrapM = VMATERIAL
                VMIDividerM = VMATERIAL
                '# VANITY ELEVATED MATCHING INTERIOR #
                VEMIBackM = VMATERIAL
                VEMIGableM = VMATERIAL
                VEMITopStrapM = VMATERIAL
                VEMITopM = VMATERIAL
                VEMIBtmM = VMATERIAL
                VEMIFFShelfM = VMATERIAL
                VEMIStrapM = VMATERIAL
                VEMIDividerM = VMATERIAL
                '# BASE MICROWAVE OPEN SHELF 1 DRAWER #
                BMO1DUBackM = VMATERIAL
                BMO1DUGableM = VMATERIAL
                BMO1DUTopM = VMATERIAL
                BMO1DUBtmM = VMATERIAL
                '# BASE WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                BWRBackM = VMATERIAL
                BWRGableM = VMATERIAL
                BWRTopBtmM = VMATERIAL
                BWRFFSHM = VMATERIAL
                BWRValanceM = VSMATERIAL
                BWRVN = "Ordered Joe G."
                '# HUTCH MATCHING INTERIOR #
                HMIBackM = VMATERIAL
                HMIGableM = VMATERIAL
                HMITopFFShelfM = VMATERIAL
                HMIBtmFFShelfM = VMATERIAL
                HMITopM = VMATERIAL
                HMIBtmM = VMATERIAL
                HMIStrapM = VMATERIAL
                HMIDividerM = VMATERIAL
                HMIAdjShelfM = VMATERIAL
                '# CANOPY #
                RHFanSHM = "Ply"
                RHFrontM = VMATERIAL
                RHSFrontM = VMATERIAL
                RHSTopM = VMATERIAL
                RHIPanelM = VMATERIAL
                '# OVEN PANEL MATERIAL #
                OV1M = VMATERIAL
                '# FLOATING SHELF MATERIAL #
                FShelfM = VMATERIAL
                '# WINDOW PANEL MATERIAL #
                WPanelM = VMATERIAL
                '# WINDOW BOX MATERIAL #
                WTPanelM = VMATERIAL
                WFrontM = VMATERIAL
                '# DESK SUPPORT GABLE MATERIAL #
                DSGableM = VMATERIAL
                '# FANCY BRACKET MATERIAL #
                FBracketM = VMATERIAL
                '# FANCY VALANCE MATERIAL #
                FValanceM = VMATERIAL
                '# PLUMBING COVER PANEL MATERIAL #
                PCPanelM = VMATERIAL
                '# EDGE CODES VENEER AND PVC #
                VeneerEdgeCode = VEDGECODE
                PVCEdgeCode = PEDGECODE

            Case "JAVA GLOW PVC"
                Dim VMATERIAL As String = "Java G.Mel"
                Dim VSMATERIAL As String = "Java G.PVC"
                Dim VEDGECODE As String = "9513"
                Dim PEDGECODE As String = "9513"
                'TALL UTILITY 2 UNITS UPPER
                TU2UUMIBackM = VMATERIAL 'SETS TALL UTILITY UPPER BACK MATERIAL TO WHITE MELAMINE
                TU2UUMIGableM = VMATERIAL 'SETS TALL UTILITY UPPER GABLE MATERIAL TO WHITE MELAMINE
                TU2UUMITopBtmM = VMATERIAL 'SETS TALL UTILITY UPPER TOP + BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMITopM = VMATERIAL 'SETS TALL UTILITY UPPER TOP MATERIAL TO WHITE MELAMINE
                TU2UUMIBotM = VMATERIAL 'SETS TALL UTILITY UPPER BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMIAdjShelfM = VMATERIAL 'SETS TALL UTILITY UPPER ADJUSTABLE SHELF BACK MATERIAL TO WHITE MELAMINE
                '# UPPER MATCHING INTERIOR #
                UMIBackM = VMATERIAL
                UMIGableM = VMATERIAL
                UMITopLightM = VMATERIAL
                UMIBtmM = VMATERIAL
                UMITopBtmM = VMATERIAL
                UMIAdjShelfM = VMATERIAL
                '# UPPER MICROWAVE MATCHING INTERIOR #
                UMMIBackM = VMATERIAL
                UMMIGableM = VMATERIAL
                UMMITopBtmM = VMATERIAL
                UMMIMicroSHM = VMATERIAL
                '# UPPER CORNER DIAGONAL MATCHING INTERIOR #
                UCDMILBackM = VMATERIAL
                UCDMISBackM = VMATERIAL
                UCDMIGableM = VMATERIAL
                UCDMITopBtmM = VMATERIAL
                UCDMIAdjShelfM = VMATERIAL
                '# UPPER WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                UWRBackM = VMATERIAL
                UWRGableM = VMATERIAL
                UWRTopBtmM = VMATERIAL
                UWRFFSHM = VMATERIAL
                UWRFShelfM = VMATERIAL
                UWRDivM = VMATERIAL
                '# UPPER END SHELF #
                UESLGableM = VMATERIAL
                UESSGableM = VMATERIAL
                UESTopM = VMATERIAL
                UESFixedShelfM = VMATERIAL
                '# BASE MATCHING INTERIOR #
                BMIBackM = VMATERIAL
                BMIGableM = VMATERIAL
                BMITopStrapM = VMATERIAL
                BMIBtmM = VMATERIAL
                BMIAdjShelfM = VMATERIAL
                BMITopM = VMATERIAL
                BMIStrapM = VMATERIAL
                BMIDividerM = VMATERIAL
                '# BASE PENINSULA OPEN #
                BPOPTopBtmFSHM = VMATERIAL
                BPOPSGableM = VMATERIAL
                BPOPLGableM = VMATERIAL
                '# VANITY MATCHING INTERIOR #
                VMIBackM = VMATERIAL
                VMIGableM = VMATERIAL
                VMITopStrapM = VMATERIAL
                VMITopM = VMATERIAL
                VMIBtmM = VMATERIAL
                VMIFFShelfM = VMATERIAL
                VMIStrapM = VMATERIAL
                VMIDividerM = VMATERIAL
                '# VANITY ELEVATED MATCHING INTERIOR #
                VEMIBackM = VMATERIAL
                VEMIGableM = VMATERIAL
                VEMITopStrapM = VMATERIAL
                VEMITopM = VMATERIAL
                VEMIBtmM = VMATERIAL
                VEMIFFShelfM = VMATERIAL
                VEMIStrapM = VMATERIAL
                VEMIDividerM = VMATERIAL
                '# BASE MICROWAVE OPEN SHELF 1 DRAWER #
                BMO1DUBackM = VMATERIAL
                BMO1DUGableM = VMATERIAL
                BMO1DUTopM = VMATERIAL
                BMO1DUBtmM = VMATERIAL
                '# BASE WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                BWRBackM = VMATERIAL
                BWRGableM = VMATERIAL
                BWRTopBtmM = VMATERIAL
                BWRFFSHM = VMATERIAL
                BWRValanceM = VSMATERIAL
                BWRVN = "Ordered Joe G."
                '# HUTCH MATCHING INTERIOR #
                HMIBackM = VMATERIAL
                HMIGableM = VMATERIAL
                HMITopFFShelfM = VMATERIAL
                HMIBtmFFShelfM = VMATERIAL
                HMITopM = VMATERIAL
                HMIBtmM = VMATERIAL
                HMIStrapM = VMATERIAL
                HMIDividerM = VMATERIAL
                HMIAdjShelfM = VMATERIAL
                '# CANOPY #
                RHFanSHM = "Ply"
                RHFrontM = VMATERIAL
                RHSFrontM = VMATERIAL
                RHSTopM = VMATERIAL
                RHIPanelM = VMATERIAL
                '# OVEN PANEL MATERIAL #
                OV1M = VMATERIAL
                '# FLOATING SHELF MATERIAL #
                FShelfM = VMATERIAL
                '# WINDOW PANEL MATERIAL #
                WPanelM = VMATERIAL
                '# WINDOW BOX MATERIAL #
                WTPanelM = VMATERIAL
                WFrontM = VMATERIAL
                '# DESK SUPPORT GABLE MATERIAL #
                DSGableM = VMATERIAL
                '# FANCY BRACKET MATERIAL #
                FBracketM = VMATERIAL
                '# FANCY VALANCE MATERIAL #
                FValanceM = VMATERIAL
                '# PLUMBING COVER PANEL MATERIAL #
                PCPanelM = VMATERIAL
                '# EDGE CODES VENEER AND PVC #
                VeneerEdgeCode = VEDGECODE
                PVCEdgeCode = PEDGECODE

            Case "MAJESTIC WALNUT PVC"
                Dim VMATERIAL As String = "Maj.W.Mel"
                Dim VSMATERIAL As String = "Maj.W.PVC"
                Dim VEDGECODE As String = "5476"
                Dim PEDGECODE As String = "5476"
                'TALL UTILITY 2 UNITS UPPER
                TU2UUMIBackM = VMATERIAL 'SETS TALL UTILITY UPPER BACK MATERIAL TO WHITE MELAMINE
                TU2UUMIGableM = VMATERIAL 'SETS TALL UTILITY UPPER GABLE MATERIAL TO WHITE MELAMINE
                TU2UUMITopBtmM = VMATERIAL 'SETS TALL UTILITY UPPER TOP + BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMITopM = VMATERIAL 'SETS TALL UTILITY UPPER TOP MATERIAL TO WHITE MELAMINE
                TU2UUMIBotM = VMATERIAL 'SETS TALL UTILITY UPPER BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMIAdjShelfM = VMATERIAL 'SETS TALL UTILITY UPPER ADJUSTABLE SHELF BACK MATERIAL TO WHITE MELAMINE
                '# UPPER MATCHING INTERIOR #
                UMIBackM = VMATERIAL
                UMIGableM = VMATERIAL
                UMITopLightM = VMATERIAL
                UMIBtmM = VMATERIAL
                UMITopBtmM = VMATERIAL
                UMIAdjShelfM = VMATERIAL
                '# UPPER MICROWAVE MATCHING INTERIOR #
                UMMIBackM = VMATERIAL
                UMMIGableM = VMATERIAL
                UMMITopBtmM = VMATERIAL
                UMMIMicroSHM = VMATERIAL
                '# UPPER CORNER DIAGONAL MATCHING INTERIOR #
                UCDMILBackM = VMATERIAL
                UCDMISBackM = VMATERIAL
                UCDMIGableM = VMATERIAL
                UCDMITopBtmM = VMATERIAL
                UCDMIAdjShelfM = VMATERIAL
                '# UPPER WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                UWRBackM = VMATERIAL
                UWRGableM = VMATERIAL
                UWRTopBtmM = VMATERIAL
                UWRFFSHM = VMATERIAL
                UWRFShelfM = VMATERIAL
                UWRDivM = VMATERIAL
                '# UPPER END SHELF #
                UESLGableM = VMATERIAL
                UESSGableM = VMATERIAL
                UESTopM = VMATERIAL
                UESFixedShelfM = VMATERIAL
                '# BASE MATCHING INTERIOR #
                BMIBackM = VMATERIAL
                BMIGableM = VMATERIAL
                BMITopStrapM = VMATERIAL
                BMIBtmM = VMATERIAL
                BMIAdjShelfM = VMATERIAL
                BMITopM = VMATERIAL
                BMIStrapM = VMATERIAL
                BMIDividerM = VMATERIAL
                '# BASE PENINSULA OPEN #
                BPOPTopBtmFSHM = VMATERIAL
                BPOPSGableM = VMATERIAL
                BPOPLGableM = VMATERIAL
                '# VANITY MATCHING INTERIOR #
                VMIBackM = VMATERIAL
                VMIGableM = VMATERIAL
                VMITopStrapM = VMATERIAL
                VMITopM = VMATERIAL
                VMIBtmM = VMATERIAL
                VMIFFShelfM = VMATERIAL
                VMIStrapM = VMATERIAL
                VMIDividerM = VMATERIAL
                '# VANITY ELEVATED MATCHING INTERIOR #
                VEMIBackM = VMATERIAL
                VEMIGableM = VMATERIAL
                VEMITopStrapM = VMATERIAL
                VEMITopM = VMATERIAL
                VEMIBtmM = VMATERIAL
                VEMIFFShelfM = VMATERIAL
                VEMIStrapM = VMATERIAL
                VEMIDividerM = VMATERIAL
                '# BASE MICROWAVE OPEN SHELF 1 DRAWER #
                BMO1DUBackM = VMATERIAL
                BMO1DUGableM = VMATERIAL
                BMO1DUTopM = VMATERIAL
                BMO1DUBtmM = VMATERIAL
                '# BASE WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                BWRBackM = VMATERIAL
                BWRGableM = VMATERIAL
                BWRTopBtmM = VMATERIAL
                BWRFFSHM = VMATERIAL
                BWRValanceM = VSMATERIAL
                BWRVN = "Ordered Joe G."
                '# HUTCH MATCHING INTERIOR #
                HMIBackM = VMATERIAL
                HMIGableM = VMATERIAL
                HMITopFFShelfM = VMATERIAL
                HMIBtmFFShelfM = VMATERIAL
                HMITopM = VMATERIAL
                HMIBtmM = VMATERIAL
                HMIStrapM = VMATERIAL
                HMIDividerM = VMATERIAL
                HMIAdjShelfM = VMATERIAL
                '# CANOPY #
                RHFanSHM = "Ply"
                RHFrontM = VMATERIAL
                RHSFrontM = VMATERIAL
                RHSTopM = VMATERIAL
                RHIPanelM = VMATERIAL
                '# OVEN PANEL MATERIAL #
                OV1M = VMATERIAL
                '# FLOATING SHELF MATERIAL #
                FShelfM = VMATERIAL
                '# WINDOW PANEL MATERIAL #
                WPanelM = VMATERIAL
                '# WINDOW BOX MATERIAL #
                WTPanelM = VMATERIAL
                WFrontM = VMATERIAL
                '# DESK SUPPORT GABLE MATERIAL #
                DSGableM = VMATERIAL
                '# FANCY BRACKET MATERIAL #
                FBracketM = VMATERIAL
                '# FANCY VALANCE MATERIAL #
                FValanceM = VMATERIAL
                '# PLUMBING COVER PANEL MATERIAL #
                PCPanelM = VMATERIAL
                '# EDGE CODES VENEER AND PVC #
                VeneerEdgeCode = VEDGECODE
                PVCEdgeCode = PEDGECODE

            Case "MYSTIC PVC"
                Dim VMATERIAL As String = "Mys.Mel"
                Dim VSMATERIAL As String = "Mys.PVC"
                Dim VEDGECODE As String = "6464"
                Dim PEDGECODE As String = "6464"
                'TALL UTILITY 2 UNITS UPPER
                TU2UUMIBackM = VMATERIAL 'SETS TALL UTILITY UPPER BACK MATERIAL TO WHITE MELAMINE
                TU2UUMIGableM = VMATERIAL 'SETS TALL UTILITY UPPER GABLE MATERIAL TO WHITE MELAMINE
                TU2UUMITopBtmM = VMATERIAL 'SETS TALL UTILITY UPPER TOP + BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMITopM = VMATERIAL 'SETS TALL UTILITY UPPER TOP MATERIAL TO WHITE MELAMINE
                TU2UUMIBotM = VMATERIAL 'SETS TALL UTILITY UPPER BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMIAdjShelfM = VMATERIAL 'SETS TALL UTILITY UPPER ADJUSTABLE SHELF BACK MATERIAL TO WHITE MELAMINE
                '# UPPER MATCHING INTERIOR #
                UMIBackM = VMATERIAL
                UMIGableM = VMATERIAL
                UMITopLightM = VMATERIAL
                UMIBtmM = VMATERIAL
                UMITopBtmM = VMATERIAL
                UMIAdjShelfM = VMATERIAL
                '# UPPER MICROWAVE MATCHING INTERIOR #
                UMMIBackM = VMATERIAL
                UMMIGableM = VMATERIAL
                UMMITopBtmM = VMATERIAL
                UMMIMicroSHM = VMATERIAL
                '# UPPER CORNER DIAGONAL MATCHING INTERIOR #
                UCDMILBackM = VMATERIAL
                UCDMISBackM = VMATERIAL
                UCDMIGableM = VMATERIAL
                UCDMITopBtmM = VMATERIAL
                UCDMIAdjShelfM = VMATERIAL
                '# UPPER WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                UWRBackM = VMATERIAL
                UWRGableM = VMATERIAL
                UWRTopBtmM = VMATERIAL
                UWRFFSHM = VMATERIAL
                UWRFShelfM = VMATERIAL
                UWRDivM = VMATERIAL
                '# UPPER END SHELF #
                UESLGableM = VMATERIAL
                UESSGableM = VMATERIAL
                UESTopM = VMATERIAL
                UESFixedShelfM = VMATERIAL
                '# BASE MATCHING INTERIOR #
                BMIBackM = VMATERIAL
                BMIGableM = VMATERIAL
                BMITopStrapM = VMATERIAL
                BMIBtmM = VMATERIAL
                BMIAdjShelfM = VMATERIAL
                BMITopM = VMATERIAL
                BMIStrapM = VMATERIAL
                BMIDividerM = VMATERIAL
                '# BASE PENINSULA OPEN #
                BPOPTopBtmFSHM = VMATERIAL
                BPOPSGableM = VMATERIAL
                BPOPLGableM = VMATERIAL
                '# VANITY MATCHING INTERIOR #
                VMIBackM = VMATERIAL
                VMIGableM = VMATERIAL
                VMITopStrapM = VMATERIAL
                VMITopM = VMATERIAL
                VMIBtmM = VMATERIAL
                VMIFFShelfM = VMATERIAL
                VMIStrapM = VMATERIAL
                VMIDividerM = VMATERIAL
                '# VANITY ELEVATED MATCHING INTERIOR #
                VEMIBackM = VMATERIAL
                VEMIGableM = VMATERIAL
                VEMITopStrapM = VMATERIAL
                VEMITopM = VMATERIAL
                VEMIBtmM = VMATERIAL
                VEMIFFShelfM = VMATERIAL
                VEMIStrapM = VMATERIAL
                VEMIDividerM = VMATERIAL
                '# BASE MICROWAVE OPEN SHELF 1 DRAWER #
                BMO1DUBackM = VMATERIAL
                BMO1DUGableM = VMATERIAL
                BMO1DUTopM = VMATERIAL
                BMO1DUBtmM = VMATERIAL
                '# BASE WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                BWRBackM = VMATERIAL
                BWRGableM = VMATERIAL
                BWRTopBtmM = VMATERIAL
                BWRFFSHM = VMATERIAL
                BWRValanceM = VSMATERIAL
                BWRVN = "Ordered Joe G."
                '# HUTCH MATCHING INTERIOR #
                HMIBackM = VMATERIAL
                HMIGableM = VMATERIAL
                HMITopFFShelfM = VMATERIAL
                HMIBtmFFShelfM = VMATERIAL
                HMITopM = VMATERIAL
                HMIBtmM = VMATERIAL
                HMIStrapM = VMATERIAL
                HMIDividerM = VMATERIAL
                HMIAdjShelfM = VMATERIAL
                '# CANOPY #
                RHFanSHM = "Ply"
                RHFrontM = VMATERIAL
                RHSFrontM = VMATERIAL
                RHSTopM = VMATERIAL
                RHIPanelM = VMATERIAL
                '# OVEN PANEL MATERIAL #
                OV1M = VMATERIAL
                '# FLOATING SHELF MATERIAL #
                FShelfM = VMATERIAL
                '# WINDOW PANEL MATERIAL #
                WPanelM = VMATERIAL
                '# WINDOW BOX MATERIAL #
                WTPanelM = VMATERIAL
                WFrontM = VMATERIAL
                '# DESK SUPPORT GABLE MATERIAL #
                DSGableM = VMATERIAL
                '# FANCY BRACKET MATERIAL #
                FBracketM = VMATERIAL
                '# FANCY VALANCE MATERIAL #
                FValanceM = VMATERIAL
                '# PLUMBING COVER PANEL MATERIAL #
                PCPanelM = VMATERIAL
                '# EDGE CODES VENEER AND PVC #
                VeneerEdgeCode = VEDGECODE
                PVCEdgeCode = PEDGECODE

            Case "NATURAL MAPLE PVC"
                Dim VMATERIAL As String = "Nat.M.Mel"
                Dim VSMATERIAL As String = "Nat.M.PVC"
                Dim VEDGECODE As String = "6116"
                Dim PEDGECODE As String = "6116"
                'TALL UTILITY 2 UNITS UPPER
                TU2UUMIBackM = VMATERIAL 'SETS TALL UTILITY UPPER BACK MATERIAL TO WHITE MELAMINE
                TU2UUMIGableM = VMATERIAL 'SETS TALL UTILITY UPPER GABLE MATERIAL TO WHITE MELAMINE
                TU2UUMITopBtmM = VMATERIAL 'SETS TALL UTILITY UPPER TOP + BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMITopM = VMATERIAL 'SETS TALL UTILITY UPPER TOP MATERIAL TO WHITE MELAMINE
                TU2UUMIBotM = VMATERIAL 'SETS TALL UTILITY UPPER BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMIAdjShelfM = VMATERIAL 'SETS TALL UTILITY UPPER ADJUSTABLE SHELF BACK MATERIAL TO WHITE MELAMINE
                '# UPPER MATCHING INTERIOR #
                UMIBackM = VMATERIAL
                UMIGableM = VMATERIAL
                UMITopLightM = VMATERIAL
                UMIBtmM = VMATERIAL
                UMITopBtmM = VMATERIAL
                UMIAdjShelfM = VMATERIAL
                '# UPPER MICROWAVE MATCHING INTERIOR #
                UMMIBackM = VMATERIAL
                UMMIGableM = VMATERIAL
                UMMITopBtmM = VMATERIAL
                UMMIMicroSHM = VMATERIAL
                '# UPPER CORNER DIAGONAL MATCHING INTERIOR #
                UCDMILBackM = VMATERIAL
                UCDMISBackM = VMATERIAL
                UCDMIGableM = VMATERIAL
                UCDMITopBtmM = VMATERIAL
                UCDMIAdjShelfM = VMATERIAL
                '# UPPER WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                UWRBackM = VMATERIAL
                UWRGableM = VMATERIAL
                UWRTopBtmM = VMATERIAL
                UWRFFSHM = VMATERIAL
                UWRFShelfM = VMATERIAL
                UWRDivM = VMATERIAL
                '# UPPER END SHELF #
                UESLGableM = VMATERIAL
                UESSGableM = VMATERIAL
                UESTopM = VMATERIAL
                UESFixedShelfM = VMATERIAL
                '# BASE MATCHING INTERIOR #
                BMIBackM = VMATERIAL
                BMIGableM = VMATERIAL
                BMITopStrapM = VMATERIAL
                BMIBtmM = VMATERIAL
                BMIAdjShelfM = VMATERIAL
                BMITopM = VMATERIAL
                BMIStrapM = VMATERIAL
                BMIDividerM = VMATERIAL
                '# BASE PENINSULA OPEN #
                BPOPTopBtmFSHM = VMATERIAL
                BPOPSGableM = VMATERIAL
                BPOPLGableM = VMATERIAL
                '# VANITY MATCHING INTERIOR #
                VMIBackM = VMATERIAL
                VMIGableM = VMATERIAL
                VMITopStrapM = VMATERIAL
                VMITopM = VMATERIAL
                VMIBtmM = VMATERIAL
                VMIFFShelfM = VMATERIAL
                VMIStrapM = VMATERIAL
                VMIDividerM = VMATERIAL
                '# VANITY ELEVATED MATCHING INTERIOR #
                VEMIBackM = VMATERIAL
                VEMIGableM = VMATERIAL
                VEMITopStrapM = VMATERIAL
                VEMITopM = VMATERIAL
                VEMIBtmM = VMATERIAL
                VEMIFFShelfM = VMATERIAL
                VEMIStrapM = VMATERIAL
                VEMIDividerM = VMATERIAL
                '# BASE MICROWAVE OPEN SHELF 1 DRAWER #
                BMO1DUBackM = VMATERIAL
                BMO1DUGableM = VMATERIAL
                BMO1DUTopM = VMATERIAL
                BMO1DUBtmM = VMATERIAL
                '# BASE WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                BWRBackM = VMATERIAL
                BWRGableM = VMATERIAL
                BWRTopBtmM = VMATERIAL
                BWRFFSHM = VMATERIAL
                BWRValanceM = VSMATERIAL
                BWRVN = "Ordered Joe G."
                '# HUTCH MATCHING INTERIOR #
                HMIBackM = VMATERIAL
                HMIGableM = VMATERIAL
                HMITopFFShelfM = VMATERIAL
                HMIBtmFFShelfM = VMATERIAL
                HMITopM = VMATERIAL
                HMIBtmM = VMATERIAL
                HMIStrapM = VMATERIAL
                HMIDividerM = VMATERIAL
                HMIAdjShelfM = VMATERIAL
                '# CANOPY #
                RHFanSHM = "Ply"
                RHFrontM = VMATERIAL
                RHSFrontM = VMATERIAL
                RHSTopM = VMATERIAL
                RHIPanelM = VMATERIAL
                '# OVEN PANEL MATERIAL #
                OV1M = VMATERIAL
                '# FLOATING SHELF MATERIAL #
                FShelfM = VMATERIAL
                '# WINDOW PANEL MATERIAL #
                WPanelM = VMATERIAL
                '# WINDOW BOX MATERIAL #
                WTPanelM = VMATERIAL
                WFrontM = VMATERIAL
                '# DESK SUPPORT GABLE MATERIAL #
                DSGableM = VMATERIAL
                '# FANCY BRACKET MATERIAL #
                FBracketM = VMATERIAL
                '# FANCY VALANCE MATERIAL #
                FValanceM = VMATERIAL
                '# PLUMBING COVER PANEL MATERIAL #
                PCPanelM = VMATERIAL
                '# EDGE CODES VENEER AND PVC #
                VeneerEdgeCode = VEDGECODE
                PVCEdgeCode = PEDGECODE

            Case "PINK MAPLE PVC"
                Dim VMATERIAL As String = "Pink.M.Mel"
                Dim VSMATERIAL As String = "Pink.M.PVC"
                Dim VEDGECODE As String = "8567"
                Dim PEDGECODE As String = "8567"
                'TALL UTILITY 2 UNITS UPPER
                TU2UUMIBackM = VMATERIAL 'SETS TALL UTILITY UPPER BACK MATERIAL TO WHITE MELAMINE
                TU2UUMIGableM = VMATERIAL 'SETS TALL UTILITY UPPER GABLE MATERIAL TO WHITE MELAMINE
                TU2UUMITopBtmM = VMATERIAL 'SETS TALL UTILITY UPPER TOP + BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMITopM = VMATERIAL 'SETS TALL UTILITY UPPER TOP MATERIAL TO WHITE MELAMINE
                TU2UUMIBotM = VMATERIAL 'SETS TALL UTILITY UPPER BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMIAdjShelfM = VMATERIAL 'SETS TALL UTILITY UPPER ADJUSTABLE SHELF BACK MATERIAL TO WHITE MELAMINE
                '# UPPER MATCHING INTERIOR #
                UMIBackM = VMATERIAL
                UMIGableM = VMATERIAL
                UMITopLightM = VMATERIAL
                UMIBtmM = VMATERIAL
                UMITopBtmM = VMATERIAL
                UMIAdjShelfM = VMATERIAL
                '# UPPER MICROWAVE MATCHING INTERIOR #
                UMMIBackM = VMATERIAL
                UMMIGableM = VMATERIAL
                UMMITopBtmM = VMATERIAL
                UMMIMicroSHM = VMATERIAL
                '# UPPER CORNER DIAGONAL MATCHING INTERIOR #
                UCDMILBackM = VMATERIAL
                UCDMISBackM = VMATERIAL
                UCDMIGableM = VMATERIAL
                UCDMITopBtmM = VMATERIAL
                UCDMIAdjShelfM = VMATERIAL
                '# UPPER WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                UWRBackM = VMATERIAL
                UWRGableM = VMATERIAL
                UWRTopBtmM = VMATERIAL
                UWRFFSHM = VMATERIAL
                UWRFShelfM = VMATERIAL
                UWRDivM = VMATERIAL
                '# UPPER END SHELF #
                UESLGableM = VMATERIAL
                UESSGableM = VMATERIAL
                UESTopM = VMATERIAL
                UESFixedShelfM = VMATERIAL
                '# BASE MATCHING INTERIOR #
                BMIBackM = VMATERIAL
                BMIGableM = VMATERIAL
                BMITopStrapM = VMATERIAL
                BMIBtmM = VMATERIAL
                BMIAdjShelfM = VMATERIAL
                BMITopM = VMATERIAL
                BMIStrapM = VMATERIAL
                BMIDividerM = VMATERIAL
                '# BASE PENINSULA OPEN #
                BPOPTopBtmFSHM = VMATERIAL
                BPOPSGableM = VMATERIAL
                BPOPLGableM = VMATERIAL
                '# VANITY MATCHING INTERIOR #
                VMIBackM = VMATERIAL
                VMIGableM = VMATERIAL
                VMITopStrapM = VMATERIAL
                VMITopM = VMATERIAL
                VMIBtmM = VMATERIAL
                VMIFFShelfM = VMATERIAL
                VMIStrapM = VMATERIAL
                VMIDividerM = VMATERIAL
                '# VANITY ELEVATED MATCHING INTERIOR #
                VEMIBackM = VMATERIAL
                VEMIGableM = VMATERIAL
                VEMITopStrapM = VMATERIAL
                VEMITopM = VMATERIAL
                VEMIBtmM = VMATERIAL
                VEMIFFShelfM = VMATERIAL
                VEMIStrapM = VMATERIAL
                VEMIDividerM = VMATERIAL
                '# BASE MICROWAVE OPEN SHELF 1 DRAWER #
                BMO1DUBackM = VMATERIAL
                BMO1DUGableM = VMATERIAL
                BMO1DUTopM = VMATERIAL
                BMO1DUBtmM = VMATERIAL
                '# BASE WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                BWRBackM = VMATERIAL
                BWRGableM = VMATERIAL
                BWRTopBtmM = VMATERIAL
                BWRFFSHM = VMATERIAL
                BWRValanceM = VSMATERIAL
                BWRVN = "Ordered Joe G."
                '# HUTCH MATCHING INTERIOR #
                HMIBackM = VMATERIAL
                HMIGableM = VMATERIAL
                HMITopFFShelfM = VMATERIAL
                HMIBtmFFShelfM = VMATERIAL
                HMITopM = VMATERIAL
                HMIBtmM = VMATERIAL
                HMIStrapM = VMATERIAL
                HMIDividerM = VMATERIAL
                HMIAdjShelfM = VMATERIAL
                '# CANOPY #
                RHFanSHM = "Ply"
                RHFrontM = VMATERIAL
                RHSFrontM = VMATERIAL
                RHSTopM = VMATERIAL
                RHIPanelM = VMATERIAL
                '# OVEN PANEL MATERIAL #
                OV1M = VMATERIAL
                '# FLOATING SHELF MATERIAL #
                FShelfM = VMATERIAL
                '# WINDOW PANEL MATERIAL #
                WPanelM = VMATERIAL
                '# WINDOW BOX MATERIAL #
                WTPanelM = VMATERIAL
                WFrontM = VMATERIAL
                '# DESK SUPPORT GABLE MATERIAL #
                DSGableM = VMATERIAL
                '# FANCY BRACKET MATERIAL #
                FBracketM = VMATERIAL
                '# FANCY VALANCE MATERIAL #
                FValanceM = VMATERIAL
                '# PLUMBING COVER PANEL MATERIAL #
                PCPanelM = VMATERIAL
                '# EDGE CODES VENEER AND PVC #
                VeneerEdgeCode = VEDGECODE
                PVCEdgeCode = PEDGECODE

            Case "PORTLAND CHERRY PVC"
                Dim VMATERIAL As String = "Por.Ch.Mel"
                Dim VSMATERIAL As String = "Por.Ch.PVC"
                Dim VEDGECODE As String = "9774"
                Dim PEDGECODE As String = "9774"
                'TALL UTILITY 2 UNITS UPPER
                TU2UUMIBackM = VMATERIAL 'SETS TALL UTILITY UPPER BACK MATERIAL TO WHITE MELAMINE
                TU2UUMIGableM = VMATERIAL 'SETS TALL UTILITY UPPER GABLE MATERIAL TO WHITE MELAMINE
                TU2UUMITopBtmM = VMATERIAL 'SETS TALL UTILITY UPPER TOP + BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMITopM = VMATERIAL 'SETS TALL UTILITY UPPER TOP MATERIAL TO WHITE MELAMINE
                TU2UUMIBotM = VMATERIAL 'SETS TALL UTILITY UPPER BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMIAdjShelfM = VMATERIAL 'SETS TALL UTILITY UPPER ADJUSTABLE SHELF BACK MATERIAL TO WHITE MELAMINE
                '# UPPER MATCHING INTERIOR #
                UMIBackM = VMATERIAL
                UMIGableM = VMATERIAL
                UMITopLightM = VMATERIAL
                UMIBtmM = VMATERIAL
                UMITopBtmM = VMATERIAL
                UMIAdjShelfM = VMATERIAL
                '# UPPER MICROWAVE MATCHING INTERIOR #
                UMMIBackM = VMATERIAL
                UMMIGableM = VMATERIAL
                UMMITopBtmM = VMATERIAL
                UMMIMicroSHM = VMATERIAL
                '# UPPER CORNER DIAGONAL MATCHING INTERIOR #
                UCDMILBackM = VMATERIAL
                UCDMISBackM = VMATERIAL
                UCDMIGableM = VMATERIAL
                UCDMITopBtmM = VMATERIAL
                UCDMIAdjShelfM = VMATERIAL
                '# UPPER WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                UWRBackM = VMATERIAL
                UWRGableM = VMATERIAL
                UWRTopBtmM = VMATERIAL
                UWRFFSHM = VMATERIAL
                UWRFShelfM = VMATERIAL
                UWRDivM = VMATERIAL
                '# UPPER END SHELF #
                UESLGableM = VMATERIAL
                UESSGableM = VMATERIAL
                UESTopM = VMATERIAL
                UESFixedShelfM = VMATERIAL
                '# BASE MATCHING INTERIOR #
                BMIBackM = VMATERIAL
                BMIGableM = VMATERIAL
                BMITopStrapM = VMATERIAL
                BMIBtmM = VMATERIAL
                BMIAdjShelfM = VMATERIAL
                BMITopM = VMATERIAL
                BMIStrapM = VMATERIAL
                BMIDividerM = VMATERIAL
                '# BASE PENINSULA OPEN #
                BPOPTopBtmFSHM = VMATERIAL
                BPOPSGableM = VMATERIAL
                BPOPLGableM = VMATERIAL
                '# VANITY MATCHING INTERIOR #
                VMIBackM = VMATERIAL
                VMIGableM = VMATERIAL
                VMITopStrapM = VMATERIAL
                VMITopM = VMATERIAL
                VMIBtmM = VMATERIAL
                VMIFFShelfM = VMATERIAL
                VMIStrapM = VMATERIAL
                VMIDividerM = VMATERIAL
                '# VANITY ELEVATED MATCHING INTERIOR #
                VEMIBackM = VMATERIAL
                VEMIGableM = VMATERIAL
                VEMITopStrapM = VMATERIAL
                VEMITopM = VMATERIAL
                VEMIBtmM = VMATERIAL
                VEMIFFShelfM = VMATERIAL
                VEMIStrapM = VMATERIAL
                VEMIDividerM = VMATERIAL
                '# BASE MICROWAVE OPEN SHELF 1 DRAWER #
                BMO1DUBackM = VMATERIAL
                BMO1DUGableM = VMATERIAL
                BMO1DUTopM = VMATERIAL
                BMO1DUBtmM = VMATERIAL
                '# BASE WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                BWRBackM = VMATERIAL
                BWRGableM = VMATERIAL
                BWRTopBtmM = VMATERIAL
                BWRFFSHM = VMATERIAL
                BWRValanceM = VSMATERIAL
                BWRVN = "Ordered Joe G."
                '# HUTCH MATCHING INTERIOR #
                HMIBackM = VMATERIAL
                HMIGableM = VMATERIAL
                HMITopFFShelfM = VMATERIAL
                HMIBtmFFShelfM = VMATERIAL
                HMITopM = VMATERIAL
                HMIBtmM = VMATERIAL
                HMIStrapM = VMATERIAL
                HMIDividerM = VMATERIAL
                HMIAdjShelfM = VMATERIAL
                '# CANOPY #
                RHFanSHM = "Ply"
                RHFrontM = VMATERIAL
                RHSFrontM = VMATERIAL
                RHSTopM = VMATERIAL
                RHIPanelM = VMATERIAL
                '# OVEN PANEL MATERIAL #
                OV1M = VMATERIAL
                '# FLOATING SHELF MATERIAL #
                FShelfM = VMATERIAL
                '# WINDOW PANEL MATERIAL #
                WPanelM = VMATERIAL
                '# WINDOW BOX MATERIAL #
                WTPanelM = VMATERIAL
                WFrontM = VMATERIAL
                '# DESK SUPPORT GABLE MATERIAL #
                DSGableM = VMATERIAL
                '# FANCY BRACKET MATERIAL #
                FBracketM = VMATERIAL
                '# FANCY VALANCE MATERIAL #
                FValanceM = VMATERIAL
                '# PLUMBING COVER PANEL MATERIAL #
                PCPanelM = VMATERIAL
                '# EDGE CODES VENEER AND PVC #
                VeneerEdgeCode = VEDGECODE
                PVCEdgeCode = PEDGECODE

            Case "RED APPLE PVC"
                Dim VMATERIAL As String = "Red.A.Mel"
                Dim VSMATERIAL As String = "Red.A.PVC"
                Dim VEDGECODE As String = "2668"
                Dim PEDGECODE As String = "2668"
                'TALL UTILITY 2 UNITS UPPER
                TU2UUMIBackM = VMATERIAL 'SETS TALL UTILITY UPPER BACK MATERIAL TO WHITE MELAMINE
                TU2UUMIGableM = VMATERIAL 'SETS TALL UTILITY UPPER GABLE MATERIAL TO WHITE MELAMINE
                TU2UUMITopBtmM = VMATERIAL 'SETS TALL UTILITY UPPER TOP + BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMITopM = VMATERIAL 'SETS TALL UTILITY UPPER TOP MATERIAL TO WHITE MELAMINE
                TU2UUMIBotM = VMATERIAL 'SETS TALL UTILITY UPPER BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMIAdjShelfM = VMATERIAL 'SETS TALL UTILITY UPPER ADJUSTABLE SHELF BACK MATERIAL TO WHITE MELAMINE
                '# UPPER MATCHING INTERIOR #
                UMIBackM = VMATERIAL
                UMIGableM = VMATERIAL
                UMITopLightM = VMATERIAL
                UMIBtmM = VMATERIAL
                UMITopBtmM = VMATERIAL
                UMIAdjShelfM = VMATERIAL
                '# UPPER MICROWAVE MATCHING INTERIOR #
                UMMIBackM = VMATERIAL
                UMMIGableM = VMATERIAL
                UMMITopBtmM = VMATERIAL
                UMMIMicroSHM = VMATERIAL
                '# UPPER CORNER DIAGONAL MATCHING INTERIOR #
                UCDMILBackM = VMATERIAL
                UCDMISBackM = VMATERIAL
                UCDMIGableM = VMATERIAL
                UCDMITopBtmM = VMATERIAL
                UCDMIAdjShelfM = VMATERIAL
                '# UPPER WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                UWRBackM = VMATERIAL
                UWRGableM = VMATERIAL
                UWRTopBtmM = VMATERIAL
                UWRFFSHM = VMATERIAL
                UWRFShelfM = VMATERIAL
                UWRDivM = VMATERIAL
                '# UPPER END SHELF #
                UESLGableM = VMATERIAL
                UESSGableM = VMATERIAL
                UESTopM = VMATERIAL
                UESFixedShelfM = VMATERIAL
                '# BASE MATCHING INTERIOR #
                BMIBackM = VMATERIAL
                BMIGableM = VMATERIAL
                BMITopStrapM = VMATERIAL
                BMIBtmM = VMATERIAL
                BMIAdjShelfM = VMATERIAL
                BMITopM = VMATERIAL
                BMIStrapM = VMATERIAL
                BMIDividerM = VMATERIAL
                '# BASE PENINSULA OPEN #
                BPOPTopBtmFSHM = VMATERIAL
                BPOPSGableM = VMATERIAL
                BPOPLGableM = VMATERIAL
                '# VANITY MATCHING INTERIOR #
                VMIBackM = VMATERIAL
                VMIGableM = VMATERIAL
                VMITopStrapM = VMATERIAL
                VMITopM = VMATERIAL
                VMIBtmM = VMATERIAL
                VMIFFShelfM = VMATERIAL
                VMIStrapM = VMATERIAL
                VMIDividerM = VMATERIAL
                '# VANITY ELEVATED MATCHING INTERIOR #
                VEMIBackM = VMATERIAL
                VEMIGableM = VMATERIAL
                VEMITopStrapM = VMATERIAL
                VEMITopM = VMATERIAL
                VEMIBtmM = VMATERIAL
                VEMIFFShelfM = VMATERIAL
                VEMIStrapM = VMATERIAL
                VEMIDividerM = VMATERIAL
                '# BASE MICROWAVE OPEN SHELF 1 DRAWER #
                BMO1DUBackM = VMATERIAL
                BMO1DUGableM = VMATERIAL
                BMO1DUTopM = VMATERIAL
                BMO1DUBtmM = VMATERIAL
                '# BASE WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                BWRBackM = VMATERIAL
                BWRGableM = VMATERIAL
                BWRTopBtmM = VMATERIAL
                BWRFFSHM = VMATERIAL
                BWRValanceM = VSMATERIAL
                BWRVN = "Ordered Joe G."
                '# HUTCH MATCHING INTERIOR #
                HMIBackM = VMATERIAL
                HMIGableM = VMATERIAL
                HMITopFFShelfM = VMATERIAL
                HMIBtmFFShelfM = VMATERIAL
                HMITopM = VMATERIAL
                HMIBtmM = VMATERIAL
                HMIStrapM = VMATERIAL
                HMIDividerM = VMATERIAL
                HMIAdjShelfM = VMATERIAL
                '# CANOPY #
                RHFanSHM = "Ply"
                RHFrontM = VMATERIAL
                RHSFrontM = VMATERIAL
                RHSTopM = VMATERIAL
                RHIPanelM = VMATERIAL
                '# OVEN PANEL MATERIAL #
                OV1M = VMATERIAL
                '# FLOATING SHELF MATERIAL #
                FShelfM = VMATERIAL
                '# WINDOW PANEL MATERIAL #
                WPanelM = VMATERIAL
                '# WINDOW BOX MATERIAL #
                WTPanelM = VMATERIAL
                WFrontM = VMATERIAL
                '# DESK SUPPORT GABLE MATERIAL #
                DSGableM = VMATERIAL
                '# FANCY BRACKET MATERIAL #
                FBracketM = VMATERIAL
                '# FANCY VALANCE MATERIAL #
                FValanceM = VMATERIAL
                '# PLUMBING COVER PANEL MATERIAL #
                PCPanelM = VMATERIAL
                '# EDGE CODES VENEER AND PVC #
                VeneerEdgeCode = VEDGECODE
                PVCEdgeCode = PEDGECODE

            Case "SILKEN MAPLE PVC"
                Dim VMATERIAL As String = "Sil.M.Mel"
                Dim VSMATERIAL As String = "Sil.M.PVC"
                Dim VEDGECODE As String = "5557"
                Dim PEDGECODE As String = "5557"
                'TALL UTILITY 2 UNITS UPPER
                TU2UUMIBackM = VMATERIAL 'SETS TALL UTILITY UPPER BACK MATERIAL TO WHITE MELAMINE
                TU2UUMIGableM = VMATERIAL 'SETS TALL UTILITY UPPER GABLE MATERIAL TO WHITE MELAMINE
                TU2UUMITopBtmM = VMATERIAL 'SETS TALL UTILITY UPPER TOP + BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMITopM = VMATERIAL 'SETS TALL UTILITY UPPER TOP MATERIAL TO WHITE MELAMINE
                TU2UUMIBotM = VMATERIAL 'SETS TALL UTILITY UPPER BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMIAdjShelfM = VMATERIAL 'SETS TALL UTILITY UPPER ADJUSTABLE SHELF BACK MATERIAL TO WHITE MELAMINE
                '# UPPER MATCHING INTERIOR #
                UMIBackM = VMATERIAL
                UMIGableM = VMATERIAL
                UMITopLightM = VMATERIAL
                UMIBtmM = VMATERIAL
                UMITopBtmM = VMATERIAL
                UMIAdjShelfM = VMATERIAL
                '# UPPER MICROWAVE MATCHING INTERIOR #
                UMMIBackM = VMATERIAL
                UMMIGableM = VMATERIAL
                UMMITopBtmM = VMATERIAL
                UMMIMicroSHM = VMATERIAL
                '# UPPER CORNER DIAGONAL MATCHING INTERIOR #
                UCDMILBackM = VMATERIAL
                UCDMISBackM = VMATERIAL
                UCDMIGableM = VMATERIAL
                UCDMITopBtmM = VMATERIAL
                UCDMIAdjShelfM = VMATERIAL
                '# UPPER WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                UWRBackM = VMATERIAL
                UWRGableM = VMATERIAL
                UWRTopBtmM = VMATERIAL
                UWRFFSHM = VMATERIAL
                UWRFShelfM = VMATERIAL
                UWRDivM = VMATERIAL
                '# UPPER END SHELF #
                UESLGableM = VMATERIAL
                UESSGableM = VMATERIAL
                UESTopM = VMATERIAL
                UESFixedShelfM = VMATERIAL
                '# BASE MATCHING INTERIOR #
                BMIBackM = VMATERIAL
                BMIGableM = VMATERIAL
                BMITopStrapM = VMATERIAL
                BMIBtmM = VMATERIAL
                BMIAdjShelfM = VMATERIAL
                BMITopM = VMATERIAL
                BMIStrapM = VMATERIAL
                BMIDividerM = VMATERIAL
                '# BASE PENINSULA OPEN #
                BPOPTopBtmFSHM = VMATERIAL
                BPOPSGableM = VMATERIAL
                BPOPLGableM = VMATERIAL
                '# VANITY MATCHING INTERIOR #
                VMIBackM = VMATERIAL
                VMIGableM = VMATERIAL
                VMITopStrapM = VMATERIAL
                VMITopM = VMATERIAL
                VMIBtmM = VMATERIAL
                VMIFFShelfM = VMATERIAL
                VMIStrapM = VMATERIAL
                VMIDividerM = VMATERIAL
                '# VANITY ELEVATED MATCHING INTERIOR #
                VEMIBackM = VMATERIAL
                VEMIGableM = VMATERIAL
                VEMITopStrapM = VMATERIAL
                VEMITopM = VMATERIAL
                VEMIBtmM = VMATERIAL
                VEMIFFShelfM = VMATERIAL
                VEMIStrapM = VMATERIAL
                VEMIDividerM = VMATERIAL
                '# BASE MICROWAVE OPEN SHELF 1 DRAWER #
                BMO1DUBackM = VMATERIAL
                BMO1DUGableM = VMATERIAL
                BMO1DUTopM = VMATERIAL
                BMO1DUBtmM = VMATERIAL
                '# BASE WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                BWRBackM = VMATERIAL
                BWRGableM = VMATERIAL
                BWRTopBtmM = VMATERIAL
                BWRFFSHM = VMATERIAL
                BWRValanceM = VSMATERIAL
                BWRVN = "Ordered Joe G."
                '# HUTCH MATCHING INTERIOR #
                HMIBackM = VMATERIAL
                HMIGableM = VMATERIAL
                HMITopFFShelfM = VMATERIAL
                HMIBtmFFShelfM = VMATERIAL
                HMITopM = VMATERIAL
                HMIBtmM = VMATERIAL
                HMIStrapM = VMATERIAL
                HMIDividerM = VMATERIAL
                HMIAdjShelfM = VMATERIAL
                '# CANOPY #
                RHFanSHM = "Ply"
                RHFrontM = VMATERIAL
                RHSFrontM = VMATERIAL
                RHSTopM = VMATERIAL
                RHIPanelM = VMATERIAL
                '# OVEN PANEL MATERIAL #
                OV1M = VMATERIAL
                '# FLOATING SHELF MATERIAL #
                FShelfM = VMATERIAL
                '# WINDOW PANEL MATERIAL #
                WPanelM = VMATERIAL
                '# WINDOW BOX MATERIAL #
                WTPanelM = VMATERIAL
                WFrontM = VMATERIAL
                '# DESK SUPPORT GABLE MATERIAL #
                DSGableM = VMATERIAL
                '# FANCY BRACKET MATERIAL #
                FBracketM = VMATERIAL
                '# FANCY VALANCE MATERIAL #
                FValanceM = VMATERIAL
                '# PLUMBING COVER PANEL MATERIAL #
                PCPanelM = VMATERIAL
                '# EDGE CODES VENEER AND PVC #
                VeneerEdgeCode = VEDGECODE
                PVCEdgeCode = PEDGECODE

            Case "SILVER PVC"
                Dim VMATERIAL As String = "Silv.Mel"
                Dim VSMATERIAL As String = "Silv.PVC"
                Dim VEDGECODE As String = "151"
                Dim PEDGECODE As String = "151"
                'TALL UTILITY 2 UNITS UPPER
                TU2UUMIBackM = VMATERIAL 'SETS TALL UTILITY UPPER BACK MATERIAL TO WHITE MELAMINE
                TU2UUMIGableM = VMATERIAL 'SETS TALL UTILITY UPPER GABLE MATERIAL TO WHITE MELAMINE
                TU2UUMITopBtmM = VMATERIAL 'SETS TALL UTILITY UPPER TOP + BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMITopM = VMATERIAL 'SETS TALL UTILITY UPPER TOP MATERIAL TO WHITE MELAMINE
                TU2UUMIBotM = VMATERIAL 'SETS TALL UTILITY UPPER BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMIAdjShelfM = VMATERIAL 'SETS TALL UTILITY UPPER ADJUSTABLE SHELF BACK MATERIAL TO WHITE MELAMINE
                '# UPPER MATCHING INTERIOR #
                UMIBackM = VMATERIAL
                UMIGableM = VMATERIAL
                UMITopLightM = VMATERIAL
                UMIBtmM = VMATERIAL
                UMITopBtmM = VMATERIAL
                UMIAdjShelfM = VMATERIAL
                '# UPPER MICROWAVE MATCHING INTERIOR #
                UMMIBackM = VMATERIAL
                UMMIGableM = VMATERIAL
                UMMITopBtmM = VMATERIAL
                UMMIMicroSHM = VMATERIAL
                '# UPPER CORNER DIAGONAL MATCHING INTERIOR #
                UCDMILBackM = VMATERIAL
                UCDMISBackM = VMATERIAL
                UCDMIGableM = VMATERIAL
                UCDMITopBtmM = VMATERIAL
                UCDMIAdjShelfM = VMATERIAL
                '# UPPER WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                UWRBackM = VMATERIAL
                UWRGableM = VMATERIAL
                UWRTopBtmM = VMATERIAL
                UWRFFSHM = VMATERIAL
                UWRFShelfM = VMATERIAL
                UWRDivM = VMATERIAL
                '# UPPER END SHELF #
                UESLGableM = VMATERIAL
                UESSGableM = VMATERIAL
                UESTopM = VMATERIAL
                UESFixedShelfM = VMATERIAL
                '# BASE MATCHING INTERIOR #
                BMIBackM = VMATERIAL
                BMIGableM = VMATERIAL
                BMITopStrapM = VMATERIAL
                BMIBtmM = VMATERIAL
                BMIAdjShelfM = VMATERIAL
                BMITopM = VMATERIAL
                BMIStrapM = VMATERIAL
                BMIDividerM = VMATERIAL
                '# BASE PENINSULA OPEN #
                BPOPTopBtmFSHM = VMATERIAL
                BPOPSGableM = VMATERIAL
                BPOPLGableM = VMATERIAL
                '# VANITY MATCHING INTERIOR #
                VMIBackM = VMATERIAL
                VMIGableM = VMATERIAL
                VMITopStrapM = VMATERIAL
                VMITopM = VMATERIAL
                VMIBtmM = VMATERIAL
                VMIFFShelfM = VMATERIAL
                VMIStrapM = VMATERIAL
                VMIDividerM = VMATERIAL
                '# VANITY ELEVATED MATCHING INTERIOR #
                VEMIBackM = VMATERIAL
                VEMIGableM = VMATERIAL
                VEMITopStrapM = VMATERIAL
                VEMITopM = VMATERIAL
                VEMIBtmM = VMATERIAL
                VEMIFFShelfM = VMATERIAL
                VEMIStrapM = VMATERIAL
                VEMIDividerM = VMATERIAL
                '# BASE MICROWAVE OPEN SHELF 1 DRAWER #
                BMO1DUBackM = VMATERIAL
                BMO1DUGableM = VMATERIAL
                BMO1DUTopM = VMATERIAL
                BMO1DUBtmM = VMATERIAL
                '# BASE WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                BWRBackM = VMATERIAL
                BWRGableM = VMATERIAL
                BWRTopBtmM = VMATERIAL
                BWRFFSHM = VMATERIAL
                BWRValanceM = VSMATERIAL
                BWRVN = "Ordered Joe G."
                '# HUTCH MATCHING INTERIOR #
                HMIBackM = VMATERIAL
                HMIGableM = VMATERIAL
                HMITopFFShelfM = VMATERIAL
                HMIBtmFFShelfM = VMATERIAL
                HMITopM = VMATERIAL
                HMIBtmM = VMATERIAL
                HMIStrapM = VMATERIAL
                HMIDividerM = VMATERIAL
                HMIAdjShelfM = VMATERIAL
                '# CANOPY #
                RHFanSHM = "Ply"
                RHFrontM = VMATERIAL
                RHSFrontM = VMATERIAL
                RHSTopM = VMATERIAL
                RHIPanelM = VMATERIAL
                '# OVEN PANEL MATERIAL #
                OV1M = VMATERIAL
                '# FLOATING SHELF MATERIAL #
                FShelfM = VMATERIAL
                '# WINDOW PANEL MATERIAL #
                WPanelM = VMATERIAL
                '# WINDOW BOX MATERIAL #
                WTPanelM = VMATERIAL
                WFrontM = VMATERIAL
                '# DESK SUPPORT GABLE MATERIAL #
                DSGableM = VMATERIAL
                '# FANCY BRACKET MATERIAL #
                FBracketM = VMATERIAL
                '# FANCY VALANCE MATERIAL #
                FValanceM = VMATERIAL
                '# PLUMBING COVER PANEL MATERIAL #
                PCPanelM = VMATERIAL
                '# EDGE CODES VENEER AND PVC #
                VeneerEdgeCode = VEDGECODE
                PVCEdgeCode = PEDGECODE

            Case "VANILLA STIX PVC"
                Dim VMATERIAL As String = "Van.St.Mel"
                Dim VSMATERIAL As String = "Van.St.PVC"
                Dim VEDGECODE As String = "5942"
                Dim PEDGECODE As String = "5942"
                'TALL UTILITY 2 UNITS UPPER
                TU2UUMIBackM = VMATERIAL 'SETS TALL UTILITY UPPER BACK MATERIAL TO WHITE MELAMINE
                TU2UUMIGableM = VMATERIAL 'SETS TALL UTILITY UPPER GABLE MATERIAL TO WHITE MELAMINE
                TU2UUMITopBtmM = VMATERIAL 'SETS TALL UTILITY UPPER TOP + BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMITopM = VMATERIAL 'SETS TALL UTILITY UPPER TOP MATERIAL TO WHITE MELAMINE
                TU2UUMIBotM = VMATERIAL 'SETS TALL UTILITY UPPER BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMIAdjShelfM = VMATERIAL 'SETS TALL UTILITY UPPER ADJUSTABLE SHELF BACK MATERIAL TO WHITE MELAMINE
                '# UPPER MATCHING INTERIOR #
                UMIBackM = VMATERIAL
                UMIGableM = VMATERIAL
                UMITopLightM = VMATERIAL
                UMIBtmM = VMATERIAL
                UMITopBtmM = VMATERIAL
                UMIAdjShelfM = VMATERIAL
                '# UPPER MICROWAVE MATCHING INTERIOR #
                UMMIBackM = VMATERIAL
                UMMIGableM = VMATERIAL
                UMMITopBtmM = VMATERIAL
                UMMIMicroSHM = VMATERIAL
                '# UPPER CORNER DIAGONAL MATCHING INTERIOR #
                UCDMILBackM = VMATERIAL
                UCDMISBackM = VMATERIAL
                UCDMIGableM = VMATERIAL
                UCDMITopBtmM = VMATERIAL
                UCDMIAdjShelfM = VMATERIAL
                '# UPPER WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                UWRBackM = VMATERIAL
                UWRGableM = VMATERIAL
                UWRTopBtmM = VMATERIAL
                UWRFFSHM = VMATERIAL
                UWRFShelfM = VMATERIAL
                UWRDivM = VMATERIAL
                '# UPPER END SHELF #
                UESLGableM = VMATERIAL
                UESSGableM = VMATERIAL
                UESTopM = VMATERIAL
                UESFixedShelfM = VMATERIAL
                '# BASE MATCHING INTERIOR #
                BMIBackM = VMATERIAL
                BMIGableM = VMATERIAL
                BMITopStrapM = VMATERIAL
                BMIBtmM = VMATERIAL
                BMIAdjShelfM = VMATERIAL
                BMITopM = VMATERIAL
                BMIStrapM = VMATERIAL
                BMIDividerM = VMATERIAL
                '# BASE PENINSULA OPEN #
                BPOPTopBtmFSHM = VMATERIAL
                BPOPSGableM = VMATERIAL
                BPOPLGableM = VMATERIAL
                '# VANITY MATCHING INTERIOR #
                VMIBackM = VMATERIAL
                VMIGableM = VMATERIAL
                VMITopStrapM = VMATERIAL
                VMITopM = VMATERIAL
                VMIBtmM = VMATERIAL
                VMIFFShelfM = VMATERIAL
                VMIStrapM = VMATERIAL
                VMIDividerM = VMATERIAL
                '# VANITY ELEVATED MATCHING INTERIOR #
                VEMIBackM = VMATERIAL
                VEMIGableM = VMATERIAL
                VEMITopStrapM = VMATERIAL
                VEMITopM = VMATERIAL
                VEMIBtmM = VMATERIAL
                VEMIFFShelfM = VMATERIAL
                VEMIStrapM = VMATERIAL
                VEMIDividerM = VMATERIAL
                '# BASE MICROWAVE OPEN SHELF 1 DRAWER #
                BMO1DUBackM = VMATERIAL
                BMO1DUGableM = VMATERIAL
                BMO1DUTopM = VMATERIAL
                BMO1DUBtmM = VMATERIAL
                '# BASE WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                BWRBackM = VMATERIAL
                BWRGableM = VMATERIAL
                BWRTopBtmM = VMATERIAL
                BWRFFSHM = VMATERIAL
                BWRValanceM = VSMATERIAL
                BWRVN = "Ordered Joe G."
                '# HUTCH MATCHING INTERIOR #
                HMIBackM = VMATERIAL
                HMIGableM = VMATERIAL
                HMITopFFShelfM = VMATERIAL
                HMIBtmFFShelfM = VMATERIAL
                HMITopM = VMATERIAL
                HMIBtmM = VMATERIAL
                HMIStrapM = VMATERIAL
                HMIDividerM = VMATERIAL
                HMIAdjShelfM = VMATERIAL
                '# CANOPY #
                RHFanSHM = "Ply"
                RHFrontM = VMATERIAL
                RHSFrontM = VMATERIAL
                RHSTopM = VMATERIAL
                RHIPanelM = VMATERIAL
                '# OVEN PANEL MATERIAL #
                OV1M = VMATERIAL
                '# FLOATING SHELF MATERIAL #
                FShelfM = VMATERIAL
                '# WINDOW PANEL MATERIAL #
                WPanelM = VMATERIAL
                '# WINDOW BOX MATERIAL #
                WTPanelM = VMATERIAL
                WFrontM = VMATERIAL
                '# DESK SUPPORT GABLE MATERIAL #
                DSGableM = VMATERIAL
                '# FANCY BRACKET MATERIAL #
                FBracketM = VMATERIAL
                '# FANCY VALANCE MATERIAL #
                FValanceM = VMATERIAL
                '# PLUMBING COVER PANEL MATERIAL #
                PCPanelM = VMATERIAL
                '# EDGE CODES VENEER AND PVC #
                VeneerEdgeCode = VEDGECODE
                PVCEdgeCode = PEDGECODE

            Case "WHITE ASH PVC"
                Dim VMATERIAL As String = "W.Ash Mel"
                Dim VSMATERIAL As String = "W.Ash PVC"
                Dim VEDGECODE As String = "White"
                Dim PEDGECODE As String = "White"
                'TALL UTILITY 2 UNITS UPPER
                TU2UUMIBackM = VMATERIAL 'SETS TALL UTILITY UPPER BACK MATERIAL TO WHITE MELAMINE
                TU2UUMIGableM = VMATERIAL 'SETS TALL UTILITY UPPER GABLE MATERIAL TO WHITE MELAMINE
                TU2UUMITopBtmM = VMATERIAL 'SETS TALL UTILITY UPPER TOP + BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMITopM = VMATERIAL 'SETS TALL UTILITY UPPER TOP MATERIAL TO WHITE MELAMINE
                TU2UUMIBotM = VMATERIAL 'SETS TALL UTILITY UPPER BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMIAdjShelfM = VMATERIAL 'SETS TALL UTILITY UPPER ADJUSTABLE SHELF BACK MATERIAL TO WHITE MELAMINE
                '# UPPER MATCHING INTERIOR #
                UMIBackM = VMATERIAL
                UMIGableM = VMATERIAL
                UMITopLightM = VMATERIAL
                UMIBtmM = VMATERIAL
                UMITopBtmM = VMATERIAL
                UMIAdjShelfM = VMATERIAL
                '# UPPER MICROWAVE MATCHING INTERIOR #
                UMMIBackM = VMATERIAL
                UMMIGableM = VMATERIAL
                UMMITopBtmM = VMATERIAL
                UMMIMicroSHM = VMATERIAL
                '# UPPER CORNER DIAGONAL MATCHING INTERIOR #
                UCDMILBackM = VMATERIAL
                UCDMISBackM = VMATERIAL
                UCDMIGableM = VMATERIAL
                UCDMITopBtmM = VMATERIAL
                UCDMIAdjShelfM = VMATERIAL
                '# UPPER WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                UWRBackM = VMATERIAL
                UWRGableM = VMATERIAL
                UWRTopBtmM = VMATERIAL
                UWRFFSHM = VMATERIAL
                UWRFShelfM = VMATERIAL
                UWRDivM = VMATERIAL
                '# UPPER END SHELF #
                UESLGableM = VMATERIAL
                UESSGableM = VMATERIAL
                UESTopM = VMATERIAL
                UESFixedShelfM = VMATERIAL
                '# BASE MATCHING INTERIOR #
                BMIBackM = VMATERIAL
                BMIGableM = VMATERIAL
                BMITopStrapM = VMATERIAL
                BMIBtmM = VMATERIAL
                BMIAdjShelfM = VMATERIAL
                BMITopM = VMATERIAL
                BMIStrapM = VMATERIAL
                BMIDividerM = VMATERIAL
                '# BASE PENINSULA OPEN #
                BPOPTopBtmFSHM = VMATERIAL
                BPOPSGableM = VMATERIAL
                BPOPLGableM = VMATERIAL
                '# VANITY MATCHING INTERIOR #
                VMIBackM = VMATERIAL
                VMIGableM = VMATERIAL
                VMITopStrapM = VMATERIAL
                VMITopM = VMATERIAL
                VMIBtmM = VMATERIAL
                VMIFFShelfM = VMATERIAL
                VMIStrapM = VMATERIAL
                VMIDividerM = VMATERIAL
                '# VANITY ELEVATED MATCHING INTERIOR #
                VEMIBackM = VMATERIAL
                VEMIGableM = VMATERIAL
                VEMITopStrapM = VMATERIAL
                VEMITopM = VMATERIAL
                VEMIBtmM = VMATERIAL
                VEMIFFShelfM = VMATERIAL
                VEMIStrapM = VMATERIAL
                VEMIDividerM = VMATERIAL
                '# BASE MICROWAVE OPEN SHELF 1 DRAWER #
                BMO1DUBackM = VMATERIAL
                BMO1DUGableM = VMATERIAL
                BMO1DUTopM = VMATERIAL
                BMO1DUBtmM = VMATERIAL
                '# BASE WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                BWRBackM = VMATERIAL
                BWRGableM = VMATERIAL
                BWRTopBtmM = VMATERIAL
                BWRFFSHM = VMATERIAL
                BWRValanceM = VSMATERIAL
                BWRVN = "Ordered Joe G."
                '# HUTCH MATCHING INTERIOR #
                HMIBackM = VMATERIAL
                HMIGableM = VMATERIAL
                HMITopFFShelfM = VMATERIAL
                HMIBtmFFShelfM = VMATERIAL
                HMITopM = VMATERIAL
                HMIBtmM = VMATERIAL
                HMIStrapM = VMATERIAL
                HMIDividerM = VMATERIAL
                HMIAdjShelfM = VMATERIAL
                '# CANOPY #
                RHFanSHM = "Ply"
                RHFrontM = VMATERIAL
                RHSFrontM = VMATERIAL
                RHSTopM = VMATERIAL
                RHIPanelM = VMATERIAL
                '# OVEN PANEL MATERIAL #
                OV1M = VMATERIAL
                '# FLOATING SHELF MATERIAL #
                FShelfM = VMATERIAL
                '# WINDOW PANEL MATERIAL #
                WPanelM = VMATERIAL
                '# WINDOW BOX MATERIAL #
                WTPanelM = VMATERIAL
                WFrontM = VMATERIAL
                '# DESK SUPPORT GABLE MATERIAL #
                DSGableM = VMATERIAL
                '# FANCY BRACKET MATERIAL #
                FBracketM = VMATERIAL
                '# FANCY VALANCE MATERIAL #
                FValanceM = VMATERIAL
                '# PLUMBING COVER PANEL MATERIAL #
                PCPanelM = VMATERIAL
                '# EDGE CODES VENEER AND PVC #
                VeneerEdgeCode = VEDGECODE
                PVCEdgeCode = PEDGECODE

            Case "WHITE CRYSTAL PVC"
                Dim VMATERIAL As String = "555 White"
                Dim VSMATERIAL As String = "W.Cr.PVC"
                Dim VEDGECODE As String = "White"
                Dim PEDGECODE As String = "White"
                'TALL UTILITY 2 UNITS UPPER
                TU2UUMIBackM = VMATERIAL 'SETS TALL UTILITY UPPER BACK MATERIAL TO WHITE MELAMINE
                TU2UUMIGableM = VMATERIAL 'SETS TALL UTILITY UPPER GABLE MATERIAL TO WHITE MELAMINE
                TU2UUMITopBtmM = VMATERIAL 'SETS TALL UTILITY UPPER TOP + BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMITopM = VMATERIAL 'SETS TALL UTILITY UPPER TOP MATERIAL TO WHITE MELAMINE
                TU2UUMIBotM = VMATERIAL 'SETS TALL UTILITY UPPER BOTTOM MATERIAL TO WHITE MELAMINE
                TU2UUMIAdjShelfM = VMATERIAL 'SETS TALL UTILITY UPPER ADJUSTABLE SHELF BACK MATERIAL TO WHITE MELAMINE
                '# UPPER MATCHING INTERIOR #
                UMIBackM = VMATERIAL
                UMIGableM = VMATERIAL
                UMITopLightM = VMATERIAL
                UMIBtmM = VMATERIAL
                UMITopBtmM = VMATERIAL
                UMIAdjShelfM = VMATERIAL
                '# UPPER MICROWAVE MATCHING INTERIOR #
                UMMIBackM = VMATERIAL
                UMMIGableM = VMATERIAL
                UMMITopBtmM = VMATERIAL
                UMMIMicroSHM = VMATERIAL
                '# UPPER CORNER DIAGONAL MATCHING INTERIOR #
                UCDMILBackM = VMATERIAL
                UCDMISBackM = VMATERIAL
                UCDMIGableM = VMATERIAL
                UCDMITopBtmM = VMATERIAL
                UCDMIAdjShelfM = VMATERIAL
                '# UPPER WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                UWRBackM = VMATERIAL
                UWRGableM = VMATERIAL
                UWRTopBtmM = VMATERIAL
                UWRFFSHM = VMATERIAL
                UWRFShelfM = VMATERIAL
                UWRDivM = VMATERIAL
                '# UPPER END SHELF #
                UESLGableM = VMATERIAL
                UESSGableM = VMATERIAL
                UESTopM = VMATERIAL
                UESFixedShelfM = VMATERIAL
                '# BASE MATCHING INTERIOR #
                BMIBackM = VMATERIAL
                BMIGableM = VMATERIAL
                BMITopStrapM = VMATERIAL
                BMIBtmM = VMATERIAL
                BMIAdjShelfM = VMATERIAL
                BMITopM = VMATERIAL
                BMIStrapM = VMATERIAL
                BMIDividerM = VMATERIAL
                '# BASE PENINSULA OPEN #
                BPOPTopBtmFSHM = VMATERIAL
                BPOPSGableM = VMATERIAL
                BPOPLGableM = VMATERIAL
                '# VANITY MATCHING INTERIOR #
                VMIBackM = VMATERIAL
                VMIGableM = VMATERIAL
                VMITopStrapM = VMATERIAL
                VMITopM = VMATERIAL
                VMIBtmM = VMATERIAL
                VMIFFShelfM = VMATERIAL
                VMIStrapM = VMATERIAL
                VMIDividerM = VMATERIAL
                '# VANITY ELEVATED MATCHING INTERIOR #
                VEMIBackM = VMATERIAL
                VEMIGableM = VMATERIAL
                VEMITopStrapM = VMATERIAL
                VEMITopM = VMATERIAL
                VEMIBtmM = VMATERIAL
                VEMIFFShelfM = VMATERIAL
                VEMIStrapM = VMATERIAL
                VEMIDividerM = VMATERIAL
                '# BASE MICROWAVE OPEN SHELF 1 DRAWER #
                BMO1DUBackM = VMATERIAL
                BMO1DUGableM = VMATERIAL
                BMO1DUTopM = VMATERIAL
                BMO1DUBtmM = VMATERIAL
                '# BASE WINE RACK LATTICE, UPPER WINE RACK LATTICE HALF #
                BWRBackM = VMATERIAL
                BWRGableM = VMATERIAL
                BWRTopBtmM = VMATERIAL
                BWRFFSHM = VMATERIAL
                BWRValanceM = VSMATERIAL
                BWRVN = "Ordered Joe G."
                '# HUTCH MATCHING INTERIOR #
                HMIBackM = VMATERIAL
                HMIGableM = VMATERIAL
                HMITopFFShelfM = VMATERIAL
                HMIBtmFFShelfM = VMATERIAL
                HMITopM = VMATERIAL
                HMIBtmM = VMATERIAL
                HMIStrapM = VMATERIAL
                HMIDividerM = VMATERIAL
                HMIAdjShelfM = VMATERIAL
                '# CANOPY #
                RHFanSHM = "Ply"
                RHFrontM = VMATERIAL
                RHSFrontM = VMATERIAL
                RHSTopM = VMATERIAL
                RHIPanelM = VMATERIAL
                '# OVEN PANEL MATERIAL #
                OV1M = VMATERIAL
                '# FLOATING SHELF MATERIAL #
                FShelfM = VMATERIAL
                '# WINDOW PANEL MATERIAL #
                WPanelM = VMATERIAL
                '# WINDOW BOX MATERIAL #
                WTPanelM = VMATERIAL
                WFrontM = VMATERIAL
                '# DESK SUPPORT GABLE MATERIAL #
                DSGableM = VMATERIAL
                '# FANCY BRACKET MATERIAL #
                FBracketM = VMATERIAL
                '# FANCY VALANCE MATERIAL #
                FValanceM = VMATERIAL
                '# PLUMBING COVER PANEL MATERIAL #
                PCPanelM = VMATERIAL
                '# EDGE CODES VENEER AND PVC #
                VeneerEdgeCode = VEDGECODE
                PVCEdgeCode = PEDGECODE
            Case Else
        End Select

        '#####################################
        '# STANDARD VENEER DOOR FINISH RULES #
        '#####################################
        '
        Dim VFinish = CKWOPLANNERSDoorFinishBox
        Select Case VFinish

            Case "BLACK SEMIGLOSS"
                PVCEdgeCode = "Black"
                VeneerEdgeCode = "Black"

            Case "GREY LIGHT"
                PVCEdgeCode = "105"
                VeneerEdgeCode = "M/V"

            Case "OC-26"
                PVCEdgeCode = "112"
                VeneerEdgeCode = "M/V"

            Case "2111-60"
                PVCEdgeCode = "190"
                VeneerEdgeCode = "M/V"

            Case "CARAMEL MAPLE", "CARAMEL OAK"
                PVCEdgeCode = "205"

            Case "FRUITWOOD OAK"
                PVCEdgeCode = "386"

            Case "TULIP MAPLE"
                PVCEdgeCode = "1037"

            Case "MANGO MAPLE"
                PVCEdgeCode = "1117"

            Case "CHARCOAL MAPLE", "CHARCOAL OAK", "CHOCOLATE MAPLE", "CHOCOLATE OAK", "EBONY MAPLE", "EBONY OAK"
                PVCEdgeCode = "1130"

            Case "GREY STONE MAPLE", "GREY STONE OAK"
                PVCEdgeCode = "1156"

            Case "2124-30", "CSP-110"
                PVCEdgeCode = "1315"
                VeneerEdgeCode = "M/V"

            Case "ANTIQUE WHITE", "CC-490", "CREAM MATTE", "DOVE MATTE", "DOVE WHITE", "OC-20", "OC-46"
                PVCEdgeCode = "1438"
                VeneerEdgeCode = "M/V"

            Case "TULIP OAK"
                PVCEdgeCode = "2567"

            Case "PECAN MAPLE", "PECAN OAK"
                PVCEdgeCode = "2668"

            Case "GINGER MAPLE", "GINGER OAK", "MOCHA MAPLE", "MOCHA OAK"
                PVCEdgeCode = "5038"

            Case "GRAPHITE MAPLE", "GRAPHITE OAK", "SLATE MAPLE", "SLATE OAK"
                PVCEdgeCode = "5741"

            Case "NATURAL MAPLE", "NUTMEG MAPLE"
                PVCEdgeCode = "6116"

            Case "BUTTERNUT MAPLE", "BUTTERNUT OAK"
                PVCEdgeCode = "6117"

            Case "ESPRESSO MAPLE", "ESPRESSO OAK"
                PVCEdgeCode = "6119"

            Case "ASPEN MAPLE", "ASPEN OAK"
                PVCEdgeCode = "6120"

            Case "KHAKI MAPLE", "KHAKI OAK"
                PVCEdgeCode = "6120"

            Case "BROWN CHERRY CHERRY", "COCOA MAPLE", "COCOA OAK", "WALNUT MAPLE"
                PVCEdgeCode = "6121"

            Case "ALMOND MATTE", "MUSHROOM"
                PVCEdgeCode = "7002"
                VeneerEdgeCode = "M/V"

            Case "2126-60", "2134-60", "FOG GREY"
                PVCEdgeCode = "7006"
                VeneerEdgeCode = "M/V"

            Case "AF-95"
                PVCEdgeCode = "7043"
                VeneerEdgeCode = "M/V"

            Case "GREY DARK"
                PVCEdgeCode = "7071"
                VeneerEdgeCode = "M/V"

            Case "CARAMEL PINE"
                PVCEdgeCode = "7806"

            Case "NATURAL OAK", "NUTMEG OAK"
                PVCEdgeCode = "8001"

            Case "PEACH OAK"
                PVCEdgeCode = "8050"

            Case "CHESTNUT CHERRY", "NOCE WALNUT", "WILD WALNUT MAPLE"
                PVCEdgeCode = "8103"

            Case "ANTIQUE BROWN MAPLE", "ANTIQUE BROWN OAK", "CINNAMON CHERRY"
                PVCEdgeCode = "8114"

            Case "MEDIUM BROWN MAPLE", "MEDIUM BROWN OAK"
                PVCEdgeCode = "8336"

            Case "HARVEST MAPLE"
                PVCEdgeCode = "8597"

            Case "WHITE WASH OAK"
                PVCEdgeCode = "8670"

            Case "OLIVE MAPLE", "OLIVE OAK"
                PVCEdgeCode = "8693"

            Case "PEACH MAPLE", "WHITE WASH MAPLE"
                PVCEdgeCode = "8699"

            Case "COGNAC CHERRY"
                PVCEdgeCode = "8725"

            Case "NATURAL CHERRY"
                PVCEdgeCode = "8876"

            Case "HARVEST OAK"
                PVCEdgeCode = "8883"

            Case "BURGUNDY MAPLE", "BURGUNDY OAK", "RASBERRY MAPLE", "RASBERRY OAK", "RED CHERRY CHERRY"
                PVCEdgeCode = "9774"

            Case "CAFE MAPLE"
                PVCEdgeCode = "20416"

            Case "TOFFEE MAPLE", "TOFFEE OAK"
                PVCEdgeCode = "40439"

            Case "PEARL SEMI-GLOSS", "WHITE MATTE MDF"
                PVCEdgeCode = "WHITE"
                VeneerEdgeCode = "M/V"

            Case "CC-30", "CC-40", "OC-17", "OC-25", "OC-30", "OC-65"
                PVCEdgeCode = "WHITE"
                VeneerEdgeCode = "WHITE"

            Case "ESCARPMENT", "CC-518"
                PVCEdgeCode = "7213"
            Case Else
        End Select

        '####################################
        '# UPPER EDGING RULES AND VARIABLES #
        '####################################
        '
        Dim UGX As Double
        Dim UGY As Double
        UGX = UGableX
        UGY = UGableY

        Dim UTX As Double
        Dim UTY As Double
        UTX = UTopX
        UTY = UTopY

        Dim UTHX As Double
        Dim UTHY As Double
        UTHX = UTopHingeX
        UTHY = UTopHingeY

        Dim UTLX As Double
        Dim UTLY As Double
        UTLX = UTopLightX
        UTLY = UTopLightY

        Dim UBX As Double
        Dim UBY As Double
        UBX = UBtmX
        UBY = UBtmY

        Dim UTBX As Double
        Dim UTBY As Double
        UTBX = UTopBtmX
        UTBY = UTopBtmY

        Dim UTFDX As Double
        Dim UTFDY As Double
        UTFDX = UFDividerX
        UTFDY = UFDividerY

        Dim UTDX As Double
        Dim UTDY As Double
        UTDX = UDividerX
        UTDY = UDividerY

        Dim UASX As Double
        Dim UASY As Double
        UASX = UAdjShelfX
        UASY = UAdjShelfY

        '#############################
        '# UPPER GABLE EDGE SEQUENCE #
        '#############################
        '
        If (UGX > UGY) Then
            UGEdgeSeq = "E2S"
            UGEdgeCode = BMEdgeCode
            UGEdgeSeq2 = "E1L"
            UGEdgeCode2 = PVCEdgeCode
        End If
        If (UGX < UGY) Then
            UGEdgeSeq = "E2L"
            UGEdgeCode = BMEdgeCode
            UGEdgeSeq2 = "E1S"
            UGEdgeCode2 = PVCEdgeCode
        End If

        '#######################################
        '# UPPER TOP + FAN SHELF EDGE SEQUENCE #
        '#######################################
        '
        '###############
        '# GROUP RULES #
        '###############
        '
        If (CKWOPLANNERSGroupBox = "GROUP1") Then

            If (UTX > UTY) Then
                UTEdgeSeq = "E1L"
                UTEdgeCode = BMEdgeCode
            End If

            If (UTX < UTY) Then
                UTEdgeSeq = "E1S"
                UTEdgeCode = BMEdgeCode
            End If

        Else

            If (UTX > UTY) Then
                UTEdgeSeq = "E1L"
                UTEdgeCode = PVCEdgeCode
            End If

            If (UTX < UTY) Then
                UTEdgeSeq = "E1S"
                UTEdgeCode = PVCEdgeCode
            End If

        End If

        '#################################
        '# UPPER TOP HINGE EDGE SEQUENCE #
        '#################################
        '
        '###############
        '# GROUP RULES #
        '###############
        '
        If (CKWOPLANNERSGroupBox = "GROUP1") Then

            If (UTHX > UTHY) Then
                UTHEdgeSeq = "E1L"
                UTHEdgeCode = BMEdgeCode
            End If

            If (UTHX < UTHY) Then
                UTHEdgeSeq = "E1S"
                UTHEdgeCode = BMEdgeCode
            End If

        Else

            If (UTHX > UTHY) Then
                UTHEdgeSeq = "E1L"
                UTHEdgeCode = PVCEdgeCode
            End If

            If (UTHX < UTHY) Then
                UTHEdgeSeq = "E1S"
                UTHEdgeCode = PVCEdgeCode
            End If

        End If

        '#################################
        '# UPPER TOP LIGHT EDGE SEQUENCE #
        '#################################
        '
        '###############
        '# GROUP RULES #
        '###############
        '
        If (CKWOPLANNERSGroupBox = "GROUP1") Then

            If (UTLX > UTLY) Then
                UTLEdgeSeq = "E1L"
                UTLEdgeCode = BMEdgeCode
            End If

            If (UTLX < UTLY) Then
                UTLEdgeSeq = "E1S"
                UTLEdgeCode = BMEdgeCode
            End If

        Else

            If (UTLX > UTLY) Then
                UTLEdgeSeq = "E1L"
                UTLEdgeCode = PVCEdgeCode
            End If

            If (UTLX < UTLY) Then
                UTLEdgeSeq = "E1S"
                UTLEdgeCode = PVCEdgeCode
            End If

        End If

        '##############################
        '# UPPER BOTTOM EDGE SEQUENCE #
        '##############################
        '
        '###############
        '# GROUP RULES #
        '###############
        '
        If (CKWOPLANNERSGroupBox = "GROUP1") Then

            If (UBX > UBY) Then
                UBEdgeSeq = "E1L"
                UBEdgeCode = BMEdgeCode
            End If

            If (UBX < UBY) Then
                UBEdgeSeq = "E1S"
                UBEdgeCode = BMEdgeCode
            End If

        Else

            If (UBX > UBY) Then
                UBEdgeSeq = "E1L"
                UBEdgeCode = PVCEdgeCode
            End If

            If (UBX < UBY) Then
                UBEdgeSeq = "E1S"
                UBEdgeCode = PVCEdgeCode
            End If

        End If

        '####################################
        '# UPPER TOP + BOTTOM EDGE SEQUENCE #
        '####################################
        '
        '###############
        '# GROUP RULES #
        '###############
        '
        If (CKWOPLANNERSGroupBox = "GROUP1") Then

            If (UTBX > UTBY) Then
                UTBEdgeSeq = "E1L"
                UTBEdgeCode = BMEdgeCode
            End If

            If (UTBX < UTBY) Then
                UTBEdgeSeq = "E1S"
                UTBEdgeCode = BMEdgeCode
            End If

        Else

            If (UTBX > UTBY) Then
                UTBEdgeSeq = "E1L"
                UTBEdgeCode = PVCEdgeCode
            End If

            If (UTBX < UTBY) Then
                UTBEdgeSeq = "E1S"
                UTBEdgeCode = PVCEdgeCode
            End If

        End If

        '##########################################
        '# UPPER TRAY FIXED DIVIDER EDGE SEQUENCE #
        '##########################################
        '
        '###############
        '# GROUP RULES #
        '###############
        '
        If (CKWOPLANNERSGroupBox = "GROUP1") Then

            If (UTFDX > UTFDY) Then
                UTFDEdgeSeq = "E1L"
                UTFDEdgeCode = BMEdgeCode
            End If

            If (UTFDX < UTFDY) Then
                UTFDEdgeSeq = "E1S"
                UTFDEdgeCode = BMEdgeCode
            End If

        Else

            If (UTFDX > UTFDY) Then
                UTFDEdgeSeq = "E1L"
                UTFDEdgeCode = PVCEdgeCode
            End If

            If (UTFDX < UTFDY) Then
                UTFDEdgeSeq = "E1S"
                UTFDEdgeCode = PVCEdgeCode
            End If

        End If

        '####################################
        '# UPPER TRAY DIVIDER EDGE SEQUENCE #
        '####################################
        '
        '###############
        '# GROUP RULES #
        '###############
        '
        If (CKWOPLANNERSGroupBox = "GROUP1") Then

            If (UTDX > UTDY) Then
                UTDEdgeSeq = "E1L"
                UTDEdgeCode = BMEdgeCode
            End If

            If (UTDX < UTDY) Then
                UTDEdgeSeq = "E1S"
                UTDEdgeCode = BMEdgeCode
            End If

        Else

            If (UTDX > UTDY) Then
                UTDEdgeSeq = "E1L"
                UTDEdgeCode = PVCEdgeCode
            End If

            If (UTDX < UTDY) Then
                UTDEdgeSeq = "E1S"
                UTDEdgeCode = PVCEdgeCode
            End If

        End If

        '########################################
        '# UPPER ADJUSTABLE SHELF EDGE SEQUENCE #
        '########################################
        '
        If (UASX > UASY) Then
            UASEdgeSeq = "E1L"
            UASEdgeCode = BMEdgeCode
        End If

        If (UASX < UASY) Then
            UASEdgeSeq = "E1S"
            UASEdgeCode = BMEdgeCode
        End If

        '######################################################
        '# UPPER MATCHING INTERIOR EDGING RULES AND VARIABLES #
        '######################################################
        '
        Dim UMIGX As Double
        Dim UMIGY As Double
        UMIGX = UMIGableX
        UMIGY = UMIGableY

        Dim UMITLX As Double
        Dim UMITLY As Double
        UMITLX = UMITopLightX
        UMITLY = UMITopLightY

        Dim UMIBX As Double
        Dim UMIBY As Double
        UMIBX = UMIBtmX
        UMIBY = UMIBtmY

        Dim UMITBX As Double
        Dim UMITBY As Double
        UMITBX = UMITopBtmX
        UMITBY = UMITopBtmY

        Dim UMIASX As Double
        Dim UMIASY As Double
        UMIASX = UMIAdjShelfX
        UMIASY = UMIAdjShelfY

        '###############################################
        '# UPPER MATCHING INTERIOR GABLE EDGE SEQUENCE #
        '###############################################
        '
        If (UMIGX > UMIGY) Then
            UMIGEdgeSeq = "E2S"
            UMIGEdgeCode = VeneerEdgeCode
            UMIGEdgeSeq2 = "E1L"
            UMIGEdgeCode2 = VeneerEdgeCode
        End If

        If (UMIGX < UMIGY) Then
            UMIGEdgeSeq = "E2L"
            UMIGEdgeCode = VeneerEdgeCode
            UMIGEdgeSeq2 = "E1S"
            UMIGEdgeCode2 = VeneerEdgeCode
        End If

        '###################################################
        '# UPPER MATCHING INTERIOR TOP LIGHT EDGE SEQUENCE #
        '###################################################
        '

        '###############
        '# GROUP RULES #
        '###############
        '
        If (CKWOPLANNERSGroupBox = "GROUP1") Then

            If (UMITLX > UMITLY) Then
                UMITLEdgeSeq = "E1L"
                UMITLEdgeCode = VeneerEdgeCode
            End If

            If (UMITLX < UMITLY) Then
                UMITLEdgeSeq = "E1S"
                UMITLEdgeCode = VeneerEdgeCode
            End If

        Else

            If (UMITLX > UMITLY) Then
                UMITLEdgeSeq = "E1L"
                UMITLEdgeCode = VeneerEdgeCode
            End If

            If (UMITLX < UMITLY) Then
                UMITLEdgeSeq = "E1S"
                UMITLEdgeCode = VeneerEdgeCode
            End If

        End If

        '################################################
        '# UPPER MATCHING INTERIOR BOTTOM EDGE SEQUENCE #
        '################################################
        '

        '###############
        '# GROUP RULES #
        '###############
        '
        If (CKWOPLANNERSGroupBox = "GROUP1") Then

            If (UMIBX > UMIBY) Then
                UMIBEdgeSeq = "E1L"
                UMIBEdgeCode = VeneerEdgeCode
            End If

            If (UMIBX < UMIBY) Then
                UMIBEdgeSeq = "E1S"
                UMIBEdgeCode = VeneerEdgeCode
            End If

        Else

            If (UMIBX > UMIBY) Then
                UMIBEdgeSeq = "E1L"
                UMIBEdgeCode = VeneerEdgeCode
            End If

            If (UMIBX < UMIBY) Then
                UMIBEdgeSeq = "E1S"
                UMIBEdgeCode = VeneerEdgeCode
            End If

        End If

        '######################################################
        '# UPPER MATCHING INTERIOR TOP + BOTTOM EDGE SEQUENCE #
        '######################################################
        '

        '###############
        '# GROUP RULES #
        '###############
        '
        If (CKWOPLANNERSGroupBox = "GROUP1") Then

            If (UMITBX > UMITBY) Then
                UMITBEdgeSeq = "E1L"
                UMITBEdgeCode = VeneerEdgeCode
            End If

            If (UMITBX < UMITBY) Then
                UMITBEdgeSeq = "E1S"
                UMITBEdgeCode = VeneerEdgeCode
            End If

        Else

            If (UMITBX > UMITBY) Then
                UMITBEdgeSeq = "E1L"
                UMITBEdgeCode = VeneerEdgeCode
            End If

            If (UMITBX < UMITBY) Then
                UMITBEdgeSeq = "E1S"
                UMITBEdgeCode = VeneerEdgeCode
            End If

        End If

        '##########################################################
        '# UPPER MATCHING INTERIOR ADJUSTABLE SHELF EDGE SEQUENCE #
        '##########################################################
        '
        If (UMIASX > UMIASY) Then
            UMIASEdgeSeq = "E1L"
            UMIASEdgeCode = VeneerEdgeCode
        End If

        If (UMIASX < UMIASY) Then
            UMIASEdgeSeq = "E1S"
            UMIASEdgeCode = VeneerEdgeCode
        End If


        '####################################
        '# UPPER EDGING RULES AND VARIABLES #
        '####################################
        '
        Dim UMGX As Double
        Dim UMGY As Double
        UMGX = UMGableX
        UMGY = UMGableY

        Dim UMTX As Double
        Dim UMTY As Double
        UMTX = UMTopX
        UMTY = UMTopY

        Dim UMBX As Double
        Dim UMBY As Double
        UMBX = UMBtmX
        UMBY = UMBtmY

        Dim UMASX As Double
        Dim UMASY As Double
        UMASX = UMAdjShelfX
        UMASY = UMAdjShelfY

        '#############################
        '# UPPER GABLE EDGE SEQUENCE #
        '#############################
        '
        If (UMGX > UMGY) Then
            UMGEdgeSeq = "E2S"
            UMGEdgeCode = BMEdgeCode
            UMGEdgeSeq2 = "E1L"
            UMGEdgeCode2 = PVCEdgeCode
        End If

        If (UMGX < UMGY) Then
            UMGEdgeSeq = "E2L"
            UMGEdgeCode = BMEdgeCode
            UMGEdgeSeq2 = "E1S"
            UMGEdgeCode2 = PVCEdgeCode
        End If

        '####################################
        '# UPPER TOP + BOTTOM EDGE SEQUENCE #
        '####################################
        '
        '###############
        '# GROUP RULES #
        '###############
        '
        If (CKWOPLANNERSGroupBox = "GROUP1") Then

            If (UMTX > UMTY) Then
                UMTEdgeSeq = "E1L"
                UMTEdgeCode = BMEdgeCode
            End If

            If (UMTX < UMTY) Then
                UMTEdgeSeq = "E1S"
                UMTEdgeCode = BMEdgeCode
            End If

            If (UMBX > UMBY) Then
                UMBEdgeSeq = "E1L"
                UMBEdgeCode = PVCEdgeCode
            End If

            If (UMBX < UMBY) Then
                UMBEdgeSeq = "E1S"
                UMBEdgeCode = PVCEdgeCode
            End If

        Else

            If (UMTX > UMTY) Then
                UMTEdgeSeq = "E1L"
                UMTEdgeCode = PVCEdgeCode
            End If

            If (UMTX < UMTY) Then
                UMTEdgeSeq = "E1S"
                UMTEdgeCode = PVCEdgeCode
            End If

            If (UMBX > UMBY) Then
                UMBEdgeSeq = "E1L"
                UMBEdgeCode = PVCEdgeCode
            End If
            If (UMBX < UMBY) Then
                UMBEdgeSeq = "E1S"
                UMBEdgeCode = PVCEdgeCode
            End If

        End If

        '########################################
        '# UPPER ADJUSTABLE SHELF EDGE SEQUENCE #
        '########################################
        '
        If (UMASX > UMASY) Then
            UMASEdgeSeq = "E1L"
            UMASEdgeCode = BMEdgeCode
        End If

        If (UMASX < UMASY) Then
            UMASEdgeSeq = "E1S"
            UMASEdgeCode = BMEdgeCode
        End If

        '######################################################
        '# UPPER MATCHING INTERIOR EDGING RULES AND VARIABLES #
        '######################################################
        '
        Dim UMMIGX As Double
        Dim UMMIGY As Double
        UMMIGX = UMMIGableX
        UMMIGY = UMMIGableY

        Dim UMMITBX As Double
        Dim UMMITBY As Double
        UMMITBX = UMMITopBtmX
        UMMITBY = UMMITopBtmY

        Dim UMMIASX As Double
        Dim UMMIASY As Double
        UMMIASX = UMMIMicroSHX
        UMMIASY = UMMIMicroSHY

        '###############################################
        '# UPPER MATCHING INTERIOR GABLE EDGE SEQUENCE #
        '###############################################
        '
        If (UMMIGX > UMMIGY) Then
            UMMIGEdgeSeq = "E2S"
            UMMIGEdgeCode = VeneerEdgeCode
            UMMIGEdgeSeq2 = "E1L"
            UMMIGEdgeCode2 = VeneerEdgeCode
        End If

        If (UMMIGX < UMMIGY) Then
            UMMIGEdgeSeq = "E2L"
            UMMIGEdgeCode = VeneerEdgeCode
            UMMIGEdgeSeq2 = "E1S"
            UMMIGEdgeCode2 = VeneerEdgeCode
        End If

        '######################################################
        '# UPPER MATCHING INTERIOR TOP + BOTTOM EDGE SEQUENCE #
        '######################################################
        '

        '###############
        '# GROUP RULES #
        '###############
        '
        If (CKWOPLANNERSGroupBox = "GROUP1") Then

            If (UMMITBX > UMMITBY) Then
                UMMITBEdgeSeq = "E1L"
                UMMITBEdgeCode = VeneerEdgeCode
            End If

            If (UMMITBX < UMMITBY) Then
                UMMITBEdgeSeq = "E1S"
                UMMITBEdgeCode = VeneerEdgeCode
            End If

        Else

            If (UMMITBX > UMMITBY) Then
                UMMITBEdgeSeq = "E1L"
                UMMITBEdgeCode = VeneerEdgeCode
            End If

            If (UMMITBX < UMMITBY) Then
                UMMITBEdgeSeq = "E1S"
                UMMITBEdgeCode = VeneerEdgeCode
            End If

        End If

        '##########################################################
        '# UPPER MATCHING INTERIOR ADJUSTABLE SHELF EDGE SEQUENCE #
        '##########################################################
        '
        If (UMMIASX > UMMIASY) Then
            UMMIMSEdgeSeq = "E2S"
            UMMIMSEdgeCode = VeneerEdgeCode
            UMMIMSEdgeSeq2 = "E1L"
            UMMIMSEdgeCode2 = VeneerEdgeCode
        End If

        If (UMMIASX < UMMIASY) Then
            UMMIMSEdgeSeq = "E2L"
            UMMIMSEdgeCode = VeneerEdgeCode
            UMMIMSEdgeSeq2 = "E1S"
            UMMIMSEdgeCode2 = VeneerEdgeCode
        End If


        '####################################################
        '# UPPER CORNER DIAGONAL EDGING RULES AND VARIABLES #
        '####################################################
        '
        Dim UCDGX As Double
        Dim UCDGY As Double
        UCDGX = UCDGableX
        UCDGY = UCDGableY

        Dim UCDTBX As Double
        Dim UCDTBY As Double
        UCDTBX = UCDTopBtmX
        UCDTBY = UCDTopBtmY

        Dim UCDASX As Double
        Dim UCDASY As Double
        UCDASX = UCDAdjShelfX
        UCDASY = UCDAdjShelfY

        '#############################################
        '# UPPER CORNER DIAGONAL GABLE EDGE SEQUENCE #
        '#############################################
        '
        If (CKWOPLANNERSSpeciesBox = "PVC") Then

            If (UCDGX > UCDGY) Then
                UCDGEdgeSeq = "E2S"
                UCDGEdgeCode = BMEdgeCode
                UCDGEdgeSeq2 = "E1L"
                UCDGEdgeCode2 = PVCEdgeCode
            End If

            If (UCDGX < UCDGY) Then
                UCDGEdgeSeq = "E2L"
                UCDGEdgeCode = BMEdgeCode
                UCDGEdgeSeq2 = "E1S"
                UCDGEdgeCode2 = PVCEdgeCode
            End If

        Else

            If (UCDGX > UCDGY) Then
                UCDGEdgeSeq = "E2S"
                UCDGEdgeCode = BMEdgeCode
                UCDGEdgeSeq2 = "E1L"
                UCDGEdgeCode2 = VeneerEdgeCode
            End If

            If (UCDGX < UCDGY) Then
                UCDGEdgeSeq = "E2L"
                UCDGEdgeCode = BMEdgeCode
                UCDGEdgeSeq2 = "E1S"
                UCDGEdgeCode2 = VeneerEdgeCode
            End If

        End If

        '####################################################
        '# UPPER CORNER DIAGONAL TOP + BOTTOM EDGE SEQUENCE #
        '####################################################
        '
        If (CKWOPLANNERSGroupBox = "GROUP1") Then
            UCDTBEdgeSeq = "E1M"
            UCDTBEdgeCode = BMEdgeCode
        Else
            UCDTBEdgeSeq = "E1M"
            UCDTBEdgeCode = PVCEdgeCode
        End If

        '########################################################
        '# UPPER CORNER DIAGONAL ADJUSTABLE SHELF EDGE SEQUENCE #
        '########################################################
        '
        UCDASEdgeSeq = "E1M"
        UCDASEdgeCode = BMEdgeCode


        '######################################################################
        '# UPPER CORNER DIAGONAL MATCHING INTERIOR EDGING RULES AND VARIABLES #
        '######################################################################
        '
        Dim UCDMIGX As Double
        Dim UCDMIGY As Double
        UCDMIGX = UCDMIGableX
        UCDMIGY = UCDMIGableY

        Dim UCDMITBX As Double
        Dim UCDMITBY As Double
        UCDMITBX = UCDMITopBtmX
        UCDMITBY = UCDMITopBtmY

        '###############################################
        '# UPPER MATCHING INTERIOR GABLE EDGE SEQUENCE #
        '###############################################
        '
        If (UCDMIGX > UCDMIGY) Then
            UCDMIGEdgeSeq = "E2S"
            UCDMIGEdgeCode = VeneerEdgeCode
            UCDMIGEdgeSeq2 = "E1L"
            UCDMIGEdgeCode2 = VeneerEdgeCode
        End If

        If (UCDMIGX < UCDMIGY) Then
            UCDMIGEdgeSeq = "E2L"
            UCDMIGEdgeCode = VeneerEdgeCode
            UCDMIGEdgeSeq2 = "E1S"
            UCDMIGEdgeCode2 = VeneerEdgeCode
        End If

        '######################################################
        '# UPPER MATCHING INTERIOR TOP + BOTTOM EDGE SEQUENCE #
        '######################################################
        '
        UCDMITBEdgeSeq = "E1M"
        UCDMITBEdgeCode = VeneerEdgeCode

        '##########################################################
        '# UPPER MATCHING INTERIOR ADJUSTABLE SHELF EDGE SEQUENCE #
        '##########################################################
        '
        UCDMIASEdgeSeq = "E1M"
        UCDMIASEdgeCode = VeneerEdgeCode

        '####################################################
        '# UPPER END SHELF UPPER EDGING RULES AND VARIABLES #
        '####################################################
        '
        Dim UESLGX As Double = UESLGableX
        Dim UESLGY As Double = UESLGableY

        Dim UESSGX As Double = UESSGableX
        Dim UESSGY As Double = UESSGableY

        Dim UESTX As Double = UESTopX
        Dim UESTY As Double = UESTopY

        Dim UESFSHX As Double = UESFixedShelfX
        Dim UESFSHY As Double = UESFixedShelfY

        '###############################################
        '# UPPER END SHELF UPPER L-GABLE EDGE SEQUENCE #
        '###############################################
        '
        If (UESLGX > UESLGY) Then
            UESLGEdgeSeq = "E2S"
            UESLGEdgeCode = VeneerEdgeCode
            UESLGEdgeSeq2 = "E1L"
            UESLGEdgeCode2 = VeneerEdgeCode
        End If

        If (UESLGX < UESLGY) Then
            UESLGEdgeSeq = "E2L"
            UESLGEdgeCode = VeneerEdgeCode
            UESLGEdgeSeq2 = "E1S"
            UESLGEdgeCode2 = VeneerEdgeCode
        End If

        '###############################################
        '# UPPER END SHELF UPPER S-GABLE EDGE SEQUENCE #
        '###############################################
        '
        If (UESSGX > UESSGY) Then
            UESSGEdgeSeq = "E2S"
            UESSGEdgeCode = VeneerEdgeCode
            UESSGEdgeSeq2 = "E1L"
            UESSGEdgeCode2 = VeneerEdgeCode
        End If

        If (UESSGX < UESSGY) Then
            UESSGEdgeSeq = "E2L"
            UESSGEdgeCode = VeneerEdgeCode
            UESSGEdgeSeq2 = "E1S"
            UESSGEdgeCode2 = VeneerEdgeCode
        End If

        '###################################################
        '# UPPER END SHELF TOP + FIXED SHELF EDGE SEQUENCE #
        '###################################################
        '
        If (CutlistForm.FShBox.Text = "DIAGONAL") Then
            UESTEdgeSeq = "E3S"
            UESTEdgeCode = VeneerEdgeCode
            UESFSHEdgeSeq = "E3S"
            UESFSHEdgeCode = VeneerEdgeCode
        End If

        If (CutlistForm.FShBox.Text = "ROUND") Then
            UESTEdgeSeq = "E1L"
            UESTEdgeCode = VeneerEdgeCode
            UESFSHEdgeSeq = "E1L"
            UESFSHEdgeCode = VeneerEdgeCode
        End If

        If (CutlistForm.FShBox.Text = "TRIANGLE") Then
            UESTEdgeSeq = "E1M"
            UESTEdgeCode = VeneerEdgeCode
            UESFSHEdgeSeq = "E1M"
            UESFSHEdgeCode = VeneerEdgeCode
        End If

        '###################################
        '# BASE EDGING RULES AND VARIABLES #
        '###################################
        '
        Dim BGX As Double
        Dim BGY As Double
        BGX = BGableX
        BGY = BGableY

        Dim BTSX As Double
        Dim BTSY As Double
        BTSX = BTopStrapX
        BTSY = BTopStrapY

        Dim BRTX As Double
        Dim BRTY As Double
        BRTX = BTopX
        BRTY = BTopY

        Dim BRBX As Double
        Dim BRBY As Double
        BRBX = BBtmX
        BRBY = BBtmY

        Dim BTX As Double
        Dim BTY As Double
        BTX = BTopX
        BTY = BTopY

        Dim BBX As Double
        Dim BBY As Double
        BBX = BBtmX
        BBY = BBtmY

        Dim BASX As Double
        Dim BASY As Double
        BASX = BAdjShelfX
        BASY = BAdjShelfY

        Dim BSX As Double
        Dim BSY As Double
        BSX = BStrapX
        BSY = BStrapY

        Dim BDX As Double
        Dim BDY As Double
        BDX = BDividerX
        BDY = BDividerY

        Dim BFFSX As Double
        Dim BFFSY As Double
        BFFSX = BFFShelfX
        BFFSY = BFFShelfY

        '############################
        '# BASE GABLE EDGE SEQUENCE #
        '############################
        '
        BBackEdgeSeq = "E1S"
        BBackEdgeCode = BMEdgeCode

        '############################
        '# BASE GABLE EDGE SEQUENCE #
        '############################
        '
        If (BGX > BGY) Then
            BGEdgeSeq = "E1L"
            BGEdgeCode = PVCEdgeCode
        End If
        If (BGX < BGY) Then
            BGEdgeSeq = "E1S"
            BGEdgeCode = PVCEdgeCode
        End If

        '################################
        '# BASE TOP STRAP EDGE SEQUENCE #
        '################################
        '
        '###############
        '# GROUP RULES #
        '###############
        '
        If (CKWOPLANNERSGroupBox = "GROUP1") Then

            If (BTSX > BTSY) Then
                BTSEdgeSeq = "E1L"
                BTSEdgeCode = BMEdgeCode
            End If

            If (BTSX < BTSY) Then
                BTSEdgeSeq = "E1S"
                BTSEdgeCode = BMEdgeCode
            End If

        Else

            If (BTSX > BTSY) Then
                BTSEdgeSeq = "E1L"
                BTSEdgeCode = PVCEdgeCode
            End If

            If (BTSX < BTSY) Then
                BTSEdgeSeq = "E1S"
                BTSEdgeCode = PVCEdgeCode
            End If

        End If

        '##########################
        '# BASE TOP EDGE SEQUENCE #
        '##########################
        '
        '###############
        '# GROUP RULES #
        '###############
        '
        If (CKWOPLANNERSGroupBox = "GROUP1") Then

            If (BTX > BTY) Then
                BTEdgeSeq = "E1L"
                BTEdgeCode = BMEdgeCode
            End If

            If (BTX < BTY) Then
                BTEdgeSeq = "E1S"
                BTEdgeCode = BMEdgeCode
            End If

        Else

            If (BTX > BTY) Then
                BTSEdgeSeq = "E1L"
                BTSEdgeCode = PVCEdgeCode
            End If

            If (BTX < BTY) Then
                BTSEdgeSeq = "E1S"
                BTSEdgeCode = PVCEdgeCode
            End If

        End If

        '################################
        '# BASE RANGE TOP EDGE SEQUENCE #
        '################################
        '
        If (BRTX > BRTY) Then
            BRTEdgeSeq = "E1L"
            BRTEdgeCode = PVCEdgeCode
        End If

        If (BRTX < BRTY) Then
            BRTEdgeSeq = "E1S"
            BRTEdgeCode = PVCEdgeCode
        End If

        '#############################
        '# BASE BOTTOM EDGE SEQUENCE #
        '#############################
        '
        If (BBX > BBY) Then
            BBEdgeSeq = "E1L"
            BBEdgeCode = BMEdgeCode
        End If

        If (BBX < BBY) Then
            BBEdgeSeq = "E1S"
            BBEdgeCode = BMEdgeCode
        End If

        '#######################################
        '# BASE FULL FIXED SHELF EDGE SEQUENCE #
        '#######################################
        '
        '###############
        '# GROUP RULES #
        '###############
        '
        If (BFFSX > BFFSY) Then
            BFFSHEdgeSeq = "E1L"
            BFFSHEdgeCode = PVCEdgeCode
        End If

        If (BFFSX < BFFSY) Then
            BFFSHEdgeSeq = "E1S"
            BFFSHEdgeCode = PVCEdgeCode
        End If

        '#######################################
        '# BASE ADJUSTABLE SHELF EDGE SEQUENCE #
        '#######################################
        '
        If (BASX > BASY) Then
            BASEdgeSeq = "E1L"
            BASEdgeCode = BMEdgeCode
        End If

        If (BASX < BASY) Then
            BASEdgeSeq = "E1S"
            BASEdgeCode = BMEdgeCode
        End If

        '############################
        '# BASE STRAP EDGE SEQUENCE #
        '############################
        '
        If (BSX > BSY) Then
            BSEdgeSeq = "E1L"
            BSEdgeCode = PVCEdgeCode
        End If

        If (BSX < BSY) Then
            BSEdgeSeq = "E1S"
            BSEdgeCode = PVCEdgeCode
        End If

        '##############################
        '# BASE DIVIDER EDGE SEQUENCE #
        '##############################
        '
        BDEdgeSeq = "E1S"
        BDEdgeCode = PVCEdgeCode

        '##################################################
        '# BASE PENINSULA OPEN EDGING RULES AND VARIABLES #
        '##################################################
        '
        '##################################################################
        '# BASE PENINSULA OPEN TOP, BOTTOM, AND FIXED SHELF EDGE SEQUENCE #
        '##################################################################
        '
        BPOPTBEdgeSeq = "E1L"
        BPOPTBEdgeCode = VeneerEdgeCode

        '#############################################
        '# BASE PENINSULA OPEN S-GABLE EDGE SEQUENCE #
        '#############################################
        '
        BPOPSGEdgeSeq = "E1L"
        BPOPSGEdgeCode = VeneerEdgeCode

        '#############################################
        '# BASE PENINSULA OPEN L-GABLE EDGE SEQUENCE #
        '#############################################
        '
        BPOPLGEdgeSeq = "E1L"
        BPOPLGEdgeCode = VeneerEdgeCode

        '#####################################################
        '# BASE MATCHING INTERIOR EDGING RULES AND VARIABLES #
        '#####################################################
        '
        Dim BMIGX As Double
        Dim BMIGY As Double
        BMIGX = BMIGableX
        BMIGY = BMIGableY

        Dim BMITSX As Double
        Dim BMITSY As Double
        BMITSX = BMITopStrapX
        BMITSY = BMITopStrapY

        Dim BMITX As Double
        Dim BMITY As Double
        BMITX = BMITopX
        BMITY = BMITopY

        Dim BMIBX As Double
        Dim BMIBY As Double
        BMIBX = BMIBtmX
        BMIBY = BMIBtmY

        Dim BMIASX As Double
        Dim BMIASY As Double
        BMIASX = BMIAdjShelfX
        BMIASY = BMIAdjShelfY

        Dim BMISX As Double
        Dim BMISY As Double
        BMISX = BMIStrapX
        BMISY = BMIStrapY

        Dim BMIDX As Double
        Dim BMIDY As Double
        BMIDX = BMIDividerX
        BMIDY = BMIDividerY

        '##############################################
        '# BASE MATCHING INTERIOR GABLE EDGE SEQUENCE #
        '##############################################
        '
        If (BMIGX > BMIGY) Then
            BMIGEdgeSeq = "E1L"
            BMIGEdgeCode = VeneerEdgeCode
        End If

        If (BMIGX < BMIGY) Then
            BMIGEdgeSeq = "E1S"
            BMIGEdgeCode = VeneerEdgeCode
        End If

        '##################################################
        '# BASE MATCHING INTERIOR TOP STRAP EDGE SEQUENCE #
        '##################################################
        '
        If (BMITSX > BMITSY) Then
            BMITSEdgeSeq = "E1L"
            BMITSEdgeCode = VeneerEdgeCode
        End If

        If (BMITSX < BMITSY) Then
            BMITSEdgeSeq = "E1S"
            BMITSEdgeCode = VeneerEdgeCode
        End If

        '############################################
        '# BASE MATCHING INTERIOR TOP EDGE SEQUENCE #
        '############################################
        '
        If (BMITX > BMITY) Then
            BMITEdgeSeq = "E1L"
            BMITEdgeCode = VeneerEdgeCode
        End If

        If (BMITX < BMITY) Then
            BMITEdgeSeq = "E1S"
            BMITEdgeCode = VeneerEdgeCode
        End If

        '##############################################
        '# BASE MATCHING INTERIOR STRAP EDGE SEQUENCE #
        '##############################################
        '
        If (BMISX > BMISY) Then
            BMISEdgeSeq = "E1L"
            BMISEdgeCode = VeneerEdgeCode
        End If

        If (BMISX < BMISY) Then
            BMISEdgeSeq = "E1S"
            BMISEdgeCode = VeneerEdgeCode
        End If

        '###############################################
        '# BASE MATCHING INTERIOR BOTTOM EDGE SEQUENCE #
        '###############################################
        '
        If (BMIBX > BMIBY) Then
            BMIBEdgeSeq = "E1L"
            BMIBEdgeCode = VeneerEdgeCode
        End If

        If (BMIBX < BMIBY) Then
            BMIBEdgeSeq = "E1S"
            BMIBEdgeCode = VeneerEdgeCode
        End If

        '################################################
        '# BASE MATCHING INTERIOR DIVIDER EDGE SEQUENCE #
        '################################################
        '
        If (BMIDX > BMIDY) Then
            BMIDEdgeSeq = "E1L"
            BMIDEdgeCode = VeneerEdgeCode
        End If

        If (BMIDX < BMIDY) Then
            BMIDEdgeSeq = "E1S"
            BMIDEdgeCode = VeneerEdgeCode
        End If

        '#########################################################
        '# BASE MATCHING INTERIOR ADJUSTABLE SHELF EDGE SEQUENCE #
        '#########################################################
        '
        If (BMIASX > BMIASY) Then
            BMIASEdgeSeq = "E1L"
            BMIASEdgeCode = VeneerEdgeCode
        End If

        If (BMIASX < BMIASY) Then
            BMIASEdgeSeq = "E1S"
            BMIASEdgeCode = VeneerEdgeCode
        End If

        '#################################################################
        '# BASE MICROWAVE OPEN SHELF 1 DRAWER EDGING RULES AND VARIABLES #
        '#################################################################
        '
        Dim BMO1DUGX As Double
        Dim BMO1DUGY As Double
        BMO1DUGX = BMO1DUGableX
        BMO1DUGY = BMO1DUGableY

        Dim BMO1DUTX As Double
        Dim BMO1DUTY As Double
        BMO1DUTX = BMO1DUTopX
        BMO1DUTY = BMO1DUTopY

        Dim BMO1DUBX As Double
        Dim BMO1DUBY As Double
        BMO1DUBX = BMO1DUBtmX
        BMO1DUBY = BMO1DUBtmY

        '##########################################################
        '# BASE MICROWAVE OPEN SHELF 1 DRAWER GABLE EDGE SEQUENCE #
        '##########################################################
        '
        If (BMO1DUGX > BMO1DUGY) Then
            BMO1DUGEdgeSeq = "E2S"
            BMO1DUGEdgeCode = VeneerEdgeCode
            BMO1DUGEdgeSeq2 = "E1L"
            BMO1DUGEdgeCode2 = VeneerEdgeCode
        End If

        If (BMO1DUGX < BMO1DUGY) Then
            BMO1DUGEdgeSeq = "E2L"
            BMO1DUGEdgeCode = VeneerEdgeCode
            BMO1DUGEdgeSeq2 = "E1S"
            BMO1DUGEdgeCode2 = VeneerEdgeCode
        End If

        '########################################################
        '# BASE MICROWAVE OPEN SHELF 1 DRAWER TOP EDGE SEQUENCE #
        '########################################################
        '

        '###############
        '# GROUP RULES #
        '###############
        '
        If (CKWOPLANNERSGroupBox = "GROUP1") Then

            If (BMO1DUTX > BMO1DUTY) Then
                BMO1DUTEdgeSeq = "E1L"
                BMO1DUTEdgeCode = VeneerEdgeCode
            End If

            If (BMO1DUTX < BMO1DUTY) Then
                BMO1DUTEdgeSeq = "E1S"
                BMO1DUTEdgeCode = VeneerEdgeCode
            End If

        Else

            If (BMO1DUBX > BMO1DUBY) Then
                BMO1DUBEdgeSeq = "E1L"
                BMO1DUBEdgeCode = VeneerEdgeCode
            End If

            If (BMO1DUBX < BMO1DUBY) Then
                BMO1DUBEdgeSeq = "E1S"
                BMO1DUBEdgeCode = VeneerEdgeCode
            End If

        End If

        '###########################################################
        '# BASE MICROWAVE OPEN SHELF 1 DRAWER BOTTOM EDGE SEQUENCE #
        '###########################################################
        '

        '###############
        '# GROUP RULES #
        '###############
        '
        If (CKWOPLANNERSGroupBox = "GROUP1") Then

            If (BMO1DUBX > BMO1DUBY) Then
                BMO1DUBEdgeSeq = "E1L"
                BMO1DUBEdgeCode = VeneerEdgeCode
            End If

            If (BMO1DUBX < BMO1DUBY) Then
                BMO1DUBEdgeSeq = "E1S"
                BMO1DUBEdgeCode = VeneerEdgeCode
            End If

        Else

            If (BMO1DUBX > BMO1DUBY) Then
                BMO1DUBEdgeSeq = "E1L"
                BMO1DUBEdgeCode = VeneerEdgeCode
            End If

            If (BMO1DUBX < BMO1DUBY) Then
                BMO1DUBEdgeSeq = "E1S"
                BMO1DUBEdgeCode = VeneerEdgeCode
            End If

        End If

        '#################################################################
        '# BASE MICROWAVE OPEN SHELF 1 DRAWER EDGING RULES AND VARIABLES #
        '#################################################################
        '
        Dim BMO1DBGX As Double
        Dim BMO1DBGY As Double
        BMO1DBGX = BMO1DBGableX
        BMO1DBGY = BMO1DBGableY

        Dim BMO1DBTSX As Double
        Dim BMO1DBTSY As Double
        BMO1DBTSX = BMO1DBTopStrapX
        BMO1DBTSY = BMO1DBTopStrapY

        Dim BMO1DBBX As Double
        Dim BMO1DBBY As Double
        BMO1DBBX = BMO1DBBtmX
        BMO1DBBY = BMO1DBBtmY

        '##########################################################
        '# BASE MICROWAVE OPEN SHELF 1 DRAWER GABLE EDGE SEQUENCE #
        '##########################################################
        '
        If (BMO1DBGX > BMO1DBGY) Then
            BMO1DBGEdgeSeq = "E1L"
            BMO1DBGEdgeCode = VeneerEdgeCode
        End If

        If (BMO1DBGX < BMO1DBGY) Then
            BMO1DBGEdgeSeq = "E1S"
            BMO1DBGEdgeCode = VeneerEdgeCode
        End If

        '##############################################################
        '# BASE MICROWAVE OPEN SHELF 1 DRAWER TOP STRAP EDGE SEQUENCE #
        '##############################################################
        '
        If (BMO1DBTSX > BMO1DBTSY) Then
            BMO1DBTSEdgeSeq = "E1L"
            BMO1DBTSEdgeCode = VeneerEdgeCode
        End If

        If (BMO1DBTSX < BMO1DBTSY) Then
            BMO1DBTSEdgeSeq = "E1S"
            BMO1DBTSEdgeCode = VeneerEdgeCode
        End If

        '###########################################################
        '# BASE MICROWAVE OPEN SHELF 1 DRAWER BOTTOM EDGE SEQUENCE #
        '###########################################################
        '
        If (BMO1DBBX > BMO1DBBY) Then
            BMO1DBBEdgeSeq = "E1L"
            BMO1DBBEdgeCode = VeneerEdgeCode
        End If

        If (BMO1DBBX < BMO1DBBY) Then
            BMO1DBBEdgeSeq = "E1S"
            BMO1DBBEdgeCode = VeneerEdgeCode
        End If

        '#########################################
        '# OVEN PANEL EDGING RULES AND VARIABLES #
        '#########################################
        '
        '###############
        '# GROUP RULES #
        '###############
        '
        If (CKWOPLANNERSGroupBox = "GROUP1") Then

            Dim OVPECase = CKWOPLANNERSSpeciesBox
            Select Case OVPECase
                Case "MAPLE", "OAK", "CHERRY", "PINE", "WALNUT"
                    OVPLEdgeSeq = "E4S"
                    OVPLEdgeCode = VeneerEdgeCode
                Case "MDF", "PVC"
                    OVPLEdgeSeq = ""
                    OVPLEdgeCode = ""
                Case Else
            End Select

        Else

            OVPLEdgeSeq = ""
            OVPLEdgeCode = ""

        End If

        '##################################################
        '# TALL UTILITY 1 UNIT EDGING RULES AND VARIABLES #
        '##################################################
        '
        Dim TU1UTX As Double = TU1UTopX
        Dim TU1UTY As Double = TU1UTopY

        Dim TU1UBX As Double = TU1UBotX
        Dim TU1UBY As Double = TU1UBotY

        Dim TU1UFFSHX As Double = TU1UFFSX
        Dim TU1UFFSHY As Double = TU1UFFSY

        Dim TU1USFSHX As Double = TU1USFSX
        Dim TU1USFSHY As Double = TU1USFSY

        Dim TU1USX As Double = TU1UStrapX
        Dim TU1USY As Double = TU1UStrapY

        Dim TU1UASX As Double = TU1UAdjShelfX
        Dim TU1UASY As Double = TU1UAdjShelfY

        '###########################################
        '# TALL UTILITY 1 UNIT GABLE EDGE SEQUENCE #
        '###########################################
        '
        TU1UGEdgeSeq = "E1L"
        TU1UGEdgeCode = PVCEdgeCode

        '###########################################
        '# TALL UTILITY 1 UNIT STRAP EDGE SEQUENCE #
        '###########################################
        '
        If (TU1USX > TU1USY) Then
            TU1USEdgeSeq = "E1L"
            TU1USEdgeCode = PVCEdgeCode
        End If

        If (TU1USX < TU1USY) Then
            TU1USEdgeSeq = "E1S"
            TU1USEdgeCode = PVCEdgeCode
        End If

        '#########################################
        '# TALL UTILITY 1 UNIT TOP EDGE SEQUENCE #
        '#########################################
        '
        If (CKWOPLANNERSGroupBox = "GROUP1") Then

            If (TU1UTX > TU1UTY) Then
                TU1UTEdgeSeq = "E1L"
                TU1UTEdgeCode = BMEdgeCode
            End If

            If (TU1UTX < TU1UTY) Then
                TU1UTEdgeSeq = "E1S"
                TU1UTEdgeCode = BMEdgeCode
            End If

        End If

        If (CKWOPLANNERSGroupBox = "GROUP2") Then

            If (TU1UTX > TU1UTY) Then
                TU1UTEdgeSeq = "E1L"
                TU1UTEdgeCode = PVCEdgeCode
            End If

            If (TU1UTX < TU1UTY) Then
                TU1UTEdgeSeq = "E1S"
                TU1UTEdgeCode = PVCEdgeCode
            End If

        End If
        '############################################
        '# TALL UTILITY 1 UNIT BOTTOM EDGE SEQUENCE #
        '############################################
        '
        If (TU1UBX > TU1UBY) Then
            TU1UBEdgeSeq = "E1L"
            TU1UBEdgeCode = BMEdgeCode
        End If

        If (TU1UBX < TU1UBY) Then
            TU1UBEdgeSeq = "E1S"
            TU1UBEdgeCode = BMEdgeCode
        End If

        '######################################################
        '# TALL UTILITY 1 UNIT FULL FIXED SHELF EDGE SEQUENCE #
        '######################################################
        '
        If (TU1UFFSHX > TU1UFFSHY) Then
            TU1UFFSEdgeSeq = "E1L"
            TU1UFFSEdgeCode = PVCEdgeCode
        End If

        If (TU1UFFSHX < TU1UFFSHY) Then
            TU1UFFSEdgeSeq = "E1S"
            TU1UFFSEdgeCode = PVCEdgeCode
        End If

        '###########################################
        '# TALL UTILITY 1 UNIT STRAP EDGE SEQUENCE #
        '###########################################
        '
        If (TU1USX > TU1USY) Then
            TU1USEdgeSeq = "E1L"
            TU1USEdgeCode = PVCEdgeCode
        End If

        If (TU1USX < TU1USY) Then
            TU1USEdgeSeq = "E1S"
            TU1USEdgeCode = PVCEdgeCode
        End If

        '#########################################################
        '# TALL UTILITY 1 UNIT SHALLOW FIXED SHELF EDGE SEQUENCE #
        '#########################################################
        '
        If (TU1USFSHX > TU1USFSHY) Then
            TU1USFSEdgeSeq = "E1L"
            TU1USFSEdgeCode = BMEdgeCode
        End If

        If (TU1USFSHX < TU1USFSHY) Then
            TU1USFSEdgeSeq = "E1S"
            TU1USFSEdgeCode = BMEdgeCode
        End If

        '######################################################
        '# TALL UTILITY 1 UNIT ADJUSTABLE SHELF EDGE SEQUENCE #
        '######################################################
        '
        If (TU1UASX > TU1UASY) Then
            TU1UASEdgeSeq = "E1L"
            TU1UASEdgeCode = BMEdgeCode
        End If

        If (TU1UASX < TU1UASY) Then
            TU1UASEdgeSeq = "E1S"
            TU1UASEdgeCode = BMEdgeCode
        End If


        '#####################################################################################################################################################################################
        '# TALL UTILITY 2 UNIT UPPER EDGING RULES AND VARIABLES #
        '########################################################
        '
        Dim TU2UUGX As Double = TU2UUGableX
        Dim TU2UUGY As Double = TU2UUGableY

        Dim TU2UUTX As Double = TU2UUTopX
        Dim TU2UUTY As Double = TU2UUTopY

        Dim TU2UUBX As Double = TU2UUBotX
        Dim TU2UUBY As Double = TU2UUBotY

        Dim TU2UUTBX As Double = TU2UUTopBtmX
        Dim TU2UUTBY As Double = TU2UUTopBtmY

        Dim TU2UUASX As Double = TU2UUAdjShelfX
        Dim TU2UUASY As Double = TU2UUAdjShelfY

        '#####################################################################################################################################################################################
        '# TALL UTILITY 2 UNIT UPPER MI EDGING RULES AND VARIABLES #
        '########################################################
        '
        Dim TU2UUMIGX As Double = TU2UUMIGableX
        Dim TU2UUMIGY As Double = TU2UUMIGableY

        Dim TU2UUMITX As Double = TU2UUMITopX
        Dim TU2UUMITY As Double = TU2UUMITopY

        Dim TU2UUMIBX As Double = TU2UUMIBotX
        Dim TU2UUMIBY As Double = TU2UUMIBotY

        Dim TU2UUMITBX As Double = TU2UUMITopBtmX
        Dim TU2UUMITBY As Double = TU2UUMITopBtmY

        Dim TU2UUMIASX As Double = TU2UUMIAdjShelfX
        Dim TU2UUMIASY As Double = TU2UUMIAdjShelfY

        '######################################
        '# TALL UTILITY 2 UNIT TALL VARIABLES #
        '######################################
        '
        Dim TU2UBTX As Double = TU2UBTopX
        Dim TU2UBTY As Double = TU2UBTopY

        Dim TU2UBBX As Double = TU2UBBotX
        Dim TU2UBBY As Double = TU2UBBotY

        Dim TU2UBTBX As Double = TU2UBTopBtmX
        Dim TU2UBTBY As Double = TU2UBTopBtmY

        Dim TU2UBSFSHX As Double = TU2UBSFSX
        Dim TU2UBSFSHY As Double = TU2UBSFSY

        Dim TU2UBASX As Double = TU2UBAdjShelfX
        Dim TU2UBASY As Double = TU2UBAdjShelfY

        '#################################################
        '# TALL UTILITY 2 UNIT UPPER GABLE EDGE SEQUENCE #
        '#################################################
        '
        If (TU2UUGX > TU2UUGY) Then
            TU2UUGEdgeSeq = "E2S"
            TU2UUGEdgeCode = BMEdgeCode
            TU2UUGEdgeSeq2 = "E1L"
            TU2UUGEdgeCode2 = PVCEdgeCode
        Else
            TU2UUGEdgeSeq = "E2L"
            TU2UUGEdgeCode = BMEdgeCode
            TU2UUGEdgeSeq2 = "E1S"
            TU2UUGEdgeCode2 = PVCEdgeCode
        End If

        '###############
        '# GROUP RULES #
        '###############
        '
        If (CKWOPLANNERSGroupBox = "GROUP1") Then
            '###############################################
            '# TALL UTILITY 2 UNIT UPPER TOP EDGE SEQUENCE #
            '###############################################
            '
            If (TU2UUTX > TU2UUTY) Then
                TU2UUTEdgeSeq = "E1L"
                TU2UUTEdgeCode = BMEdgeCode
            End If

            If (TU2UUTX < TU2UUTY) Then
                TU2UUTEdgeSeq = "E1S"
                TU2UUTEdgeCode = BMEdgeCode
            End If

            '##################################################
            '# TALL UTILITY 2 UNIT UPPER BOTTOM EDGE SEQUENCE #
            '##################################################
            '
            If (TU2UUBX > TU2UUBY) Then
                TU2UUBEdgeSeq = "E1L"
                TU2UUBEdgeCode = PVCEdgeCode
            End If

            If (TU2UUBX < TU2UUBY) Then
                TU2UUBEdgeSeq = "E1S"
                TU2UUBEdgeCode = PVCEdgeCode
            End If

            '############################################################
            '# TALL UTILITY 2 UNIT UPPER ADJUSTABLE SHELF EDGE SEQUENCE #
            '############################################################
            '
            If (TU2UUASX > TU2UUASY) Then
                TU2UUASEdgeSeq = "E1L"
                TU2UUASEdgeCode = BMEdgeCode
            End If

            If (TU2UUASX < TU2UUASY) Then
                TU2UUASEdgeSeq = "E1S"
                TU2UUASEdgeCode = BMEdgeCode
            End If

        Else

            '#######################################################
            '# TALL UTILITY 2 UNIT TALL TOP + BOTTOM EDGE SEQUENCE #
            '#######################################################
            '
            If (TU2UUTBX > TU2UUTBY) Then
                TU2UUTBEdgeSeq = "E1L"
                TU2UUTBEdgeCode = PVCEdgeCode
            End If
            If (TU2UUTBX < TU2UUTBY) Then
                TU2UUTBEdgeSeq = "E1S"
                TU2UUTBEdgeCode = PVCEdgeCode
            End If

            '############################################################
            '# TALL UTILITY 2 UNIT UPPER ADJUSTABLE SHELF EDGE SEQUENCE #
            '############################################################
            '
            If (TU2UUASX > TU2UUASY) Then
                TU2UUASEdgeSeq = "E1L"
                TU2UUASEdgeCode = PVCEdgeCode
            End If

            If (TU2UUASX < TU2UUASY) Then
                TU2UUASEdgeSeq = "E1S"
                TU2UUASEdgeCode = PVCEdgeCode
            End If

        End If

        '#################################################
        '# TALL UTILITY 2 UNIT UPPER MI GABLE EDGE SEQUENCE #
        '#################################################
        '
        If (TU2UUMIGX > TU2UUMIGY) Then
            TU2UUMIGEdgeSeq = "E2S"
            TU2UUMIGEdgeCode = VeneerEdgeCode
            TU2UUMIGEdgeSeq2 = "E1L"
            TU2UUMIGEdgeCode2 = VeneerEdgeCode
        End If
        If (TU2UUMIGX < TU2UUMIGY) Then
            TU2UUMIGEdgeSeq = "E2L"
            TU2UUMIGEdgeCode = VeneerEdgeCode
            TU2UUMIGEdgeSeq2 = "E1S"
            TU2UUMIGEdgeCode2 = VeneerEdgeCode
        End If

        '##########
        '# GROUP1 #
        '##########
        '
        If (CKWOPLANNERSGroupBox = "GROUP1") Then
            '###############################################
            '# TALL UTILITY 2 UNIT UPPER TOP EDGE SEQUENCE #
            '###############################################
            '
            If (TU2UUMITX > TU2UUMITY) Then
                TU2UUMITEdgeSeq = "E1L"
                TU2UUMITEdgeCode = VeneerEdgeCode
            End If
            If (TU2UUMITX < TU2UUMITY) Then
                TU2UUMITEdgeSeq = "E1S"
                TU2UUMITEdgeCode = VeneerEdgeCode
            End If

            '##################################################
            '# TALL UTILITY 2 UNIT UPPER BOTTOM EDGE SEQUENCE #
            '##################################################
            '
            If (TU2UUMIBX > TU2UUMIBY) Then
                TU2UUMIBEdgeSeq = "E1L"
                TU2UUMIBEdgeCode = VeneerEdgeCode
            End If
            If (TU2UUMIBX < TU2UUMIBY) Then
                TU2UUMIBEdgeSeq = "E1S"
                TU2UUMIBEdgeCode = VeneerEdgeCode
            End If

            '############################################################
            '# TALL UTILITY 2 UNIT UPPER ADJUSTABLE SHELF EDGE SEQUENCE #
            '############################################################
            '
            If (TU2UUMIASX > TU2UUMIASY) Then
                TU2UUMIASEdgeSeq = "E1L"
                TU2UUMIASEdgeCode = VeneerEdgeCode
            End If
            If (TU2UUMIASX < TU2UUMIASY) Then
                TU2UUMIASEdgeSeq = "E1S"
                TU2UUMIASEdgeCode = VeneerEdgeCode
            End If
        End If

        '##########
        '# GROUP2 #
        '##########
        '
        If (CKWOPLANNERSGroupBox = "GROUP2") Then
            '#######################################################
            '# TALL UTILITY 2 UNIT TALL TOP + BOTTOM EDGE SEQUENCE #
            '#######################################################
            '
            If (TU2UUMITBX > TU2UUMITBY) Then
                TU2UUMITBEdgeSeq = "E1L"
                TU2UUMITBEdgeCode = VeneerEdgeCode
            End If
            If (TU2UUMITBX < TU2UUMITBY) Then
                TU2UUMITBEdgeSeq = "E1S"
                TU2UUMITBEdgeCode = VeneerEdgeCode
            End If

            '############################################################
            '# TALL UTILITY 2 UNIT UPPER ADJUSTABLE SHELF EDGE SEQUENCE #
            '############################################################
            '
            If (TU2UUMIASX > TU2UUMIASY) Then
                TU2UUMIASEdgeSeq = "E1L"
                TU2UUMIASEdgeCode = VeneerEdgeCode
            End If
            If (TU2UUMIASX < TU2UUMIASY) Then
                TU2UUMIASEdgeSeq = "E1S"
                TU2UUMIASEdgeCode = VeneerEdgeCode
            End If
        End If

        '################################################
        '# TALL UTILITY 2 UNIT TALL GABLE EDGE SEQUENCE #
        '################################################
        '
        TU2UBGEdgeSeq = "E1L"
        TU2UBGEdgeCode = PVCEdgeCode

        '###############
        '# GROUP RULES #
        '###############
        '
        If (CKWOPLANNERSGroupBox = "GROUP1") Then
            '##############################################
            '# TALL UTILITY 2 UNIT TALL TOP EDGE SEQUENCE #
            '##############################################
            '
            If (TU2UBTX > TU2UBTY) Then
                TU2UBTEdgeSeq = "E1L"
                TU2UBTEdgeCode = PVCEdgeCode
            End If

            If (TU2UBTX < TU2UBTY) Then
                TU2UBTEdgeSeq = "E1S"
                TU2UBTEdgeCode = PVCEdgeCode
            End If

            '#################################################
            '# TALL UTILITY 2 UNIT TALL BOTTOM EDGE SEQUENCE #
            '#################################################
            '
            If (TU2UBBX > TU2UBBY) Then
                TU2UBBEdgeSeq = "E1L"
                TU2UBBEdgeCode = BMEdgeCode
            End If

            If (TU2UBBX < TU2UBBY) Then
                TU2UBBEdgeSeq = "E1S"
                TU2UBBEdgeCode = BMEdgeCode
            End If

            '##############################################################
            '# TALL UTILITY 2 UNIT TALL SHALLOW FIXED SHELF EDGE SEQUENCE #
            '##############################################################
            '
            If (TU2UBSFSHX > TU2UBSFSHY) Then
                TU2UBSFSEdgeSeq = "E1L"
                TU2UBSFSEdgeCode = BMEdgeCode
            End If

            If (TU2UBSFSHX < TU2UBSFSHY) Then
                TU2UBSFSEdgeSeq = "E1S"
                TU2UBSFSEdgeCode = BMEdgeCode
            End If

            '###########################################################
            '# TALL UTILITY 2 UNIT TALL ADJUSTABLE SHELF EDGE SEQUENCE #
            '###########################################################
            '
            If (TU2UBASX > TU2UBASY) Then
                TU2UBASEdgeSeq = "E1L"
                TU2UBASEdgeCode = BMEdgeCode
            End If

            If (TU2UBASX < TU2UBASY) Then
                TU2UBASEdgeSeq = "E1S"
                TU2UBASEdgeCode = BMEdgeCode
            End If

        Else

            '##############################################
            '# TALL UTILITY 2 UNIT TALL TOP EDGE SEQUENCE #
            '##############################################
            '
            If (TU2UBTX > TU2UBTY) Then
                TU2UBTEdgeSeq = "E1L"
                TU2UBTEdgeCode = PVCEdgeCode
            End If

            If (TU2UBTX < TU2UBTY) Then
                TU2UBTEdgeSeq = "E1S"
                TU2UBTEdgeCode = PVCEdgeCode
            End If

            '#################################################
            '# TALL UTILITY 2 UNIT TALL BOTTOM EDGE SEQUENCE #
            '#################################################
            '
            If (TU2UBTX > TU2UBTY) Then
                TU2UBBEdgeSeq = "E1L"
                TU2UBBEdgeCode = BMEdgeCode
            End If

            If (TU2UBTX < TU2UBTY) Then
                TU2UBBEdgeSeq = "E1S"
                TU2UBBEdgeCode = BMEdgeCode
            End If

            '##############################################################
            '# TALL UTILITY 2 UNIT TALL SHALLOW FIXED SHELF EDGE SEQUENCE #
            '##############################################################
            '
            If (TU2UBSFSHX > TU2UBSFSHY) Then
                TU2UBSFSEdgeSeq = "E1L"
                TU2UBSFSEdgeCode = PVCEdgeCode
            End If

            If (TU2UBSFSHX < TU2UBSFSHY) Then
                TU2UBSFSEdgeSeq = "E1S"
                TU2UBSFSEdgeCode = PVCEdgeCode
            End If

            '###########################################################
            '# TALL UTILITY 2 UNIT TALL ADJUSTABLE SHELF EDGE SEQUENCE #
            '###########################################################
            '
            If (TU2UBASX > TU2UBASY) Then
                TU2UBASEdgeSeq = "E1L"
                TU2UBASEdgeCode = BMEdgeCode
            End If

            If (TU2UBASX < TU2UBASY) Then
                TU2UBASEdgeSeq = "E1S"
                TU2UBASEdgeCode = BMEdgeCode
            End If

        End If

        '#####################################################################################################################################################################################
        '# UPPER WINE RACK EDGING RULES AND VARIABLES #
        '##############################################
        '
        Dim UWRGX As Double
        Dim UWRGY As Double
        UWRGX = UWRGableX
        UWRGY = UWRGableY

        Dim UWRTBX As Double
        Dim UWRTBY As Double
        UWRTBX = UTopBtmX
        UWRTBY = UTopBtmY

        Dim UWRFFSX As Double
        Dim UWRFFSY As Double
        UWRFFSX = UWRFFSHX
        UWRFFSY = UWRFFSHY

        Dim UWRFSX As Double
        Dim UWRFSY As Double
        UWRFSX = UWRFShelfX
        UWRFSY = UWRFShelfY

        Dim UWRDX As Double
        Dim UWRDY As Double
        UWRDX = UWRDivX
        UWRDY = UWRDivY

        '#######################################
        '# UPPER WINE RACK GABLE EDGE SEQUENCE #
        '#######################################
        '
        If (UWRGX > UWRGY) Then
            UWRGEdgeSeq = "E2S"
            UWRGEdgeCode = VeneerEdgeCode
            UWRGEdgeSeq2 = "E1L"
            UWRGEdgeCode2 = VeneerEdgeCode
        End If

        If (UWRGX < UWRGY) Then
            UWRGEdgeSeq = "E2L"
            UWRGEdgeCode = VeneerEdgeCode
            UWRGEdgeSeq2 = "E1S"
            UWRGEdgeCode2 = VeneerEdgeCode
        End If

        '##############################################
        '# UPPER WINE RACK TOP + BOTTOM EDGE SEQUENCE #
        '##############################################
        '
        '###############
        '# GROUP RULES #
        '###############
        '
        If (CKWOPLANNERSGroupBox = "GROUP1") Then

            If (UWRTBX > UWRTBY) Then
                UWRTBEdgeSeq = "E1L"
                UWRTBEdgeCode = VeneerEdgeCode
            End If

            If (UWRTBX < UWRTBY) Then
                UWRTBEdgeSeq = "E1S"
                UWRTBEdgeCode = VeneerEdgeCode
            End If

        Else

            If (UWRTBX > UWRTBY) Then
                UWRTBEdgeSeq = "E1L"
                UWRTBEdgeCode = VeneerEdgeCode
            End If

            If (UWRTBX < UWRTBY) Then
                UWRTBEdgeSeq = "E1S"
                UWRTBEdgeCode = VeneerEdgeCode
            End If

        End If

        '##################################################
        '# UPPER WINE RACK FULL FIXED SHELF EDGE SEQUENCE #
        '##################################################
        '
        If (UWRFFSX > UWRFFSY) Then
            UWRFFSEdgeSeq = "E1L"
            UWRFFSEdgeCode = VeneerEdgeCode
        End If

        If (UWRFFSX < UWRFFSY) Then
            UWRFFSEdgeSeq = "E1S"
            UWRFFSEdgeCode = VeneerEdgeCode
        End If

        '##################################################
        '# UPPER WINE RACK FIXED SHELF EDGE SEQUENCE #
        '##################################################
        '
        If (UWRFSX > UWRFSY) Then
            UWRFSEdgeSeq = "E1L"
            UWRFSEdgeCode = VeneerEdgeCode
        End If

        If (UWRFSX < UWRFSY) Then
            UWRFSEdgeSeq = "E1S"
            UWRFSEdgeCode = VeneerEdgeCode
        End If

        '##################################################
        '# UPPER WINE RACK DIVIDER EDGE SEQUENCE #
        '##################################################
        '
        If (UWRDX > UWRDY) Then
            UWRDEdgeSeq = "E2S"
            UWRDEdgeCode = VeneerEdgeCode
            UWRDEdgeSeq2 = "E1L"
            UWRDEdgeCode2 = VeneerEdgeCode
        End If

        If (UWRDX < UWRDY) Then
            UWRDEdgeSeq = "E2L"
            UWRDEdgeCode = VeneerEdgeCode
            UWRDEdgeSeq2 = "E1S"
            UWRDEdgeCode2 = VeneerEdgeCode
        End If



        '#####################################################################################################################################################################################
        '# BASE WINE RACK EDGING RULES AND VARIABLES #
        '#############################################
        '
        Dim BWRGX As Double
        Dim BWRGY As Double
        BWRGX = BWRGableX
        BWRGY = BWRGableY

        Dim BWRTBX As Double
        Dim BWRTBY As Double
        BWRTBX = BWRTopBtmX
        BWRTBY = BWRTopBtmY

        Dim BWRFFSX As Double
        Dim BWRFFSY As Double
        BWRFFSX = BWRFFSHX
        BWRFFSY = BWRFFSHY

        '######################################
        '# BASE WINE RACK GABLE EDGE SEQUENCE #
        '######################################
        '
        If (BWRGX > BWRGY) Then
            BWRGEdgeSeq = "E1L"
            BWRGEdgeCode = VeneerEdgeCode
        End If

        If (BWRGX < BWRGY) Then
            BWRGEdgeSeq = "E1S"
            BWRGEdgeCode = VeneerEdgeCode
        End If

        '#############################################
        '# BASE WINE RACK TOP + BOTTOM EDGE SEQUENCE #
        '#############################################
        '
        '###############
        '# GROUP RULES #
        '###############
        '
        If (CKWOPLANNERSGroupBox = "GROUP1") Then

            If (BWRTBX > BWRTBY) Then
                BWRTBEdgeSeq = "E1L"
                BWRTBEdgeCode = VeneerEdgeCode
            End If

            If (BWRTBX < BWRTBY) Then
                BWRTBEdgeSeq = "E1S"
                BWRTBEdgeCode = VeneerEdgeCode
            End If

        End If

        '#####################
        '# GROUP2 AND GROUP3 #
        '#####################
        '
        If (CKWOPLANNERSGroupBox = "GROUP2" Or CKWOPLANNERSGroupBox = "GROUP3") Then

            If (BWRTBX > BWRTBY) Then
                BWRTBEdgeSeq = "E1L"
                BWRTBEdgeCode = VeneerEdgeCode
            End If

            If (BWRTBX < BWRTBY) Then
                BWRTBEdgeSeq = "E1S"
                BWRTBEdgeCode = VeneerEdgeCode
            End If

        End If

        '#################################################
        '# BASE WINE RACK FULL FIXED SHELF EDGE SEQUENCE #
        '#################################################
        '
        If (BWRFFSX > BWRFFSY) Then
            BWRFFSEdgeSeq = "E1L"
            BWRFFSEdgeCode = VeneerEdgeCode
        End If

        If (BWRFFSX < BWRFFSY) Then
            BWRFFSEdgeSeq = "E1S"
            BWRFFSEdgeCode = VeneerEdgeCode
        End If

        '#####################################################################################################################################################################################
        '# HUTCH EDGING RULES AND VARIABLES #
        '####################################
        '
        Dim HGX As Double
        Dim HGY As Double
        HGX = HGableX
        HGY = HGableY

        Dim HTFFSHX As Double
        Dim HTFFSHY As Double
        HTFFSHX = HTopFFShelfX
        HTFFSHY = HTopFFShelfY

        Dim HBFFSHX As Double
        Dim HBFFSHY As Double
        HBFFSHX = HBtmFFShelfX
        HBFFSHY = HBtmFFShelfY

        Dim HTX As Double
        Dim HTY As Double
        HTX = HTopX
        HTY = HTopY

        Dim HBX As Double
        Dim HBY As Double
        HBX = HBtmX
        HBY = HBtmY

        Dim HDX As Double
        Dim HDY As Double
        HDX = HDividerX
        HDY = HDividerY

        Dim HSX As Double
        Dim HSY As Double
        HSX = HStrapX
        HSY = HStrapY

        Dim HASX As Double
        Dim HASY As Double
        HASX = HAdjShelfX
        HASY = HAdjShelfY

        '#############################
        '# HUTCH GABLE EDGE SEQUENCE #
        '#############################
        '
        If (HGX > HGY) Then
            HGEdgeSeq = "E2S"
            HGEdgeCode = BMEdgeCode
            HGEdgeSeq2 = "E1L"
            HGEdgeCode2 = PVCEdgeCode
        End If

        If (HGX < HGY) Then
            HGEdgeSeq = "E2L"
            HGEdgeCode = BMEdgeCode
            HGEdgeSeq2 = "E1S"
            HGEdgeCode2 = PVCEdgeCode
        End If

        '#########################################
        '# HUTCH TOP + FULL FIXED SHELF SEQUENCE #
        '#########################################
        '
        If (HTFFSHX > HTFFSHY) Then
            HTFFSEdgeSeq = "E1L"
            HTFFSEdgeCode = PVCEdgeCode
        End If

        If (HTFFSHX < HTFFSHY) Then
            HTFFSEdgeSeq = "E1S"
            HTFFSEdgeCode = PVCEdgeCode
        End If

        '#################################################
        '# HUTCH BOTTOM + FULL FIXED SHELF EDGE SEQUENCE #
        '#################################################
        '
        If (HBFFSHX > HBFFSHY) Then
            HBFFSEdgeSeq = "E1L"
            HBFFSEdgeCode = PVCEdgeCode
        End If

        If (HBFFSHX < HBFFSHY) Then
            HBFFSEdgeSeq = "E1S"
            HBFFSEdgeCode = PVCEdgeCode
        End If

        '###########################
        '# HUTCH TOP EDGE SEQUENCE #
        '###########################
        '
        '###############
        '# GROUP RULES #
        '###############
        '
        If (CKWOPLANNERSGroupBox = "GROUP1") Then

            If (HTX > HTY) Then
                HTEdgeSeq = "E1L"
                HTEdgeCode = BMEdgeCode
            End If

            If (HTX < HTY) Then
                HTEdgeSeq = "E1S"
                HTEdgeCode = BMEdgeCode
            End If

        Else

            If (HTX > HTY) Then
                HTEdgeSeq = "E1L"
                HTEdgeCode = PVCEdgeCode
            End If

            If (HTX < HTY) Then
                HTEdgeSeq = "E1S"
                HTEdgeCode = PVCEdgeCode
            End If

        End If

        '##############################
        '# HUTCH BOTTOM EDGE SEQUENCE #
        '##############################
        '
        '###############
        '# GROUP RULES #
        '###############
        '
        If (CKWOPLANNERSGroupBox = "GROUP1") Then

            If (HBX > HBY) Then
                HBEdgeSeq = "E1L"
                HBEdgeCode = BMEdgeCode
            End If

            If (HBX < HBY) Then
                HBEdgeSeq = "E1S"
                HBEdgeCode = BMEdgeCode
            End If

        Else

            If (HBX > HBY) Then
                HBEdgeSeq = "E1L"
                HBEdgeCode = PVCEdgeCode
            End If

            If (HBX < HBY) Then
                HBEdgeSeq = "E1S"
                HBEdgeCode = PVCEdgeCode
            End If

        End If

        '##############################
        '# HUTCH STRAP EDGE SEQUENCE #
        '##############################
        '
        If (HSX > HSY) Then
            HSEdgeSeq = "E1L"
            HSEdgeCode = PVCEdgeCode
        End If

        If (HSX < HSY) Then
            HSEdgeSeq = "E1S"
            HSEdgeCode = PVCEdgeCode
        End If

        '###############################
        '# HUTCH DIVIDER EDGE SEQUENCE #
        '###############################
        '
        If (HDX > HDY) Then
            HDEdgeSeq = "E1L"
            HDEdgeCode = PVCEdgeCode
        End If

        If (HDX < HDY) Then
            HDEdgeSeq = "E1S"
            HDEdgeCode = PVCEdgeCode
        End If


        '########################################
        '# HUTCH ADJUSTABLE SHELF EDGE SEQUENCE #
        '########################################
        '
        If (HASX > HASY) Then
            HASEdgeSeq = "E1L"
            HASEdgeCode = BMEdgeCode
        End If

        If (HASX < HASY) Then
            HASEdgeSeq = "E1S"
            HASEdgeCode = BMEdgeCode
        End If

        '#####################################################################################################################################################################################
        '# HUTCH MATCHING INTERIOR EDGING RULES AND VARIABLES #
        '######################################################
        '
        Dim HMIGX As Double
        Dim HMIGY As Double
        HMIGX = HMIGableX
        HMIGY = HMIGableY

        Dim HMITFFSHX As Double
        Dim HMITFFSHY As Double
        HMITFFSHX = HMITopFFShelfX
        HMITFFSHY = HMITopFFShelfY

        Dim HMIBFFSHX As Double
        Dim HMIBFFSHY As Double
        HMIBFFSHX = HMIBtmFFShelfX
        HMIBFFSHY = HMIBtmFFShelfY

        Dim HMITX As Double
        Dim HMITY As Double
        HMITX = HMITopX
        HMITY = HMITopY

        Dim HMIBX As Double
        Dim HMIBY As Double
        HMIBX = HMIBtmX
        HMIBY = HMIBtmY

        Dim HMIDX As Double
        Dim HMIDY As Double
        HMIDX = HMIDividerX
        HMIDY = HMIDividerY

        Dim HMISX As Double
        Dim HMISY As Double
        HMISX = HMIStrapX
        HMISY = HMIStrapY

        Dim HMIASX As Double
        Dim HMIASY As Double
        HMIASX = HMIAdjShelfX
        HMIASY = HMIAdjShelfY

        '###############################################
        '# HUTCH MATCHING INTERIOR GABLE EDGE SEQUENCE #
        '###############################################
        '
        If (HMIGX > HMIGY) Then
            HMIGEdgeSeq = "E2S"
            HMIGEdgeCode = VeneerEdgeCode
            HMIGEdgeSeq2 = "E1L"
            HMIGEdgeCode2 = VeneerEdgeCode
        End If

        If (HMIGX < HMIGY) Then
            HMIGEdgeSeq = "E2L"
            HMIGEdgeCode = VeneerEdgeCode
            HMIGEdgeSeq2 = "E1S"
            HMIGEdgeCode2 = VeneerEdgeCode
        End If

        '###########################################################
        '# HUTCH MATCHING INTERIOR TOP + FULL FIXED SHELF SEQUENCE #
        '###########################################################
        '
        If (HMITFFSHX > HMITFFSHY) Then
            HMITFFSEdgeSeq = "E1L"
            HMITFFSEdgeCode = VeneerEdgeCode
        End If

        If (HMITFFSHX < HMITFFSHY) Then
            HMITFFSEdgeSeq = "E1S"
            HMITFFSEdgeCode = VeneerEdgeCode
        End If

        '###################################################################
        '# HUTCH MATCHING INTERIOR BOTTOM + FULL FIXED SHELF EDGE SEQUENCE #
        '###################################################################
        '
        If (HMIBFFSHX > HMIBFFSHY) Then
            HMIBFFSEdgeSeq = "E1L"
            HMIBFFSEdgeCode = VeneerEdgeCode
        End If

        If (HMIBFFSHX < HMIBFFSHY) Then
            HMIBFFSEdgeSeq = "E1S"
            HMIBFFSEdgeCode = VeneerEdgeCode
        End If

        '#############################################
        '# HUTCH MATCHING INTERIOR TOP EDGE SEQUENCE #
        '#############################################
        '
        If (HMITX > HMITY) Then
            HMITEdgeSeq = "E1L"
            HMITEdgeCode = VeneerEdgeCode
        End If

        If (HMITX < HMITY) Then
            HMITEdgeSeq = "E1S"
            HMITEdgeCode = VeneerEdgeCode
        End If

        '################################################
        '# HUTCH MATCHING INTERIOR BOTTOM EDGE SEQUENCE #
        '################################################
        '
        If (HMIBX > HMIBY) Then
            HMIBEdgeSeq = "E1L"
            HMIBEdgeCode = VeneerEdgeCode
        End If

        If (HMIBX < HMIBY) Then
            HMIBEdgeSeq = "E1S"
            HMIBEdgeCode = VeneerEdgeCode
        End If

        '#################################################
        '# HUTCH MATCHING INTERIOR DIVIDER EDGE SEQUENCE #
        '#################################################
        '
        If (HMIDX > HMIDY) Then
            HMIDEdgeSeq = "E1L"
            HMIDEdgeCode = VeneerEdgeCode
        End If

        If (HMIDX < HMIDY) Then
            HMIDEdgeSeq = "E1S"
            HMIDEdgeCode = VeneerEdgeCode
        End If

        '###############################################
        '# HUTCH MATCHING INTERIOR STRAP EDGE SEQUENCE #
        '###############################################
        '
        If (HMISX > HMISY) Then
            HMISEdgeSeq = "E1L"
            HMISEdgeCode = VeneerEdgeCode
        End If

        If (HMISX < HMISY) Then
            HMISEdgeSeq = "E1S"
            HMISEdgeCode = VeneerEdgeCode
        End If

        '##########################################################
        '# HUTCH MATCHING INTERIOR ADJUSTABLE SHELF EDGE SEQUENCE #
        '##########################################################
        '
        If (HMIASX > HMIASY) Then
            HMIASEdgeSeq = "E1L"
            HMIASEdgeCode = VeneerEdgeCode
        End If

        If (HMIASX < HMIASY) Then
            HMIASEdgeSeq = "E1S"
            HMIASEdgeCode = VeneerEdgeCode
        End If

        '#####################################################################################################################################################################################
        '# VANITY EDGING RULES AND VARIABLES #
        '#####################################
        '
        Dim VGX As Double
        Dim VGY As Double
        VGX = VGableX
        VGY = VGableY

        Dim VTSX As Double
        Dim VTSY As Double
        VTSX = VTopStrapX
        VTSY = VTopStrapY

        Dim VSX As Double
        Dim VSY As Double
        VSX = VStrapX
        VSY = VStrapY

        Dim VTX As Double
        Dim VTY As Double
        VTX = VTopX
        VTY = VTopY

        Dim VBX As Double
        Dim VBY As Double
        VBX = VBtmX
        VBY = VBtmY

        Dim VFFSHX As Double
        Dim VFFSHY As Double
        VFFSHX = VFFShelfX
        VFFSHY = VFFShelfY

        '##############################
        '# VANITY GABLE EDGE SEQUENCE #
        '##############################
        '
        If (VGX > VGY) Then
            VGEdgeSeq = "E1L"
            VGEdgeCode = PVCEdgeCode
        End If

        If (VGX < VGY) Then
            VGEdgeSeq = "E1S"
            VGEdgeCode = PVCEdgeCode
        End If

        '##################################
        '# VANITY TOP STRAP EDGE SEQUENCE #
        '##################################
        '
        '###############
        '# GROUP RULES #
        '###############
        '
        If (CKWOPLANNERSGroupBox = "GROUP1") Then

            If (VTSX > VTSY) Then
                VTSEdgeSeq = "E1L"
                VTSEdgeCode = BMEdgeCode
            End If

            If (VTSX < VTSY) Then
                VTSEdgeSeq = "E1S"
                VTSEdgeCode = BMEdgeCode
            End If

        Else

            If (VTSX > VTSY) Then
                VTSEdgeSeq = "E1L"
                VTSEdgeCode = PVCEdgeCode
            End If

            If (VTSX < VTSY) Then
                VTSEdgeSeq = "E1S"
                VTSEdgeCode = PVCEdgeCode
            End If

        End If

        '############################
        '# VANITY TOP EDGE SEQUENCE #
        '############################
        '
        '###############
        '# GROUP RULES #
        '###############
        '
        If (CKWOPLANNERSGroupBox = "GROUP1") Then

            If (VTX > VTY) Then
                VTEdgeSeq = "E1L"
                VTEdgeCode = BMEdgeCode
            End If

            If (VTX < VTY) Then
                VTEdgeSeq = "E1S"
                VTEdgeCode = BMEdgeCode
            End If

        Else

            If (VTX > VTY) Then
                VTEdgeSeq = "E1L"
                VTEdgeCode = PVCEdgeCode
            End If

            If (VTX < VTY) Then
                VTEdgeSeq = "E1S"
                VTEdgeCode = PVCEdgeCode
            End If

        End If

        '###############################
        '# VANITY BOTTOM EDGE SEQUENCE #
        '###############################
        '
        '###############
        '# GROUP RULES #
        '###############
        '
        If (VBX > VBY) Then
            VBEdgeSeq = "E1L"
            VBEdgeCode = BMEdgeCode
        End If

        If (VBX < VBY) Then
            VBEdgeSeq = "E1S"
            VBEdgeCode = BMEdgeCode
        End If

        '#########################################
        '# VANITY FULL FIXED SHELF EDGE SEQUENCE #
        '#########################################
        '
        If (VFFSHX > VFFSHY) Then
            VFFSHEdgeSeq = "E1L"
            VFFSHEdgeCode = PVCEdgeCode
        End If

        If (VFFSHX < VFFSHY) Then
            VFFSHEdgeSeq = "E1S"
            VFFSHEdgeCode = PVCEdgeCode
        End If

        '################################
        '# VANITY DIVIDER EDGE SEQUENCE #
        '################################
        '
        VDEdgeSeq = "E1S"
        VDEdgeCode = PVCEdgeCode

        '##############################
        '# VANITY STRAP EDGE SEQUENCE #
        '##############################
        '
        If (VSX > VSY) Then
            VSEdgeSeq = "E1L"
            VSEdgeCode = PVCEdgeCode
        End If

        If (VSX < VSY) Then
            VSEdgeSeq = "E1S"
            VSEdgeCode = PVCEdgeCode
        End If

        '#####################################################################################################################################################################################
        '# VANITY MATCHING INTERIOR EDGING RULES AND VARIABLES #
        '#######################################################
        '
        Dim VMIGX As Double
        Dim VMIGY As Double
        VMIGX = VMIGableX
        VMIGY = VMIGableY

        Dim VMITSX As Double
        Dim VMITSY As Double
        VMITSX = VMITopStrapX
        VMITSY = VMITopStrapY

        Dim VMISX As Double
        Dim VMISY As Double
        VMISX = VMIStrapX
        VMISY = VMIStrapY

        Dim VMITX As Double
        Dim VMITY As Double
        VMITX = VMITopX
        VMITY = VMITopY

        Dim VMIBX As Double
        Dim VMIBY As Double
        VMIBX = VMIBtmX
        VMIBY = VMIBtmY

        Dim VMIFFSHX As Double
        Dim VMIFFSHY As Double
        VMIFFSHX = VMIFFShelfX
        VMIFFSHY = VMIFFShelfY

        '################################################
        '# VANITY MATCHING INTERIOR GABLE EDGE SEQUENCE #
        '################################################
        '
        If (VMIGX > VMIGY) Then
            VMIGEdgeSeq = "E1L"
            VMIGEdgeCode = VeneerEdgeCode
        End If

        If (VMIGX < VMIGY) Then
            VMIGEdgeSeq = "E1S"
            VMIGEdgeCode = VeneerEdgeCode
        End If

        '####################################################
        '# VANITY MATCHING INTERIOR TOP STRAP EDGE SEQUENCE #
        '####################################################
        '
        If (VMITSX > VMITSY) Then
            VMITSEdgeSeq = "E1L"
            VMITSEdgeCode = VeneerEdgeCode
        End If

        If (VMITSX < VMITSY) Then
            VMITSEdgeSeq = "E1S"
            VMITSEdgeCode = VeneerEdgeCode
        End If

        '##############################################
        '# VANITY MATCHING INTERIOR TOP EDGE SEQUENCE #
        '##############################################
        '
        If (VMITX > VMITY) Then
            VMITEdgeSeq = "E1L"
            VMITEdgeCode = VeneerEdgeCode
        End If

        If (VMITX < VMITY) Then
            VMITEdgeSeq = "E1S"
            VMITEdgeCode = VeneerEdgeCode
        End If

        '#################################################
        '# VANITY MATCHING INTERIOR BOTTOM EDGE SEQUENCE #
        '#################################################
        '
        If (VMIBX > VMIBY) Then
            VMIBEdgeSeq = "E1L"
            VMIBEdgeCode = VeneerEdgeCode
        End If

        If (VMIBX < VMIBY) Then
            VMIBEdgeSeq = "E1S"
            VMIBEdgeCode = VeneerEdgeCode
        End If

        '###########################################################
        '# VANITY MATCHING INTERIOR FULL FIXED SHELF EDGE SEQUENCE #
        '###########################################################
        '
        If (VMIFFSHX > VMIFFSHY) Then
            VMIFFSHEdgeSeq = "E1L"
            VMIFFSHEdgeCode = VeneerEdgeCode
        End If

        If (VMIFFSHX < VMIFFSHY) Then
            VMIFFSHEdgeSeq = "E1S"
            VMIFFSHEdgeCode = VeneerEdgeCode
        End If

        '##################################################
        '# VANITY MATCHING INTERIOR DIVIDER EDGE SEQUENCE #
        '##################################################
        '
        VMIDEdgeSeq = "E1S"
        VMIDEdgeCode = VeneerEdgeCode

        '################################################
        '# VANITY MATCHING INTERIOR STRAP EDGE SEQUENCE #
        '################################################
        '
        If (VMISX > VMISY) Then
            VMISEdgeSeq = "E1L"
            VMISEdgeCode = VeneerEdgeCode
        End If

        If (VMISX < VMISY) Then
            VMISEdgeSeq = "E1S"
            VMISEdgeCode = VeneerEdgeCode
        End If

        '#####################################################################################################################################################################################
        '# VANITY ELEVATED EDGING RULES AND VARIABLES #
        '##############################################
        '
        Dim VEGX As Double
        Dim VEGY As Double
        VEGX = VEGableX
        VEGY = VEGableY

        Dim VETSX As Double
        Dim VETSY As Double
        VETSX = VETopStrapX
        VETSY = VETopStrapY

        Dim VESX As Double
        Dim VESY As Double
        VESX = VEStrapX
        VESY = VEStrapY

        Dim VETX As Double
        Dim VETY As Double
        VETX = VETopX
        VETY = VETopY

        Dim VEBX As Double
        Dim VEBY As Double
        VEBX = VEBtmX
        VEBY = VEBtmY

        Dim VEFFSHX As Double
        Dim VEFFSHY As Double
        VEFFSHX = VEFFShelfX
        VEFFSHY = VEFFShelfY

        '#######################################
        '# VANITY ELEVATED GABLE EDGE SEQUENCE #
        '#######################################
        '
        If (VEGX > VEGY) Then
            VEGEdgeSeq = "E1L"
            VEGEdgeCode = PVCEdgeCode
        End If

        If (VEGX < VEGY) Then
            VEGEdgeSeq = "E1S"
            VEGEdgeCode = PVCEdgeCode
        End If

        '###########################################
        '# VANITY ELEVATED TOP STRAP EDGE SEQUENCE #
        '###########################################
        '
        '###############
        '# GROUP RULES #
        '###############
        '
        If (CKWOPLANNERSGroupBox = "GROUP1") Then

            If (VETSX > VETSY) Then
                VETSEdgeSeq = "E1L"
                VETSEdgeCode = BMEdgeCode
            End If

            If (VETSX < VETSY) Then
                VETSEdgeSeq = "E1S"
                VETSEdgeCode = BMEdgeCode
            End If

        Else

            If (VETSX > VETSY) Then
                VETSEdgeSeq = "E1L"
                VETSEdgeCode = PVCEdgeCode
            End If

            If (VETSX < VETSY) Then
                VETSEdgeSeq = "E1S"
                VETSEdgeCode = PVCEdgeCode
            End If

        End If

        '#####################################
        '# VANITY ELEVATED TOP EDGE SEQUENCE #
        '#####################################
        '
        '###############
        '# GROUP RULES #
        '###############
        '
        If (CKWOPLANNERSGroupBox = "GROUP1") Then

            If (VETX > VETY) Then
                VETEdgeSeq = "E1L"
                VETEdgeCode = BMEdgeCode
            End If

            If (VETX < VETY) Then
                VETEdgeSeq = "E1S"
                VETEdgeCode = BMEdgeCode
            End If

        Else

            If (VETX > VETY) Then
                VETEdgeSeq = "E1L"
                VETEdgeCode = PVCEdgeCode
            End If

            If (VETX < VETY) Then
                VETEdgeSeq = "E1S"
                VETEdgeCode = PVCEdgeCode
            End If

        End If

        '########################################
        '# VANITY ELEVATED BOTTOM EDGE SEQUENCE #
        '########################################
        '
        '##########
        '# GROUP RULES #
        '##########
        '
        If (VEBX > VEBY) Then
            VEBEdgeSeq = "E1L"
            VEBEdgeCode = BMEdgeCode
        End If

        If (VEBX < VEBY) Then
            VEBEdgeSeq = "E1S"
            VEBEdgeCode = BMEdgeCode
        End If

        '##################################################
        '# VANITY ELEVATED FULL FIXED SHELF EDGE SEQUENCE #
        '##################################################
        '
        If (VEFFSHX > VEFFSHY) Then
            VEFFSHEdgeSeq = "E1L"
            VEFFSHEdgeCode = PVCEdgeCode
        End If

        If (VEFFSHX < VEFFSHY) Then
            VEFFSHEdgeSeq = "E1S"
            VEFFSHEdgeCode = PVCEdgeCode
        End If

        '#########################################
        '# VANITY ELEVATED DIVIDER EDGE SEQUENCE #
        '#########################################
        '
        VEDEdgeSeq = "E1S"
        VEDEdgeCode = PVCEdgeCode

        '#######################################
        '# VANITY ELEVATED STRAP EDGE SEQUENCE #
        '#######################################
        '
        If (VESX > VESY) Then
            VESEdgeSeq = "E1L"
            VESEdgeCode = PVCEdgeCode
        End If

        If (VESX < VESY) Then
            VESEdgeSeq = "E1S"
            VESEdgeCode = PVCEdgeCode
        End If

        '#####################################################################################################################################################################################
        '# VANITY ELEVATED MATCHING INTERIOR EDGING RULES AND VARIABLES #
        '################################################################
        '
        Dim VEMIGX As Double
        Dim VEMIGY As Double
        VEMIGX = VEMIGableX
        VEMIGY = VEMIGableY

        Dim VEMITSX As Double
        Dim VEMITSY As Double
        VEMITSX = VEMITopStrapX
        VEMITSY = VEMITopStrapY

        Dim VEMISX As Double
        Dim VEMISY As Double
        VEMISX = VEMIStrapX
        VEMISY = VEMIStrapY

        Dim VEMITX As Double
        Dim VEMITY As Double
        VEMITX = VEMITopX
        VEMITY = VEMITopY

        Dim VEMIBX As Double
        Dim VEMIBY As Double
        VEMIBX = VEMIBtmX
        VEMIBY = VEMIBtmY

        Dim VEMIFFSHX As Double
        Dim VEMIFFSHY As Double
        VEMIFFSHX = VEMIFFShelfX
        VEMIFFSHY = VEMIFFShelfY

        '#########################################################
        '# VANITY ELEVATED MATCHING INTERIOR GABLE EDGE SEQUENCE #
        '#########################################################
        '
        If (VEMIGX > VEMIGY) Then
            VEMIGEdgeSeq = "E1L"
            VEMIGEdgeCode = VeneerEdgeCode
        End If

        If (VEMIGX < VEMIGY) Then
            VEMIGEdgeSeq = "E1S"
            VEMIGEdgeCode = VeneerEdgeCode
        End If

        '#############################################################
        '# VANITY ELEVATED MATCHING INTERIOR TOP STRAP EDGE SEQUENCE #
        '#############################################################
        '
        If (VEMITSX > VEMITSY) Then
            VEMITSEdgeSeq = "E1L"
            VEMITSEdgeCode = VeneerEdgeCode
        End If

        If (VEMITSX < VEMITSY) Then
            VEMITSEdgeSeq = "E1S"
            VEMITSEdgeCode = VeneerEdgeCode
        End If

        '#######################################################
        '# VANITY ELEVATED MATCHING INTERIOR TOP EDGE SEQUENCE #
        '#######################################################
        '
        If (VEMITX > VEMITY) Then
            VEMITEdgeSeq = "E1L"
            VEMITEdgeCode = VeneerEdgeCode
        End If

        If (VEMITX < VEMITY) Then
            VEMITEdgeSeq = "E1S"
            VEMITEdgeCode = VeneerEdgeCode
        End If

        '##########################################################
        '# VANITY ELEVATED MATCHING INTERIOR BOTTOM EDGE SEQUENCE #
        '##########################################################
        '
        If (VEMIBX > VEMIBY) Then
            VEMIBEdgeSeq = "E1L"
            VEMIBEdgeCode = VeneerEdgeCode
        End If

        If (VEMIBX < VEMIBY) Then
            VEMIBEdgeSeq = "E1S"
            VEMIBEdgeCode = VeneerEdgeCode
        End If

        '####################################################################
        '# VANITY ELEVATED MATCHING INTERIOR FULL FIXED SHELF EDGE SEQUENCE #
        '####################################################################
        '
        If (VEMIFFSHX > VEMIFFSHY) Then
            VEMIFFSHEdgeSeq = "E1L"
            VEMIFFSHEdgeCode = VeneerEdgeCode
        End If

        If (VEMIFFSHX < VEMIFFSHY) Then
            VEMIFFSHEdgeSeq = "E1S"
            VEMIFFSHEdgeCode = VeneerEdgeCode
        End If

        '###########################################################
        '# VANITY ELEVATED MATCHING INTERIOR DIVIDER EDGE SEQUENCE #
        '###########################################################
        '
        VEMIDEdgeSeq = "E1S"
        VEMIDEdgeCode = VeneerEdgeCode

        '#########################################################
        '# VANITY ELEVATED MATCHING INTERIOR STRAP EDGE SEQUENCE #
        '#########################################################
        '
        If (VEMISX > VEMISY) Then
            VEMISEdgeSeq = "E1L"
            VEMISEdgeCode = VeneerEdgeCode
        End If

        If (VEMISX < VEMISY) Then
            VEMISEdgeSeq = "E1S"
            VEMISEdgeCode = VeneerEdgeCode
        End If

        '#####################################################################################################################################################################################
        '# CANOPY EDGING RULES AND VARIABLES #
        '#####################################
        '
        Dim RHGX As Double
        Dim RHGY As Double
        RHGX = RHGableX
        RHGY = RHGableY

        Dim RHTX As Double
        Dim RHTY As Double
        RHTX = RHTopX
        RHTY = RHTopY

        Dim RHFX As Double
        Dim RHFY As Double
        RHFX = RHFrontX
        RHFY = RHFrontY

        Dim RHBX As Double
        Dim RHBY As Double
        RHBX = RHBackX
        RHBY = RHBackY

        '##############################
        '# CANOPY GABLE EDGE SEQUENCE #
        '##############################
        '
        If (RHGX > RHGY) Then
            RHGEdgeSeq = "E2S"
            RHGEdgeCode = VeneerEdgeCode
            RHGEdgeSeq2 = "E1L"
            RHGEdgeCode2 = VeneerEdgeCode
        End If
        If (RHGX < RHGY) Then
            RHGEdgeSeq = "E2L"
            RHGEdgeCode = VeneerEdgeCode
            RHGEdgeSeq2 = "E1S"
            RHGEdgeCode2 = VeneerEdgeCode
        End If

        '############################
        '# CANOPY TOP EDGE SEQUENCE #
        '############################
        '
        If (RHTX > RHTY) Then
            RHTEdgeSeq = "E1L"
            RHTEdgeCode = VeneerEdgeCode
        End If
        If (RHTX < RHTY) Then
            RHTEdgeSeq = "E1S"
            RHTEdgeCode = VeneerEdgeCode
        End If

        '###############################
        '# CANOPY FRONT EDGE SEQUENCE #
        '###############################
        '
        If (CKWOPLANNERSSpeciesBox = "MAPLE" Or CKWOPLANNERSSpeciesBox = "OAK" Or CKWOPLANNERSSpeciesBox = "CHERRY" Or CKWOPLANNERSSpeciesBox = "PINE" Or CKWOPLANNERSSpeciesBox = "WALNUT") Then
            RHFEdgeSeq = "E4S"
            RHFEdgeCode = VeneerEdgeCode
        Else
            RHFEdgeSeq = ""
            RHFEdgeCode = ""
        End If


        '#############################
        '# CANOPY BACK EDGE SEQUENCE #
        '#############################
        '
        '
        If (RHBX > RHBY) Then
            RHBEdgeSeq = "E1S"
            RHBEdgeCode = VeneerEdgeCode
        End If

        If (RHBX < RHBY) Then
            RHBEdgeSeq = "E1L"
            RHBEdgeCode = VeneerEdgeCode
        End If

        '#####################################################################################################################################################################################
        '# BASE INTEGRATED APPLIANCE EDGING RULES AND VARIABLES #
        '########################################################
        '
        Dim BIAAGX As Double
        Dim BIAAGY As Double
        BIAAGX = BIAAGableX
        BIAAGY = BIAAGableY

        Dim BIAATSX As Double
        Dim BIAATSY As Double
        BIAATSX = BIAATopStrapX
        BIAATSY = BIAATopStrapY

        Dim BIAABX As Double
        Dim BIAABY As Double
        BIAABX = BIAABottomX
        BIAABY = BIAABottomY

        Dim BIABGX As Double
        Dim BIABGY As Double
        BIABGX = BIABGableX
        BIABGY = BIABGableY

        Dim BIABTSX As Double
        Dim BIABTSY As Double
        BIABTSX = BIABTopStrapX
        BIABTSY = BIABTopStrapY

        Dim BIABBX As Double
        Dim BIABBY As Double
        BIABBX = BIABBottomX
        BIABBY = BIABBottomY

        '############################################################
        '# BASE INTEGRATED APPLIANCE, APPLIANCE GABLE EDGE SEQUENCE #
        '############################################################
        '
        If (BIAAGX > BIAAGY) Then
            BIAAGEdgeSeq = "E2S"
            BIAAGEdgeCode = BMEdgeCode
            BIAAGEdgeSeq2 = "E1L"
            BIAAGEdgeCode2 = PVCEdgeCode
        End If
        If (BIAAGX < BIAAGY) Then
            BIAAGEdgeSeq = "E2L"
            BIAAGEdgeCode = BMEdgeCode
            BIAAGEdgeSeq2 = "E1S"
            BIAAGEdgeCode2 = PVCEdgeCode
        End If

        '################################################################
        '# BASE INTEGRATED APPLIANCE, APPLIANCE TOP STRAP EDGE SEQUENCE #
        '################################################################
        '
        If (BIAATSX > BIAATSY) Then
            BIAATSEdgeSeq = "E1L"
            BIAATSEdgeCode = PVCEdgeCode
        End If

        If (BIAATSX < BIAATSY) Then
            BIAATSEdgeSeq = "E1S"
            BIAATSEdgeCode = PVCEdgeCode
        End If

        '#############################################################
        '# BASE INTEGRATED APPLIANCE, APPLIANCE BOTTOM EDGE SEQUENCE #
        '#############################################################
        '
        If (BIAABX > BIAABY) Then
            BIAABEdgeSeq = "E1L"
            BIAABEdgeCode = PVCEdgeCode
        End If

        If (BIAABX < BIAABY) Then
            BIAABEdgeSeq = "E1S"
            BIAABEdgeCode = PVCEdgeCode
        End If

        '#######################################################
        '# BASE INTEGRATED APPLIANCE, BASE GABLE EDGE SEQUENCE #
        '#######################################################
        '
        If (BIABGX > BIABGY) Then
            BIABGEdgeSeq = "E1L"
            BIABGEdgeCode = PVCEdgeCode
        End If
        If (BIABGX < BIABGY) Then
            BIABGEdgeSeq = "E1S"
            BIABGEdgeCode = PVCEdgeCode
        End If

        '###########################################################
        '# BASE INTEGRATED APPLIANCE, BASE TOP STRAP EDGE SEQUENCE #
        '###########################################################
        '
        If (BIAATSX > BIAATSY) Then
            BIAATSEdgeSeq = "E1L"
            BIAATSEdgeCode = PVCEdgeCode
        End If

        If (BIAATSX < BIAATSY) Then
            BIAATSEdgeSeq = "E1S"
            BIAATSEdgeCode = PVCEdgeCode
        End If

        '########################################################
        '# BASE INTEGRATED APPLIANCE, BASE BOTTOM EDGE SEQUENCE #
        '########################################################
        '
        If (BIAABX > BIAABY) Then
            BIAABEdgeSeq = "E1L"
            BIAABEdgeCode = BMEdgeCode
        End If

        If (BIAABX < BIAABY) Then
            BIAABEdgeSeq = "E1S"
            BIAABEdgeCode = BMEdgeCode
        End If

        '#####################################################################################################################################################################################
        '# WRITE AND APPEND TO CUT LIST #
        '################################
        '
        Dim CabType = CutlistForm.CabCodeBox1.Text

        '##################################
        '# SELECT CUTLIST FROM USER INPUT #
        '##################################
        '
        Select Case CabType
            '#####################################################################################################################################################################################
            '# UPPER CUTLIST PARTS AND ALTERNATIVES #
            '########################################
            '
            '#############################
            '# UPPER # 9 LINES #
            '#############################
            '
            Case "UPPER"
                Dim UCase = CutlistForm.UpperCabCodeBox.Text
                Select Case UCase
                    Case "U", "UB", "UO", "US", "UU", "USHF", "USHK", "USHL", "USHS", "UUHF", "UUHK", "UUHL", "UUHS"
                        CabCode = CutlistForm.UpperCabCodeBox.Text & CutlistForm.OptionsVar
                        If (CutlistForm.MatchingInteriorCheck.Checked) Then
                            If (UMIAdjShelfQ = 0) Then
                                DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                                DGV.Rows.Add(CabSize1)
                                DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                                DGV.Rows.Add(UMIBackQ, "Back", UMIBackY, UMIBackX, UMIBackZ, UMIBackM)
                                '# TOP LIGHT CHECK #
                                If (CutlistForm.TLightCheck.Checked) Then
                                    DGV.Rows.Add(UMITopLightQ, "Top Light", UMITopLightY, UMITopLightX, UMITopLightZ, UMITopLightM, UMITLEdgeSeq, UMITLEdgeCode)
                                    DGV.Rows.Add(UMIBtmQ, "Bottom", UMIBtmY, UMIBtmX, UMIBtmZ, UMIBtmM, UMIBEdgeSeq, UMIBEdgeCode)
                                Else
                                    DGV.Rows.Add(UMITopBtmQ, "TopBtm", UMITopBtmY, UMITopBtmX, UMITopBtmZ, UMITopBtmM, UMITBEdgeSeq, UMITBEdgeCode)
                                End If
                                DGV.Rows.Add(UMIGableQ, "Gable", UMIGableY, UMIGableX, UMIGableZ, UMIGableM, UMIGEdgeSeq, UMIGEdgeCode, UMIGEdgeSeq2, UMIGEdgeCode)
                                DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                            Else
                                DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                                DGV.Rows.Add(CabSize1)
                                DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                                DGV.Rows.Add(UMIBackQ, "Back", UMIBackY, UMIBackX, UMIBackZ, UMIBackM)
                                '# TOP LIGHT CHECK #
                                If (CutlistForm.TLightCheck.Checked) Then
                                    DGV.Rows.Add(UMITopLightQ, "Top Light", UMITopLightY, UMITopLightX, UMITopLightZ, UMITopLightM, UMITLEdgeSeq, UMITLEdgeCode)
                                    DGV.Rows.Add(UMIBtmQ, "Bottom", UMIBtmY, UMIBtmX, UMIBtmZ, UMIBtmM, UMIBEdgeSeq, UMIBEdgeCode)
                                Else
                                    DGV.Rows.Add(UMITopBtmQ, "TopBtm", UMITopBtmY, UMITopBtmX, UMITopBtmZ, UMITopBtmM, UMITBEdgeSeq, UMITBEdgeCode)
                                End If
                                '# GLASS SHELF CHECK #
                                If (CutlistForm.GlassCheck.Checked) Then
                                    DGV.Rows.Add(UMIAdjShelfQ, "Adj.Sh", "", "", "", "Glass", "", "", "", "", "Ordered Joe G")
                                Else
                                    If (CKWOPLANNER.SSpeciesBox1.Text = "MDF") Then
                                        DGV.Rows.Add(UMIAdjShelfQ, "Adj.Sh", UMIAdjShelfY, UMIAdjShelfX, UMIAdjShelfZ, UMIAdjShelfM)
                                    Else
                                        DGV.Rows.Add(UMIAdjShelfQ, "Adj.Sh", UMIAdjShelfY, UMIAdjShelfX, UMIAdjShelfZ, UMIAdjShelfM, UMIASEdgeSeq, UMIASEdgeCode)
                                    End If
                                End If
                                DGV.Rows.Add(UMIGableQ, "Gable", UMIGableY, UMIGableX, UMIGableZ, UMIGableM, UMIGEdgeSeq, UMIGEdgeCode, UMIGEdgeSeq2, UMIGEdgeCode)
                                DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                            End If
                        Else
                            If (UAdjShelfQ = 0) Then
                                DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                                DGV.Rows.Add(CabSize1)
                                DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                                DGV.Rows.Add(UBackQ, "Back", UBackY, UBackX, UBackZ, UBackM)
                                '# TOP LIGHT CHECK #
                                If (CutlistForm.TLightCheck.Checked) Then
                                    DGV.Rows.Add(UTopLightQ, "Top Light", UTopLightY, UTopLightX, UTopLightZ, UTopLightM, UTLEdgeSeq, UTLEdgeCode)
                                    DGV.Rows.Add(UBtmQ, "Bottom", UBtmY, UBtmX, UBtmZ, UBtmM, UBEdgeSeq, UBEdgeCode)
                                Else
                                    DGV.Rows.Add(UTopBtmQ, "TopBtm", UTopBtmY, UTopBtmX, UTopBtmZ, UTopBtmM, UTBEdgeSeq, UTBEdgeCode)
                                End If
                                DGV.Rows.Add(UGableQ, "Gable", UGableY, UGableX, UGableZ, UGableM, UGEdgeSeq, UGEdgeCode, UGEdgeSeq2, UGEdgeCode2)
                                DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                            Else
                                DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                                DGV.Rows.Add(CabSize1)
                                DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                                DGV.Rows.Add(UBackQ, "Back", UBackY, UBackX, UBackZ, UBackM)
                                '# TOP LIGHT CHECK #
                                If (CutlistForm.TLightCheck.Checked) Then
                                    DGV.Rows.Add(UTopLightQ, "Top Light", UTopLightY, UTopLightX, UTopLightZ, UTopLightM, UTLEdgeSeq, UTLEdgeCode)
                                    DGV.Rows.Add(UBtmQ, "Bottom", UBtmY, UBtmX, UBtmZ, UBtmM, UBEdgeSeq, UBEdgeCode)
                                Else
                                    DGV.Rows.Add(UTopBtmQ, "TopBtm", UTopBtmY, UTopBtmX, UTopBtmZ, UTopBtmM, UTBEdgeSeq, UTBEdgeCode)
                                End If
                                '# GLASS SHELF CHECK #
                                If (CutlistForm.GlassCheck.Checked) Then
                                    DGV.Rows.Add(UAdjShelfQ, "Adj.Sh", "", "", "", "Glass", "", "", "", "", "Ordered Joe G")
                                Else
                                    '# MDF SPECIES CHECK #
                                    If (CKWOPLANNER.SSpeciesBox1.Text = "MDF") Then
                                        DGV.Rows.Add(UAdjShelfQ, "Adj.Sh", UAdjShelfY, UAdjShelfX, UAdjShelfZ, UAdjShelfM)
                                    Else
                                        DGV.Rows.Add(UAdjShelfQ, "Adj.Sh", UAdjShelfY, UAdjShelfX, UAdjShelfZ, UAdjShelfM, UASEdgeSeq, UASEdgeCode)
                                    End If

                                End If
                                DGV.Rows.Add(UGableQ, "Gable", UGableY, UGableX, UGableZ, UGableM, UGEdgeSeq, UGEdgeCode, UGEdgeSeq2, UGEdgeCode2)
                                DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                            End If
                        End If

                    Case "UL", "USL", "UUL", "UDL", "UUDL", "USDL"
                        CabCode = CutlistForm.UpperCabCodeBox.Text & CutlistForm.OptionsVar
                        '# UPPER MATCHING INTERIOR CHECK #
                        If (CutlistForm.MatchingInteriorCheck.Checked) Then
                            If (UMIAdjShelfQ = 0) Then
                                DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                                DGV.Rows.Add(CabSize1)
                                DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                                DGV.Rows.Add(UMIBackQ, "Back", UMIBackY, UMIBackX, UMIBackZ, UMIBackM)
                                '# TOP LIGHT CHECK #
                                If (CutlistForm.TLightCheck.Checked) Then
                                    DGV.Rows.Add(UMITopBtmQ, "Top Light", UMITopBtmY, UMITopBtmX, UMITopBtmZ, UMITopBtmM, UMITBEdgeSeq, UMITBEdgeCode)
                                Else
                                    DGV.Rows.Add(UMITopBtmQ, "Top Hinge", UMITopBtmY, UMITopBtmX, UMITopBtmZ, UMITopBtmM, UMITBEdgeSeq, UMITBEdgeCode)
                                End If
                                DGV.Rows.Add(UMITopBtmQ, "Bottom", UMITopBtmY, UMITopBtmX, UMITopBtmZ, UMITopBtmM, UMITBEdgeSeq, UMITBEdgeCode)
                                DGV.Rows.Add(UMIGableQ, "Gable", UMIGableY, UMIGableX, UMIGableZ, UMIGableM, UMIGEdgeSeq, UMIGEdgeCode, UMIGEdgeSeq2, UMIGEdgeCode)
                                DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                            Else
                                DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                                DGV.Rows.Add(CabSize1)
                                DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                                DGV.Rows.Add(UMIBackQ, "Back", UMIBackY, UMIBackX, UMIBackZ, UMIBackM)
                                '# TOP LIGHT CHECK #
                                If (CutlistForm.TLightCheck.Checked) Then
                                    DGV.Rows.Add(UMITopBtmQ, "Top Light", UMITopBtmY, UMITopBtmX, UMITopBtmZ, UMITopBtmM, UMITBEdgeSeq, UMITBEdgeCode)
                                Else
                                    DGV.Rows.Add(UMITopBtmQ, "Top Hinge", UMITopBtmY, UMITopBtmX, UMITopBtmZ, UMITopBtmM, UMITBEdgeSeq, UMITBEdgeCode)
                                End If
                                DGV.Rows.Add(UMITopBtmQ, "Bottom", UMITopBtmY, UMITopBtmX, UMITopBtmZ, UMITopBtmM, UMITBEdgeSeq, UMITBEdgeCode)
                                '# GLASS SHELF CHECK #
                                If (CutlistForm.GlassCheck.Checked) Then
                                    DGV.Rows.Add(UMIAdjShelfQ, "Adj.Sh", "", "", "", "Glass", "", "", "", "", "Ordered Joe G")
                                Else
                                    '# MDF SPECIES CHECK #
                                    If (CKWOPLANNER.SSpeciesBox1.Text = "MDF") Then
                                        DGV.Rows.Add(UMIAdjShelfQ, "Adj.Sh", UMIAdjShelfY, UMIAdjShelfX, UMIAdjShelfZ, UMIAdjShelfM)
                                    Else
                                        DGV.Rows.Add(UMIAdjShelfQ, "Adj.Sh", UMIAdjShelfY, UMIAdjShelfX, UMIAdjShelfZ, UMIAdjShelfM, UMIASEdgeSeq, UMIASEdgeCode)
                                    End If
                                End If
                                DGV.Rows.Add(UMIGableQ, "Gable", UMIGableY, UMIGableX, UMIGableZ, UMIGableM, UMIGEdgeSeq, UMIGEdgeCode, UMIGEdgeSeq2, UMIGEdgeCode)
                                DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                            End If
                        Else
                            If (UAdjShelfQ = 0) Then
                                DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                                DGV.Rows.Add(CabSize1)
                                DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                                DGV.Rows.Add(UBackQ, "Back", UBackY, UBackX, UBackZ, UBackM)
                                '# TOP LIGHT CHECK #
                                If (CutlistForm.TLightCheck.Checked) Then
                                    DGV.Rows.Add(UTopHingeQ, "Top Light", UTopHingeY, UTopHingeX, UTopHingeZ, UTopHingeM, UTHEdgeSeq, UTHEdgeCode)
                                Else
                                    DGV.Rows.Add(UTopHingeQ, "Top Hinge", UTopHingeY, UTopHingeX, UTopHingeZ, UTopHingeM, UTHEdgeSeq, UTHEdgeCode)
                                End If
                                DGV.Rows.Add(UBtmQ, "Bottom", UBtmY, UBtmX, UBtmZ, UBtmM, UBEdgeSeq, UBEdgeCode)
                                DGV.Rows.Add(UGableQ, "Gable", UGableY, UGableX, UGableZ, UGableM, UGEdgeSeq, UGEdgeCode, UGEdgeSeq2, UGEdgeCode2)
                                DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                            Else
                                DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                                DGV.Rows.Add(CabSize1)
                                DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                                DGV.Rows.Add(UBackQ, "Back", UBackY, UBackX, UBackZ, UBackM)
                                '# TOP LIGHT CHECK #
                                If (CutlistForm.TLightCheck.Checked) Then
                                    DGV.Rows.Add(UTopHingeQ, "Top Light", UTopHingeY, UTopHingeX, UTopHingeZ, UTopHingeM, UTHEdgeSeq, UTHEdgeCode)
                                Else
                                    DGV.Rows.Add(UTopHingeQ, "Top Hinge", UTopHingeY, UTopHingeX, UTopHingeZ, UTopHingeM, UTHEdgeSeq, UTHEdgeCode)
                                End If
                                DGV.Rows.Add(UBtmQ, "Bottom", UBtmY, UBtmX, UBtmZ, UBtmM, UBEdgeSeq, UBEdgeCode)
                                '# GLASS SHELF CHECK #
                                If (CutlistForm.GlassCheck.Checked) Then
                                    DGV.Rows.Add(UAdjShelfQ, "Adj.Sh", "", "", "", "Glass", "", "", "", "", "Ordered Joe G")
                                Else
                                    '# MDF SPECIES CHECK #
                                    If (CKWOPLANNER.SSpeciesBox1.Text = "MDF") Then
                                        DGV.Rows.Add(UAdjShelfQ, "Adj.Sh", UAdjShelfY, UAdjShelfX, UAdjShelfZ, UAdjShelfM)
                                    Else
                                        DGV.Rows.Add(UAdjShelfQ, "Adj.Sh", UAdjShelfY, UAdjShelfX, UAdjShelfZ, UAdjShelfM, UASEdgeSeq, UASEdgeCode)
                                    End If

                                End If
                                DGV.Rows.Add(UGableQ, "Gable", UGableY, UGableX, UGableZ, UGableM, UGEdgeSeq, UGEdgeCode, UGEdgeSeq2, UGEdgeCode2)
                                DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                            End If
                        End If

                    Case "UA"
                        '# MATCHING INTERIOR CHECK #
                        If (CutlistForm.MatchingInteriorCheck.Checked) Then
                            CabCode = CutlistForm.UpperCabCodeBox.Text & CutlistForm.OptionsVar
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(UMIBackQ, "Back", UMIBackY, UMIBackX, UMIBackZ, UMIBackM)
                            '# TOP LIGHT CHECK #
                            If (CutlistForm.TLightCheck.Checked) Then
                                DGV.Rows.Add(UTopQ, "Top Light", UMITopBtmY, UMITopBtmX, UMITopBtmZ, UMITopBtmM, UMITBEdgeSeq, UMITBEdgeCode)
                                DGV.Rows.Add(UTopQ, "Bottom", UMITopBtmY, UMITopBtmX, UMITopBtmZ, UMITopBtmM, UMITBEdgeSeq, UMITBEdgeCode)
                            Else
                                DGV.Rows.Add(UTopBtmQ, "TopBot", UMITopBtmY, UMITopBtmX, UMITopBtmZ, UMITopBtmM, UMITBEdgeSeq, UMITBEdgeCode)
                            End If
                            DGV.Rows.Add(UMIGableQ, "Gable", UMIGableY, UMIGableX, UMIGableZ, UMIGableM, UMIGEdgeSeq, UMIGEdgeCode, UMIGEdgeSeq2, UMIGEdgeCode2)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                        Else
                            CabCode = CutlistForm.UpperCabCodeBox.Text & CutlistForm.OptionsVar
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(UBackQ, "Back", UBackY, UBackX, UBackZ, UBackM)
                            '# TOP LIGHT CHECK #
                            If (CutlistForm.TLightCheck.Checked) Then
                                DGV.Rows.Add(UTopQ, "Top Light", UTopY, UTopX, UTopZ, UTopBtmM, UTEdgeSeq, UTEdgeCode)
                                DGV.Rows.Add(UTopQ, "Bottom", UTopY, UTopX, UTopZ, UTopBtmM, UTEdgeSeq, UTEdgeCode)
                            Else
                                DGV.Rows.Add(UTopBtmQ, "TopBot", UTopY, UTopX, UTopZ, UTopBtmM, UTEdgeSeq, UTEdgeCode)
                            End If
                            DGV.Rows.Add(UFanSHQ, "Fan Sh", UFanSHY, UFanSHX, UFanSHZ, UFanSHM)
                            DGV.Rows.Add(UGableQ, "Gable", UGableY, UGableX, UGableZ, UGableM, UGEdgeSeq, UGEdgeCode, UGEdgeSeq2, UGEdgeCode2)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                        End If

                    Case "UTDF"
                        '# MATCHING INTERIOR CHECK #
                        If (CutlistForm.MatchingInteriorCheck.Checked) Then
                            '# FILL ME IN LATER #
                        Else
                            CabCode = CutlistForm.UpperCabCodeBox.Text & CutlistForm.OptionsVar
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(UBackQ, "Back", UBackY, UBackX, UBackZ, UBackM)
                            '# TOP LIGHT CHECK #
                            If (CutlistForm.TLightCheck.Checked) Then
                                DGV.Rows.Add(UTopBtmQ, "Top Light", UTopBtmY, UTopBtmX, UTopBtmZ, UTopBtmM, UTBEdgeSeq, UTBEdgeCode)
                                DGV.Rows.Add(UTopBtmQ, "Bottom", UTopBtmY, UTopBtmX, UTopBtmZ, UTopBtmM, UTBEdgeSeq, UTBEdgeCode)
                            Else
                                DGV.Rows.Add(UTopBtmQ, "TopBtm", UTopBtmY, UTopBtmX, UTopBtmZ, UTopBtmM, UTBEdgeSeq, UTBEdgeCode)
                            End If
                            DGV.Rows.Add(UFDividerQ, "F.Divider", UFDividerY, UFDividerX, UFDividerZ, UFDividerM, UTFDEdgeSeq, UTFDEdgeCode)
                            DGV.Rows.Add(UDividerQ, "Divider", UDividerY, UDividerX, UDividerZ, UDividerM, UTDEdgeSeq, UTDEdgeCode)
                            DGV.Rows.Add(UGableQ, "Gable", UGableY, UGableX, UGableZ, UGableM, UGEdgeSeq, UGEdgeCode, UGEdgeSeq2, UGEdgeCode2)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                        End If

                    Case "UTDH"
                        '# MATCHING INTERIOR CHECK #
                        If (CutlistForm.MatchingInteriorCheck.Checked) Then
                            '# FILL IN LATER #
                        Else
                            CabCode = CutlistForm.UpperCabCodeBox.Text & CutlistForm.OptionsVar
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(UBackQ, "Back", UBackY, UBackX, UBackZ, UBackM)
                            '# TOP LIGHT CHECK #
                            If (CutlistForm.TLightCheck.Checked) Then
                                DGV.Rows.Add(UTopBtmQ, "Top Light", UTopBtmY, UTopBtmX, UTopBtmZ, UTopBtmM, UTBEdgeSeq, UTBEdgeCode)
                                DGV.Rows.Add(UTopBtmQ, "Bottom", UTopBtmY, UTopBtmX, UTopBtmZ, UTopBtmM, UTBEdgeSeq, UTBEdgeCode)
                            Else
                                DGV.Rows.Add(UTopBtmQ, "TopBtm", UTopBtmY, UTopBtmX, UTopBtmZ, UTopBtmM, UTBEdgeSeq, UTBEdgeCode)
                            End If
                            DGV.Rows.Add(UFDividerQ, "F.Divider", UFDividerY, UFDividerX, UFDividerZ, UFDividerM, UTFDEdgeSeq, UTFDEdgeCode)
                            DGV.Rows.Add(UDividerQ, "Divider", UDividerY, UDividerX, UDividerZ, UDividerM, UTDEdgeSeq, UTDEdgeCode)
                            '# GLASS SHELF CHECK #
                            If (CutlistForm.GlassCheck.Checked) Then
                                DGV.Rows.Add(UAdjShelfQ, "Adj.Sh", "", "", "", "Glass", "", "", "", "", "Ordered Joe G")
                            Else
                                '# MDF SPECIES CHECK #
                                If (CKWOPLANNER.SSpeciesBox1.Text = "MDF") Then
                                    DGV.Rows.Add(UAdjShelfQ, "Adj.Sh", UAdjShelfY, UAdjShelfX, UAdjShelfZ, UAdjShelfM)
                                Else
                                    DGV.Rows.Add(UAdjShelfQ, "Adj.Sh", UAdjShelfY, UAdjShelfX, UAdjShelfZ, UAdjShelfM, UASEdgeSeq, UASEdgeCode)
                                End If

                            End If
                            DGV.Rows.Add(UGableQ, "Gable", UGableY, UGableX, UGableZ, UGableM, UGEdgeSeq, UGEdgeCode, UGEdgeSeq2, UGEdgeCode2)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                        End If

                    Case "UWRL"
                        CabCode = CutlistForm.UpperCabCodeBox.Text & CutlistForm.OptionsVar
                        DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                        DGV.Rows.Add(CabSize1)
                        DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                        DGV.Rows.Add(UWRBackQ, "Back", UWRBackY, UWRBackX, UWRBackZ, UWRBackM)
                        If (CutlistForm.TLightCheck.Checked) Then
                            DGV.Rows.Add(UWRTopBtmQ, "Top Light", UWRTopBtmY, UWRTopBtmX, UWRTopBtmZ, UWRTopBtmM, UWRTBEdgeSeq, UWRTBEdgeCode)
                            DGV.Rows.Add(UWRTopBtmQ, "Bottom", UWRTopBtmY, UWRTopBtmX, UWRTopBtmZ, UWRTopBtmM, UWRTBEdgeSeq, UWRTBEdgeCode)
                        Else
                            DGV.Rows.Add(UWRTopBtmQ, "TopBtm", UWRTopBtmY, UWRTopBtmX, UWRTopBtmZ, UWRTopBtmM, UWRTBEdgeSeq, UWRTBEdgeCode)
                        End If
                        DGV.Rows.Add(UWRGableQ, "Gable", UWRGableY, UWRGableX, UWRGableZ, UWRGableM, UWRGEdgeSeq, UWRGEdgeCode, UWRGEdgeSeq2, UWRGEdgeCode2)
                        DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")

                    Case "UWRLH"
                        CabCode = CutlistForm.UpperCabCodeBox.Text & CutlistForm.OptionsVar
                        DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                        DGV.Rows.Add(CabSize1)
                        DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                        DGV.Rows.Add(UWRBackQ, "Back", UWRBackY, UWRBackX, UWRBackZ, UWRBackM)
                        If (CutlistForm.TLightCheck.Checked) Then
                            DGV.Rows.Add(UWRTopBtmQ, "Top Light", UWRTopBtmY, UWRTopBtmX, UWRTopBtmZ, UWRTopBtmM, UWRTBEdgeSeq, UWRTBEdgeCode)
                            DGV.Rows.Add(UWRTopBtmQ, "Bottom", UWRTopBtmY, UWRTopBtmX, UWRTopBtmZ, UWRTopBtmM, UWRTBEdgeSeq, UWRTBEdgeCode)
                        Else
                            DGV.Rows.Add(UWRTopBtmQ, "TopBtm", UWRTopBtmY, UWRTopBtmX, UWRTopBtmZ, UWRTopBtmM, UWRTBEdgeSeq, UWRTBEdgeCode)
                        End If
                        DGV.Rows.Add(UWRFFSHQ, "F.F.Sh", UWRFFSHY, UWRFFSHX, UWRFFSHZ, UWRFFSHM, UWRFFSEdgeSeq, UWRFFSEdgeCode)
                        DGV.Rows.Add(UWRGableQ, "Gable", UWRGableY, UWRGableX, UWRGableZ, UWRGableM, UWRGEdgeSeq, UWRGEdgeCode, UWRGEdgeSeq2, UWRGEdgeCode2)
                        DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")

                    Case "UWRSQ"
                        CabCode = CutlistForm.UpperCabCodeBox.Text & CutlistForm.OptionsVar
                        DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                        DGV.Rows.Add(CabSize1)
                        DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                        DGV.Rows.Add(UWRBackQ, "Back", UWRBackY, UWRBackX, UWRBackZ, UWRBackM)
                        If (CutlistForm.TLightCheck.Checked) Then
                            DGV.Rows.Add(UWRTopBtmQ, "Top Light", UWRTopBtmY, UWRTopBtmX, UWRTopBtmZ, UWRTopBtmM, UWRTBEdgeSeq, UWRTBEdgeCode)
                            DGV.Rows.Add(UWRTopBtmQ, "Bottom", UWRTopBtmY, UWRTopBtmX, UWRTopBtmZ, UWRTopBtmM, UWRTBEdgeSeq, UWRTBEdgeCode)
                        Else
                            DGV.Rows.Add(UWRTopBtmQ, "TopBtm", UWRTopBtmY, UWRTopBtmX, UWRTopBtmZ, UWRTopBtmM, UWRTBEdgeSeq, UWRTBEdgeCode)
                        End If
                        DGV.Rows.Add(UWRFShelfQ, "F.Sh", UWRFShelfY, UWRFShelfX, UWRFShelfZ, UWRFShelfM, UWRFSEdgeSeq, UWRFSEdgeCode)
                        DGV.Rows.Add(UWRDivQ, "Divider", UWRDivY, UWRDivX, UWRDivZ, UWRDivM, UWRDEdgeSeq, UWRDEdgeCode)
                        DGV.Rows.Add(UWRGableQ, "Gable", UWRGableY, UWRGableX, UWRGableZ, UWRGableM, UWRGEdgeSeq, UWRGEdgeCode, UWRGEdgeSeq2, UWRGEdgeCode2)
                        DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")

                End Select

            '#############################
            '# UPPER MICROWAVE # ? LINES #
            '#############################
            '
            Case "UPPER MICROWAVE"
                Dim CH1 = (UMGableX + 1) / 10
                Dim CH2 = (UMMIGableX + 1) / 10
                Dim TH = CH1 + CH2
                CabCode = "UM"
                If (UMAdjShelfQ = 0) Then
                    Dim CabCode2 As String = "U"
                    DGV.Rows.Add(CabCode)
                    DGV.Rows.Add(VarWidthI & "-" & TH & "-" & VarDepthI)
                    DGV.Rows.Add(CabCode2, "Qty", VarAmountI)
                    DGV.Rows.Add(VarWidthI & "-" & CH1 & "-" & VarDepthI)
                    DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                    DGV.Rows.Add(UMBackQ, "Back", UMBackY, UMBackX, UMBackZ, UMBackM)
                    DGV.Rows.Add(UMTopQ, "Top", UMTopY, UMTopX, UMTopZ, UMTopM, UMTEdgeSeq, UMTEdgeCode)
                    DGV.Rows.Add(UMBtmQ, "Bottom", UMBtmY, UMBtmX, UMBtmZ, UMBtmM, UMBEdgeSeq, UMBEdgeCode)
                    DGV.Rows.Add(UMGableQ, "Gable", UMGableY, UMGableX, UMGableZ, UMGableM, UMGEdgeSeq, UMGEdgeCode, UMGEdgeSeq2, UMGEdgeCode2)
                    DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                    CabCode2 = "M"
                    DGV.Rows.Add(CabCode)
                    DGV.Rows.Add(VarWidthI & "-" & TH & "-" & VarDepthI)
                    DGV.Rows.Add(CabCode2, "Qty", VarAmountI)
                    DGV.Rows.Add(VarWidthI & "-" & CH2 & "-" & VarDepthI)
                    DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                    DGV.Rows.Add(UMMIBackQ, "Back", UMMIBackY, UMMIBackX, UMMIBackZ, UMMIBackM)
                    DGV.Rows.Add(UMMITopBtmQ, "TopBtm", UMMITopBtmY, UMMITopBtmX, UMMITopBtmZ, UMMITopBtmM, UMMITBEdgeSeq, UMMITBEdgeCode)
                    DGV.Rows.Add(UMMIMicroSHQ, "Micro.Sh", UMMIMicroSHY, UMMIMicroSHX, UMMIMicroSHZ, UMMIMicroSHM, UMMIMSEdgeSeq, UMMIMSEdgeCode, UMMIMSEdgeSeq2, UMMIMSEdgeCode2)
                    DGV.Rows.Add(UMMIGableQ, "Gable", UMMIGableY, UMMIGableX, UMMIGableZ, UMMIGableM, UMMIGEdgeSeq, UMMIGEdgeCode, UMMIGEdgeSeq2, UMMIGEdgeCode2)
                    DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                Else
                    Dim CabCode2 As String = "U"
                    DGV.Rows.Add(CabCode)
                    DGV.Rows.Add(VarWidthI & "-" & TH & "-" & VarDepthI)
                    DGV.Rows.Add(CabCode2, "Qty", VarAmountI)
                    DGV.Rows.Add(VarWidthI & "-" & CH1 & "-" & VarDepthI)
                    DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                    DGV.Rows.Add(UMBackQ, "Back", UMBackY, UMBackX, UMBackZ, UMBackM)
                    DGV.Rows.Add(UMTopQ, "Top", UMTopY, UMTopX, UMTopZ, UMTopM, UMTEdgeSeq, UMTEdgeCode)
                    DGV.Rows.Add(UMBtmQ, "Bottom", UMBtmY, UMBtmX, UMBtmZ, UMBtmM, UMBEdgeSeq, UMBEdgeCode)
                    DGV.Rows.Add(UMAdjShelfQ, "Adj.Sh", UMAdjShelfY, UMAdjShelfX, UMAdjShelfZ, UMAdjShelfM, UMASEdgeSeq, UMASEdgeCode)
                    DGV.Rows.Add(UMGableQ, "Gable", UMGableY, UMGableX, UMGableZ, UMGableM, UMGEdgeSeq, UMGEdgeCode, UMGEdgeSeq2, UMGEdgeCode2)
                    DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                    CabCode2 = "M"
                    DGV.Rows.Add(CabCode)
                    DGV.Rows.Add(VarWidthI & "-" & TH & "-" & VarDepthI)
                    DGV.Rows.Add(CabCode2, "Qty", VarAmountI)
                    DGV.Rows.Add(VarWidthI & "-" & CH2 & "-" & VarDepthI)
                    DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                    DGV.Rows.Add(UMMIBackQ, "Back", UMMIBackY, UMMIBackX, UMMIBackZ, UMMIBackM)
                    DGV.Rows.Add(UMMITopBtmQ, "TopBtm", UMMITopBtmY, UMMITopBtmX, UMMITopBtmZ, UMMITopBtmM, UMMITBEdgeSeq, UMMITBEdgeCode)
                    DGV.Rows.Add(UMMIMicroSHQ, "Micro.Sh", UMMIMicroSHY, UMMIMicroSHX, UMMIMicroSHZ, UMMIMicroSHM, UMMIMSEdgeSeq, UMMIMSEdgeCode, UMMIMSEdgeSeq2, UMMIMSEdgeCode2)
                    DGV.Rows.Add(UMMIGableQ, "Gable", UMMIGableY, UMMIGableX, UMMIGableZ, UMMIGableM, UMMIGEdgeSeq, UMMIGEdgeCode, UMMIGEdgeSeq2, UMMIGEdgeCode2)
                    DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                End If

            '#############################################
            '# UPPER CORNER DIAGONAL # 7 LINES / 8 LINES #
            '#############################################
            '
            Case "UPPER CORNER DIAGONAL"
                CabCode = "UCD" & CutlistForm.OptionsVar
                '# MATCHING INTERIOR CHECK #
                If (CutlistForm.MatchingInteriorCheck.Checked) Then
                    If (UCDMIAdjShelfQ = 0) Then
                        DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                        DGV.Rows.Add(CabSize1)
                        DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                        DGV.Rows.Add(UCDMILBackQ, "L-Back", UCDMILBackY, UCDMILBackX, UCDMILBackZ, UCDMILBackM)
                        DGV.Rows.Add(UCDMISBackQ, "S-Back", UCDMISBackY, UCDMISBackX, UCDMISBackZ, UCDMISBackM)
                        '# TOP LIGHT CHECK #
                        If (CutlistForm.TLightCheck.Checked) Then
                            DGV.Rows.Add(UCDMITopBtmQ, "Top Light", UCDMITopBtmY, UCDMITopBtmX, UCDMITopBtmZ, UCDMITopBtmM, UCDMITBEdgeSeq, UCDMITBEdgeCode)
                            DGV.Rows.Add(UCDMITopBtmQ, "Bottom", UCDMITopBtmY, UCDMITopBtmX, UCDMITopBtmZ, UCDMITopBtmM, UCDMITBEdgeSeq, UCDMITBEdgeCode)
                        Else
                            DGV.Rows.Add(UCDMITopBtmQ, "TopBtm", UCDMITopBtmY, UCDMITopBtmX, UCDMITopBtmZ, UCDMITopBtmM, UCDMITBEdgeSeq, UCDMITBEdgeCode)
                        End If
                        DGV.Rows.Add(UCDMIGableQ, "Gable", UCDMIGableY, UCDMIGableX, UCDMIGableZ, UCDMIGableM, UCDMIGEdgeSeq, UCDMIGEdgeCode, UCDMIGEdgeSeq2, UCDMIGEdgeCode2)
                        DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                    Else
                        DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                        DGV.Rows.Add(CabSize1)
                        DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                        DGV.Rows.Add(UCDMILBackQ, "L-Back", UCDMILBackY, UCDMILBackX, UCDMILBackZ, UCDMILBackM)
                        DGV.Rows.Add(UCDMISBackQ, "S-Back", UCDMISBackY, UCDMISBackX, UCDMISBackZ, UCDMISBackM)
                        '# TOP LIGHT CHECK #
                        If (CutlistForm.TLightCheck.Checked) Then
                            DGV.Rows.Add(UCDMITopBtmQ, "Top Light", UCDMITopBtmY, UCDMITopBtmX, UCDMITopBtmZ, UCDMITopBtmM, UCDMITBEdgeSeq, UCDMITBEdgeCode)
                            DGV.Rows.Add(UCDMITopBtmQ, "Bottom", UCDMITopBtmY, UCDMITopBtmX, UCDMITopBtmZ, UCDMITopBtmM, UCDMITBEdgeSeq, UCDMITBEdgeCode)
                        Else
                            DGV.Rows.Add(UCDMITopBtmQ, "TopBtm", UCDMITopBtmY, UCDMITopBtmX, UCDMITopBtmZ, UCDMITopBtmM, UCDMITBEdgeSeq, UCDMITBEdgeCode)
                        End If
                        '# TOP LIGHT CHECK #
                        If (CutlistForm.GlassCheck.Checked) Then
                            DGV.Rows.Add(UCDMIAdjShelfQ, "Adj.Sh", "", "", "", "Glass", "", "", "", "", "Ordered Joe G")
                        Else
                            '# MDF SPECIES CHECK #
                            If (CKWOPLANNERSSpeciesBox = "MDF") Then
                                DGV.Rows.Add(UCDMIAdjShelfQ, "Adj.Sh", UCDMIAdjShelfY, UCDMIAdjShelfX, UCDMIAdjShelfZ, UCDMIAdjShelfM)
                            Else
                                DGV.Rows.Add(UCDMIAdjShelfQ, "Adj.Sh", UCDMIAdjShelfY, UCDMIAdjShelfX, UCDMIAdjShelfZ, UCDMIAdjShelfM, UCDMIASEdgeSeq, UCDMIASEdgeCode)
                            End If
                        End If
                        DGV.Rows.Add(UCDMIGableQ, "Gable", UCDMIGableY, UCDMIGableX, UCDMIGableZ, UCDMIGableM, UCDMIGEdgeSeq, UCDMIGEdgeCode, UCDMIGEdgeSeq2, UCDMIGEdgeCode2)
                        DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                    End If
                Else
                    If (UCDAdjShelfQ = 0) Then
                        DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                        DGV.Rows.Add(CabSize1)
                        DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                        DGV.Rows.Add(UCDBackQ, "Back", UCDBackY, UCDBackX, UCDBackZ, UCDBackM)
                        '# TOP LIGHT CHECK #
                        If (CutlistForm.TLightCheck.Checked) Then
                            DGV.Rows.Add(UCDTopBtmQ, "Top Light", UCDTopBtmY, UCDTopBtmX, UCDTopBtmZ, UCDTopBtmM, UCDTBEdgeSeq, UCDASEdgeCode)
                            DGV.Rows.Add(UCDTopBtmQ, "Bottom", UCDTopBtmY, UCDTopBtmX, UCDTopBtmZ, UCDTopBtmM, UCDTBEdgeSeq, UCDASEdgeCode)
                        Else
                            DGV.Rows.Add(UCDTopBtmQ, "TopBtm", UCDTopBtmY, UCDTopBtmX, UCDTopBtmZ, UCDTopBtmM, UCDTBEdgeSeq, UCDASEdgeCode)
                        End If
                        DGV.Rows.Add(UCDGableQ, "Gable", UCDGableY, UCDGableX, UCDGableZ, UCDGableM, UCDGEdgeSeq, UCDGEdgeCode, UCDGEdgeSeq2, UCDGEdgeCode2)
                        DGV.Rows.Add(UCDBackStrapQ, "Back Strap", UCDBackStrapY, UCDBackStrapX, UCDBackStrapZ, UCDBackStrapM)
                        DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                    Else
                        DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                        DGV.Rows.Add(CabSize1)
                        DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                        DGV.Rows.Add(UCDBackQ, "Back", UCDBackY, UCDBackX, UCDBackZ, UCDBackM)
                        '# TOP LIGHT CHECK #
                        If (CutlistForm.TLightCheck.Checked) Then
                            DGV.Rows.Add(UCDTopBtmQ, "Top Light", UCDTopBtmY, UCDTopBtmX, UCDTopBtmZ, UCDTopBtmM, UCDTBEdgeSeq, UCDASEdgeCode)
                            DGV.Rows.Add(UCDTopBtmQ, "Bottom", UCDTopBtmY, UCDTopBtmX, UCDTopBtmZ, UCDTopBtmM, UCDTBEdgeSeq, UCDASEdgeCode)
                        Else
                            DGV.Rows.Add(UCDTopBtmQ, "TopBtm", UCDTopBtmY, UCDTopBtmX, UCDTopBtmZ, UCDTopBtmM, UCDTBEdgeSeq, UCDASEdgeCode)
                        End If
                        If (CutlistForm.GlassCheck.Checked) Then
                            DGV.Rows.Add(UCDAdjShelfQ, "Adj.Sh", "", "", "", "Glass", "", "", "", "", "Ordered Joe G")
                        Else
                            '# MDF SPECIES CHECK #
                            If (CKWOPLANNERSSpeciesBox = "MDF") Then
                                DGV.Rows.Add(UCDAdjShelfQ, "Adj.Sh", UCDAdjShelfY, UCDAdjShelfX, UCDAdjShelfZ, UCDAdjShelfM)
                            Else
                                DGV.Rows.Add(UCDAdjShelfQ, "Adj.Sh", UCDAdjShelfY, UCDAdjShelfX, UCDAdjShelfZ, UCDAdjShelfM, UCDASEdgeSeq, UCDASEdgeCode)
                            End If
                        End If
                        DGV.Rows.Add(UCDGableQ, "Gable", UCDGableY, UCDGableX, UCDGableZ, UCDGableM, UCDGEdgeSeq, UCDGEdgeCode, UCDGEdgeSeq2, UCDGEdgeCode2)
                        DGV.Rows.Add(UCDBackStrapQ, "Back Strap", UCDBackStrapY, UCDBackStrapX, UCDBackStrapZ, UCDBackStrapM)
                        DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                    End If
                End If

            '############################
            '# UPPER END SHELF # 8 #
            '############################
            '
            Case "UPPER END SHELF"
                CabCode = "UES-" & CutlistForm.FShType
                DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                DGV.Rows.Add(CabSize1)
                DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                DGV.Rows.Add(UESTopQ, "Top", UESTopY, UESTopX, UESTopZ, UESTopM, UESTEdgeSeq, UESTEdgeCode)
                DGV.Rows.Add(UESFixedShelfQ, "F.Sh", UESFixedShelfY, UESFixedShelfX, UESFixedShelfZ, UESFixedShelfM, UESFSHEdgeSeq, UESFSHEdgeCode)
                DGV.Rows.Add(UESLGableQ, "L-Gable", UESLGableY, UESLGableX, UESLGableZ, UESLGableM, UESLGEdgeSeq, UESLGEdgeCode, UESLGEdgeSeq2, UESLGEdgeCode2)
                DGV.Rows.Add(UESSGableQ, "S-Gable", UESSGableY, UESSGableX, UESSGableZ, UESSGableM, UESSGEdgeSeq, UESSGEdgeCode, UESSGEdgeSeq2, UESSGEdgeCode2)
                DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")

                '##################
                '# BASE # 9 LINES #
                '##################
                '
            Case "BASE"
                Dim BaseCase = CutlistForm.BaseCabCodeBox.Text
                Dim BaseMICase = CutlistForm.BaseCabCodeBox.Text
                '# MATCHING INTERIOR CHECK #
                If (CutlistForm.MatchingInteriorCheck.Checked) Then
                    Select Case BaseMICase
                        Case "B"
                            CabCode = CutlistForm.BaseCabCodeBox.Text & CutlistForm.OptionsVar
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(BMIBackQ, "Back", BMIBackY, BMIBackX, BMIBackZ, BMIBackM)
                            DGV.Rows.Add(BMITopStrapQ, "Top Strap", BMITopStrapY, BMITopStrapX, BMITopStrapZ, BMITopStrapM, BMITSEdgeSeq, BMITSEdgeCode)
                            DGV.Rows.Add(BMIBtmQ, "Bottom", BMIBtmY, BMIBtmX, BMIBtmZ, BMIBtmM, BMIBEdgeSeq, BMIBEdgeCode)
                            If (CutlistForm.GlassCheck.Checked = True) Then
                            Else
                                If (CKWOPLANNERSSpeciesBox = "MDF") Then
                                    DGV.Rows.Add(BMIAdjShelfQ, "Adj.Sh", BMIAdjShelfY, BMIAdjShelfX, BMIAdjShelfZ, BMIAdjShelfM)
                                Else
                                    DGV.Rows.Add(BMIAdjShelfQ, "Adj.Sh", BMIAdjShelfY, BMIAdjShelfX, BMIAdjShelfZ, BMIAdjShelfM, BMIASEdgeSeq, BMIASEdgeCode)
                                End If
                            End If
                            DGV.Rows.Add(BMIGableQ, "Gable", BMIGableY, BMIGableX, BMIGableZ, BMIGableM, BMIGEdgeSeq, BMIGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")

                        Case "B1D", "B2D", "B3D", "B4D"
                            CabCode = CutlistForm.BaseCabCodeBox.Text & CutlistForm.OptionsVar & "-" & Hardware
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(BMIBackQ, "Back", BMIBackY, BMIBackX, BMIBackZ, BMIBackM)
                            DGV.Rows.Add(BMITopStrapQ, "Top Strap", BMITopStrapY, BMITopStrapX, BMITopStrapZ, BMITopStrapM, BMITSEdgeSeq, BMITSEdgeCode)
                            DGV.Rows.Add(BMIBtmQ, "Bottom", BMIBtmY, BMIBtmX, BMIBtmZ, BMIBtmM, BMIBEdgeSeq, BMIBEdgeCode)
                            DGV.Rows.Add(BMIGableQ, "Gable", BMIGableY, BMIGableX, BMIGableZ, BMIGableM, BMIGEdgeSeq, BMIGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")

                        Case "BPO"
                            CabCode = CutlistForm.BaseCabCodeBox.Text & CutlistForm.OptionsVar
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(BMIBackQ, "Back", BMIBackY, BMIBackX, BMIBackZ, BMIBackM)
                            DGV.Rows.Add(BMITopStrapQ, "Top Strap", BMITopStrapY, BMITopStrapX, BMITopStrapZ, BMITopStrapM, BMITSEdgeSeq, BMITSEdgeCode)
                            DGV.Rows.Add(BMIBtmQ, "Bottom", BMIBtmY, BMIBtmX, BMIBtmZ, BMIBtmM, BMIBEdgeSeq, BMIBEdgeCode)
                            DGV.Rows.Add(BMIGableQ, "Gable", BMIGableY, BMIGableX, BMIGableZ, BMIGableM, BMIGEdgeSeq, BMIGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")

                        Case "BS"
                            CabCode = CutlistForm.BaseCabCodeBox.Text & CutlistForm.OptionsVar
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(BMIBackQ, "Back", BMIBackY, BMIBackX, BMIBackZ, BMIBackM)
                            DGV.Rows.Add(BMITopStrapQ, "Top Strap", BMITopStrapY, BMITopStrapX, BMITopStrapZ, BMITopStrapM, BMITSEdgeSeq, BMITSEdgeCode)
                            DGV.Rows.Add(BMIBtmQ, "Bottom", BMIBtmY, BMIBtmX, BMIBtmZ, BMIBtmM, BMIBEdgeSeq, BMIBEdgeCode)
                            DGV.Rows.Add(BMIStrapQ, "Sink Strap", BMIStrapY, BMIStrapX, BMIStrapZ, BMIStrapM)
                            DGV.Rows.Add(BMIGableQ, "Gable", BMIGableY, BMIGableX, BMIGableZ, BMIGableM, BMIGEdgeSeq, BMIGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")

                        Case "BSFF"
                            CabCode = CutlistForm.BaseCabCodeBox.Text & CutlistForm.OptionsVar
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(BMIBackQ, "Back", BMIBackY, BMIBackX, BMIBackZ, BMIBackM)
                            DGV.Rows.Add(BMITopStrapQ, "Top Strap", BMITopStrapY, BMITopStrapX, BMITopStrapZ, BMITopStrapM, BMITSEdgeSeq, BMITSEdgeCode)
                            DGV.Rows.Add(BMIBtmQ, "Bottom", BMIBtmY, BMIBtmX, BMIBtmZ, BMIBtmM, BMIBEdgeSeq, BMIBEdgeCode)
                            DGV.Rows.Add(BMIStrapQ, "Strap", BMIStrapY, BMIStrapX, BMIStrapZ, BMIStrapM, BMISEdgeSeq, BMISEdgeCode)
                            DGV.Rows.Add(BMIGableQ, "Gable", BMIGableY, BMIGableX, BMIGableZ, BMIGableM, BMIGEdgeSeq, BMIGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")

                        Case "BSFM"
                            CabCode = CutlistForm.BaseCabCodeBox.Text & CutlistForm.OptionsVar
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(BMIBackQ, "Back", BMIBackY, BMIBackX, BBackZ, BBackM)
                            DGV.Rows.Add(BMITopQ, "Top", BMITopY, BMITopX, BMITopZ, "PLY", BRTEdgeSeq, BRTEdgeCode)
                            DGV.Rows.Add(BMIBtmQ, "Bottom", BMIBtmY, BMIBtmX, BMIBtmZ, BMIBtmM, BMIBEdgeSeq, BMIBEdgeCode)
                            DGV.Rows.Add(BMIGableQ, "Gable", BMIGableY, BMIGableX, BMIGableZ, BMIGableM, BMIGEdgeSeq, BMIGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")

                        Case "BSFM2D"
                            CabCode = CutlistForm.BaseCabCodeBox.Text & CutlistForm.OptionsVar & "-" & Hardware
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(BBackQ, "Back", BBackY, BBackX, BBackZ, BBackM)
                            DGV.Rows.Add(BTopQ, "Top", BTopY, BTopX, BTopZ, "PLY", BRTEdgeSeq, BRTEdgeCode)
                            DGV.Rows.Add(BBtmQ, "Bottom", BBtmY, BBtmX, BBtmZ, BBtmM, BBEdgeSeq, BBEdgeCode)
                            DGV.Rows.Add(BGableQ, "Gable", BGableY, BGableX, BGableZ, BGableM, BGEdgeSeq, BGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")

                        Case "BTD"
                            CabCode = CutlistForm.BaseCabCodeBox.Text & CutlistForm.OptionsVar & "-" & Hardware
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(BMIBackQ, "Back", BMIBackY, BMIBackX, BMIBackZ, BMIBackM)
                            DGV.Rows.Add(BMITopStrapQ, "Top Strap", BMITopStrapY, BMITopStrapX, BMITopStrapZ, BMITopStrapM, BMITSEdgeSeq, BMITSEdgeCode)
                            DGV.Rows.Add(BMIBtmQ, "Bottom", BMIBtmY, BMIBtmX, BMIBtmZ, BMIBtmM, BMIBEdgeSeq, BMIBEdgeCode)
                            DGV.Rows.Add(BMIStrapQ, "Strap", BMIStrapY, BMIStrapX, BMIStrapZ, BMIStrapM, BMISEdgeSeq, BMISEdgeCode)
                            DGV.Rows.Add(BMIGableQ, "Gable", BMIGableY, BMIGableX, BMIGableZ, BMIGableM, BMIGEdgeSeq, BMIGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")

                        Case "BTDD"
                            CabCode = CutlistForm.BaseCabCodeBox.Text & CutlistForm.OptionsVar & "-" & Hardware
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(BMIBackQ, "Back", BMIBackY, BMIBackX, BMIBackZ, BMIBackM)
                            DGV.Rows.Add(BMITopQ, "Top", BMITopY, BMITopX, BMITopZ, BMITopM, BMITEdgeSeq, BMITEdgeCode)
                            DGV.Rows.Add(BMIBtmQ, "Bottom", BMIBtmY, BMIBtmX, BMIBtmZ, BMIBtmM, BMIBEdgeSeq, BMIBEdgeCode)
                            DGV.Rows.Add(BMIStrapQ, "Strap", BMIStrapY, BMIStrapX, BMIStrapZ, BMIStrapM, BMISEdgeSeq, BMISEdgeCode)
                            DGV.Rows.Add(BMIDividerQ, "Divider", BMIDividerY, BMIDividerX, BMIDividerZ, BMIDividerM, BMIDEdgeSeq, BMIDEdgeCode)
                            DGV.Rows.Add(BMIGableQ, "Gable", BMIGableY, BMIGableX, BMIGableZ, BMIGableM, BMIGEdgeSeq, BMIGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")

                        Case "BWRL"
                            CabCode = CutlistForm.BaseCabCodeBox.Text & CutlistForm.OptionsVar
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(BWRBackQ, "Back", BWRBackY, BWRBackX, BWRBackZ, BWRBackM)
                            DGV.Rows.Add(BWRTopBtmQ, "TopBtm", BWRTopBtmY, BWRTopBtmX, BWRTopBtmZ, BWRTopBtmM, BWRTBEdgeSeq, BWRTBEdgeCode)
                            DGV.Rows.Add(BWRGableQ, "Gable", BWRGableY, BWRGableX, BWRGableZ, BWRGableM, BWRGEdgeSeq, BWRGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")

                        Case "BWRLH"
                            CabCode = CutlistForm.BaseCabCodeBox.Text & CutlistForm.OptionsVar
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(BWRBackQ, "Back", BWRBackY, BWRBackX, BWRBackZ, BWRBackM)
                            DGV.Rows.Add(BWRTopBtmQ, "TopBtm", BWRTopBtmY, BWRTopBtmX, BWRTopBtmZ, BWRTopBtmM, BWRTBEdgeSeq, BWRTBEdgeCode)
                            DGV.Rows.Add(BWRFFSHQ, "F.F.Sh", BWRFFSHY, BWRFFSHX, BWRFFSHZ, BWRFFSHM, BWRFFSEdgeSeq, BWRFFSEdgeCode)
                            DGV.Rows.Add(BWRGableQ, "Gable", BWRGableY, BWRGableX, BWRGableZ, BWRGableM, BWRGEdgeSeq, BWRGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")

                        Case "BWRLHV"
                            CabCode = CutlistForm.BaseCabCodeBox.Text & CutlistForm.OptionsVar
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(BWRBackQ, "Back", BWRBackY, BWRBackX, BWRBackZ, BWRBackM)
                            DGV.Rows.Add(BWRTopBtmQ, "TopBtm", BWRTopBtmY, BWRTopBtmX, BWRTopBtmZ, BWRTopBtmM, BWRTBEdgeSeq, BWRTBEdgeCode)
                            DGV.Rows.Add(BWRFFSHQ, "F.F.Sh", BWRFFSHY, BWRFFSHX, BWRFFSHZ, BWRFFSHM, BWRFFSEdgeSeq, BWRFFSEdgeCode)
                            DGV.Rows.Add(BWRGableQ, "Gable", BWRGableY, BWRGableX, BWRGableZ, BWRGableM, BWRGEdgeSeq, BWRGEdgeCode)
                            DGV.Rows.Add(BWRValanceQ, "Valance", BWRValanceY, BWRValanceX, BWRValanceZ, BWRValanceM, "", "", "", "", BWRVN)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                    End Select
                Else
                    Select Case BaseCase
                        Case "B"
                            CabCode = CutlistForm.BaseCabCodeBox.Text & CutlistForm.OptionsVar
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(BBackQ, "Back", BBackY, BBackX, BBackZ, BBackM)
                            DGV.Rows.Add(BTopStrapQ, "Top Strap", BTopStrapY, BTopStrapX, BTopStrapZ, BTopStrapM, BTSEdgeSeq, BTSEdgeCode)
                            DGV.Rows.Add(BBtmQ, "Bottom", BBtmY, BBtmX, BBtmZ, BBtmM, BBEdgeSeq, BBEdgeCode)
                            If (CutlistForm.GlassCheck.Checked = True) Then
                                DGV.Rows.Add(BAdjShelfQ, "Adj.Sh", "", "", "", "Glass", "", "", "", "", "Ordered Joe G")
                            Else
                                DGV.Rows.Add(BAdjShelfQ, "Adj.Sh", BAdjShelfY, BAdjShelfX, BAdjShelfZ, BAdjShelfM, BASEdgeSeq, BASEdgeCode)
                            End If
                            DGV.Rows.Add(BGableQ, "Gable", BGableY, BGableX, BGableZ, BGableM, BGEdgeSeq, BGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")

                        Case "BR"
                            CabCode = CutlistForm.BaseCabCodeBox.Text & CutlistForm.OptionsVar
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(BBackQ, "Back", BBackY, BBackX, BBackZ, BBackM)
                            DGV.Rows.Add(BTopQ, "Top", BTopY, BTopX, BTopZ, "PLY", BRTEdgeSeq, BRTEdgeCode)
                            DGV.Rows.Add(BBtmQ, "Bottom", BBtmY, BBtmX, BBtmZ, BBtmM, BBEdgeSeq, BBEdgeCode)
                            DGV.Rows.Add(BGableQ, "Gable", BGableY, BGableX, BGableZ, BGableM, BGEdgeSeq, BGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")

                        Case "BR2D"
                            CabCode = CutlistForm.BaseCabCodeBox.Text & CutlistForm.OptionsVar & "-" & Hardware
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(BBackQ, "Back", BBackY, BBackX, BBackZ, BBackM)
                            DGV.Rows.Add(BTopQ, "Top", BTopY, BTopX, BTopZ, "PLY", BRTEdgeSeq, BRTEdgeCode)
                            DGV.Rows.Add(BBtmQ, "Bottom", BBtmY, BBtmX, BBtmZ, BBtmM, BBEdgeSeq, BBEdgeCode)
                            DGV.Rows.Add(BGableQ, "Gable", BGableY, BGableX, BGableZ, BGableM, BGEdgeSeq, BGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")

                        Case "BSFF"
                            CabCode = CutlistForm.BaseCabCodeBox.Text & CutlistForm.OptionsVar
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(BBackQ, "Back", BBackY, BBackX, BBackZ, BBackM)
                            DGV.Rows.Add(BTopStrapQ, "Top Strap", BTopStrapY, BTopStrapX, BTopStrapZ, BTopStrapM, BTSEdgeSeq, BTSEdgeCode)
                            DGV.Rows.Add(BBtmQ, "Bottom", BBtmY, BBtmX, BBtmZ, BBtmM, BBEdgeSeq, BBEdgeCode)
                            DGV.Rows.Add(BStrapQ, "Strap", BStrapY, BStrapX, BStrapZ, BStrapM)
                            DGV.Rows.Add(BGableQ, "Gable", BGableY, BGableX, BGableZ, BGableM, BGEdgeSeq, BGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")

                        Case "BSFM"
                            CabCode = CutlistForm.BaseCabCodeBox.Text & CutlistForm.OptionsVar
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(BBackQ, "Back", BBackY, BBackX, BBackZ, BBackM)
                            DGV.Rows.Add(BTopQ, "Top", BTopY, BTopX, BTopZ, "PLY", BRTEdgeSeq, BRTEdgeCode)
                            DGV.Rows.Add(BBtmQ, "Bottom", BBtmY, BBtmX, BBtmZ, BBtmM, BBEdgeSeq, BBEdgeCode)
                            DGV.Rows.Add(BGableQ, "Gable", BGableY, BGableX, BGableZ, BGableM, BGEdgeSeq, BGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")

                        Case "BSFM2D"
                            CabCode = CutlistForm.BaseCabCodeBox.Text & CutlistForm.OptionsVar & "-" & Hardware
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(BBackQ, "Back", BBackY, BBackX, BBackZ, BBackM)
                            DGV.Rows.Add(BTopQ, "Top", BTopY, BTopX, BTopZ, "PLY", BRTEdgeSeq, BRTEdgeCode)
                            DGV.Rows.Add(BBtmQ, "Bottom", BBtmY, BBtmX, BBtmZ, BBtmM, BBEdgeSeq, BBEdgeCode)
                            DGV.Rows.Add(BGableQ, "Gable", BGableY, BGableX, BGableZ, BGableM, BGEdgeSeq, BGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")

                        Case "BS"
                            CabCode = CutlistForm.BaseCabCodeBox.Text & CutlistForm.OptionsVar
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(BBackQ, "Back", BBackY, BBackX, BBackZ, BBackM)
                            DGV.Rows.Add(BTopStrapQ, "Top Strap", BTopStrapY, BTopStrapX, BTopStrapZ, BTopStrapM, BTSEdgeSeq, BTSEdgeCode)
                            DGV.Rows.Add(BBtmQ, "Bottom", BBtmY, BBtmX, BBtmZ, BBtmM, BBEdgeSeq, BBEdgeCode)
                            DGV.Rows.Add(BStrapQ, "Sink Strap", BStrapY, BStrapX, BStrapZ, BStrapM)
                            DGV.Rows.Add(BGableQ, "Gable", BGableY, BGableX, BGableZ, BGableM, BGEdgeSeq, BGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")

                        Case "B1D", "B2D", "B3D", "B4D"
                            CabCode = CutlistForm.BaseCabCodeBox.Text & CutlistForm.OptionsVar & "-" & Hardware
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(BBackQ, "Back", BBackY, BBackX, BBackZ, BBackM)
                            DGV.Rows.Add(BTopStrapQ, "Top Strap", BTopStrapY, BTopStrapX, BTopStrapZ, BTopStrapM, BTSEdgeSeq, BTSEdgeCode)
                            DGV.Rows.Add(BBtmQ, "Bottom", BBtmY, BBtmX, BBtmZ, BBtmM, BBackEdgeSeq, BBackEdgeCode)
                            DGV.Rows.Add(BGableQ, "Gable", BGableY, BGableX, BGableZ, BGableM, BGEdgeSeq, BGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")

                        Case "BTD"
                            CabCode = CutlistForm.BaseCabCodeBox.Text & CutlistForm.OptionsVar & "-" & Hardware
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(BBackQ, "Back", BBackY, BBackX, BBackZ, BBackM)
                            DGV.Rows.Add(BTopStrapQ, "Top Strap", BTopStrapY, BTopStrapX, BTopStrapZ, BTopStrapM, BTSEdgeSeq, BTSEdgeCode)
                            DGV.Rows.Add(BBtmQ, "Bottom", BBtmY, BBtmX, BBtmZ, BBtmM, BBEdgeSeq, BBEdgeCode)
                            DGV.Rows.Add(BStrapQ, "Strap", BStrapY, BStrapX, BStrapZ, BStrapM, BSEdgeSeq, BSEdgeCode)
                            DGV.Rows.Add(BGableQ, "Gable", BGableY, BGableX, BGableZ, BGableM, BGEdgeSeq, BGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")

                        Case "BTDD"
                            CabCode = CutlistForm.BaseCabCodeBox.Text & CutlistForm.OptionsVar & "-" & Hardware
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(BBackQ, "Back", BBackY, BBackX, BBackZ, BBackM)
                            DGV.Rows.Add(BTopQ, "Top", BTopY, BTopX, BTopZ, BTopM, BTEdgeSeq, BTEdgeCode)
                            DGV.Rows.Add(BBtmQ, "Bottom", BBtmY, BBtmX, BBtmZ, BBtmM, BBEdgeSeq, BBEdgeCode)
                            DGV.Rows.Add(BStrapQ, "Strap", BStrapY, BStrapX, BStrapZ, BStrapM, BSEdgeSeq, BSEdgeCode)
                            DGV.Rows.Add(BDividerQ, "Divider", BDividerY, BDividerX, BDividerZ, BDividerM, BDEdgeSeq, BDEdgeCode)
                            DGV.Rows.Add(BGableQ, "Gable", BGableY, BGableX, BGableZ, BGableM, BGEdgeSeq, BGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                    End Select
                End If

            Case "BASE END ANGLED DOOR"

                If (CutlistForm.OrientationBox.Text = "LEFT") Then
                    CabCode = "BEAD-L" & CutlistForm.OptionsVar
                    DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                    DGV.Rows.Add(CabSize1)
                    DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                    DGV.Rows.Add(BBackQ, "Back", BBackY, BBackX, BBackZ, BBackM)
                    DGV.Rows.Add(BBtmQ, "Top", BBtmY, BBtmX, BBtmZ, BBtmM, "E1M", BTSEdgeCode)
                    DGV.Rows.Add(BBtmQ, "Bottom", BBtmY, BBtmX, BBtmZ, BBtmM, "E1M", BBEdgeCode)
                    If (CutlistForm.GlassCheck.Checked = True) Then
                        DGV.Rows.Add(BEADAdjShelfQ, "Adj.Sh", "", "", "", "Glass", "", "", "", "", "Ordered Joe G")
                    Else
                        DGV.Rows.Add(BEADAdjShelfQ, "Adj.Sh", BEADAdjShelfY, BEADAdjShelfX, BEADAdjShelfZ, BAdjShelfM, "E1M", BASEdgeCode)
                    End If
                    DGV.Rows.Add(BEADLGableQ, "L-GableR", BGableY, BGableX, BGableZ, BGableM, "E1M", BGEdgeCode)
                    DGV.Rows.Add(BEADSGableQ, "S-GableL", BEADSGableY, BEADSGableX, BEADSGableZ, BGableM, "E1M", BGEdgeCode)
                ElseIf (CutlistForm.OrientationBox.Text = "RIGHT") Then
                    CabCode = "BEAD-R" & CutlistForm.OptionsVar
                    DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                    DGV.Rows.Add(CabSize1)
                    DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                    DGV.Rows.Add(BBackQ, "Back", BBackY, BBackX, BBackZ, BBackM)
                    DGV.Rows.Add(BBtmQ, "Top", BBtmY, BBtmX, BBtmZ, BBtmM, "E1M", BTSEdgeCode)
                    DGV.Rows.Add(BBtmQ, "Bottom", BBtmY, BBtmX, BBtmZ, BBtmM, "E1M", BBEdgeCode)
                    If (CutlistForm.GlassCheck.Checked = True) Then
                        DGV.Rows.Add(BEADAdjShelfQ, "Adj.Sh", "", "", "", "Glass", "", "", "", "", "Ordered Joe G")
                    Else
                        DGV.Rows.Add(BEADAdjShelfQ, "Adj.Sh", BEADAdjShelfY, BEADAdjShelfX, BEADAdjShelfZ, BAdjShelfM, "E1M", BASEdgeCode)
                    End If
                    DGV.Rows.Add(BEADLGableQ, "L-GableL", BGableY, BGableX, BGableZ, BGableM, "E1M", BGEdgeCode)
                    DGV.Rows.Add(BEADSGableQ, "S-GableR", BEADSGableY, BEADSGableX, BEADSGableZ, BGableM, "E1M", BGEdgeCode)
                End If
                DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")

            Case "BASE PENINSULA OPEN"
                Dim BPOPCase2 = CutlistForm.OrientationBox.Text
                Select Case BPOPCase2
                    Case "LEFT"
                        CabCode = "BPOP-L"
                        DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                        DGV.Rows.Add(CabSize1)
                        DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                        DGV.Rows.Add(BPOPTopBtmFSHQ, "TopBtm + F.SH", BPOPTopBtmFSHY, BPOPTopBtmFSHX, BPOPTopBtmFSHZ, BPOPTopBtmFSHM, BPOPTBEdgeSeq, BPOPTBEdgeCode)
                        DGV.Rows.Add(BPOPSGableQ, "S-Gable-L", BPOPSGableY, BPOPSGableX, BPOPSGableZ, BPOPSGableM, BPOPSGEdgeSeq, BPOPSGEdgeCode)
                        DGV.Rows.Add(BPOPLGableQ, "L-Gable-R", BPOPLGableY, BPOPLGableX, BPOPLGableZ, BPOPLGableM, BPOPLGEdgeSeq, BPOPLGEdgeCode)
                        DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                    Case "RIGHT"
                        CabCode = "BPOP-R"
                        DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                        DGV.Rows.Add(CabSize1)
                        DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                        DGV.Rows.Add(BPOPTopBtmFSHQ, "TopBtm + F.SH", BPOPTopBtmFSHY, BPOPTopBtmFSHX, BPOPTopBtmFSHZ, BPOPTopBtmFSHM, BPOPTBEdgeSeq, BPOPTBEdgeCode)
                        DGV.Rows.Add(BPOPSGableQ, "S-Gable-R", BPOPSGableY, BPOPSGableX, BPOPSGableZ, BPOPSGableM, BPOPSGEdgeSeq, BPOPSGEdgeCode)
                        DGV.Rows.Add(BPOPLGableQ, "L-Gable-L", BPOPLGableY, BPOPLGableX, BPOPLGableZ, BPOPLGableM, BPOPLGEdgeSeq, BPOPLGEdgeCode)
                        DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                End Select

            '#################################
            '# OVEN PANEL 1 CUTOUT # 4 LINES #
            '#################################
            '
            Case "OVEN PANEL 1 CUTOUT"
                CabCode = "OVP1"
                If (CKWOPLANNERSSpeciesBox = "PVC") Then
                    DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                    DGV.Rows.Add(CabSize1)
                    DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                    DGV.Rows.Add(OV1Q, "Panel", OV1Y, OV1X, OV1Z, OV1M, "", "", "", "", OV1N)
                    DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                Else
                    DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                    DGV.Rows.Add(CabSize1)
                    DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                    DGV.Rows.Add(OV1Q, "Panel", OV1Y, OV1X, OV1Z, OV1M, OVPLEdgeSeq, OVPLEdgeCode)
                    DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                End If

            '#################################
            '# OVEN PANEL 2 CUTOUT # 4 LINES #
            '#################################
            '
            Case "OVEN PANEL 2 CUTOUT"
                CabCode = "OVP2"
                If (CKWOPLANNERSSpeciesBox = "PVC") Then
                    DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                    DGV.Rows.Add(CabSize1)
                    DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                    DGV.Rows.Add(OV1Q, "Panel", OV1Y, OV1X, OV1Z, OV1M, "", "", "", "", OV1N)
                    DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                Else
                    DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                    DGV.Rows.Add(CabSize1)
                    DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                    DGV.Rows.Add(OV1Q, "Panel", OV1Y, OV1X, OV1Z, OV1M, OVPLEdgeSeq, OVPLEdgeCode)
                    DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                End If

            '#################################
            '# OVEN PANEL 3 CUTOUT # 4 LINES #
            '#################################
            '
            Case "OVEN PANEL 3 CUTOUT"
                CabCode = "OVP3"
                If (CKWOPLANNERSSpeciesBox = "PVC") Then
                    DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                    DGV.Rows.Add(CabSize1)
                    DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                    DGV.Rows.Add(OV1Q, "Panel", OV1Y, OV1X, OV1Z, OV1M, "", "", "", "", OV1N)
                    DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                Else
                    DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                    DGV.Rows.Add(CabSize1)
                    DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                    DGV.Rows.Add(OV1Q, "Panel", OV1Y, OV1X, OV1Z, OV1M, OVPLEdgeSeq, OVPLEdgeCode)
                    DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                End If

            '##################################
            '# TALL UTILITY 1 UNIT # 11 LINES #
            '##################################
            '
            Case "TALL UTILITY 1 UNIT"
                If (CutlistForm.MatchingInteriorCheck.Checked) Then
                    CabCode = "TUMI"
                    DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                    DGV.Rows.Add(CabSize1)
                    DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                    DGV.Rows.Add(TU1UBackQ, "Back", TU1UBackY, TU1UBackX, TU1UBackZ, TU1UBackM)
                    DGV.Rows.Add(TU1UTopQ, "Top", TU1UTopY, TU1UTopX, TU1UTopZ, TU1UTopM, TU1UTEdgeSeq, TU1UTEdgeCode)
                    DGV.Rows.Add(TU1UBotQ, "Bottom", TU1UBotY, TU1UBotX, TU1UBotZ, TU1UBotM, TU1UBEdgeSeq, TU1UBEdgeCode)
                    DGV.Rows.Add(TU1UFFSQ, "F.F.SH", TU1UFFSY, TU1UFFSX, TU1UFFSZ, TU1UFFSM, TU1UFFSEdgeSeq, TU1UFFSEdgeCode)
                    DGV.Rows.Add(TU1USFSQ, "S.F.Sh", TU1USFSY, TU1USFSX, TU1USFSZ, TU1USFSM, TU1USFSEdgeSeq, TU1USFSEdgeCode)
                    DGV.Rows.Add(TU1UStrapQ, "Strap", TU1UStrapY, TU1UStrapX, TU1UStrapZ, TU1UStrapM, TU1USEdgeSeq, TU1USEdgeCode)
                    DGV.Rows.Add(TU1UAdjShelfQ, "Adj.Sh", TU1UAdjShelfY, TU1UAdjShelfX, TU1UAdjShelfZ, TU1UAdjShelfM, TU1UASEdgeSeq, TU1UASEdgeCode)
                    DGV.Rows.Add(TU1UGableQ, "Gable", TU1UGableY, TU1UGableX, TU1UGableZ, TU1UGableM, TU1UGEdgeSeq, TU1UGEdgeCode)
                    DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                Else
                    CabCode = "TU"
                    DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                    DGV.Rows.Add(CabSize1)
                    DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                    DGV.Rows.Add(TU1UBackQ, "Back", TU1UBackY, TU1UBackX, TU1UBackZ, TU1UBackM)
                    DGV.Rows.Add(TU1UTopQ, "Top", TU1UTopY, TU1UTopX, TU1UTopZ, TU1UTopM, TU1UTEdgeSeq, TU1UTEdgeCode)
                    DGV.Rows.Add(TU1UBotQ, "Bottom", TU1UBotY, TU1UBotX, TU1UBotZ, TU1UBotM, TU1UBEdgeSeq, TU1UBEdgeCode)
                    DGV.Rows.Add(TU1UFFSQ, "F.F.SH", TU1UFFSY, TU1UFFSX, TU1UFFSZ, TU1UFFSM, TU1UFFSEdgeSeq, TU1UFFSEdgeCode)
                    DGV.Rows.Add(TU1USFSQ, "S.F.Sh", TU1USFSY, TU1USFSX, TU1USFSZ, TU1USFSM, TU1USFSEdgeSeq, TU1USFSEdgeCode)
                    DGV.Rows.Add(TU1UStrapQ, "Strap", TU1UStrapY, TU1UStrapX, TU1UStrapZ, TU1UStrapM, TU1USEdgeSeq, TU1USEdgeCode)
                    DGV.Rows.Add(TU1UAdjShelfQ, "Adj.Sh", TU1UAdjShelfY, TU1UAdjShelfX, TU1UAdjShelfZ, TU1UAdjShelfM, TU1UASEdgeSeq, TU1UASEdgeCode)
                    DGV.Rows.Add(TU1UGableQ, "Gable", TU1UGableY, TU1UGableX, TU1UGableZ, TU1UGableM, TU1UGEdgeSeq, TU1UGEdgeCode)
                    DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                End If

            Case "TALL UTILITY 2 UNIT"
                Dim TUCase5 = CutlistForm.TUCabBox.Text
                Dim BCBCase = CutlistForm.BaseCabCodeBox2.Text
                Select Case TUCase5
                    Case "TU"
                        '###############
                        '# GROUP RULES #
                        '###############
                        '
                        Dim CH1 = (TU2UUGableX + 1) / 10
                        Dim CH2 = (TU2UBGableX) / 10
                        Dim TH = CH1 + CH2
                        If (CutlistForm.MatchingInteriorCheck.Checked) Then
                            If (CKWOPLANNERSGroupBox = "GROUP1") Then
                                CabCode = CutlistForm.TUCabBox.Text
                                Dim CabCode2 As String = "U"
                                DGV.Rows.Add(CabCode)
                                DGV.Rows.Add(VarWidthI & "-" & TH & "-" & VarDepthI)
                                DGV.Rows.Add(CabCode2, "Qty", VarAmountI)
                                DGV.Rows.Add(VarWidthI & "-" & CH1 & "-" & VarDepthI)
                                DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                                DGV.Rows.Add(TU2UUMIBackQ, "U-Back", TU2UUBackY, TU2UUBackX, TU2UUBackZ, TU2UUBackM)
                                DGV.Rows.Add(TU2UUMITopQ, "U-Top", TU2UUTopY, TU2UUTopX, TU2UUTopZ, TU2UUTopM, TU2UUTEdgeSeq, TU2UUTEdgeCode)
                                DGV.Rows.Add(TU2UUMIBotQ, "U-Bottom", TU2UUBotY, TU2UUBotX, TU2UUBotZ, TU2UUBotM, TU2UUBEdgeSeq, TU2UUBEdgeCode)
                                DGV.Rows.Add(TU2UUMIAdjShelfQ, "U-Adj.Sh", TU2UUAdjShelfY, TU2UUAdjShelfX, TU2UUAdjShelfZ, TU2UUAdjShelfM, TU2UUASEdgeSeq, TU2UUASEdgeCode)
                                DGV.Rows.Add(TU2UUMIGableQ, "U-Gable", TU2UUGableY, TU2UUGableX, TU2UUGableZ, TU2UUGableM, TU2UUGEdgeSeq, TU2UUGEdgeCode, TU2UUGEdgeSeq2, TU2UUGEdgeCode2)
                                DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                                CabCode2 = "T"
                                DGV.Rows.Add(CabCode)
                                DGV.Rows.Add(VarWidthI & "-" & TH & "-" & VarDepthI)
                                DGV.Rows.Add(CabCode2, "Qty", VarAmountI)
                                DGV.Rows.Add(VarWidthI & "-" & CH2 & "-" & VarDepthI)
                                DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                                DGV.Rows.Add(TU2UBBackQ, "T-Back", TU2UBBackY, TU2UBBackX, TU2UBBackZ, TU2UBBackM)
                                DGV.Rows.Add(TU2UBTopQ, "T-Top", TU2UBTopY, TU2UBTopX, TU2UBTopZ, TU2UBTopM, TU2UBTEdgeSeq, TU2UBTEdgeCode)
                                DGV.Rows.Add(TU2UBBotQ, "T-Bottom", TU2UBBotY, TU2UBBotX, TU2UBBotZ, TU2UBBotM, TU2UBBEdgeSeq, TU2UBBEdgeCode)
                                DGV.Rows.Add(TU2UBSFSQ, "T-S.F.Sh", TU2UBSFSY, TU2UBSFSX, TU2UBSFSZ, TU2UBSFSM, TU2UBSFSEdgeSeq, TU2UBSFSEdgeCode)
                                DGV.Rows.Add(TU2UBAdjShelfQ, "T-Adj.Sh", TU2UBAdjShelfY, TU2UBAdjShelfX, TU2UBAdjShelfZ, TU2UBAdjShelfM, TU2UBASEdgeSeq, TU2UBASEdgeCode)
                                DGV.Rows.Add(TU2UBGableQ, "T-Gable", TU2UBGableY, TU2UBGableX, TU2UBGableZ, TU2UBGableM, TU2UBGEdgeSeq, TU2UBGEdgeCode)
                                DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                            End If

                            '##########
                            '# GROUP2 #
                            '##########
                            '
                            If (CKWOPLANNERSGroupBox = "GROUP2") Then
                                CabCode = CutlistForm.TUCabBox.Text
                                Dim CabCode2 As String = "U-"
                                DGV.Rows.Add(CabCode)
                                DGV.Rows.Add(VarWidthI & "-" & TH & "-" & VarDepthI)
                                DGV.Rows.Add(CabCode2, "Qty", VarAmountI)
                                DGV.Rows.Add(VarWidthI & "-" & CH1 & "-" & VarDepthI)
                                DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                                DGV.Rows.Add(TU2UUMIBackQ, "U-Back", TU2UUBackY, TU2UUBackX, TU2UUBackZ, TU2UUBackM)
                                DGV.Rows.Add(TU2UUMITopBtmQ, "U-TopBtm", TU2UUTopBtmY, TU2UUTopBtmX, TU2UUTopBtmZ, TU2UUTopBtmM, TU2UUTBEdgeSeq, TU2UUTBEdgeCode)
                                DGV.Rows.Add(TU2UUMIAdjShelfQ, "U-Adj.Sh", TU2UUAdjShelfY, TU2UUAdjShelfX, TU2UUAdjShelfZ, TU2UUAdjShelfM, TU2UUASEdgeSeq, TU2UUASEdgeCode)
                                DGV.Rows.Add(TU2UUMIGableQ, "U-Gable", TU2UUGableY, TU2UUGableX, TU2UUGableZ, TU2UUGableM, TU2UUGEdgeSeq, TU2UUGEdgeCode, TU2UUGEdgeSeq2, TU2UUGEdgeCode2)
                                DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                                CabCode2 = "T"
                                DGV.Rows.Add(CabCode)
                                DGV.Rows.Add(VarWidthI & "-" & TH & "-" & VarDepthI)
                                DGV.Rows.Add(CabCode2, "Qty", VarAmountI)
                                DGV.Rows.Add(VarWidthI & "-" & CH2 & "-" & VarDepthI)
                                DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                                DGV.Rows.Add(TU2UBBackQ, "T-Back", TU2UBBackY, TU2UBBackX, TU2UBBackZ, TU2UBBackM)
                                DGV.Rows.Add(TU2UBTopBtmQ, "T-TopBtm", TU2UBTopBtmY, TU2UBTopBtmX, TU2UBTopBtmZ, TU2UBTopBtmM, TU2UBTBEdgeSeq, TU2UBTBEdgeCode)
                                DGV.Rows.Add(TU2UBSFSQ, "T-S.F.Sh", TU2UBSFSY, TU2UBSFSX, TU2UBSFSZ, TU2UBSFSM, TU2UBSFSEdgeSeq, TU2UBSFSEdgeCode)
                                DGV.Rows.Add(TU2UBAdjShelfQ, "T-Adj.Sh", TU2UBAdjShelfY, TU2UBAdjShelfX, TU2UBAdjShelfZ, TU2UBAdjShelfM, TU2UBASEdgeSeq, TU2UBASEdgeCode)
                                DGV.Rows.Add(TU2UBGableQ, "T-Gable", TU2UBGableY, TU2UBGableX, TU2UBGableZ, TU2UBGableM, TU2UBASEdgeSeq, TU2UBASEdgeCode)
                                DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                            End If
                        Else
                            If (CKWOPLANNERSGroupBox = "GROUP1") Then
                                CabCode = CutlistForm.TUCabBox.Text
                                Dim CabCode2 As String = "U"
                                DGV.Rows.Add(CabCode)
                                DGV.Rows.Add(VarWidthI & "-" & TH & "-" & VarDepthI)
                                DGV.Rows.Add(CabCode2, "Qty", VarAmountI)
                                DGV.Rows.Add(VarWidthI & "-" & CH1 & "-" & VarDepthI)
                                DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                                DGV.Rows.Add(TU2UUBackQ, "U-Back", TU2UUBackY, TU2UUBackX, TU2UUBackZ, TU2UUBackM)
                                DGV.Rows.Add(TU2UUTopQ, "U-Top", TU2UUTopY, TU2UUTopX, TU2UUTopZ, TU2UUTopM, TU2UUTEdgeSeq, TU2UUTEdgeCode)
                                DGV.Rows.Add(TU2UUBotQ, "U-Bottom", TU2UUBotY, TU2UUBotX, TU2UUBotZ, TU2UUBotM, TU2UUBEdgeSeq, TU2UUBEdgeCode)
                                DGV.Rows.Add(TU2UUAdjShelfQ, "U-Adj.Sh", TU2UUAdjShelfY, TU2UUAdjShelfX, TU2UUAdjShelfZ, TU2UUAdjShelfM, TU2UUASEdgeSeq, TU2UUASEdgeCode)
                                DGV.Rows.Add(TU2UUGableQ, "U-Gable", TU2UUGableY, TU2UUGableX, TU2UUGableZ, TU2UUGableM, TU2UUGEdgeSeq, TU2UUGEdgeCode, TU2UUGEdgeSeq2, TU2UUGEdgeCode2)
                                DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                                CabCode2 = "T"
                                DGV.Rows.Add(CabCode)
                                DGV.Rows.Add(VarWidthI & "-" & TH & "-" & VarDepthI)
                                DGV.Rows.Add(CabCode2, "Qty", VarAmountI)
                                DGV.Rows.Add(VarWidthI & "-" & CH2 & "-" & VarDepthI)
                                DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                                DGV.Rows.Add(TU2UBBackQ, "T-Back", TU2UBBackY, TU2UBBackX, TU2UBBackZ, TU2UBBackM)
                                DGV.Rows.Add(TU2UBTopQ, "T-Top", TU2UBTopY, TU2UBTopX, TU2UBTopZ, TU2UBTopM, TU2UBTEdgeSeq, TU2UBTEdgeCode)
                                DGV.Rows.Add(TU2UBBotQ, "T-Bottom", TU2UBBotY, TU2UBBotX, TU2UBBotZ, TU2UBBotM, TU2UBBEdgeSeq, TU2UBBEdgeCode)
                                DGV.Rows.Add(TU2UBSFSQ, "T-S.F.Sh", TU2UBSFSY, TU2UBSFSX, TU2UBSFSZ, TU2UBSFSM, TU2UBSFSEdgeSeq, TU2UBSFSEdgeCode)
                                DGV.Rows.Add(TU2UBAdjShelfQ, "T-Adj.Sh", TU2UBAdjShelfY, TU2UBAdjShelfX, TU2UBAdjShelfZ, TU2UBAdjShelfM, TU2UBASEdgeSeq, TU2UBASEdgeCode)
                                DGV.Rows.Add(TU2UBGableQ, "T-Gable", TU2UBGableY, TU2UBGableX, TU2UBGableZ, TU2UBGableM, TU2UBGEdgeSeq, TU2UBGEdgeCode)
                                DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                            End If

                            '##########
                            '# GROUP2 #
                            '##########
                            '
                            If (CKWOPLANNERSGroupBox = "GROUP2") Then
                                CabCode = CutlistForm.TUCabBox.Text
                                Dim CabCode2 As String = "U-"
                                DGV.Rows.Add(CabCode)
                                DGV.Rows.Add(VarWidthI & "-" & TH & "-" & VarDepthI)
                                DGV.Rows.Add(CabCode2, "Qty", VarAmountI)
                                DGV.Rows.Add(CabSize1)
                                DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                                DGV.Rows.Add(TU2UUBackQ, "U-Back", TU2UUBackY, TU2UUBackX, TU2UUBackZ, TU2UUBackM)
                                DGV.Rows.Add(TU2UUTopBtmQ, "U-TopBtm", TU2UUTopBtmY, TU2UUTopBtmX, TU2UUTopBtmZ, TU2UUTopBtmM, TU2UUTBEdgeSeq, TU2UUTBEdgeCode)
                                DGV.Rows.Add(TU2UUAdjShelfQ, "U-Adj.Sh", TU2UUAdjShelfY, TU2UUAdjShelfX, TU2UUAdjShelfZ, TU2UUAdjShelfM, TU2UUASEdgeSeq, TU2UUASEdgeCode)
                                DGV.Rows.Add(TU2UUGableQ, "U-Gable", TU2UUGableY, TU2UUGableX, TU2UUGableZ, TU2UUGableM, TU2UUGEdgeSeq, TU2UUGEdgeCode, TU2UUGEdgeSeq2, TU2UUGEdgeCode2)
                                DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                                CabCode2 = "T"
                                DGV.Rows.Add(CabCode)
                                DGV.Rows.Add(VarWidthI & "-" & TH & "-" & VarDepthI)
                                DGV.Rows.Add(CabCode2, "Qty", VarAmountI)
                                DGV.Rows.Add(CabSize2)
                                DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                                DGV.Rows.Add(TU2UBBackQ, "T-Back", TU2UBBackY, TU2UBBackX, TU2UBBackZ, TU2UBBackM)
                                DGV.Rows.Add(TU2UBTopBtmQ, "T-TopBtm", TU2UBTopBtmY, TU2UBTopBtmX, TU2UBTopBtmZ, TU2UBTopBtmM, TU2UBTBEdgeSeq, TU2UBTBEdgeCode)
                                DGV.Rows.Add(TU2UBSFSQ, "T-S.F.Sh", TU2UBSFSY, TU2UBSFSX, TU2UBSFSZ, TU2UBSFSM, TU2UBSFSEdgeSeq, TU2UBSFSEdgeCode)
                                DGV.Rows.Add(TU2UBAdjShelfQ, "T-Adj.Sh", TU2UBAdjShelfY, TU2UBAdjShelfX, TU2UBAdjShelfZ, TU2UBAdjShelfM, TU2UBASEdgeSeq, TU2UBASEdgeCode)
                                DGV.Rows.Add(TU2UBGableQ, "T-Gable", TU2UBGableY, TU2UBGableX, TU2UBGableZ, TU2UBGableM, TU2UBASEdgeSeq, TU2UBASEdgeCode)
                                DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                            End If
                        End If

                    Case "TURS"
                        Dim CH1 = (TU2UUGableX + 1) / 10
                        Dim CH1MI = (TU2UUMIGableX + 1) / 10
                        Dim CH2 = (TU2UBGableX) / 10
                        Dim TH = CH1 + CH2
                        '################
                        '# GROUP1 RULES #
                        '################
                        '
                        If (CutlistForm.MatchingInteriorCheck.Checked) Then
                            If (CKWOPLANNERSGroupBox = "GROUP1") Then
                                CabCode = CutlistForm.TUCabBox.Text
                                Dim CabCode2 As String = "U"
                                DGV.Rows.Add(CabCode)
                                DGV.Rows.Add(VarWidthI & "-" & TH & "-" & VarDepthI)
                                DGV.Rows.Add(CabCode2, "Qty", VarAmountI)
                                DGV.Rows.Add(VarWidthI & "-" & CH1MI & "-" & VarDepthI)
                                DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                                DGV.Rows.Add(TU2UUMIBackQ, "U-Back", TU2UUMIBackY, TU2UUMIBackX, TU2UUMIBackZ, TU2UUMIBackM)
                                DGV.Rows.Add(TU2UUMITopQ, "U-Top", TU2UUMITopY, TU2UUMITopX, TU2UUMITopZ, TU2UUMITopM, TU2UUMITEdgeSeq, TU2UUMITEdgeCode)
                                DGV.Rows.Add(TU2UUMIBotQ, "U-Bottom", TU2UUMIBotY, TU2UUMIBotX, TU2UUMIBotZ, TU2UUMIBotM, TU2UUMIBEdgeSeq, TU2UUMIBEdgeCode)
                                DGV.Rows.Add(TU2UUMIAdjShelfQ, "U-Adj.Sh", TU2UUMIAdjShelfY, TU2UUMIAdjShelfX, TU2UUMIAdjShelfZ, TU2UUMIAdjShelfM, TU2UUMIASEdgeSeq, TU2UUMIASEdgeCode)
                                DGV.Rows.Add(TU2UUMIGableQ, "U-Gable", TU2UUMIGableY, TU2UUMIGableX, TU2UUMIGableZ, TU2UUMIGableM, TU2UUMIGEdgeSeq, TU2UUMIGEdgeCode, TU2UUMIGEdgeSeq2, TU2UUMIGEdgeCode2)
                                DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                                Select Case BCBCase
                                    Case "B1D", "B2D", "B3D", "B4D", "BTD", "BTDD"
                                        CabCode2 = CutlistForm.BaseCabCodeBox2.Text & "-" & Hardware
                                    Case Else
                                        CabCode2 = CutlistForm.BaseCabCodeBox2.Text
                                End Select
                                DGV.Rows.Add(CabCode)
                                DGV.Rows.Add(VarWidthI & "-" & TH & "-" & VarDepthI)
                                DGV.Rows.Add(CabCode2, "Qty", VarAmountI)
                                DGV.Rows.Add(VarWidthI & "-" & CH2 & "-" & VarDepthI)
                                DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                                DGV.Rows.Add(TU2UBBackQ, "B-Back", TU2UBBackY, TU2UBBackX, TU2UBBackZ, TU2UBBackM)
                                DGV.Rows.Add(TU2UBTopQ, "B-Top", TU2UBTopY, TU2UBTopX, TU2UBTopZ, TU2UBTopM, TU2UBTEdgeSeq, TU2UBTEdgeCode)
                                DGV.Rows.Add(TU2UBBotQ, "B-Bottom", TU2UBBotY, TU2UBBotX, TU2UBBotZ, TU2UBBotM, TU2UBBEdgeSeq, TU2UBBEdgeCode)
                                DGV.Rows.Add(TU2UBAdjShelfQ, "B-Adj.Sh", TU2UBAdjShelfY, TU2UBAdjShelfX, TU2UBAdjShelfZ, TU2UBAdjShelfM, TU2UBASEdgeSeq, TU2UBASEdgeCode)
                                DGV.Rows.Add(TU2UBGableQ, "B-Gable", TU2UBGableY, TU2UBGableX, TU2UBGableZ, TU2UBGableM, TU2UBGEdgeSeq, TU2UBGEdgeCode)
                                DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                            End If

                            '##########
                            '# GROUP2 #
                            '##########
                            '
                            If (CKWOPLANNERSGroupBox = "GROUP2") Then
                                CabCode = CutlistForm.TUCabBox.Text
                                Dim CabCode2 As String = "U"
                                DGV.Rows.Add(CabCode)
                                DGV.Rows.Add(VarWidthI & "-" & TH & "-" & VarDepthI)
                                DGV.Rows.Add(CabCode2, "Qty", VarAmountI)
                                DGV.Rows.Add(VarWidthI & "-" & CH1MI & "-" & VarDepthI)
                                DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                                DGV.Rows.Add(TU2UUMIBackQ, "U-Back", TU2UUMIBackY, TU2UUMIBackX, TU2UUMIBackZ, TU2UUMIBackM)
                                DGV.Rows.Add(TU2UUMITopBtmQ, "U-TopBtm", TU2UUMITopBtmY, TU2UUMITopBtmX, TU2UUMITopBtmZ, TU2UUMITopBtmM, TU2UUMITBEdgeSeq, TU2UUMITBEdgeCode)
                                DGV.Rows.Add(TU2UUMIAdjShelfQ, "U-Adj.Sh", TU2UUMIAdjShelfY, TU2UUMIAdjShelfX, TU2UUMIAdjShelfZ, TU2UUMIAdjShelfM, TU2UUMIASEdgeSeq, TU2UUMIASEdgeCode)
                                DGV.Rows.Add(TU2UUMIGableQ, "U-Gable", TU2UUMIGableY, TU2UUMIGableX, TU2UUMIGableZ, TU2UUMIGableM, TU2UUMIGEdgeSeq, TU2UUMIGEdgeCode, TU2UUMIGEdgeSeq2, TU2UUMIGEdgeCode2)
                                DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                                Select Case BCBCase
                                    Case "B1D", "B2D", "B3D", "B4D", "BTD", "BTDD"
                                        CabCode2 = CutlistForm.BaseCabCodeBox2.Text & "-" & Hardware
                                    Case Else
                                        CabCode2 = CutlistForm.BaseCabCodeBox2.Text
                                End Select
                                DGV.Rows.Add(CabCode)
                                DGV.Rows.Add(VarWidthI & "-" & TH & "-" & VarDepthI)
                                DGV.Rows.Add(CabCode2, "Qty", VarAmountI)
                                DGV.Rows.Add(VarWidthI & "-" & CH2 & "-" & VarDepthI)
                                DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                                DGV.Rows.Add(TU2UBBackQ, "T-Back", TU2UBBackY, TU2UBBackX, TU2UBBackZ, TU2UBBackM)
                                DGV.Rows.Add(TU2UBTopBtmQ, "T-TopBtm", TU2UBTopBtmY, TU2UBTopBtmX, TU2UBTopBtmZ, TU2UBTopBtmM, TU2UBTBEdgeSeq, TU2UBTBEdgeCode)
                                DGV.Rows.Add(TU2UBSFSQ, "T-S.F.Sh", TU2UBSFSY, TU2UBSFSX, TU2UBSFSZ, TU2UBSFSM, TU2UBSFSEdgeSeq, TU2UBSFSEdgeCode)
                                DGV.Rows.Add(TU2UBAdjShelfQ, "T-Adj.Sh", TU2UBAdjShelfY, TU2UBAdjShelfX, TU2UBAdjShelfZ, TU2UBAdjShelfM, TU2UBASEdgeSeq, TU2UBASEdgeCode)
                                DGV.Rows.Add(TU2UBGableQ, "T-Gable", TU2UBGableY, TU2UBGableX, TU2UBGableZ, TU2UBGableM, TU2UBASEdgeSeq, TU2UBASEdgeCode)
                                DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                            End If
                        Else
                            If (CKWOPLANNERSGroupBox = "GROUP1") Then
                                CabCode = CutlistForm.TUCabBox.Text
                                Dim CabCode2 As String = "U"
                                DGV.Rows.Add(CabCode)
                                DGV.Rows.Add(VarWidthI & "-" & TH & "-" & VarDepthI)
                                DGV.Rows.Add(CabCode2, "Qty", VarAmountI)
                                DGV.Rows.Add(VarWidthI & "-" & CH1 & "-" & VarDepthI)
                                DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                                DGV.Rows.Add(TU2UUBackQ, "U-Back", TU2UUBackY, TU2UUBackX, TU2UUBackZ, TU2UUBackM)
                                DGV.Rows.Add(TU2UUTopQ, "U-Top", TU2UUTopY, TU2UUTopX, TU2UUTopZ, TU2UUTopM, TU2UUTEdgeSeq, TU2UUTEdgeCode)
                                DGV.Rows.Add(TU2UUBotQ, "U-Bottom", TU2UUBotY, TU2UUBotX, TU2UUBotZ, TU2UUBotM, TU2UUBEdgeSeq, TU2UUBEdgeCode)
                                DGV.Rows.Add(TU2UUAdjShelfQ, "U-Adj.Sh", TU2UUAdjShelfY, TU2UUAdjShelfX, TU2UUAdjShelfZ, TU2UUAdjShelfM, TU2UUASEdgeSeq, TU2UUASEdgeCode)
                                DGV.Rows.Add(TU2UUGableQ, "U-Gable", TU2UUGableY, TU2UUGableX, TU2UUGableZ, TU2UUGableM, TU2UUGEdgeSeq, TU2UUGEdgeCode, TU2UUGEdgeSeq2, TU2UUGEdgeCode2)
                                DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                                Select Case BCBCase
                                    Case "B1D", "B2D", "B3D", "B4D", "BTD", "BTDD"
                                        CabCode2 = CutlistForm.BaseCabCodeBox2.Text & "-" & Hardware
                                    Case Else
                                        CabCode2 = CutlistForm.BaseCabCodeBox2.Text
                                End Select
                                DGV.Rows.Add(CabCode)
                                DGV.Rows.Add(VarWidthI & "-" & TH & "-" & VarDepthI)
                                DGV.Rows.Add(CabCode2, "Qty", VarAmountI)
                                DGV.Rows.Add(VarWidthI & "-" & CH2 & "-" & VarDepthI)
                                DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                                DGV.Rows.Add(TU2UBBackQ, "T-Back", TU2UBBackY, TU2UBBackX, TU2UBBackZ, TU2UBBackM)
                                DGV.Rows.Add(TU2UBTopQ, "T-Top", TU2UBTopY, TU2UBTopX, TU2UBTopZ, TU2UBTopM, TU2UBTEdgeSeq, TU2UBTEdgeCode)
                                DGV.Rows.Add(TU2UBBotQ, "T-Bottom", TU2UBBotY, TU2UBBotX, TU2UBBotZ, TU2UBBotM, TU2UBBEdgeSeq, TU2UBBEdgeCode)
                                DGV.Rows.Add(TU2UBSFSQ, "T-S.F.Sh", TU2UBSFSY, TU2UBSFSX, TU2UBSFSZ, TU2UBSFSM, TU2UBSFSEdgeSeq, TU2UBSFSEdgeCode)
                                DGV.Rows.Add(TU2UBAdjShelfQ, "T-Adj.Sh", TU2UBAdjShelfY, TU2UBAdjShelfX, TU2UBAdjShelfZ, TU2UBAdjShelfM, TU2UBASEdgeSeq, TU2UBASEdgeCode)
                                DGV.Rows.Add(TU2UBGableQ, "T-Gable", TU2UBGableY, TU2UBGableX, TU2UBGableZ, TU2UBGableM, TU2UBGEdgeSeq, TU2UBGEdgeCode)
                                DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                            End If

                            '##########
                            '# GROUP2 #
                            '##########
                            '
                            If (CKWOPLANNERSGroupBox = "GROUP2") Then
                                CabCode = CutlistForm.TUCabBox.Text
                                Dim CabCode2 As String = "U"
                                DGV.Rows.Add(CabCode)
                                DGV.Rows.Add(VarWidthI & "-" & TH & "-" & VarDepthI)
                                DGV.Rows.Add(CabCode2, "Qty", VarAmountI)
                                DGV.Rows.Add(VarWidthI & "-" & CH1 & "-" & VarDepthI)
                                DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                                DGV.Rows.Add(TU2UUBackQ, "U-Back", TU2UUBackY, TU2UUBackX, TU2UUBackZ, TU2UUBackM)
                                DGV.Rows.Add(TU2UUTopBtmQ, "U-TopBtm", TU2UUTopBtmY, TU2UUTopBtmX, TU2UUTopBtmZ, TU2UUTopBtmM, TU2UUTBEdgeSeq, TU2UUTBEdgeCode)
                                DGV.Rows.Add(TU2UUAdjShelfQ, "U-Adj.Sh", TU2UUAdjShelfY, TU2UUAdjShelfX, TU2UUAdjShelfZ, TU2UUAdjShelfM, TU2UUASEdgeSeq, TU2UUASEdgeCode)
                                DGV.Rows.Add(TU2UUGableQ, "U-Gable", TU2UUGableY, TU2UUGableX, TU2UUGableZ, TU2UUGableM, TU2UUGEdgeSeq, TU2UUGEdgeCode, TU2UUGEdgeSeq2, TU2UUGEdgeCode2)
                                DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                                Select Case BCBCase
                                    Case "B1D", "B2D", "B3D", "B4D", "BTD", "BTDD"
                                        CabCode2 = CutlistForm.BaseCabCodeBox2.Text & "-" & Hardware
                                    Case Else
                                        CabCode2 = CutlistForm.BaseCabCodeBox2.Text
                                End Select
                                DGV.Rows.Add(CabCode)
                                DGV.Rows.Add(VarWidthI & "-" & TH & "-" & VarDepthI)
                                DGV.Rows.Add(CabCode2, "Qty", VarAmountI)
                                DGV.Rows.Add(VarWidthI & "-" & CH2 & "-" & VarDepthI)
                                DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                                DGV.Rows.Add(TU2UBBackQ, "T-Back", TU2UBBackY, TU2UBBackX, TU2UBBackZ, TU2UBBackM)
                                DGV.Rows.Add(TU2UBTopBtmQ, "T-TopBtm", TU2UBTopBtmY, TU2UBTopBtmX, TU2UBTopBtmZ, TU2UBTopBtmM, TU2UBTBEdgeSeq, TU2UBTBEdgeCode)
                                DGV.Rows.Add(TU2UBSFSQ, "T-S.F.Sh", TU2UBSFSY, TU2UBSFSX, TU2UBSFSZ, TU2UBSFSM, TU2UBSFSEdgeSeq, TU2UBSFSEdgeCode)
                                DGV.Rows.Add(TU2UBAdjShelfQ, "T-Adj.Sh", TU2UBAdjShelfY, TU2UBAdjShelfX, TU2UBAdjShelfZ, TU2UBAdjShelfM, TU2UBASEdgeSeq, TU2UBASEdgeCode)
                                DGV.Rows.Add(TU2UBGableQ, "T-Gable", TU2UBGableY, TU2UBGableX, TU2UBGableZ, TU2UBGableM, TU2UBASEdgeSeq, TU2UBASEdgeCode)
                                DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                            End If
                        End If
                    Case Else
                End Select

            '################################
            '# HUTCH DRAWER # 9 OR 10 LINES #
            '################################
            '
            Case "HUTCH DRAWER"
                CabCode = "HD" & CutlistForm.OptionsVar & "-" & Hardware
                '# MATCHING INTERIOR CHECK #
                If (CutlistForm.MatchingInteriorCheck.Checked) Then
                    If (HMIAdjShelfQ = 0 Or CutlistForm.GlassCheck.Checked = True) Then
                        DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                        DGV.Rows.Add(CabSize1)
                        DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                        DGV.Rows.Add(HMIBackQ, "Back", HMIBackY, HMIBackX, HMIBackZ, HMIBackM)
                        DGV.Rows.Add(HMITopQ, "Top", HMITopY, HMITopX, HMITopZ, HMITopM, HMITEdgeSeq, HMITEdgeCode)
                        DGV.Rows.Add(HMITopFFShelfQ, "T-F.F.Sh", HMITopFFShelfY, HMITopFFShelfX, HMITopFFShelfZ, HMITopFFShelfM, HMITFFSEdgeSeq, HMITFFSEdgeCode)
                        DGV.Rows.Add(HMIBtmQ, "Bottom", HMIBtmY, HMIBtmX, HMIBtmZ, HMIBtmM, HMIBEdgeSeq, HMIBEdgeCode)
                        DGV.Rows.Add(HMIStrapQ, "Strap", HMIStrapY, HMIStrapX, HMIStrapZ, HMIStrapM, HMISEdgeSeq, HMISEdgeCode)
                        DGV.Rows.Add(HMIGableQ, "Gable", HMIGableY, HMIGableX, HMIGableZ, HMIGableM, HMIGEdgeSeq, HMIGEdgeCode, HMIGEdgeSeq2, HMIGEdgeCode2)
                        DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                    Else
                        DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                        DGV.Rows.Add(CabSize1)
                        DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                        DGV.Rows.Add(HMIBackQ, "Back", HMIBackY, HMIBackX, HMIBackZ, HMIBackM)
                        DGV.Rows.Add(HMITopQ, "Top", HMITopY, HMITopX, HMITopZ, HMITopM, HMITEdgeSeq, HMITEdgeCode)
                        DGV.Rows.Add(HMITopFFShelfQ, "T-F.F.Sh", HMITopFFShelfY, HMITopFFShelfX, HMITopFFShelfZ, HMITopFFShelfM, HMITFFSEdgeSeq, HMITFFSEdgeCode)
                        DGV.Rows.Add(HMIBtmQ, "Bottom", HMIBtmY, HMIBtmX, HMIBtmZ, HMIBtmM, HMIBEdgeSeq, HMIBEdgeCode)
                        DGV.Rows.Add(HMIStrapQ, "Strap", HMIStrapY, HMIStrapX, HMIStrapZ, HMIStrapM, HMISEdgeSeq, HMISEdgeCode)
                        If (CKWOPLANNERSSpeciesBox = "MDF") Then
                            DGV.Rows.Add(HMIAdjShelfQ, "Adj.Sh", HMIAdjShelfY, HMIAdjShelfX, HMIAdjShelfZ, HMIAdjShelfM)
                        Else
                            DGV.Rows.Add(HMIAdjShelfQ, "Adj.Sh", HMIAdjShelfY, HMIAdjShelfX, HMIAdjShelfZ, HMIAdjShelfM, HMIASEdgeSeq, HMIASEdgeCode)
                        End If
                        DGV.Rows.Add(HMIGableQ, "Gable", HMIGableY, HMIGableX, HMIGableZ, HMIGableM, HMIGEdgeSeq, HMIGEdgeCode, HMIGEdgeSeq2, HMIGEdgeCode2)
                        DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                    End If
                Else
                    If (HAdjShelfQ = 0 Or CutlistForm.GlassCheck.Checked = True) Then
                        DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                        DGV.Rows.Add(CabSize1)
                        DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                        DGV.Rows.Add(HBackQ, "Back", HBackY, HBackX, HBackZ, HBackM)
                        DGV.Rows.Add(HTopQ, "Top", HTopY, HTopX, HTopZ, HTopM, HTEdgeSeq, HTEdgeCode)
                        DGV.Rows.Add(HTopFFShelfQ, "T-F.F.Sh", HTopFFShelfY, HTopFFShelfX, HTopFFShelfZ, HTopFFShelfM, HTFFSEdgeSeq, HTFFSEdgeCode)
                        DGV.Rows.Add(HBtmQ, "Bottom", HBtmY, HBtmX, HBtmZ, HBtmM, HBEdgeSeq, HBEdgeCode)
                        DGV.Rows.Add(HStrapQ, "Strap", HStrapY, HStrapX, HStrapZ, HStrapM, HSEdgeSeq, HSEdgeCode)
                        DGV.Rows.Add(HGableQ, "Gable", HGableY, HGableX, HGableZ, HGableM, HGEdgeSeq, HGEdgeCode, HGEdgeSeq2, HGEdgeCode2)
                        DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                    Else
                        DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                        DGV.Rows.Add(CabSize1)
                        DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                        DGV.Rows.Add(HBackQ, "Back", HBackY, HBackX, HBackZ, HBackM)
                        DGV.Rows.Add(HTopQ, "Top", HTopY, HTopX, HTopZ, HTopM, HTEdgeSeq, HTEdgeCode)
                        DGV.Rows.Add(HTopFFShelfQ, "T-F.F.Sh", HTopFFShelfY, HTopFFShelfX, HTopFFShelfZ, HTopFFShelfM, HTFFSEdgeSeq, HTFFSEdgeCode)
                        DGV.Rows.Add(HBtmQ, "Bottom", HBtmY, HBtmX, HBtmZ, HBtmM, HBEdgeSeq, HBEdgeCode)
                        DGV.Rows.Add(HStrapQ, "Strap", HStrapY, HStrapX, HStrapZ, HStrapM, HSEdgeSeq, HSEdgeCode)
                        DGV.Rows.Add(HAdjShelfQ, "Adj.Sh", HAdjShelfY, HAdjShelfX, HAdjShelfZ, HAdjShelfM, HASEdgeSeq, HASEdgeCode)
                        DGV.Rows.Add(HGableQ, "Gable", HGableY, HGableX, HGableZ, HGableM, HGEdgeSeq, HGEdgeCode, HGEdgeSeq2, HGEdgeCode2)
                        DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                    End If
                End If

            '######################################
            '# HUTCH DRAWER STACK # 9 OR 10 LINES #
            '######################################
            '
            Case "HUTCH DRAWER STACK"
                CabCode = "HDS" & CutlistForm.OptionsVar & "-" & Hardware
                If (CutlistForm.MatchingInteriorCheck.Checked) Then
                    If (HMIAdjShelfQ = 0 Or CutlistForm.GlassCheck.Checked = True) Then
                        DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                        DGV.Rows.Add(CabSize1)
                        DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                        DGV.Rows.Add(HMIBackQ, "Back", HMIBackY, HMIBackX, HMIBackZ, HMIBackM)
                        DGV.Rows.Add(HMITopQ, "Top", HMITopY, HMITopX, HMITopZ, HMITopM, HMITEdgeSeq, HMITEdgeCode)
                        DGV.Rows.Add(HMITopFFShelfQ, "T-F.F.Sh", HMITopFFShelfY, HMITopFFShelfX, HMITopFFShelfZ, HMITopFFShelfM, HMITFFSEdgeSeq, HMITFFSEdgeCode)
                        DGV.Rows.Add(HMIBtmQ, "Bottom", HMIBtmY, HMIBtmX, HMIBtmZ, HMIBtmM, HMIBEdgeSeq, HMIBEdgeCode)
                        DGV.Rows.Add(HMIStrapQ, "Strap", HMIStrapY, HMIStrapX, HMIStrapZ, HMIStrapM, HMISEdgeSeq, HMISEdgeCode)
                        DGV.Rows.Add(HMIGableQ, "Gable", HMIGableY, HMIGableX, HMIGableZ, HMIGableM, HMIGEdgeSeq, HMIGEdgeCode, HMIGEdgeSeq2, HMIGEdgeCode2)
                        DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                    Else
                        DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                        DGV.Rows.Add(CabSize1)
                        DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                        DGV.Rows.Add(HMIBackQ, "Back", HMIBackY, HMIBackX, HMIBackZ, HMIBackM)
                        DGV.Rows.Add(HMITopQ, "Top", HMITopY, HMITopX, HMITopZ, HMITopM, HMITEdgeSeq, HMITEdgeCode)
                        DGV.Rows.Add(HMITopFFShelfQ, "T-F.F.Sh", HMITopFFShelfY, HMITopFFShelfX, HMITopFFShelfZ, HMITopFFShelfM, HMITFFSEdgeSeq, HMITFFSEdgeCode)
                        DGV.Rows.Add(HMIBtmQ, "Bottom", HMIBtmY, HMIBtmX, HMIBtmZ, HMIBtmM, HMIBEdgeSeq, HMIBEdgeCode)
                        DGV.Rows.Add(HMIStrapQ, "Strap", HMIStrapY, HMIStrapX, HMIStrapZ, HMIStrapM, HMISEdgeSeq, HMISEdgeCode)
                        If (CKWOPLANNERSSpeciesBox = "MDF") Then
                            DGV.Rows.Add(HMIAdjShelfQ, "Adj.Sh", HMIAdjShelfY, HMIAdjShelfX, HMIAdjShelfZ, HMIAdjShelfM)
                        Else
                            DGV.Rows.Add(HMIAdjShelfQ, "Adj.Sh", HMIAdjShelfY, HMIAdjShelfX, HMIAdjShelfZ, HMIAdjShelfM, HMIASEdgeSeq, HMIASEdgeCode)
                        End If
                        DGV.Rows.Add(HMIGableQ, "Gable", HMIGableY, HMIGableX, HMIGableZ, HMIGableM, HMIGEdgeSeq, HMIGEdgeCode, HMIGEdgeSeq2, HMIGEdgeCode2)
                        DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                    End If
                Else
                    If (HAdjShelfQ = 0 Or CutlistForm.GlassCheck.Checked = True) Then
                        DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                        DGV.Rows.Add(CabSize1)
                        DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                        DGV.Rows.Add(HBackQ, "Back", HBackY, HBackX, HBackZ, HBackM)
                        DGV.Rows.Add(HTopQ, "Top", HTopY, HTopX, HTopZ, HTopM, HTEdgeSeq, HTEdgeCode)
                        DGV.Rows.Add(HTopFFShelfQ, "T-F.F.Sh", HTopFFShelfY, HTopFFShelfX, HTopFFShelfZ, HTopFFShelfM, HTFFSEdgeSeq, HTFFSEdgeCode)
                        DGV.Rows.Add(HBtmQ, "Bottom", HBtmY, HBtmX, HBtmZ, HBtmM, HBEdgeSeq, HBEdgeCode)
                        DGV.Rows.Add(HStrapQ, "Strap", HStrapY, HStrapX, HStrapZ, HStrapM, HSEdgeSeq, HSEdgeCode)
                        DGV.Rows.Add(HGableQ, "Gable", HGableY, HGableX, HGableZ, HGableM, HGEdgeSeq, HGEdgeCode, HGEdgeSeq2, HGEdgeCode2)
                        DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                    Else
                        DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                        DGV.Rows.Add(CabSize1)
                        DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                        DGV.Rows.Add(HBackQ, "Back", HBackY, HBackX, HBackZ, HBackM)
                        DGV.Rows.Add(HTopQ, "Top", HTopY, HTopX, HTopZ, HTopM, HTEdgeSeq, HTEdgeCode)
                        DGV.Rows.Add(HTopFFShelfQ, "T-F.F.Sh", HTopFFShelfY, HTopFFShelfX, HTopFFShelfZ, HTopFFShelfM, HTFFSEdgeSeq, HTFFSEdgeCode)
                        DGV.Rows.Add(HBtmQ, "Bottom", HBtmY, HBtmX, HBtmZ, HBtmM, HBEdgeSeq, HBEdgeCode)
                        DGV.Rows.Add(HStrapQ, "Strap", HStrapY, HStrapX, HStrapZ, HStrapM, HSEdgeSeq, HSEdgeCode)
                        DGV.Rows.Add(HAdjShelfQ, "Adj.Sh", HAdjShelfY, HAdjShelfX, HAdjShelfZ, HAdjShelfM, HASEdgeSeq, HASEdgeCode)
                        DGV.Rows.Add(HGableQ, "Gable", HGableY, HGableX, HGableZ, HGableM, HGEdgeSeq, HGEdgeCode, HGEdgeSeq2, HGEdgeCode2)
                        DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                    End If
                End If

            '########################################
            '# HUTCH DOUBLE DRAWER # 10 OR 11 LINES #
            '########################################
            '
            Case "HUTCH DOUBLE DRAWER"
                CabCode = "HDD" & CutlistForm.OptionsVar & "-" & Hardware
                If (CutlistForm.MatchingInteriorCheck.Checked) Then
                    If (HMIAdjShelfQ = 0 Or CutlistForm.GlassCheck.Checked = True) Then
                        DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                        DGV.Rows.Add(CabSize1)
                        DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                        DGV.Rows.Add(HMIBackQ, "Back", HMIBackY, HMIBackX, HMIBackZ, HMIBackM)
                        DGV.Rows.Add(HMITopFFShelfQ, "T-F.F.Sh", HMITopFFShelfY, HMITopFFShelfX, HMITopFFShelfZ, HMITopFFShelfM, HMITFFSEdgeSeq, HMITFFSEdgeCode)
                        DGV.Rows.Add(HMIBtmFFShelfQ, "B-F.F.Sh", HMIBtmFFShelfY, HMIBtmFFShelfX, HMIBtmFFShelfZ, HMIBtmFFShelfM, HMIBFFSEdgeSeq, HMIBFFSEdgeCode)
                        DGV.Rows.Add(HMIDividerQ, "Divider", HMIDividerY, HMIDividerX, HMIDividerZ, HMIDividerM, HMIDEdgeSeq, HMIDEdgeCode)
                        DGV.Rows.Add(HMIGableQ, "Gable", HMIGableY, HMIGableX, HMIGableZ, HMIGableM, HMIGEdgeSeq, HMIGEdgeCode, HMIGEdgeSeq2, HMIGEdgeCode2)
                        DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                    Else
                        DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                        DGV.Rows.Add(CabSize1)
                        DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                        DGV.Rows.Add(HMIBackQ, "Back", HMIBackY, HMIBackX, HMIBackZ, HMIBackM)
                        DGV.Rows.Add(HMITopFFShelfQ, "T-F.F.Sh", HMITopFFShelfY, HMITopFFShelfX, HMITopFFShelfZ, HMITopFFShelfM, HMITFFSEdgeSeq, HMITFFSEdgeCode)
                        DGV.Rows.Add(HMIBtmFFShelfQ, "B-F.F.Sh", HMIBtmFFShelfY, HMIBtmFFShelfX, HMIBtmFFShelfZ, HMIBtmFFShelfM, HMIBFFSEdgeSeq, HMIBFFSEdgeCode)
                        DGV.Rows.Add(HMIDividerQ, "Divider", HMIDividerY, HMIDividerX, HMIDividerZ, HMIDividerM, HMIDEdgeSeq, HMIDEdgeCode)
                        If (CKWOPLANNERSSpeciesBox = "MDF") Then
                            DGV.Rows.Add(HMIAdjShelfQ, "Adj.Sh", HMIAdjShelfY, HMIAdjShelfX, HMIAdjShelfZ, HMIAdjShelfM)
                        Else
                            DGV.Rows.Add(HMIAdjShelfQ, "Adj.Sh", HMIAdjShelfY, HMIAdjShelfX, HMIAdjShelfZ, HMIAdjShelfM, HMIASEdgeSeq, HMIASEdgeCode)
                        End If
                        DGV.Rows.Add(HMIGableQ, "Gable", HMIGableY, HMIGableX, HMIGableZ, HMIGableM, HMIGEdgeSeq, HMIGEdgeCode, HMIGEdgeSeq2, HMIGEdgeCode2)
                        DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                    End If
                Else
                    If (HAdjShelfQ = 0 Or CutlistForm.GlassCheck.Checked = True) Then
                        DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                        DGV.Rows.Add(CabSize1)
                        DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                        DGV.Rows.Add(HBackQ, "Back", HBackY, HBackX, HBackZ, HBackM)
                        DGV.Rows.Add(HTopFFShelfQ, "T-F.F.Sh", HTopFFShelfY, HTopFFShelfX, HTopFFShelfZ, HTopFFShelfM, HTFFSEdgeSeq, HTFFSEdgeCode)
                        DGV.Rows.Add(HBtmFFShelfQ, "B-F.F.Sh", HBtmFFShelfY, HBtmFFShelfX, HBtmFFShelfZ, HBtmFFShelfM, HBFFSEdgeSeq, HBFFSEdgeCode)
                        DGV.Rows.Add(HDividerQ, "Divider", HDividerY, HDividerX, HDividerZ, HDividerM, HDEdgeSeq, HDEdgeCode)
                        DGV.Rows.Add(HGableQ, "Gable", HGableY, HGableX, HGableZ, HGableM, HGEdgeSeq, HGEdgeCode, HGEdgeSeq2, HGEdgeCode2)
                        DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                    Else
                        DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                        DGV.Rows.Add(CabSize1)
                        DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                        DGV.Rows.Add(HBackQ, "Back", HBackY, HBackX, HBackZ, HBackM)
                        DGV.Rows.Add(HTopFFShelfQ, "T-F.F.Sh", HTopFFShelfY, HTopFFShelfX, HTopFFShelfZ, HTopFFShelfM, HTFFSEdgeSeq, HTFFSEdgeCode)
                        DGV.Rows.Add(HBtmFFShelfQ, "B-F.F.Sh", HBtmFFShelfY, HBtmFFShelfX, HBtmFFShelfZ, HBtmFFShelfM, HBFFSEdgeSeq, HBFFSEdgeCode)
                        DGV.Rows.Add(HDividerQ, "Divider", HDividerY, HDividerX, HDividerZ, HDividerM, HDEdgeSeq, HDEdgeCode)
                        DGV.Rows.Add(HAdjShelfQ, "Adj.Sh", HAdjShelfY, HAdjShelfX, HAdjShelfZ, HAdjShelfM, HASEdgeSeq, HASEdgeCode)
                        DGV.Rows.Add(HGableQ, "Gable", HGableY, HGableX, HGableZ, HGableM, HGEdgeSeq, HGEdgeCode, HGEdgeSeq2, HGEdgeCode2)
                        DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                    End If
                End If

            '##############################################
            '# HUTCH DOUBLE DRAWER STACK # 10 OR 11 LINES #
            '##############################################
            '
            Case "HUTCH DOUBLE DRAWER STACK"
                CabCode = "HDDS" & CutlistForm.OptionsVar & "-" & Hardware
                If (CutlistForm.MatchingInteriorCheck.Checked) Then
                    If (HMIAdjShelfQ = 0 Or CutlistForm.GlassCheck.Checked = True) Then
                        DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                        DGV.Rows.Add(CabSize1)
                        DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                        DGV.Rows.Add(HMIBackQ, "Back", HMIBackY, HMIBackX, HMIBackZ, HMIBackM)
                        DGV.Rows.Add(HMITopQ, "Top", HMITopY, HMITopX, HMITopZ, HMITopM, HMITEdgeSeq, HMITEdgeCode)
                        DGV.Rows.Add(HMITopFFShelfQ, "T-F.F.Sh", HMITopFFShelfY, HMITopFFShelfX, HMITopFFShelfZ, HMITopFFShelfM, HMITFFSEdgeSeq, HMITFFSEdgeCode)
                        DGV.Rows.Add(HMIBtmQ, "Bottom", HMIBtmY, HMIBtmX, HMIBtmZ, HMIBtmM, HMIBEdgeSeq, HMIBEdgeCode)
                        DGV.Rows.Add(HMIBtmFFShelfQ, "B-F.F.Sh", HMIBtmFFShelfY, HMIBtmFFShelfX, HMIBtmFFShelfZ, HMIBtmFFShelfM, HMIBFFSEdgeSeq, HMIBFFSEdgeCode)
                        DGV.Rows.Add(HMIDividerQ, "Divider", HMIDividerY, HMIDividerX, HMIDividerZ, HMIDividerM, HMIDEdgeSeq, HMIDEdgeCode)
                        DGV.Rows.Add(HMIGableQ, "Gable", HMIGableY, HMIGableX, HMIGableZ, HMIGableM, HMIGEdgeSeq, HMIGEdgeCode, HMIGEdgeSeq2, HMIGEdgeCode2)
                        DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                    Else
                        DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                        DGV.Rows.Add(CabSize1)
                        DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                        DGV.Rows.Add(HMIBackQ, "Back", HMIBackY, HMIBackX, HMIBackZ, HMIBackM)
                        DGV.Rows.Add(HMITopQ, "Top", HMITopY, HMITopX, HMITopZ, HMITopM, HMITEdgeSeq, HMITEdgeCode)
                        DGV.Rows.Add(HMITopFFShelfQ, "T-F.F.Sh", HMITopFFShelfY, HMITopFFShelfX, HMITopFFShelfZ, HMITopFFShelfM, HMITFFSEdgeSeq, HMITFFSEdgeCode)
                        DGV.Rows.Add(HMIBtmQ, "Bottom", HMIBtmY, HMIBtmX, HMIBtmZ, HMIBtmM, HMIBEdgeSeq, HMIBEdgeCode)
                        DGV.Rows.Add(HMIBtmFFShelfQ, "B-F.F.Sh", HMIBtmFFShelfY, HMIBtmFFShelfX, HMIBtmFFShelfZ, HMIBtmFFShelfM, HMIBFFSEdgeSeq, HMIBFFSEdgeCode)
                        DGV.Rows.Add(HMIDividerQ, "Divider", HMIDividerY, HMIDividerX, HMIDividerZ, HMIDividerM, HMIDEdgeSeq, HMIDEdgeCode)
                        If (CKWOPLANNERSSpeciesBox = "MDF") Then
                            DGV.Rows.Add(HMIAdjShelfQ, "Adj.Sh", HMIAdjShelfY, HMIAdjShelfX, HMIAdjShelfZ, HMIAdjShelfM)
                        Else
                            DGV.Rows.Add(HMIAdjShelfQ, "Adj.Sh", HMIAdjShelfY, HMIAdjShelfX, HMIAdjShelfZ, HMIAdjShelfM, HMIASEdgeSeq, HMIASEdgeCode)
                        End If
                        DGV.Rows.Add(HMIGableQ, "Gable", HMIGableY, HMIGableX, HMIGableZ, HMIGableM, HMIGEdgeSeq, HMIGEdgeCode, HMIGEdgeSeq2, HMIGEdgeCode2)
                        DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                    End If
                Else
                    If (HAdjShelfQ = 0 Or CutlistForm.GlassCheck.Checked = True) Then
                        DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                        DGV.Rows.Add(CabSize1)
                        DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                        DGV.Rows.Add(HBackQ, "Back", HBackY, HBackX, HBackZ, HBackM)
                        DGV.Rows.Add(HTopQ, "Top", HTopY, HTopX, HTopZ, HTopM, HTEdgeSeq, HTEdgeCode)
                        DGV.Rows.Add(HTopFFShelfQ, "T-F.F.Sh", HTopFFShelfY, HTopFFShelfX, HTopFFShelfZ, HTopFFShelfM, HTFFSEdgeSeq, HTFFSEdgeCode)
                        DGV.Rows.Add(HBtmQ, "Bottom", HBtmY, HBtmX, HBtmZ, HBtmM, HBEdgeSeq, HBEdgeCode)
                        DGV.Rows.Add(HBtmFFShelfQ, "B-F.F.Sh", HBtmFFShelfY, HBtmFFShelfX, HBtmFFShelfZ, HBtmFFShelfM, HBFFSEdgeSeq, HBFFSEdgeCode)
                        DGV.Rows.Add(HDividerQ, "Divider", HDividerY, HDividerX, HDividerZ, HDividerM, HDEdgeSeq, HDEdgeCode)
                        DGV.Rows.Add(HGableQ, "Gable", HGableY, HGableX, HGableZ, HGableM, HGEdgeSeq, HGEdgeCode, HGEdgeSeq2, HGEdgeCode2)
                        DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                    Else
                        DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                        DGV.Rows.Add(CabSize1)
                        DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                        DGV.Rows.Add(HBackQ, "Back", HBackY, HBackX, HBackZ, HBackM)
                        DGV.Rows.Add(HTopQ, "Top", HTopY, HTopX, HTopZ, HTopM, HTEdgeSeq, HTEdgeCode)
                        DGV.Rows.Add(HTopFFShelfQ, "T-F.F.Sh", HTopFFShelfY, HTopFFShelfX, HTopFFShelfZ, HTopFFShelfM, HTFFSEdgeSeq, HTFFSEdgeCode)
                        DGV.Rows.Add(HBtmQ, "Bottom", HBtmY, HBtmX, HBtmZ, HBtmM, HBEdgeSeq, HBEdgeCode)
                        DGV.Rows.Add(HBtmFFShelfQ, "B-F.F.Sh", HBtmFFShelfY, HBtmFFShelfX, HBtmFFShelfZ, HBtmFFShelfM, HBFFSEdgeSeq, HBFFSEdgeCode)
                        DGV.Rows.Add(HDividerQ, "Divider", HDividerY, HDividerX, HDividerZ, HDividerM, HDEdgeSeq, HDEdgeCode)
                        DGV.Rows.Add(HAdjShelfQ, "Adj.Sh", HAdjShelfY, HAdjShelfX, HAdjShelfZ, HAdjShelfM, HASEdgeSeq, HASEdgeCode)
                        DGV.Rows.Add(HGableQ, "Gable", HGableY, HGableX, HGableZ, HGableM, HGEdgeSeq, HGEdgeCode, HGEdgeSeq2, HGEdgeCode2)
                        DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                    End If
                End If

            '####################
            '# VANITY # 7 LINES #
            '####################
            '
            Case "VANITY"
                Dim VanCase = CutlistForm.VanityCodeBox.Text
                If (CutlistForm.MatchingInteriorCheck.Checked) Then
                    Select Case VanCase
                        Case "V"
                            CabCode = CutlistForm.VanityCodeBox.Text & CutlistForm.OptionsVar
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(VMIBackQ, "Back", VMIBackY, VMIBackX, VMIBackZ, VMIBackM)
                            DGV.Rows.Add(VMITopStrapQ, "Top Strap", VMITopStrapY, VMITopStrapX, VMITopStrapZ, VMITopStrapM, VMITSEdgeSeq, VMITSEdgeCode)
                            DGV.Rows.Add(VMIBtmQ, "Bottom", VMIBtmY, VMIBtmX, VMIBtmZ, VMIBtmM, VMIBEdgeSeq, VMIBEdgeCode)
                            DGV.Rows.Add(VMIGableQ, "Gable", VMIGableY, VMIGableX, VMIGableZ, VMIGableM, VMIGEdgeSeq, VMIGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                        Case "VS"
                            CabCode = CutlistForm.VanityCodeBox.Text & CutlistForm.OptionsVar
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(VMIBackQ, "Back", VMIBackY, VMIBackX, VMIBackZ, VMIBackM)
                            DGV.Rows.Add(VMITopStrapQ, "Top Strap", VMITopStrapY, VMITopStrapX, VMITopStrapZ, VMITopStrapM, VMITSEdgeSeq, VMITSEdgeCode)
                            DGV.Rows.Add(VMIBtmQ, "Bottom", VMIBtmY, VMIBtmX, VMIBtmZ, VMIBtmM, VMIBEdgeSeq, VMIBEdgeCode)
                            DGV.Rows.Add(VMIStrapQ, "Sink Strap", VMIStrapY, VMIStrapX, VMIStrapZ, VMIStrapM)
                            DGV.Rows.Add(VMIGableQ, "Gable", VMIGableY, VMIGableX, VMIGableZ, VMIGableM, VMIGEdgeSeq, VMIGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                        Case "V2D", "V3D"
                            CabCode = CutlistForm.VanityCodeBox.Text & CutlistForm.OptionsVar & "-" & Hardware
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(VMIBackQ, "Back", VMIBackY, VMIBackX, VMIBackZ, VMIBackM)
                            DGV.Rows.Add(VMITopStrapQ, "Top Strap", VMITopStrapY, VMITopStrapX, VMITopStrapZ, VMITopStrapM, VMITSEdgeSeq, VMITSEdgeCode)
                            DGV.Rows.Add(VMIBtmQ, "Bottom", VMIBtmY, VMIBtmX, VMIBtmZ, VMIBtmM, VMIBEdgeSeq, VMIBEdgeCode)
                            DGV.Rows.Add(VMIGableQ, "Gable", VMIGableY, VMIGableX, VMIGableZ, VMIGableM, VMIGEdgeSeq, VMIGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                        Case "VBD"
                            CabCode = CutlistForm.VanityCodeBox.Text & CutlistForm.OptionsVar & "-" & Hardware
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(VMIBackQ, "Back", VMIBackY, VMIBackX, VMIBackZ, VMIBackM)
                            DGV.Rows.Add(VMITopStrapQ, "Top Strap", VMITopStrapY, VMITopStrapX, VMITopStrapZ, VMITopStrapM, VMITSEdgeSeq, VMITSEdgeCode)
                            DGV.Rows.Add(VMIBtmQ, "Bottom", VMIBtmY, VMIBtmX, VMIBtmZ, VMIBtmM, VMIBEdgeSeq, VMIBEdgeCode)
                            DGV.Rows.Add(VMIFFShelfQ, "F.F.Sh", VMIFFShelfY, VMIFFShelfX, VMIFFShelfZ, VMIFFShelfM, VMIFFSHEdgeSeq, VMIFFSHEdgeCode)
                            DGV.Rows.Add(VMIGableQ, "Gable", VMIGableY, VMIGableX, VMIGableZ, VMIGableM, VMIGEdgeSeq, VMIGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                        Case "VTD"
                            CabCode = CutlistForm.VanityCodeBox.Text & CutlistForm.OptionsVar & "-" & Hardware
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(VMIBackQ, "Back", VMIBackY, VMIBackX, VMIBackZ, VMIBackM)
                            DGV.Rows.Add(VMITopStrapQ, "Top Strap", VMITopStrapY, VMITopStrapX, VMITopStrapZ, VMITopStrapM, VMITSEdgeSeq, VMITSEdgeCode)
                            DGV.Rows.Add(VMIBtmQ, "Bottom", VMIBtmY, VMIBtmX, VMIBtmZ, VMIBtmM, VMIBEdgeSeq, VMIBEdgeCode)
                            DGV.Rows.Add(VMIStrapQ, "Strap", VMIStrapY, VMIStrapX, VMIStrapZ, VMIStrapM, VMISEdgeSeq, VMISEdgeCode)
                            DGV.Rows.Add(VMIGableQ, "Gable", VMIGableY, VMIGableX, VMIGableZ, VMIGableM, VMIGEdgeSeq, VMIGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                        Case "VTDD"
                            CabCode = CutlistForm.VanityCodeBox.Text & CutlistForm.OptionsVar & "-" & Hardware
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(VMIBackQ, "Back", VMIBackY, VMIBackX, VMIBackZ, VMIBackM)
                            DGV.Rows.Add(VMITopQ, "Top", VMITopY, VMITopX, VMITopZ, VMITopM, VMITEdgeSeq, VMITEdgeCode)
                            DGV.Rows.Add(VMIBtmQ, "Bottom", VMIBtmY, VMIBtmX, VMIBtmZ, VMIBtmM, VMIBEdgeSeq, VMIBEdgeCode)
                            DGV.Rows.Add(VMIStrapQ, "Strap", VMIStrapY, VMIStrapX, VMIStrapZ, VMIStrapM, VMISEdgeSeq, VMISEdgeCode)
                            DGV.Rows.Add(VMIDividerQ, "Divider", VMIDividerY, VMIDividerX, VMIDividerZ, VMIDividerM, VMIDEdgeSeq, VMIDEdgeCode)
                            DGV.Rows.Add(VMIGableQ, "Gable", VMIGableY, VMIGableX, VMIGableZ, VMIGableM, VMIGEdgeSeq, VMIGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                        Case Else
                    End Select
                Else
                    Select Case VanCase
                        Case "V"
                            CabCode = CutlistForm.VanityCodeBox.Text & "-"
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(VBackQ, "Back", VBackY, VBackX, VBackZ, VBackM)
                            DGV.Rows.Add(VTopStrapQ, "Top Strap", VTopStrapY, VTopStrapX, VTopStrapZ, VTopStrapM, VTSEdgeSeq, VTSEdgeCode)
                            DGV.Rows.Add(VBtmQ, "Bottom", VBtmY, VBtmX, VBtmZ, VBtmM, VBEdgeSeq, VBEdgeCode)
                            DGV.Rows.Add(VGableQ, "Gable", VGableY, VGableX, VGableZ, VGableM, VGEdgeSeq, VGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                        Case "VS"
                            CabCode = CutlistForm.VanityCodeBox.Text & "-"
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(VBackQ, "Back", VBackY, VBackX, VBackZ, VBackM)
                            DGV.Rows.Add(VTopStrapQ, "Top Strap", VTopStrapY, VTopStrapX, VTopStrapZ, VTopStrapM, VTSEdgeSeq, VTSEdgeCode)
                            DGV.Rows.Add(VBtmQ, "Bottom", VBtmY, VBtmX, VBtmZ, VBtmM, VBEdgeSeq, VBEdgeCode)
                            DGV.Rows.Add(VStrapQ, "Sink Strap", VStrapY, VStrapX, VStrapZ, VStrapM)
                            DGV.Rows.Add(VGableQ, "Gable", VGableY, VGableX, VGableZ, VGableM, VGEdgeSeq, VGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                        Case "V2D", "V3D"
                            CabCode = CutlistForm.VanityCodeBox.Text & "-" & Hardware
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(VBackQ, "Back", VBackY, VBackX, VBackZ, VBackM)
                            DGV.Rows.Add(VTopStrapQ, "Top Strap", VTopStrapY, VTopStrapX, VTopStrapZ, VTopStrapM, VTSEdgeSeq, VTSEdgeCode)
                            DGV.Rows.Add(VBtmQ, "Bottom", VBtmY, VBtmX, VBtmZ, VBtmM, VBEdgeSeq, VBEdgeCode)
                            DGV.Rows.Add(VGableQ, "Gable", VGableY, VGableX, VGableZ, VGableM, VGEdgeSeq, VGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                        Case "VBD"
                            CabCode = CutlistForm.VanityCodeBox.Text & "-" & Hardware
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(VBackQ, "Back", VBackY, VBackX, VBackZ, VBackM)
                            DGV.Rows.Add(VTopStrapQ, "Top Strap", VTopStrapY, VTopStrapX, VTopStrapZ, VTopStrapM, VTSEdgeSeq, VTSEdgeCode)
                            DGV.Rows.Add(VBtmQ, "Bottom", VBtmY, VBtmX, VBtmZ, VBtmM, VBEdgeSeq, VBEdgeCode)
                            DGV.Rows.Add(VFFShelfQ, "F.F.Sh", VFFShelfY, VFFShelfX, VFFShelfZ, VFFShelfM, VFFSHEdgeSeq, VFFSHEdgeCode)
                            DGV.Rows.Add(VGableQ, "Gable", VGableY, VGableX, VGableZ, VGableM, VGEdgeSeq, VGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                        Case "VTD"
                            CabCode = CutlistForm.VanityCodeBox.Text & "-" & Hardware
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(VBackQ, "Back", VBackY, VBackX, VBackZ, VBackM)
                            DGV.Rows.Add(VTopStrapQ, "Top Strap", VTopStrapY, VTopStrapX, VTopStrapZ, VTopStrapM, VTSEdgeSeq, VTSEdgeCode)
                            DGV.Rows.Add(VBtmQ, "Bottom", VBtmY, VBtmX, VBtmZ, VBtmM, VBEdgeSeq, VBEdgeCode)
                            DGV.Rows.Add(VStrapQ, "Strap", VStrapY, VStrapX, VStrapZ, VStrapM, VSEdgeSeq, VSEdgeCode)
                            DGV.Rows.Add(VGableQ, "Gable", VGableY, VGableX, VGableZ, VGableM, VGEdgeSeq, VGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                        Case "VTDD"
                            CabCode = CutlistForm.VanityCodeBox.Text & "-" & Hardware
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(VBackQ, "Back", VBackY, VBackX, VBackZ, VBackM)
                            DGV.Rows.Add(VTopQ, "Top", VTopY, VTopX, VTopZ, VTopM, VTEdgeSeq, VTEdgeCode)
                            DGV.Rows.Add(VBtmQ, "Bottom", VBtmY, VBtmX, VBtmZ, VBtmM, VBEdgeSeq, VBEdgeCode)
                            DGV.Rows.Add(VStrapQ, "Strap", VStrapY, VStrapX, VStrapZ, VStrapM, VSEdgeSeq, VSEdgeCode)
                            DGV.Rows.Add(VDividerQ, "Divider", VDividerY, VDividerX, VDividerZ, VDividerM, VDEdgeSeq, VDEdgeCode)
                            DGV.Rows.Add(VGableQ, "Gable", VGableY, VGableX, VGableZ, VGableM, VGEdgeSeq, VGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                        Case Else
                    End Select
                End If

            '#############################
            '# VANITY ELEVATED # 7 LINES #
            '#############################
            '
            Case "VANITY ELEVATED"
                If (CutlistForm.BackGrooveBox.Text = "3mm") Then
                    BackGrove = "⅛-"
                Else
                    BackGrove = "¼-"
                End If
                Dim VanCase = CutlistForm.VanityCodeBox2.Text
                If (CutlistForm.MatchingInteriorCheck.Checked) Then
                    Select Case VanCase
                        Case "VE"
                            CabCode = CutlistForm.VanityCodeBox2.Text & CutlistForm.OptionsVar & BackGrove
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(VEMIBackQ, "Back", VEMIBackY, VEMIBackX, VEMIBackZ, VEMIBackM)
                            DGV.Rows.Add(VEMITopStrapQ, "Top Strap", VEMITopStrapY, VEMITopStrapX, VEMITopStrapZ, VEMITopStrapM, VEMITSEdgeSeq, VEMITSEdgeCode)
                            DGV.Rows.Add(VEMIBtmQ, "Bottom", VEMIBtmY, VEMIBtmX, VEMIBtmZ, VEMIBtmM, VEMIBEdgeSeq, VEMIBEdgeCode)
                            DGV.Rows.Add(VEMIGableQ, "Gable", VEMIGableY, VEMIGableX, VEMIGableZ, VEMIGableM, VEMIGEdgeSeq, VEMIGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")

                        Case "VES"
                            CabCode = CutlistForm.VanityCodeBox2.Text & CutlistForm.OptionsVar & BackGrove
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(VEMIBackQ, "Back", VEMIBackY, VEMIBackX, VEMIBackZ, VEMIBackM)
                            DGV.Rows.Add(VEMITopStrapQ, "Top Strap", VEMITopStrapY, VEMITopStrapX, VEMITopStrapZ, VEMITopStrapM, VEMITSEdgeSeq, VEMITSEdgeCode)
                            DGV.Rows.Add(VEMIBtmQ, "Bottom", VEMIBtmY, VEMIBtmX, VEMIBtmZ, VEMIBtmM, VEMIBEdgeSeq, VEMIBEdgeCode)
                            DGV.Rows.Add(VEMIStrapQ, "Sink Strap", VEMIStrapY, VEMIStrapX, VEMIStrapZ, VEMIStrapM)
                            DGV.Rows.Add(VEMIGableQ, "Gable", VEMIGableY, VEMIGableX, VEMIGableZ, VEMIGableM, VEMIGEdgeSeq, VEMIGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")

                        Case "VE2D", "VE3D"
                            CabCode = CutlistForm.VanityCodeBox2.Text & CutlistForm.OptionsVar & "-" & Hardware & BackGrove
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(VEMIBackQ, "Back", VEMIBackY, VEMIBackX, VEMIBackZ, VEMIBackM)
                            DGV.Rows.Add(VEMITopStrapQ, "Top Strap", VEMITopStrapY, VEMITopStrapX, VEMITopStrapZ, VEMITopStrapM, VEMITSEdgeSeq, VEMITSEdgeCode)
                            DGV.Rows.Add(VEMIBtmQ, "Bottom", VEMIBtmY, VEMIBtmX, VEMIBtmZ, VEMIBtmM, VEMIBEdgeSeq, VEMIBEdgeCode)
                            DGV.Rows.Add(VEMIGableQ, "Gable", VEMIGableY, VEMIGableX, VEMIGableZ, VEMIGableM, VEMIGEdgeSeq, VEMIGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")

                        Case "VEBD"
                            CabCode = CutlistForm.VanityCodeBox2.Text & CutlistForm.OptionsVar & "-" & Hardware & BackGrove
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(VEMIBackQ, "Back", VEMIBackY, VEMIBackX, VEMIBackZ, VEMIBackM)
                            DGV.Rows.Add(VEMITopStrapQ, "Top Strap", VEMITopStrapY, VEMITopStrapX, VEMITopStrapZ, VEMITopStrapM, VEMITSEdgeSeq, VEMITSEdgeCode)
                            DGV.Rows.Add(VEMIBtmQ, "Bottom", VEMIBtmY, VEMIBtmX, VEMIBtmZ, VEMIBtmM, VEMIBEdgeSeq, VEMIBEdgeCode)
                            DGV.Rows.Add(VEMIFFShelfQ, "F.F.Sh", VEMIFFShelfY, VEMIFFShelfX, VEMIFFShelfZ, VEMIFFShelfM, VEMIFFSHEdgeSeq, VEMIFFSHEdgeCode)
                            DGV.Rows.Add(VEMIGableQ, "Gable", VEMIGableY, VEMIGableX, VEMIGableZ, VEMIGableM, VEMIGEdgeSeq, VEMIGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")

                        Case "VETD"
                            CabCode = CutlistForm.VanityCodeBox2.Text & CutlistForm.OptionsVar & "-" & Hardware & BackGrove
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(VEMIBackQ, "Back", VEMIBackY, VEMIBackX, VEMIBackZ, VEMIBackM)
                            DGV.Rows.Add(VEMITopStrapQ, "Top Strap", VEMITopStrapY, VEMITopStrapX, VEMITopStrapZ, VEMITopStrapM, VEMITSEdgeSeq, VEMITSEdgeCode)
                            DGV.Rows.Add(VEMIBtmQ, "Bottom", VEMIBtmY, VEMIBtmX, VEMIBtmZ, VEMIBtmM, VEMIBEdgeSeq, VEMIBEdgeCode)
                            DGV.Rows.Add(VEMIStrapQ, "Strap", VEMIStrapY, VEMIStrapX, VEMIStrapZ, VEMIStrapM, VEMISEdgeSeq, VEMISEdgeCode)
                            DGV.Rows.Add(VEMIGableQ, "Gable", VEMIGableY, VEMIGableX, VEMIGableZ, VEMIGableM, VEMIGEdgeSeq, VEMIGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")

                        Case "VETDD"
                            CabCode = CutlistForm.VanityCodeBox2.Text & CutlistForm.OptionsVar & "-" & Hardware & BackGrove
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(VEMIBackQ, "Back", VEMIBackY, VEMIBackX, VEMIBackZ, VEMIBackM)
                            DGV.Rows.Add(VEMITopQ, "Top", VEMITopY, VEMITopX, VEMITopZ, VEMITopM, VEMITEdgeSeq, VEMITEdgeCode)
                            DGV.Rows.Add(VEMIBtmQ, "Bottom", VEMIBtmY, VEMIBtmX, VEMIBtmZ, VEMIBtmM, VEMIBEdgeSeq, VEMIBEdgeCode)
                            DGV.Rows.Add(VEMIStrapQ, "Strap", VEMIStrapY, VEMIStrapX, VEMIStrapZ, VEMIStrapM, VEMISEdgeSeq, VEMISEdgeCode)
                            DGV.Rows.Add(VEMIDividerQ, "Divider", VEMIDividerY, VEMIDividerX, VEMIDividerZ, VEMIDividerM, VEMIDEdgeSeq, VEMIDEdgeCode)
                            DGV.Rows.Add(VEMIGableQ, "Gable", VEMIGableY, VEMIGableX, VEMIGableZ, VEMIGableM, VEMIGEdgeSeq, VEMIGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                        Case Else
                    End Select
                Else
                    Select Case VanCase
                        Case "VE"
                            CabCode = CutlistForm.VanityCodeBox2.Text & CutlistForm.OptionsVar & BackGrove
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(VEBackQ, "Back", VEBackY, VEBackX, VEBackZ, VEBackM)
                            DGV.Rows.Add(VETopStrapQ, "Top Strap", VETopStrapY, VETopStrapX, VETopStrapZ, VETopStrapM, VETSEdgeSeq, VETSEdgeCode)
                            DGV.Rows.Add(VEBtmQ, "Bottom", VEBtmY, VEBtmX, VEBtmZ, VEBtmM, VEBEdgeSeq, VEBEdgeCode)
                            DGV.Rows.Add(VEGableQ, "Gable", VEGableY, VEGableX, VEGableZ, VEGableM, VEGEdgeSeq, VEGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                        Case "VES"
                            CabCode = CutlistForm.VanityCodeBox2.Text & CutlistForm.OptionsVar & BackGrove
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(VEBackQ, "Back", VEBackY, VEBackX, VEBackZ, VEBackM)
                            DGV.Rows.Add(VETopStrapQ, "Top Strap", VETopStrapY, VETopStrapX, VETopStrapZ, VETopStrapM, VETSEdgeSeq, VETSEdgeCode)
                            DGV.Rows.Add(VEBtmQ, "Bottom", VEBtmY, VEBtmX, VEBtmZ, VEBtmM, VEBEdgeSeq, VEBEdgeCode)
                            DGV.Rows.Add(VEStrapQ, "Sink Strap", VEStrapY, VEStrapX, VEStrapZ, VEStrapM)
                            DGV.Rows.Add(VEGableQ, "Gable", VEGableY, VEGableX, VEGableZ, VEGableM, VEGEdgeSeq, VEGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                        Case "VE2D", "VE3D"
                            CabCode = CutlistForm.VanityCodeBox2.Text & CutlistForm.OptionsVar & "-" & Hardware & BackGrove
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(VEBackQ, "Back", VEBackY, VEBackX, VEBackZ, VEBackM)
                            DGV.Rows.Add(VETopStrapQ, "Top Strap", VETopStrapY, VETopStrapX, VETopStrapZ, VETopStrapM, VETSEdgeSeq, VETSEdgeCode)
                            DGV.Rows.Add(VEBtmQ, "Bottom", VEBtmY, VEBtmX, VEBtmZ, VEBtmM, VEBEdgeSeq, VEBEdgeCode)
                            DGV.Rows.Add(VEGableQ, "Gable", VEGableY, VEGableX, VEGableZ, VEGableM, VEGEdgeSeq, VEGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                        Case "VEBD"
                            CabCode = CutlistForm.VanityCodeBox2.Text & CutlistForm.OptionsVar & "-" & Hardware & BackGrove
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(VEBackQ, "Back", VEBackY, VEBackX, VEBackZ, VEBackM)
                            DGV.Rows.Add(VETopStrapQ, "Top Strap", VETopStrapY, VETopStrapX, VETopStrapZ, VETopStrapM, VETSEdgeSeq, VETSEdgeCode)
                            DGV.Rows.Add(VEBtmQ, "Bottom", VEBtmY, VEBtmX, VEBtmZ, VEBtmM, VEBEdgeSeq, VEBEdgeCode)
                            DGV.Rows.Add(VEFFShelfQ, "F.F.Sh", VEFFShelfY, VEFFShelfX, VEFFShelfZ, VEFFShelfM, VEFFSHEdgeSeq, VEFFSHEdgeCode)
                            DGV.Rows.Add(VEGableQ, "Gable", VEGableY, VEGableX, VEGableZ, VEGableM, VEGEdgeSeq, VEGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                        Case "VETD"
                            CabCode = CutlistForm.VanityCodeBox2.Text & CutlistForm.OptionsVar & "-" & Hardware & BackGrove
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(VEBackQ, "Back", VEBackY, VEBackX, VEBackZ, VEBackM)
                            DGV.Rows.Add(VETopStrapQ, "Top Strap", VETopStrapY, VETopStrapX, VETopStrapZ, VETopStrapM, VETSEdgeSeq, VETSEdgeCode)
                            DGV.Rows.Add(VEBtmQ, "Bottom", VEBtmY, VEBtmX, VEBtmZ, VEBtmM, VEBEdgeSeq, VEBEdgeCode)
                            DGV.Rows.Add(VEStrapQ, "Strap", VEStrapY, VEStrapX, VEStrapZ, VEStrapM, VESEdgeSeq, VESEdgeCode)
                            DGV.Rows.Add(VEGableQ, "Gable", VEGableY, VEGableX, VEGableZ, VEGableM, VEGEdgeSeq, VEGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                        Case "VETDD"
                            CabCode = CutlistForm.VanityCodeBox2.Text & CutlistForm.OptionsVar & "-" & Hardware & BackGrove
                            DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                            DGV.Rows.Add(CabSize1)
                            DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                            DGV.Rows.Add(VEBackQ, "Back", VEBackY, VEBackX, VEBackZ, VEBackM)
                            DGV.Rows.Add(VETopQ, "Top", VETopY, VETopX, VETopZ, VETopM, VETEdgeSeq, VETEdgeCode)
                            DGV.Rows.Add(VEBtmQ, "Bottom", VEBtmY, VEBtmX, VEBtmZ, VEBtmM, VEBEdgeSeq, VEBEdgeCode)
                            DGV.Rows.Add(VEStrapQ, "Strap", VEStrapY, VEStrapX, VEStrapZ, VEStrapM, VESEdgeSeq, VESEdgeCode)
                            DGV.Rows.Add(VEDividerQ, "Divider", VEDividerY, VEDividerX, VEDividerZ, VEDividerM, VEDEdgeSeq, VEDEdgeCode)
                            DGV.Rows.Add(VEGableQ, "Gable", VEGableY, VEGableX, VEGableZ, VEGableM, VEGEdgeSeq, VEGEdgeCode)
                            DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                        Case Else
                    End Select
                End If

            '############################
            '# CANOPY SINGLE # 11 LINES #
            '############################
            '
            Case "CANOPY HOOD"
                CabCode = "HOOD"
                DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                DGV.Rows.Add(CabSize1)
                DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                DGV.Rows.Add(RHBackQ, "Back", RHBackY, RHBackX2, RHBackZ, RHBackM, RHBEdgeSeq, RHBEdgeCode)
                DGV.Rows.Add(RHTopQ, "Top", RHTopY, RHTopX, RHTopZ, RHTopM, RHTEdgeSeq, RHTEdgeCode)
                DGV.Rows.Add(RHGableQ, "Gable", RHGableY, RHGableX, RHGableZ, RHGableM, RHGEdgeSeq, RHGEdgeCode, RHGEdgeSeq2, RHGEdgeCode2)
                DGV.Rows.Add(RHFanSHQ, "Fan Shelf", RHFanSHY, RHFanSHX, RHFanSHZ, RHFanSHM)
                DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")

            '############################
            '# CANOPY SINGLE # 11 LINES #
            '############################
            '
            Case "CANOPY SINGLE"
                Dim CanopyCode As String = ""
                CanopyCode = CutlistForm.CanopyCodeBox.Text & CutlistForm.MoldCodeBox.Text
                CabCode = CanopyCode & "-S"
                DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                DGV.Rows.Add(CabSize1)
                DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                DGV.Rows.Add(RHBackQ, "Back", RHBackY, RHBackX, RHBackZ, RHBackM, RHBEdgeSeq, RHBEdgeCode)
                DGV.Rows.Add(RHTopQ, "Top", RHTopY, RHTopX, RHTopZ, RHTopM, RHTEdgeSeq, RHTEdgeCode)
                DGV.Rows.Add(RHGableQ, "Gable", RHGableY, RHGableX, RHGableZ, RHGableM, RHGEdgeSeq, RHGEdgeCode, RHGEdgeSeq2, RHGEdgeCode2)
                DGV.Rows.Add(RHFrontQ, "Front", RHFrontY, RHFrontX, RHFrontZ, RHFrontM, RHFEdgeSeq, RHFEdgeCode)
                DGV.Rows.Add(RHFanSHQ, "Fan Shelf", RHFanSHY, RHFanSHX, RHFanSHZ, RHFanSHM)
                DGV.Rows.Add(RHSFrontQ, "S-Front", RHSFrontY, RHSFrontX, RHSFrontZ, RHSFrontM)
                DGV.Rows.Add(RHSTopQ, "S-Top", RHSTopY, RHSTopX, RHSTopZ, RHSTopM)
                DGV.Rows.Add(RHIPanelQ, "IPanel", RHIPanelY, RHIPanelX, RHIPanelZ, RHIPanelM)
                DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")

            '############################
            '# CANOPY DOUBLE # 11 LINES #
            '############################
            '
            Case "CANOPY DOUBLE"
                Dim CanopyCode As String = ""
                CanopyCode = CutlistForm.CanopyCodeBox.Text & CutlistForm.MoldCodeBox.Text
                CabCode = CanopyCode & "-D"
                DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                DGV.Rows.Add(CabSize1)
                DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                DGV.Rows.Add(RHBackQ, "Back", RHBackY, RHBackX, RHBackZ, RHBackM, RHBEdgeSeq, RHBEdgeCode)
                DGV.Rows.Add(RHTopQ, "Top", RHTopY, RHTopX, RHTopZ, RHTopM, RHTEdgeSeq, RHTEdgeCode)
                DGV.Rows.Add(RHGableQ, "Gable", RHGableY, RHGableX, RHGableZ, RHGableM, RHGEdgeSeq, RHGEdgeCode, RHGEdgeSeq2, RHGEdgeCode2)
                DGV.Rows.Add(RHFrontQ, "Front", RHFrontY, RHFrontX, RHFrontZ, RHFrontM, RHFEdgeSeq, RHFEdgeCode)
                DGV.Rows.Add(RHFanSHQ, "Fan Shelf", RHFanSHY, RHFanSHX, RHFanSHZ, RHFanSHM)
                DGV.Rows.Add(RHSFrontQ, "S-Front", RHSFrontY, RHSFrontX, RHSFrontZ, RHSFrontM)
                DGV.Rows.Add(RHSTopQ, "S-Top", RHSTopY, RHSTopX, RHSTopZ, RHSTopM)
                DGV.Rows.Add(RHIPanelQ, "IPanel", RHIPanelY, RHIPanelX, RHIPanelZ, RHIPanelM)
                DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")

            '############################
            '# CANOPY TRIPLE # 12 LINES #
            '############################
            '
            Case "CANOPY TRIPLE"
                Dim CanopyCode As String = ""
                CanopyCode = CutlistForm.CanopyCodeBox.Text & CutlistForm.MoldCodeBox.Text
                CabCode = CanopyCode & "-T"
                DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                DGV.Rows.Add(CabSize1)
                DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                DGV.Rows.Add(RHBackQ, "Back", RHBackY, RHBackX, RHBackZ, RHBackM, RHBEdgeSeq, RHBEdgeCode)
                DGV.Rows.Add(RHTopQ, "Top", RHTopY, RHTopX, RHTopZ, RHTopM, RHTEdgeSeq, RHTEdgeCode)
                DGV.Rows.Add(RHGableQ, "Gable", RHGableY, RHGableX, RHGableZ, RHGableM, RHGEdgeSeq, RHGEdgeCode, RHGEdgeSeq2, RHGEdgeCode2)
                DGV.Rows.Add(RHFrontQ, "Front", RHFrontY, RHFrontX, RHFrontZ, RHFrontM, RHFEdgeSeq, RHFEdgeCode)
                DGV.Rows.Add(RHFanSHQ, "Fan Shelf", RHFanSHY, RHFanSHX, RHFanSHZ, RHFanSHM)
                DGV.Rows.Add(RHSFrontQ, "S-Front", RHSFrontY, RHSFrontX, RHSFrontZ, RHSFrontM)
                DGV.Rows.Add(RHSTopQ, "S-Top", RHSTopY, RHSTopX, RHSTopZ, RHSTopM)
                DGV.Rows.Add(RHIPanelQ, "IPanel", RHIPanelY, RHIPanelX, RHIPanelZ, RHIPanelM)
                DGV.Rows.Add(RHIPanel2Q, "IPanel", RHIPanel2Y, RHIPanel2X, RHIPanel2Z, RHIPanelM)
                DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")

            '########################################
            '# BASE INTEGRATED APPLIANCE # 17 LINES #
            '########################################
            '
            Case "BASE INTEGRATED APPLIANCE"
                Dim CH1 = (BIAAGableX + 1) / 10
                Dim CH2 = (BIABGableX) / 10
                Dim TH = CH1 + CH2
                Dim CabCode2 = "AP"
                If (CutlistForm.BaseCabCodeBox.Text = "B" Or CutlistForm.BaseCabCodeBox.Text = "BS") Then
                    CabCode = "BIA" & CutlistForm.ExtCode
                    DGV.Rows.Add(CabCode)
                    DGV.Rows.Add(VarWidthI & "-" & TH & "-" & VarDepthI)
                    DGV.Rows.Add(CabCode2, "Qty", VarAmountI)
                    DGV.Rows.Add(VarWidthI & "-" & CH1 & "-" & VarDepthI)
                    DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                    DGV.Rows.Add(BIAATopStrapQ, "AP-TopStrap", BIAATopStrapY, BIAATopStrapX, BIAATopStrapZ, BIAATopStrapM, BIAATSEdgeSeq, BIAATSEdgeCode)
                    DGV.Rows.Add(BIAABottomQ, "AP-Bottom", BIAABottomY, BIAABottomX, BIAABottomZ, BIAABottomM, BIAABEdgeSeq, BIAABEdgeCode)
                    DGV.Rows.Add(BIAAGableQ, "AP-Gable", BIAAGableY, BIAAGableX, BIAAGableZ, BIAAGableM, BIAAGEdgeSeq, BIAAGEdgeCode, BIAAGEdgeSeq2, BIAAGEdgeCode2)
                    DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                    CabCode2 = CutlistForm.BaseCabCodeBox.Text
                    DGV.Rows.Add(CabCode)
                    DGV.Rows.Add(VarWidthI & "-" & TH & "-" & VarDepthI)
                    DGV.Rows.Add(CabCode2, "Qty", VarAmountI)
                    DGV.Rows.Add(VarWidthI & "-" & CH2 & "-" & VarDepthI)
                    DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                    DGV.Rows.Add(BIABBackQ, "B-Back", BIABBackY, BIABBackX, BIABBackZ, BIABBackM)
                    DGV.Rows.Add(BIABTopStrapQ, "B-TopStrap", BIABTopStrapY, BIABTopStrapX, BIABTopStrapZ, BIABTopStrapM, BIABTSEdgeSeq, BIABTSEdgeCode)
                    DGV.Rows.Add(BIABBottomQ, "B-Bottom", BIABBottomY, BIABBottomX, BIABBottomZ, BIABBottomM, BIABBEdgeSeq, BIABBEdgeCode)
                    DGV.Rows.Add(BIABGableQ, "B-Gable", BIABGableY, BIABGableX, BIABGableZ, BIABGableM, BIABGEdgeSeq, BIABGEdgeCode)
                    DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                Else
                    CabCode = "BIA" & CutlistForm.ExtCode & Hardware
                    DGV.Rows.Add(CabCode)
                    DGV.Rows.Add(VarWidthI & "-" & TH & "-" & VarDepthI)
                    DGV.Rows.Add(CabCode2, "Qty", VarAmountI)
                    DGV.Rows.Add(VarWidthI & "-" & CH1 & "-" & VarDepthI)
                    DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                    DGV.Rows.Add(BIAATopStrapQ, "AP-TopStrap", BIAATopStrapY, BIAATopStrapX, BIAATopStrapZ, BIAATopStrapM, BIAATSEdgeSeq, BIAATSEdgeCode)
                    DGV.Rows.Add(BIAABottomQ, "AP-Bottom", BIAABottomY, BIAABottomX, BIAABottomZ, BIAABottomM, BIAABEdgeSeq, BIAABEdgeCode)
                    DGV.Rows.Add(BIAAGableQ, "AP-Gable", BIAAGableY, BIAAGableX, BIAAGableZ, BIAAGableM, BIAAGEdgeSeq, BIAAGEdgeCode, BIAAGEdgeSeq2, BIAAGEdgeCode2)
                    DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                    CabCode2 = CutlistForm.BaseCabCodeBox.Text & "-" & Hardware
                    DGV.Rows.Add(CabCode)
                    DGV.Rows.Add(VarWidthI & "-" & TH & "-" & VarDepthI)
                    DGV.Rows.Add(CabCode2, "Qty", VarAmountI)
                    DGV.Rows.Add(VarWidthI & "-" & CH2 & "-" & VarDepthI)
                    DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                    DGV.Rows.Add(BIABBackQ, "B-Back", BIABBackY, BIABBackX, BIABBackZ, BIABBackM)
                    DGV.Rows.Add(BIABTopStrapQ, "B-TopStrap", BIABTopStrapY, BIABTopStrapX, BIABTopStrapZ, BIABTopStrapM, BIABTSEdgeSeq, BIABTSEdgeCode)
                    DGV.Rows.Add(BIABBottomQ, "B-Bottom", BIABBottomY, BIABBottomX, BIABBottomZ, BIABBottomM, BIABBEdgeSeq, BIABBEdgeCode)
                    DGV.Rows.Add(BIABGableQ, "B-Gable", BIABGableY, BIABGableX, BIABGableZ, BIABGableM, BIABGEdgeSeq, BIABGEdgeCode)
                    DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                End If


            '############################
            '# FLOATING SHELF # 5 LINES #
            '############################
            '
            Case "FLOATING SHELF"
                CabCode = "FLOATING SHELF"
                DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                DGV.Rows.Add(CabSize1)
                DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                DGV.Rows.Add(FShelfQ, "Panel", FShelfY, FShelfX, FShelfZ, FShelfM)
                DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")

            '##########################
            '# WINDOW PANEL # 5 LINES #
            '##########################
            '
            Case "WINDOW PANEL"
                CabCode = "WINDOW PANEL"
                DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                DGV.Rows.Add(CabSize1)
                DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                DGV.Rows.Add(WPanelQ, "Panel", WPanelY, WPanelX, WPanelZ, WPanelM)
                DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")

            '##########################
            '# WINDOW BOX # 6 LINES #
            '##########################
            '
            Case "WINDOW BOX"
                CabCode = "WINDOW BOX"
                DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                DGV.Rows.Add(CabSize1)
                DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                DGV.Rows.Add(WTPanelQ, "Top Panel", WTPanelY, WTPanelX, WTPanelZ, WTPanelM)
                DGV.Rows.Add(WFrontQ, "Front", WFrontY, WFrontX, WFrontZ, WFrontM)
                DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")

            '############################
            '# PLUMBING COVER # 5 LINES #
            '############################
            '
            Case "PLUMBING COVER"
                CabCode = "PLUMBING COVER"
                DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                DGV.Rows.Add(CabSize1)
                DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                DGV.Rows.Add(PCPanelQ, "Panel", PCPanelY, PCPanelX, PCPanelZ, PCPanelM)
                DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")

            '##########################
            '# WINDOW PANEL # 6 LINES #
            '##########################
            '
            Case "SUPPORT BOX"
                CabCode = "SUPPORT BOX"
                DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                DGV.Rows.Add(CabSize1)
                DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                DGV.Rows.Add(SBGableQ, "Gable", SBGableY, SBGableX, SBGableZ, SBGableM)
                DGV.Rows.Add(SBRailQ, "Rails", SBRailY, SBRailX, SBRailZ, SBRailM)
                DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")

            '##############################
            '# SPECIAL CABINETS # 0 LINES #
            '##############################
            '
            Case "BASE MICROWAVE OPEN 1 DRAWER", "TALL INTEGRATED APPLIANCE 3 UNIT"
                If (CutlistForm.SPCheckBox.Checked = False) Then
                    Select Case MsgBox("This Cabinet Requires Special Properties! Would You Like To Enable Special Checkbox?", MsgBoxStyle.YesNo, "ATTENTION")
                        Case MsgBoxResult.Yes
                            CutlistForm.SPCheckBox.Checked = True
                            Exit Sub
                        Case MsgBoxResult.No
                            Exit Sub
                    End Select
                End If

            '################################
            '# DESK SUPPORT GABLE # 5 LINES #
            '################################
            '
            Case "DESK SUPPORT GABLE"
                CabCode = "SUPPORT GABLE"
                DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                DGV.Rows.Add(CabSize1)
                DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                DGV.Rows.Add(DSGableQ, "Panel", DSGableY, DSGableX, DSGableZ, DSGableM)
                DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")

            '#####################################
            '# FANCY BRACKET (MATTAMY) # 5 LINES #
            '#####################################
            '
            Case "FANCY BRACKET"
                CabCode = "FANCY BRACKET"
                DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                DGV.Rows.Add("19.5-19.5-3.2")
                DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                DGV.Rows.Add(FBracketQ, "Panel", FBracketY, FBracketX, FBracketZ, FBracketM)
                DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")

            '#####################################
            '# FANCY BRACKET (MATTAMY) # 5 LINES #
            '#####################################
            '
            Case "FANCY VALANCE", "FUNITURE VALANCE"
                If (CutlistForm.CabCodeBox1.Text = "FANCY VALANCE") Then
                    CabCode = "FANCY VALANCE"
                    DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                    DGV.Rows.Add(CabSize1)
                    DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                    DGV.Rows.Add(FValanceQ, "Panel", FValanceY, FValanceX, FValanceZ, FValanceM)
                    DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                Else
                    CabCode = "FURNITURE VALANCE"
                    DGV.Rows.Add(CabCode, "Qty", VarAmountI)
                    DGV.Rows.Add(CabSize1)
                    DGV.Rows.Add("Qty", "Part", "W", "H", "D", "Material", "Edge", "Code", "Edge", "Code", "Notes")
                    DGV.Rows.Add(FValanceQ, "Panel", FValanceY, FValanceX, FValanceZ, FValanceM)
                    DGV.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "")
                End If

                '################################
                '# NEW CABINET SLOT #           #
                '################################
                '
            Case Else
        End Select

        '###################
        '# CLEAR TEXTBOXES #
        '###################
        '
        '################################
        '# CABINET MAIN CLASS BOX CLEAR #
        '################################
        '
        CutlistForm.CabCodeBox1.Text = ""

        '###############################
        '# CABINET SUB CLASS BOX CLEAR #
        '###############################
        '

        CutlistForm.ResetBooleanFields()

        '##################
        '# DISPOSE IMAGES #
        '##################
        '
        CutlistForm.PictureBox1.Image.Dispose()
        CutlistForm.PictureBox1.Image = My.Resources.None

        '##########################
        '# COPY PROGRAM DIRECTORY #
        '##########################
        '
        ClassFunctions.CopyDirectory(CutlistForm.lblCabCode.Text, CutlistForm.lblCabCode2.Text)

        '##########################################
        '# CLEAR AND RESET CABINET LABELS TO NONE #
        '##########################################
        '
        CutlistForm.lblCabCode2.Text = "NONE"
        CutlistForm.ClearAllFields()
    End Sub

End Class
'12028 LINES