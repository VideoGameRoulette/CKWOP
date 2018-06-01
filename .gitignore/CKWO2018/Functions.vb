Imports System.IO

Public Class Functions

    Public Shared Function NewWorkOrder()
        Form1.DataGridView1.Rows.Clear()
        Form1.WONum.Text = "0000"
        Form1.CabCode.Text = ""
        Form1.widthBox.Text = 0
        Form1.heightBox.Text = 0
        Form1.heightBox2.Text = 0
        Form1.heightBox3.Text = 0
        Form1.depthBox.Text = 0
        Form1.depthBox2.Text = 0
        Form1.depthBox3.Text = 0
        Form1.amountBox.Text = 0
        Form1.CabCode.Text = "None"
        Form1.CabCode2.Text = "None"
        Form1.CabCode3.Text = "None"
        Form1.roomBox.Text = ""
        Form1.doorStyleBox.Text = ""
        Form1.doorFinishBox.Text = ""
        Form1.speciesBox.Text = ""
        Form1.groupBox.Text = ""
        Form1.materialBox.Text = ""
        Form1.hardwareBox.Text = ""
        NewWorkOrder = "New Work Order Initiallized"
    End Function

    Public Shared Function CopyDirectory()
        '#######################################
        '# CABINET CODE VARIABLE TO SEARCH FOR #
        '#######################################
        '
        Dim Height1 As Double = Form1.heightBox.Text
        Dim Height2 As Double = Form1.heightBox2.Text
        Dim Height3 As Double = Form1.heightBox3.Text
        Dim TotalHeight As Double = Height1 + Height2 + Height3

        Dim CabCodeDir As String = Form1.CabCode.Text & "-" & Form1.widthBox.Text & "-" & TotalHeight & "-" & Form1.depthBox.Text ' STOCK PROGRAM FOLDER NAME
        Dim CabCodeDir2 As String = Form1.CabCode.Text & "-TEMPLATE" ' TEMPLATE FOLDER NAME
        Dim TemplateDir As String = Form1.CabCode.Text & "-" & Form1.widthBox.Text & "-" & TotalHeight & "-" & Form1.depthBox.Text & "-TEMPLATE" ' RENAME CABINET CODE VARIABLE OUTPUT FOR TEMPLATE FOLDERS

        'Dim CabCodeDir As String = Form1.lblOutput.Text ' STOCK PROGRAM FOLDER NAME
        'Dim CabCodeDir2 As String = Form1.lblOutput2.Text ' TEMPLATE FOLDER NAME
        'Dim TemplateDir As String = Form1.lblOutput3.Text ' RENAME CABINET CODE VARIABLE OUTPUT FOR TEMPLATE FOLDERS

        '########################################
        '# SOURCE PATH OF CABINET CODE VARIABLE #
        '########################################
        '
        Dim SourcePath As String = Form1.PrgmPath & CabCodeDir & "\" 'CABINET PROGRAM DIRECTORY
        Dim SourcePath2 As String = Form1.TempPath & CabCodeDir2 & "\" 'CABINET TEMPLATE DIRECTORY

        '#############################################
        '# DESTINATION PATH OF CABINET CODE VARIABLE #
        '#############################################
        '
        Dim DestinationPath As String = Form1.WOPath & Form1.roomBox.Text

        '################################################
        '# GET DIRECTORY PATHS TO COPY CABINET PROGRAMS #
        '################################################
        '
        Dim newDirectory As String = System.IO.Path.Combine(DestinationPath, Path.GetFileName(Path.GetDirectoryName(SourcePath))) ' COPY STOCK PROGRAM OF CABINET TO WORK ORDER FOLDER 
        Dim newDirectory2 As String = System.IO.Path.Combine(DestinationPath, Path.GetFileName(Path.GetDirectoryName(SourcePath2))) ' COPY TEMPLATE PROGRAM OF CABINET TO WORK ORDER FOLDER 

        If Not (Directory.Exists(SourcePath)) Then
            '########################################################
            '# COPY TEMPLATE PROGRAM FOLDER TO WORK ORDER DIRECTORY #
            '########################################################
            '
            If Not Directory.Exists(newDirectory2) Then
                Directory.CreateDirectory(newDirectory2)
            End If
            Try
                Microsoft.VisualBasic.FileIO.FileSystem.CopyDirectory(SourcePath2, newDirectory2, True)
            Catch error9 As Exception
            End Try
            Try
                My.Computer.FileSystem.RenameDirectory(newDirectory2, TemplateDir)
            Catch error10 As Exception
                MsgBox(error10.ToString)
            End Try
        Else
            '#####################################################
            '# COPY STOCK PROGRAM FOLDER TO WORK ORDER DIRECTORY #
            '#####################################################
            '
            If Not (Directory.Exists(newDirectory)) Then
                Directory.CreateDirectory(newDirectory)
            End If
            Try
                Microsoft.VisualBasic.FileIO.FileSystem.CopyDirectory(SourcePath, newDirectory, True)
            Catch error10 As Exception
            End Try
        End If
        CopyDirectory = "Directory Copied."
    End Function

    Public Shared Function ClearFields()
        Form1.CabCode.Text = "None"
        Form1.CabCode2.Text = "None"
        Form1.CabCode3.Text = "None"
        Form1.widthBox.Text = 0
        Form1.heightBox.Text = 0
        Form1.heightBox2.Text = 0
        Form1.heightBox3.Text = 0
        Form1.depthBox.Text = 0
        Form1.depthBox2.Text = 0
        Form1.depthBox3.Text = 0
        Form1.amountBox.Text = 0
        ClearFields = "Fields Cleaered For Next Cabinet"
    End Function

End Class
