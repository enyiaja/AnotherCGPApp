Imports System
Imports System.IO
Imports System.Security
Imports System.Security.Cryptography

Public Class Form1
    Dim no As Integer = 0
    Public Structure Course
        Dim lblNo As Label
        Dim txtName As TextBox
        Dim cmbUnits As ComboBox
        Dim txtScore As TextBox
        Dim txtLetter As TextBox
        Dim cmbGrade As ComboBox
    End Structure
    Dim x As Integer = 2
    Dim y As Integer = 6
    Dim x1 As Integer = 33
    Dim x2 As Integer = 223
    Dim x3 As Integer = 274
    Dim x4 As Integer = 343
    Dim x5 As Integer = 399
    Dim Run(no) As Course


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim lblNolbl As Label = New Label
        lblNolbl.Text = "s/n"
        lblNolbl.Location = New Point(x + 12, -24)

        Dim txtNamelbl As Label = New Label
        txtNamelbl.Text = "Course Title"
        txtNamelbl.Location = New Point(x1 + 12, -24)

        Dim cmbUnitslbl As Label = New Label
        cmbUnitslbl.Text = "Course Units"
        cmbUnitslbl.Location = New Point(x2 + 12, -24)

        Dim txtScorelbl As Label = New Label
        txtScorelbl.Text = "Score"
        txtScorelbl.Location = New Point(x3 + 12, -24)

        Dim cmbGradelbl As Label = New Label
        cmbGradelbl.Text = "Score"
        cmbGradelbl.Location = New Point(x4 + 12, -24)

        Dim txtLetterlbl As Label = New Label
        txtLetterlbl.Text = "Grade"
        txtLetterlbl.Location = New Point(x5 + 12, -24)



        With Run(0)
            .lblNo = New Label
            With .lblNo
                .Location = New Point(x, -24)
            End With
            .txtName = New TextBox
            With .txtName
                .Location = New Point(x1, -24)
            End With
            .cmbUnits = New ComboBox
            With .cmbUnits
                .Location = New Point(x2, -24)
            End With
            .txtScore = New TextBox
            With .txtScore
                .Location = New Point(x3, -24)
            End With
            .cmbGrade = New ComboBox
            With .cmbGrade
                .Location = New Point(x4, -24)
            End With
            .txtLetter = New TextBox
            With .txtLetter
                .Location = New Point(x5, -24)
            End With
        End With
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub btnRemove_Click(sender As Object, e As EventArgs) Handles btnRemove.Click
        ' Button to Remove Last Course
        Panel1.Controls.Remove(Run(no).lblNo)
        Panel1.Controls.Remove(Run(no).txtName)
        Panel1.Controls.Remove(Run(no).cmbUnits)
        Panel1.Controls.Remove(Run(no).txtScore)
        Panel1.Controls.Remove(Run(no).cmbGrade)
        Panel1.Controls.Remove(Run(no).txtLetter)
        no -= 1
        txtNo.Text = no
        ReDim Preserve Run(UBound(Run) - 1)
        y -= 30

    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        'Button to add a course
        no += 1
        txtNo.Text = no
        'y += 30
        ReDim Preserve Run(no + 1)
        With Run(no)
            .lblNo = New Label
            With .lblNo
                .Text = no
                .Size = New Size(20, 13)

                .Location = New Point(x, Run(no - 1).lblNo.Location.Y + 30)
                .Parent = Panel1
            End With

            .txtName = New TextBox
            With .txtName
                .Size = New Size(174, 20)
                .Location = New Point(x1, Run(no - 1).txtName.Location.Y + 30)
                .Parent = Panel1
            End With

            .cmbUnits = New ComboBox
            With .cmbUnits
                For i = 1 To 6
                    .Items.Add(i)
                Next
                .Size = New Size(40, 21)
                .Location = New Point(x2, Run(no - 1).cmbUnits.Location.Y + 30)
                .Parent = Panel1
            End With

            .txtScore = New TextBox
            With .txtScore
                .Size = New Size(59, 20)
                .Location = New Point(x3, Run(no - 1).txtScore.Location.Y + 30)
                .Parent = Panel1
                AddHandler .TextChanged, AddressOf ScoreChanged
                AddHandler .CursorChanged, AddressOf ScoreChanged
            End With


            .cmbGrade = New ComboBox
            With .cmbGrade
                For i = 2 To 5 Step 0.5
                    .Items.Add(i)
                Next
                .Size = New Size(40, 21)
                .Location = New Point(x4, Run(no - 1).cmbGrade.Location.Y + 30)
                .Parent = Panel1
                'AddHandler .DropDownClosed, AddressOf GradeChanged
            End With


            .txtLetter = New TextBox
            With .txtLetter
                .ReadOnly = True
                .Size = New Size(33, 20)
                .Location = New Point(x5, Run(no - 1).txtLetter.Location.Y + 30)
                .Parent = Panel1
            End With
        End With
    End Sub

    Public Function LetterFill()
        For i = 0 To no
            With Run(i)
                If .txtName.Text <> "" Then
                    If .txtScore.Text <> "" And Val(.txtScore.Text) <= 100 And Val(.txtScore.Text) >= 0 Then
                        If Val(.txtScore.Text) >= 80 Or Val(.cmbGrade.Text) = 5 Then
                            .txtLetter.Text = "A"
                            .cmbGrade.Text = "5"
                        ElseIf Val(.cmbGrade.Text) = 4.5 Then
                            .txtLetter.Text = "B+"
                            .cmbGrade.Text = "4.5"
                        ElseIf Val(.txtScore.Text) >= 60 Or Val(.cmbGrade.Text) = 4 Then
                            .txtLetter.Text = "B"
                            .cmbGrade.Text = "4"
                        ElseIf Val(.cmbGrade.Text) = 3.5 Then
                            .txtLetter.Text = "C+"
                            .cmbGrade.Text = "3.5"
                        ElseIf Val(.txtScore.Text) >= 50 Or Val(.cmbGrade.Text) = 3 Then
                            .txtLetter.Text = "C"
                            .cmbGrade.Text = "3"
                        ElseIf Val(.txtScore.Text) >= 45 Then
                            .txtLetter.Text = "D"
                            .cmbGrade.Text = "2"
                        ElseIf Val(.txtScore.Text) >= 40 Then
                            .txtLetter.Text = "E"
                            .cmbGrade.Text = "1"
                        ElseIf Val(.txtScore.Text) >= 0 Then
                            .txtLetter.Text = "F"
                            .cmbGrade.Text = "0"
                        End If
                    Else
                        .txtScore.Clear()
                        .cmbGrade.Text = ""
                        'MsgBox("Value must be between 0 and 100", MsgBoxStyle.Exclamation, "Score Field Message")
                    End If
                End If
            End With
        Next
    End Function

    Private Sub ScoreChanged(sender As Object, e As EventArgs)
        For i = 0 To no
            With Run(i)
                If .txtScore.Text = "" And .cmbGrade.Text = "" Then
                    .txtLetter.Clear()
                End If
            End With
            LetterFill()
        Next
    End Sub


End Class
