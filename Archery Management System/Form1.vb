Option Explicit On
Imports System.IO
'Here are lots of the global variables used within the program.
'Each of these have a purpose and are used in multiple different subs.
Public Class frmScreen
    Dim sFileLoc As String = "G:\Computing\Coursework\Archery Management System"
    Dim clBow As String = " "
    Dim clGender As String = " "
    Dim logon As String
    Dim bRound As Boolean
    Dim sUser As String
    Dim sPass1 As String
    Dim sPass2 As String
    Dim bPassCorrect As Boolean = False
    Dim iValidScore As Boolean = False
    Dim aScore(12) As Integer
    Dim fFileName As String
    Dim iAge As Integer
    Dim bAge As Boolean = False
    Dim bSortRequest As Boolean = False
    Dim iTempScore As Integer = 0
    Dim bWorc As Boolean = False
    Dim f1 As String = 0
    Dim f2 As String = 0
    Dim f3 As String = 0
    Private fs As Object
    'Change to the New User screen
    Private Sub btnNew_Click(sender As Object, e As EventArgs) Handles btnNew.Click
        TabControl.SelectedIndex = 1
    End Sub
    'Change to the Existing User screen
    Private Sub btnExist_Click(sender As Object, e As EventArgs) Handles btnExist.Click
        TabControl.SelectedIndex = 2
    End Sub
    'Change from the New User to Existing User screen.
    Private Sub btnExist1_Click(sender As Object, e As EventArgs) Handles btnExist1.Click
        TabControl.SelectedIndex = 2
    End Sub
    'Change from the Existing User to the New User screen.
    Private Sub btnNUser_Click(sender As Object, e As EventArgs) Handles btnNUser.Click
        TabControl.SelectedIndex = 1
    End Sub
    'Set sUser to the text entered by the user. Establishes the text of the variable for use later.
    Private Sub txtUser_TextChanged(sender As Object, e As EventArgs) Handles txtUser.TextChanged
        sUser = txtUser.Text

    End Sub
    'Check the validity of the user's passwords and change the colour to show this.
    Private Sub txtPass2_TextChanged(sender As Object, e As EventArgs) Handles txtPass2.TextChanged
        sPass1 = txtPass1.Text
        sPass2 = txtPass2.Text
        If sPass1 = sPass2 And sPass1.Length > 7 Then
            txtPass1.BackColor = Color.Green
            txtPass2.BackColor = Color.Green
            Dim bPassCorrect As Boolean = True
        Else : txtPass1.BackColor = Color.Red
            txtPass2.BackColor = Color.Red
            Dim bPassCorrect As Boolean = False
        End If
    End Sub
    'Check the validity of the user's passwords and change the colour to show this. This is done for the first password.
    'This is repeated in case the user changes the first password after the second.
    Private Sub txtPass1_TextChanged(sender As Object, e As EventArgs) Handles txtPass1.TextChanged
        sPass1 = txtPass1.Text
        sPass2 = txtPass2.Text
        If sPass1 = sPass2 And sPass1.Length > 7 Then
            txtPass1.BackColor = Color.Green
            txtPass2.BackColor = Color.Green
            Dim bPassCorrect As Boolean = True
        Else : txtPass1.BackColor = Color.Red
            txtPass2.BackColor = Color.Red
            Dim bPassCorrect As Boolean = False
        End If
    End Sub
    'This forst section will use a set of nested selections and iterations to confirm that the password is valid.
    Private Sub btnCreate_Click_1(sender As Object, e As EventArgs) Handles btnCreate.Click
        Dim ValidPassword As Boolean = False
        If sPass1 = sPass2 And txtPass1.Text.Length > 7 Then
            Dim password As String = txtPass1.Text
            Dim contInt As Boolean = False
            Dim contAlph As Boolean = False
            Dim pos As Integer = 0
            For Each c As Char In password
                If IsNumeric(password(pos)) Then
                    contInt = True
                End If
                If Not IsNumeric(password.Chars(pos)) Then
                    contAlph = True
                End If
                pos = pos + 1
            Next
            If contInt = True And contAlph = True Then
                ValidPassword = True
            End If
        End If

        'This section will attempt to create a user profile if the password is valid and the username is not taken.
        If ValidPassword = True Then
            If txtUser.Text.Length > 3 Then ' This checks that the Username is longer than 3 characters.
                If My.Computer.FileSystem.FileExists(sFileLoc & "\Users\" & sUser & ".txt") Then
                    MsgBox("Username taken, please select another.")
                Else
                    File.Create(sFileLoc & "\Users\" & sUser & ".txt").Close() ' Creates a file for a profile and closes it.
                    Try ' Catch exceptions incase the program is unable to create an account.
                        My.Computer.FileSystem.WriteAllText(sFileLoc & "\Users\" & sUser & ".txt", sUser, True)
                        File.AppendAllText(sFileLoc & "\Users\" & sUser & ".txt", Environment.NewLine + sPass1)
                    Catch ex As System.IO.IOException
                        MsgBox("Unable to create file with your information")
                        Close()
                    End Try

                    TabControl.SelectedIndex = 3 ' Go to the Account Info screen.
                End If
                'These messages appear if one of the initial IF statements were false.
            Else
                MsgBox("Unable to create, check that the passwords match or the username is long enough.")
            End If
        Else
            MsgBox("Password must contain a letter, a number and be longer than 7 characters long.")
        End If

    End Sub

    'This takes the user's input and checks it to user records to load user information.
    Private Sub btnSignIn_Click(sender As Object, e As EventArgs) Handles btnSignIn.Click
        sPass1 = sExistPass.Text
        sUser = sExistUser.Text
        If My.Computer.FileSystem.FileExists(sFileLoc & "\Users\" & sUser & ".txt") Then 'Find the user's file
            sPass2 = System.IO.File.ReadAllLines(sFileLoc & "\Users\" & sUser & ".txt")(1) 'Find the user's Password
            'This checks the password entered for the user profile.
            If sPass1 = sPass2 Then
                'Retrieve the colour selection for the user
                cName = System.IO.File.ReadAllLines(sFileLoc & "\Users\" & sUser & ".txt")(7)
                bcColor = Color.FromName(cName)
                Me.BackColor = bcColor ' Applies the colour that the user has stored.
                TabControl.SelectedIndex = 4 'Go to the Home screen
                'Messages are displayed if any of the conditions are not met.
            Else
                MsgBox("Your password is invalid.")
            End If

        Else
            MsgBox("Invalid information, please check your information and try again.")

        End If
    End Sub
    Dim bcColor As Color
    Dim cName As String = "Orange"
    'Used to create the text that is stored in the user file to create the colour for the background.
    Sub Colour()
        If bcOrange.Checked = True Then 'Checks if this radio button is selected.
            cName = "Orange" 'Set the variable text to this string
        ElseIf bcRed.Checked = True Then
            cName = "Crimson"
        ElseIf bcBlue.Checked = True Then
            cName = "LightBlue"
        ElseIf bcYellow.Checked = True Then
            cName = "Gold"
        ElseIf bcGreen.Checked = True Then
            cName = "Green"
        End If
        bcColor = Color.FromName(cName) 'Sets the variable to this line of code.
    End Sub
    'This sub is responsible for storing the data entered by the user.
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Colour() ' Calls the subroutine above this, called Colour().
        'This condition checks that all data entered is present and valid.
        If txtName.Text.Length > 1 And clBow.Length > 1 And clGender.Length > 1 And sGNAS.Text.Length = 7 And txtAge.Text.Length > 1 Then
            My.Computer.FileSystem.DeleteFile(sFileLoc & "\Users\" & sUser & ".txt") 'Delete the user's file
            'These lines store all of the user's data to their file.
            My.Computer.FileSystem.WriteAllText(sFileLoc & "\Users\" & sUser & ".txt", sUser, True)
            File.AppendAllText(sFileLoc & "\Users\" & sUser & ".txt", Environment.NewLine + sPass1)
            File.AppendAllText(sFileLoc & "\Users\" & sUser & ".txt", Environment.NewLine + txtName.Text)
            File.AppendAllText(sFileLoc & "\Users\" & sUser & ".txt", Environment.NewLine + clBow)
            File.AppendAllText(sFileLoc & "\Users\" & sUser & ".txt", Environment.NewLine + sGNAS.Text)
            File.AppendAllText(sFileLoc & "\Users\" & sUser & ".txt", Environment.NewLine + clGender)
            File.AppendAllText(sFileLoc & "\Users\" & sUser & ".txt", Environment.NewLine + txtAge.Text)
            File.AppendAllText(sFileLoc & "\Users\" & sUser & ".txt", Environment.NewLine + cName)
            Me.BackColor = bcColor 'Applies the colour chosen by the user.
            TabControl.SelectedIndex = 4 'Go to the Home screen
        Else
            MsgBox("You must fill in all fields.") 'Displayed if the condition is not met.
        End If

    End Sub
    'Safely closes the program is the button is clicked.
    Private Sub btnQuit_Click(sender As Object, e As EventArgs) Handles btnQuit.Click
        Close()
    End Sub
    'Sets bRound to false if the round is a Portsmouth round.
    Private Sub cboxPort_CheckedChanged(sender As Object, e As EventArgs) Handles cboxPort.CheckedChanged
        cboxWorc.Checked = False 'Disables the Worcester box.
        bRound = False 'Sets the variable to False.
    End Sub
    'Sets bRound to True if the round is a Worcester round.
    Private Sub cboxWorc_CheckedChanged(sender As Object, e As EventArgs) Handles cboxWorc.CheckedChanged
        cboxPort.Checked = False 'Disables the Portsmouth round.
        bRound = True 'Sets the variable to False.
    End Sub
    'This subroutine enables the program to store the top 10 scores of the users in the correct file and order. 
    Sub Score()
        Dim fFile As String = sFileLoc & "\Scores\" & fFileName & ".txt"
        Dim aUser(10) As String
        For i As Integer = 0 To 9 'This will read all of the scores and users into arrays.
            aScore(i) = System.IO.File.ReadAllLines(fFile)(i)
            aUser(i) = System.IO.File.ReadAllLines(fFile)(i + 11)
        Next

        If aScore(9) < txtScore.Text Then 'The user's score is higher than the last if it should be on the list.
            MsgBox("You have a new high score")
            aScore(9) = txtScore.Text 'Last score is replaced by user's score.
            aUser(9) = sUser 'Last User is replaced by the User

            'This sorts the arrays into descending order.
            'Every swap of score swapps the user to keep them in order with their score.
            Dim temp As Integer
            Dim sTemp As String
            For x As Integer = 0 To 9
                For y As Integer = 0 To (9 - x)
                    If aScore(y) < aScore(y + 1) Then
                        temp = aScore(y + 1)
                        aScore(y + 1) = aScore(y)
                        aScore(y) = temp
                        sTemp = aUser(y + 1)
                        aUser(y + 1) = aUser(y)
                        aUser(y) = sTemp

                    End If
                Next
            Next

            File.WriteAllText(fFile, aScore(0)) 'Write the top score to the file.
            For i As Integer = 1 To 9 'Write the last 9 scores to the file.
                File.AppendAllText(fFile, Environment.NewLine + aScore(i).ToString)
            Next
            File.AppendAllText(fFile, Environment.NewLine + "0") 'Add a 0 to the file.
            For i As Integer = 0 To 9
                File.AppendAllText(fFile, Environment.NewLine + aUser(i)) 'Add all of the usernames to the file.
            Next
        End If

    End Sub
    'Go to the Score screen.
    Private Sub btnScore_Click(sender As Object, e As EventArgs) Handles btnScore.Click
        TabControl.SelectedIndex = 5
        labelAge.Text = System.IO.File.ReadAllLines(sFileLoc & "\Users\" & sUser & ".txt")(6) 'Set label text to user's stored age.
    End Sub
    Sub clearsheet()
        Dim aArrow(,) As Label = {{A11, A21, A31, A41, A51, A61, E1, H1, G1, T1}, {A12, A22, A32, A42, A52, A62, E2, H2, G2, T2}, {A13, A23, A33, A43, A53, A63, E3, H3, G3, T3}, {A14, A24, A34, A44, A54, A64, E4, H4, G4, T4}, {A15, A25, A35, A45, A55, A65, E5, H5, G5, T5}, {A16, A26, A36, A46, A56, A66, E6, G6, H6, T6}, {A17, A27, A37, A47, A57, A67, E7, H7, G7, T7}, {A18, A28, A38, A48, A58, A68, E8, G8, H8, T8}, {A19, A29, A39, A49, A59, A69, E9, H9, G9, T9}, {A110, A120, A130, A140, A150, A160, E10, G10, H10, T10}}
        For Each Space As Label In aArrow
            Space.Text = 0
            Space.BackColor = Color.White
            Space.ForeColor = Color.Black
        Next
        iTempScore = 11
    End Sub
    'Go to the Score Sheet Screen.
    Private Sub btnSheet_Click(sender As Object, e As EventArgs) Handles btnSheet.Click
        clearsheet()
        TabControl.SelectedIndex = 6
    End Sub

    'Go to the Account Info screen
    Private Sub btnAccount_Click(sender As Object, e As EventArgs) Handles btnAccount.Click
        TabControl.SelectedIndex = 3
    End Sub
    'Go to the Leader Board screen.
    Private Sub btnSub_Click(sender As Object, e As EventArgs) Handles btnSub.Click
        FileSelect() 'Use the "FileSelect" subroutine.
        TabControl.SelectedIndex = 7
    End Sub

    'Forces the user input of score to be a number only.
    Private Sub txtScore_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtScore.KeyPress
        If Not IsNumeric(e.KeyChar) AndAlso e.KeyChar <> ControlChars.Back Then
            e.Handled = True
        End If

    End Sub
    'Forces the user input of Age to be a number only.
    Private Sub txtAge_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtAge.KeyPress
        If Not IsNumeric(e.KeyChar) AndAlso e.KeyChar <> ControlChars.Back Then
            e.Handled = True
        End If
    End Sub
 
    'Apply changes when user selects a Portsmouth score sheet.
    Private Sub btnPortSheet_Click(sender As Object, e As EventArgs) Handles btnPortSheet.Click
        tbleSheet.Visible = True 'Make the table visible.
        btn1.Visible = True 'Make all of the buttons visible.
        btn2.Visible = True
        btn3.Visible = True
        btn4.Visible = True
        btn5.Visible = True
        btn6.Visible = True
        btn7.Visible = True
        btn8.Visible = True
        btn9.Visible = True
        btn10.Visible = True
        btnM.Visible = True
        btnX.Visible = True
        btnWorcSheet.Visible = False 'Make the round buttons invisible.
        btnPortSheet.Visible = False
        iTempScore = 11
    End Sub
    'Apply changes when user selects a Worcester score sheet.
    Private Sub btnWorcSheet_Click(sender As Object, e As EventArgs) Handles btnWorcSheet.Click
        bWorc = True 'Set a variable to True.
        tbleSheet.Visible = True 'Make the table visible.
        btn1.Visible = True
        btn1.BackColor = Color.Black 'Move and change the buttons to apply to the Worcester round.
        btn1.ForeColor = Color.White
        btn1.Location = New Point(210, 411)
        btn2.Visible = True
        btn2.BackColor = Color.Black
        btn2.ForeColor = Color.White
        btn2.Location = New Point(170, 411)
        btn3.Visible = True
        btn3.BackColor = Color.Black
        btn3.ForeColor = Color.White
        btn3.Location = New Point(130, 411)
        btn4.Visible = True
        btn4.BackColor = Color.Black
        btn4.ForeColor = Color.White
        btn4.Location = New Point(90, 411)
        btn5.Visible = True
        btn5.BackColor = Color.White
        btn5.Location = New Point(50, 411)
        btnM.Visible = True
        btnM.Location = New Point(130, 440)
        btnWorcSheet.Visible = False 'Make the round buttons invisible.
        btnPortSheet.Visible = False
        iTempScore = 11
    End Sub
    'Set next label to "X".
    Private Sub btnX_Click(sender As Object, e As EventArgs) Handles btnX.Click
        DirectCast(Me.Controls.Find("A" & iTempScore, True)(0), Label).BackColor = Color.Gold
        DirectCast(Me.Controls.Find("A" & iTempScore, True)(0), Label).Text = "X"
        iTempScore = iTempScore + 10
    End Sub
    'Set next label to "10".
    Private Sub btn10_Click(sender As Object, e As EventArgs) Handles btn10.Click
        DirectCast(Me.Controls.Find("A" & iTempScore, True)(0), Label).BackColor = Color.Gold
        DirectCast(Me.Controls.Find("A" & iTempScore, True)(0), Label).Text = "10"
        iTempScore = iTempScore + 10
    End Sub
    'Set next label to "9".
    Private Sub btn9_Click(sender As Object, e As EventArgs) Handles btn9.Click
        DirectCast(Me.Controls.Find("A" & iTempScore, True)(0), Label).BackColor = Color.Gold
        DirectCast(Me.Controls.Find("A" & iTempScore, True)(0), Label).Text = "9"
        iTempScore = iTempScore + 10
    End Sub
    'Set next label to "8".
    Private Sub btn8_Click(sender As Object, e As EventArgs) Handles btn8.Click
        DirectCast(Me.Controls.Find("A" & iTempScore, True)(0), Label).BackColor = Color.Red
        DirectCast(Me.Controls.Find("A" & iTempScore, True)(0), Label).Text = "8"
        iTempScore = iTempScore + 10
    End Sub
    'Set next label to "7".
    Private Sub btn7_Click(sender As Object, e As EventArgs) Handles btn7.Click
        DirectCast(Me.Controls.Find("A" & iTempScore, True)(0), Label).BackColor = Color.Red
        DirectCast(Me.Controls.Find("A" & iTempScore, True)(0), Label).Text = "7"
        iTempScore = iTempScore + 10
    End Sub
    'Set next label to "6".
    Private Sub btn6_Click(sender As Object, e As EventArgs) Handles btn6.Click
        DirectCast(Me.Controls.Find("A" & iTempScore, True)(0), Label).BackColor = Color.Cyan
        DirectCast(Me.Controls.Find("A" & iTempScore, True)(0), Label).Text = "6"
        iTempScore = iTempScore + 10
    End Sub
    'Set next label to "5".
    Private Sub btn5_Click(sender As Object, e As EventArgs) Handles btn5.Click
        If bWorc = True Then
            DirectCast(Me.Controls.Find("A" & iTempScore, True)(0), Label).BackColor = Color.White
            DirectCast(Me.Controls.Find("A" & iTempScore, True)(0), Label).ForeColor = Color.Black
        Else
            DirectCast(Me.Controls.Find("A" & iTempScore, True)(0), Label).BackColor = Color.Cyan
        End If
        DirectCast(Me.Controls.Find("A" & iTempScore, True)(0), Label).Text = "5"
        iTempScore = iTempScore + 10
    End Sub
    'Set next label to "4".
    Private Sub btn4_Click(sender As Object, e As EventArgs) Handles btn4.Click
        DirectCast(Me.Controls.Find("A" & iTempScore, True)(0), Label).BackColor = Color.Black
        DirectCast(Me.Controls.Find("A" & iTempScore, True)(0), Label).ForeColor = Color.White
        DirectCast(Me.Controls.Find("A" & iTempScore, True)(0), Label).Text = "4"
        iTempScore = iTempScore + 10
    End Sub
    'Set next label to "3".
    Private Sub btn3_Click(sender As Object, e As EventArgs) Handles btn3.Click
        DirectCast(Me.Controls.Find("A" & iTempScore, True)(0), Label).BackColor = Color.Black
        DirectCast(Me.Controls.Find("A" & iTempScore, True)(0), Label).ForeColor = Color.White
        DirectCast(Me.Controls.Find("A" & iTempScore, True)(0), Label).Text = "3"
        iTempScore = iTempScore + 10
    End Sub
    'Set next label to "2".
    Private Sub btn2_Click(sender As Object, e As EventArgs) Handles btn2.Click
        If bWorc = True Then
            DirectCast(Me.Controls.Find("A" & iTempScore, True)(0), Label).BackColor = Color.Black
            DirectCast(Me.Controls.Find("A" & iTempScore, True)(0), Label).ForeColor = Color.White
        Else
            DirectCast(Me.Controls.Find("A" & iTempScore, True)(0), Label).BackColor = Color.White
            DirectCast(Me.Controls.Find("A" & iTempScore, True)(0), Label).ForeColor = Color.Black
        End If
        DirectCast(Me.Controls.Find("A" & iTempScore, True)(0), Label).Text = "2"
        iTempScore = iTempScore + 10
    End Sub
    'Set next label to "1".
    Private Sub btn1_Click(sender As Object, e As EventArgs) Handles btn1.Click
        If bWorc = True Then
            DirectCast(Me.Controls.Find("A" & iTempScore, True)(0), Label).BackColor = Color.Black
            DirectCast(Me.Controls.Find("A" & iTempScore, True)(0), Label).ForeColor = Color.White
        Else
            DirectCast(Me.Controls.Find("A" & iTempScore, True)(0), Label).BackColor = Color.White
            DirectCast(Me.Controls.Find("A" & iTempScore, True)(0), Label).ForeColor = Color.Black
        End If
        DirectCast(Me.Controls.Find("A" & iTempScore, True)(0), Label).Text = "1"
        iTempScore = iTempScore + 10
    End Sub
    'Set next label to "M".
    Private Sub btnM_Click(sender As Object, e As EventArgs) Handles btnM.Click
        DirectCast(Me.Controls.Find("A" & iTempScore, True)(0), Label).BackColor = Color.Brown
        DirectCast(Me.Controls.Find("A" & iTempScore, True)(0), Label).Text = "M"
        iTempScore = iTempScore + 10
    End Sub
    Dim iRound As New Integer
    Dim iArrow As Integer = 0
    Dim iHit As Integer = 0
    Dim iGold As Integer = 0
    Dim iRT As Integer = 0
    Sub roundend(ByRef iArrow As Double, ByRef iGold As Double, ByRef iHit As Double)
        Dim aArrow(,) As Label = {{A11, A21, A31, A41, A51, A61}, {A12, A22, A32, A42, A52, A62}, {A13, A23, A33, A43, A53, A63}, {A14, A24, A34, A44, A54, A64}, {A15, A25, A35, A45, A55, A65}, {A16, A26, A36, A46, A56, A66}, {A17, A27, A37, A47, A57, A67}, {A18, A28, A38, A48, A58, A68}, {A19, A29, A39, A49, A59, A69}, {A110, A120, A130, A140, A150, A160}}
        iArrow = 0
        iHit = 0
        iGold = 0
        Dim isGold As Boolean = False
        For i As Integer = 0 To 5

            If IsNumeric(aArrow(iRound, i).Text) Then
                Dim iScore As Integer = Convert.ToInt16(aArrow(iRound, i).Text)
                iArrow = iArrow + iScore
                iHit = iHit + 1
                iRT = iRT + iScore
                If iScore = 9 Or iScore = 10 Then
                    isGold = True
                End If
            ElseIf aArrow(iRound, i).Text = "X" Then
                iArrow = iArrow + 10
                iHit = iHit + 1
                iRT = iRT + 10
                isGold = True
            End If
            If isGold = True Then
                iGold = iGold + 1
                isGold = False
            End If
        Next
    End Sub
    Private Sub A61_TextChanged(sender As Object, e As EventArgs) Handles A61.TextChanged
        iRound = 0
        roundend(iArrow, iGold, iHit)
        E1.Text = iArrow
        G1.Text = iGold
        H1.Text = iHit
        T1.Text = iRT
        iTempScore = 2
    End Sub

    Private Sub A62_TextChanged(sender As Object, e As EventArgs) Handles A62.TextChanged
        iRound = 1
        roundend(iArrow, iGold, iHit)

        E2.Text = iArrow
        G2.Text = iGold
        H2.Text = iHit
        T2.Text = iRT
        iTempScore = 3
    End Sub

    Private Sub A63_TextChanged(sender As Object, e As EventArgs) Handles A63.TextChanged
        iRound = 2
        roundend(iArrow, iGold, iHit)

        E3.Text = iArrow
        G3.Text = iGold
        H3.Text = iHit
        T3.Text = iRT
        iTempScore = 4
    End Sub

    Private Sub A64_TextChanged(sender As Object, e As EventArgs) Handles A64.TextChanged
        iRound = 3
        roundend(iArrow, iGold, iHit)

        E4.Text = iArrow
        G4.Text = iGold
        H4.Text = iHit
        T4.Text = iRT
        iTempScore = 5
    End Sub

    Private Sub A65_TextChanged(sender As Object, e As EventArgs) Handles A65.TextChanged
        iRound = 4
        roundend(iArrow, iGold, iHit)

        E5.Text = iArrow
        G5.Text = iGold
        H5.Text = iHit
        T5.Text = iRT
        iTempScore = 6
    End Sub

    Private Sub A66_TextChanged(sender As Object, e As EventArgs) Handles A66.TextChanged
        iRound = 5
        roundend(iArrow, iGold, iHit)

        E6.Text = iArrow
        G6.Text = iGold
        H6.Text = iHit
        T6.Text = iRT
        iTempScore = 7
    End Sub

    Private Sub A67_TextChanged(sender As Object, e As EventArgs) Handles A67.TextChanged
        iRound = 6
        roundend(iArrow, iGold, iHit)

        E7.Text = iArrow
        G7.Text = iGold
        H7.Text = iHit
        T7.Text = iRT
        iTempScore = 8
    End Sub

    Private Sub A68_TextChanged(sender As Object, e As EventArgs) Handles A68.TextChanged
        iRound = 7
        roundend(iArrow, iGold, iHit)

        E8.Text = iArrow
        G8.Text = iGold
        H8.Text = iHit
        T8.Text = iRT
        iTempScore = 9
    End Sub

    Private Sub A69_TextChanged(sender As Object, e As EventArgs) Handles A69.TextChanged
        iRound = 8
        roundend(iArrow, iGold, iHit)

        E9.Text = iArrow
        G9.Text = iGold
        H9.Text = iHit
        T9.Text = iArrow + T8.Text
        iTempScore = 100
    End Sub

    Private Sub A610_TextChanged(sender As Object, e As EventArgs) Handles A160.TextChanged
        iRound = 9
        roundend(iArrow, iGold, iHit)

        E10.Text = iArrow
        G10.Text = iGold
        H10.Text = iHit
        T10.Text = iRT
        iTempScore = 0
        btnSubScore.Visible = True
        btnX.Visible = False
        btn10.Visible = False
        btn9.Visible = False
        btn8.Visible = False
        btn7.Visible = False
        btn6.Visible = False
        btn5.Visible = False
        btn4.Visible = False
        btn3.Visible = False
        btn2.Visible = False
        btn1.Visible = False
        btnM.Visible = False
    End Sub

    Private Sub btnSubScore_Click(sender As Object, e As EventArgs) Handles btnSubScore.Click
        txtScore.Text = T10.Text
        iAge = System.IO.File.ReadAllLines(sFileLoc & "\Users\" & sUser & ".txt")(6)
        If bWorc = True Then
            cboxWorc.Checked = True
        Else
            cboxPort.Checked = True
        End If
        FileSelect()
        TabControl.SelectedIndex = 7
    End Sub
    Sub FileSelect()
        Dim bBow As Boolean = False
        If System.IO.File.ReadAllLines(sFileLoc & "\Users\" & sUser & ".txt")(3) = "Recurve" Then
            bBow = True
        End If
        If cboxPort.Checked Then
            If txtScore.Text < 600 Then
                If iAge < 18 Then
                    If bBow = True Then
                        fFileName = "Jr Recurve Portsmouth"
                        Score()
                    Else
                        fFileName = "Jr Compound Portsmouth"
                        Score()
                    End If
                Else
                    If bBow = True Then
                        fFileName = "Sr Recurve Portsmouth"
                        Score()
                    Else
                        fFileName = "Sr Compound Portsmouth"
                        Score()
                    End If
                End If
            End If
        ElseIf cboxWorc.Checked Then
            If txtScore.Text < 300 Then
                If iAge < 18 Then
                    If bBow = True Then
                        fFileName = "Jr Recurve Worcester"
                        Score()
                    Else
                        fFileName = "Jr Compound Worcester"
                        Score()
                    End If
                Else
                    If bBow = True Then
                        fFileName = "Sr Recurve Worcester"
                        Score()
                    Else
                        fFileName = "Sr Compound Worcester"
                        Score()
                    End If
                End If
            Else
                MsgBox("Score is too high.")
            End If
        Else
            MsgBox("Score is too high")
        End If
    End Sub

    Private Sub cbSr_CheckedChanged(sender As Object, e As EventArgs) Handles cbSr.CheckedChanged
        f1 = "Sr "
        cbJr.Checked = False
    End Sub

    Private Sub cbJr_CheckedChanged(sender As Object, e As EventArgs) Handles cbJr.CheckedChanged
        f1 = "Jr "
        cbSr.Checked = False
    End Sub

    Private Sub cbRecurve_CheckedChanged(sender As Object, e As EventArgs) Handles cbRecurve.CheckedChanged
        f2 = "Recurve "
        cbCompound.Checked = False
    End Sub

    Private Sub cbCompound_CheckedChanged(sender As Object, e As EventArgs) Handles cbCompound.CheckedChanged
        f2 = "Compound "
        cbRecurve.Checked = False
    End Sub

    Private Sub cbPort_CheckedChanged(sender As Object, e As EventArgs) Handles cbPort.CheckedChanged
        f3 = "Portsmouth"
        cbWorc.Checked = False
    End Sub

    Private Sub cbWorc_CheckedChanged(sender As Object, e As EventArgs) Handles cbWorc.CheckedChanged
        f3 = "Worcester"
        cbPort.Checked = False
    End Sub

    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click
        TabControl.SelectedIndex = 4
    End Sub

    Private Sub btnFind_Click(sender As Object, e As EventArgs) Handles btnFind.Click
        fFileName = (f1 + f2 + f3)
        Dim changed As Boolean = False
        If f1 = "0" And f3 = "0" And f3 = "0" Then
            MsgBox("You need to select all three options.")
            changed = True
        Else
            Dim fFile As String = sFileLoc & "\Scores\" & fFileName & ".txt"
            Dim aScores() As Label = {S1, S2, S3, S4, S5, S6, S7, S8, S9, S10}
            Dim aUsers() As Label = {U1, U2, U3, U4, U5, U6, U7, U8, U9, U10}
            For i As Integer = 0 To 9
                aScores(i).Text = System.IO.File.ReadAllLines(fFile)(i)
                aUsers(i).Text = System.IO.File.ReadAllLines(fFile)(i + 11)
                changed = True
            Next
        End If
        If changed = False Then
            MsgBox("You have unchecked a box, please check all boxes.")
        End If
        changed = False
    End Sub

    Private Sub cbRec_CheckedChanged(sender As Object, e As EventArgs) Handles cbRec.CheckedChanged
        cbComp.Checked = False
        clBow = "Recurve"
    End Sub

    Private Sub cbComp_CheckedChanged(sender As Object, e As EventArgs) Handles cbComp.CheckedChanged
        cbRec.Checked = False
        clBow = "Compound"
    End Sub

    Private Sub cbMale_CheckedChanged(sender As Object, e As EventArgs) Handles cbMale.CheckedChanged
        cbFemale.Checked = False
        clGender = "Male"
    End Sub

    Private Sub cbFemale_CheckedChanged(sender As Object, e As EventArgs) Handles cbFemale.CheckedChanged
        cbMale.Checked = False
        clGender = "Female"
    End Sub

    Private Sub sGNAS_KeyPress(sender As Object, e As KeyPressEventArgs) Handles sGNAS.KeyPress
        If Not IsNumeric(e.KeyChar) AndAlso e.KeyChar <> ControlChars.Back Then
            e.Handled = True
        End If
    End Sub

    Private Sub btnCancelScore_Click(sender As Object, e As EventArgs) Handles btnCancelScore.Click
        TabControl.SelectedIndex = 4
    End Sub
End Class
