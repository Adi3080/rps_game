Public Class Form1


    Public strPlayerchoice As String = "nochoiceYet"
    Public strComputerchoice As String = "nochoiceYet"

    Public binPlayerWins As Boolean = "False"
    Public binComputerWins As Boolean = "False"
    Public intNbrOfPlayerWins As Integer = 0
    Public intNbrOfPlayerLosses As Integer = 0
    Public intNbrOfPlayerTies As Integer = 0S


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label2.Text = "0"
        Label3.Text = "0"
        Label4.Text = "0"

        PictureBox1.Image = My.Resources.ResourceManager.GetObject("Playerjpg")
        If PictureBox1.Image Is Nothing Then
            MessageBox.Show("RPS player image not found in program resources library")
        End If
        PictureBox2.Image = My.Resources.ResourceManager.GetObject("Computerjpg")
        If PictureBox2.Image Is Nothing Then
            MessageBox.Show("RPS player image not found in program resources library")
        End If
        Button8.ForeColor = Color.DarkTurquoise
        Button8.FlatStyle = FlatStyle.Standard
        Button8.TextAlign = ContentAlignment.MiddleCenter
        Button8.Text = "Play now " & vbCrLf & "And Win!"

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        strPlayerchoice = "Rock"

        PictureBox1.Image = My.Resources.ResourceManager.GetObject("Rockplayerjpg")
        PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage
        If PictureBox1.Image Is Nothing Then
            MessageBox.Show("Player rock image not found in program resource library ")
        End If
        Button4.Enabled = False
        Button5.Enabled = False
        Button3.Enabled = False
        makerandomcoputerchoice()
        CheckForWinner()
        DisplayGameResult()


    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        strPlayerchoice = "Paper"

        PictureBox1.Image = My.Resources.ResourceManager.GetObject("Paperplayerjpg")
        PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage
        If PictureBox1.Image Is Nothing Then
            MessageBox.Show("Player rock image not found in program resource library ")
        End If
        Button3.Enabled = False
        Button5.Enabled = False
        Button4.Enabled = False
        makerandomcoputerchoice()
        CheckForWinner()
        DisplayGameResult()



    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        strPlayerchoice = "Scissor"

        PictureBox1.Image = My.Resources.ResourceManager.GetObject("Scissorplayerjpg")
        PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage
        If PictureBox1.Image Is Nothing Then
            MessageBox.Show("Player rock image not found in program resource library ")
        End If
        Button4.Enabled = False
        Button3.Enabled = False
        Button5.Enabled = False
        makerandomcoputerchoice()
        CheckForWinner()
        DisplayGameResult()


    End Sub
    Public Sub makerandomcoputerchoice()
        Randomize()
        Dim intcomputerRandomchoice As Integer = Int((3 * Rnd()) + 1)
        If intcomputerRandomchoice = 1 Then
            strComputerchoice = "Rock"
        ElseIf intcomputerRandomchoice = 2 Then
            strComputerchoice = "Paper"
        ElseIf intcomputerRandomchoice = 3 Then
            strComputerchoice = "Scissor"
        Else
            MessageBox.Show("Error in computer choice random number generator ")
        End If
        If strComputerchoice = "Rock" Then
            PictureBox2.Image = My.Resources.ResourceManager.GetObject("Rockcomputerjpg")
            PictureBox2.SizeMode = PictureBoxSizeMode.StretchImage
        End If
        If strComputerchoice = "Paper" Then
            PictureBox2.Image = My.Resources.ResourceManager.GetObject("Papercomputerjpg")
            PictureBox2.SizeMode = PictureBoxSizeMode.StretchImage
        End If
        If strComputerchoice = "Scissor" Then
            PictureBox2.Image = My.Resources.ResourceManager.GetObject("Scissorcomputerjpg")
            PictureBox2.SizeMode = PictureBoxSizeMode.StretchImage
        End If
    End Sub
    Public Sub CheckForWinner()
        Select Case strPlayerchoice
            Case "Rock"
                CheckRockVersusComputerSelection()
            Case "Paper"
                CheckPaperVersusComputerSelection()
            Case "Scissor"
                CheckScissorVersusComputerSelection()
            Case Else
                MessageBox.Show("Error:Player has not selected Rock Paper or Scissors")

        End Select
    End Sub
    Public Sub CheckRockVersusComputerSelection()
        If strComputerchoice = "Rock" Then
            binComputerWins = "True"
            binPlayerWins = "True"

        ElseIf strComputerchoice = "Paper" Then
            binComputerWins = "True"

        ElseIf strComputerchoice = "Scissor" Then
            binPlayerWins = "True"

        Else
            MessageBox.Show("Check Winner Error: Computer has not selected rock paper or scissor ")
        End If
    End Sub
    Public Sub CheckPaperVersusComputerSelection()
        If strComputerchoice = "Rock" Then
            binPlayerWins = "True"

        ElseIf strComputerchoice = "Paper" Then
            binComputerWins = "True"
            binPlayerWins = "True"


        ElseIf strComputerchoice = "Scissor" Then
            binComputerWins = "True"

        Else
            MessageBox.Show("Check Winner Error: Computer has not selected rock paper or scissor ")
        End If
    End Sub
    Public Sub CheckScissorVersusComputerSelection()
        If strComputerchoice = "Rock" Then
            binComputerWins = "True"

        ElseIf strComputerchoice = "Paper" Then
            binPlayerWins = "True"

        ElseIf strComputerchoice = "Scissor" Then
            binPlayerWins = "True"
            binComputerWins = "True"


        Else
            MessageBox.Show("Check Winner Error: Computer has not selected rock paper or scissor ")
        End If
    End Sub
    Public Sub DisplayGameResult()
        If strPlayerchoice = "Rock" And binPlayerWins = "True" And binComputerWins = "false" Then
            intNbrOfPlayerWins = intNbrOfPlayerWins + 1
            Label2.Text = CStr(intNbrOfPlayerWins)
            Button8.ForeColor = Color.DarkTurquoise
            Button8.FlatStyle = FlatStyle.Standard
            Button8.TextAlign = ContentAlignment.MiddleCenter
            Button8.Text = "Player " & vbCrLf & "Wins!"
            PictureBox1.Image = My.Resources.ResourceManager.GetObject("GIF_PLAYER_ROCK_WINS")
        End If
        If strPlayerchoice = "Paper" And binPlayerWins = "True" And binComputerWins = "false" Then
            intNbrOfPlayerWins = intNbrOfPlayerWins + 1
            Label2.Text = CStr(intNbrOfPlayerWins)
            Button8.ForeColor = Color.DarkTurquoise
            Button8.FlatStyle = FlatStyle.Standard
            Button8.TextAlign = ContentAlignment.MiddleCenter
            Button8.Text = "Player " & vbCrLf & "Wins!"
            PictureBox1.Image = My.Resources.ResourceManager.GetObject("GIF_PLAYER_PAPER_WINS")
        End If
        If strPlayerchoice = "Scissor" And binPlayerWins = "True" And binComputerWins = "false" Then
            intNbrOfPlayerWins = intNbrOfPlayerWins + 1
            Label2.Text = CStr(intNbrOfPlayerWins)
            Button8.ForeColor = Color.DarkTurquoise
            Button8.FlatStyle = FlatStyle.Standard
            Button8.TextAlign = ContentAlignment.MiddleCenter
            Button8.Text = "Player " & vbCrLf & "Wins!"
            PictureBox1.Image = My.Resources.ResourceManager.GetObject("GIF_PLAYER_SCISSOR_WINS")
        End If
        If strComputerchoice = "Rock" And binComputerWins = "true" And binPlayerWins = "false" Then
            intNbrOfPlayerLosses = intNbrOfPlayerLosses + 1
            Label3.Text = CStr(intNbrOfPlayerLosses)
            Button8.ForeColor = Color.Red
            Button8.FlatStyle = FlatStyle.Standard
            Button8.TextAlign = ContentAlignment.MiddleCenter
            Button8.Text = "Computer " & vbCrLf & "Wins!"
            PictureBox1.Image = My.Resources.ResourceManager.GetObject("GIF_COMPUTER_ROCK_WINS")
        End If
        If strComputerchoice = "Paper" And binComputerWins = "true" And binPlayerWins = "false" Then
            intNbrOfPlayerLosses = intNbrOfPlayerLosses + 1
            Label3.Text = CStr(intNbrOfPlayerLosses)
            Button8.ForeColor = Color.Red
            Button8.FlatStyle = FlatStyle.Standard
            Button8.TextAlign = ContentAlignment.MiddleCenter
            Button8.Text = "Computer  " & vbCrLf & "Wins!"
            PictureBox1.Image = My.Resources.ResourceManager.GetObject("GIF_COMPUTER_PAPER_WINS")
        End If
        If strComputerchoice = "Scissor" And binComputerWins = "true" And binPlayerWins = "false" Then
            intNbrOfPlayerLosses = intNbrOfPlayerLosses + 1
            Label3.Text = CStr(intNbrOfPlayerLosses)
            Button8.ForeColor = Color.Red
            Button8.FlatStyle = FlatStyle.Standard
            Button8.TextAlign = ContentAlignment.MiddleCenter
            Button8.Text = "Computer " & vbCrLf & "Wins"
            PictureBox1.Image = My.Resources.ResourceManager.GetObject("GIF_COMPUTER_SCISSOR_WINS")
        End If
        If strComputerchoice = "Scissor" And strPlayerchoice = "Scissor" Then
            intNbrOfPlayerTies = intNbrOfPlayerTies + 1
            Label4.Text = CStr(intNbrOfPlayerTies)
            Button8.ForeColor = Color.Yellow
            Button8.FlatStyle = FlatStyle.Standard
            Button8.TextAlign = ContentAlignment.MiddleCenter
            Button8.Text = "It is a  " & vbCrLf & "Tie game "
        End If
        If strComputerchoice = "Rock" And strPlayerchoice = "Rock" Then
            intNbrOfPlayerTies = intNbrOfPlayerTies + 1
            Label4.Text = CStr(intNbrOfPlayerTies)
            Button8.ForeColor = Color.Yellow
            Button8.FlatStyle = FlatStyle.Standard
            Button8.TextAlign = ContentAlignment.MiddleCenter
            Button8.Text = "It is a  " & vbCrLf & "Tie game "
        End If
        If strComputerchoice = "Paper" And strPlayerchoice = "Paper" Then
            intNbrOfPlayerTies = intNbrOfPlayerTies + 1
            Label4.Text = CStr(intNbrOfPlayerTies)
            Button8.ForeColor = Color.Yellow
            Button8.FlatStyle = FlatStyle.Standard
            Button8.TextAlign = ContentAlignment.MiddleCenter
            Button8.Text = "It is a  " & vbCrLf & "Tie game "
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim confirm_msg As Integer
        confirm_msg = MessageBox.Show("Are you sure you want to exit?", "Exit amd close",
        MessageBoxButtons.YesNo, MessageBoxIcon.Information)
        If confirm_msg = vbYes Then
            Me.Close()
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click


        PictureBox1.Image = My.Resources.ResourceManager.GetObject("Playerjpg")
        If PictureBox1.Image Is Nothing Then
            MessageBox.Show("RPS player image not found in program resources library")
        End If
        PictureBox2.Image = My.Resources.ResourceManager.GetObject("Computerjpg")
        If PictureBox2.Image Is Nothing Then
            MessageBox.Show("RPS player image not found in program resources library")
        End If
        Button8.ForeColor = Color.DarkTurquoise
        Button8.FlatStyle = FlatStyle.Standard
        Button8.TextAlign = ContentAlignment.MiddleCenter
        Button8.Text = "Play now " & vbCrLf & "And Win!"
        Button4.Enabled = True
        Button5.Enabled = True
        Button3.Enabled = True
    End Sub
End Class
