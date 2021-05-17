Public Class Form1
    '*************************************************************************************************
    'Programmer Name: Judie Hu
    'Date: May 2018
    'Purpose: Write a program which will train kids to respond to things quickly
    'Language: Visual Basic
    '*************************************************************************************************
    'Variables: 
    '    RndPic - move a picture at random 
    '    MOVEBY - selected picture moves down by 
    '    YLOC - y location of selected picture
    '    AWAY - away score
    '    Hit - hit score
    '*************************************************************************************************
    'Components Used(pictures, sounds, movies etc...):
    '   16-magenta balloon.png, 18-sapphire-blue balloon.png, 32-silver balloon.png, green balloon.png, purple balloon.png(5 colourful balloons)
    '   background image.jpg (background image)
    '   floop.wav (hit sound)
    '   peeeooop_x.wav (away sound)
    '   balloon.ico (icon image)
    '**************************************************************************************************

    Dim RndPIC, MOVEBY, YLoc, AWAY, HIT As Integer

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Randomize()
        MOVEBY = 20
        AWAY = 0
        HIT = 0
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        RndPIC = Int((5 - 1 + 1) * Rnd() + 1)
        If RndPIC = 1 Then
            YLoc = PictureBox1.Location.Y + MOVEBY
            PictureBox1.Location = New Point(PictureBox1.Location.X, YLoc)
        ElseIf RndPIC = 2 Then
            YLoc = PictureBox2.Location.Y + MOVEBY
            PictureBox2.Location = New Point(PictureBox2.Location.X, YLoc)
        ElseIf RndPIC = 3 Then
            YLoc = PictureBox3.Location.Y + MOVEBY
            PictureBox3.Location = New Point(PictureBox3.Location.X, YLoc)
        ElseIf RndPIC = 4 Then
            YLoc = PictureBox4.Location.Y + MOVEBY
            PictureBox4.Location = New Point(PictureBox4.Location.X, YLoc)
        ElseIf RndPIC = 5 Then
            YLoc = PictureBox5.Location.Y + MOVEBY
            PictureBox5.Location = New Point(PictureBox5.Location.X, YLoc)
        End If

        If PictureBox1.Location.Y > Me.Height Then
            PictureBox1.Location = New Point(PictureBox1.Location.X, 0)
            AWAY = AWAY + 1
            My.Computer.Audio.Play(My.Resources.peeeooop_x, AudioPlayMode.Background)
        ElseIf PictureBox2.Location.Y > Me.Height Then
            PictureBox2.Location = New Point(PictureBox2.Location.X, 0)
            AWAY = AWAY + 1
            My.Computer.Audio.Play(My.Resources.peeeooop_x, AudioPlayMode.Background)
        ElseIf PictureBox3.Location.Y > Me.Height Then
            PictureBox3.Location = New Point(PictureBox3.Location.X, 0)
            AWAY = AWAY + 1
            My.Computer.Audio.Play(My.Resources.peeeooop_x, AudioPlayMode.Background)
        ElseIf PictureBox4.Location.Y > Me.Height Then
            PictureBox4.Location = New Point(PictureBox4.Location.X, 0)
            AWAY = AWAY + 1
            My.Computer.Audio.Play(My.Resources.peeeooop_x, AudioPlayMode.Background)
        ElseIf PictureBox5.Location.Y > Me.Height Then
            PictureBox5.Location = New Point(PictureBox5.Location.X, 0)
            AWAY = AWAY + 1
            My.Computer.Audio.Play(My.Resources.peeeooop_x, AudioPlayMode.Background)
        End If

        lblAway.Text = Str(AWAY)
        If AWAY = 5 Then
            lblGameOver.Visible = True
            Timer1.Enabled = False
        End If
        If AWAY = 5 Then
            PictureBox1.Enabled = False
            PictureBox2.Enabled = False
            PictureBox3.Enabled = False
            PictureBox4.Enabled = False
            PictureBox5.Enabled = False
        End If
    End Sub

    Private Sub cmdEasy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEasy.Click
        Timer1.Interval = 100
        Timer1.Enabled = True
    End Sub

    Private Sub cmdMedium_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMedium.Click
        Timer1.Interval = 50
        Timer1.Enabled = True
    End Sub

    Private Sub cmdCrazy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCrazy.Click
        Timer1.Interval = 1
        Timer1.Enabled = True
    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click
        PictureBox1.Location = New Point(PictureBox1.Location.X, 0)
        HIT = HIT + 1
        lblHit.Text = Str(HIT)
        My.Computer.Audio.Play(My.Resources.floop, AudioPlayMode.Background)
    End Sub

    Private Sub PictureBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox2.Click
        PictureBox2.Location = New Point(PictureBox2.Location.X, 0)
        HIT = HIT + 1
        lblHit.Text = Str(HIT)
        My.Computer.Audio.Play(My.Resources.floop, AudioPlayMode.Background)
    End Sub

    Private Sub PictureBox3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox3.Click
        PictureBox3.Location = New Point(PictureBox3.Location.X, 0)
        HIT = HIT + 1
        lblHit.Text = Str(HIT)
        My.Computer.Audio.Play(My.Resources.floop, AudioPlayMode.Background)
    End Sub

    Private Sub PictureBox4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox4.Click
        PictureBox4.Location = New Point(PictureBox4.Location.X, 0)
        HIT = HIT + 1
        lblHit.Text = Str(HIT)
        My.Computer.Audio.Play(My.Resources.floop, AudioPlayMode.Background)
    End Sub

    Private Sub PictureBox5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox5.Click
        PictureBox5.Location = New Point(PictureBox5.Location.X, 0)
        HIT = HIT + 1
        lblHit.Text = Str(HIT)
        My.Computer.Audio.Play(My.Resources.floop, AudioPlayMode.Background)
    End Sub

    Private Sub cmdQuit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdQuit.Click
        Timer1.Enabled = False
    End Sub

    Private Sub cmdAgain_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAgain.Click
        AWAY = 0
        HIT = 0
        lblHit.Text = Str(HIT)
        lblAway.Text = Str(AWAY)
        PictureBox1.Location = New Point(PictureBox1.Location.X, 0)
        PictureBox2.Location = New Point(PictureBox2.Location.X, 0)
        PictureBox3.Location = New Point(PictureBox3.Location.X, 0)
        PictureBox4.Location = New Point(PictureBox4.Location.X, 0)
        PictureBox5.Location = New Point(PictureBox5.Location.X, 0)
        PictureBox1.Enabled = True
        PictureBox2.Enabled = True
        PictureBox3.Enabled = True
        PictureBox4.Enabled = True
        PictureBox5.Enabled = True
        lblGameOver.Visible = False
    End Sub
End Class
