Imports Fly.My.Resources.resFly
'Mosca (fly)
'Completed: 9:28 PM 1/13/2021
'Chris Morningstar
'Although the heavy work on this was done in 2017, today I removed the need for the external
'gif files and compiled a release of the fly.  The fly pictures are in the resFly resource 
'file now and fly.exe can run on any machine with the .net 3.5 framework installed.  
'
Public Class Fly
    'Fly Variables
    Private Dot_Ro As Long = 60         'Dot's distance from the mouse pointer
    Private Dot_Theta As Double = 0       'Dot's initial angle
    Private Dot_Speed As Double           'Dot's absolute Angular speed
    Private Dot_Direction As Long = -1   'Dot's direction (1=clockwise)
    Private Dot_x As Double = 0           'Dot's original position
    Private Dot_y As Double = 0
    Private alpha As Double
    Private mult As Double                'Angle from the fly to the mouse
    Private picX As Double = 20           'Ausiliary variable to define the angle
    Private picY As Double = 100          'Fly's coords.
    Private TpicX As Double = 20           'Ausiliary variable to define the angle
    Private TpicY As Double = 100          'Fly's coords.
    Private intStep As Integer = 10     'Pixels
    Private speed As Integer = 100      'u-seconds
    Private sFlyToUse(9) As String      'sFlyToUse specifies the right picture;
    Private LastUsed As Integer = 0
    'Private img() As String             'img pre-caches images.
    Dim mBit As Integer
    Private ns As Boolean                       'I added These, but ns is still kind of a mystery.
    Private mouseX As Long = MousePosition.X    'and these
    Private mouseY As Long = MousePosition.Y    'and these
    Dim ImgPath As String = "flys\"             'This line is no longer needed.
    Dim Debugging As Boolean = False
    'End Fly Variables
    Dim myFly As My.Resources.resFly

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ns = Me.Visible
        Call InitializeFly()
        Timer1.Interval = speed
        Call moveFly()

        If Debugging = True Then
            mouseX = Screen.PrimaryScreen.Bounds.Width / 2
            mouseY = Screen.PrimaryScreen.Bounds.Height / 2
        End If
        'setInterval('ChangeDotDirection()', speed*50);
        'Timer1.Interval = speed * 50
        ChangeDotDirection()


        'Me.Location = New Point(0 - Me.Width, My.Computer.Screen.WorkingArea.Bottom - Me.Height + 5)
    End Sub

    ' Moves the fly around the screen
    Private Sub moveFly()
        '
        'This is an asumption.
        picX = Me.Location.X
        picY = Me.Location.Y
        ' moves the fly in a new position...
        calcNewPos()
        'ns is not declared and I'm not sure if pic or flydiv represents "Me"
        If (ns) Then 'not totally sure on what I think ns means...
            Me.Location = New Point(picX, picY)
        Else
            Me.Location = New Point(picX - Me.Width / 2, picY - Me.Height / 2)
        End If
        ' ... and changes the image.
        alpha = -180 * alpha / Math.PI
        alpha += 22.5
        Dim Ok As Boolean = False
        Dim i As Integer = 0

        'This will have to be followed.
        Do While ((i < 4) AndAlso Not Ok)
            If (alpha < ((Math.PI * -1) + (45 * i))) Then
                display((mult * (i + 1)))
                Ok = True
            End If
            i = i + 1
        Loop
    End Sub
    'Calculate new position
    Private Sub calcNewPos()
        '	/*
        '		All this calculations make the Dot
        '		to come near the mouse-pointer,
        '		and the fly to follow the dot.
        '	*/
        Dim dist = Math.Sqrt(Math.Pow(mouseY - picY, 2) + Math.Pow(mouseX - picX, 2))
        Dot_Speed = (Math.PI / 15)
        'Debug.Print(Dot_Speed)
        Dot_Theta = (Dot_Theta + (Dot_Direction * Dot_Speed))
        'Dot_Theta = (Dot_Theta + (-1 * Dot_Speed))
        Dot_x = (mouseX + (Dot_Ro * Math.Cos(Dot_Theta)))
        Dot_y = (mouseY + (Dot_Ro * Math.Sin(Dot_Theta)))
        '
        'ChangeDotDirection() will have to be a little more random as it can't happen everytime, but should happen
        'more often
        '
        'I'm pretty sure the code is right, but the fly actually reached his destination and boom... divide by zero.
        'so i'm going to help him to not have that issue again.
        If (Dot_x - picX) = 0 Then
            'Congrats mr fly... you made it to where you were going.
            'Debug.Print("Dot_x = picX")
            ChangeDotDirection()
            'ns = False
            Randomize()
            Dot_x = Dot_x + CInt(Math.Floor((10 - 1 + 1) * Rnd())) + 1
            'Dot_x = CLng(Math.Floor((Screen.PrimaryScreen.Bounds.Width() - 0 + 1) * Rnd())) + 0
            'Dot_y = CLng(Math.Floor((Screen.PrimaryScreen.Bounds.Height() - 0 + 1) * Rnd())) + 0
        End If
        Dim arg As Long = ((Dot_y - picY) / (Dot_x - picX))
        If (Dot_x - picX) < 0 Then
            mult = -1
        Else
            mult = 1
        End If
        alpha = Math.Atan(arg)
        Dim dx = (mult * (intStep * Math.Cos(alpha)))
        Dim dy = (mult * (intStep * Math.Sin(alpha)))

        picX = (picX + dx)
        picY = (picY + dy)

    End Sub

    ' Shows the proper image for the fly.
    Private Sub display(ByVal direction As Integer)
        Dim NewDirection As Integer

        If direction < 0 Then
            NewDirection = Math.Abs(direction) + 4
        Else
            NewDirection = direction
        End If

        'direction must be from -4 to 4, but not 0.

        'Here we need to follow what 'direction' is.
        'PictureBox1.ImageLocation = sFlyToUse(direction)
        If LastUsed <> NewDirection Then
            Debug.Print(Application.StartupPath & "\" & sFlyToUse(NewDirection))
            'Old line that requires external images... replaced 1/13/2021 
            'PictureBox1.Image = Image.FromFile(Application.StartupPath & "\" & sFlyToUse(NewDirection))
            'Pictures are in the resource file...
            PictureBox1.Image = My.Resources.resFly.ResourceManager.GetObject(sFlyToUse(NewDirection))
            LastUsed = NewDirection
        End If
    End Sub

    Private Sub InitializeFly()
        Dot_x = MousePosition.X 'Create an initial target.
        Dot_y = MousePosition.Y

        'using external .gif files...
        'sFlyToUse(0) = ""
        'sFlyToUse(1) = ImgPath & "DxR.gif"  'dir[1]="Mosca_2.gif";
        'sFlyToUse(2) = ImgPath & "R.gif"    'dir[2]="Mosca_1.gif";
        'sFlyToUse(3) = ImgPath & "UxR.gif"  'dir[3]="Mosca_8.gif";
        'sFlyToUse(4) = ImgPath & "U.gif"    'dir[4]="Mosca_7.gif";
        'sFlyToUse(5) = ImgPath & "UxL.gif"  'dir[-1]="Mosca_6.gif";
        'sFlyToUse(6) = ImgPath & "L.gif"    'dir[-2]="Mosca_5.gif";
        'sFlyToUse(7) = ImgPath & "DxL.gif"  'dir[-3]="Mosca_4.gif";
        'sFlyToUse(8) = ImgPath & "D.gif"    'dir[-4]="Mosca_3.gif";

        sFlyToUse(0) = ""
        sFlyToUse(1) = "DxR"  'dir[1]="Mosca_2.gif";
        sFlyToUse(2) = "R"    'dir[2]="Mosca_1.gif";
        sFlyToUse(3) = "UxR"  'dir[3]="Mosca_8.gif";
        sFlyToUse(4) = "U"    'dir[4]="Mosca_7.gif";
        sFlyToUse(5) = "UxL"  'dir[-1]="Mosca_6.gif";
        sFlyToUse(6) = "L"    'dir[-2]="Mosca_5.gif";
        sFlyToUse(7) = "DxL"  'dir[-3]="Mosca_4.gif";
        sFlyToUse(8) = "D"    'dir[-4]="Mosca_3.gif";
    End Sub

    'Changes Dot's turning direction
    Private Sub ChangeDotDirection()
        Dot_Direction = -Dot_Direction
        Debug.Print(Dot_Direction)
        Dot_Theta = Dot_Theta + Math.PI
    End Sub

    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        'function getMousePosition(e) {
        'mouseY=(ns)?e.pageY:window.event.y + document.body.scrollTop;
        'mouseX=(ns)?e.pageX:window.event.x + document.body.scrollLeft;
        '}
        'This gives another hint about ns...
        If Not Debugging Then
            mouseX = MousePosition.X
            mouseY = MousePosition.Y
        End If
        moveFly()

    End Sub

    'Added partly for AMXrnd but it's a reuseable routine to get random numbers.
    Private Function GetRandom(ByVal Min As Integer, ByVal Max As Integer) As Integer
        ' by making Generator static, we preserve the same instance '
        ' (i.e., do not create new instances with the same seed over and over) '
        ' between calls '
        Static Generator As System.Random = New System.Random()
        GetRandom = Generator.Next(Min, Max)
        'or
        '- Initialize the random-number generator.
        'Randomize()
        '- Generate random value between 1 and 6.
        'GetRandom = CInt(Int((Max * Rnd()) + Min))
    End Function

    Private Sub PictureBox1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseMove
        Dim direct As Integer = GetRandom(-1, 1)
        If Dot_Direction <> direct Then
            ChangeDotDirection()
        End If
    End Sub
End Class