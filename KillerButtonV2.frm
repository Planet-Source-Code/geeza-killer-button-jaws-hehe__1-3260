VERSION 5.00
Begin VB.Form frmKillerButton 
   BackColor       =   &H80000018&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   5520
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNoBut 
      Caption         =   "no"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3540
      TabIndex        =   1
      Top             =   2100
      Width           =   375
   End
   Begin VB.Timer tmrFollow 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3540
      Top             =   2460
   End
   Begin VB.CommandButton cmdYesBut 
      Caption         =   "Yes!  I love this program!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1500
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1860
      Width           =   1575
   End
   Begin VB.Image Up_LeftPic 
      Height          =   570
      Left            =   4320
      Picture         =   "KILLER~3.frx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image LeftPic 
      Height          =   570
      Left            =   3780
      Picture         =   "KILLER~3.frx":0432
      Top             =   0
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image Down_LeftPic 
      Height          =   570
      Left            =   3240
      Picture         =   "KILLER~3.frx":086D
      Top             =   0
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image DownPic 
      Height          =   570
      Left            =   2700
      Picture         =   "KILLER~3.frx":0CA5
      Top             =   0
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image Down_RightPic 
      Height          =   570
      Left            =   2160
      Picture         =   "KILLER~3.frx":10DD
      Top             =   0
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image RightPic 
      Height          =   570
      Left            =   1620
      Picture         =   "KILLER~3.frx":150F
      Top             =   0
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image Up_RightPic 
      Height          =   570
      Left            =   1080
      Picture         =   "KILLER~3.frx":194A
      Top             =   0
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image UpPic 
      Height          =   570
      Left            =   540
      Picture         =   "KILLER~3.frx":1D82
      Top             =   0
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image FacePic 
      Height          =   570
      Left            =   0
      Picture         =   "KILLER~3.frx":21B7
      Top             =   0
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Would you like to Register me?"
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   1440
      Width           =   2295
   End
End
Attribute VB_Name = "frmKillerButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    'Get the current cursor Hot-Spot position

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Const a_Radius = 30 'Acceptable Radius the cursor can be
                'within for the button to 'grab' the cursor
Const HWND_TOPMOST = -1
Dim XnY As POINTAPI, ExitDo As Boolean

Private Sub cmdNoBut_Click()
    cmdYesBut.ZOrder 0  'Set the follower button to infront
    cmdYesBut.Font.Size = 8
    cmdYesBut.Caption = "Grrrr!" & Chr(13) & Chr(10) & "Register now!!"
    cmdYesBut.Picture = FacePic.Picture
    tmrFollow.Enabled = True  'Start the button moving!
End Sub

Private Sub cmdYesBut_Click()
    ExitDo = True
    'Stop the Do..Loop from running, though you don't need
    'this if you're going to unload the form like this
    
    tmrFollow.Enabled = False
    MsgBox "                 Why thankyou! :)" & Chr(13) & Chr(10) & _
        "Killer Button and Images made By GEEZA" & Chr(13) & Chr(10) & _
        "                GEEZA1@aol.com", vbApplicationModal + vbInformation, "hehe!"
    Unload Me
    End
End Sub

Private Sub DirectionPic(ByVal LeftRight As Integer, UpDown As Integer)
    If UpDown = 0 Then
        If LeftRight = 1 Then
            cmdYesBut.Picture = LeftPic.Picture
        Else
            cmdYesBut.Picture = RightPic.Picture
        End If
    ElseIf UpDown = 1 Then
        If LeftRight = 0 Then
            cmdYesBut.Picture = UpPic.Picture
        ElseIf LeftRight = 1 Then
            cmdYesBut.Picture = Up_LeftPic.Picture
        Else
            cmdYesBut.Picture = Up_RightPic.Picture
        End If
    Else
        If LeftRight = 0 Then
            cmdYesBut.Picture = DownPic.Picture
        ElseIf LeftRight = 1 Then
            cmdYesBut.Picture = Down_LeftPic.Picture
        Else
            cmdYesBut.Picture = Down_RightPic.Picture
        End If
    End If
End Sub

Private Sub tmrFollow_Timer()
    Dim Direction As Integer
    
    DoEvents
    GetCursorPos XnY
    XnY.X = ScaleX(XnY.X, vbPixels, vbTwips) 'Change the dimensions from Pixels
    XnY.Y = ScaleY(XnY.Y, vbPixels, vbTwips) 'to Twips

    If (cmdYesBut.Left + cmdYesBut.Width / 2 + Me.Left > XnY.X + a_Radius) Or (cmdYesBut.Left + cmdYesBut.Width / 2 + Me.Left < XnY.X - a_Radius) Then
        'Movement in X
        If cmdYesBut.Left < 0 Then
            cmdYesBut.Left = 0
            Me.Left = Me.Left - 15  'push window
            Direction = 1 'left
        ElseIf cmdYesBut.Left + cmdYesBut.Width > Me.Width Then
            cmdYesBut.Left = Me.Width - cmdYesBut.Width
            Me.Left = Me.Left + 15  'push window
            Direction = 2 'right
        Else
            If cmdYesBut.Left + cmdYesBut.Width / 2 + Me.Left < XnY.X Then
                cmdYesBut.Left = cmdYesBut.Left + 30
                Direction = 2
            Else
                cmdYesBut.Left = cmdYesBut.Left - 30
                Direction = 1
            End If
        End If
    End If
        
    If Not (cmdYesBut.Top + cmdYesBut.Height / 2 + Me.Top > XnY.Y - a_Radius) Or (cmdYesBut.Top + cmdYesBut.Height / 2 + Me.Top > XnY.Y + a_Radius) Then
        If cmdYesBut.Top < 0 Then
            cmdYesBut.Top = 0
            Me.Top = Me.Top - 15
            Call DirectionPic(Direction, 1)
        ElseIf cmdYesBut.Top + cmdYesBut.Height > Me.Height Then
            cmdYesBut.Top = Me.Height - cmdYesBut.Height
            Me.Top = Me.Top + 15
            Call DirectionPic(Direction, 2)
        Else
            If cmdYesBut.Top + cmdYesBut.Height / 2 + Me.Top < XnY.Y Then
                cmdYesBut.Top = cmdYesBut.Top + 30
                Call DirectionPic(Direction, 2)
            Else
                cmdYesBut.Top = cmdYesBut.Top - 30
                Call DirectionPic(Direction, 1)
            End If
        End If
    ElseIf Direction = 0 Then
        'Within a_Radius twips of the center
        '(pretty long IF statements huh?!)
        tmrFollow.Enabled = False
        Call StickButton(Me, cmdYesBut, cmdYesBut.Width / 2, cmdYesBut.Height / 2)
    Else: Call DirectionPic(Direction, 0)
    End If
End Sub

Private Sub StickButton(ByVal Form As Form, ByVal Button As CommandButton, DpX As Long, DpY As Long)
    Do
        DoEvents    'So it doesn't 'Hang' the program
        GetCursorPos XnY
        XnY.X = ScaleX(XnY.X, vbPixels, vbTwips)
        XnY.Y = ScaleY(XnY.Y, vbPixels, vbTwips)
        
        If XnY.X - DpX <= Form.Left Then
            Button.Left = 0
            Me.Left = Me.Left - 15
            If XnY.Y - DpY <= Form.Top Then
                Button.Top = 0
                Me.Top = Me.Top - 15
                Button.Picture = Up_LeftPic.Picture
            ElseIf XnY.Y + (Button.Height - DpY) >= Form.Top + Form.Height Then
                Button.Top = Form.Height - Button.Height
                Me.Top = Me.Top + 15
                Button.Picture = Down_LeftPic.Picture
            Else
                Button.Top = XnY.Y - DpY - Form.Top
                Button.Picture = LeftPic.Picture
            End If
        ElseIf XnY.X + Button.Width - DpX >= Form.Left + Form.Width Then
            Button.Left = Form.Width - Button.Width
            Me.Left = Me.Left + 15
            If XnY.Y - DpY <= Form.Top Then
                Button.Top = 0
                Me.Top = Me.Top - 15
                Button.Picture = Up_RightPic.Picture
            ElseIf XnY.Y + (Button.Height - DpY) >= Form.Top + Form.Height Then
                Button.Top = Form.Height - Button.Height
                Me.Top = Me.Top + 15
                Button.Picture = Down_RightPic.Picture
            Else
                Button.Top = XnY.Y - DpY - Form.Top
                Button.Picture = RightPic.Picture
            End If
        Else
            Button.Left = XnY.X - DpX - Form.Left
            If XnY.Y - DpY <= Form.Top Then
                Button.Top = 0
                Me.Top = Me.Top - 15
                Button.Picture = UpPic.Picture
            ElseIf XnY.Y + (Button.Height - DpY) >= Form.Top + Form.Height Then
                Button.Top = Form.Height - Button.Height
                Me.Top = Me.Top + 15
                Button.Picture = DownPic.Picture
            Else
                Button.Top = XnY.Y - DpY - Form.Top
                Button.Picture = FacePic.Picture
            End If
        End If
        If ExitDo Then Exit Do
    Loop  'Stick the button to the cursor until ExitDo is true
    'And they wont be able to click anything else on the form!! hehe!
End Sub


'Ok, if the user clicks the 'No' button, then the killer
'button awakens, and goes after your cursor!
'When it finally gets the cursor, it grabs hold and wont
'let you click anything else on the form.

'It does seem a bit wrong, now that the button is always
'confined to the form due to graphical buttons not working
'well with no parents (0, well how would you feel? lol)
'So you can either use both parts, or just one by modifying
'some of the lines of code, and removing parts

'Well, i hope you like it!  i know everyone laughed when
'i showed them version 1 i made in college!
'All i ask is that you please put a reference to me
'into any projects you use this with e.g. in your about
'boxes, or even better in any message box that comes up
'after clicking the KilerButton
'('Killer button written by GEEZA!  GEEZA1@aol.com, the caffine junkie!')
'Oh and maybe change the pictures if you have time,
'everyone's projects have to be a little different :)

'this has taken me ages to write,
'and a lot of swearing too :)
'i'll have to improve the engines sometime though...

'Anyone have any ideas about merging the Java 'Eyes' effect
'into the project, for the face?

'Ok, it's 3am, i'm gonna go before i pass out,
'and computer keyboards arn't comfy, trust me!
'night Y 'all!
