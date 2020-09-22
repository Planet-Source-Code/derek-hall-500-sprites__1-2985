VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "BitBlt Masking Example"
   ClientHeight    =   2940
   ClientLeft      =   180
   ClientTop       =   -1665
   ClientWidth     =   3600
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   196
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   240
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAFPS 
      Height          =   285
      Left            =   4440
      TabIndex        =   11
      Top             =   1800
      Width           =   615
   End
   Begin VB.HScrollBar hscBallCount 
      Height          =   195
      LargeChange     =   10
      Left            =   1560
      Max             =   500
      TabIndex        =   8
      Top             =   2040
      Value           =   100
      Width           =   1455
   End
   Begin VB.HScrollBar hscGlobalSpeed 
      Height          =   195
      LargeChange     =   5
      Left            =   120
      Max             =   99
      TabIndex        =   7
      Top             =   2040
      Value           =   25
      Width           =   1335
   End
   Begin VB.CheckBox chkCollision 
      Caption         =   "Collision Detection"
      Height          =   255
      Left            =   5160
      TabIndex        =   5
      Top             =   1800
      Width           =   2175
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   195
      LargeChange     =   5
      Left            =   5160
      Max             =   50
      Min             =   1
      TabIndex        =   2
      Top             =   2040
      Value           =   20
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox picBackground 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      Height          =   675
      Left            =   60
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   66
      TabIndex        =   0
      Top             =   0
      Width           =   1050
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   11040
      Top             =   8280
   End
   Begin VB.PictureBox picFace 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   465
      Left            =   7200
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   1
      Top             =   1800
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label lblBallCount 
      Alignment       =   2  'Center
      Caption         =   "Sprite Count = 100"
      Height          =   255
      Left            =   1560
      TabIndex        =   10
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label lblMaxSpeed 
      Alignment       =   2  'Center
      Caption         =   "Max Speed = 25"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "Sprite speed per Frame"
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label lblCollision 
      Height          =   255
      Left            =   5160
      TabIndex        =   6
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label txtFrames 
      Alignment       =   2  'Center
      Caption         =   "Frames Per Sec"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   7680
      TabIndex        =   3
      ToolTipText     =   "Milliseconds"
      Top             =   2040
      Width           =   735
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuStart 
         Caption         =   "Start"
      End
      Begin VB.Menu mnuResetSprites 
         Caption         =   "Reset Sprites"
      End
      Begin VB.Menu mnuStop 
         Caption         =   "Stop"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this can be improved alot, but will give some of
'you an idea on how to make things happen
  
  
  Const SpriteSize = 30
  Const BallCounter = 500
  Dim BallCount As Integer
  Dim GlobalSpeed As Integer
  Dim CurrentPicX(BallCounter) As Integer
  Dim CurrentPicY(BallCounter) As Integer
  Dim BallDead(BallCounter) As Boolean
  Dim i As Integer, j As Integer
  Dim Xmode(BallCounter) As Boolean, Ymode(BallCounter) As Boolean
  Dim Xspeed(BallCounter) As Integer, Yspeed(BallCounter) As Integer
  Dim GoNow As Boolean

Private Sub Form_Load()
  GoNow = False
  frmMain.Height = 6600
  frmMain.Width = 9600
  GlobalSpeed = 25
  BallCount = 100
  Form_Resize
End Sub

Private Sub Form_Resize()
  If frmMain.WindowState = vbMinimized Then Exit Sub
  picBackground.Left = 4
  picBackground.Top = 4
  If frmMain.Height < 3000 Then frmMain.Height = 3000
  If frmMain.Width < 3000 Then frmMain.Width = 3000
  picBackground.Height = (frmMain.Height / 15) - 90
  picBackground.Width = (frmMain.Width / 15) - 16

  HScroll1.Top = picBackground.Height + 8
  txtFrames.Top = (HScroll1.Top + HScroll1.Height) + 3
  hscGlobalSpeed.Top = HScroll1.Top
  lblMaxSpeed.Top = txtFrames.Top
  hscBallCount.Top = HScroll1.Top
  lblBallCount.Top = txtFrames.Top
  chkCollision.Top = HScroll1.Top - 2
  Label1.Top = txtFrames.Top + 5
  txtAFPS.Top = txtFrames.Top
End Sub


Private Sub Form_Unload(Cancel As Integer)
  Unload Me
  End
End Sub

Private Sub hscGlobalSpeed_Change()
  GlobalSpeed = hscGlobalSpeed.Value
  lblMaxSpeed.Caption = "Max Speed = " & hscGlobalSpeed.Value + 1
End Sub

Private Sub HScroll1_Change()
    
    Timer1.Interval = 1000 / HScroll1.Value
    txtFrames.Caption = HScroll1.Value & " Frames Per Sec"

End Sub

Private Sub HScBallCount_Change()
  BallCount = hscBallCount.Value
  lblBallCount.Caption = "Sprite Count = " & BallCount + 1
End Sub

Private Sub mnuResetSprites_Click()
  For i = 0 To BallCounter
    BallDead(i) = False
    CurrentPicX(i) = 0
    CurrentPicY(i) = 0
  Next i
End Sub

Private Sub mnuStart_Click()

    Select Case mnuStart.Caption
        Case "Start"
            mnuStart.Caption = "Stop"
            'Timer1.Enabled = True
            mnuResetSprites_Click
            GoNow = True
        Case "Stop"
            mnuStart.Caption = "Start"
            Timer1.Enabled = False
            picBackground = LoadPicture
            GoNow = False
    End Select

fullspeed
End Sub

Private Sub mnuStop_Click()
mnuStop.Checked = Not mnuStop.Checked
End Sub



Sub CollisionCheck(CurrentSprite As Integer)
    Dim SpriteCounter As Integer, CurrentGapY As Integer, CurrentGapX As Integer

    For SpriteCounter = 0 To BallCount
        If Not SpriteCounter = CurrentSprite And Not BallDead(SpriteCounter) Then
            If CurrentPicY(CurrentSprite) > CurrentPicY(SpriteCounter) Then CurrentGapY = CurrentPicY(CurrentSprite) - CurrentPicY(SpriteCounter)
            If CurrentPicX(CurrentSprite) > CurrentPicX(SpriteCounter) Then CurrentGapX = CurrentPicX(CurrentSprite) - CurrentPicX(SpriteCounter)
            If CurrentPicY(CurrentSprite) < CurrentPicY(SpriteCounter) Then CurrentGapY = CurrentPicY(SpriteCounter) - CurrentPicY(CurrentSprite)
            If CurrentPicX(CurrentSprite) < CurrentPicX(SpriteCounter) Then CurrentGapX = CurrentPicX(SpriteCounter) - CurrentPicX(CurrentSprite)
           
            If CurrentGapX < SpriteSize And CurrentGapX > -SpriteSize Then
                If CurrentGapY < SpriteSize And CurrentGapY > -SpriteSize Then
                   ' txtWeHit = txtWeHit + 1
                    BallDead(SpriteCounter) = True
                    'BallDead(CurrentSprite) = True
                End If
            End If
        End If
      Next SpriteCounter
End Sub


Sub fullspeed()
Dim AFPS As Integer, CurrentTime As Variant
AFPS = 0
CurrentTime = Timer
Do While GoNow = True
 picBackground = LoadPicture
    For i = 0 To BallCount
        If Not BallDead(i) Then
            If CurrentPicY(i) > picBackground.Height - SpriteSize Then
                Ymode(i) = False
                Yspeed(i) = Int((GlobalSpeed * Rnd) + 1)
            End If
            If CurrentPicY(i) < 1 Then
               Ymode(i) = True
                Yspeed(i) = Int((GlobalSpeed * Rnd) + 1)
            End If
            If CurrentPicX(i) > picBackground.Width - SpriteSize Then
                Xmode(i) = False
                Xspeed(i) = Int((GlobalSpeed * Rnd) + 1)
            End If
            If CurrentPicX(i) < 1 Then
                Xmode(i) = True
                Xspeed(i) = Int((GlobalSpeed * Rnd) + 1)
            End If
            If Ymode(i) Then
                CurrentPicY(i) = CurrentPicY(i) + Yspeed(i)
            Else
               CurrentPicY(i) = CurrentPicY(i) - Yspeed(i)
            End If
            If Xmode(i) Then
                CurrentPicX(i) = CurrentPicX(i) + Xspeed(i)
            Else
                CurrentPicX(i) = CurrentPicX(i) - Xspeed(i)
            End If
            If chkCollision.Value = 1 Then CollisionCheck i
            If Not BallDead(i) Then
                BitBlt picBackground.hDC, CurrentPicX(i), CurrentPicY(i), 27, 27, picFace.hDC, 27, 0, SRCAND
                BitBlt picBackground.hDC, CurrentPicX(i), CurrentPicY(i), 27, 27, picFace.hDC, 0, 0, SRCPAINT
            End If
            Label1.Caption = Timer
        End If ' BallDead
    Next i
    AFPS = AFPS + 1 ' Actual Frames Per Second
    If Timer > CurrentTime + 0.99 Then
        txtAFPS = AFPS
        AFPS = 0
        CurrentTime = Timer
    End If
    DoEvents
Loop
   picBackground.Refresh

End Sub
