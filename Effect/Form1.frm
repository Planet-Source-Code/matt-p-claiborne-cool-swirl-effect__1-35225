VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Swirl"
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   ScaleHeight     =   580
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   676
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Swirl"
      Height          =   975
      Left            =   6360
      TabIndex        =   17
      Top             =   7560
      Width           =   2175
      Begin VB.CheckBox Check4 
         Caption         =   "Normalize Speeds"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   18
         Text            =   "2000"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Numer Of ""Dots"""
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Options"
      Height          =   975
      Left            =   8640
      TabIndex        =   13
      Top             =   7560
      Width           =   1455
      Begin VB.CheckBox Check3 
         Caption         =   "Invert Speed"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Invert Spin"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   450
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Solid Colors"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   135
      Left            =   3720
      Max             =   20
      TabIndex        =   9
      Top             =   7920
      Value           =   3
      Width           =   2415
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   135
      Left            =   3720
      Max             =   40
      TabIndex        =   8
      Top             =   8280
      Value           =   7
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Color"
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   7560
      Width           =   3375
      Begin VB.HScrollBar HSColor 
         Height          =   135
         Index           =   2
         Left            =   720
         Max             =   255
         TabIndex        =   4
         Top             =   720
         Value           =   255
         Width           =   2535
      End
      Begin VB.HScrollBar HSColor 
         Height          =   135
         Index           =   1
         Left            =   720
         Max             =   255
         TabIndex        =   3
         Top             =   480
         Value           =   55
         Width           =   2535
      End
      Begin VB.HScrollBar HSColor 
         Height          =   135
         Index           =   0
         Left            =   720
         Max             =   255
         TabIndex        =   2
         Top             =   240
         Value           =   55
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "Blue:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Green:"
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
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Red:"
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
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   7335
      Left            =   120
      ScaleHeight     =   485
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   661
      TabIndex        =   0
      Top             =   120
      Width           =   9975
      Begin VB.Timer Timer2 
         Interval        =   1
         Left            =   3600
         Top             =   1560
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   3000
         Top             =   960
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Speed"
      Height          =   975
      Left            =   3600
      TabIndex        =   10
      Top             =   7560
      Width           =   2655
      Begin VB.Label Label5 
         Caption         =   "Circular Speed:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   120
         TabIndex        =   12
         Top             =   560
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Outward Speed:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   120
         TabIndex        =   11
         Top             =   200
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type SwirlParticles
Angle As Single
AngleADD As Single
X As Integer
Y As Integer
CurX As Integer
CurY As Integer
Red As Byte
Green As Byte
Blue As Byte
Rad As Single
RadADD As Single
End Type

Dim Swirl() As SwirlParticles
Dim TMax As Integer
Dim PicWidth As Integer
Dim PicHeight As Integer
Dim PicWH As Integer
Dim PicWHd2 As Integer
Dim FPS As Integer
Dim Running As Boolean

Private Sub Check1_Click()
SwirlColorsRed
SwirlColorsGreen
SwirlColorsBlue
End Sub

Private Sub Check2_Click()
Dim X As Integer

For X = 0 To TMax
Swirl(X).AngleADD = -Swirl(X).AngleADD
Next

End Sub


Private Sub Check3_Click()
Dim X As Integer

For X = 0 To TMax
Swirl(X).RadADD = -Swirl(X).RadADD
Next
End Sub

Private Sub Check4_Click()

SwirlAngleRad

End Sub

Private Sub Form_Load()

Dim X As Integer

PicWidth = Picture1.ScaleWidth
PicHeight = Picture1.ScaleHeight
PicWH = Sqr(PicWidth ^ 2 + PicHeight ^ 2)
Randomize Time
TMax = 1000
ReDim Swirl(TMax) As SwirlParticles
For X = 0 To TMax
Swirl(X).Angle = Int(Rnd * 361)
Swirl(X).AngleADD = (Rnd * 0.07)
Swirl(X).X = PicWidth / 2
Swirl(X).Y = PicHeight / 2
Swirl(X).Rad = Rnd * PicWH / 2
Swirl(X).RadADD = Rnd * 3
Swirl(X).Red = (Rnd * 55)
Swirl(X).Green = (Rnd * 55)
Swirl(X).Blue = (Rnd * 255)
Next
Text1.Text = TMax
Me.Show
'Running = True
PicWHd2 = PicWH / 2

'SwirlScreen

Timer2.Enabled = True
End Sub




Private Sub HSColor_Change(Index As Integer)
Select Case Index
Case 0
SwirlColorsRed
Case 1
SwirlColorsGreen
Case 2
SwirlColorsBlue
End Select
End Sub


Private Sub HSColor_Scroll(Index As Integer)
Select Case Index
Case 0
SwirlColorsRed
Case 1
SwirlColorsGreen
Case 2
SwirlColorsBlue
End Select
End Sub

Private Sub HScroll1_Change()
SwirlAngle
End Sub

Private Sub HScroll1_Scroll()
SwirlAngle
End Sub

Private Sub HScroll2_Change()
SwirlRad
End Sub

Private Sub HScroll2_Scroll()
SwirlRad
End Sub

Private Sub Text1_Change()
On Error GoTo ERROR
Dim X As Integer

If Not Text1.Text = "" Then
If Int(Text1.Text) = Text1.Text Then


    Timer2.Enabled = False
    TMax = Text1.Text
    ReDim Swirl(TMax) As SwirlParticles
    
    For X = 0 To TMax
        Swirl(X).Angle = Int(Rnd * 361)
        Swirl(X).X = PicWidth / 2
        Swirl(X).Y = PicHeight / 2
        Swirl(X).Rad = Rnd * PicWH / 2
    Next
    
    SwirlAngle
    SwirlRad
    SwirlColorsRGB
    Picture1.Cls
    Timer2.Enabled = True

End If
End If

Exit Sub
ERROR:
Text1.Text = TMax

End Sub

Private Sub Timer1_Timer()
Me.Caption = "Swirl - " & FPS & " FPS"
FPS = 0
End Sub

Private Sub Timer2_Timer()
Dim X As Integer
For X = 0 To TMax

Swirl(X).Angle = Swirl(X).Angle + Swirl(X).AngleADD
If Swirl(X).Angle > 360 Then Swirl(X).Angle = 0


If Swirl(X).Rad > PicWHd2 Then Swirl(X).Rad = 0
If Swirl(X).Rad < 0 Then Swirl(X).Rad = PicWHd2
Swirl(X).Rad = Swirl(X).Rad + Swirl(X).RadADD

SetPixel Picture1.hdc, Swirl(X).CurX, Swirl(X).CurY, 0
Swirl(X).CurX = Swirl(X).X + Swirl(X).Rad * (Cos(Swirl(X).Angle))
Swirl(X).CurY = Swirl(X).Y + Swirl(X).Rad * (Sin(Swirl(X).Angle))

SetPixel Picture1.hdc, Swirl(X).CurX, Swirl(X).CurY, RGB(Swirl(X).Red, Swirl(X).Green, Swirl(X).Blue)
Next
FPS = FPS + 1
End Sub

Private Sub SwirlColorsRed()
Dim X As Integer

If Check1.Value = 0 Then
    Randomize Time
    For X = 0 To TMax
        Swirl(X).Red = Rnd * HSColor(0).Value
    Next
Else
    For X = 0 To TMax
        Swirl(X).Red = HSColor(0).Value
    Next
End If

End Sub

Private Sub SwirlColorsGreen()
Dim X As Integer

If Check1.Value = 0 Then
    Randomize Time
    For X = 0 To TMax
        Swirl(X).Green = Rnd * HSColor(1).Value
    Next
Else
    For X = 0 To TMax
        Swirl(X).Green = HSColor(1).Value
    Next
End If

End Sub

Private Sub SwirlColorsBlue()
Dim X As Integer

If Check1.Value = 0 Then
    Randomize Time
    For X = 0 To TMax
        Swirl(X).Blue = Rnd * HSColor(2).Value
    Next
Else
    For X = 0 To TMax
        Swirl(X).Blue = HSColor(2).Value
    Next
End If

End Sub

Private Sub SwirlAngle()
Dim X As Integer
If Check4.Value = 0 Then
    If Check2.Value = 0 Then
        For X = 0 To TMax
            Swirl(X).AngleADD = (Rnd * HScroll1.Value) / 100
        Next
    Else
        For X = 0 To TMax
            Swirl(X).AngleADD = -((Rnd * HScroll1.Value) / 100)
        Next
    End If
Else
    If Check2.Value = 0 Then
        For X = 0 To TMax
            Swirl(X).AngleADD = (HScroll1.Value) / 100
        Next
    Else
        For X = 0 To TMax
            Swirl(X).AngleADD = -((HScroll1.Value) / 100)
        Next
    End If
End If
End Sub

Private Sub SwirlRad()
Dim X As Integer
If Check4.Value = 0 Then
    If Check3.Value = 0 Then
        For X = 0 To TMax
            Swirl(X).RadADD = (Rnd * HScroll2.Value) / 2
        Next
    Else
        For X = 0 To TMax
            Swirl(X).RadADD = -((Rnd * HScroll2.Value) / 2)
        Next
    End If
Else
    If Check3.Value = 0 Then
        For X = 0 To TMax
            Swirl(X).RadADD = (HScroll2.Value) / 2
        Next
    Else
        For X = 0 To TMax
            Swirl(X).RadADD = -((HScroll2.Value) / 2)
        Next
    End If
End If
End Sub
Private Sub SwirlAngleRad()
Dim X As Integer
If Check4.Value = 0 Then
    If Check3.Value = 0 Then
        For X = 0 To TMax
            Swirl(X).RadADD = (Rnd * HScroll2.Value) / 2
            Swirl(X).AngleADD = (Rnd * HScroll1.Value) / 100
        Next
    Else
        For X = 0 To TMax
            Swirl(X).RadADD = -((Rnd * HScroll2.Value) / 2)
            Swirl(X).AngleADD = -(Rnd * HScroll1.Value) / 100
        Next
    End If
Else
    If Check3.Value = 0 Then
        For X = 0 To TMax
            Swirl(X).RadADD = (HScroll2.Value) / 2
            Swirl(X).AngleADD = (HScroll1.Value) / 100
        Next
    Else
        For X = 0 To TMax
            Swirl(X).RadADD = -((HScroll2.Value) / 2)
            Swirl(X).AngleADD = -(HScroll1.Value) / 100
        Next
    End If
End If
End Sub
Private Sub SwirlColorsRGB()
Dim X As Integer

If Check1.Value = 0 Then
    Randomize Time
    For X = 0 To TMax
        Swirl(X).Red = Rnd * HSColor(0).Value
        Swirl(X).Green = Rnd * HSColor(1).Value
        Swirl(X).Blue = Rnd * HSColor(2).Value
    Next
Else
    For X = 0 To TMax
        Swirl(X).Red = HSColor(0).Value
        Swirl(X).Green = HSColor(1).Value
        Swirl(X).Blue = HSColor(2).Value
    Next
End If

End Sub
