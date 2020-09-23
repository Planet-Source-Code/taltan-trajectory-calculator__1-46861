VERSION 5.00
Begin VB.Form Trajectory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trajectory calculator"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11655
   Icon            =   "Trajectory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   514
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   777
   Begin VB.CheckBox Check2 
      Caption         =   "Formulas follow"
      Height          =   255
      Left            =   240
      TabIndex        =   29
      Top             =   6840
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Formulas"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Reset"
      Height          =   255
      Left            =   5040
      TabIndex        =   27
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   ">>"
      Height          =   255
      Left            =   6360
      TabIndex        =   24
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "<<"
      Height          =   255
      Left            =   3720
      TabIndex        =   23
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Timer fall 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3120
      Top             =   5880
   End
   Begin VB.PictureBox plot 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H008080FF&
      Height          =   4815
      Left            =   0
      ScaleHeight     =   319
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   775
      TabIndex        =   22
      Top             =   0
      Width           =   11655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2640
      TabIndex        =   21
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   8760
      Top             =   6480
   End
   Begin VB.Frame Frame1 
      Caption         =   "Projectile data"
      Height          =   2295
      Left            =   3960
      TabIndex        =   6
      Top             =   5280
      Width           =   7455
      Begin VB.Label CuRa 
         BackColor       =   &H80000009&
         Caption         =   "0"
         Height          =   255
         Left            =   1560
         TabIndex        =   20
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Current range"
         Height          =   255
         Left            =   1560
         TabIndex        =   19
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label CuAl 
         BackColor       =   &H80000009&
         Caption         =   "0"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Current altitude"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label MaxAl 
         BackColor       =   &H80000009&
         Caption         =   "0"
         Height          =   255
         Left            =   5880
         TabIndex        =   16
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label TimeImp 
         BackColor       =   &H80000009&
         Caption         =   "0"
         Height          =   255
         Left            =   4440
         TabIndex        =   15
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label TimeAir 
         BackColor       =   &H80000009&
         Caption         =   "0"
         Height          =   255
         Left            =   3000
         TabIndex        =   14
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label speed 
         BackColor       =   &H80000009&
         Caption         =   "0"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lenn 
         BackColor       =   &H80000009&
         Caption         =   "0"
         Height          =   255
         Left            =   1560
         TabIndex        =   12
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Max altitude (apex)"
         Height          =   375
         Left            =   5880
         TabIndex        =   11
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Time in air"
         Height          =   255
         Left            =   3000
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Time to impact"
         Height          =   255
         Left            =   4440
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Range"
         Height          =   375
         Left            =   1560
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Speed"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "FIRE"
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   7200
      Width           =   1215
   End
   Begin VB.TextBox ang 
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Text            =   "30"
      Top             =   6240
      Width           =   1215
   End
   Begin VB.TextBox vel 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Text            =   "100"
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "m    |"
      Height          =   255
      Left            =   11250
      TabIndex        =   26
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label LenM 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   10560
      TabIndex        =   25
      Top             =   4920
      Width           =   90
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Angle:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   6240
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Velocity:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   5400
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Mesurements are in meters, seconds and radians."
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   4920
      Width           =   3495
   End
End
Attribute VB_Name = "Trajectory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ContiniueFlight As Boolean
Dim Meter As Integer
Private FTime As Double

Dim FlightPath As Flight

Private Type Flight
    X As Integer
    Y As Integer
    OldX As Integer
    OldY As Integer
End Type

Private Sub Check1_Click()
    If Check1.value = 1 Then
        Issues.Show
        frmExample.Show
    Else
        Issues.Hide
        frmExample.Hide
    End If
End Sub

Private Sub Check2_Click()
    If Check2 = "1" Then
        Issues.Timer1 = True
        frmExample.Timer1 = True
        
    Else
        Issues.Timer1 = False
        frmExample.Timer1 = False
        
    End If
End Sub

Private Sub Command1_Click()
    
    Dim a, b, c, d, e
    
    If ang > 90 Then
        MsgBox "Angle must be under 90 degrees.", vbExclamation, "Error"
        Exit Sub
    ElseIf ang < 1 Then
        MsgBox "Angle must be over 0 degrees.", vbExclamation, "Error"
        Exit Sub
    End If
    plot.Width = 100 * vel
    With FlightPath
        .OldX = 0
        .OldY = 0
        .X = 0
        .Y = 0
    End With
    plot.Cls
    ContiniueFlight = True
     Command1.Enabled = False
     Command2.Enabled = True
    ' ang.Enabled = False
    ' vel.Enabled = False
    
    a = vel * Cos(ConvDegToRad(ang))
    b = vel * Sin(ConvDegToRad(ang))
    c = b / 9.81
    d = 2 * c
    e = a * d
    
    MaxAl = Apex(vel, ConvDegToRad(ang))
    lenn = e
    speed = vel & " m/s"
    TimeAir = TotalTimeInAir(vel, ConvDegToRad(ang))
    If TimeAir > 2 Then
        TimeImp = Int(TotalTimeInAir(vel, ConvDegToRad(ang)))
        Timer1 = True
    End If
    FTime = 0
    '  fall.Enabled = True
    plot.Cls
    Do While FlightPath.Y > -1 And ContiniueFlight = True
        Update

        Loop
ContiniueFlight = True
    End Sub

Private Function ConvDegToRad(ByVal Deg As Double) As Double
    ConvDegToRad = Deg / 180 * PI
End Function

Public Function Apex(ByVal MuzzleVelocity As Double, ByVal FireAngle As Double) As Double
    Apex = (MuzzleVelocity ^ 2 * Sin(FireAngle) ^ 2) / (2 * 9.81)
End Function

Public Function TotalTimeInAir(ByVal MuzzleVelocityT As Double, ByVal FireAngleT As Double) As Double
    TotalTimeInAir = (2 * MuzzleVelocityT * Sin(FireAngleT) / 9.81)
End Function

Private Sub Command2_Click()
    Timer1 = False
    Command1.Enabled = True
    Command2.Enabled = False
    ContiniueFlight = False
End Sub

Private Sub Command3_Click()
    Issues.Show
End Sub

Private Sub Command4_Click()
    If plot.Left = 0 Then
        Exit Sub
    End If
    plot.Move plot.Left + 5, plot.Top
    Meter = Meter - "4,84610552763819095477386934673" ' hmm... lol :D
    LenM.Caption = Meter
End Sub

Private Sub Command5_Click()
    plot.Move plot.Left - 5, plot.Top
    Meter = Meter + "4,84610552763819095477386934673"
    LenM.Caption = Meter
End Sub

Private Sub Command6_Click()
    plot.Move 0, 0
    Meter = "771,5"
    LenM.Caption = "771"
End Sub

Private Sub Form_Load()
    Meter = "771,5"
    ContiniueFlight = True
    Load Issues
    Load frmExample
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmExample.Timer1 = False
    Issues.Timer1 = False
    Unload frmExample
    Unload Issues
    ContiniueFlight = False
End Sub

Private Sub Timer1_Timer()
    TimeImp = TimeImp - 0.01
    If TimeImp = 0 Then
        Timer1 = False
        Command1.Enabled = True
        Command2.Enabled = False
        ang.Enabled = True
        vel.Enabled = True
    End If
End Sub

Public Sub Update()
    On Error Resume Next
    'Calculate the X and Y co-ordinates of the projectile.
    FlightPath.Y = vel * FTime * Sin(ConvDegToRad(ang)) - (0.5 * 9.81 * FTime ^ 2)
    FlightPath.X = (vel * Cos(ConvDegToRad(ang))) * FTime
    
    plot.Line (FlightPath.OldX, plot.Height - FlightPath.OldY)-(FlightPath.X, plot.Height - FlightPath.Y)
    
    FlightPath.OldX = FlightPath.X
    FlightPath.OldY = FlightPath.Y
    
    FTime = FTime + 0.01
    DoEvents
End Sub
