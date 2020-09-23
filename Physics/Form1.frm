VERSION 5.00
Begin VB.Form Issues 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Find it out"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7485
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3120
      Top             =   1920
   End
   Begin VB.CommandButton Command2 
      Caption         =   "!"
      Height          =   375
      Left            =   2400
      TabIndex        =   10
      Top             =   2400
      Width           =   255
   End
   Begin VB.PictureBox pic2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   5160
      ScaleHeight     =   615
      ScaleWidth      =   2295
      TabIndex        =   7
      Top             =   480
      Width           =   2295
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2760
      ScaleHeight     =   615
      ScaleWidth      =   2175
      TabIndex        =   6
      Top             =   480
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "The air time"
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "The range"
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "The apex"
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "The launch angle"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "The launch velocity"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "Method 2"
      Height          =   255
      Left            =   5160
      TabIndex        =   9
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Method 1"
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "What do you want to find out?"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2145
   End
End
Attribute VB_Name = "Issues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Showing As Integer
Private Sub Command1_Click(Index As Integer)
    pic2.Cls
    Select Case Index
        Case 1
            Showing = 1
            pic.Picture = LoadPicture(App.Path & "\vel1.bmp")
            pic2.Print "You can also find the velocity by"
            pic2.Print "using the apex and range."
            pic2.Print "But this is not documented here."
        Case 2
            Showing = 2
            pic.Picture = LoadPicture(App.Path & "\ang1.bmp")
            pic2.Print "    X"
        Case 3
            Showing = 3
            pic.Picture = LoadPicture(App.Path & "\apex1.bmp")
            pic2.Print "    X"
        Case 4
            Showing = 4
            pic.Picture = LoadPicture(App.Path & "\ran1.bmp")
            pic2.Print "    X"
        Case 5
            Showing = 5
            pic.Picture = LoadPicture(App.Path & "\air1.bmp")
            pic2.Print "    X"
    End Select
End Sub

Private Sub Command2_Click()
    MsgBox "Since some calculators works with DEG as standard you don't need " & vbCrLf & "to convert the degrees into radians. Check what your calculator does before " & vbCrLf & "calculating. If you get wrong answers you must convert the degrees into radians by doing so: " & vbCrLf & "degrees / 180 * PI."
End Sub

Private Sub Form_Load()
    Command1_Click (1)
    Me.Left = Trajectory.Left
End Sub

Public Function Iknow()
    Iknow = Showing
End Function

Private Sub Timer1_Timer()
    Me.Left = Trajectory.Left
    Me.Top = (Trajectory.Top + Trajectory.Height)
End Sub
