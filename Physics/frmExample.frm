VERSION 5.00
Begin VB.Form frmExample 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Example"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3075
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmExample.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   3075
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2520
      Top             =   2520
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Forward"
      Height          =   255
      Left            =   720
      TabIndex        =   10
      ToolTipText     =   "This will forward the angle and velocity into "
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "GO"
      Height          =   1215
      Left            =   2640
      TabIndex        =   8
      Top             =   600
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Caption         =   "What you know"
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2415
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   960
         TabIndex        =   7
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   960
         TabIndex        =   6
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Range:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Angle:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Velocity:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.Label findout 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   45
   End
   Begin VB.Label what 
      Caption         =   "Enter the number you know and those you do not know you leave blank."
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2865
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim biS
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    On Error GoTo errh:
    
    If biS = "110" Then
        findout = "( " & Text1 & "^2 * sin " & "(2 * " & Text2 & ")" & " ) / 9.81" & vbCrLf & _
        ((Text1 ^ 2) * Sin(2 * ConvDegToRad(Text2))) / 9.81
    ElseIf biS = "101" Then
        a = ConvRadToDeg(Atn(((9.81 * Text3) / (Text1 ^ 2)) / Sqr(-((9.81 * Text3) / (Text1 ^ 2)) * ((9.81 * Text3) / (Text1 ^ 2)) + 1)))
        findout = "sin-1(( 9.81 * " & Text3 & ") / ( " & Text1 & "^2))" & vbCrLf & (a / 2)
        
    ElseIf biS = "011" Then
        findout = "sqr( (" & Text3 & " * 9.81) / ( sin (2 * " & Text2 & ") )" & vbCrLf & _
        Sqr((Text3 * 9.81) / (Sin(2 * ConvDegToRad(Text2))))
    End If
    
errh:
    If Err Then: MsgBox Err.Description
End Sub

Private Sub Command3_Click()
Trajectory.ang = Text2
Trajectory.vel = Text1
End Sub

Private Sub Form_Load()
    biS = "000"
    Me.Left = (Issues.Left + Issues.Width)
    Me.Top = Issues.Top
End Sub

Private Sub Text1_Change()
    If Text1 = "" Then
        a = Mid(biS, 2, 2)
        biS = "0" & a
    Else
        a = Mid(biS, 2, 2)
        biS = "1" & a
    End If
    'MsgBox biS
    DoBox
End Sub

Private Sub Text2_Change()
    a = biS
    If Text2 = "" Then
        biS = Mid(a, 1, 1) & "0" & Mid(a, 3, 1)
    Else
        biS = Mid(a, 1, 1) & "1" & Mid(a, 3, 1)
    End If
    DoBox
    'MsgBox biS
End Sub

Private Sub Text3_Change()
    If Text3 = "" Then
        a = Mid(biS, 1, 2)
        biS = a & "0"
    Else
        a = Mid(biS, 1, 2)
        biS = a & "1"
    End If
    'MsgBox biS
    DoBox
End Sub

Sub DoBox()
    Select Case biS
            
        Case "000"
            Text1.Enabled = True
            Text2.Enabled = True
            Text3.Enabled = True
            
        Case "100"
            Text2.Enabled = True
            Text3.Enabled = True
            
        Case "110"
            Text3.Enabled = False
            
        Case "101"
            Text2.Enabled = False
            
        Case "001"
            Text1.Enabled = True
            Text2.Enabled = True
            
        Case "010"
            Text1.Enabled = True
            Text3.Enabled = True
            
        Case "011"
            Text1.Enabled = False
    End Select
    
End Sub

Private Sub Timer1_Timer()
    Me.Left = (Issues.Left + Issues.Width)
    Me.Top = Issues.Top
End Sub
