VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fading Form"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   720
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   1733
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "That was a cool effect huh? Click the close button to see it fade out! A great effect for splash screens."
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'I did not write the transparency module.
'It is from another author on psource
'They deserve credit

Dim goingDown As Boolean, transparency As Integer

Private Sub Command1_Click()
goingDown = True
Me.Enabled = True 'stops the user from clicking on
                  'something after fading starts
Timer1.Enabled = True
End Sub

Private Sub Form_Load()
MakeTransparent Me.hWnd, 0
transparency = 0
goingDown = False
Me.Enabled = False
End Sub

Private Sub Timer1_Timer()
If goingDown = True Then
    If transparency > 0 Then
        MakeTransparent Me.hWnd, transparency
        transparency = transparency - 5
    Else
        Unload Me
    End If
Else
    If transparency < 256 Then
        MakeTransparent Me.hWnd, transparency
        transparency = transparency + 5
    Else
        MakeOpaque Me.hWnd
        Me.Enabled = True
        Timer1.Enabled = False
    End If
End If
End Sub
