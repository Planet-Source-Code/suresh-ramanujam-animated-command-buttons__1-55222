VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Animated Command Buttons - By Suresh Ramanujam"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Wana Quit?"
      Height          =   975
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2040
      Top             =   2280
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5175
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Button With Caption"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3240
         TabIndex        =   4
         Top             =   1320
         Width           =   1725
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Button Without Caption"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   1995
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************
'                      Animated Command Buttons            *
'                                By                        *
'                          Suresh Ramanujam                *
'                                                          *
'                                                          *
'                               Contact:                   *
'                       www.planetsourcecode.com           *
'                    sureshramanujam@rediffmail.com        *
'                                                          *
'***********************************************************

Dim i As Integer    'counter to control frames

Private Sub Command1_Click()
 'Alert
 MsgBox "Gotcha!! At last i made it!!!!!!", , "Animated Command Buttons"
End Sub

Private Sub Command2_Click()
'Alert
ans = MsgBox("Howzat folks, really kooooooooollll????. Wait for more code. Bye!!!!", vbYesNo, "Animated Command Buttons")

If ans = vbYes Then
    End
End If
End Sub

Private Sub Form_Load()
i = 1
'Alert
MsgBox "This is a simple code, which teaches you, how to animate ur command buttons. For more info. go through the readme.txt file. Please don forget to rate my code.", , "Animated Command Buttons"
End Sub

' This animates the command button
Private Sub Timer1_Timer()
If i > 12 Then              'check the counter
    i = 1                   're-initialise
Else
    'load frame to command button1
    Command1.Picture = LoadPicture(App.Path & "\mail" & i & ".bmp")
    
    'load frame to command button1
    Command2.Picture = LoadPicture(App.Path & "\mail" & i & ".bmp")
    
    i = i + 1
End If
End Sub
