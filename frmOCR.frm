VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOCR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "A Novice Attempt At Optical Character Recognition"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11460
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   11460
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   2880
      TabIndex        =   19
      Top             =   5760
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CheckBox cbEnableAutoTrain 
      Caption         =   "Enable Auto-Train"
      Height          =   195
      Left            =   5640
      TabIndex        =   16
      Top             =   5520
      Width           =   2295
   End
   Begin VB.TextBox txtMaxBadSectors 
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Text            =   "0"
      Top             =   5760
      Width           =   735
   End
   Begin VB.TextBox txtToleranceLevel 
      Height          =   375
      Left            =   2880
      TabIndex        =   12
      Text            =   "30"
      Top             =   5280
      Width           =   735
   End
   Begin VB.CheckBox cbEnableDebug 
      Caption         =   "Enable Debug [Slower]"
      Height          =   195
      Left            =   5640
      TabIndex        =   11
      Top             =   5280
      Value           =   1  'Checked
      Width           =   2535
   End
   Begin VB.PictureBox pbDebug2 
      BackColor       =   &H8000000E&
      Height          =   2655
      Left            =   8400
      ScaleHeight     =   173
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   189
      TabIndex        =   10
      Top             =   3240
      Width           =   2895
   End
   Begin VB.PictureBox pbDebug1 
      BackColor       =   &H8000000E&
      Height          =   2655
      Left            =   8400
      ScaleHeight     =   173
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   189
      TabIndex        =   9
      Top             =   240
      Width           =   2895
   End
   Begin VB.TextBox txtSampleRate 
      Height          =   405
      Left            =   120
      TabIndex        =   7
      Text            =   "2,2"
      Top             =   5280
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Train"
      Height          =   735
      Left            =   4800
      TabIndex        =   4
      Top             =   960
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reset"
      Height          =   735
      Left            =   4800
      TabIndex        =   3
      Top             =   120
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Output:"
      Height          =   2295
      Left            =   4800
      TabIndex        =   2
      Top             =   2880
      Width           =   3375
      Begin VB.Label lblStatus 
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1920
         Width           =   3015
      End
      Begin VB.Label Label2 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label1 
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdRecognize 
      Caption         =   "Recognize"
      Height          =   735
      Left            =   4800
      TabIndex        =   1
      Top             =   1800
      Width           =   3375
   End
   Begin VB.PictureBox OCRImage 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000E&
      Height          =   5055
      Left            =   120
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   333
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   301
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label Label7 
      Caption         =   "Match:"
      Height          =   255
      Left            =   8400
      TabIndex        =   18
      Top             =   3000
      Width           =   2895
   End
   Begin VB.Label Label6 
      Caption         =   "Current Letter:"
      Height          =   255
      Left            =   8400
      TabIndex        =   17
      Top             =   0
      Width           =   2895
   End
   Begin VB.Label Label5 
      Caption         =   "Max # of non-matching sectors."
      Height          =   375
      Left            =   960
      TabIndex        =   15
      Top             =   5760
      Width           =   1815
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      Caption         =   "Tolerance Level"
      Height          =   255
      Left            =   3720
      TabIndex        =   13
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "<-- Enter sampling rate.    (1=Highest Rate) (x,y)"
      Height          =   495
      Left            =   960
      TabIndex        =   8
      Top             =   5280
      Width           =   1815
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmOCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Analyzer As New OCRAnalyzer
Private Sub cbEnableDebug_Click()
    If Me.cbEnableDebug Then
        EnableDebug
    Else
        DisableDebug
    End If
End Sub
Public Function InDebugMode() As Boolean
    If Me.cbEnableDebug Then
        Let InDebugMode = True
    End If
End Function
Private Sub EnableDebug()
    With Me
        If frmOCR.WindowState <> vbMaximized Then
            Do Until .Width = 11580
                Let .Width = .Width + 5
                DoEvents
            Loop
        End If
        Let .Label6.Visible = True
        Let .Label7.Visible = True
        Let .pbDebug1.Visible = True
        Let .pbDebug2.Visible = True
    End With
    Let frmOCR.Caption = frmOCR.Caption & "   [DEBUG MODE]"
End Sub
Private Sub DisableDebug()
    With Me
        Let .Label6.Visible = False
        Let .Label7.Visible = False
        Let .pbDebug1.Visible = False
        Let .pbDebug2.Visible = False
        If frmOCR.WindowState <> vbMaximized Then
            Do Until .Width = 8370
                Let .Width = .Width - 5
                DoEvents
            Loop
        End If
    End With
    Let frmOCR.Caption = Replace(frmOCR.Caption, "   [DEBUG MODE]", "")
End Sub
'8370
'11580
Private Sub cmdRecognize_Click()
    Let Me.Label1.Caption = Analyzer.Analyze(Me.OCRImage)
End Sub
Private Sub Command1_Click()
    Let Me.Label2.Caption = ""
    Let Me.Label1.Caption = ""
    Set Me.OCRImage = LoadPicture
    Me.pbDebug1.Cls
    Me.pbDebug2.Cls
    Me.ProgressBar1.Value = 0
End Sub
Private Sub Command2_Click()
    Let Me.Label1.Caption = Analyzer.Analyze(Me.OCRImage, True)
End Sub
Private Sub Form_Resize()
    Static LastMaximized As Boolean
    If Me.WindowState <> vbMaximized And LastMaximized Then
        If Me.InDebugMode Then
            Let Me.Width = 11580
        Else
            Let Me.Width = 8370
        End If
    End If
    If Me.WindowState = vbMaximized Then
        Let LastMaximized = True
    Else
        Let LastMaximized = False
    End If
End Sub
Private Sub OCRImage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Or Button = vbRightButton Then
        If Me.ProgressBar1.Value <> 0 Then
            Let Me.ProgressBar1.Value = 0
        End If
    End If
    If Button = vbLeftButton Then
        'Stop
        OCRImage.DrawWidth = 8
        OCRImage.PSet (X, Y)
        OCRImage.DrawWidth = 1
    ElseIf Button = vbRightButton Then
        OCRImage.DrawWidth = 30
        OCRImage.PSet (X, Y), QBColor(15)
        OCRImage.DrawWidth = 1
    End If
End Sub
