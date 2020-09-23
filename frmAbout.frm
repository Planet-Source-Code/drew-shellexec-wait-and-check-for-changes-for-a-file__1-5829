VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Vertical Splitter Sample"
   ClientHeight    =   3645
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5235
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2515.843
   ScaleMode       =   0  'User
   ScaleWidth      =   4915.936
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   480
      Left            =   240
      Picture         =   "frmAbout.frx":0442
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   480
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   435
      Left            =   4035
      TabIndex        =   0
      Top             =   2730
      Width           =   1080
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Visual Code Server."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   165
      TabIndex        =   7
      Top             =   2940
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.StotzerSoftware.com"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   495
      MouseIcon       =   "frmAbout.frx":0884
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   2745
      Width           =   2340
   End
   Begin VB.Label Label3 
      Caption         =   $"frmAbout.frx":0B8E
      Height          =   825
      Left            =   165
      TabIndex        =   6
      Top             =   2745
      Width           =   3765
   End
   Begin VB.Label Label1 
      Caption         =   "Developed by Stotzer Software 1999"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   915
      MouseIcon       =   "frmAbout.frx":0C6D
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2160
      Width           =   3390
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   4803.25
      Y1              =   1770.408
      Y2              =   1770.408
   End
   Begin VB.Label lblDescription 
      Caption         =   $"frmAbout.frx":0F77
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   915
      TabIndex        =   2
      Top             =   810
      Width           =   4050
   End
   Begin VB.Label lblTitle 
      Caption         =   "ShellExec, Wait, and Check"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   915
      TabIndex        =   3
      Top             =   315
      Width           =   3930
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   4789.164
      Y1              =   1770.408
      Y2              =   1780.761
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'******************************************************************
' Windows API Declarations for ShellExec
Private Const SW_SHOW           As Integer = 5
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'******************************************************************
    

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Label1_Click()
    ShellExec "http://www.StotzerSoftware.com"
End Sub

Private Sub Label2_Click()
    ShellExec "http://www.StotzerSoftware.com"
End Sub


Private Sub ShellExec(sFilename As String, Optional oForm As Variant)
    Dim lRC                 As Long
    
    '--- SHELLEXECUTE THE DOCUMENT TO THE ASSOCIATED APP ---'
    If IsMissing(oForm) Then
        lRC = ShellExecute(0&, "open", sFilename, vbNullString, CurDir$, SW_SHOW)
    Else
        lRC = ShellExecute(oForm.hwnd, "open", sFilename, vbNullString, CurDir$, SW_SHOW)
    End If
    If lRC < 32 Then
        MsgBox "Unable to open this file: " & vbLf & "    " & sFilename, vbExclamation, "Error Opening File"
        Exit Sub
    End If
    
End Sub


