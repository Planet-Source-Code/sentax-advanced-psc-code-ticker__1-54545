VERSION 5.00
Begin VB.Form frmWait 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Please Wait..."
   ClientHeight    =   990
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4935
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblWorldID 
      Alignment       =   2  'Center
      Caption         =   "Visual Basic..."
      Height          =   255
      Left            =   75
      TabIndex        =   1
      Top             =   510
      Width           =   4785
   End
   Begin VB.Label lblStatusTitle 
      Alignment       =   2  'Center
      Caption         =   "Downloading World..."
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   90
      Width           =   2805
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const HWND_TOPMOST = -1

Private Sub Form_Load()

'THIS FORMS PURPOSE IN LIFE IS TO WAIT TO BE CLOSED AND LOOK GOOD AT IT

SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

End Sub
