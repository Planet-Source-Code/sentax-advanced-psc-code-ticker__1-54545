VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PSC - Code Ticker Options"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6180
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
   ScaleHeight     =   2745
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboRefreshMin 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3570
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   1875
      Width           =   690
   End
   Begin VB.CheckBox chkDBasic 
      Caption         =   "Dark Basic"
      Height          =   240
      Left            =   3435
      TabIndex        =   14
      Top             =   1560
      Width           =   1860
   End
   Begin VB.CheckBox chkLISP 
      Caption         =   "LISP"
      Height          =   240
      Left            =   3435
      TabIndex        =   13
      Top             =   1320
      Width           =   1860
   End
   Begin VB.CheckBox chkdotNet 
      Caption         =   ".NET"
      Height          =   240
      Left            =   3435
      TabIndex        =   12
      Top             =   1080
      Width           =   1860
   End
   Begin VB.CheckBox chkColdFusion 
      Caption         =   "Cold Fusion"
      Height          =   240
      Left            =   3435
      TabIndex        =   11
      Top             =   840
      Width           =   1860
   End
   Begin VB.CheckBox chkPHP 
      Caption         =   "PHP"
      Height          =   240
      Left            =   3435
      TabIndex        =   10
      Top             =   585
      Width           =   1860
   End
   Begin VB.CheckBox chkDelphi 
      Caption         =   "Delphi"
      Height          =   240
      Left            =   3435
      TabIndex        =   9
      Top             =   330
      Width           =   1860
   End
   Begin VB.CheckBox chkPerl 
      Caption         =   "Perl"
      Height          =   240
      Left            =   810
      TabIndex        =   8
      Top             =   1560
      Width           =   1860
   End
   Begin VB.CheckBox chkSQL 
      Caption         =   "SQL"
      Height          =   240
      Left            =   810
      TabIndex        =   7
      Top             =   1320
      Width           =   1860
   End
   Begin VB.CheckBox chkASP 
      Caption         =   "ASP"
      Height          =   240
      Left            =   810
      TabIndex        =   6
      Top             =   1080
      Width           =   1860
   End
   Begin VB.CheckBox chkCPP 
      Caption         =   "C/C++"
      Height          =   240
      Left            =   810
      TabIndex        =   5
      Top             =   840
      Width           =   1860
   End
   Begin VB.CheckBox chkJavaScript 
      Caption         =   "Java/Javascript"
      Height          =   240
      Left            =   810
      TabIndex        =   4
      Top             =   585
      Width           =   1860
   End
   Begin VB.CheckBox chkVB 
      Caption         =   "Visual Basic"
      Height          =   240
      Left            =   810
      TabIndex        =   2
      Top             =   330
      Width           =   1860
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   3090
      TabIndex        =   1
      Top             =   2325
      Width           =   1155
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   345
      Left            =   1800
      TabIndex        =   0
      Top             =   2325
      Width           =   1155
   End
   Begin VB.Label Label2 
      Caption         =   "Refresh Selected World Every:             Minutes"
      Height          =   225
      Left            =   915
      TabIndex        =   15
      Top             =   1950
      Width           =   4140
   End
   Begin VB.Label Label1 
      Caption         =   "Please select the worlds you want to download."
      Height          =   315
      Left            =   1035
      TabIndex        =   3
      Top             =   30
      Width           =   4110
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
If frmMain.bLoading = True Then End
Unload Me
End Sub

Private Sub cmdSave_Click()
'CHECK WHICH CHECK BOXES ARE CHECKED... 2 MANY CHECK WORDS IN THERE. :)

GatherWorlds

If Len(sCodeWorldURLS) = 0 Then
    MsgBox "You need to have at least 1 world selected!", vbInformation, "Error"
    Exit Sub
End If

Unload Me
End Sub


Public Sub GatherWorlds()
'FUNCTION THAT GENERATES THE STRING THAT HOLDS THE ACTIVE WORLDS

frmMain.bLoading = True

sCodeWorldURLS = ""

frmMain.cboWorlds.Clear
frmMain.cboWorlds.AddItem "All Pre-Selected Worlds"

frmMain.lsvwNew.ListItems.Clear

If chkVB.Value = 1 Then
    sCodeWorldURLS = sCodeWorldURLS & "|1"
    frmMain.cboWorlds.AddItem GetWorldIDText(1)
End If

If chkJavaScript.Value = 1 Then
    sCodeWorldURLS = sCodeWorldURLS & "|2"
    frmMain.cboWorlds.AddItem GetWorldIDText(2)
End If

If chkCPP.Value = 1 Then
    sCodeWorldURLS = sCodeWorldURLS & "|3"
    frmMain.cboWorlds.AddItem GetWorldIDText(3)
End If

If chkASP.Value = 1 Then
    sCodeWorldURLS = sCodeWorldURLS & "|4"
    frmMain.cboWorlds.AddItem GetWorldIDText(4)
End If

If chkSQL.Value = 1 Then
    sCodeWorldURLS = sCodeWorldURLS & "|5"
    frmMain.cboWorlds.AddItem GetWorldIDText(5)
End If

If chkPerl.Value = 1 Then
    sCodeWorldURLS = sCodeWorldURLS & "|6"
    frmMain.cboWorlds.AddItem GetWorldIDText(6)
End If

If chkDelphi.Value = 1 Then
    sCodeWorldURLS = sCodeWorldURLS & "|7"
    frmMain.cboWorlds.AddItem GetWorldIDText(7)
End If

If chkPHP.Value = 1 Then
    sCodeWorldURLS = sCodeWorldURLS & "|8"
    frmMain.cboWorlds.AddItem GetWorldIDText(8)
End If

If chkColdFusion.Value = 1 Then
    sCodeWorldURLS = sCodeWorldURLS & "|9"
    frmMain.cboWorlds.AddItem GetWorldIDText(9)
End If

If chkdotNet.Value = 1 Then
    sCodeWorldURLS = sCodeWorldURLS & "|10"
    frmMain.cboWorlds.AddItem GetWorldIDText(10)
End If

If chkLISP.Value = 1 Then
    sCodeWorldURLS = sCodeWorldURLS & "|13"
    frmMain.cboWorlds.AddItem GetWorldIDText(13)
End If

If chkDBasic.Value = 1 Then
    sCodeWorldURLS = sCodeWorldURLS & "|14"
    frmMain.cboWorlds.AddItem GetWorldIDText(14)
End If

frmMain.cboWorlds.ListIndex = 0

'REMOVES THE LEADING "|" FROM THE STRING

If Mid(sCodeWorldURLS, 1, 1) = "|" Then sCodeWorldURLS = Mid(sCodeWorldURLS, 2, Len(sCodeWorldURLS) - 1)


frmMain.iRefreshSelWorld = FormatNumber(cboRefreshMin.Text, 0) * 60
frmMain.iRefreshSelWorldTimer = 0

frmMain.bLoading = False

End Sub

Private Sub Form_Load()

'LOADS COMBO BOX

cboRefreshMin.AddItem "1"
cboRefreshMin.AddItem "3"
cboRefreshMin.AddItem "5"
cboRefreshMin.AddItem "10"
cboRefreshMin.AddItem "15"
cboRefreshMin.AddItem "20"
cboRefreshMin.AddItem "30"
cboRefreshMin.AddItem "60"

cboRefreshMin.ListIndex = 0

If sCodeWorldURLS <> "" Then
    CheckFoundWorlds
End If
End Sub

Public Sub CheckFoundWorlds()
'ANALYZES THE WORLD CODE STRING AND THEN CHECKS THE CHECKBOXES
Dim aResult() As String, i As Long

aResult = Split(sCodeWorldURLS, "|")

For i = LBound(aResult) To UBound(aResult)
    
    Select Case aResult(i)
    
        Case "1":
            chkVB.Value = 1
        Case "2":
            chkJavaScript.Value = 1
        Case "3":
            chkCPP.Value = 1
        Case "4":
            chkASP.Value = 1
        Case "5":
            chkSQL.Value = 1
        Case "6":
            chkPerl.Value = 1
        Case "7":
            chkDelphi.Value = 1
        Case "8":
            chkPHP.Value = 1
        Case "9":
            chkColdFusion.Value = 1
        Case "10":
            chkdotNet.Value = 1
        Case "13":
            chkLISP.Value = 1
        Case "14":
            chkDBasic.Value = 1
            
    End Select
    
Next

'FINDS THE REFRESH INTERVAL AND SELECT IT IN COMBO BOX
For i = 0 To frmOptions.cboRefreshMin.ListCount - 1
    cboRefreshMin.ListIndex = i
    If cboRefreshMin.Text = frmMain.iRefreshSelWorld / 60 Then Exit For
Next

End Sub
