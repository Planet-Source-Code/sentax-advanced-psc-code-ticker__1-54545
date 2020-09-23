VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "PSC - Ticker"
   ClientHeight    =   3915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7470
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimerOneSec 
      Interval        =   1000
      Left            =   6990
      Top             =   1785
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3915
      Left            =   0
      ScaleHeight     =   3885
      ScaleWidth      =   7440
      TabIndex        =   1
      Top             =   0
      Width           =   7470
      Begin VB.CommandButton cmdOptions 
         Caption         =   "Options"
         Height          =   345
         Left            =   6045
         TabIndex        =   8
         Top             =   285
         Width           =   1365
      End
      Begin VB.ComboBox cboWorlds 
         Height          =   360
         Left            =   30
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   285
         Width           =   6015
      End
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   6420
         Top             =   1590
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin VB.PictureBox picCodeHolder 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3210
         Left            =   30
         ScaleHeight     =   3180
         ScaleWidth      =   7350
         TabIndex        =   3
         Top             =   660
         Width           =   7380
         Begin MSComctlLib.ListView lsvwNew 
            Height          =   3225
            Left            =   -15
            TabIndex        =   6
            Top             =   -30
            Width           =   7380
            _ExtentX        =   13018
            _ExtentY        =   5689
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Num"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "ID"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "World"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Title"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Author"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Time"
               Object.Width           =   1235
            EndProperty
         End
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         X1              =   6900
         X2              =   7050
         Y1              =   180
         Y2              =   180
      End
      Begin VB.Label lblMini 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   6810
         TabIndex        =   5
         Top             =   -15
         Width           =   330
      End
      Begin VB.Label lblClose 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "X"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   7125
         TabIndex        =   4
         Top             =   -15
         Width           =   330
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Planet Source Code - Code Ticker"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   -15
         TabIndex        =   2
         Top             =   -15
         Width           =   7470
      End
   End
   Begin VB.Label lblStandBy 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PSC Ticker: 0 Code Samples "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   3690
      TabIndex        =   0
      Top             =   1665
      Width           =   2460
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public bLoading As Boolean
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const HWND_TOPMOST = -1

Private iTimerMinimize As Integer
Public iRefreshSelWorldTimer As Integer
Public iRefreshSelWorld As Integer
Private bShowForm As Boolean

Private Sub cboWorlds_Change()
iTimerMinimize = 0
If bLoading = False Then ChangeWorldViews
End Sub

Private Sub cboWorlds_Click()
iTimerMinimize = 0
If bLoading = False Then ChangeWorldViews
End Sub

Private Sub cmdOptions_Click()
'LOADS OPTIONS FROM BUTTON AND THEN UPDATES LIST

frmOptions.Show 1

cboWorlds.ListIndex = 0
ChangeWorldViews

End Sub

Private Sub Form_Load()
bLoading = True

'LOAD OPTIONS DIALOG
LoadOptions

'LOAD FORM IN STANDBY POSITION
lblStandBy.Visible = True
picBack.Visible = False
Me.Height = 240
Me.Width = 2460
Me.Move Screen.Width - Me.Width - 1200, 75
lblStandBy.Left = 0
lblStandBy.Top = 0

'SET ON TOP OF EVERYTHING
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

'DOWNLOAD THE PSC FILES
ScanHTTPFiles True

bLoading = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
iTimerMinimize = 0
End Sub

Private Sub lblClose_Click()
End
End Sub

Private Sub lblMini_Click()
'SENDS FOR TO STANDBY

lblStandBy.Visible = True
picBack.Visible = False
Me.Height = 240
Me.Width = 2460
Me.Move Screen.Width - Me.Width - 1200, 75
'TimerOneSec.Enabled = False
lblStandBy.Left = 0
lblStandBy.Top = 0
bShowForm = False
End Sub

Private Sub lblOlderTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
iTimerMinimize = 0
End Sub

Private Sub lblStandBy_Click()
'SENDS FORM TO VIEW MODE

lblStandBy.Visible = False
picBack.Visible = True
Me.Height = 3915
Me.Width = 7470
Me.Move Screen.Width - Me.Width - 1200, 75
iTimerMinimize = 0
bShowForm = True
End Sub

Private Sub lblTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
iTimerMinimize = 0
End Sub

Private Sub lsvwNew_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
iTimerMinimize = 0
End Sub

Private Sub lsvwNew_DblClick()
'OPENS NEW IE WINDOW WITH CODE LINK

lblMini_Click

Dim iSel As Integer
iSel = lsvwNew.SelectedItem

Shell "C:\Program Files\Internet Explorer\iexplore.exe " & "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=" & lsvwNew.ListItems(iSel).ListSubItems(1).Text & "&lngWId=" & GetWorldText2ID(lsvwNew.ListItems(iSel).ListSubItems(2).Text), vbNormalFocus

End Sub

Private Sub lsvwNew_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
iTimerMinimize = 0
End Sub

Private Sub lsvwOlder_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
iTimerMinimize = 0
End Sub

Private Sub picBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
iTimerMinimize = 0
End Sub

Private Sub picCodeHolder_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
iTimerMinimize = 0
End Sub

Private Sub TimerOneSec_Timer()

'ACTIVATED EVERY SECOND AND CHECKS FOR UNACTIVE MOUSE TO MAKE FORM STANDBY
'ALSO COUNTS UP FOR REFRESHING WORLD SELECTED IN COMBO BOX

If iTimerMinimize = 5 And bShowForm = True Then
    lblStandBy.Visible = True
    picBack.Visible = False
    Me.Height = 240
    Me.Width = 2460
    Me.Move Screen.Width - Me.Width - 1200, 75
    lblStandBy.Left = 0
    lblStandBy.Top = 0
    bShowForm = False
Else
    iTimerMinimize = iTimerMinimize + 1
    bShowForm = True
End If

lblStandBy.Caption = "PSC Ticker: " & lsvwNew.ListItems.Count & " code samples"

If iRefreshSelWorldTimer >= iRefreshSelWorld And lblStandBy.Visible = True And bLoading = False Then
    'CHANGES/REFRESHES THE CURRENT WORLD SELECTED IN COMBO BOX
    ChangeWorldViews
    iRefreshSelWorldTimer = 0
Else
    iRefreshSelWorldTimer = iRefreshSelWorldTimer + 1
End If

End Sub

Public Sub LoadOptions()
'LOAD OPTIONS DIALOG
'I THOUGHT I WAS GOING TO DO MORE IN THIS FUNCTION :)

frmOptions.Show 1
End Sub

Public Sub ScanHTTPFiles(bFirstScan As Boolean)
'THIS FUNCTION SPLITS THE STRING GENERATED THROUGH THE OPTIONS DIALOG
'THEN SCANS EACH WORLD ONE BY ONE

Dim aResult() As String, i As Long

aResult = Split(sCodeWorldURLS, "|")

If bFirstScan = True Then
    frmWait.Visible = True
End If

For i = LBound(aResult) To UBound(aResult)
    
    ScanFile FormatNumber(aResult(i), 0)
    
Next

Unload frmWait

End Sub

Function ScanFile(lngWID As Integer)
'THIS FUNCTION DOWNLOADS THE ACTUAL TICKER PAGE AND SAVES TO HD
'AFTER SUCCESS OF DOWNLOAD IT THEN PARSES THE FILE FOR CODE SAMPLES

If Dir(App.Path & "\CODEPAGES\", vbDirectory) = "" Then
    MkDir (App.Path & "\CODEPAGES\")
End If

frmWait.lblStatusTitle.Caption = "Downloading World..."
frmWait.lblWorldID.Caption = GetWorldIDText(lngWID)
DoEvents

If DownloadFile("http://www.planet-source-code.com/vb/linktous/ScrollingCode.asp?lngWId=" & lngWID & "&blnHideChannelSubscribe=true&blnLaunchLinkInNewWindow=true&blnShowTickerWorldTitle=false", App.Path & "\CODEPAGES\PSC-ID" & lngWID & ".asp") = True Then
    ParseFile App.Path & "\CODEPAGES\PSC-ID" & lngWID & ".asp", GetWorldIDText(lngWID)
Else
    MsgBox "There was a problem downloading one of the code files!"
    End
End If

lblStandBy.Caption = "PSC Ticker: " & lsvwNew.ListItems.Count & " code samples"

End Function

Public Function DownloadFile(strURL As String, _
                             strDestination As String, _
                             Optional UserName As String = Empty, _
                             Optional Password As String = Empty) _
                             As Boolean

'THIS FUNCTION WAS TAKEN FROM THE PSC CODE FOUND HERE:
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=7335&lngWId=1
'THANKS FOR THE GOOD CODE
'I HAD TO COMMENT OUT THE MOVIE AND LABEL UPDATES JUST TO DO THE DOWNLOAD

' Funtion DownloadFile: Download a file via HTTP
'
' Author:   Jeff Cockayne
'
' Inputs:   strURL String; the source URL of the file
'           strDestination; valid Win95/NT path to where you want it
'           (i.e. "C:\Program Files\My Stuff\Purina.pdf")
'
' Returns:  Boolean; Was the download successful?

Const CHUNK_SIZE As Long = 1024 ' Download chunk size
Const ROLLBACK As Long = 4096   ' Bytes to roll back on resume
                                ' You can be less conservative,
                                ' and roll back less, but I
                                ' don't recommend it.
Dim bData() As Byte             ' Data var
Dim blnResume As Boolean        ' True if resuming download
Dim intFile As Integer          ' FreeFile var
Dim lngBytesReceived As Long    ' Bytes received so far
Dim lngFileLength As Long       ' Total length of file in bytes
Dim lngX                        ' Temp long var
Dim sglLastTime As Single          ' Time last chunk received
Dim sglRate As Single           ' Var to hold transfer rate
Dim sglTime As Single           ' Var to hold time remaining
Dim strFile As String           ' Temp filename var
Dim strHeader As String         ' HTTP header store
Dim strHost As String           ' HTTP Host

On Local Error GoTo InternetErrorHandler

' Start with Cancel flag = False
CancelSearch = False

' Get just filename (without dirs) for display
strFile = ReturnFileOrFolder(strDestination, True)
strHost = ReturnFileOrFolder(strURL, True, True)
              
SourceLabel = Empty
TimeLabel = Empty
ToLabel = Empty
RateLabel = Empty

' Pre-open the AVI
'With Animation1
'    .AutoPlay = True
'    .Open App.Path & "\DOWNLD2.AVI"
'End With

' Show the download status form
'Show
' Move form into view
'Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2

StartDownload:

If blnResume Then
    StatusLabel = "Resuming download..."
    lngBytesReceived = lngBytesReceived - ROLLBACK
    If lngBytesReceived < 0 Then lngBytesReceived = 0
Else
    StatusLabel = "Getting file information..."
End If
' Give the system time to update the form gracefully
DoEvents

' Download file
With Inet1
    .URL = strURL
    .UserName = UserName
    .Password = Password
    ' GET file, sending the magic resume input header...
    .Execute , "GET", , "Range: bytes=" & CStr(lngBytesReceived) & "-" & vbCrLf
    
    ' While initiating connection, yield CPU to Windows
    While .StillExecuting
        DoEvents
        ' If user pressed Cancel button on StatusForm
        ' then fail, cancel, and exit this download
        If CancelSearch Then GoTo ExitDownload
    Wend

    'StatusLabel = "Saving:"
    'SourceLabel = FitText(SourceLabel, strHost & " from " & .RemoteHost)
    'ToLabel = FitText(ToLabel, strDestination)

    ' Get first header ("HTTP/X.X XXX ...")
    strHeader = .GetHeader
End With

' Trap common HTTP response codes
Select Case Mid(strHeader, 10, 3)
    Case "200"  ' OK
        ' If resuming, however, this is a failure
        If blnResume Then
            ' Delete partially downloaded file
            Kill strDestination
            ' Prompt
            If MsgBox("The server is unable to resume this download." & _
                      vbCr & vbCr & _
                      "Do you want to continue anyway?", _
                      vbExclamation + vbYesNo, _
                      "Unable to Resume Download") = vbYes Then
                    ' Yes - continue anyway:
                    ' Set resume flag to False
                    blnResume = False
                Else
                    ' No - cancel
                    CancelSearch = True
                    GoTo ExitDownload
                End If
            End If
            
    Case "206"  ' 206=Partial Content, which is GREAT when resuming!
    
    Case "204"  ' No content
        MsgBox "Nothing to download!", _
               vbInformation, _
               "No Content"
        CancelSearch = True
        GoTo ExitDownload
        
    Case "401"  ' Not authorized
        MsgBox "Authorization failed!", _
               vbCritical, _
               "Unauthorized"
        CancelSearch = True
        GoTo ExitDownload
    
    Case "404"  ' File Not Found
        MsgBox "The file, " & _
               """" & Inet1.URL & """" & _
               " was not found!", _
               vbCritical, _
               "File Not Found"
        CancelSearch = True
        GoTo ExitDownload
        
    Case vbCrLf ' Empty header
        MsgBox "Cannot establish connection." & vbCr & vbCr & _
               "Check your Internet connection and try again.", _
               vbExclamation, _
               "Cannot Establish Connection"
        CancelSearch = True
        GoTo ExitDownload
        
    Case Else
        ' Miscellaneous unexpected errors
        strHeader = Left(strHeader, InStr(strHeader, vbCr))
        If strHeader = Empty Then strHeader = "<nothing>"
        MsgBox "The server returned the following response:" & _
               vbCr & vbCr & _
               strHeader, _
               vbCritical, _
               "Error Downloading File"
        CancelSearch = True
        GoTo ExitDownload
End Select

' Get file length with "Content-Length" header request
If blnResume = False Then
    ' Set timer for gauging download speed
    sglLastTime = Timer - 1
    strHeader = Inet1.GetHeader("Content-Length")
    lngFileLength = Val(strHeader)
    If lngFileLength = 0 Then
        GoTo ExitDownload
    End If
End If

' Check for available disk space first...
' If on a physical or mapped drive. Can't with a UNC path.
If Mid(strDestination, 2, 2) = ":\" Then
    If DiskFreeSpace(Left(strDestination, _
                          InStr(strDestination, "\"))) < lngFileLength Then
        ' Not enough free space to download file
        MsgBox "There is not enough free space on disk for this file." _
               & vbCr & vbCr & "Please free up some disk space and try again.", _
               vbCritical, _
               "Insufficient Disk Space"
        GoTo ExitDownload
    End If
End If

' Prepare display
'
' Progress Bar
'With ProgressBar
'    .Value = 0
'    .Max = lngFileLength
'End With

' Give system a chance to show AVI
DoEvents

' Reset bytes received counter if not resuming
If blnResume = False Then lngBytesReceived = 0


On Local Error GoTo FileErrorHandler

' Create destination directory, if necessary
strHeader = ReturnFileOrFolder(strDestination, False)
If Dir(strHeader, vbDirectory) = Empty Then
    MkDir strHeader
End If

' If no errors occurred, then spank the file to disk
intFile = FreeFile()        ' Set intFile to an unused file.
' Open a file to write to.
Open strDestination For Binary Access Write As #intFile
' If resuming, then seek byte position in downloaded file
' where we last left off...
If blnResume Then Seek #intFile, lngBytesReceived + 1
Do
    ' Get chunks...
    bData = Inet1.GetChunk(CHUNK_SIZE, icByteArray)
    Put #intFile, , bData   ' Put it into our destination file
    If CancelSearch Then Exit Do
    lngBytesReceived = lngBytesReceived + UBound(bData, 1) + 1
    sglRate = lngBytesReceived / (Timer - sglLastTime)
    sglTime = (lngFileLength - lngBytesReceived) / sglRate
    TimeLabel = FormatTime(sglTime) & _
                   " (" & _
                   FormatFileSize(lngBytesReceived) & _
                   " of " & _
                   FormatFileSize(lngFileLength) & _
                   " copied)"
    RateLabel = FormatFileSize(sglRate, "###.0") & "/Sec"
    'ProgressBar.Value = lngBytesReceived
    'Me.Caption = Format((lngBytesReceived / lngFileLength), "##0%") & _
    '             " of " & strFile & " Completed"
Loop While UBound(bData, 1) > 0       ' Loop while there's still data...
Close #intFile

ExitDownload:
' Success if the # of bytes transferred = content length
If lngBytesReceived = lngFileLength Then
    StatusLabel = "Download completed!"
    DownloadFile = True
Else
    If Dir(strDestination) = Empty Then
        CancelSearch = True
    Else
        ' Resume? (If not cancelled)
        If CancelSearch = False Then
            If MsgBox("The connection with the server was reset." & _
                      vbCr & vbCr & _
                      "Click ""Retry"" to resume downloading the file." & _
                      vbCr & "(Approximate time remaining: " & FormatTime(sglTime) & ")" & _
                      vbCr & vbCr & _
                      "Click ""Cancel"" to cancel downloading the file.", _
                      vbExclamation + vbRetryCancel, _
                      "Download Incomplete") = vbRetry Then
                    ' Yes
                    blnResume = True
                    GoTo StartDownload
            End If
        End If
    End If
    ' No or unresumable failure:
    ' Delete partially downloaded file
    If Not Dir(strDestination) = Empty Then Kill strDestination
    DownloadFile = False
End If

CleanUp:
' Close AVI
'Animation1.Close

' Make sure that the Internet connection is closed...
Inet1.Cancel
' ...and exit this function
'Unload Me

Exit Function

InternetErrorHandler:
    ' Err# 9 occurs when UBound(bData,1) < 0
    If Err.Number = 9 Then Resume Next
    ' Other errors...
    MsgBox "Error: " & Err.Description & " occurred.", _
           vbCritical, _
           "Error Downloading File"
    Err.Clear
    GoTo ExitDownload
    
FileErrorHandler:
    MsgBox "Cannot write file to disk." & _
           vbCr & vbCr & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, _
           "Error Downloading File"
    CancelSearch = True
    Err.Clear
    GoTo ExitDownload
    
End Function

Public Sub ParseFile(sFile As String, sWorld As String)
'THIS IS PART I WANTED TO DO THE MOST.  PARSE THE ASP FILE AFTER DOWNLOADING
'AFTER PARSING THE FILE IT THEN INSERTS THE INFO INTO THE LIST VIEW FOR VIEWING
'IF THE HTML CODE CHANGES IN THE ASP FILE THIS WILL MOST LIKELY NEED TO BE MODIFIED

Dim sFileText As String
Dim i As Integer
Dim i2 As Integer
Dim sCodeID As String
Dim sCodeTitle As String
Dim sCodeBy As String

sFileText = FileText(sFile)
'MsgBox sFileText


For i = 1 To Len(sFileText)
    If Mid(sFileText, i, 10) = "txtCodeId=" Then
        i = i + 10
        i2 = i
        sCodeID = ""
        For i2 = i To i + 20
            If IsNumeric(Mid(sFileText, i2, 1)) = True Then
                sCodeID = sCodeID & Mid(sFileText, i2, 1)
            Else
                Exit For
            End If
        Next
        
        i = i2
        
        For i2 = i To i + 30
            If Mid(sFileText, i2, 1) = ">" Then
                Exit For
            End If
        Next
        
        i = i2 + 1
        sCodeTitle = ""
        
        For i2 = i To i + 100 'GIVE ENOUGH SPACES FOR LONG TITLES
            If Mid(sFileText, i2, 1) <> "<" Then
                sCodeTitle = sCodeTitle & Mid(sFileText, i2, 1)
            Else
                Exit For
            End If
        Next
        
        i = i2
        
        For i2 = i To i + 50
            If Mid(sFileText, i2, 2) = "By" Then
                Exit For
            End If
        Next
        
        i = i2 + 3
        sCodeBy = ""
        
        For i2 = i To i + 50
            If Mid(sFileText, i2, 6) <> "&nbsp;" Then
                sCodeBy = sCodeBy & Mid(sFileText, i2, 1)
            Else
                Exit For
            End If
        Next
        
        i = i2 + 14
        sCodeDate = ""
        
        For i2 = i To i + 5
            If Mid(sFileText, i2, 1) <> "<" Then
                sCodeDate = sCodeDate & Mid(sFileText, i2, 1)
            Else
                Exit For
            End If
        Next
        
        lsvwNew.ListItems.Add , , lsvwNew.ListItems.Count + 1
        lsvwNew.ListItems(lsvwNew.ListItems.Count).ListSubItems.Add 1, , sCodeID
        lsvwNew.ListItems(lsvwNew.ListItems.Count).ListSubItems.Add 2, , sWorld
        lsvwNew.ListItems(lsvwNew.ListItems.Count).ListSubItems.Add 3, , sCodeTitle
        lsvwNew.ListItems(lsvwNew.ListItems.Count).ListSubItems.Add 4, , sCodeBy
        lsvwNew.ListItems(lsvwNew.ListItems.Count).ListSubItems.Add 5, , sCodeDate
        
        lsvwNew.Refresh
        
    End If
Next
End Sub


Function FileText(ByVal filename As String) As String
'READS A FILE TO THE END

Dim handle As Integer
handle = FreeFile
Open filename$ For Binary As #handle
FileText = Space$(LOF(handle))
Get #handle, , FileText
Close #handle

End Function


Public Sub ChangeWorldViews()
'IF THE COMBO BOX IS CLICKED OR CHANGED THIS IS FIRED
'ALSO THIS IS FIRED FOR AUTO REFRESHES

If lblStandBy.Visible = False Then frmWait.Visible = True

lsvwNew.ListItems.Clear

If cboWorlds.ListIndex = 0 Then
    Dim aResult() As String, i As Long
    
    aResult = Split(sCodeWorldURLS, "|")
    
    For i = LBound(aResult) To UBound(aResult)
        iTimerMinimize = 0
        ScanFile FormatNumber(aResult(i), 0)
    Next
Else
    iTimerMinimize = 0
    ScanFile FormatNumber(GetWorldText2ID(cboWorlds.Text), 0)
   
End If

If lblStandBy.Visible = False Then Unload frmWait

End Sub
