Attribute VB_Name = "Globals"
'SOME FUNCTIONS HERE WERE TAKEN FROM THE PSC CODE SAMPLE MENTIONED EARLIER


Global sCodeWorldURLS As String
Private CancelSearch As Boolean


'****************************************************************
'Windows API/Global Declarations for :FreeDiskSpace
'****************************************************************
Private Declare Function GetDiskFreeSpace Lib "kernel32" _
                         Alias "GetDiskFreeSpaceA" _
                         (ByVal lpRootPathName As String, _
                          lpSectorsPerCluster As Long, _
                          lpBytesPerSector As Long, _
                          lpNumberOfFreeClusters As Long, _
                          lpTotalNumberOfClusters As Long) As Long


Function GetWorldIDText(iID As Integer) As String
'RETURNS THE STRING FOR A WORLD ID

Select Case iID

    Case 1:
        GetWorldIDText = "Visual Basic"
    Case 2:
        GetWorldIDText = "Java/Javascript"
    Case 3:
        GetWorldIDText = "C/C++"
    Case 4:
        GetWorldIDText = "ASP (Active Server Pages)"
    Case 5:
        GetWorldIDText = "SQL"
    Case 6:
        GetWorldIDText = "PERL"
    Case 7:
        GetWorldIDText = "DELPHI"
    Case 8:
        GetWorldIDText = "PHP"
    Case 9:
        GetWorldIDText = "Cold Fusion"
    Case 10:
        GetWorldIDText = ".Net"
    Case 13:
        GetWorldIDText = "LISP"
    Case 14:
        GetWorldIDText = "Dark Basic"
    Case Else:
        GetWorldIDText = "Unknown..."
End Select

End Function

Function GetWorldText2ID(sText As String) As Integer
'FINDS THE ID FOR A WORLD STRING

Select Case sText

    Case "Visual Basic":
        GetWorldText2ID = 1
    Case "Java/Javascript":
        GetWorldText2ID = 2
    Case "C/C++":
        GetWorldText2ID = 3
    Case "ASP (Active Server Pages)":
        GetWorldText2ID = 4
    Case "SQL":
        GetWorldText2ID = 5
    Case "PERL":
        GetWorldText2ID = 6
    Case "DELPHI":
        GetWorldText2ID = 7
    Case "PHP":
        GetWorldText2ID = 8
    Case "Cold Fusion":
        GetWorldText2ID = 9
    Case ".Net":
        GetWorldText2ID = 10
    Case "LISP":
        GetWorldText2ID = 13
    Case "Dark Basic":
        GetWorldText2ID = 14
    Case Else:
        GetWorldText2ID = 1
End Select

End Function


Public Function FormatFileSize(ByVal dblFileSize As Double, _
                               Optional ByVal strFormatMask As String) _
                               As String

' FormatFileSize:   Formats dblFileSize in bytes into
'                   X GB or X MB or X KB or X bytes depending
'                   on size (a la Win9x Properties tab)

Select Case dblFileSize
    Case 0 To 1023              ' Bytes
        FormatFileSize = Format(dblFileSize) & " bytes"
    Case 1024 To 1048575        ' KB
        If strFormatMask = Empty Then strFormatMask = "###0"
        FormatFileSize = Format(dblFileSize / 1024#, strFormatMask) & " KB"
    Case 1024# ^ 2 To 1073741823 ' MB
        If strFormatMask = Empty Then strFormatMask = "###0.0"
        FormatFileSize = Format(dblFileSize / (1024# ^ 2), strFormatMask) & " MB"
    Case Is > 1073741823#       ' GB
        If strFormatMask = Empty Then strFormatMask = "###0.0"
        FormatFileSize = Format(dblFileSize / (1024# ^ 3), strFormatMask) & " GB"
End Select

End Function

Public Function FormatTime(ByVal sglTime As Single) As String
                           
' FormatTime:   Formats time in seconds to time in
'               Hours and/or Minutes and/or Seconds

' Determine how to display the time
Select Case sglTime
    Case 0 To 59    ' Seconds
        FormatTime = Format(sglTime, "0") & " sec"
    Case 60 To 3599 ' Minutes Seconds
        FormatTime = Format(Int(sglTime / 60), "#0") & _
                     " min " & _
                     Format(sglTime Mod 60, "0") & " sec"
    Case Else       ' Hours Minutes
        FormatTime = Format(Int(sglTime / 3600), "#0") & _
                     " hr " & _
                     Format(sglTime / 60 Mod 60, "0") & " min"
End Select

End Function

Public Function DiskFreeSpace(strDrive As String) As Double

' DiskFreeSpace:    returns the amount of free space on a drive
'                   in Windows9x/2000/NT4+

Dim SectorsPerCluster As Long
Dim BytesPerSector As Long
Dim NumberOfFreeClusters As Long
Dim TotalNumberOfClusters As Long
Dim FreeBytes As Long
Dim spaceInt As Integer

strDrive = QualifyPath(strDrive)

' Call the API function
GetDiskFreeSpace strDrive, _
                 SectorsPerCluster, _
                 BytesPerSector, _
                 NumberOFreeClusters, _
                 TotalNumberOfClusters

' Calculate the number of free bytes
DiskFreeSpace = NumberOFreeClusters * SectorsPerCluster * BytesPerSector

End Function


Public Function QualifyPath(strPath As String) As String

' Make sure the path ends in "\"
QualifyPath = IIf(Right(strPath, 1) = "\", strPath, strPath & "\")

End Function


Public Function ReturnFileOrFolder(FullPath As String, _
                                   ReturnFile As Boolean, _
                                   Optional IsURL As Boolean = False) _
                                   As String

' ReturnFileOrFolder:   Returns the filename or path of an
'                       MS-DOS file or URL.
'
' Author:   Jeff Cockayne 4.30.99
'
' Inputs:   FullPath:   String; the full path
'           ReturnFile: Boolean; return filename or path?
'                       (True=filename, False=path)
'           IsURL:      Boolean; Pass True if path is a URL.
'
' Returns:  String:     the filename or path
'

Dim intDelimiterIndex As Integer

intDelimiterIndex = InStrRev(FullPath, IIf(IsURL, "/", "\"))
If intDelimiterIndex = 0 Then
    ReturnFileOrFolder = FullPath
Else
    ReturnFileOrFolder = IIf(ReturnFile, _
                         Right(FullPath, Len(FullPath) - intDelimiterIndex), _
                         Left(FullPath, intDelimiterIndex))
End If

End Function
