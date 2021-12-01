Attribute VB_Name = "Module1"
'****************************************************************
'Windows API/Global Declarations for :FileFound()
'****************************************************************

Public Declare Function FindFirstFile& Lib "kernel32" _
       Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData _
       As WIN32_FIND_DATA)

Public Declare Function FindClose Lib "kernel32" _
       (ByVal hFindFile As Long) As Long


Public Const MAX_PATH = 260

Type FILETIME ' 8 Bytes
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type


Type WIN32_FIND_DATA ' 318 Bytes
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved¯ As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type
Declare Function mciSendString Lib "winmm.dll" Alias _
    "mciSendStringA" (ByVal lpstrCommand As String, ByVal _
    lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal _
    hwndCallback As Long) As Long
    
    
Global ValidFile As Boolean

Function FileFound(strFileName As String) As Boolean

Dim lpFindFileData As WIN32_FIND_DATA
Dim hFindFirst As Long

     
       hFindFirst = FindFirstFile(strFileName, lpFindFileData)

              If hFindFirst > 0 Then
                      FindClose hFindFirst
                      ValidFile = True
              Else
                      ValidFile = False
              End If

End Function

