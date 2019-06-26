Attribute VB_Name = "modSelFile"
Option Explicit



Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OpenFileName) As Long

Private Type OpenFileName
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

'Purpose     :  Allows the user to select a file name from a local or network directory.
'Inputs      :  sInitDir            The initial directory of the file dialog.
'               sFileFilters        A file filter string, with the following format:
'                                   eg. "Excel Files;*.xls|Text Files;*.txt|Word Files;*.doc"
'               [sTitle]            The dialog title
'               [lParentHwnd]       The handle to the parent dialog that is calling this function.
'Outputs     :  Returns the selected path and file name or a zero length string if the user pressed cancel


Function BrowseForFile(sInitDir As String, Optional ByVal sFileFilters As String, Optional sTitle As String = "Open File", Optional lParentHwnd As Long) As String
    Dim tFileBrowse As OpenFileName
    Const clMaxLen As Long = 254
    
    tFileBrowse.lStructSize = Len(tFileBrowse)
    
    'Replace friendly deliminators with nulls
    sFileFilters = Replace(sFileFilters, "|", vbNullChar)
    sFileFilters = Replace(sFileFilters, ";", vbNullChar)
    If Right$(sFileFilters, 1) <> vbNullChar Then
        'Add final delimiter
        sFileFilters = sFileFilters & vbNullChar
    End If
    
    'Select a filter
    tFileBrowse.lpstrFilter = sFileFilters & "All Files (*.*)" & vbNullChar & "*.*" & vbNullChar
    'create a buffer for the file
    tFileBrowse.lpstrFile = String(clMaxLen, " ")
    'set the maximum length of a returned file
    tFileBrowse.nMaxFile = clMaxLen + 1
    'Create a buffer for the file title
    tFileBrowse.lpstrFileTitle = Space$(clMaxLen)
    'Set the maximum length of a returned file title
    tFileBrowse.nMaxFileTitle = clMaxLen + 1
    'Set the initial directory
    tFileBrowse.lpstrInitialDir = sInitDir
    'Set the parent handle
    tFileBrowse.hwndOwner = lParentHwnd
    'Set the title
    tFileBrowse.lpstrTitle = sTitle
    
    'No flags
    tFileBrowse.flags = 0

    'Show the dialog
    If GetOpenFileName(tFileBrowse) Then
        BrowseForFile = Trim$(tFileBrowse.lpstrFile)
        If Right$(BrowseForFile, 1) = vbNullChar Then
            'Remove trailing null
            BrowseForFile = Left$(BrowseForFile, Len(BrowseForFile) - 1)
        End If
    End If
End Function

'Sub Test()
'    BrowseForFile "c:\", "Excel File (*.xls);*.xls", "Open Workbook"
'End Sub

