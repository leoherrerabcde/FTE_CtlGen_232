Attribute VB_Name = "ModListViewProperties"
' LHE 29 Sept 2006
' Archivo .bas Creado por LHE
' Rutinas de Manejo de ListView

Option Explicit



Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub InvalidateRect Lib "user32" (ByVal hWnd As Long, ByVal t As Long, ByVal bErase As Long)
Private Declare Sub ValidateRect Lib "user32" (ByVal hWnd As Long, ByVal t As Long)


Private Type POINT
   X As Long
   Y As Long
End Type

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Const LVM_FIRST As Long = &H1000
Private Const LVM_GETITEM As Long = LVM_FIRST + 5
Private Const LVM_FINDITEM As Long = LVM_FIRST + 13
Private Const LVM_ENSUREVISIBLE = LVM_FIRST + 19
Private Const LVM_SETCOLUMNWIDTH As Long = LVM_FIRST + 30
Private Const LVM_GETTOPINDEX = LVM_FIRST + 39
Private Const LVM_SETITEMSTATE As Long = LVM_FIRST + 43
Private Const LVM_GETITEMSTATE As Long = LVM_FIRST + 44
Private Const LVM_GETITEMTEXT As Long = LVM_FIRST + 45
Private Const LVM_SORTITEMS As Long = LVM_FIRST + 48
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE As Long = LVM_FIRST + 54
Private Const LVM_GETEXTENDEDLISTVIEWSTYLE As Long = LVM_FIRST + 55
Private Const LVM_SETCOLUMNORDERARRAY = LVM_FIRST + 58
Private Const LVM_GETCOLUMNORDERARRAY = LVM_FIRST + 59

Private Const LVS_EX_GRIDLINES As Long = &H1
Private Const LVS_EX_SUBITEMIMAGES As Long = &H2
Private Const LVS_EX_CHECKBOXES As Long = &H4
Private Const LVS_EX_TRACKSELECT As Long = &H8
Private Const LVS_EX_HEADERDRAGDROP As Long = &H10
Private Const LVS_EX_FULLROWSELECT As Long = &H20

Private Const LVFI_PARAM As Long = 1

Private Const LVIF_TEXT As Long = 1
Private Const LVIF_IMAGE As Long = 2
Private Const LVIF_PARAM As Long = 4
Private Const LVIF_STATE As Long = 8
Private Const LVIF_INDENT As Long = &H10
Private Const LVIF_NORECOMPUTE As Long = &H800
Private Const LVIS_STATEIMAGEMASK As Long = &HF000&

Private Type LV_ITEM
   Mask As Long
   Index As Long
   SubItem As Long
   State As Long
   StateMask As Long
   Text As String
   TextMax As Long
   Icon As Long
   Param As Long
   Indent As Long
End Type

Private Type LV_FINDINFO
   Flags As Long
   pSz As String
   lParam As Long
   pt As POINT
   vkDirection As Long
End Type
'--- ListView Set Column Width Messages ---'
Public Enum LVSCW_Styles
   LVSCW_AUTOSIZE = -1
   LVSCW_AUTOSIZE_USEHEADER = -2
End Enum

Public Enum LVStylesEx
   Checkboxes = LVS_EX_CHECKBOXES
   FullRowSelect = LVS_EX_FULLROWSELECT
   GridLines = LVS_EX_GRIDLINES
   HeaderDragDrop = LVS_EX_HEADERDRAGDROP
   SubItemImages = LVS_EX_SUBITEMIMAGES
   TrackSelect = LVS_EX_TRACKSELECT
End Enum

'--- Sorting Variables ---'
Public Enum LVItemTypes
   lvDate = 0
   lvNumber = 1
   lvBinary = 2
   lvAlphabetic = 3
End Enum
Public Enum LVSortTypes
   lvAscending = 0
   lvDescending = 1
End Enum

Enum ImageSizingTypes
   [sizeNone] = 0
   [sizeCheckBox]
   [sizeIcon]
End Enum


Public Function LVSetStyleEx(lv As ListView, ByVal NewStyle As LVStylesEx, ByVal NewVal As Boolean) As Boolean
   
   Dim nStyle As Long
   
   ' get the current ListView style
   nStyle = SendMessage(lv.hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, ByVal 0&)
   
   If NewVal Then
      ' set the extended style bit
      nStyle = nStyle Or NewStyle
   Else
      ' remove the extended style bit
      nStyle = nStyle Xor NewStyle
   End If
   
   ' set the new ListView style
   LVSetStyleEx = CBool(SendMessage(lv.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, ByVal nStyle))

End Function

Sub SaveColumnWideListView(ByRef lvLstView As ListView, _
                            ByRef lsApp As String, _
                            ByRef lsSection As String)

Dim ls_Key          As String
Dim i               As Integer

    With lvLstView
        For i = 1 To .ColumnHeaders.Count
            ls_Key = "Width Column " & Trim$(Str(i))
            SaveSetting lsApp, lsSection, ls_Key, .ColumnHeaders(i).Width
        Next
    End With
    

End Sub


Function LstVwFindItemChecked(LV_LstVw As ListView) As Integer

Dim i           As Integer
Dim LV_Checked  As Integer

    LV_Checked = 0
    With LV_LstVw
        For i = 1 To .ListItems.Count
            If .ListItems(i).Checked = True Then
                If LV_Checked Then
                    LstVwFindItemChecked = 0
                    Exit Function
                Else
                    LV_Checked = i
                End If
            End If
            
        Next
        LstVwFindItemChecked = LV_Checked
    End With
    
End Function

Sub LoadColumnWideListView(ByRef lvLstView As ListView, _
                            ByRef lsApp As String, _
                            ByRef lsSection As String)

Dim ls_Key          As String
Dim i               As Integer
Dim lWidth          As Integer
'Dim lsWidth         As String

    With lvLstView
        For i = 1 To .ColumnHeaders.Count
            ls_Key = "Width Column " & Trim$(Str(i))
            lWidth = ChangeRegionalConfig(GetSetting(lsApp, lsSection, ls_Key, .ColumnHeaders(i).Width))
            .ColumnHeaders(i).Width = lWidth
        Next
    End With
    

End Sub


Sub AddItemListView(ByRef lvLstVw As ListView, _
                    ByRef lsData() As String, _
                    Optional ByVal VerPrimero As Boolean, _
                    Optional ByVal Index As Integer)
                    
Dim LV_LstSubItems  As ListSubItems
Dim i               As Integer
    
    Set LV_LstSubItems = lvLstVw.ListItems.Add(, , lsData(0))
    
    lvLstVw.ListItems(lvLstVw.ListItems.Count).Tag = Index
    
    ValidateRect lvLstVw.hWnd, 0&
    
    If VerPrimero = False Then
        lvLstVw.ListItems(lvLstVw.ListItems.Count).EnsureVisible
    End If
    
    With LV_LstSubItems
        For i = 1 To UBound(lsData)
            .Add , , lsData(i)
        Next
    End With
    
End Sub

Sub AddColumListView(ByRef lvLstVw As ListView, _
                    ByRef lsData() As String)
                    
Dim LV_LstSubItems   As ListSubItems
Dim i               As Integer
    
    With lvLstVw
        .ColumnHeaders.Clear
        For i = 0 To UBound(lsData)
            .ColumnHeaders.Add , , lsData(i)
        Next
        .ListItems.Clear
    End With
    
End Sub

Function Contar_Items_Chequeados(ByRef LstVw As ListView) As Integer

Dim i           As Integer
Dim LV_Count    As Integer

    LV_Count = 0
    With LstVw
        For i = 1 To .ListItems.Count
            If LVItemChecked(LstVw, i) = True Then
                LV_Count = LV_Count + 1
            End If
        Next
    End With
    
    Contar_Items_Chequeados = LV_Count
    
End Function

Function FormatHora(lsHora As String) As String

    FormatHora = Format(lsHora, "hh:mm:ss")
    
End Function

Function GetSubItems(ByRef LstVw As ListView, Index As Integer) As String

    GetSubItems = LstVw.SelectedItem.ListSubItems(Index).Text
    
End Function


    

Function ChangeRegionalConfig(lsNumber As String) As String

        'If InStr(lsNumber, ".") Then
            ChangeRegionalConfig = Replace(lsNumber, ".", ",")
        'Else
        '    ChangeRegionalConfig = Replace(lsNumber, ",", ".")
        'End If

End Function

Function GetIntegerFromStr(lsNumber As String) As Double

    If IsNumeric(GetIntegerFromStr) = True Then
        If lsNumber < 32768 Then
            GetIntegerFromStr = lsNumber
        Else
            GetIntegerFromStr = ChangeRegionalConfig(lsNumber)
        End If
    Else
        GetIntegerFromStr = ChangeRegionalConfig(lsNumber)
    End If
    
End Function

Public Sub LVSetColWidth(lv As ListView, ByVal ColumnIndex As Long, ByVal Style As LVSCW_Styles)
   '------------------------------------------------------------------------------
   '--- If you include the header in the sizing then the last column will
   '--- automatically size to fill the remaining listview width.
   '------------------------------------------------------------------------------
   With lv
      ' verify that the listview is in report view and that the column exists
      If .View = lvwReport Then
         If ColumnIndex >= 1 And ColumnIndex <= .ColumnHeaders.Count Then
            Call SendMessage(.hWnd, LVM_SETCOLUMNWIDTH, ColumnIndex - 1, ByVal Style)
         End If
      End If
   End With
End Sub

Public Function LVItemChecked(lv As ListView, ByVal Index As Long) As Boolean
   Dim nRet As Long
   Const MaskBit As Long = &H1000   '(2 ^ 12)
   
   ' get current statemask bits
   nRet = SendMessage(lv.hWnd, LVM_GETITEMSTATE, Index - 1, ByVal LVIS_STATEIMAGEMASK)
   
   ' return what the Checked bit is set to
   LVItemChecked = (((nRet \ MaskBit) - 1) <> 0)
End Function

Public Sub LVSetAllColWidths(lv As ListView, ByVal Style As LVSCW_Styles)
   Dim ColumnIndex As Long
   '--- loop through all of the columns in the listview and size each
   With lv
      For ColumnIndex = 1 To .ColumnHeaders.Count
         LVSetColWidth lv, ColumnIndex, Style
      Next ColumnIndex
   End With
End Sub


Sub SetHighlightColumn(lv As ListView, _
                               clrHighlight As SystemColorConstants, _
                               clrDefault As OLE_COLOR, _
                               nColumn As Long, _
                               nSizingType As ImageSizingTypes, _
                               Picture1)

   Dim cnt     As Long  'counter
   Dim cl      As Long  'columnheader left
   Dim cw      As Long  'columnheader width
         
   On Local Error GoTo SetHighlightColumn_Error
   
   If lv.View = lvwReport Then
   
     'set up the listview properties
      With lv
        .Picture = Nothing  'clear picture
        .Refresh
        .Visible = 1
        .PictureAlignment = lvwTile
      End With  ' lv
        
     'set up the picture box properties
      With Picture1
         .AutoRedraw = False       'clear/reset picture
         .Picture = Nothing
         .BackColor = clrDefault
         .Height = 1
         .AutoRedraw = True        'assure image draws
         .BorderStyle = vbBSNone   'other attributes
         .ScaleMode = vbTwips
         '.Top = Form1.Top - 10000  'move it off screen
         .Visible = False
         .Height = 1               'only need a 1 pixel high picture
         .Width = Screen.Width
            
        'draw a box in the highlight colour
        'at location of the column passed
         cl = lv.ColumnHeaders(nColumn).Left
         cw = lv.ColumnHeaders(nColumn).Left + _
              lv.ColumnHeaders(nColumn).Width
         Picture1.Line (cl, 0)-(cw, 210), clrHighlight, BF
         
         .AutoSize = True
      End With  'Picture1
     
     'set the lv picture to the
     'Picture1 image
      lv.Refresh
      lv.Picture = Picture1.Image
      
   Else
    
      lv.Picture = Nothing
        
   End If  'lv.View = lvwReport

SetHighlightColumn_Exit:
On Local Error GoTo 0
Exit Sub
    
SetHighlightColumn_Error:

  'clear the listview's picture and exit
   With lv
      .Picture = Nothing
      .Refresh
   End With
   
   Resume SetHighlightColumn_Exit
    
End Sub


