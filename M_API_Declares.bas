Attribute VB_Name = "M_API_Declares"
  Option Explicit
  ' demo project showing how to impliment advanced subclassing techniques in VB
  ' by Bryan Stafford of New Vision Software - newvision@mvps.org
  ' this demo is released into the public domain "as is" without
  ' warranty or guaranty of any kind.  In other words, use at
  ' your own risk.
  
  ' IMPORTANT NOTE:  if you don't have the DbgwProc.dll Debug Object from
  ' the MS site, you should get it and use it with this project while in
  ' the IDE.  before compiling an app that uses the debug object, you should
  ' set the DEBUGWINDOWPROC conditional compiliation variable equal to zero.


  Public Const WM_EXITSIZEMOVE As Long = &H232&

  Public Const GWL_WNDPROC As Long = (-4&)

  Public Declare Function IsWindow Lib "user32" (ByVal hWnd&) As Long

  Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes&)

  Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
  Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
  Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd&, ByVal lpString$) As Long

  Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd&, _
                                                              ByVal nIndex&, ByVal dwNewLong&) As Long

  Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc&, _
                                               ByVal hWnd&, ByVal Msg&, ByVal wParam&, ByVal lParam&) As Long

  ' the function we use to send our custom message to the window
  Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd&, ByVal wMsg&, ByVal wParam&, lParam As Any) As Long


  Public Declare Function GetTickCount& Lib "kernel32" ()
  
