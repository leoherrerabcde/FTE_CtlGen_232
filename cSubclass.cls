VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
  
  
  ' if we are using the debug dll, we need a debug hook variable
  #If DEBUGWINDOWPROC Then
    Private m_DWPHook As WindowProcHook
  #End If
  ' original window proc address
  Private m_WinProcOld As Long
  
  ' used to demonstate sending application defined messages
  Public Enum CUSTOM_MESSAGES
    NV_CUSTOM_MESSAGE = &H8FF1&
    NV_INSTRUMENT_MESSAGE = &H8FF2&
    NV_PROJECT_CHANGE_MSG = &H8FF3&
  End Enum
  
  ' subclassed hWnd
  Private m_hWnd As Long
    
  ' our events
  Event FormMove()
  Event CustomMessage()
  Event InstrumentMessage(wParam As Long, lParam As Long)
  Event ProjectChange(wParam As Long, lParam As Long)
  
Friend Function WindowProc(ByVal hWindow&, ByVal uMsg&, ByVal wParam&, ByVal lParam&) As Long
  ' this function is called from the MAdvancedSubclass BAS module.  all messages are for
  ' the subclasses hWnd are passed here to be processed before passing them on to VB

  Select Case uMsg
    ' message sent when the window is finished being moved
    Case WM_EXITSIZEMOVE
      RaiseEvent FormMove
              
    ' our custom message
    Case NV_CUSTOM_MESSAGE
      RaiseEvent CustomMessage

      ' this is a private message so we don't need to pass it on to VB
      Exit Function
      
    Case NV_INSTRUMENT_MESSAGE
        RaiseEvent InstrumentMessage(wParam&, lParam&)
        
        Exit Function
    
    Case NV_PROJECT_CHANGE_MSG
        RaiseEvent ProjectChange(wParam&, lParam&)
        Exit Function
  End Select
  
  ' pass the messages on to VB
  WindowProc = CallWindowProc(m_WinProcOld, hWindow, uMsg, wParam, lParam)

End Function

Public Sub SubClass(ByVal hWnd&)
  ' this routine subclasses the window
  
  ' first check with the OS to make sure the hWnd is a valid window
  If IsWindow(hWnd) Then
    ' we store a pointer to this instance of the class in an extra property
    ' we add to the window using the SetProp API function.  if the property exists,
    ' this window is already subclassed and we don't want to overwrite the property
    If GetProp(hWnd, "nvAdvSubcls") Then Exit Sub
    
    ' store a pointer to this instance of the class object in an added property
    If SetProp(hWnd, ByVal "nvAdvSubcls", ObjPtr(Me)) Then
      ' if the SetProp call succeeded, subclass the window...
      
      ' we use a conditional compilation argument to tell whether or not we should use the
      ' debug dll.  when you want to compile a release, set the DEBUGWINDOWPROC variable in
      ' the project properties dialog to zero.
      #If DEBUGWINDOWPROC Then
        ' this is the procedure for setting a debug hook...
        Set m_DWPHook = CreateWindowProcHook
        With m_DWPHook
        
          ' our windowproc is in the MAdvancedSubclass BAS module
          .SetMainProc AddressOf MAdvancedSubclass.WindProc
          
          ' replace the windowproc for this window with ours
          m_WinProcOld = SetWindowLong(hWnd, GWL_WNDPROC, .ProcAddress)
          .SetDebugProc m_WinProcOld
        End With
      #Else
        ' not using the debug dll so just replace the windowproc for this window with ours.
        ' our windowproc is in the MAdvancedSubclass BAS module
        m_WinProcOld = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf MAdvancedSubclass.WindProc)
      #End If
  
      ' store the subclassed hWnd so we will have it when it is time to un-subclass
      m_hWnd = hWnd
    End If
  End If
  
End Sub

Private Sub UnSubClass()
  ' this routine un-subclasses the window
  
  ' make sure the window still exists....
  If IsWindow(m_hWnd) Then
    ' if we have an address for the old windowproc, return the proc to the original state...
    If m_WinProcOld Then
      ' un-subclass the window
      SetWindowLong m_hWnd, GWL_WNDPROC, m_WinProcOld
      
      ' remove the added property
      RemoveProp m_hWnd, "nvAdvSubcls"
      
      ' set the variables to zero to avoid any misshaps
      m_WinProcOld = 0
      m_hWnd = 0
    End If
    
    ' if we are using the debug dll, destroy the instance
    #If DEBUGWINDOWPROC Then
      Set m_DWPHook = Nothing
    #End If
  End If
  
End Sub

Private Sub Class_Terminate()
  ' we un-subclass the window when we are terminated
  UnSubClass

End Sub
