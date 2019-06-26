Attribute VB_Name = "MAdvancedSubclass"
  Option Private Module
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


' our windowproc and the function that resolves the unreferenced pointer that
' we stored on the window using SetProp.  this pair of functions gets the pointer
' using GetProp and turns it into a referenced object that can be called using the
' standard VB syntax.  the WindProc function then passes the messages and params to
' the WindowProc function in our cSubclass module.  since the object pointer is
' stored directly on the subclassed window, we are guaranteed to be accessing the
' correct instance of the cSubclass object associated with the window.  this way,
' you can have more than one form using the subclass object and still have all of
' the messages go to the correct place.
Public Function WindProc(ByVal hWnd&, ByVal uMsg&, ByVal wParam&, ByVal lParam&) As Long
  WindProc = cSubclassFromhWnd(hWnd).WindowProc(hWnd, uMsg, wParam, lParam)
End Function

Private Function cSubclassFromhWnd(ByVal hWnd As Long) As cSubclass
  ' resolve a dumb pointer into a referenced object....
  
  Dim SubclassEx As cSubclass, pObj As Long
    
  ' retrieve the pointer from the property we set in the subclass routine
  pObj = GetProp(hWnd, ByVal "nvAdvSubcls")

  ' copy the pointer into the local variable.  if you end your app durring this
  ' process, vb will crash when it tries to destroy the extra object reference
  ' so don't end your app now.
  CopyMemory SubclassEx, pObj, 4&
  
  ' set a reference to the object
  Set cSubclassFromhWnd = SubclassEx
  
  ' clear the object variable so VB won't try to
  ' decrement the reference count on the object
  CopyMemory SubclassEx, 0&, 4&
  
End Function
