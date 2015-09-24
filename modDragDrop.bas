Attribute VB_Name = "modDragDrop"
Option Compare Database
Option Explicit

'modDragDrop
'2011-09 jl

'Usage instructions/documentation, files and samples as well as
'the related FeedMe functionality can be found here:
'   http://www.utteraccess.com/forum/index.php?showtopic=1973842


'V1.0.1 2011-09-29
'  Modifications from original:
'    - Added link resolution
'    - Added DragDropRemoveDuplicates()
'  Known Bugs
'    - Link resolution doesn't seem to work on
'      some .lnk files.
'      You can modify pfGetShortcutTargets to
'      turn off shortcut target resolution and
'      always return .lnk or .url files.


Private Declare Function SetWindowLong _
  Lib "User32" Alias "SetWindowLongA" _
    (ByVal hWnd As Long, _
     ByVal nIndex As Long, _
     ByVal dwNewLong As Long _
     ) As Long

Private Declare Function GetWindowLong _
  Lib "User32" Alias "GetWindowLongA" _
   (ByVal hWnd As Long, _
    ByVal nIndex As Long _
    ) As Long

Private Declare Sub DragAcceptFiles _
  Lib "shell32.dll" _
    (ByVal hWnd As Long, _
     ByVal fAccept As Long)

Private Declare Sub DragFinish _
  Lib "shell32.dll" _
    (ByVal hDrop As Long)

Private Declare Function DragQueryFile _
  Lib "shell32.dll" Alias "DragQueryFileA" _
    (ByVal hDrop As Long, _
     ByVal iFile As Long, _
     ByVal lpszFile As String, _
     ByVal cch As Long _
     ) As Long

Private Declare Function CallWindowProc _
  Lib "User32" Alias "CallWindowProcA" _
    (ByVal lpPrevWndFunc As Long, _
     ByVal hWnd As Long, _
     ByVal Msg As Long, _
     ByVal wParam As Long, _
     ByVal lParam As Long _
     ) As Long

Const GWL_WNDPROC As Long = -4
Const GWL_EXSTYLE = -20

Const WM_DROPFILES = &H233

Const WS_EX_ACCEPTFILES = &H10

Private currFrm As Access.Form 'reference to the Form currently
                               'hooked to the callback.  This
                               'is required to pass the dropped
                               'files to the form
Private currHwndFrm As Long 'handle to the currently hooked form
                            'this is used to verify there isn't
                            'two forms hooked at the same time
                            '(for example overlapping controls
                            'where the mouse jumps from one
                            'to the other)
Private prevWndProc As Long 'handle to the window's previous
                            'WindowProc.  This is only set while
                            'currHwndFrm is valid

Private hookOn As Boolean


Public Enum DragDropReturnFormat
  ReturnSemicolonDelimitedString = 0
  ReturnVariantArray = 1
  ReturnLineFormattedString = 2
End Enum


Public Function DragDropRemoveDuplicates(ByVal FileList As String) As String
  'accepts a semicolon delimited list string and removes
  'any duplicate instances of the substrings
  
  Dim v As Variant
  Dim i As Integer
  Dim s As String
  
  v = Split(FileList, ";")
  FileList = ";" & FileList & ";"
  For i = 0 To UBound(v)
    If InStr(1, s, ";" & v(i) & ";") = 0 Then
      s = s & ";" & v(i) & ";"
    End If
  Next i
  
  While InStr(1, s, ";;") <> 0
    s = Replace(s, ";;", ";")
  Wend
  
  DragDropRemoveDuplicates = Mid(s, 2, Len(s) - 2)
  
End Function

Public Function DragDropFormatAs( _
    retFormat As DragDropReturnFormat, _
    FileList As String _
    ) As Variant
  'returns the default semicolon delimited list in
  'the specified format
  Dim Ret As Variant
  
  Select Case retFormat
  
    Case DragDropReturnFormat.ReturnSemicolonDelimitedString
      Ret = CStr(FileList)
  
    Case DragDropReturnFormat.ReturnVariantArray
      Ret = Split(FileList, ";")
    
    Case DragDropReturnFormat.ReturnLineFormattedString
      Ret = CStr(Replace(FileList, ";", vbCrLf))
      
  End Select
  
  DragDropFormatAs = Ret

End Function


Public Function DragDropInitForm(hWndFrm As Long)
  'inits the form for drag/drop
  Dim lExStyle As Long
  
  'get the current extended window style
  lExStyle = GetWindowLong(hWndFrm, GWL_EXSTYLE)
  'add the flag for accepting dragged files
  lExStyle = lExStyle Or WS_EX_ACCEPTFILES
  'set the new extended window style
  SetWindowLong hWndFrm, GWL_EXSTYLE, lExStyle
  
  'register the form for drag/drop acceptance
  DragAcceptFiles hWndFrm, True
  
End Function


Public Function DragDropSetHook(hookOn As Boolean, hWndFrm As Long)
  'toggles the hook on or off
  If hookOn Then
    dragDropHookOn hWndFrm
    'Debug.Print "HookOn"
  Else
    dragDropHookOff
    'Debug.Print "HookOff"
  End If
End Function

Public Function DragDropCallback( _
  ByVal hWnd As Long, _
  ByVal Msg As Long, _
  ByVal wParam As Long, _
  ByVal lParam As Long _
  ) As Long
  'callback for drag/drops
  
  'as in all callbacks, we're screwed if there's an error
  'and need to be absolutely certain this procedure will
  'exit
  On Error Resume Next

  'We want to turn this hook off as soon as possible after
  'the files are dropped, otherwise it clogs up the rest
  'of the project operations.  It just so happens that the
  'MouseMove event for a control accepting the dropped files
  'fires exactly one time when the files are physically
  'dropped on the form - that one time is where this hook
  'is turned on, so it's almost a safe bet that we can
  'turn this directly back off as soon as it's called.
  '
  'If you experience occasional times where you drop files
  'but they do not register, set the DDH_MAXCALLS constant
  'to a higher number.  In Access 2010 in particular this
  'number seems to want to be set around 50 or so for
  'it to work every time.  You might need to go higher

  
  Const DDH_MAXCALLS = 50
  Static intCallCount As Integer
  
  If Msg <> WM_DROPFILES Then
    'the hWnd parameter passed to this function by the OS is
    'the handle to the form that was hooked.  If the message
    'is not for a file drop, we'll send the message to that
    'form's standard message procedure
    CallWindowProc ByVal prevWndProc, _
                   ByVal hWnd, _
                   ByVal Msg, _
                   ByVal wParam, _
                   ByVal lParam
  Else
    'we have a file drop, handle that
    dragDropQueryFiles wParam
  End If
  
  'check the callcount and unhook if required
  intCallCount = intCallCount + 1
  If intCallCount >= DDH_MAXCALLS Then
    intCallCount = 0
    DragDropSetHook False, hWnd
  End If
  
End Function


Private Sub dragDropQueryFiles(hDrop As Long)
  'passes a semicolon delimited list of the dropped
  'files to the currFrm.DragDropFiles() sub
  
  Const MAX_PATH = 255
  Dim Ret As String 'function return
  Dim s As String 'temp/various
  Dim iCount As Integer 'count of files dropped
  Dim iPathLen As Integer 'length of the current path
  Dim i As Integer 'temp/various
  
  'get the count of files dropped
  s = String(MAX_PATH, 0)
  iCount = DragQueryFile(hDrop, &HFFFFFFFF, s, Len(s))
  
  'iterate the filecount and build the return
  For i = 0 To iCount - 1
    s = String(MAX_PATH, 0)
    iPathLen = DragQueryFile(hDrop, i, s, MAX_PATH)
    Ret = Ret & ";" & Trim(Left(s, iPathLen))
  Next i
    
  DragFinish hDrop
  
  
  Ret = Mid(Ret, 2)
  
  Ret = pfGetShortcutTargets(Ret)
  
  currFrm.DragDropFiles Ret
    
End Sub

Private Function pfGetShortcutTargets(ByVal files As String) As String
  Dim v As Variant
  Dim i As Integer
  Dim Ret As String
  
  v = Split(files, ";")
  
  For i = 0 To UBound(v)
  
    If (Right(CStr(v(i)), 4) = ".lnk") Or (Right(CStr(v(i)), 4) = ".url") Then
      Ret = Ret & ";" & pfGetShortcutTarget(CStr(v(i)))
    Else
      Ret = Ret & ";" & CStr(v(i))
    End If
  
  Next i
  
  pfGetShortcutTargets = Mid(Ret, 2)
  
End Function

Private Function pfGetShortcutTarget(sLink As String, Optional bIncludeArguments As Boolean = True) As String
  On Error GoTo err_proc
  Dim Ret As String
  With CreateObject("Shell.Application").NameSpace(0).ParseName(sLink).GetLink
    'Debug.Print .Path
    Ret = .Path
    If Ret = "" Then
      'PIDL (Computer, Recycle Bin, Games, etc)
      Ret = sLink
    Else
      If bIncludeArguments Then Ret = Trim(Ret & " " & .Arguments)
    End If
  End With
exit_proc:
  pfGetShortcutTarget = Ret
  Exit Function
err_proc:
  If Err.Number = 445 Then
    'object doesn't support this action (.Arguments failure on .url link)
    Resume exit_proc
  Else
    'Debug.Print Err.Number & " " & Err.Description
    'known errors are 91: failure to create object and 70: permission denied
    'also: office applications return non-usable path
    'PIDL locations return "" so we apply the original link instead
    Ret = sLink
  End If
  Resume exit_proc
End Function



Private Function dragDropHookOn(hWndFrm As Long)
  'turns the callback on
  If hookOn Then Exit Function
  'set the current form and handle for later use
  currHwndFrm = hWndFrm
  Set currFrm = dragDropGetFrmFromHWnd(hWndFrm)
  'use SetWindowLong to set the new callback address
  'and return the previous callback address to prevWndProc
  prevWndProc = SetWindowLong(hWndFrm, GWL_WNDPROC, AddressOf DragDropCallback)
  hookOn = True
End Function

Private Function dragDropHookOff()
  'turns the callback off
  
  'set the window's callback address to it's previous value
  SetWindowLong currHwndFrm, GWL_WNDPROC, prevWndProc
  'clear the window and callback settings until next callback init
  prevWndProc = 0
  currHwndFrm = 0
  Set currFrm = Nothing
  hookOn = False
End Function


Private Function dragDropGetFrmFromHWnd(hWnd As Long) As Access.Form
  'retrieves a form reference from the form's handle
  'this reference is later used to call the form's procedure and pass
  'it the dropped file list
  Dim frm As Access.Form
  For Each frm In Access.Forms
    If frm.hWnd = hWnd Then Exit For
  Next frm
  Set dragDropGetFrmFromHWnd = frm
End Function




