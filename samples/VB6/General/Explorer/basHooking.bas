Attribute VB_Name = "basHooking"
Option Explicit

  Private hKeyboardHook As Long
  Private pKeyboardHookConsumers() As Long


  Private Declare Function CallNextHookEx Lib "user32.dll" (ByVal hhk As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal sz As Long)
  Private Declare Function GetCurrentThreadId Lib "kernel32.dll" () As Long
  Private Declare Function SetWindowsHookEx Lib "user32.dll" Alias "SetWindowsHookExW" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadId As Long) As Long
  Private Declare Function UnhookWindowsHookEx Lib "user32.dll" (ByVal hhk As Long) As Long
  Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByVal pDest As Long, ByVal sz As Long)


' registers <Obj> for a keyboard hook
Sub InstallKeyboardHook(ByVal Obj As IHook)
  Const WH_KEYBOARD = 2
  Dim exists As Boolean
  Dim i As Integer
  Dim iAvailableSlot As Integer
  Dim pObj As Long

  ' if we didn't yet, register the hook now
  If hKeyboardHook = 0 Then hKeyboardHook = SetWindowsHookEx(WH_KEYBOARD, AddressOf KeyboardProc, 0&, GetCurrentThreadId)
  If hKeyboardHook Then
    ' now search the array of listeners for a free slot
    iAvailableSlot = -1
    pObj = ObjPtr(Obj)
    On Error GoTo initArray
    i = LBound(pKeyboardHookConsumers)
    For i = LBound(pKeyboardHookConsumers) To UBound(pKeyboardHookConsumers)
      If pKeyboardHookConsumers(i) = pObj Then
        ' the object that wanted this hook, already has it
        exists = True
      ElseIf pKeyboardHookConsumers(i) = 0 Then
        ' we found a free slot
        iAvailableSlot = i
        Exit For
      End If
    Next

    If Not exists Then
      If iAvailableSlot = 0 Then
        ' no free slot was found, so resize the array
        ReDim Preserve pKeyboardHookConsumers(UBound(pKeyboardHookConsumers) + 1) As Long
        iAvailableSlot = UBound(pKeyboardHookConsumers)
      End If
      ' store the new listener
      pKeyboardHookConsumers(iAvailableSlot) = pObj
    End If
  End If
  Exit Sub

initArray:
  ReDim pKeyboardHookConsumers(0)
  Resume Next
End Sub

' the callback method for a keyboard hook
Function KeyboardProc(ByVal hookCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim eatIt As Boolean
  Dim i As Integer
  Dim Obj As IHook

  On Error GoTo MyError
  If hookCode >= 0 Then
    ' call each listener before we forward this message to the next hook
    For i = LBound(pKeyboardHookConsumers) To UBound(pKeyboardHookConsumers)
      If pKeyboardHookConsumers(i) Then
        Set Obj = GetObjectFromPtr(pKeyboardHookConsumers(i))
        Obj.KeyboardProcBefore hookCode, wParam, lParam, eatIt
      End If
    Next
    If eatIt Then
      ' bon appetit :)
      KeyboardProc = 1
    Else
      ' let the next hook do his job
      KeyboardProc = CallNextHookEx(hKeyboardHook, hookCode, wParam, lParam)
      ' call each listener again
      For i = LBound(pKeyboardHookConsumers) To UBound(pKeyboardHookConsumers)
        If pKeyboardHookConsumers(i) Then
          Set Obj = GetObjectFromPtr(pKeyboardHookConsumers(i))
          Obj.KeyboardProcAfter hookCode, wParam, lParam
        End If
      Next
    End If
  Else
    KeyboardProc = CallNextHookEx(hKeyboardHook, hookCode, wParam, lParam)
  End If
  Exit Function

MyError:
  KeyboardProc = CallNextHookEx(hKeyboardHook, hookCode, wParam, lParam)
End Function

' removes <Obj> from the listeners of the keyboard hook
Sub RemoveKeyboardHook(ByVal Obj As IHook)
  Dim countRefs As Integer
  Dim i As Integer
  Dim pObj As Long

  ' find the object in the array of listeners
  pObj = ObjPtr(Obj)
  On Error GoTo MyError
  For i = LBound(pKeyboardHookConsumers) To UBound(pKeyboardHookConsumers)
    If pKeyboardHookConsumers(i) = pObj Then
      ' found it, free its slot
      pKeyboardHookConsumers(i) = 0
    ElseIf pKeyboardHookConsumers(i) Then
      ' there's still another object listening
      countRefs = countRefs + 1
    End If
  Next

  If countRefs = 0 Then
    ' all listeners are gone, so unregister the hook
    UnhookWindowsHookEx hKeyboardHook
    hKeyboardHook = 0
  End If

MyError:
End Sub


' returns the object <Ptr> points to
Private Function GetObjectFromPtr(ByVal Ptr As Long) As Object
  Dim ret As Object

  ' get the object <Ptr> points to
  CopyMemory VarPtr(ret), VarPtr(Ptr), LenB(Ptr)
  Set GetObjectFromPtr = ret
  ' free memory
  ZeroMemory VarPtr(ret), 4
End Function
