Attribute VB_Name = "basMain"
Option Explicit

  Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal sz As Long)


' returns the higher 16 bits of <val>
Function HiWord(ByVal val As Long) As Integer
  Dim ret As Integer

  CopyMemory VarPtr(ret), VarPtr(val) + LenB(ret), LenB(ret)

  HiWord = ret
End Function

' returns the lower 16 bits of <val>
Function LoWord(ByVal val As Long) As Integer
  Dim ret As Integer

  CopyMemory VarPtr(ret), VarPtr(val), LenB(ret)

  LoWord = ret
End Function

' merges <Lo> and <Hi> to a 32 bit number
Function MakeDWord(ByVal Lo As Integer, ByVal Hi As Integer) As Long
  Dim ret As Long

  CopyMemory VarPtr(ret), VarPtr(Lo), LenB(Lo)
  CopyMemory VarPtr(ret) + LenB(Lo), VarPtr(Hi), LenB(Hi)

  MakeDWord = ret
End Function
