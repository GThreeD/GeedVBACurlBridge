VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "WSC_ISender"
Option Explicit

Public Function Connect(ByVal wsUrl As String, ByVal timeoutMs As Long, ByVal verifyPeer As Boolean, ByVal verifyHost As Boolean) As Boolean
End Function

Public Function SendText(ByVal textValue As String) As Boolean
End Function

Public Function PollOnce(ByVal timeoutMs As Long) As Boolean
End Function

Public Sub Close()
End Sub

Public Property Get IsConnected() As Boolean
End Property

Public Property Get LastError() As String
End Property