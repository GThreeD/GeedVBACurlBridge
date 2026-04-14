VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "WSC_TestSink"
Option Explicit
Implements WSC_IEventSink

Private Sub WSC_IEventSink_OnMessage(ByVal textValue As String)
    Debug.Print "Message=" & textValue
End Sub

Private Sub WSC_IEventSink_OnError(ByVal errorText As String)
    Debug.Print "Error=" & errorText
End Sub

Private Sub WSC_IEventSink_OnOpen()
    Debug.Print "Opened."
End Sub

Private Sub WSC_IEventSink_OnClose()
    Debug.Print "Closed."
End Sub