VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "WSC_Client"
Option Explicit
Implements WSC_ISender

Private mHandle As LongPtr
Private mConnected As Boolean
Private mLastError As String
Private mUrl As String
Private mSinks As Collection

Private Sub Class_Initialize()
    Set mSinks = New Collection
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Me.Disconnect
    On Error GoTo 0
End Sub

Public Function Connect(ByVal wsUrl As String, Optional ByVal timeoutMs As Long = 5000, Optional ByVal verifyPeer As Boolean = True, Optional ByVal verifyHost As Boolean = True) As Boolean
    Connect = WSC_ISender_Connect(wsUrl, timeoutMs, verifyPeer, verifyHost)
End Function

Public Function SendText(ByVal textValue As String) As Boolean
    SendText = WSC_ISender_SendText(textValue)
End Function

Public Function PollOnce(Optional ByVal timeoutMs As Long = 25) As Boolean
    PollOnce = WSC_ISender_PollOnce(timeoutMs)
End Function

Public Sub Disconnect()
    WSC_ISender_Close
End Sub

Public Property Get IsConnected() As Boolean
    IsConnected = mConnected
End Property

Public Property Get LastError() As String
    LastError = mLastError
End Property

Public Property Get Url() As String
    Url = mUrl
End Property

Public Sub RegisterSink(ByVal sink As WSC_IEventSink)
    Dim currentSink As WSC_IEventSink

    If sink Is Nothing Then Exit Sub

    For Each currentSink In mSinks
        If ObjPtr(currentSink) = ObjPtr(sink) Then
            Exit Sub
        End If
    Next currentSink

    mSinks.Add sink
End Sub

Public Sub UnregisterSink(ByVal sink As WSC_IEventSink)
    Dim currentSink As WSC_IEventSink
    Dim rebuilt As Collection

    If sink Is Nothing Then Exit Sub

    Set rebuilt = New Collection
    For Each currentSink In mSinks
        If ObjPtr(currentSink) <> ObjPtr(sink) Then
            rebuilt.Add currentSink
        End If
    Next currentSink

    Set mSinks = rebuilt
End Sub

Private Function WSC_ISender_Connect(ByVal wsUrl As String, ByVal timeoutMs As Long, ByVal verifyPeer As Boolean, ByVal verifyHost As Boolean) As Boolean
    Dim handleValue As LongPtr
    Dim errorText As String

    Me.Disconnect
    mLastError = vbNullString

    If Not WSCB_Open(wsUrl, timeoutMs, verifyPeer, verifyHost, handleValue, errorText) Then
        mLastError = errorText
        NotifyError mLastError
        Exit Function
    End If

    mHandle = handleValue
    mConnected = True
    mUrl = wsUrl

    NotifyOpen
    WSC_ISender_Connect = True
End Function

Private Function WSC_ISender_SendText(ByVal textValue As String) As Boolean
    Dim errorText As String

    If Not mConnected Or mHandle = 0 Then
        mLastError = "Not connected."
        NotifyError mLastError
        Exit Function
    End If

    If Not WSCB_SendText(mHandle, textValue, errorText) Then
        mLastError = errorText
        NotifyError mLastError
        Exit Function
    End If

    WSC_ISender_SendText = True
End Function

Private Function WSC_ISender_PollOnce(ByVal timeoutMs As Long) As Boolean
    Dim didReceive As Boolean
    Dim peerClosed As Boolean
    Dim textValue As String
    Dim errorText As String

    If Not mConnected Or mHandle = 0 Then
        mLastError = "Not connected."
        Exit Function
    End If

    If Not WSCB_TryReceiveText(mHandle, timeoutMs, didReceive, textValue, peerClosed, errorText) Then
        mLastError = errorText
        NotifyError mLastError
        InternalClose True
        Exit Function
    End If

    If peerClosed Then
        InternalClose True
        Exit Function
    End If

    If didReceive Then
        NotifyMessage textValue
        WSC_ISender_PollOnce = True
    End If
End Function

Private Sub WSC_ISender_Close()
    InternalClose True
End Sub

Private Property Get WSC_ISender_IsConnected() As Boolean
    WSC_ISender_IsConnected = mConnected
End Property

Private Property Get WSC_ISender_LastError() As String
    WSC_ISender_LastError = mLastError
End Property

Private Sub InternalClose(ByVal raiseCloseEvent As Boolean)
    Dim wasConnected As Boolean
    wasConnected = mConnected

    If mHandle <> 0 Then
        WSCB_Close mHandle
    End If

    mHandle = 0
    mConnected = False
    mUrl = vbNullString

    If raiseCloseEvent And wasConnected Then
        NotifyClose
    End If
End Sub

Private Sub NotifyMessage(ByVal textValue As String)
    Dim sink As WSC_IEventSink

    For Each sink In mSinks
        On Error Resume Next
        sink.OnMessage textValue
        On Error GoTo 0
    Next sink
End Sub

Private Sub NotifyError(ByVal errorText As String)
    Dim sink As WSC_IEventSink

    For Each sink In mSinks
        On Error Resume Next
        sink.OnError errorText
        On Error GoTo 0
    Next sink
End Sub

Private Sub NotifyOpen()
    Dim sink As WSC_IEventSink

    For Each sink In mSinks
        On Error Resume Next
        sink.OnOpen
        On Error GoTo 0
    Next sink
End Sub

Private Sub NotifyClose()
    Dim sink As WSC_IEventSink

    For Each sink In mSinks
        On Error Resume Next
        sink.OnClose
        On Error GoTo 0
    Next sink
End Sub