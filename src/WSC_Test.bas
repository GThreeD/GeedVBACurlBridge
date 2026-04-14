Attribute VB_Name = "WSC_Test"
Option Explicit

Private mClient As WSC_Client
Private mSink As WSC_TestSink

Public Sub WSC_Test_BridgeInfo()
    Debug.Print "BridgeName=" & WSCB_BridgeName()
    Debug.Print "BridgeApi=" & CStr(WSCB_BridgeApiVersion())
    Debug.Print "Libcurl=" & WSCB_LibcurlVersion()
    End Sub

Public Sub WSC_Test_Open()
    Set mClient = New WSC_Client
    Set mSink = New WSC_TestSink

    mClient.RegisterSink mSink

    Debug.Print "Connect=" & CStr(mClient.Connect("ws://127.0.0.1:8787/ws", 5000, True, True))
    If Not mClient.IsConnected Then
        Debug.Print "Error=" & mClient.LastError
    End If
End Sub

Public Sub WSC_Test_SendOne()
    Dim ok As Boolean

    If mClient Is Nothing Then
        WSC_Test_Open
    End If

    ok = mClient.SendText("1")
    Debug.Print "Send=" & CStr(ok)

    If Not ok Then
        Debug.Print "Error=" & mClient.LastError
    End If
End Sub

Public Sub WSC_Test_PollOnce()
    Dim didReceive As Boolean

    If mClient Is Nothing Then
        Debug.Print "Client not initialized."
        Exit Sub
    End If

    didReceive = mClient.PollOnce(500)
    Debug.Print "Polled=" & CStr(didReceive)

    If Len(mClient.LastError) > 0 Then
        Debug.Print "Error=" & mClient.LastError
    End If
End Sub

Public Sub WSC_Test_Roundtrip()
    WSC_Test_Open
    WSC_Test_SendOne
    WSC_Test_PollOnce
End Sub

Public Sub WSC_Test_Close()
    If Not mClient Is Nothing Then
        mClient.Disconnect
        Set mClient = Nothing
    End If

    Set mSink = Nothing
End Sub

Public Sub WSC_Test_Shutdown()
    WSC_Test_Close
    WSCB_GlobalShutdown
End Sub