Attribute VB_Name = "WSC_Bridge"
Option Explicit

Private Const WSC_LIB As String = "libcurl_vba_bridge.dylib"

Public Const WSC_BRIDGE_API_EXPECTED As Long = @WSC_BRIDGE_API_VERSION@
Public Const WSC_VERSION_MAJOR As Long = @WSC_VERSION_MAJOR@
Public Const WSC_VERSION_MINOR As Long = @WSC_VERSION_MINOR@
Public Const WSC_VERSION_PATCH As Long = @WSC_VERSION_PATCH@
Public Const WSC_VERSION_BUILD As Long = @WSC_VERSION_BUILD@
Public Const WSC_VERSION_STRING As String = "@WSC_VERSION_STRING@"

Private Const WSC_OK As Long = 0
Private Const WSC_AGAIN As Long = 81
Private Const WSC_GOT_NOTHING As Long = 52

Private Const WSC_ERRBUF_LEN As Long = 1024
Private Const WSC_RECVBUF_LEN As Long = 65536

Private Const RTLD_NOW As Long = &H2
Private Const RTLD_GLOBAL As Long = &H8

Private mGlobalInitialized As Boolean

Private Declare PtrSafe Function dlopen Lib "libc.dylib" (ByVal path As String, ByVal mode As Long) As LongPtr
Private Declare PtrSafe Function dlerror Lib "libc.dylib" () As String

Private Declare PtrSafe Function wsc_bridge_api_version_native Lib "libcurl_vba_bridge.dylib" Alias "wsc_bridge_api_version" () As Long
Private Declare PtrSafe Function wsc_bridge_name_native Lib "libcurl_vba_bridge.dylib" Alias "wsc_bridge_name" () As LongPtr

Private Declare PtrSafe Function wsc_global_init_native Lib "libcurl_vba_bridge.dylib" Alias "wsc_global_init" () As Long
Private Declare PtrSafe Sub wsc_global_cleanup_native Lib "libcurl_vba_bridge.dylib" Alias "wsc_global_cleanup" ()
Private Declare PtrSafe Function wsc_libcurl_version_native Lib "libcurl_vba_bridge.dylib" Alias "wsc_libcurl_version" () As LongPtr

Private Declare PtrSafe Function wsc_open_native Lib "libcurl_vba_bridge.dylib" Alias "wsc_open" ( _
    ByVal url As String, _
    ByVal timeoutMs As Long, _
    ByVal verifyPeer As Long, _
    ByVal verifyHost As Long, _
    ByRef outHandle As LongPtr, _
    ByVal errBuf As String, _
    ByVal errBufLen As Long) As Long

Private Declare PtrSafe Sub wsc_close_native Lib "libcurl_vba_bridge.dylib" Alias "wsc_close" ( _
    ByVal h As LongPtr)

Private Declare PtrSafe Function wsc_send_text_native Lib "libcurl_vba_bridge.dylib" Alias "wsc_send_text_utf8" ( _
    ByVal h As LongPtr, _
    ByVal bufPtr As LongPtr, _
    ByVal bufLen As LongPtr, _
    ByRef sentBytes As LongPtr, _
    ByVal errBuf As String, _
    ByVal errBufLen As Long) As Long

Private Declare PtrSafe Function wsc_recv_text_native Lib "libcurl_vba_bridge.dylib" Alias "wsc_recv_text_utf8" ( _
    ByVal h As LongPtr, _
    ByVal outBufPtr As LongPtr, _
    ByVal outBufLen As LongPtr, _
    ByRef receivedBytes As LongPtr, _
    ByVal errBuf As String, _
    ByVal errBufLen As Long) As Long

Private Declare PtrSafe Function wsc_last_error_text_native Lib "libcurl_vba_bridge.dylib" Alias "wsc_last_error_text" ( _
    ByVal code As Long) As LongPtr

Private Declare PtrSafe Function memcpy Lib "/usr/lib/libSystem.B.dylib" ( _
    ByVal destination As LongPtr, _
    ByVal source As LongPtr, _
    ByVal byteCount As LongPtr) As LongPtr

Private Sub EnsureLibraryLoaded()
    Static isLoaded As Boolean
    Static libHandle As LongPtr
    
    If isLoaded Then Exit Sub

    Dim libPath As String
    libpath = Environ("HOME") & "/curl/libcurl_vba_bridge.dylib"

    If Dir(libPath) = "" Then
        MsgBox "Bibliothek nicht gefunden: " & libpath, vbCritical
        Exit Sub
    End If

    libHandle = dlopen(libPath, RTLD_NOW Or RTLD_GLOBAL)

    If libHandle = 0 Then
        MsgBox "dlopen fehlgeschlagen: " & dlerror(), vbCritical
        Exit Sub
    End If

    isLoaded = true
End Sub

Public Function WSCB_BridgeApiVersion() As Long
    EnsureLibraryLoaded
    WSCB_BridgeApiVersion = wsc_bridge_api_version_native()
End Function

Public Function WSCB_BridgeName() As String
    EnsureLibraryLoaded
    Dim p As LongPtr
    p = wsc_bridge_name_native()
    If p <> 0 Then
        WSCB_BridgeName = WSCB_CStringToString(p)
    End If
End Function

Public Function WSCB_LibcurlVersion() As String
    EnsureLibraryLoaded
    Dim p As LongPtr
    p = wsc_libcurl_version_native()
    If p <> 0 Then
        WSCB_LibcurlVersion = WSCB_CStringToString(p)
    End If
End Function

Public Function WSCB_EnsureCompatibleBridge(ByRef errorText As String) As Boolean
    EnsureLibraryLoaded
    Dim apiVersion As Long

    apiVersion = WSCB_BridgeApiVersion()
    If apiVersion <> WSC_BRIDGE_API_EXPECTED Then
        errorText = "Bridge API mismatch. Expected=" & CStr(WSC_BRIDGE_API_EXPECTED) & ", Actual=" & CStr(apiVersion)
        Exit Function
    End If

    WSCB_EnsureCompatibleBridge = True
End Function

Public Function WSCB_EnsureGlobalInit(ByRef errorText As String) As Boolean
    EnsureLibraryLoaded
    Dim rc As Long

    If mGlobalInitialized Then
        WSCB_EnsureGlobalInit = True
        Exit Function
    End If

    rc = wsc_global_init_native()
    If rc <> WSC_OK Then
        errorText = "wsc_global_init failed: " & CStr(rc)
        Exit Function
    End If

    mGlobalInitialized = True
    WSCB_EnsureGlobalInit = True
End Function

Public Sub WSCB_GlobalShutdown()
    EnsureLibraryLoaded
    If mGlobalInitialized Then
        On Error Resume Next
        wsc_global_cleanup_native
        On Error GoTo 0
        mGlobalInitialized = False
    End If
End Sub

Public Function WSCB_Open(ByVal wsUrl As String, ByVal timeoutMs As Long, ByVal verifyPeer As Boolean, ByVal verifyHost As Boolean, ByRef outHandle As LongPtr, ByRef errorText As String) As Boolean
    EnsureLibraryLoaded
    Dim rc As Long
    Dim errBuf As String
    Dim compatible As Boolean
    Dim initialized As Boolean

    errorText = vbNullString
    outHandle = 0

    compatible = WSCB_EnsureCompatibleBridge(errorText)
    If Not compatible Then Exit Function

    initialized = WSCB_EnsureGlobalInit(errorText)
    If Not initialized Then Exit Function

    errBuf = String$(WSC_ERRBUF_LEN, vbNullChar)

    rc = wsc_open_native(wsUrl, timeoutMs, Abs(verifyPeer), IIf(verifyHost, 2, 0), outHandle, errBuf, Len(errBuf))
    If rc <> WSC_OK Or outHandle = 0 Then
        errorText = WSCB_TrimNulls(errBuf)
        If Len(errorText) = 0 Then
            errorText = "wsc_open failed: " & WSCB_ErrorText(rc)
        End If
        Exit Function
    End If

    WSCB_Open = True
End Function

Public Function WSCB_SendText(ByVal handle As LongPtr, ByVal textValue As String, ByRef errorText As String) As Boolean
    Dim rc As Long
    Dim errBuf As String
    Dim sentBytes As LongPtr
    Dim utf8() As Byte
    Dim bufPtr As LongPtr
    Dim bufLen As LongPtr

    EnsureLibraryLoaded
    errorText = vbNullString

    If handle = 0 Then
        errorText = "Invalid handle."
        Exit Function
    End If

    utf8 = WSCB_StringToWireBytes(textValue)
    errBuf = String$(WSC_ERRBUF_LEN, vbNullChar)

    If (Not Not utf8) <> 0 Then
        bufPtr = VarPtr(utf8(LBound(utf8)))
        bufLen = UBound(utf8) - LBound(utf8) + 1
    Else
        bufPtr = 0
        bufLen = 0
    End If

    rc = wsc_send_text_native(handle, bufPtr, bufLen, sentBytes, errBuf, Len(errBuf))
    If rc <> WSC_OK Then
        errorText = WSCB_TrimNulls(errBuf)
        If Len(errorText) = 0 Then
            errorText = "wsc_send_text failed: " & WSCB_ErrorText(rc)
        End If
        Exit Function
    End If

    WSCB_SendText = True
End Function

Public Function WSCB_TryReceiveText(ByVal handle As LongPtr, ByVal timeoutMs As Long, ByRef didReceive As Boolean, ByRef textValue As String, ByRef peerClosed As Boolean, ByRef errorText As String) As Boolean
    Dim startedAt As Double
    Dim rc As Long
    Dim recvCount As LongPtr
    Dim errBuf As String
    Dim outBuf() As Byte

    EnsureLibraryLoaded

    didReceive = False
    peerClosed = False
    textValue = vbNullString
    errorText = vbNullString

    If handle = 0 Then
        errorText = "Invalid handle."
        Exit Function
    End If

    startedAt = Timer
    ReDim outBuf(0 To WSC_RECVBUF_LEN - 1)

    Do
        errBuf = String$(WSC_ERRBUF_LEN, vbNullChar)
        recvCount = 0

        rc = wsc_recv_text_native(handle, VarPtr(outBuf(0)), WSC_RECVBUF_LEN, recvCount, errBuf, Len(errBuf))

        If rc = WSC_OK Then
            If recvCount > 0 Then
                didReceive = True
                textValue = WSCB_WireBytesToString(outBuf, CLng(recvCount))
                WSCB_TryReceiveText = True
                Exit Function
            End If

        ElseIf rc = WSC_AGAIN Then
            DoEvents

        ElseIf rc = WSC_GOT_NOTHING Then
            peerClosed = True
            WSCB_TryReceiveText = True
            Exit Function

        Else
            errorText = WSCB_TrimNulls(errBuf)
            If Len(errorText) = 0 Then
                errorText = "wsc_recv_text failed: " & WSCB_ErrorText(rc)
            End If
            Exit Function
        End If

        If WSCB_ElapsedMs(startedAt) >= timeoutMs Then
            WSCB_TryReceiveText = True
            Exit Function
        End If
    Loop
End Function

Public Sub WSCB_Close(ByVal handle As LongPtr)
    EnsureLibraryLoaded
    If handle <> 0 Then
        On Error Resume Next
        wsc_close_native handle
        On Error GoTo 0
    End If
End Sub

Private Function WSCB_ErrorText(ByVal rc As Long) As String
    EnsureLibraryLoaded
    Dim p As LongPtr
    p = wsc_last_error_text_native(rc)

    If p = 0 Then
        WSCB_ErrorText = "curl error " & CStr(rc)
    Else
        WSCB_ErrorText = WSCB_CStringToString(p)
    End If
End Function

Private Function WSCB_CStringToString(ByVal pText As LongPtr) As String
    Dim bytes() As Byte
    Dim i As Long
    Dim oneByte(0 To 0) As Byte

    If pText = 0 Then Exit Function

    ReDim bytes(0 To 4095)

    For i = 0 To 4095
        memcpy VarPtr(oneByte(0)), pText + i, 1
        If oneByte(0) = 0 Then Exit For
        bytes(i) = oneByte(0)
    Next i

    If i = 0 Then
        WSCB_CStringToString = vbNullString
    Else
        ReDim Preserve bytes(0 To i - 1)
        WSCB_CStringToString = WSCB_Utf8BytesToString(bytes)
    End If
End Function

Private Function WSCB_TrimNulls(ByVal textValue As String) As String
    EnsureLibraryLoaded
    Dim p As Long
    p = InStr(1, textValue, vbNullChar)

    If p > 0 Then
        WSCB_TrimNulls = Left$(textValue, p - 1)
    Else
        WSCB_TrimNulls = textValue
    End If
End Function

Private Function WSCB_ElapsedMs(ByVal startedAt As Double) As Long
    EnsureLibraryLoaded
    Dim nowValue As Double
    nowValue = Timer

    If nowValue >= startedAt Then
        WSCB_ElapsedMs = CLng((nowValue - startedAt) * 1000#)
    Else
        WSCB_ElapsedMs = CLng(((86400# - startedAt) + nowValue) * 1000#)
    End If
End Function

Private Function WSCB_StringToWireBytes(ByVal textValue As String) As Byte()
    Dim result() As Byte
    Dim i As Long
    Dim ch As Long

    If Len(textValue) = 0 Then
        ReDim result(0 To -1)
        WSCB_StringToWireBytes = result
        Exit Function
    End If

    ReDim result(0 To Len(textValue) - 1)

    For i = 1 To Len(textValue)
        ch = AscW(Mid$(textValue, i, 1))
        If ch < 0 Or ch > 127 Then
            Err.Raise vbObjectError + 7301, "WSCB_StringToWireBytes", _
                "Non-ASCII character in wire text at position " & CStr(i) & "."
        End If
        result(i - 1) = ch
    Next i

    WSCB_StringToWireBytes = result
End Function

Private Function WSCB_WireBytesToString(ByRef bytes() As Byte, ByVal byteCount As Long) As String
    Dim result As String
    Dim i As Long

    If byteCount <= 0 Then
        WSCB_WireBytesToString = vbNullString
        Exit Function
    End If

    result = String$(byteCount, vbNullChar)

    For i = 1 To byteCount
        Mid$(result, i, 1) = Chr$(bytes(LBound(bytes) + i - 1))
    Next i

    WSCB_WireBytesToString = result
End Function