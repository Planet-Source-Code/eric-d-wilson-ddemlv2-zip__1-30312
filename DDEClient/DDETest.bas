Attribute VB_Name = "modDDETest"
'*************************************************************************
'
' Copyright (c) 2001
'
' Author:   Eric D. Wilson
' Email:    edwilson97@yahoo.com
'
' NOTE:
' This code module contains all of the DDEML declarations that I use throughout
' the application. I've tried to comment any declaration/type changes I've made.
'
'*************************************************************************

Option Explicit

Global g_lInstID As Long
Global g_hService As Long
Global g_hService2 As Long
Global g_hTopic As Long
Global g_hTopic2 As Long
Global g_hItem As Long
Global g_hDDEConv As Long
Global g_hDDEConvList As Long
Global g_hDDEPrevConv As Long
Global g_aryConvID() As Long

' See note in cmdRequest_Click() concerning the magic number.
Public Const MAGIC_NUMBER = 3

'*************************************************************************
' DDEML Return Values
'*************************************************************************
Public Const DMLERR_NO_ERROR = 0
Public Const DMLERR_ADVACKTIMEOUT = &H4000
Public Const DMLERR_BUSY = &H4001
Public Const DMLERR_DATAACKTIMEOUT = &H4002
Public Const DMLERR_DLL_NOT_INITIALIZED = &H4003
Public Const DMLERR_DLL_USAGE = &H4004
Public Const DMLERR_EXECACKTIMEOUT = &H4005
Public Const DMLERR_INVALIDPARAMETER = &H4006
Public Const DMLERR_LOW_MEMORY = &H4007
Public Const DMLERR_MEMORY_ERROR = &H4008
Public Const DMLERR_NOTPROCESSED = &H4009
Public Const DMLERR_NO_CONV_ESTABLISHED = &H400A
Public Const DMLERR_POKEACKTIMEOUT = &H400B
Public Const DMLERR_POSTMSG_FAILED = &H400C
Public Const DMLERR_REENTRANCY = &H400D
Public Const DMLERR_SERVER_DIED = &H400E
Public Const DMLERR_SYS_ERROR = &H400F
Public Const DMLERR_UNADVACKTIMEOUT = &H4010
Public Const DMLERR_UNFOUND_QUEUE_ID = &H4011

'*************************************************************************
' DDEML Flags
'*************************************************************************
Public Const XCLASS_BOOL = &H1000&
Public Const XCLASS_DATA = &H2000&
Public Const XCLASS_FLAGS = &H4000&
Public Const XCLASS_NOTIFICATION = &H8000&
Public Const XTYPF_NOBLOCK = &H2&    ' CBR_BLOCK doesn't seem to work
Public Const XTYP_ADVDATA = (&H10& Or XCLASS_FLAGS)
Public Const XTYP_ADVREQ = (&H20& Or XCLASS_DATA Or XTYPF_NOBLOCK)
Public Const XTYP_ADVSTART = (XCLASS_BOOL Or &H30&)
Public Const XTYP_ADVSTOP = (XCLASS_NOTIFICATION Or &H40&)
Public Const XTYP_CONNECT = (XCLASS_BOOL Or &H60& Or XTYPF_NOBLOCK)
Public Const XTYP_CONNECT_CONFIRM = (XCLASS_NOTIFICATION Or &H70& Or XTYPF_NOBLOCK)
Public Const XTYP_DISCONNECT = (XCLASS_NOTIFICATION Or &HC0& Or XTYPF_NOBLOCK)
Public Const XTYP_ERROR = (XCLASS_NOTIFICATION Or &H0& Or XTYPF_NOBLOCK)
Public Const XTYP_EXECUTE = (XCLASS_FLAGS Or &H50&)
Public Const XTYP_MASK = &HF0&
Public Const XTYP_MONITOR = (XCLASS_NOTIFICATION Or &HF0& Or XTYPF_NOBLOCK)
Public Const XTYP_POKE = (XCLASS_FLAGS Or &H90&)
Public Const XTYP_REGISTER = (XCLASS_NOTIFICATION Or &HA0& Or XTYPF_NOBLOCK)
Public Const XTYP_REQUEST = (XCLASS_DATA Or &HB0&)
Public Const XTYP_SHIFT = 4  '  shift to turn XTYP_ into an index
Public Const XTYP_UNREGISTER = (XCLASS_NOTIFICATION Or &HD0& Or XTYPF_NOBLOCK)
Public Const XTYP_WILDCONNECT = (XCLASS_DATA Or &HE0& Or XTYPF_NOBLOCK)
Public Const XTYP_XACT_COMPLETE = (XCLASS_NOTIFICATION Or &H80&)
Public Const CP_WINANSI = 1004      ' Default codepage for DDE conversations.
Public Const CP_WINUNICODE = 1200
Public Const CF_TEXT = 1
Public Const CBF_SKIP_ALLNOTIFICATIONS = &H3C0000
Public Const APPCLASS_MONITOR = &H1
Public Const APPCMD_CLIENTONLY = &H10&
Public Const MF_CALLBACKS = &H8000000
Public Const MF_CONV = &H40000000
Public Const MF_ERRORS = &H10000000
Public Const MF_HSZ_INFO = &H1000000
Public Const MF_LINKS = &H20000000
Public Const MF_POSTMSGS = &H4000000
Public Const MF_SENDMSGS = &H2000000
Public Const TIMEOUT_ASYNC = &HFFFF
Public Const QID_SYNC = &HFFFF
Public Const DDE_FACK = &H8000
Public Const DDE_FBUSY = &H4000
Public Const DDE_FNOTPROCESSED = &H0
Public Const EC_ENABLEALL = 0

'*************************************************************************
' DDEML Type Declarations
'*************************************************************************
Public Type SECURITY_QUALITY_OF_SERVICE
    Length As Long
    Impersonationlevel As Integer
    ContextTrackingMode As Integer
    EffectiveOnly As Long
End Type

Public Type CONVCONTEXT
    cb As Long
    wFlags As Long
    wCountryID As Long
    iCodePage As Long
    dwLangID As Long
    dwSecurity As Long
    qos As SECURITY_QUALITY_OF_SERVICE
End Type

Public Type CONVINFO
    cb As Long
    hUser As Long
    hConvPartner As Long
    hszSvcPartner As Long
    hszServiceReq As Long
    hszTopic As Long
    hszItem As Long
    wFmt As Long
    wType As Long
    wStatus As Long
    wConvst As Long
    wLastError As Long
    hConvList As Long
    ConvCtxt As CONVCONTEXT
    hwnd As Long
    hwndPartner As Long
End Type

'*************************************************************************
' DDEML Function Declarations
'*************************************************************************
Public Declare Function DdeInitialize Lib "user32" Alias "DdeInitializeA" _
    (pidInst As Long, _
    ByVal pfnCallback As Long, _
    ByVal afCmd As Long, _
    ByVal ulRes As Long) As Integer
    
' Removed the alias.
Public Declare Function DdeUninitialize Lib "user32" _
    (ByVal idInst As Long) As Long
    
' Removed the alias.
Public Declare Function DdeConnect Lib "user32" _
    (ByVal idInst As Long, _
    ByVal hszService As Long, _
    ByVal hszTopic As Long, _
    pCC As Any) As Long
    
' Removed the alias.
Public Declare Function DdeDisconnect Lib "user32" _
    (ByVal hConv As Long) As Long
    
Public Declare Function DdeCreateStringHandle Lib "user32" Alias "DdeCreateStringHandleA" _
    (ByVal idInst As Long, _
    ByVal psz As String, _
    ByVal iCodePage As Long) As Long
    
' Removed the alias.
Public Declare Function DdeFreeStringHandle Lib "user32" _
    (ByVal idInst As Long, _
    ByVal hsz As Long) As Long
            
' Removed the alias and changed the first parameter from "ByRef pData as Byte"
' to "ByVal pData as String".
Public Declare Function DdeClientTransaction Lib "user32" _
    (ByVal pData As String, _
    ByVal cbData As Long, _
    ByVal hConv As Long, _
    ByVal hszItem As Long, _
    ByVal wFmt As Long, _
    ByVal wType As Long, _
    ByVal dwTimeout As Long, _
    pdwResult As Long) As Long
    
' The API loader provides an alias of "DdeGetDataA" for this function.
' You need to remove it because the DLL entry point can't be found for
' the alias.
Public Declare Function DdeGetData Lib "user32" _
    (ByVal hData As Long, _
    ByVal pDst As String, _
    ByVal cbMax As Long, _
    ByVal cbOff As Long) As Long

Public Declare Function DdeQueryConvInfo Lib "user32" _
    (ByVal hConv As Long, _
    ByVal idTransaction As Long, _
    pConvInfo As CONVINFO) As Long

Public Declare Function DdeQueryNextServer Lib "user32" _
    (ByVal hConvList As Long, _
    ByVal hConvPrev As Long) As Long

Public Declare Function DdeConnectList Lib "user32" _
    (ByVal idInst As Long, _
    ByVal hszService As Long, _
    ByVal hszTopic As Long, _
    ByVal hConvList As Long, _
    pCC As CONVCONTEXT) As Long

Public Declare Function DdeDisconnectList Lib "user32" _
    (ByVal hConvList As Long) As Long

Public Declare Function DdeQueryString Lib "user32" _
    Alias "DdeQueryStringA" _
    (ByVal idInst As Long, _
    ByVal hsz As Long, _
    ByVal psz As String, _
    ByVal cchMax As Long, _
    ByVal iCodePage As Long) As Long

' Removed the alias.
Public Declare Function DdeFreeDataHandle Lib "user32" _
    (ByVal hData As Long) As Long

' Removed the alias.
Public Declare Function DdeGetLastError Lib "user32" _
    (ByVal idInst As Long) As Long

Public Declare Function DdeEnableCallback Lib "user32" _
    (ByVal idInst As Long, _
    ByVal hConv As Long, _
    ByVal wCmd As Long) As Long

Public Function DDECallback(ByVal uType As Long, ByVal uFmt As Long, ByVal hConv As Long, ByVal hszString1 As Long, ByVal hszString2 As Long, ByVal hData As Long, ByVal dwData1 As Long, ByVal dwData2 As Long) As Long
    
Dim lSize As Long
Dim sBuffer As String
Dim Ret As Long
    
    Debug.Print "In client callback. uType: " & uType
    
    Select Case uType
        
        ' This is th eevent you'll receive when a server sends you a advisment.
        Case XTYP_ADVDATA
            Debug.Print "XTYP_ADVDATA"
            
            lSize = DdeGetData(hData, vbNullString, 0, 0)
            
            ' If size is 0 then there's no data to grab.
            If (lSize > 0) Then
                
                ' Allocate a buffer for the return data.
                sBuffer = String$(lSize - MAGIC_NUMBER, 0)
                            
                ' Grab the data.
                lSize = DdeGetData(hData, sBuffer, Len(sBuffer), 0)
    
                ' Print the contents of the buffer.
                frmDDEApp.txtValue.Text = sBuffer
            End If
            
        Case XTYP_ADVSTART
            Debug.Print "XTYP_ADVSTART"
            
        Case XTYP_ADVSTOP
            Debug.Print "XTYP_ADVSTOP"
        
        Case XTYP_CONNECT
            Debug.Print "XTYP_CONNECT"
            
        Case XTYP_CONNECT_CONFIRM
            Debug.Print "XTYP_CONNECT_CONFIRM"
            
        Case XTYP_DISCONNECT
            Debug.Print "XTYP_DISCONNECT"
            
        Case XTYP_ERROR
            Debug.Print "XTYP_ERROR"
            
        Case XTYP_EXECUTE
            Debug.Print "XTYP_EXECUTE"
            
        Case XTYP_MASK
            Debug.Print "XTYP_MASK"
            
        Case XTYP_MONITOR
            Debug.Print "XTYP_MONITOR"
            
        Case XTYP_POKE
            Debug.Print "XTYP_POKE"
            
        Case XTYP_REGISTER
            Debug.Print "XTYP_REGISTER"
            g_hService2 = hszString2
            
            lSize = DdeQueryString(g_lInstID, hszString2, vbNullString, 0, CP_WINANSI)
            sBuffer = Space(lSize)
            DdeQueryString g_lInstID, hszString2, sBuffer, lSize + 1, CP_WINANSI

            sBuffer = UCase(sBuffer)
            
        Case XTYP_REQUEST
            Debug.Print "XTYP_REQUEST"
            
        Case XTYP_SHIFT
            Debug.Print "XTYP_SHIFT"
            
        Case XTYP_UNREGISTER
            Debug.Print "XTYP_UNREGISTER"
            
        Case XTYP_WILDCONNECT
            Debug.Print "XTYP_WILDCONNECT"
            
        Case XTYP_XACT_COMPLETE
            Debug.Print "XTYP_XACT_COMPLETE"
                
    End Select
    
    DDECallback = 0

End Function

