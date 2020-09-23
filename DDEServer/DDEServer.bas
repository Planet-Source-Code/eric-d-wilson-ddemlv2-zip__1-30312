Attribute VB_Name = "modDDEServer"
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

' The DDE Instance ID and Item variables need to be global since there used by
' both the form and the code module.
Global lInstID As Long             ' DDE instance identifier.
Global hszDDEItemAdvise As Long    ' String handle for the item name.

'*************************************************************************
' DDEML Server Constants
'*************************************************************************
Public Const DDE_SERVER = "MyServer"
Public Const DDE_TOPIC = "MyTopic"
Public Const DDE_ADVISE = "MyAdvise"
Public Const DDE_REQUEST = "MyRequest"
Public Const DDE_POKE = "MyPoke"
Public Const DDE_COMMAND1 = "MAX"
Public Const DDE_COMMAND2 = "MIN"
Public Const DDE_COMMAND3 = "NORMAL"

' For some reason when performing an advise the data seems to be truncated by three
' characters. So to alleviate that problem we use a magic number constant.
Public Const MAGIC_NUMBER = 3

' This is just a string that we'll return whenever a client performs a DDE
' request if the input window on the form is empty. Otherwise we'll return
' whatever is in the window.
Public Const DDE_REQUEST_STRING = "The server input window is empty."

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
Public Const APPCMD_FILTERINITS = &H20&
Public Const XCLASS_BOOL = &H1000
Public Const XCLASS_DATA = &H2000
Public Const XCLASS_FLAGS = &H4000
Public Const XCLASS_NOTIFICATION = &H8000
Public Const XTYPF_NOBLOCK = &H2    ' CBR_BLOCK doesn't seem to work
Public Const XTYP_ADVDATA = (&H10 Or XCLASS_FLAGS)
Public Const XTYP_ADVREQ = (&H20 Or XCLASS_DATA Or XTYPF_NOBLOCK)
Public Const XTYP_ADVSTART = (XCLASS_BOOL Or &H30)
Public Const XTYP_ADVSTOP = (XCLASS_NOTIFICATION Or &H40)
Public Const XTYP_CONNECT = (XCLASS_BOOL Or &H60 Or XTYPF_NOBLOCK)
Public Const XTYP_CONNECT_CONFIRM = (XCLASS_NOTIFICATION Or &H70 Or XTYPF_NOBLOCK)
Public Const XTYP_DISCONNECT = (XCLASS_NOTIFICATION Or &HC0 Or XTYPF_NOBLOCK)
Public Const XTYP_ERROR = (XCLASS_NOTIFICATION Or &H0 Or XTYPF_NOBLOCK)
Public Const XTYP_EXECUTE = (XCLASS_FLAGS Or &H50)
Public Const XTYP_MASK = &HF0
Public Const XTYP_MONITOR = (XCLASS_NOTIFICATION Or &HF0 Or XTYPF_NOBLOCK)
Public Const XTYP_POKE = (XCLASS_FLAGS Or &H90)
Public Const XTYP_REGISTER = (XCLASS_NOTIFICATION Or &HA0 Or XTYPF_NOBLOCK)
Public Const XTYP_REQUEST = (XCLASS_DATA Or &HB0)
Public Const XTYP_SHIFT = 4  '  shift to turn XTYP_ into an index
Public Const XTYP_UNREGISTER = (XCLASS_NOTIFICATION Or &HD0 Or XTYPF_NOBLOCK)
Public Const XTYP_WILDCONNECT = (XCLASS_DATA Or &HE0 Or XTYPF_NOBLOCK)
Public Const XTYP_XACT_COMPLETE = (XCLASS_NOTIFICATION Or &H80)
Public Const CP_WINANSI = 1004      ' Default codepage for DDE conversations.
Public Const CP_WINUNICODE = 1200
Public Const CF_TEXT = 1
Public Const CBF_SKIP_ALLNOTIFICATIONS = &H3C0000
Public Const TIMEOUT_ASYNC = &HFFFF
Public Const DNS_REGISTER = &H1
Public Const DNS_UNREGISTER = &H2
Public Const DDE_FACK = &H8000
Public Const DDE_FBUSY = &H4000
Public Const DDE_FNOTPROCESSED = &H0
Public Const HDATA_APPOWNED = &H1

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
    
Public Declare Function DdeNameService Lib "user32" _
            (ByVal idInst As Long, _
            ByVal hsz1 As Long, _
            ByVal hsz2 As Long, _
            ByVal afCmd As Long) As Long

' Removed the alias.
Public Declare Function DdeConnect Lib "user32" _
    (ByVal idInst As Long, _
    ByVal hszService As Long, _
    ByVal hszTopic As Long, _
    pCC As Any) As Long
    
' Removed the alias.
Public Declare Function DdeDisconnect Lib "user32" _
    (ByVal hConv As Long) As Long
    
Public Declare Function DdeCreateStringHandle Lib "user32" _
    Alias "DdeCreateStringHandleA" _
    (ByVal idInst As Long, _
    ByVal psz As String, _
    ByVal iCodePage As Long) As Long
    
' Removed the alias.
Public Declare Function DdeFreeStringHandle Lib "user32" _
    (ByVal idInst As Long, _
    ByVal hsz As Long) As Long

' Removed the alias.
Public Declare Function DdeQueryString Lib "user32" _
    Alias "DdeQueryStringA" _
    (ByVal idInst As Long, _
    ByVal hsz As Long, _
    ByVal psz As String, _
    ByVal cchMax As Long, _
    ByVal iCodePage As Long) As Long

Public Declare Function DdeCreateDataHandle Lib "user32" _
    (ByVal idInst As Long, _
    ByVal pSrc As String, _
    ByVal cb As Long, _
    ByVal cbOff As Long, _
    ByVal hszItem As Long, _
    ByVal wFmt As Long, _
    ByVal afCmd As Long) As Long
            
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

' Removed the alias.
Public Declare Function DdeFreeDataHandle Lib "user32" _
    (ByVal hData As Long) As Long

' Removed the alias.
Public Declare Function DdeGetLastError Lib "user32" _
    (ByVal idInst As Long) As Long
    
Public Declare Function DdePostAdvise Lib "user32" _
    (ByVal idInst As Long, _
    ByVal hszTopic As Long, _
    ByVal hszItem As Long) As Long

Function DDECallback(ByVal uType As Long, _
                     ByVal uFmt As Long, _
                     ByVal hConv As Long, _
                     ByVal hszString1 As Long, _
                     ByVal hszString2 As Long, _
                     ByVal hData As Long, _
                     ByVal dwData1 As Long, _
                     ByVal dwData2 As Long) As Long
        
    Dim lSize As Long
    Dim lRet As Long
    Dim sBuffer As String
    Dim sRequest As String
    
    Select Case uType
        ' Process the connect transaction.
        Case XTYP_CONNECT
            ' Just return a positive acknowledgement. If we don't the conversation will
            ' never be completed between us and the client.
            lRet = DDE_FACK
            
        ' Process the request transaction.
        Case XTYP_REQUEST
            ' What's the size of the string?
            lSize = DdeQueryString(lInstID, hszString2, vbNullString, 0, CP_WINANSI)
            
            ' Allocate space for the string.
            sBuffer = Space(lSize)
            
            ' Grab the string.
            DdeQueryString lInstID, hszString2, sBuffer, lSize + 1, CP_WINANSI
                        
            ' Check to see if the client is requesting something we can supply.
            If (sBuffer = DDE_REQUEST) Then
                If (frmDDEServer.txtData.Text = "") Then
                    sRequest = DDE_REQUEST_STRING
                Else
                    sRequest = frmDDEServer.txtData.Text
                End If
                
                ' Create a data object and populate it with our string. Then return the object.
                lRet = DdeCreateDataHandle(lInstID, sRequest, Len(sRequest), 0, hszString2, CF_TEXT, 0)
            Else
                ' The client didn't ask nicely so we're not going to process the request.
                lRet = DDE_FNOTPROCESSED
            End If
            
        ' Process the execute transaction.
        Case XTYP_EXECUTE
            ' What's the size of the string?
            lSize = DdeGetData(ByVal hData, vbNullString, 0, 0)
            
            ' Allocate space for the buffer.
            sBuffer = Space(lSize)
            
            ' Grab the DDE data object.
            DdeGetData ByVal hData, sBuffer, lSize, 0
            
            sBuffer = UCase(sBuffer)
            
            ' Set the default return.
            lRet = DDE_FACK
            
            ' Did the client specify a command that we understand?
            If (sBuffer = DDE_COMMAND1) Then
                frmDDEServer.WindowState = vbMaximized
            ElseIf (sBuffer = DDE_COMMAND2) Then
                frmDDEServer.WindowState = vbMinimized
            ElseIf (sBuffer = DDE_COMMAND3) Then
                frmDDEServer.WindowState = vbNormal
            Else
                lRet = DDE_FNOTPROCESSED
            End If
            
        ' Process the poke request.
        Case XTYP_POKE
            lSize = DdeQueryString(lInstID, hszString2, vbNullString, 0, CP_WINANSI)
            sBuffer = Space(lSize)
            DdeQueryString lInstID, hszString2, sBuffer, lSize + 1, CP_WINANSI

            If (sBuffer = DDE_POKE) Then
                ' Since the client is sending data for an item that we support we can
                ' grab the data.
                lSize = DdeGetData(ByVal hData, vbNullString, 0, 0)
            
                ' Allocate space for the buffer.
                sBuffer = Space(lSize)
            
                ' Grab the DDE data object.
                DdeGetData ByVal hData, sBuffer, lSize, 0
                            
                ' Make sure the window is visible.
                frmDDEServer.WindowState = vbNormal
                
                ' Add the data to the text box.
                frmDDEServer.txtData.Text = sBuffer
                
                lRet = DDE_FACK
            Else
                lRet = DDE_FNOTPROCESSED
            End If
                                    
        Case XTYP_ADVREQ
            Debug.Print "We got the advise request. Conv: " & hConv
                                    
            If (frmDDEServer.txtData.Text = "") Then
                sRequest = DDE_REQUEST_STRING
            Else
                sRequest = frmDDEServer.txtData.Text
            End If
            
            ' Now create the data handle for the changed data.
            lRet = DdeCreateDataHandle(lInstID, sRequest, Len(sRequest) + MAGIC_NUMBER, 0, hszDDEItemAdvise, CF_TEXT, 0)
                        
        Case XTYP_ADVSTART
            Debug.Print "Start advise request made. Conv: " & hConv
            
            ' Enable the "Advise" button.
            frmDDEServer.cmdAdvise.Enabled = True
            
            lRet = DDE_FACK
            
        Case XTYP_ADVSTOP
            Debug.Print "Stop advise request made. Conv: " & hConv
            
            ' Disable the "Advise" button.
            frmDDEServer.cmdAdvise.Enabled = False
            
            lRet = DDE_FACK
            
    End Select
    
    ' Set the final callback return.
    DDECallback = lRet
    
End Function



