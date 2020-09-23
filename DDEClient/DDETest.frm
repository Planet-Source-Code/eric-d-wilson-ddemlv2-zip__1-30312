VERSION 5.00
Begin VB.Form frmDDEApp 
   Caption         =   "DDEML Test Application"
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   ScaleHeight     =   3510
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   3255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   7455
      Begin VB.ComboBox cboItem 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "DDETest.frx":0000
         Left            =   120
         List            =   "DDETest.frx":0010
         TabIndex        =   10
         Top             =   1920
         Width           =   3375
      End
      Begin VB.CommandButton cmdUninitialize 
         Caption         =   "Uninitialize"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5640
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdInitialize 
         Caption         =   "Initialize"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3720
         TabIndex        =   0
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdStopAdv 
         Caption         =   "Stop Advise"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5640
         TabIndex        =   5
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CommandButton cmdStartAdv 
         Caption         =   "Start Advise"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5640
         TabIndex        =   3
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5640
         TabIndex        =   7
         Top             =   2520
         Width           =   1695
      End
      Begin VB.TextBox txtService 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   3375
      End
      Begin VB.TextBox txtTopic 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   3375
      End
      Begin VB.TextBox txtValue 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   2640
         Width           =   3375
      End
      Begin VB.CommandButton cmdExecute 
         Caption         =   "Execute"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3720
         TabIndex        =   2
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton cmdPoke 
         Caption         =   "Poke"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3720
         TabIndex        =   4
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CommandButton cmdRequest 
         Caption         =   "Request"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3720
         TabIndex        =   6
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Service:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Topic:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Item:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Value:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2400
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmDDEApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************
'
' Copyright (c) 2001
'
' Author:   Eric D. Wilson
' Email:    edwilson97@yahoo.com
'
' NOTE:
' This application demonstrates how to program DDE through the DDE Management Library.
' This is an alternative to the standard form based DDE mechanism that is provided with
' Visual Basic. There is very little documentation on using the DDEML through Visual Basic
' so it takes a little trial and error. A perfect example is the function declaration  for
' DdeGetData(). The decalration that you get from the API loader is incorrect. You have to
' remove the Alias otherwise it won't work. There are some other changes that I had to make
' like variable types. I've tried to add a comment for all the ones I've changed.
'
' This app provides three DDE communication mechanisms including Execute, Request, and
' Poke. There are two other aspects that include the DDE subsystem initialization and the
' DDE conversation connection. Once you initialize the DDE subsystem you can open and close
' DDE conversations as often as you want without having to uninitialize and reinitialize.
'
' Here are some tests you can run.
'
' UPDATE:   11/30/2001
' This application has been updated to incorporate an advise loop example. By default the
' application is setup to work agaist the DDEML demo server. If you want to run the examples
' against Excel, IE, ... all you have to do is initialize DDE and then clear out the textbox
' controls and enter the information specific to the other DDE servers that you want to
' interface with.
' ************************************************************************
'
' Execute:
'
' Start Internet Explorer. On the DDETest form add the following data.
'
'   Service:    iexplore
'   Topic:      WWW_OpenURL
'   Value:      http://www.microsoft.com
'
' Now press the execute button and you should be taken to the Microsoft homepage. Note
' that eventhough the command executed you received a failure return code. I can't explain
' it. It's Microsoft logic. :-) Actually if you check out the Spyworks DDE definition you'll
' see that there are a host of other parameters to include.
'
' ************************************************************************
'
' Poke:
'
' Start Microsoft Excel. Create Book1 if it doesn't come up by default.
' On the DDETest form add the following data.
'
'   Service:    excel
'   Topic:      book1
'   Item:       R1C1    - That stands for row 1, colum 1 (cell A1).
'   Value:      This is a test.
'
' When you press the Poke button that string value is added to the cell A1 on the visible sheet
' of book1. This time you should get a successful return value.
'
'*************************************************************************
'
' Request:
'
' With Microsoft Excel still open add the following data to cell A5.
'
'   "This is another test"
'
' Make sure to hit enter or click on another cell after adding text.
' If the cell that your trying to request data of still has the focus when
' you make the request, the request will fail.
'
' Now add the following data to the DDETest form.
'
'   Service:    excel
'   Topic:      book1
'   Item:       R5C1    - That stands for row 5, colum 1 (cell A5).
'
' When you press the Request button you should see the contents of cell A5 printed
' to the debug (Immediate) window.
'
' Note: DDE Request can also be used in place of DDE Execute to run commands
' on a DDE service.
'
'*************************************************************************
'
' Advise:
'
' NOTE: The advise example requires the demo DDE server.
'
' An advise loop is a mechanism for receiving notification from a DDE server
' when certain data changes. This is the equivilent of a "hot link" in
' traditional VB DDE. For this example you need to start the demo DDE server
' that should be included in the zip file. Once that is done you can start the
' DDE client and initialize the DDE subsystem.
'
' Now that DDE is initialized we need to inform the DDE server that we're
' interested in receiving notifications. To do that we create an Advise Loop. On
' the client form select the "Item" drop down and choose "MyAdvise". Next, click
' the "Start Advise" button. You have now created an advise loop with the server.
'
' Now switch to the server form and type something into the data window. Once you've
' done that click the "Advise" button. You should see the data that you entered on
' the server appear in the "Value" box of the client form. That's it!
'
' When your done testing the advise loop re-select the "MyAdvise" item and click
' the "Stop Advise" button to terminate the advise loop with the server.
'
'*************************************************************************

Option Explicit

Dim bAdvise As Boolean

Private Sub cboItem_Click()
    
    ' Adjust button states.
    Select Case cboItem.Text
        Case "<None>"
            cmdExecute.Enabled = True
            cmdPoke.Enabled = False
            cmdRequest.Enabled = False
            cmdStopAdv.Enabled = False
            cmdStartAdv.Enabled = False
        
        Case "MyAdvise"
            If (bAdvise) Then
                cmdStopAdv.Enabled = True
            Else
                cmdStartAdv.Enabled = True
            End If
            cmdExecute.Enabled = False
            cmdPoke.Enabled = False
            cmdRequest.Enabled = False
            
        Case "MyPoke"
            cmdExecute.Enabled = False
            cmdPoke.Enabled = True
            cmdRequest.Enabled = False
            cmdStopAdv.Enabled = False
            cmdStartAdv.Enabled = False
        
        Case "MyRequest"
            cmdExecute.Enabled = False
            cmdPoke.Enabled = False
            cmdRequest.Enabled = True
            cmdStopAdv.Enabled = False
            cmdStartAdv.Enabled = False
    End Select
End Sub

Private Sub cmdInitialize_Click()
    
Dim oCtl As Control
    
    Debug.Print "------------------- Begin DDE Test -----------------------"
    
    g_lInstID = 0
    
    ' Initialize the DDE subsystem. This only needs to be done once.
    If DdeInitialize(g_lInstID, AddressOf DDECallback, APPCMD_CLIENTONLY Or MF_SENDMSGS Or MF_POSTMSGS, 0) Then
        
        Debug.Print "DDE Initialize Failure."
        TranslateError
    
    Else
        
        Debug.Print "DDE Initialize Success."
    
    End If
        
    ' Enable the command buttons.
    For Each oCtl In Controls
    
        If ((TypeOf oCtl Is TextBox) Or (TypeOf oCtl Is ComboBox)) And (oCtl.Enabled = False) Then
            oCtl.Enabled = True
        End If
        
    Next
    
    cmdInitialize.Enabled = False
    cmdUninitialize.Enabled = True
    cmdClear.Enabled = True
    cboItem.ListIndex = 0
    
End Sub

Private Sub cmdUninitialize_Click()
    
Dim oCtl As Control
    
    ' Make sure we don't have any open connections.
    If (g_hDDEConv <> 0) Then
        DDE_Disconnect
    End If
    
    ' Tear down the initialized instance.
    If g_lInstID Then
        
        If DdeUninitialize(g_lInstID) Then
            
            Debug.Print "DDE Uninitialize Success."
        
        Else
            
            Debug.Print "DDE Uninitialize Failure."
            TranslateError
        
        End If
        
        g_lInstID = 0
    
    End If

    Debug.Print "-------------------- End DDE Test ------------------------"

    ' Disable the command buttons and the text boxes.
    For Each oCtl In Controls
    
        If ((TypeOf oCtl Is CommandButton) And (oCtl.Enabled = True)) Or _
           ((TypeOf oCtl Is TextBox) And (oCtl.Enabled = True)) Or _
           ((TypeOf oCtl Is ComboBox) And (oCtl.Enabled = True)) Then
        
            oCtl.Enabled = False
        
        End If
        
    Next

    cmdInitialize.Enabled = True
    
End Sub

Private Sub cmdStartAdv_Click()
    If (CheckData("Advise")) Then
        DDE_StartAdvise
        bAdvise = True
    Else
        MsgBox "Please enter the required data for the transaction."
    End If
End Sub

Private Sub cmdStopAdv_Click()
    If (CheckData("Advise")) Then
        DDE_StopAdvise
        bAdvise = False
    Else
        MsgBox "Please enter the required data for the transaction."
    End If
End Sub

Private Sub cmdExecute_Click()
    
Dim lRet As Long
Dim sValue As String
            
    If (CheckData("Execute")) Then
        ' Load the buffer.
        sValue = txtValue.Text
    
        ' Create the string handles.
        DDE_CreateStringHandles txtService.Text, txtTopic.Text
        
        ' Open the conversation.
        If (g_hDDEConv = 0) Then
            g_hDDEConv = DDE_Connect
        End If
        
        If g_hDDEConv Then
        
            ' Perform the transaction.
            lRet = DdeClientTransaction(sValue, Len(sValue), g_hDDEConv, 0, 0, XTYP_EXECUTE, 2000, 0)
        
            If (lRet) Then
            
                Debug.Print "DDE Execute Success."
        
            Else
            
                Debug.Print "DDE Execute Failure."
                TranslateError
        
            End If
            
        End If
        
        ' Release the memory.
        DDE_FreeStringHandles
    Else
        MsgBox "Please enter the required data for the transaction."
    End If
End Sub

Private Sub cmdRequest_Click()

Dim lRet As Long
Dim lSize As Long
Dim sBuffer As String
Dim sFinal As String
        
    If (CheckData("Request")) Then
        DDE_CreateStringHandles txtService.Text, txtTopic.Text, cboItem.Text
        
        ' Open the conversation.
        If (g_hDDEConv = 0) Then
            g_hDDEConv = DDE_Connect
        End If
        
        If g_hDDEConv Then
        
            ' Perform the transaction.
            lRet = DdeClientTransaction(0, 0, g_hDDEConv, g_hItem, CF_TEXT, XTYP_REQUEST, 2000, 0)
            
            If (lRet) Then
            
                Debug.Print "DDE Request Success."
                
                ' Grab the data from the DDE object create during the transaction. The DDE object
                ' is part of the DDE subsystem memory. Once we get what we want we need to free
                ' the object. Check the Microsoft Platform SDK for more information on freeing
                ' DDE global memory.
                
                ' The first call returns the size of the of the string. For some reason there's
                ' always an extra 3 bytes attached to the end of the string. That's why I have a magic
                ' number.
                lSize = DdeGetData(lRet, vbNullString, 0, 0)
                
                ' Allocate a buffer for the return data.
                sBuffer = String$(lSize, 0)
                            
                ' Grab the data.
                lSize = DdeGetData(lRet, sBuffer, Len(sBuffer), 0)
    
                ' Print the contents of the buffer.
                txtValue.Text = sBuffer
                Debug.Print sBuffer
    
                ' Free the DDE subsystem resources.
                DdeFreeDataHandle lRet
                
            Else
                
                Debug.Print "DDE Request Failed"
                TranslateError
            
            End If
            
        End If
        
        DDE_FreeStringHandles
    Else
        MsgBox "Please enter the required data for the transaction."
    End If
End Sub

Private Sub cmdPoke_Click()
    
Dim lRet As Long
Dim sValue As String
            
    If (CheckData("Poke")) Then
        ' Load the buffer.
        sValue = txtValue.Text
        
        DDE_CreateStringHandles txtService.Text, txtTopic.Text, cboItem.Text
        
        ' Open the conversation.
        If (g_hDDEConv = 0) Then
            g_hDDEConv = DDE_Connect
        End If
        
        If g_hDDEConv Then
        
            ' Perform the transaction.
            lRet = DdeClientTransaction(sValue, Len(sValue), g_hDDEConv, g_hItem, CF_TEXT, XTYP_POKE, 2000, 0)
            
            If (lRet) Then
                
                Debug.Print "DDE Poke Success"
            
            Else
                
                Debug.Print "DDE Poke Failed"
                TranslateError
            
            End If
            
        End If
        
        DDE_FreeStringHandles
    Else
        MsgBox "Please enter the required data for the transaction."
    End If
End Sub

Private Sub cmdClear_Click()
    
    ' Clear out the text boxes.
    cboItem.ListIndex = 0
    txtValue.Text = ""
    
End Sub

Private Sub DDE_CreateStringHandles(ByRef sTheService As String, ByRef sTheTopic As String, Optional ByRef sTheItem As String = "")
    
    ' Create the string handles for the service and topic. DDEML will not
    ' allow you to use standard strings. NOTE: Make sure to release the
    ' string handles once you are done with them.
    g_hService = DdeCreateStringHandle(g_lInstID, sTheService, CP_WINANSI)
    g_hTopic = DdeCreateStringHandle(g_lInstID, sTheTopic, CP_WINANSI)
    
    ' Only convert the item if we were passed a string otherwise you'll get a memory
    ' error.
    If (sTheItem <> "") Then
        
        g_hItem = DdeCreateStringHandle(g_lInstID, cboItem.Text, CP_WINANSI)
    
    End If

End Sub

Private Sub DDE_FreeStringHandles()

    ' Release our string handles.
    If (g_hService <> 0) Then
    
        DdeFreeStringHandle g_lInstID, g_hService
        DdeFreeStringHandle g_lInstID, g_hTopic
    
    End If
    
    If (g_hItem <> 0) Then
    
        DdeFreeStringHandle g_lInstID, g_hItem
    
    End If
    
    g_hService = 0
    g_hTopic = 0
    g_hItem = 0

End Sub

Private Function DDE_Connect() As Long

Dim udtConvCont As CONVCONTEXT
Dim hDDEConv As Long
        
    ' Set up the conversation context structure.
    udtConvCont.iCodePage = CP_WINANSI
    udtConvCont.cb = Len(udtConvCont)
    
    hDDEConv = 0
    
    ' Open the connection to the service.
    hDDEConv = DdeConnect(g_lInstID, g_hService, g_hTopic, udtConvCont)
    
    ' Do we have a connection?
    If hDDEConv Then
        
        Debug.Print "DDE Connection Success."
        
    Else
        
        Debug.Print "DDE Connection Failure."
        TranslateError
    
    End If
    
    DDE_Connect = hDDEConv
    
End Function

Private Sub DDE_Disconnect()
            
    ' Disconnect the DDE conversation.
    If g_hDDEConv Then
        
        If DdeDisconnect(g_hDDEConv) Then
            
            Debug.Print "DDE Disconnect Success."
        
        Else
            
            Debug.Print "DDE Disconnect Failure."
            TranslateError
        
        End If
        
        g_hDDEConv = 0
    
    End If

End Sub

Private Sub DDE_StartAdvise()
    
Dim lRet As Long
Dim lTransVal As Long

    DDE_CreateStringHandles txtService.Text, txtTopic.Text, cboItem.Text
    
    ' Open the conversation.
    If (g_hDDEConv = 0) Then
        g_hDDEConv = DDE_Connect
    End If
    
    If g_hDDEConv Then
    
        ' Perform the transaction.
        lRet = DdeClientTransaction(0, 0, g_hDDEConv, g_hItem, CF_TEXT, XTYP_ADVSTART, 2000, lTransVal)
        
        If (lRet) Then
        
            Debug.Print "DDE Advise Start Success."
                        
            ' Enable the Advise Stop button and disable the Advise Start button.
            cmdStopAdv.Enabled = True
            cmdStartAdv.Enabled = False
            
        Else
        
            Debug.Print "DDE Advise Start Failure."
            
        End If

    End If
    
    DDE_FreeStringHandles
    
End Sub

Private Sub DDE_StopAdvise()

Dim lRet As Long
Dim lTransVal As Long
    
    DDE_CreateStringHandles txtService.Text, txtTopic.Text, cboItem.Text
        
    If g_hDDEConv Then
    
        lRet = DdeClientTransaction(0, 0, g_hDDEConv, g_hItem, CF_TEXT, XTYP_ADVSTOP, 2000, lTransVal)
        
        If (lRet) Then
        
            Debug.Print "DDE Advise Stop Success."
                        
            ' Disable the Advise Stop button.
            cmdStopAdv.Enabled = False
            cmdStartAdv.Enabled = True
            
        Else
        
            Debug.Print "DDE Advise Stop Failure."
            
        End If

    End If
    
    DDE_FreeStringHandles

End Sub

Private Function CheckData(sCommand As String) As Boolean
    
Dim bRet As Boolean

    Select Case sCommand
        Case "Execute"
            If (txtService.Text <> "") And (txtTopic.Text <> "") Then
                bRet = True
            End If
            
        Case "Poke", "Request", "Advise"
            If (txtService.Text <> "") And (txtTopic.Text <> "") And (cboItem.Text <> "<None>") Then
                bRet = True
            End If
            
    End Select

    CheckData = bRet
    
End Function

Private Sub TranslateError()
    
Dim lRet As Long
    
    lRet = DdeGetLastError(g_lInstID)

    Select Case lRet
        Case DMLERR_NO_ERROR
            Debug.Print "DMLERR_NO_ERROR"
            
        Case DMLERR_ADVACKTIMEOUT
            Debug.Print "DMLERR_ADVACKTIMEOUT"
            
        Case DMLERR_BUSY
            Debug.Print "DMLERR_BUSY"
        
        Case DMLERR_DATAACKTIMEOUT
            Debug.Print "DMLERR_DATAACKTIMEOUT"
        
        Case DMLERR_DLL_NOT_INITIALIZED
            Debug.Print "DMLERR_NOT_INITIALIZED"
        
        Case DMLERR_DLL_USAGE
            Debug.Print "DMLERR_USAGE"
        
        Case DMLERR_EXECACKTIMEOUT
            Debug.Print "DMLERR_EXECACKTIMEOUT"
        
        Case DMLERR_INVALIDPARAMETER
            Debug.Print "DMLERR_INVALIDPARAMETER"
        
        Case DMLERR_LOW_MEMORY
            Debug.Print "DMLERR_LOW_MEMORY"
        
        Case DMLERR_MEMORY_ERROR
            Debug.Print "DMLERR_MEMORY_ERROR"
        
        Case DMLERR_NOTPROCESSED
            Debug.Print "DMLERR_NOTPROCESSED"
        
        Case DMLERR_NO_CONV_ESTABLISHED
            Debug.Print "DMLERR_NO_CONV_ESTABLISHED"
        
        Case DMLERR_POKEACKTIMEOUT
            Debug.Print "DMLERR_POKEACKTIMEOUT"
        
        Case DMLERR_POSTMSG_FAILED
            Debug.Print "DMLERR_POSTMSG_FAILED"
        
        Case DMLERR_REENTRANCY
            Debug.Print "DMLERR_REENTRANCY"
        
        Case DMLERR_SERVER_DIED
            Debug.Print "DMLERR_SERVER_DIED"
        
        Case DMLERR_SYS_ERROR
            Debug.Print "DMLERR_SYS_ERROR"
        
        Case DMLERR_UNADVACKTIMEOUT
            Debug.Print "DMLERR_UNADVACKTIMEOUT"
        
        Case DMLERR_UNFOUND_QUEUE_ID
            Debug.Print "DMLERR_UNFOUND_QUEUE_ID"

    End Select
    
End Sub

Private Sub Form_Load()
    txtService.Text = "MyServer"
    txtTopic.Text = "MyTopic"
End Sub
