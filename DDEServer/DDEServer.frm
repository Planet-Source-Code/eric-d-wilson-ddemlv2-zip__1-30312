VERSION 5.00
Begin VB.Form frmDDEServer 
   Caption         =   "DDEML Server"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   4350
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
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
      Left            =   2400
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdvise 
      Caption         =   "Advise"
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
      Left            =   720
      TabIndex        =   3
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Frame fPokeData 
      Caption         =   "Data Window"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   4095
      Begin VB.TextBox txtData 
         Height          =   975
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "This is the DDE server window."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "frmDDEServer"
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
' This application demonstrates how to program a DDE server through the
' DDE Management Library. This is an alternative to the standard form based
' DDE mechanism that is provided with Visual Basic.

' There is very little documentation on using the DDEML through Visual Basic
' so it takes a little trial and error. A perfect example is the function
' declaration for DdeGetData(). The decalration that you get from the API loader
' is incorrect. You have to remove the Alias otherwise it won't work. There
' are some other changes that I had to make like variable types. I've tried
' to add a comment for all the ones I've changed.
'
' This app provides an example DDE server. The server is very basic. It provides
' processing for the major DDE transaction types Execute, Poke, and Request. It also
' provides DDE Advise functionality. This is similar to VB's hot/warm links. When
' data changes on the server the server can automatically update any interested
' clients.
'
' ** When you start the server it will be minimized to the task bar.
'
' Here are some tests you can run.
'
' ************************************************************************
'
' Execute:
'
' With the server running, setup the following information in a DDE client.
'
'   Service:    MyServer   - This is the server name.
'   Topic:      MyTopic    - This is the only defined topic.
'   Value:      This can be one of three commands specified in the server.
'               Acceptable values are:
'
'               MAX     - This will maximize the window.
'               MIN     - This will minimize the window.
'               NORMAL  - This will give the in between setting.
'
' Once you set up these values you can initiate an Execute transaction. The
' server window state should then be adjusted to match the command.
'
' ************************************************************************
'
' Poke:
'
' With the server running, setup the following information in a DDE client.
'
'   Service:    MyServer
'   Topic:      MyTopic
'   Item:       MyPoke          - This is the item value I created for Poke operations.
'   Value:      This is a test. - This can be any string you want.
'
' Now if you initiate a poke operation from your DDE client you should see the
' DDE server window appear (if it was minimized) and the string value that you
' specified should now show up in the "Poke Data" text box.
'
'*************************************************************************
'
' Request:
'
' With the server running setup the following information in a DDE client.
'
'   Service:    MyServer
'   Topic:      MyTopic
'   Item:       MyRequest       - This is the item value I created for Request
'                                 operations.
'
' If your not using my DDEML client you need to make sure you have a mechanism
' set up in your client to print the return value of the request operation.
'
' If you initiate the request operation you should receive the following
' string from the server if the input window is empty. Otherwise you'll
' receive whatever text is in the window:
'
'   "The server input window is empty."
'
'*************************************************************************
'
' Advise:
'
' With the server running setup the following information in a DDE client.
'
'   Service:    MyServer
'   Topic:      MyTopic
'   Item:       MyAdvise
'
' Once the advise loop os started by the client you can enter data into the
' data window, click the "Advise" button and that data should be reflected
' in the clients value window.
'
'*************************************************************************
'
' For further information concerning DDEML check the MSDN documentation
' for information on the various transaction types XTYP_??.
'
'*************************************************************************

Option Explicit

Private hszDDEServer As Long        ' String handle for the server name.
Private hszDDETopic As Long         ' String handle for the topic name.
Private hszDDEItemPoke As Long      ' String handle for the Poke item name.
Private bRunning As Boolean         ' Server running flag.

Private Sub cmdClear_Click()
    ' Clear the text box.
    txtData.Text = ""
End Sub

Private Sub cmdAdvise_Click()
    ' We have to initiate a DDEPostAdvise() in order to let all interested clients
    ' know that something has changed.
    If (DdePostAdvise(lInstID, 0, 0)) Then
        Debug.Print "DdePostAdvise() Success."
    Else
        Debug.Print "DdePostAdvise() Failed."
    End If
End Sub

Private Sub Form_Load()
    
    ' Initialize the DDE subsystem. We need to let the DDEML know what callback
    ' function we intend to use so we pass it address using the AddressOf operator.
    ' If we can't initialize the DDEML subsystem we exit.
    If DdeInitialize(lInstID, AddressOf DDECallback, APPCMD_FILTERINITS, 0) Then
        Exit Sub
    End If
   
    ' Now that the DDEML subsystem is initialized we create string handles for our
    ' server/topic name.
    hszDDEServer = DdeCreateStringHandle(lInstID, DDE_SERVER, CP_WINANSI)
    hszDDETopic = DdeCreateStringHandle(lInstID, DDE_TOPIC, CP_WINANSI)
    hszDDEItemPoke = DdeCreateStringHandle(lInstID, DDE_POKE, CP_WINANSI)
    hszDDEItemAdvise = DdeCreateStringHandle(lInstID, DDE_ADVISE, CP_WINANSI)
    
    ' Lets check to see if another DDE server has already registered with identical
    ' server/topic names. If so we'll exit. If we were to continue the DDE subsystem
    ' could become unstable when a client tried to converse with the server/topic.
    If (DdeConnect(lInstID, hszDDEServer, hszDDETopic, ByVal 0&)) Then
        MsgBox "A DDE server named " & Chr(34) & DDE_SERVER & Chr(34) & " with topic " & _
                Chr(34) & DDE_TOPIC & Chr(34) & " is already running.", vbExclamation, App.Title
        Unload Me
        Exit Sub
    End If
        
    ' We need to register the server with the DDE subsystem.
    If (DdeNameService(lInstID, hszDDEServer, 0, DNS_REGISTER)) Then
        ' Set the server running flag.
        bRunning = True
    End If
    
    Me.WindowState = vbMinimized
End Sub

Private Sub Form_Terminate()
    ' We need to release our string handles.
    DdeFreeStringHandle lInstID, hszDDEServer
    DdeFreeStringHandle lInstID, hszDDETopic
    DdeFreeStringHandle lInstID, hszDDEItemPoke
    DdeFreeStringHandle lInstID, hszDDEItemAdvise
    
    ' Unregister the DDE server.
    If bRunning Then
        DdeNameService lInstID, hszDDEServer, 0, DNS_UNREGISTER
    End If
    
    ' Break down the link with the DDE subsystem.
    DdeUninitialize lInstID
End Sub
