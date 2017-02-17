VERSION 5.00
Object = "{A2DAE4C8-A2C9-11D0-A36A-00C04FC302F2}#1.0#0"; "FinsMsgCtl.ocx"
Begin VB.Form Main_Button 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1245
   LinkTopic       =   "Form1"
   ScaleHeight     =   735
   ScaleWidth      =   1245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin FINSMSGCTLLib.FinsMsg FinsMsg1 
      Left            =   240
      Top             =   120
      _Version        =   65536
      _ExtentX        =   953
      _ExtentY        =   688
      _StockProps     =   2
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   600
      Top             =   -120
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "New Lot."
      BeginProperty Font 
         Name            =   "Leelawadee UI"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   370
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1000
   End
End
Attribute VB_Name = "Main_Button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Declare Sub ReleaseCapture Lib "User32" ()

Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Private Sub Command1_Click()
Dim i As Integer

Dim a, b, c As String
'read_id_tool a, b, c
mc_lead
'MsgBox (mc_number(i) & " " & i)


    'Main_Process.Show
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim lngReturnValue As Long

    If Button = 1 Then
        Call ReleaseCapture
        lngReturnValue = SendMessage(Main_Button.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If
End Sub

Private Sub Form_Load()

    debug_ = True         '''''''' DEBUG ''''''''

    SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, &H10
    Height = 735
    Width = 1245
    
End Sub

Private Sub Timer1_Timer1()

    If (Timer1.Interval) Then
       
        Else
           
    End If

End Sub
