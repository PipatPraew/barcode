VERSION 5.00
Begin VB.Form Main_Process 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   ScaleHeight     =   2445
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   3360
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2880
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   150
      TabIndex        =   1
      Top             =   330
      Width           =   3855
      Begin VB.CommandButton btn_Clear 
         BackColor       =   &H00808080&
         Caption         =   "Clear"
         Height          =   375
         Left            =   2760
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox tb_LotNo 
         BeginProperty Font 
            Name            =   "Leelawadee UI"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Label lb_Status 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Leelawadee UI"
         Size            =   24
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   1800
      Width           =   4215
   End
   Begin VB.Label lb_Lead 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Leelawadee UI"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   1200
      Width           =   4215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Leelawadee UI"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3950
      TabIndex        =   4
      Top             =   0
      Width           =   195
   End
   Begin VB.Label lb_CurrentTool 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "TNF Barcode"
      BeginProperty Font 
         Name            =   "Leelawadee UI"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   150
      TabIndex        =   0
      Top             =   0
      Width           =   1890
   End
End
Attribute VB_Name = "Main_Process"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Declare Sub ReleaseCapture Lib "User32" ()

Private Declare Function InternetGetConnectedState Lib "wininet" (ByRef dwflags As Long, ByVal dwReserved As Long) As Long

Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Private Sub btn_Clear_Click()
    tb_LotNo.Locked = False
    tb_LotNo.Text = ""
    tb_LotNo.SetFocus
    lb_Lead.Caption = ""
    lb_Status.Caption = ""
End Sub

Private Sub Form_Load()
    SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, &H10
    Height = 2445
    Width = 4170
End Sub

Private Sub Label1_Click()
    Me.Hide
End Sub

Private Sub tb_LotNo_Change()
    If Len(tb_LotNo.Text) >= 10 Then
        Timer1.Enabled = True
    End If
End Sub

Private Sub Timer1_Timer()

    If (Timer1.Interval) Then
        Timer1.Enabled = False
        tb_LotNo.Locked = True
        Input_LotNo = UCase(tb_LotNo)
        btn_Clear.SetFocus
        
            If InternetGetConnectedState(0, 0) = 1 Then
               data_lot = get_data_lot(Input_LotNo)
               If loaddata_Fail = True Then
                    lb_leadAlarm ("NOT LOAD DATA")
                    lb_StatusWarning ("CON'T LOAD DATA FORM DATABASE")
                    
                    Exit Sub
               End If
            Else
                lb_leadAlarm ("INTERNET ERROR")
                lb_StatusWarning ("INTERNET NOT CONNECT")
                
                Exit Sub
            End If
        process_function
    End If

End Sub

Private Sub process_function()

    read_data
    
End Sub





































