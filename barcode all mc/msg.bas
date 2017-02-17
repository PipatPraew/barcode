Attribute VB_Name = "msg"
Public Function lb_leadOK(msg As String) As String
    Main_Process.Show
    Main_Process.lb_Lead.FontSize = 23
    Main_Process.lb_Lead.ForeColor = &HFF00&
    Main_Process.lb_Lead.Caption = msg
End Function

Public Function lb_StatusOK(msg As String)
    Main_Process.Show
    Main_Process.lb_Status.FontSize = 22
    Main_Process.lb_Status.ForeColor = &HFF00&
    Main_Process.lb_Status.Caption = msg
End Function

Public Function lb_leadWarning(msg As String)
    Main_Process.Show
    Main_Process.lb_Lead.FontSize = 10
    Main_Process.lb_Lead.ForeColor = &H80000012
    Main_Process.lb_Lead.Caption = msg
End Function

Public Function lb_StatusWarning(msg As String)
    Main_Process.Show
    Main_Process.lb_Status.FontSize = 10
    Main_Process.lb_Status.ForeColor = &H80000012
    Main_Process.lb_Status.Caption = msg
End Function

Public Function lb_leadAlarm(msg As String)
    Main_Process.Show
    Main_Process.lb_Lead.FontSize = 23
    Main_Process.lb_Lead.ForeColor = &HFF&
    Main_Process.lb_Lead.Caption = msg
End Function

Public Function lb_StatusAlarm(msg As String)
    Main_Process.Show
    Main_Process.lb_Status.FontSize = 23
    Main_Process.lb_Status.ForeColor = &HFF&
    Main_Process.lb_Status.Caption = msg
End Function

Public Function lb_CurrentToolOK(msg As String)
    Main_Process.Show
    Main_Process.lb_CurrentTool.ForeColor = &HFF0000
    Main_Process.lb_CurrentTool.Caption = msg
End Function

Public Function lb_CurrentToolAlarm(msg As String)
    Main_Process.Show
    Main_Process.lb_CurrentTool.ForeColor = &HFF&
    Main_Process.lb_CurrentTool.Caption = msg
End Function

Public Sub msgEndLot()
    Main_Process.Show
    Main_Process.tb_LotNo.Locked = False
    Main_Process.tb_LotNo.Text = ""
    Main_Process.tb_LotNo.SetFocus
    Main_Process.lb_Lead.Caption = ""
    Main_Process.lb_Status.FontSize = 23
    Main_Process.lb_Status.ForeColor = &HFF0000
    Main_Process.lb_Status.Caption = "LOT END"
End Sub

Public Sub msgWarning()
    Main_Process.Show
    Main_Process.tb_LotNo.Locked = False
    Main_Process.tb_LotNo.Text = ""
    Main_Process.tb_LotNo.SetFocus
    Main_Process.lb_Lead.FontSize = 18
    Main_Process.lb_Lead.ForeColor = &HFF&
    Main_Process.lb_Lead.Caption = "SCAN BARCODE"
    Main_Process.lb_Status.Caption = ""
End Sub

