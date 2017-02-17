Attribute VB_Name = "log_file"
Option Explicit

Public Sub LogFile(txtMsg As String)
Dim IntHandle As String
Dim Today, MountYear, M, D, Y As String
    Today = Now
    M = Left(Today, 2)
    D = Mid(Today, 4, 2)
    Y = Mid(Today, 7, 4)
    MountYear = M + D + Y
    
    Dim TimeNow As String
    
    On Error GoTo error
        TimeNow = Left(Now, 2) + Mid(Now, 4, 2) + Mid(Today, 7, 4)
        IntHandle = FreeFile
        Open "C:\Save File\" & TimeNow & ".txt" For Append As IntHandle
        Print #IntHandle, txtMsg
        Close IntHandle
error:
End Sub

Public Function log_() As String
    LogFile (LOT_SCAN_NO & "/" & _
                MASSAGE1 & "/" & _
                MASSAGE2 & "/" & _
                MACHINE_NO_LINE & "/" & _
                COUNTER_TOOLIFE & "/" & _
                ID_TOOL_A & "/" & _
                CHARACTER_A & "/" & _
                ID_TOOL_B & "/" & _
                CHARACTER_B & "/" & _
                ID_TOOL_C & "/" & _
                CHARACTER_C & "/" & _
                TIME_START & "/" & _
                TIME_STOP & "/" & _
                STRIP_QTY & "/" & _
                LOT_PARAMETERS_end & "/" & _
                SASSYPACKAGE_end & "/" & _
                LEADTYPE_end & "/" & _
                FRAME_STOCK_end & "/" & _
                UNIT_QTY_end & "/" & _
                WIP_OPTN_CODE_end & "/" & _
                LOT_STATUS_end & "/" & _
                MACHINE_NO_end & vbCrLf)
End Function


