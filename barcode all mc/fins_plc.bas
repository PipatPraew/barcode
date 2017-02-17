Attribute VB_Name = "fins_plc"
'connect plc
Public Function fins_msg(fins_cmd As String) As String
    

    On Error GoTo error
    Main_Button.FinsMsg1.sendFinsCommand 0, nodeid, 0, fins_cmd
    s$ = Main_Button.FinsMsg1.receiveMessage(netid, nodeid, unitid, vbString)
    fins_error = Mid(s$, 5, 4)
    fins_msg = Mid(s$, 9, Len(s$) - 8)
error:
End Function
'read mc can run
Public Function read_run() As Boolean

    Dim ad As String
    If debugg = True Then
        ad = "01018000C8000001" 'CIO 200 -- run
    Else: ad = "0101800262000001" 'CIO 610 -- run
    End If
    
    s$ = Mid(fins_msg(ad), 4, 1)
    If s$ = "1" Then
        read_run = True
    Else: read_run = False
    End If
    
End Function

Public Sub write_run()

    Dim ad As String
    If debugg = True Then
        ad = "01028000C80000010001" 'CIO 200 -- run
    Else: ad = "01028002620000010001" 'w = 01028002620000010001  'CIO 610 -- run
    End If
    
    fins_msg (ad)
    
End Sub

Public Sub write_clear_run()

    Dim ad As String
    If debugg = True Then
        ad = "01028000C80000010000" 'CIO 200 -- run
    Else: ad = "01028002620000010000" 'w = 01028002620000010000  'CIO 610 -- run
    End If
    
    fins_msg (ad)
    
End Sub

Public Function read_end_lot() As Boolean

    Dim ad As String
    If debugg = True Then
        ad = "01018000C9000001" 'CIO 201 -- lot end
    Else: ad = "0101800263000001" 'lot end r = 0101800263000001  'CIO 611 -- lot end
    End If
    
    s$ = Mid(fins_msg(ad), 4, 1)
    If s$ = "1" Then
        read_end_lot = True
    Else: read_end_lot = False
    End If
    
End Function

Public Sub write_clear_endlot()

    Dim ad As String
    If debugg = True Then
        ad = "01028000C90000010000"  'CIO 201 -- lot end
    Else: ad = "01028002630000010000" 'w = 01028002630000010000  'CIO 611 -- lot end
    End If
    
    fins_msg (ad)
    
End Sub

Public Function read_btn_run() As Boolean

    Dim ad As String
    If debugg = True Then
        ad = "01018000CA000001"  'CIO 202 -- btn run
    Else: ad = "0101800264000001" 'btn run r = 0101800264000001  'CIO 612 -- btn run
    End If
    
    s$ = Mid(fins_msg(ad), 4, 1)
    If s$ = "1" Then
        read_btn_run = True
    Else: read_btn_run = False
    End If
    
End Function

Public Sub write_clear_btnrun()

    Dim ad As String
    If debugg = True Then
        ad = "01028000CA0000010000" 'CIO 202 -- lot end
    Else: ad = "01028002640000010000" 'w = 01028002640000010000  'CIO 612 -- btn run
    End If
    
    fins_msg (ad)
    
End Sub

'read mc number
Public Function mc_number(ByRef qty_plc As Integer) As String

    nodeid = 1

    Dim ad As String
    If debug_ = True Then
        ad = "01018000CB000001" 'CIO 203 -- mc no
    Else: ad = "0101800265000001" 'mc no = 0101800266000001  'CIO 613 -- mc no
    End If
    s$ = fins_msg(ad)
    
    If Mid(s$, 1, 1) = "0" Then
        s1$ = "DEJ-"
    End If
    If Mid(s$, 1, 1) = "1" Then
        s1$ = "TNF-"
    End If
    
    mc_number = s1$ + Mid(s$, 3, 2)
    
    Select Case mc_number
        Case "TNF-51", "TNF-55", "TNF-56", "TNF-58", "TNF-62", "TNF-65"
            qty_plc = 1
        Case "TNF-38", "TNF-42", "TNF-45", "TNF-54", "TNF-57"
            qty_plc = 2
        Case "DEJ-20", "DEJ-26", "DEJ-27", "DEJ-29", "DEJ-32", "DEJ-33", "TNF-49", "TNF-52", "TNF-53"
            qty_plc = 3
    End Select
    
End Function

Public Function read_id_tool(ByRef id_a, id_b, id_c As String) As String
    Dim fins_cmd As String

    Dim i, i1 As Integer
    
    Dim sa1 As String
    Dim s1(1 To 2) As String
    
    Dim sa2 As String
    Dim s2(1 To 4) As String
                                            
    Dim sa3(1 To 2) As String
    Dim s3(1 To 4) As String
    Dim sid(1 To 2) As String
                        
    Dim sa4 As String
    Dim s4(1 To 6) As String
                        
    mc = mc_number(qty_plc)
    
    Select Case mc
            '1 tool node 2
            Case "DEJ-20", "DEJ-26", "DEJ-27", "DEJ-29", "DEJ-32", "DEJ-33", "TNF-38", "TNF-42", "TNF-45", "TNF-54", "TNF-57"
                If debug_ = True Then
                    fins_cmd = "0101820000000002" 'R D0-->2W
                Else
                    fins_cmd = "010182138B000002" 'R D5003-->2W
                End If
                    nodeid = 2
                    qty_tool = 1
                        sa1 = fins_msg(fins_cmd)
                            For i = 1 To 2
                                If i <> 1 Then
                                    s1(i) = Mid(sa1, (i * 4) - 3, 4)
                                Else
                                    s1(i) = Mid(sa1, 1, 4)
                                End If
                            Next i
                        id_a = s1(2) & s1(1)
                        id_b = Empty
                        id_c = Empty
                    
            '2 tool node 1
            Case "TNF-51", "TNF-55", "TNF-58"
                If debug_ = True Then
                    fins_cmd = "0101820000000004" 'R D0-->4W
                Else
                    fins_cmd = "010182138B000004" 'R D5003-->4W
                End If
                    nodeid = 1
                    qty_tool = 1
                        
                        sa2 = fins_msg(fins_cmd)
                            For i = 1 To 4
                                If i <> 1 Then
                                    s2(i) = Mid(sa2, (i * 4) - 3, 4)
                                Else
                                    s2(i) = Mid(sa2, 1, 4)
                                End If
                            Next i
                        
                        id_a = s2(2) & s2(1)
                        id_b = s2(4) & s2(3)
                        id_c = Empty
                    
            '2 tool node 2,3
            Case "TNF-49", "TNF-52", "TNF-53"
                If debug_ = True Then
                    fins_cmd = "0101820000000002" 'R D0-->2W
                Else
                    fins_cmd = "010182138B000004" 'R D5003-->2W
                End If
                
                    qty_tool = 2
                
                    For i = 1 To 2
                        nodeid = i + 1
                        s3(i) = fins_msg(fins_cmd)
                    
                            For i1 = 1 To 2
                                    If i1 <> 1 Then
                                        s1(i1) = Mid(s3(i), (i1 * 4) - 3, 4)
                                    Else
                                        s1(i1) = Mid(s3(i), 1, 4)
                                    End If
                            Next i1
                        sid(i) = s1(2) & s1(1)
                    
                    Next i
                    
                    id_a = sid(1)
                    id_b = sid(2)
                    id_c = Empty
                    
              '3 tool node 1
            Case "TNF-56", "TNF-62", "TNF-65"
                If debug_ = True Then
                    fins_cmd = "0101820000000006" 'R D0-->6W
                Else
                    fins_cmd = "010182138B000006" 'R D5003-->6W
                End If
                
                nodeid = 1
                    qty_tool = 3
                        
                    sa4 = fins_msg(fins_cmd)
                        For i = 1 To 6
                            If i <> 1 Then
                                s4(i) = Mid(sa4, (i * 4) - 3, 4)
                            Else
                                s4(i) = Mid(sa4, 1, 4)
                            End If
                        Next i
                        
                    id_a = s4(2) & s4(1)
                    id_b = s4(4) & s4(3)
                    id_c = s4(6) & s4(5)
                               
    End Select

End Function

Public Function read_ch_tool(ByRef ch_a, ch_b, ch_c As String) As String
    Dim fins_cmd As String

    Dim i, i1 As Integer
    
    Dim sa1 As String
    Dim s1(1 To 2) As String
    
    Dim sa2 As String
    Dim s2(1 To 4) As String
                                            
    Dim sa3(1 To 2) As String
    Dim s3(1 To 4) As String
    Dim sid(1 To 2) As String
                        
    Dim sa4 As String
    Dim s4(1 To 6) As String
                        
    mc = mc_number(qty_plc)
    
    Select Case mc
            '1 tool node 2
            Case "DEJ-20", "DEJ-26", "DEJ-27", "DEJ-29", "DEJ-32", "DEJ-33", "TNF-38", "TNF-42", "TNF-45", "TNF-54", "TNF-57"
                If debug_ = True Then
                    fins_cmd = "0101820000000002" 'R D0-->2W
                Else
                    fins_cmd = "010182138B000002" 'R D5003-->2W
                End If
                    nodeid = 2
                    qty_tool = 1
                        sa1 = fins_msg(fins_cmd)
                            For i = 1 To 2
                                If i <> 1 Then
                                    s1(i) = Mid(sa1, (i * 4) - 3, 4)
                                Else
                                    s1(i) = Mid(sa1, 1, 4)
                                End If
                            Next i
                        ch_a = s1(2) & s1(1)
                        ch_b = Empty
                        ch_c = Empty
                    
            '2 tool node 1
            Case "TNF-51", "TNF-55", "TNF-58"
                If debug_ = True Then
                    fins_cmd = "0101820000000004" 'R D0-->4W
                Else
                    fins_cmd = "010182138B000004" 'R D5003-->4W
                End If
                    nodeid = 1
                    qty_tool = 1
                        
                        sa2 = fins_msg(fins_cmd)
                            For i = 1 To 4
                                If i <> 1 Then
                                    s2(i) = Mid(sa2, (i * 4) - 3, 4)
                                Else
                                    s2(i) = Mid(sa2, 1, 4)
                                End If
                            Next i
                        
                        ch_a = s2(2) & s2(1)
                        ch_b = s2(4) & s2(3)
                        ch_c = Empty
                    
            '2 tool node 2,3
            Case "TNF-49", "TNF-52", "TNF-53"
                If debug_ = True Then
                    fins_cmd = "0101820000000002" 'R D0-->2W
                Else
                    fins_cmd = "010182138B000004" 'R D5003-->2W
                End If
                
                    qty_tool = 2
                
                    For i = 1 To 2
                        nodeid = i + 1
                        s3(i) = fins_msg(fins_cmd)
                    
                            For i1 = 1 To 2
                                    If i1 <> 1 Then
                                        s1(i1) = Mid(s3(i), (i1 * 4) - 3, 4)
                                    Else
                                        s1(i1) = Mid(s3(i), 1, 4)
                                    End If
                            Next i1
                        sid(i) = s1(2) & s1(1)
                    
                    Next i
                    
                    ch_a = sid(1)
                    ch_b = sid(2)
                    ch_c = Empty
                    
              '3 tool node 1
            Case "TNF-56", "TNF-62", "TNF-65"
                If debug_ = True Then
                    fins_cmd = "0101820000000006" 'R D0-->6W
                Else
                    fins_cmd = "010182138B000006" 'R D5003-->6W
                End If
                
                nodeid = 1
                    qty_tool = 3
                        
                    sa4 = fins_msg(fins_cmd)
                        For i = 1 To 6
                            If i <> 1 Then
                                s4(i) = Mid(sa4, (i * 4) - 3, 4)
                            Else
                                s4(i) = Mid(sa4, 1, 4)
                            End If
                        Next i
                        
                    ch_a = s4(2) & s4(1)
                    ch_b = s4(4) & s4(3)
                    ch_c = s4(6) & s4(5)
                               
    End Select

End Function


































