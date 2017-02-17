Attribute VB_Name = "leadtype_mc"
Public Function mc_lead()
    
    Dim id_a, id_b, id_c As String
    
    read_id_tool id_a, id_b, id_c
     
    Select Case id_a
        'tnf-51,58
        Case "00001252": tool_a = "PDIP 7L PIN3"
        Case "00001253": tool_a = "PDIP 7L PIN6"
        Case "00001411": tool_a = "PDIP 7L PIN3"
        Case "00001412": tool_a = "PDIP 7L PIN6"
        'tnf-55
        Case "00001297": tool_a = "ESIP-6L"
        Case "00001339": tool_a = "ESIP-6L L-BEND"
        Case "10000500": tool_a = "EESIP-12L"
        Case "00001336": tool_a = "EESIP-12L 3ROWS"
        Case "00001395": tool_a = "EESIP-13L"
        Case "10000459": tool_a = "EESIP-13L L-BEND"
        'tnf-56
        Case "00001383": tool_a = "ESOP 11L"
        Case "00001386": tool_a = "EDIP 11L"
        Case "10000536": tool_a = "RESOP 12L"
        'TNF-62
        Case "10000484": tool_a = "ESOP 11L"
        Case "10000487": tool_a = "EDIP 11L"
        Case "00001652": tool_a = "RESOP 12L"
        'TNF-65
        Case "00001676": tool_a = "INSOP EXPOSED"
        Case "00001677": tool_a = "INSOP NON"
        Case Else
            tool_a = "ID error"
    End Select

    Select Case id_b
        'tnf-51,58
        Case "00001250": tool_b = tool_a
        Case "00001413": tool_b = tool_a
        'tnf-55
        Case "00001298": tool_b = "ESIP-6L"
        Case "00001340": tool_b = "ESIP-6L L-BEND"
        Case "10000499": tool_b = "EESIP-12L"
        Case "00001337": tool_b = "EESIP-12L 3ROWS"
        Case "00001396": tool_b = "EESIP-13L"
        Case "10000458": tool_b = "EESIP-13L L-BEND"
        'tnf-56
        Case "00001384": tool_b = "ESOP 11L"
        Case "10000501": tool_b = "EDIP 11L"
        Case "10000534": tool_b = "RESOP 12L"
        'tnf-62
        Case "10000485": tool_b = "ESOP 11L"
        Case "10000488": tool_b = "EDIP 11L"
        Case "00001648": tool_b = "RESOP 12L"
        'tnf-65
        Case "00001670": tool_b = "INSOP EXPOSED"
        Case "00001672": tool_b = "INSOP NON"
        
        Case Else
            tool_b = "ID error"
    End Select
        
    Select Case id_c
        'TNF-56
        Case "00001385": tool_c = "ESOP 11L"
        Case "00001388": tool_c = "EDIP 11L"
        Case "10000537": tool_c = "RESOP 12L"
        'TNF-62
        Case "10000486": tool_c = "ESOP 11L"
        Case "10000489": tool_c = "EDIP 11L"
        Case "00001647": tool_c = "RESOP 12L"
        'TNF-65
        Case "00001671": tool_c = "INSOP EXPOSED"
        Case "00001673": tool_c = "INSOP NON"
        Case Else
            tool_c = "ID error"
    End Select

    Dim IDerror As Boolean
        IDerror = True
    Dim x(1 To 3) As String
        x(1) = tool_a
        x(2) = tool_b
        x(3) = tool_c
        
            For i = 1 To qty_tool
                If ("ID error" = x(i)) Then
                    mc_lead = "ID error"
                    Exit Function
                Else: IDerror = False
                End If
            Next i

        If (IDerror = False) Then
        
            Select Case qty_tool
                Case 1: mc_lead = tool_a
                Case 2
                    If (tool_a = tool_b) Then
                        mc_lead = tool_a
                    Else: mc_lead = "Wrong Tool"
                    End If
                Case 3
                    If (tool_a = tool_b) Then
                        If (tool_b = tool_c) Then
                            mc_lead = tool_a
                        Else: mc_lead = "Wrong Tool"
                        End If
                    Else: mc_lead = "Wrong Tool"
                    End If
             End Select
        End If

End Function
