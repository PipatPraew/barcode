Attribute VB_Name = "leadtype_lot"
Option Explicit

Public Function leadtype(sassypackage As String)
    Dim pkg_pin, Special_mold_feature, Depop_pin As String
    
    pkg_pin = Mid(sassypackage, 1, 8)
    Depop_pin = Mid(sassypackage, 13, 2)
    Special_mold_feature = Mid(sassypackage, 12, 1)
    
        Select Case pkg_pin & Depop_pin
            Case "PDIP-07M03": leadtype = "PDIP 7L PIN3"
            Case "PDIP-07M06": leadtype = "PDIP 7L PIN6"
            Case Else:
        End Select

        Select Case pkg_pin & Special_mold_feature
            Case "INSOP24MN": leadtype = "INSOP NON"
            Case "INSOP24ME": leadtype = "INSOP EXPOSED"
            Case Else:
        End Select

Select Case pkg_pin
            
            Case "ESIP-06M": leadtype = "ESIP-6L"
            Case "ESIPL06M": leadtype = "ESIP-6L L-BEND"
            Case "EESP212M": leadtype = "EESIP-12L"
            Case "EESP312M": leadtype = "EESIP-12L 3ROWS"
            Case "EESP213M": leadtype = "EESIP-13L"
            Case "EESPL13M": leadtype = "EESIP-13L L-BEND"
                
            Case "EDIP-11M": leadtype = "EDIP 11L"
            Case "EDIP-12M": leadtype = "EDIP 12L"
            Case "ESOP-11M": leadtype = "ESOP 11L"
            Case "ESOP-12M": leadtype = "ESOP 12L"
            Case "RESOP12M": leadtype = "RESOP 12L"
            
            Case "PDIP-08M": leadtype = "PDIP-8L"
            Case "PDIPG08M": leadtype = "PDIP-8L(GW)"
                
            Case "SOIC-07M": leadtype = "SOIC-7L"
            Case "SOIC-08M": leadtype = "SOIC-8L"
            Case "SOIC-14M": leadtype = "SOIC-14L"
            Case "SOIC-16M": leadtype = "SOIC-16L"
            Case "SOICW16M": leadtype = "SOIC-16L(W)"
            Case "SOIC-18D": leadtype = "SOIC-18L(D)"
            Case "SOIC-20M": leadtype = "SOIC-20L"
            Case "SOIC-28M": leadtype = "SOIC-28L"
            Case "SOIC-32M": leadtype = "SOIC-32L"
            Case "SOIC-07H": leadtype = "SOIC-7L(H)"
            Case "SOIC-08H": leadtype = "SOIC-8L(H)"
            Case "SOIC-09H": leadtype = "SOIC-9L(H)"
            Case "SOIC-10H": leadtype = "SOIC-10L(H)"
            Case "SOIJ-08M": leadtype = "SOIK 8L"
            
            Case "QSOP-16M": leadtype = "QSOP-16L"
            Case "QSOP-20M": leadtype = "QSOP-20L"
            Case "QSOP-24M": leadtype = "QSOP-24L"
            Case "QSOP-28M": leadtype = "QSOP-28L"
                
            Case "SC70-03M": leadtype = "SC70-3L"
            Case "SC70-05M": leadtype = "SC70-5L"
            Case "SC70-06M": leadtype = "SC70-6L"
            Case "SC70-08M": leadtype = "SC70-8L"
                
            Case "SOT2-03M": leadtype = "SOT-3L"
            Case "SOT1-04M": leadtype = "SOT-4L"
            Case "SOT2-04M": leadtype = "SOT-4L"
            Case "SOT2305M": leadtype = "SOT-5L"
            Case "SOT2306M": leadtype = "SOT-6L"
            Case "SOT2308M": leadtype = "SOT-8L"
                
            Case "TSOT-05M": leadtype = "TSOT-5L"
            Case "TSOT-06M": leadtype = "TSOT-6L"
                
            Case "TSSOP08M": leadtype = "TSSOP-08L"
            Case "TSSOP14M": leadtype = "TSSOP-14L"
            Case "TSSOP16M": leadtype = "TSSOP-16L"
            Case "TSSOP20M": leadtype = "TSSOP-20L"
            Case "TSSOP24M": leadtype = "TSSOP-24L"
            Case "TSSOP28M": leadtype = "TSSOP-28L"
            Case "TSSOP31M": leadtype = "TSSOP-31L"
            Case "TSSOP38M": leadtype = "TSSOP-38L"
            Case "TSSOP48M": leadtype = "TSSOP-48L"
                
            Case "MSOP-08M": leadtype = "MSOP-8L"
            Case "MSOP-10M": leadtype = "MSOP-10L"
            Case "MSOP-12M": leadtype = "MSOP-12L"
            Case "MSOP-16M": leadtype = "MSOP-16L"
            
            Case "SSOP-20D": leadtype = "SSOP-20L"
            Case "TSOC-06M": leadtype = "TSOC-6L"
            
            Case Else:
                        
        End Select
        
        

End Function

