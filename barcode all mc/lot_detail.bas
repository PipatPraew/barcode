Attribute VB_Name = "lot_detail"
Option Explicit

Public Function get_data_lot(Input_Lot_no As String)
    get_data_lot = ""
    
    Dim objHttp As Object, sQuery As String, textResponse As String
    
On Error GoTo error
    Dim i As Integer
    
    For i = 1 To 3
        sQuery = "http://utlwebprd1/OEEWebAPI/DataForOEE/LotInfo/GetLotInfo?&logID=1231562&enNumber=41226383&token=OEE_WEB_API&busItem=[{'LotID':'" + Input_Lot_no + "'}]"
        Set objHttp = CreateObject("Microsoft.XMLHTTP")
        objHttp.Open "GET", sQuery, False
        objHttp.send
        get_data_lot = objHttp.ResponseText
        Set objHttp = Nothing
        
        If Mid(get_data_lot, Val(Len(get_data_lot)), 1) = "]" Then
            loaddata_Fail = False
            Exit For
        Else
            loaddata_Fail = True
        End If
    Next i
    
        If (loaddata_Fail = True) Then
            'LogFile (Now & "  !Load data not complete")
        End If
    
error:
End Function

Public Function lot_parameters(data_lot As String)
    lot_parameters = ""
    Dim x, r As String
    Dim i, e, Y, z As Long
    
        x = "LOT PARAMETERS"
        r = ","
        
        i = InStr(1, data_lot, x)
        z = (i + Len(x) + 11)
        e = InStr(z, data_lot, r)
        Y = e - z - 1
        
        If i <> Empty Then
            lot_parameters = Mid(data_lot, z, Y)
        End If
        
End Function

Public Function unit_qty(data_lot As String)
    unit_qty = ""
    Dim x, r As String
    Dim i, e, Y, z As Long
        
        x = "QTY"
        r = ","
        
        i = InStr(1, data_lot, x)
        z = (i + Len(x) + 11)
        e = InStr(z, data_lot, r)
        Y = e - z - 1
        
        If i <> Empty Then
            unit_qty = Mid(data_lot, z, Y)
        End If
        
End Function

Public Function lot_status(data_lot As String)
    lot_status = ""
    Dim x, r As String
    Dim i, e, Y, z As Long
    
        x = "LOT_STATUS"
        r = ","
        
        i = InStr(1, data_lot, x)
        z = (i + Len(x) + 11)
        e = InStr(z, data_lot, r)
        Y = e - z - 1
        
        If i <> Empty Then
            lot_status = Mid(data_lot, z, Y)
        End If
    
End Function

Public Function wip_optn_code(data_lot As String)
    wip_optn_code = ""
    Dim x, r As String
    Dim i, e, Y, z As Long
    
        x = "STAGE_UTL_WIP_OPTN_CODE"
        r = ","
        
        i = InStr(1, data_lot, x)
        z = (i + Len(x) + 11)
        e = InStr(z, data_lot, r)
        Y = e - z - 1
        
        If i <> Empty Then
            wip_optn_code = Mid(data_lot, z, Y)
        End If
    
End Function

Public Function sassypackage(data_lot As String)
    sassypackage = ""
    Dim x, r As String
    Dim i, e, Y, z As Long
    
        x = "SASSYPACKAGE"
        r = ","
        
        i = InStr(1, data_lot, x)
        z = (i + Len(x) + 11)
        e = InStr(z, data_lot, r)
        Y = e - z - 1
        
        If i <> Empty Then
            sassypackage = Mid(data_lot, z, Y)
        End If
    
End Function

Public Function frame_stock(data_lot As String)
    frame_stock = ""
    Dim x, r As String
    Dim i, e, Y, z As Long
    
        x = "OPTFIELD5"
        r = ","
        
        i = InStr(1, data_lot, x)
        z = (i + Len(x) + 11)
        e = InStr(z, data_lot, r)
        Y = e - z - 1
        
        If i <> Empty Then
            frame_stock = Mid(data_lot, z, Y)
        End If
    
End Function

Public Function machine_no(data_lot As String)
    machine_no = ""
    Dim x, r As String
    Dim i, e, Y, z As Long
    
        x = "MACHINE_NO"
        r = ","
        
        i = InStr(1, data_lot, x)
        z = (i + Len(x) + 11)
        e = InStr(z, data_lot, r)
        Y = e - z - 1
        
        If i <> Empty Then
            machine_no = Mid(data_lot, z, Y)
        End If
    
End Function



