Attribute VB_Name = "variable_All"

Public debug_ As Boolean
'plc
Public netid As Integer
Public nodeid As Integer
Public unitid As Integer
Public fins_error As Integer
Public mc As String
Public qty_plc As Integer
Public qty_tool As Integer
    
Public id_tool_a As String
Public id_tool_b As String
Public id_tool_c As String
    
'lot detail
Public Input_LotNo As String
Public data_lot As String
Public loaddata_Fail As Boolean

'wip
Public lot_no As String
Public pkg_code As String
Public lot_leadtype As String
Public qty_unit As String
Public stock_frame As String
Public wip_optn As String
Public lot_status As String
Public wip_mc As String
