Attribute VB_Name = "UserForm1Code"
Option Explicit

'DEFINE ALL THE LBU IP
Const AR As String = "Driver={Microsoft ODBC for Oracle};Server=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.101.2.103)(PORT=1521))(CONNECT_DATA=(SID=ABBPROD)));Uid=PCSAdmin;Pwd=PCSAdmin;"
Const AT As String = "Driver={Microsoft ODBC for Oracle};Server=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=145.241.22.68)(PORT=1521))(CONNECT_DATA=(SID=ABBPROD)));Uid=PCSAdmin;Pwd=PCSAdmin;"
Const AU As String = "Driver={Microsoft ODBC for Oracle};Server=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.128.208.91)(PORT=1521))(CONNECT_DATA=(SID=ABBPROD)));Uid=PCSAdmin;Pwd=PCSAdmin;"
Const BR As String = "Driver={Microsoft ODBC for Oracle};Server=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.100.8.102)(PORT=1521))(CONNECT_DATA=(SID=ABBPROD)));Uid=PCSAdmin;Pwd=PCSAdmin;"
Const CH As String = "Driver={Microsoft ODBC for Oracle};Server=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=138.223.3.101)(PORT=1521))(CONNECT_DATA=(SID=ABBPROD)));Uid=PCSAdmin;Pwd=PCSAdmin;"
Const CN As String = "Driver={Microsoft ODBC for Oracle};Server=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.138.112.195)(PORT=1521))(CONNECT_DATA=(SID=ABBPROD)));Uid=PCSAdmin;Pwd=PCSAdmin;"
Const CNSIT As String = "Driver={Microsoft ODBC for Oracle};Server=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.138.62.170)(PORT=1521))(CONNECT_DATA=(SID=ABBPROD)));Uid=PCSAdmin;Pwd=PCSAdmin;"
Const CZ As String = "Driver={Microsoft ODBC for Oracle};Server=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.42.26.190)(PORT=1521))(CONNECT_DATA=(SID=ABBPROD)));Uid=PCSAdmin;Pwd=PCSAdmin;"
Const DE As String = "Driver={Microsoft ODBC for Oracle};Server=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.51.40.21)(PORT=1521))(CONNECT_DATA=(SID=ABBPROD)));Uid=PCSAdmin;Pwd=PCSAdmin;"
Const DESIT As String = "Driver={Microsoft ODBC for Oracle};Server=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.49.69.62)(PORT=1521))(CONNECT_DATA=(SID=ABBPROD)));Uid=PCSAdmin;Pwd=PCSAdmin;"
Const DU As String = "Driver={Microsoft ODBC for Oracle};Server=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.29.54.37)(PORT=1521))(CONNECT_DATA=(SID=ABBPROD)));Uid=PCSAdmin;Pwd=PCSAdmin;"
Const EE As String = "Driver={Microsoft ODBC for Oracle};Server=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.28.0.41)(PORT=1521))(CONNECT_DATA=(SID=ABBPROD)));Uid=PCSAdmin;Pwd=PCSAdmin;"
Const EG As String = "Driver={Microsoft ODBC for Oracle};Server=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.29.126.50)(PORT=1521))(CONNECT_DATA=(SID=ABBPROD)));Uid=PCSAdmin;Pwd=PCSAdmin;"
Const FI As String = "Driver={Microsoft ODBC for Oracle};Server=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.58.74.154)(PORT=1521))(CONNECT_DATA=(SID=ABBPROD)));Uid=PCSAdmin;Pwd=PCSAdmin;"
Const FR As String = "Driver={Microsoft ODBC for Oracle};Server=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.51.33.15)(PORT=1522))(CONNECT_DATA=(SID=ABBPROD)));Uid=PCSAdmin;Pwd=PCSAdmin;"
Const EE1 As String = "Driver={Microsoft ODBC for Oracle};Server=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.51.40.46)(PORT=1522))(CONNECT_DATA=(SID=ABBPROD)));Uid=PCSAdmin;Pwd=PCSAdmin;"
Const GR As String = "Driver={Microsoft ODBC for Oracle};Server=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.29.0.100)(PORT=1521))(CONNECT_DATA=(SID=ABBPROD)));Uid=PCSAdmin;Pwd=PCSAdmin;"
Const IND As String = "Driver={Microsoft ODBC for Oracle};Server=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.140.47.123)(PORT=1521))(CONNECT_DATA=(SID=ABBPROD)));Uid=PCSAdmin;Pwd=PCSAdmin;"
Const IT As String = "Driver={Microsoft ODBC for Oracle};Server=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.39.14.207)(PORT=1521))(CONNECT_DATA=(SID=ABBPROD)));Uid=PCSAdmin;Pwd=PCSAdmin;"
Const KR As String = "Driver={Microsoft ODBC for Oracle};Server=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.131.2.53)(PORT=1521))(CONNECT_DATA=(SID=ABBPROD)));Uid=PCSAdmin;Pwd=PCSAdmin;"
Const MY As String = "Driver={Microsoft ODBC for Oracle};Server=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.132.6.11)(PORT=1521))(CONNECT_DATA=(SID=ABBPROD)));Uid=PCSAdmin;Pwd=PCSAdmin;"
Const NL As String = "Driver={Microsoft ODBC for Oracle};Server=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.51.40.32)(PORT=1521))(CONNECT_DATA=(SID=ABBPROD)));Uid=PCSAdmin;Pwd=PCSAdmin;"
Const NO As String = "Driver={Microsoft ODBC for Oracle};Server=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.47.61.20)(PORT=1521))(CONNECT_DATA=(SID=ABBPROD)));Uid=PCSAdmin;Pwd=PCSAdmin;"
Const PL As String = "Driver={Microsoft ODBC for Oracle};Server=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.3.106.44)(PORT=1521))(CONNECT_DATA=(SID=ABBPROD)));Uid=PCSAdmin;Pwd=PCSAdmin;"
Const RU As String = "Driver={Microsoft ODBC for Oracle};Server=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.7.42.5)(PORT=1521))(CONNECT_DATA=(SID=ABBPROD)));Uid=PCSAdmin;Pwd=PCSAdmin;"
Const SA As String = "Driver={Microsoft ODBC for Oracle};Server=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.63.197.131)(PORT=1521))(CONNECT_DATA=(SID=ABBPROD)));Uid=PCSAdmin;Pwd=PCSAdmin;"
Const SE As String = "Driver={Microsoft ODBC for Oracle};Server=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=138.227.187.84)(PORT=1521))(CONNECT_DATA=(SID=ABBPROD)));Uid=PCSAdmin;Pwd=PCSAdmin;"
Const SG As String = "Driver={Microsoft ODBC for Oracle};Server=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.134.6.35)(PORT=1521))(CONNECT_DATA=(SID=ABBPROD)));Uid=PCSAdmin;Pwd=PCSAdmin;"
Const TH As String = "Driver={Microsoft ODBC for Oracle};Server=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.135.10.12)(PORT=1521))(CONNECT_DATA=(SID=ABBPROD)));Uid=PCSAdmin;Pwd=PCSAdmin;"
Const TR As String = "Driver={Microsoft ODBC for Oracle};Server=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.29.147.123)(PORT=1521))(CONNECT_DATA=(SID=ABBPROD)));Uid=PCSAdmin;Pwd=PCSAdmin;"
Const ZA As String = "Driver={Microsoft ODBC for Oracle};Server=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.27.188.55)(PORT=1521))(CONNECT_DATA=(SID=ABBPROD)));Uid=PCSAdmin;Pwd=PCSAdmin;"
Const GRPSPG As String = "Driver={Microsoft ODBC for Oracle};Server=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.49.71.83)(PORT=1521))(CONNECT_DATA=(SID=ABBPROD)));Uid=PCSAdmin;Pwd=PCSAdmin;"

'Const GB = "Provider=SQLOLEDB.1;User ID=MEAdmin;Data Source=10.140.233.24;Initial Catalog=MEV2_85_INTTEST"
Const GB = "Provider=SQLOLEDB;Data Source=10.140.233.24;Initial Catalog=MEV2_85_INTTEST;User ID=MEAdmin;Password=welcome1234&;"


'Test instance THABB

'Const THABB As String = "Driver={Microsoft ODBC for Oracle};Server=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.140.15.52)(PORT=1521))(CONNECT_DATA=(SID=ABBPROD)));Uid=THABB;Pwd=THABB;"

'INITIALIZATION OF VARIABLES
Dim SelectCountry As String
Dim SelectTable As String
Dim DBConn As ADODB.Connection
Dim DBData As ADODB.Recordset
Dim DBField As ADODB.Field
Dim LoopCounter As Integer
Dim NumberOfLBU As Integer
Dim LBUName As String
Dim LBUConnectionString As String
Dim SQLQuery As String
Dim WS As Worksheet
Dim WB As Workbook
Dim WBName As String
Dim UpdateCell As String
Dim ErrorLBU As Range
Dim LBUErrorServer As String

Private Sub cboSelectCountry_Change()

End Sub

Private Sub cmdCancel_Click()
'UNLOADING USERFORM
Unload Me
End Sub

Private Sub cmdReset_Click()

Range("E:E").ClearContents
Range("D2").ClearContents
Workbooks("Data Collector.xlsm").Close True
End Sub

Private Sub cmdRun_Click()
    'SET OBJ VALUES FOR DBConn and DBData
    Set DBConn = New ADODB.Connection
    Set DBData = New ADODB.Recordset

    SelectCountry = cboSelectCountry.Value
    SelectTable = cboSelectTable.Value
    ActiveWorkbook.Worksheets("LBU").Activate
    
    NumberOfLBU = Range("A2", Range("A2").End(xlDown)).CountLarge
    
    Select Case SelectCountry

        Case "AR"
            DBConn.ConnectionString = AR
            Call datafetch
            
        Case "AT"
            DBConn.ConnectionString = AT
            Call datafetch
            
        Case "AU"
            DBConn.ConnectionString = AU
        Call datafetch
            
        Case "BR"
            DBConn.ConnectionString = BR
            Call datafetch
            
        Case "CH"
            DBConn.ConnectionString = CH
            Call datafetch
            
        Case "CN"
            DBConn.ConnectionString = CN
            Call datafetch
            
        Case "CNSIT"
            DBConn.ConnectionString = CNSIT
            Call datafetch
            
        Case "CZ"
            DBConn.ConnectionString = CZ
            Call datafetch
        
        Case "DE"
            DBConn.ConnectionString = DE
        Call datafetch
        
        Case "DESIT"
            DBConn.ConnectionString = DESIT
        Call datafetch
        
        Case "DU"
            DBConn.ConnectionString = DU
            Call datafetch
      
        Case "EE"
            DBConn.ConnectionString = EE
            Call datafetch
        
        Case "EG"
            DBConn.ConnectionString = EG
            Call datafetch
            
            
        Case "FI"
            DBConn.ConnectionString = FI
            Call datafetch
        
        Case "FR"
            DBConn.ConnectionString = FR
            Call datafetch
            
        Case "GB"
            DBConn.ConnectionString = GB
        Call datafetch
        
       Case "GR"
            DBConn.ConnectionString = GR
        Call datafetch
            
        Case "IND"
            DBConn.ConnectionString = IND
            Call datafetch
            
        Case "IT"
            DBConn.ConnectionString = IT
            Call datafetch
            
         Case "KR"
            DBConn.ConnectionString = KR
            Call datafetch
            
        Case "MY"
            DBConn.ConnectionString = MY
            Call datafetch
            
        Case "NL"
            DBConn.ConnectionString = NL
            Call datafetch
            
        Case "NO"
            DBConn.ConnectionString = NO
            Call datafetch
            
        Case "PL"
            DBConn.ConnectionString = PL
            Call datafetch
            
        Case "RU"
            DBConn.ConnectionString = RU
            Call datafetch
            
        Case "SA"
            DBConn.ConnectionString = SA
            Call datafetch
            
         Case "SG"
            DBConn.ConnectionString = SG
            Call datafetch
            
        Case "SE"
            DBConn.ConnectionString = SE
        Call datafetch
        
         Case "TH"
            DBConn.ConnectionString = TH
            Call datafetch
            
        Case "TR"
            DBConn.ConnectionString = TR
            Call datafetch
            
        Case "ZA"
            DBConn.ConnectionString = ZA
            Call datafetch
            
        Case "GRPSPG"
            DBConn.ConnectionString = ZA
            Call datafetch

        Case "ALL"
            Call dataall
            
        
    End Select
    
    Unload Me

    End Sub
Sub datafetch()
    'setting reference to variable
        WBName = cboSelectCountry & "_data.csv"
        
        Set WB = Workbooks.Add
        WB.SaveAs "C:\Users\" & Environ("USERNAME") & "\Desktop\Data Collector\" & WBName
        
        On Error GoTo CloseConnection
        DBConn.Open

        DBData.ActiveConnection = DBConn
        
    If Workbooks("Data Collector-SQL.xlsm").Worksheets("LBU").Range("D2").Value = "" Then
        
            If cboSelectTable = "ME_SWITCHGEARS" Then
                
                SQLQuery = "SELECT ID,NAME,CREATED_USER,CREATED_DATE,MODIFIED_USER,MODIFIED_DATE,PROJECT_ID,TYPE,TYPE_SCOPE,IS_SPEC_EXISTS,NEUTRAL_BUSBAR_DIMENSION,BBA_DEPTH,INGRESS_PROTECTION,NEUTRAL_BUS_SIZE,MAIN_BUSBAR_DIMENSION,SWITCHGEAR_ALIGNMENT,TYPE_OF_ARRANGEMENT,BOTTOM_PLATE,RATED_SHORTTIME_WITHSTAND_CAP,DIST_BUSBAR_DIMENSION,MULTIFUNCTIONAL_WALL,EARTHING_SYSTEM,BUSBAR_MATERIAL,MAIN_BUSBAR_SHORT_CCT_CURRENT,MAIN_BUSBAR_RATED_CURRENT,OVERLOAD_PROTECTION,ROOF_PLATE_TYPE,AUX_SUPPLY FROM ME_SWITCHGEARS"
            
            ElseIf cboSelectTable = "ME_PROJECTS" Then
            
                SQLQuery = "SELECT ID,NAME,CREATED_USER,CREATED_DATE,IS_SPEC_EXISTS,TYPE_SCOPE,STATE,LOCKED_BY,PROJECT_TYPE,NUMBER_OF_POLES,AMBIENT_TEMPERATURE,RATED_VOLTAGE,MODIFIED_USER,MODIFIED_DATE,HCC_COUNTRY,PROJECT_MODE,CUSTOMER_PROJ_ID,TYPICAL_NAME_PREFIX FROM ME_PROJECTS"
            
            ElseIf cboSelectTable = "ME_LOCAL_FILES" Then
            
                 SQLQuery = "SELECT LOC_OBJECT_ID,FILENAME,TYPE_SCOPE,CREATED_DATE,CREATED_USER,MODIFIED_DATE,MODIFIED_USER,IS_VISIBLE FROM ME_LOCAL_FILES"
            
            ElseIf cboSelectTable = "ME_LOCAL_PARTS" Then
            
                 SQLQuery = "SELECT ID,NAME,TYPE_SCOPE,TYPE,CREATED_DATE,MODIFIED_DATE,DESCRIPTION,ALTERNATE_PART,COST,BREAKDOWN_LIMIT,SYSTEM FROM ME_LOCAL_PARTS"
                 
            ElseIf cboSelectTable = "ME_PROJECT_REVISION" Then
            
                 SQLQuery = "SELECT PROJECT_ID,DESCRIPTION,CREATED_DATE,CREATED_USER,REVISION,MODIFIED_DATE,MODIFIED_USER,STAGE FROM ME_PROJECT_REVISION"
        
            ElseIf cboSelectTable = "ME_SWITCHGEAR_REVISION" Then
            
                 SQLQuery = "SELECT SWITCHGEAR_ID,DESCRIPTION,CREATED_DATE,CREATED_USER,PROJECT_ID,REVISION,MODIFIED_DATE FROM ME_SWITCHGEAR_REVISION"
            
            ElseIf cboSelectTable = "ME_APPLICATIONS" Then
            
                 SQLQuery = "SELECT ID,NAME,TYPE,CREATED_USER,CREATED_DATE,MODIFIED_DATE,MODIFIED_USER,DESCRIPTION,RELAY_CARD,STATE,TYPE_SCOPE,PROJECT_ID,OWNER,CUSTOMER_NAME FROM ME_APPLICATIONS"
         
            ElseIf cboSelectTable = "ME_LOOKUPS" Then
            
                 SQLQuery = "SELECT ID,NAME,TYPE,CREATED_USER,MODIFIED_USER,CREATED_DATE,MODIFIED_DATE,LOOKUP_DIS,TYPE_SCOPE,DESCRIPTION,IS_DEFAULT FROM ME_LOOKUPS"
            
            ElseIf cboSelectTable = "ME_GLOBAL_PARTS" Then
            
            SQLQuery = "SELECT ID,NAME,TYPE_SCOPE,TYPE,CREATED_DATE,CREATED_USER,MODIFIED_DATE,MODIFIED_USER,DESCRIPTION,REVISION,IS_SPEC_EXISTS,COST_MODIFIED_DATE,URL,LOCAL_PART_TEXT_4,LOCAL_PART_TEXT_3,LOCAL_PART_TEXT_2,LOCAL_PART_TEXT_1,LOCAL_PART_TEXT_5,LOCAL_PART_TEXT_6,LOCAL_PART_TEXT_7,LOCAL_PART_TEXT_8,LOCAL_PART_TEXT_9,LOCAL_PART_TEXT_10,SUPPLIER_CODE,MATERIAL_CODE,ALTERNATE_PART,COST,BREAKDOWN_LIMIT,SYSTEM,BASIC_TEXT2,BASIC_TEXT1,UNIT_OF_MEASURE,PAINTED_ITEM,MDF_CODING,STATE,WEIGHT_KG,MOUNTING_TIME,MANUFACTURER,ABB_DEF_CATEGORY FROM ME_GLOBAL_PARTS"
            
            ElseIf cboSelectTable = "ME_LOCATION_PREFERENCES" Then
            
            SQLQuery = "SELECT ID,PREFERENCE_NAME FROM ME_LOCATION_PREFERENCES"
            
            ElseIf cboSelectTable = "" Then
            MsgBox "Please Select a table name"
            Exit Sub
             
            Else
                SQLQuery = "SELECT * FROM " & cboSelectTable
                
            End If
            Debug.Print SQLQuery
            DBData.Open (SQLQuery)
            GoTo Line1:
            
        Else
            SQLQuery = Workbooks("Data Collector.xlsm").Worksheets("LBU").Range("D2").Value
        End If
        Debug.Print SQLQuery
        On Error GoTo WrongQuery
        DBData.Open (SQLQuery)
            
Line1:    On Error GoTo CloseRecordSet
    Set WS = Workbooks(WBName).Sheets.Add
    WS.Name = cboSelectCountry.Value

    
    For Each DBField In DBData.Fields
        Workbooks(WBName).Activate
        ActiveCell.Value = DBField.Name
        ActiveCell.Offset(0, 1).Select
    Next DBField
    
    ActiveSheet.Range("A1").Select
    ActiveSheet.Range("A2").CopyFromRecordset DBData
    ActiveCell.CurrentRegion.Columns.AutoFit
    ActiveWorkbook.Close True
    

CloseRecordSet:
    DBData.Close
    On Error GoTo 0
    
 DBConn.Close
 
 GoTo Finish:
    
CloseConnection:
    Workbooks("Data Collector-SQL.xlsm").Activate
    MsgBox cboSelectCountry & " Server is Down!"
    Workbooks(WBName).Close
    On Error GoTo 0
    Exit Sub
    
WrongQuery:  MsgBox "Invalid SQL Query!"
ActiveWorkbook.Close
Exit Sub

Finish:  MsgBox "Complete!"
ActiveWorkbook.Close True
End Sub
Sub dataall()
    
    'setting reference to variable
            Set WB = Workbooks.Add
            WB.SaveAs "C:\Users\" & Environ("USERNAME") & "\Desktop\Data Collector\" & "ALL_LBU_DATA.csv"
            For LoopCounter = 0 To NumberOfLBU - 2
            Workbooks("Data Collector.xlsm").Activate
            ActiveWorkbook.Worksheets("LBU").Activate
            LBUName = Workbooks("Data Collector.xlsm").Worksheets("LBU").Range("A2").Offset(LoopCounter, 0).Value
            LBUConnectionString = Workbooks("Data Collector.xlsm").Worksheets("LBU").Range("A2").Offset(LoopCounter, 2).Value
            'DBConn.ConnectionString = "Driver={Microsoft ODBC for Oracle};Server=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.128.208.91)(PORT=1521))(CONNECT_DATA=(SID=ABBPROD)));Uid=PCSAdmin;Pwd=PCSAdmin;"
            DBConn.ConnectionString = LBUConnectionString
        
        On Error GoTo CloseConnection
        DBConn.Open
        
        
        DBData.ActiveConnection = DBConn
        
        If Worksheets("LBU").Range("D2").Value = "" Then
        
               If cboSelectTable = "ME_SWITCHGEARS" Then
               
               SQLQuery = "SELECT ID,NAME,CREATED_USER,CREATED_DATE,MODIFIED_USER,MODIFIED_DATE,PROJECT_ID,TYPE,TYPE_SCOPE,IS_SPEC_EXISTS,NEUTRAL_BUSBAR_DIMENSION,BBA_DEPTH,INGRESS_PROTECTION,NEUTRAL_BUS_SIZE,MAIN_BUSBAR_DIMENSION,SWITCHGEAR_ALIGNMENT,TYPE_OF_ARRANGEMENT,BOTTOM_PLATE,RATED_SHORTTIME_WITHSTAND_CAP,DIST_BUSBAR_DIMENSION,MULTIFUNCTIONAL_WALL,EARTHING_SYSTEM,BUSBAR_MATERIAL,MAIN_BUSBAR_SHORT_CCT_CURRENT,MAIN_BUSBAR_RATED_CURRENT,OVERLOAD_PROTECTION,ROOF_PLATE_TYPE,AUX_SUPPLY FROM ME_SWITCHGEARS"
                   
                ElseIf cboSelectTable = "ME_PROJECTS" Then
                   
                       SQLQuery = "SELECT ID,NAME,CREATED_USER,CREATED_DATE,IS_SPEC_EXISTS,TYPE_SCOPE,STATE,LOCKED_BY,PROJECT_TYPE,NUMBER_OF_POLES,AMBIENT_TEMPERATURE,RATED_VOLTAGE,MODIFIED_USER,MODIFIED_DATE,HCC_COUNTRY,PROJECT_MODE,CUSTOMER_PROJ_ID FROM ME_PROJECTS"
                   
                ElseIf cboSelectTable = "ME_LOCAL_FILES" Then
                   
                        SQLQuery = "SELECT LOC_OBJECT_ID,FILENAME,TYPE_SCOPE,CREATED_DATE,CREATED_USER,MODIFIED_DATE,MODIFIED_USER,IS_VISIBLE FROM ME_LOCAL_FILES"
                     
                ElseIf cboSelectTable = "ME_LOCAL_PARTS" Then
                   
                        SQLQuery = "SELECT ID,NAME,TYPE_SCOPE,TYPE,CREATED_DATE,MODIFIED_DATE,DESCRIPTION,ALTERNATE_PART,COST,BREAKDOWN_LIMIT,SYSTEM FROM ME_LOCAL_PARTS"
                        
                ElseIf cboSelectTable = "ME_PROJECTS" Then
                   
                        SQLQuery = "SELECT PROJECT_ID,DESCRIPTION,CREATED_DATE,CREATED_USER,REVISION,MODIFIED_DATE,MODIFIED_USER,STAGE FROM ME_PROJECT_REVISION"
                
                ElseIf cboSelectTable = "ME_SWITCHGEARS" Then
           
                SQLQuery = "SELECT SWITCHGEAR_ID,DESCRIPTION,CREATED_DATE,CREATED_USER,PROJECT_ID,REVISION,MODIFIED_DATE FROM ME_SWITCHGEAR_REVISION"
           
                ElseIf cboSelectTable = "ME_APPLICATIONS" Then
            
                 SQLQuery = "SELECT ID,NAME,TYPE,CREATED_USER,CREATED_DATE,MODIFIED_DATE,MODIFIED_USER,DESCRIPTION,RELAY_CARD,STATE,TYPE_SCOPE,PROJECT_ID,OWNER,CUSTOMER_NAME FROM ME_APPLICATIONS"
        
                ElseIf cboSelectTable = "ME_LOOKUPS" Then
            
                 SQLQuery = "SELECT ID,NAME,TYPE,CREATED_USER,MODIFIED_USER,CREATED_DATE,MODIFIED_DATE,LOOKUP_DIS,TYPE_SCOPE,DESCRIPTION,IS_DEFAULT FROM ME_LOOKUPS"
                
                ElseIf cboSelectTable = "ME_GLOBAL_PARTS" Then
            
                 SQLQuery = "SELECT ID,NAME,TYPE_SCOPE,TYPE,CREATED_DATE,CREATED_USER,MODIFIED_DATE,MODIFIED_USER,DESCRIPTION,REVISION,IS_SPEC_EXISTS,COST_MODIFIED_DATE,URL,LOCAL_PART_TEXT_4,LOCAL_PART_TEXT_3,LOCAL_PART_TEXT_2,LOCAL_PART_TEXT_1,LOCAL_PART_TEXT_5,LOCAL_PART_TEXT_6,LOCAL_PART_TEXT_7,LOCAL_PART_TEXT_8,LOCAL_PART_TEXT_9,LOCAL_PART_TEXT_10,SUPPLIER_CODE,MATERIAL_CODE,ALTERNATE_PART,COST,BREAKDOWN_LIMIT,SYSTEM,BASIC_TEXT2,BASIC_TEXT1,UNIT_OF_MEASURE,PAINTED_ITEM,MDF_CODING,STATE,WEIGHT_KG,MOUNTING_TIME,MANUFACTURER,ABB_DEF_CATEGORY FROM ME_GLOBAL_PARTS"
            
                ElseIf cboSelectTable = "ME_LOCATION_PREFERENCES" Then
            
                 SQLQuery = "SELECT ID,PREFERENCE_NAME FROM ME_LOCATION_PREFERENCES"
            
                Else
                     SQLQuery = "SELECT * FROM " & cboSelectTable
                     Debug.Print SQLQuery
                End If
                DBData.Open (SQLQuery)
                GoTo Line2:
         Else
            SQLQuery = Workbooks("Data Collector.xlsm").Worksheets("LBU").Range("D2").Value
            End If
            Debug.Print SQLQuery
            
            On Error GoTo WrongQuery
            DBData.Open (SQLQuery)
            
Line2:    On Error GoTo CloseRecordSet
    Set WS = Workbooks("ALL_LBU_DATA.csv").Sheets.Add
    WS.Name = LBUName
    
    For Each DBField In DBData.Fields
        Workbooks("ALL_LBU_DATA.csv").Activate
        ActiveCell.Value = DBField.Name
        ActiveCell.Offset(0, 1).Select
    Next DBField
    
    Workbooks("ALL_LBU_DATA.csv").ActiveSheet.Range("A1").Select
    ActiveSheet.Range("A2").CopyFromRecordset DBData
    ActiveCell.CurrentRegion.Columns.AutoFit
    'ActiveCell.CurrentRegion.ClearFormats

CloseRecordSet:
    DBData.Close
    On Error GoTo 0
    
    DBConn.Close

nextcounter:

Next LoopCounter

Workbooks("ALL_LBU_DATA.csv").Close True

GoTo Complete:

CloseConnection:
    Workbooks("Data Collector.xlsm").Activate
    Set ErrorLBU = Range("A2", Range("A2").End(xlDown))
    ErrorLBU.Select
    LBUErrorServer = ErrorLBU.Find(LBUName).Select
    With ActiveCell.Offset(0, 4)
    
    .Value = "Server Down!"
    .Font.Italic = True
    .Font.Color = rgbRed
    End With
    Resume nextcounter
WrongQuery:  MsgBox "Invalid SQL Query!"
ActiveWorkbook.Close

Exit Sub

Complete:
MsgBox "Complete! Open Data Collector excel to see server status"
ActiveWorkbook.Close True

End Sub
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub frmDataCollect_Click()

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ThisWorkbook.Close SaveChanges:=True
    Application.Visible = True
    Application.Quit
End Sub


