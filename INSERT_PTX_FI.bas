Attribute VB_Name = "Module1"
'20110215 change main query
'20110215 change replace "&" เป็น " AND "
'20110215 add function PTX_FI_BACK_DT
'201110308 recheck error and add description error in main function
'110316 Ton : ปรับปรุง/แก้ไขโปรแกรมให้ไม่ค้าง และ write log และ Text File ได้ถูกต้อง
'20110422 yos add IP_NAME  ใน function ที่ ตกหล่นไป

Option Explicit
'Dim Conn As New ADODB.Connection
Dim Conn As Object
Dim count_insert, count_delete, tim_before, time_after As String
Public Sub Main()
    '110316 Ton : ปรับปรุง/แก้ไขโปรแกรมให้ไม่ค้าง และ write log และ Text File ได้ถูกต้อง
    On Error GoTo ErrHandler
    
    'Dim rs As New ADODB.Recordset
    Dim rs As Object
    Dim sql, sSqlCount, time_before, sSqlInsertLog, sSqlDelete, sSqlInsert, sSqlBackdate, subErr, sErr As String
    Dim sPTX_FI_DAILY, sPTX_FI_COMMON, sPTX_FI_22, sPTX_FI_21, sPTX_FI_1, sPTX_FI_BACK_DT As String
    Dim sDate As String
    Dim sMsg As String
    
    Set rs = CreateObject("ADODB.Connection")
    
    'Clear Output ไฟล์เดิม
    writeTextFile "", True
    
    'Connect Database
    If Not ConnectDB(App.Path & "\Config.ini") Then
        sMsg = "FAIL" & vbNewLine & "ไม่สามารถติดต่อฐานข้อมูลได้"
        writeTextFile sMsg, False
        Exit Sub
    End If
    
    If Command <> "" Then
        'ตรวจสอบวันที่ ที่ใส่เข้ามา
        sDate = Command
        'รับวันที่ในรูปแบบ YYYYMMDD เท่านั้น (Operation ส่ง - หรือ / ให้ไม่ได้)
        sDate = Mid(sDate, 1, 4) & "-" & Mid(sDate, 5, 2) & "-" & Mid(sDate, 7, 2)
        
        If Not IsDate(sDate) Then
            sMsg = "FAIL" & vbNewLine & "ระบุค่าวันที่ไม่ถูกต้อง ('" & sDate & "')"
            writeTextFile sMsg, False
            Exit Sub
        End If
    Else
        sDate = Format(Date, "YYYY") & "-" & Format(Date, "MM") & "-" & Format(Date, "DD")
    End If
    
    'ส่วนของการ gen record
    'Delete ก่อน re-run
    'sDate = "2011-03-14"
    sSqlCount = "select count(*),char(current timestamp) from sysibm.sysdummy1;"
    Set rs = Conn.Execute(sSqlCount)
    time_before = rs(1)
    Set rs = Nothing
    
    sSqlCount = " SELECT COUNT(*) AS COUNT_ROW  FROM DS_PTX WHERE DATA_SET_DATE ='" & sDate & "' AND DATA_SYSTEM_CD ='HIPO' AND FLAG_ON_OFF = '0'  WITH UR; "
    Set rs = Conn.Execute(sSqlCount)
    count_delete = rs("COUNT_ROW")
    Set rs = Nothing
    
    On Error Resume Next
    Conn.BeginTrans
    '***** เริ่ม Process Delete ข้อมูล *****
    sSqlInsertLog = " INSERT INTO DS_PTX_LOG  ( ARR_TYPE_CD, ARR_NUMBER, TRAN_DTE, TRAN_TYPE_CD, ITEM_TYPE_CD, TRAN_SEQ, ACTION_TIMESTAMP, ACTION_FLAG, REC_PAY_ITEM_DSC, DATA_SET_DATE, PROVIDER_BR_NO, IBF_IND_CD, PAY_METHOD_CD, CURR_CD, TRAN_AMT_CURR, DEBT_ARR_TYPE_CD, ISIN_CODE, DEBT_INS_NAME, ISSUER_NAME, COUNTRY_CD_ISS, ISSUE_DTE, MATURITY_DTE, ORG_TRM, ORG_TRM_U, COUPON_RATE, INT_COUNTRY_CD, UNIT_OF_TRAN, SELL_SEC_AMT, PURCHASE_DTE, OS_AMT, LU_DTE, LU_TIME, USER_LU, DATA_SYSTEM_CD, FLAG_ON_OFF, PART_ID, IP_NAME, IP_COUNTRY_CD ) "
    sSqlInsertLog = sSqlInsertLog & " SELECT ARR_TYPE_CD, ARR_NUMBER, TRAN_DTE, TRAN_TYPE_CD, ITEM_TYPE_CD, TRAN_SEQ, CURRENT TIMESTAMP,'D', REC_PAY_ITEM_DSC, DATA_SET_DATE, PROVIDER_BR_NO, IBF_IND_CD, PAY_METHOD_CD, CURR_CD, TRAN_AMT_CURR, DEBT_ARR_TYPE_CD, ISIN_CODE, DEBT_INS_NAME, ISSUER_NAME, COUNTRY_CD_ISS, ISSUE_DTE, MATURITY_DTE, ORG_TRM, ORG_TRM_U, COUPON_RATE, INT_COUNTRY_CD, UNIT_OF_TRAN, SELL_SEC_AMT, PURCHASE_DTE, OS_AMT, LU_DTE, LU_TIME, USER_LU, DATA_SYSTEM_CD, FLAG_ON_OFF, PART_ID, IP_NAME, IP_COUNTRY_CD"
    sSqlInsertLog = sSqlInsertLog & "  From DS_PTX WHERE DATA_SET_DATE ='" & sDate & "' AND DATA_SYSTEM_CD ='HIPO' AND FLAG_ON_OFF = '0' "
    Conn.Execute (sSqlInsertLog)
    
    sSqlDelete = " DELETE FROM DS_PTX WHERE DATA_SET_DATE ='" & sDate & "' AND DATA_SYSTEM_CD ='HIPO' AND FLAG_ON_OFF = '0' "
    Conn.Execute (sSqlDelete)
    '***** สิ้นสุด Process Delete ข้อมูล *****
    If Err.Number = 0 Then
        Conn.CommitTrans
    Else
        Conn.RollbackTrans
        sSqlCount = "select count(*),char(current timestamp) from sysibm.sysdummy1;"
        Set rs = Conn.Execute(sSqlCount)
        time_after = rs(1)
        Set rs = Nothing
        
        sSqlInsert = "INSERT INTO BATCH_RUN_LOG  VALUES ('" & sDate & "', 'INSERT_PTX_FI.EXE', 0, 0, 0,'" & time_before & "', '" & time_after & "', 'FAIL','Delete Rerun error'); "
        Conn.Execute (sSqlInsert)
        
        sErr = "FAIL" & vbNewLine & "Error " & Err.Number & " : " & Err.Description
        writeTextFile sErr, False
        Conn.Close
        Exit Sub
    End If
    '***** เริ่ม Process insert ข้อมูล *****
    count_insert = 0
    sErr = ""
    sPTX_FI_1 = PTX_FI_1(sDate)
    sPTX_FI_21 = PTX_FI_21(sDate)
    sPTX_FI_22 = PTX_FI_22(sDate)
    sPTX_FI_COMMON = PTX_FI_COMMON(sDate)
    sPTX_FI_BACK_DT = PTX_FI_BACK_DT(sDate)
    sPTX_FI_DAILY = PTX_FI_DAILY(sDate)
    
    If sPTX_FI_1 = "" And sPTX_FI_21 = "" And sPTX_FI_21 = "" And sPTX_FI_22 = "" And sPTX_FI_COMMON = "" And sPTX_FI_BACK_DT = "" And sPTX_FI_DAILY = "" Then
        sSqlCount = "select count(*),char(current timestamp) from sysibm.sysdummy1;"
        Set rs = Conn.Execute(sSqlCount)
        time_after = rs(1)
        Set rs = Nothing
        
        sSqlInsert = " INSERT INTO BATCH_RUN_LOG  VALUES ('" & sDate & "', 'INSERT_PTX_FI.EXE', " & count_delete & ", " & count_insert & ", 0,'" & time_before & "', '" & time_after & "', 'Complete','Insert Product FI Complete'); "
        Conn.Execute (sSqlInsert)
    Else
        'sErr = " Insert DS_PTX FAIL :  " & Err.Description & sPTX_FI_1 & sPTX_FI_21 & sPTX_FI_22 & sPTX_FI_COMMON & sPTX_FI_DAILY
        '201110308 recheck error and add description error in main function
        sErr = "Insert DS_PTX FAIL in (  " & vbNewLine
        subErr = ""
        If sPTX_FI_1 <> "" Then
            subErr = subErr & "  function PTX_FI_1 :  " & sPTX_FI_1 & vbNewLine
        End If
        If sPTX_FI_21 <> "" Then
            subErr = subErr & "  function PTX_FI_21 :  " & sPTX_FI_21 & vbNewLine
        End If
        If sPTX_FI_22 <> "" Then
            subErr = subErr & "  function PTX_FI_22 :  " & sPTX_FI_22 & vbNewLine
        End If
        If sPTX_FI_COMMON <> "" Then
            subErr = subErr & "  function PTX_FI_COMMON :  " & sPTX_FI_COMMON & vbNewLine
        End If
        If sPTX_FI_BACK_DT <> "" Then
            subErr = subErr & "  function PTX_FI_BACK_DT :  " & sPTX_FI_BACK_DT & vbNewLine
        End If
        If sPTX_FI_DAILY <> "" Then
            subErr = subErr & "  function PTX_FI_DAILY :  " & sPTX_FI_DAILY & vbNewLine
        End If
        sErr = sErr & subErr & " )"
        
        sSqlCount = "select count(*),char(current timestamp) from sysibm.sysdummy1;"
        Set rs = Conn.Execute(sSqlCount)
        time_after = rs(1)
        Set rs = Nothing

        sSqlInsert = "INSERT INTO BATCH_RUN_LOG  VALUES ('" & sDate & "', 'INSERT_PTX_FI.EXE', " & count_delete & ", " & count_insert & ", 0,'" & time_before & "', '" & time_after & "', 'FAIL','" & repQuote(Left(sErr, 500)) & "'); "
        Conn.Execute (sSqlInsert)
    End If
    '***** สิ้นสุด Process insert ข้อมูล *****
    
ErrHandler:
    If Err.Number <> 0 Then
        sMsg = "FAIL" & vbNewLine & "Error " & Err.Number & " : " & Err.Description
        writeTextFile sMsg, False
    Else
        If Trim(sErr) = "" Then
            writeTextFile "Complete", True
        Else
            writeTextFile "FAIL" & vbNewLine & sErr, True
        End If
    End If
    Conn.Close
    Set Conn = Nothing
End Sub
Private Function PTX_FI_1(iDate) As String
      ' sheet 3-PTX-FI(1)
      '110316 Ton : เพิ่ม Process การแปลงข้อมูลที่อาจมีค่า '
      On Error GoTo ErrDB
      Dim sSqlSEQ, sSqlSelect, flag_insert, sSqlSub, sHeadSqlInsert, sHeadSqlInsertLog, sSqlInsert, sSqlInsertLog As String
      Dim IP_NAME, DEBT_ARR_TYPE_CD, ISIN_CODE, DEBT_INS_NAME, ISSUER_NAME, COUNTRY_CD_ISS, ISSUE_DTE As String
      Dim MATURITY_DTE, sORG_TRM, sSplit, ORG_TRM, ORG_TRM_U, COUPON_RATE, INT_COUNTRY_CD, IP_COUNTRY_CD As String
      Dim dateInsert, endTxt As String
      Dim TRAN_SEQ As Integer
      
      Conn.BeginTrans
      
      'Dim Conn As New ADODB.Connection
      'Dim rs, rsSub1, rsSub2 As New ADODB.Recordset
      Dim rs, rsSub1, rsSub2 As Object
      Set rs = CreateObject("ADODB.Connection")
      Set rsSub1 = CreateObject("ADODB.Connection")
      Set rsSub2 = CreateObject("ADODB.Connection")
      Dim sDate As String

      sDate = iDate
      '-------------------------------------------------------------------- Find SEQ-----------------------------------------------------------------
            sSqlSEQ = " SELECT COALESCE(MAX(TRAN_SEQ), 0)  AS TRAN_SEQ "
            sSqlSEQ = sSqlSEQ & " FROM DS_PTX WHERE DATA_SET_DATE =  '" & sDate & "' AND DATA_SYSTEM_CD = 'HIPO'    WITH UR; "
            
            Set rs = Conn.Execute(sSqlSEQ)
           If Not rs.EOF Then
                  TRAN_SEQ = Trim(rs("TRAN_SEQ")) + 1
            Else
                  TRAN_SEQ = 1
            End If
            Set rs = Nothing

      '-------------------------------------------------------------------- Select Record -----------------------------------------------------------------
      sSqlSelect = " SELECT * FROM TMS_FI_WAV  "
      'sSqlSelect = sSqlSelect & "  WHERE  TRANS_DATE ='" & sDate & "' "
      sSqlSelect = sSqlSelect & "  WHERE  SETTLE_DATE ='" & sDate & "' "
      sSqlSelect = sSqlSelect & "  AND  TYPE ='T'    WITH UR; "
      Set rs = Conn.Execute(sSqlSelect)
      Do While Not rs.EOF
            flag_insert = ""
            IP_NAME = ""
            '------------------------------------------------------------------
            DEBT_ARR_TYPE_CD = ""
            ISIN_CODE = ""
            DEBT_INS_NAME = ""
            ISSUER_NAME = ""
            COUNTRY_CD_ISS = ""
            ISSUE_DTE = "null"
            MATURITY_DTE = "null"
            sORG_TRM = ""
            sSplit = ""
            ORG_TRM = "null"
            ORG_TRM_U = ""
            COUPON_RATE = "null"
            INT_COUNTRY_CD = ""
            '------------------------------------------------------------------
            '---------Select Group 1 ISIN_CODE,DEBT_INS_NAME,ISSUE_DTE,MATURITY_DTE,COUPON_RATE
                              sSqlSub = "  SELECT  SEC_FEATURE.ISIN_CODE AS ISIN_CODE ,  SEC_FEATURE.INST_DESC AS DEBT_INS_NAME, SEC_FEATURE.ISSUE_DATE AS ISSUE_DTE  "
                              sSqlSub = sSqlSub & "    ,SEC_FEATURE.MAT_DATE AS MATURITY_DTE ,SEC_FEATURE.COUPON_RATE AS COUPON_RATE, "
                              sSqlSub = sSqlSub & "     MONTH(SEC_FEATURE.MAT_DATE -  SEC_FEATURE.ISSUE_DATE ) AS DIFF_MONTH ,YEAR(SEC_FEATURE.MAT_DATE -  SEC_FEATURE.ISSUE_DATE ) AS DIFF_YEAR  "
                              sSqlSub = sSqlSub & "    FROM ESL_SEC_FEATURE AS SEC_FEATURE  WHERE SEC_FEATURE.SECURITY_ID ='" & Trim(rs("BOND_ID")) & "'   WITH UR;   "
                               
                                                            
                              Set rsSub1 = Conn.Execute(sSqlSub)
                              If Not rsSub1.EOF Then
                                          'MATURITY_DTE , ISSUE_DTE
                                          sORG_TRM = Cal_ORG_TRM(Trim(rsSub1("MATURITY_DTE")), Trim(rsSub1("ISSUE_DTE")))
                                          sSplit = Split(sORG_TRM, ";")
                                          ORG_TRM = sSplit(0)
                                          ORG_TRM_U = sSplit(1)
                                          'ISIN_CODE = Trim(rsSub1("ISIN_CODE"))
                                          If Not IsNull(rsSub1("ISIN_CODE")) Then
                                                ISIN_CODE = Replace(Trim(rsSub1("ISIN_CODE")), "&", " AND ")
                                          End If
                                          'DEBT_INS_NAME = Trim(rsSub1("DEBT_INS_NAME"))
                                          If Not IsNull(rsSub1("DEBT_INS_NAME")) Then
                                                DEBT_INS_NAME = Replace(Trim(rsSub1("DEBT_INS_NAME")), "&", " AND ")
                                          End If
                                          If IsDate(Trim(rsSub1("ISSUE_DTE"))) = True Then
                                                ISSUE_DTE = "'" & Format(Trim(rsSub1("ISSUE_DTE")), "YYYY") & "-" & Format(Trim(rsSub1("ISSUE_DTE")), "MM") & "-" & Format(Trim(rsSub1("ISSUE_DTE")), "DD") & "'"
                                          End If
                                          If IsDate(Trim(rsSub1("MATURITY_DTE"))) = True Then
                                                MATURITY_DTE = "'" & Format(Trim(rsSub1("MATURITY_DTE")), "YYYY") & "-" & Format(Trim(rsSub1("MATURITY_DTE")), "MM") & "-" & Format(Trim(rsSub1("MATURITY_DTE")), "DD") & "'"
                                          End If
                                          COUPON_RATE = Trim(rsSub1("COUPON_RATE"))
'                                         If Trim(rsSub1("DIFF_YEAR")) >= "1" Then
'                                                      ORG_TRM = Trim(rsSub1("DIFF_YEAR"))
'                                                      ORG_TRM_U = "Y"
'                                          Else
'                                                      ORG_TRM = Trim(rsSub1("DIFF_MONTH"))
'                                                      ORG_TRM_U = "M"
'                                          End If
                              End If
                              Set rsSub1 = Nothing
                              
            '---------Select Group 2 ISSUER_NAME
                              sSqlSub = "  SELECT SYENTITY.LONG_NAME  AS ISSUER_NAME  FROM ESL_SEC_FEATURE AS SEC_FEATURE "
                              sSqlSub = sSqlSub & "  INNER JOIN ESL_SYENTITY AS SYENTITY       ON     SEC_FEATURE.ENTITY_CODE = SYENTITY.ENTITY_CODE "
                              sSqlSub = sSqlSub & "   WHERE SEC_FEATURE.SECURITY_ID ='" & Trim(rs("BOND_ID")) & "'   WITH UR;  "
                              Set rsSub1 = Conn.Execute(sSqlSub)
                              If Not rsSub1.EOF And Not IsNull(rsSub1("ISSUER_NAME")) Then
                                    'ISSUER_NAME = Trim(rsSub1("ISSUER_NAME"))
                                    ISSUER_NAME = Replace(Trim(rsSub1("ISSUER_NAME")), "&", " AND ")
                              End If
                              Set rsSub1 = Nothing
                              
            '---------Select Group 3 COUNTRY_CD_ISS
                  sSqlSub = "   SELECT MAP_CODE.MAP_CD2 AS COUNTRY_CD_ISS     FROM ESL_SEC_FEATURE AS SEC_FEATURE"
                  sSqlSub = sSqlSub & "   INNER JOIN ESL_SYENTITY AS SYENTITY       ON     SEC_FEATURE.ENTITY_CODE = SYENTITY.ENTITY_CODE"
                  sSqlSub = sSqlSub & "   INNER JOIN MAP_CODE AS MAP_CODE       ON     SYENTITY.DOMI_CODE =  MAP_CODE.MAP_CD1       AND    MAP_CODE.MAP_TABLE_CD = 'MAP029' "
                  sSqlSub = sSqlSub & "   WHERE SEC_FEATURE.SECURITY_ID ='" & Trim(rs("BOND_ID")) & "'   WITH UR;  "
                              Set rsSub1 = Conn.Execute(sSqlSub)
                              If Not rsSub1.EOF Then
                                    COUNTRY_CD_ISS = Trim(rsSub1("COUNTRY_CD_ISS"))
                              End If
                              Set rsSub1 = Nothing
                              
            '---------Select Group 4 INT_COUNTRY_CD
                  sSqlSub = "   SELECT MAP_CODE3.MAP_CD2 AS INT_COUNTRY_CD  FROM ESL_SEC_FEATURE AS SEC_FEATURE"
                  sSqlSub = sSqlSub & "   INNER JOIN MAP_CODE AS MAP_CODE3      ON     SEC_FEATURE.CURRENCY =  MAP_CODE3.MAP_CD1"
                  sSqlSub = sSqlSub & "   AND    MAP_CODE3.MAP_TABLE_CD = 'MAP030'"
                  sSqlSub = sSqlSub & "   WHERE SEC_FEATURE.SECURITY_ID ='" & Trim(rs("BOND_ID")) & "'   WITH UR;  "
                              
                              Set rsSub1 = Conn.Execute(sSqlSub)
                              If Not rsSub1.EOF Then
                                    INT_COUNTRY_CD = Trim(rsSub1("INT_COUNTRY_CD"))
                              End If
                              Set rsSub1 = Nothing
                              
            '---------Select Group 5 DEBT_ARR_TYPE_CD
                  sSqlSub = "  SELECT MAP_CODE1.MAP_CD2 AS DEBT_ARR_TYPE_CD  FROM ESL_SEC_FEATURE AS SEC_FEATURE"
                  sSqlSub = sSqlSub & "   INNER JOIN MAP_CODE AS MAP_CODE1   ON     SEC_FEATURE.BOND_TYPE = MAP_CODE1.MAP_CD1      AND    MAP_CODE1.MAP_TABLE_CD = 'MAP028'"
                  sSqlSub = sSqlSub & "   WHERE SEC_FEATURE.SECURITY_ID ='" & Trim(rs("BOND_ID")) & "'   WITH UR;  "
                              Set rsSub1 = Conn.Execute(sSqlSub)
                              If Not rsSub1.EOF Then
                                    DEBT_ARR_TYPE_CD = Trim(rsSub1("DEBT_ARR_TYPE_CD"))
                              End If
                              Set rsSub1 = Nothing


            'DATE
            'dateInsert = Format(Trim(rs("TRANS_DATE")), "YYYY") & "-" & Format(Trim(rs("TRANS_DATE")), "MM") & "-" & Format(Trim(rs("TRANS_DATE")), "DD")
            dateInsert = Format(Trim(rs("SETTLE_DATE")), "YYYY") & "-" & Format(Trim(rs("SETTLE_DATE")), "MM") & "-" & Format(Trim(rs("SETTLE_DATE")), "DD")

            'TRAN_SEQ = TRAN_SEQ + 1
            sHeadSqlInsert = " INSERT INTO DS_PTX ( ARR_TYPE_CD, ARR_NUMBER, TRAN_DTE, TRAN_TYPE_CD, ITEM_TYPE_CD, TRAN_SEQ, REC_PAY_ITEM_DSC, DATA_SET_DATE, PROVIDER_BR_NO, IBF_IND_CD, PAY_METHOD_CD, CURR_CD, TRAN_AMT_CURR, DEBT_ARR_TYPE_CD, ISIN_CODE, DEBT_INS_NAME, ISSUER_NAME, COUNTRY_CD_ISS, ISSUE_DTE, MATURITY_DTE, ORG_TRM, ORG_TRM_U, COUPON_RATE, INT_COUNTRY_CD, UNIT_OF_TRAN, SELL_SEC_AMT, PURCHASE_DTE, OS_AMT, LU_DTE, LU_TIME, USER_LU, DATA_SYSTEM_CD, FLAG_ON_OFF, PART_ID, IP_NAME, IP_COUNTRY_CD) "
            sHeadSqlInsertLog = " INSERT INTO DS_PTX_LOG ( ARR_TYPE_CD, ARR_NUMBER, TRAN_DTE, TRAN_TYPE_CD, ITEM_TYPE_CD, TRAN_SEQ, ACTION_TIMESTAMP, ACTION_FLAG, REC_PAY_ITEM_DSC, DATA_SET_DATE, PROVIDER_BR_NO, IBF_IND_CD, PAY_METHOD_CD, CURR_CD, TRAN_AMT_CURR, DEBT_ARR_TYPE_CD, ISIN_CODE, DEBT_INS_NAME, ISSUER_NAME, COUNTRY_CD_ISS, ISSUE_DTE, MATURITY_DTE, ORG_TRM, ORG_TRM_U, COUPON_RATE, INT_COUNTRY_CD, UNIT_OF_TRAN, SELL_SEC_AMT, PURCHASE_DTE, OS_AMT, LU_DTE, LU_TIME, USER_LU, DATA_SYSTEM_CD, FLAG_ON_OFF, PART_ID, IP_NAME, IP_COUNTRY_CD) "
            
'                                    sSqlSub = " SELECT  MAP_CODE1.MAP_CD2 AS DEBT_ARR_TYPE_CD , SEC_FEATURE.ISIN_CODE AS ISIN_CODE ,  SEC_FEATURE.INST_TYPE AS DEBT_INS_NAME ,  SYENTITY.ENTITY_20CODE AS ISSUER_NAME , MAP_CODE2.MAP_CD2 AS COUNTRY_CD_ISS , SEC_FEATURE.ISSUE_DATE AS ISSUE_DTE ,SEC_FEATURE.MAT_DATE AS MATURITY_DTE ,SEC_FEATURE.COUPON_RATE AS COUPON_RATE , MAP_CODE3.MAP_CD2 AS INT_COUNTRY_CD ,  "
'                                    sSqlSub = sSqlSub & " SEC_FEATURE.*, SYENTITY.* ,MAP_CODE1.*,MAP_CODE2.*,MAP_CODE3.*  , MONTH(SEC_FEATURE.MAT_DATE -  SEC_FEATURE.ISSUE_DATE ) AS DIFF_MONTH ,YEAR(SEC_FEATURE.MAT_DATE -  SEC_FEATURE.ISSUE_DATE ) AS DIFF_YEAR  FROM ESL_SEC_FEATURE AS SEC_FEATURE "
'                                    sSqlSub = sSqlSub & " INNER JOIN ESL_SYENTITY AS SYENTITY ON SEC_FEATURE.ENTITY_CODE = SYENTITY.ENTITY_CODE"
'                                    sSqlSub = sSqlSub & " INNER JOIN MAP_CODE AS MAP_CODE1 ON SEC_FEATURE.BOND_TYPE = MAP_CODE1.MAP_CD1 AND MAP_CODE1.MAP_TABLE_CD = 'MAP028'"
'                                    sSqlSub = sSqlSub & " INNER JOIN MAP_CODE AS MAP_CODE2 ON SYENTITY.DOMI_CODE =  MAP_CODE2.MAP_CD1 AND MAP_CODE2.MAP_TABLE_CD = 'MAP029'"
'                                    sSqlSub = sSqlSub & " INNER JOIN MAP_CODE AS MAP_CODE3 ON SEC_FEATURE.CURRENCY =  MAP_CODE3.MAP_CD1 AND MAP_CODE3.MAP_TABLE_CD = 'MAP030'"
'                                    sSqlSub = sSqlSub & " WHERE SEC_FEATURE.SECURITY_ID ='" & Trim(rs("BOND_ID")) & "'   WITH UR;  "
'
'                                    Set rsSub1 = Conn.Execute(sSqlSub)
'
'                                    If Not rsSub1.EOF Then
'                                                DEBT_ARR_TYPE_CD = Trim(rsSub1("DEBT_ARR_TYPE_CD"))
'                                                ISIN_CODE = Trim(rsSub1("ISIN_CODE"))
'                                                DEBT_INS_NAME = Trim(rsSub1("DEBT_INS_NAME"))
'                                                ISSUER_NAME = Trim(rsSub1("ISSUER_NAME"))
'                                                COUNTRY_CD_ISS = Trim(rsSub1("COUNTRY_CD_ISS"))
'                                                If Trim(rsSub1("ISSUE_DTE")) <> "" Then
'                                                      ISSUE_DTE = "'" & Trim(rsSub1("ISSUE_DTE")) & "'"
'                                                End If
'                                                If Trim(rsSub1("MATURITY_DTE")) <> "" Then
'                                                      MATURITY_DTE = "'" & Trim(rsSub1("MATURITY_DTE")) & "'"
'                                                End If
'                                                 If Trim(rsSub1("DIFF_YEAR")) >= "1" Then
'                                                            ORG_TRM = Trim(rsSub1("DIFF_YEAR"))
'                                                            ORG_TRM_U = "Y"
'                                                Else
'                                                            ORG_TRM = Trim(rsSub1("DIFF_MONTH"))
'                                                            ORG_TRM_U = "M"
'                                                End If
'                                                COUPON_RATE = Trim(rsSub1("COUPON_RATE"))
'                                                INT_COUNTRY_CD = Trim(rsSub1("INT_COUNTRY_CD"))
'                                    End If
                              
                                    If Trim(rs("TRANS_NUM")) <> "" Then
                                    
                                          sSqlSub = " SELECT * FROM ESL_COMMON AS COMMON "
                                          sSqlSub = sSqlSub & " INNER JOIN  ESL_SYENTITY AS SYENTITY ON COMMON.CNTR_PARTY = SYENTITY.ENTITY_CODE"
                                          sSqlSub = sSqlSub & "  INNER JOIN MAP_CODE AS MAP_CODE ON SYENTITY.DOMI_CODE =  MAP_CODE.MAP_CD1 AND MAP_CODE.MAP_TABLE_CD = 'MAP029' "
                                          'sSqlSub = sSqlSub & " WHERE COMMON.TRANS_NUM ='" & Trim(rs("TRANS_NUM")) & "'   WITH UR; "
                                          'SORN
                                         sSqlSub = sSqlSub & " WHERE COMMON.TRANS_NUM ='" & Trim(rs("TRANS_NUM")) & "' and   COMMON.AS_OF_DT =  '" & sDate & "'   WITH UR; "
                                          
                                          Set rsSub2 = Conn.Execute(sSqlSub)
                                          If Not rsSub2.EOF Then
                                                     If Not IsNull(Trim(rsSub2("REMARK"))) Then
                                                            endTxt = InStr(1, Trim(rsSub2("REMARK")), "/CF")
                                                             If endTxt <> "0" Then
                                                                   IP_NAME = Mid(Trim(rsSub2("REMARK")), 1, endTxt)
                                                                   IP_NAME = Replace(IP_NAME, "&", " AND ")
'
'                                                             Else
'                                                                   IP_NAME = Trim(rsSub2("ENTITY_20CODE"))
'                                                                   IP_NAME = Replace(IP_NAME, "&", " AND ")
                                                             End If
                                                      End If
                                                      'sorn
                                                      IP_COUNTRY_CD = Trim(rsSub2("MAP_CD2"))
                                                      '20110422 yos add IP_NAME
                                                       If Not IsNull(Trim(rsSub2("ENTITY_20CODE"))) Then
                                                            IP_NAME = Trim(rsSub2("ENTITY_20CODE"))
                                                            IP_NAME = Replace(IP_NAME, "&", " AND ")
                                                       End If
                                                      
                                          Else
                                                IP_NAME = ""
                                                'sorn
                                                IP_COUNTRY_CD = ""
                                          End If
                                                                        
                                    Else
                                          sSqlSub = " SELECT * FROM ESL_SEC_FEATURE AS SEC_FEATURE"
                                          sSqlSub = sSqlSub & " INNER JOIN  ESL_SYENTITY AS SYENTITY ON SEC_FEATURE.ENTITY_CODE = SYENTITY.ENTITY_CODE"
                                          sSqlSub = sSqlSub & "  INNER JOIN MAP_CODE AS MAP_CODE ON SYENTITY.DOMI_CODE =  MAP_CODE.MAP_CD1 AND MAP_CODE.MAP_TABLE_CD = 'MAP029' "
                                          sSqlSub = sSqlSub & " WHERE SEC_FEATURE.SECURITY_ID ='" & Trim(rs("BOND_ID")) & "'   WITH UR; "
                                          
                                          Set rsSub2 = Conn.Execute(sSqlSub)
                                          If Not rsSub2.EOF Then
                                                If Not IsNull(rsSub2("ENTITY_20CODE")) Then
                                                      IP_NAME = Trim(rsSub2("ENTITY_20CODE"))
                                                      IP_NAME = Replace(IP_NAME, "&", " AND ")
                                                Else
                                                      IP_NAME = ""
                                                End If
                                                IP_COUNTRY_CD = Trim(rsSub2("MAP_CD2"))
                                          Else
                                                IP_NAME = ""
                                                IP_COUNTRY_CD = ""
                                          End If
                                          
                                    End If
    
            '110316 Ton : เพิ่ม Process การแปลงข้อมูลที่อาจมีค่า '
            ISIN_CODE = repQuote(ISIN_CODE)
            DEBT_INS_NAME = repQuote(DEBT_INS_NAME)
            ISSUER_NAME = repQuote(ISSUER_NAME)
            IP_NAME = repQuote(IP_NAME)
            '*****************************************************
            
            If Trim(rs("CHG_WAV")) <> "0" Then
                  If Trim(rs("CHG_WAV")) < "0" Then
                        'CHG_WAV < 0
                        
                        sSqlInsert = sHeadSqlInsert & "  VALUES('999999' , '" & Trim(rs("TRANS_NUM")) & "' , '" & dateInsert & "' , '270003' , '268027' ,  " & TRAN_SEQ & "   "
                        sSqlInsert = sSqlInsert & "  , '' , '" & dateInsert & "' , '' , '' , '234005' , '" & Trim(rs("CURRENCY")) & "' ,  " & Abs(rs("CHG_WAV")) & " "
                        sSqlInsert = sSqlInsert & "  , '" & DEBT_ARR_TYPE_CD & "' , '" & ISIN_CODE & "' , '" & DEBT_INS_NAME & "' , '" & ISSUER_NAME & "'  "
                        sSqlInsert = sSqlInsert & "  , '" & COUNTRY_CD_ISS & "'  , " & ISSUE_DTE & " , " & MATURITY_DTE & "  , " & ORG_TRM & "  , '" & ORG_TRM_U & "' "
                        sSqlInsert = sSqlInsert & "  ,  " & COUPON_RATE & "  ,'" & INT_COUNTRY_CD & "' , NULL , 0 , NULL , 0  , CURRENT DATE , CURRENT TIME "
                        sSqlInsert = sSqlInsert & "  , 'PLDMS'  , 'HIPO'  , '0' , '' , '" & IP_NAME & "' , '" & IP_COUNTRY_CD & "');  "
                        
                        
                        
                         sSqlInsertLog = sHeadSqlInsertLog & " VALUES( '999999' , '" & Trim(rs("TRANS_NUM")) & "' , '" & dateInsert & "' , '270003' , '268027' ,  " & TRAN_SEQ & "  , CURRENT TIMESTAMP, 'I'  "
                         sSqlInsertLog = sSqlInsertLog & "  , '' , '" & dateInsert & "' , '' , '' , '234005' , '" & Trim(rs("CURRENCY")) & "' ,  " & Abs(rs("CHG_WAV")) & "  "
                         sSqlInsertLog = sSqlInsertLog & "  , '" & DEBT_ARR_TYPE_CD & "' , '" & ISIN_CODE & "' , '" & DEBT_INS_NAME & "' , '" & ISSUER_NAME & "'  "
                         sSqlInsertLog = sSqlInsertLog & "  , '" & COUNTRY_CD_ISS & "'  , " & ISSUE_DTE & " , " & MATURITY_DTE & "  , " & ORG_TRM & "  , '" & ORG_TRM_U & "' "
                         sSqlInsertLog = sSqlInsertLog & "  ,  " & COUPON_RATE & "  ,'" & INT_COUNTRY_CD & "' , NULL , 0 , NULL , 0  , CURRENT DATE , CURRENT TIME "
                         sSqlInsertLog = sSqlInsertLog & "  , 'PLDMS'  , 'HIPO'  , '0' , '' , '" & IP_NAME & "' , '" & IP_COUNTRY_CD & "');  "

                        Conn.Execute (sSqlInsert)
                        Conn.Execute (sSqlInsertLog)

                        
                  Else
                        'CHG_WAV > 0
                        
                        sSqlInsert = sHeadSqlInsert & " VALUES ('999999', '" & Trim(rs("TRANS_NUM")) & "', '" & dateInsert & "', '270003', '268022', " & TRAN_SEQ & " "
                        sSqlInsert = sSqlInsert & " , '', '" & dateInsert & "', '', '', '234005', '" & Trim(rs("CURRENCY")) & "', " & Abs(rs("CHG_WAV")) & ""
                        sSqlInsert = sSqlInsert & " , '" & DEBT_ARR_TYPE_CD & "', '" & ISIN_CODE & "', '" & DEBT_INS_NAME & "', '" & ISSUER_NAME & "' "
                        sSqlInsert = sSqlInsert & " , '" & COUNTRY_CD_ISS & "' , " & ISSUE_DTE & ", " & MATURITY_DTE & " "
                        sSqlInsert = sSqlInsert & " ," & ORG_TRM & ", '" & ORG_TRM_U & "', " & COUPON_RATE & " ,'" & INT_COUNTRY_CD & "', NULL , 0, NULL "
                        sSqlInsert = sSqlInsert & " , 0 , current date, current time, 'PLDMS' , 'HIPO' , '0', '',  '" & IP_NAME & "' , '" & IP_COUNTRY_CD & "'); "
                        
                        sSqlInsertLog = sHeadSqlInsertLog & " VALUES ('999999', '" & Trim(rs("TRANS_NUM")) & "', '" & dateInsert & "', '270003', '268022', " & TRAN_SEQ & " , CURRENT TIMESTAMP, 'I'  "
                        sSqlInsertLog = sSqlInsertLog & " , '', '" & dateInsert & "', '', '', '234005', '" & Trim(rs("CURRENCY")) & "', " & Abs(rs("CHG_WAV")) & ""
                        sSqlInsertLog = sSqlInsertLog & " , '" & DEBT_ARR_TYPE_CD & "', '" & ISIN_CODE & "', '" & DEBT_INS_NAME & "', '" & ISSUER_NAME & "' "
                        sSqlInsertLog = sSqlInsert & " , '" & COUNTRY_CD_ISS & "' , " & ISSUE_DTE & ", " & MATURITY_DTE & " "
                        sSqlInsertLog = sSqlInsertLog & " ," & ORG_TRM & ", '" & ORG_TRM_U & "', " & COUPON_RATE & " ,'" & INT_COUNTRY_CD & "', NULL, 0, NULL  "
                        sSqlInsertLog = sSqlInsertLog & " , 0 , current date, current time, 'PLDMS' , 'HIPO' , '0', '',  '" & IP_NAME & "' , '" & IP_COUNTRY_CD & "'); "
                        
                        Conn.Execute (sSqlInsert)
                        Conn.Execute (sSqlInsertLog)
                  End If
                  TRAN_SEQ = TRAN_SEQ + 1
                  count_insert = count_insert + 1
            End If
            
            If Trim(rs("PROFIT")) <> "0" Then
                        'PROFIT
                        
                        sSqlInsert = sHeadSqlInsert & " VALUES ('999999', '" & Trim(rs("TRANS_NUM")) & "', '" & dateInsert & "', '270001', '268017', " & TRAN_SEQ & "  "
                        sSqlInsert = sSqlInsert & " , 'GAIN', '" & dateInsert & "', '', '', '', '" & Trim(rs("CURRENCY")) & "', " & Abs(rs("PROFIT")) & ""
                        sSqlInsert = sSqlInsert & " ,'', '" & ISIN_CODE & "', '" & DEBT_INS_NAME & "', '" & ISSUER_NAME & "' "
                        sSqlInsert = sSqlInsert & " , '" & COUNTRY_CD_ISS & "' , NULL ,NULL "
                        sSqlInsert = sSqlInsert & " ,NULL , '',NULL  ,''   "
                        sSqlInsert = sSqlInsert & " , NULL , 0 , NULL , 0 , current date, current time, 'PLDMS' , 'HIPO' , '0', '',  '" & IP_NAME & "' , '" & IP_COUNTRY_CD & "'); "
                        
                        sSqlInsertLog = sHeadSqlInsertLog & " VALUES ('999999', '" & Trim(rs("TRANS_NUM")) & "', '" & dateInsert & "', '270001', '268017', " & TRAN_SEQ & "  , CURRENT TIMESTAMP, 'I'  "
                        sSqlInsertLog = sSqlInsertLog & " , 'GAIN', '" & dateInsert & "', '', '', '', '" & Trim(rs("CURRENCY")) & "', " & Abs(rs("PROFIT")) & ""
                        sSqlInsertLog = sSqlInsertLog & " ,'', '" & ISIN_CODE & "', '" & DEBT_INS_NAME & "', '" & ISSUER_NAME & "' "
                        sSqlInsertLog = sSqlInsertLog & " , '" & COUNTRY_CD_ISS & "' , NULL ,NULL "
                        sSqlInsertLog = sSqlInsertLog & " ,NULL , '',NULL  ,''  "
                        sSqlInsertLog = sSqlInsertLog & " , NULL , 0 , NULL , 0 , current date, current time, 'PLDMS' , 'HIPO' , '0', '',  '" & IP_NAME & "' , '" & IP_COUNTRY_CD & "'); "
                        
                        Conn.Execute (sSqlInsert)
                        Conn.Execute (sSqlInsertLog)
                        'SORN
                         count_insert = count_insert + 1
                        TRAN_SEQ = TRAN_SEQ + 1
            End If
            
             If Trim(rs("LOSS")) <> "0" Then
                        'LOSS
                        
                        sSqlInsert = sHeadSqlInsert & " VALUES ('999999', '" & Trim(rs("TRANS_NUM")) & "', '" & dateInsert & "', '270002', '268017', " & TRAN_SEQ & " "
                        sSqlInsert = sSqlInsert & " , 'LOSS', '" & dateInsert & "', '', '', '', '" & Trim(rs("CURRENCY")) & "', " & Abs(rs("LOSS")) & ""
                        sSqlInsert = sSqlInsert & " ,'', '" & ISIN_CODE & "', '" & DEBT_INS_NAME & "', '" & ISSUER_NAME & "' "
                        sSqlInsert = sSqlInsert & " , '" & COUNTRY_CD_ISS & "' , NULL ,NULL "
                        sSqlInsert = sSqlInsert & " ,NULL , '',NULL  ,'' "
                        sSqlInsert = sSqlInsert & " , NULL , 0 , NULL , 0 , current date, current time, 'PLDMS' , 'HIPO' , '0', '',  '" & IP_NAME & "', '" & IP_COUNTRY_CD & "'); "
                  
                        sSqlInsertLog = sHeadSqlInsertLog & " VALUES ('999999', '" & Trim(rs("TRANS_NUM")) & "', '" & dateInsert & "', '270002', '268017', " & TRAN_SEQ & "  , CURRENT TIMESTAMP, 'I'  "
                        sSqlInsertLog = sSqlInsertLog & " , 'LOSS', '" & dateInsert & "', '', '', '', '" & Trim(rs("CURRENCY")) & "', " & Abs(rs("LOSS")) & ""
                        sSqlInsertLog = sSqlInsertLog & " ,'', '" & ISIN_CODE & "', '" & DEBT_INS_NAME & "', '" & ISSUER_NAME & "' "
                        sSqlInsertLog = sSqlInsertLog & " , '" & COUNTRY_CD_ISS & "' , NULL ,NULL "
                        sSqlInsertLog = sSqlInsertLog & " ,NULL , '',NULL  ,'' "
                        sSqlInsertLog = sSqlInsertLog & " , NULL , 0 , NULL , 0 , current date, current time, 'PLDMS' , 'HIPO' , '0', '', '" & IP_NAME & "'  , '" & IP_COUNTRY_CD & "'); "
                        
                        Conn.Execute (sSqlInsert)
                        Conn.Execute (sSqlInsertLog)
                        'SORN
                         count_insert = count_insert + 1
                        TRAN_SEQ = TRAN_SEQ + 1
            End If
            Set rsSub1 = Nothing
            Set rsSub2 = Nothing
            rs.MoveNext
      Loop
      Set rs = Nothing
ErrDB:
      If Err.Number <> 0 Then
            Conn.RollbackTrans
            PTX_FI_1 = Err.Description
      Else
            Conn.CommitTrans
            PTX_FI_1 = ""
      End If
End Function

Private Function PTX_FI_21(iDate) As String
    '110316 Ton : เพิ่ม Process การแปลงข้อมูลที่อาจมีค่า '
      Dim tStart_dt, sSlqHol, sSqlSEQ, sSqlSelect As String
      Dim flag_insert, IP_NAME, DEBT_ARR_TYPE_CD, ISIN_CODE, DEBT_INS_NAME, ISSUER_NAME, COUNTRY_CD_ISS As String
      Dim ISSUE_DTE, MATURITY_DTE, COUPON_RATE, INT_COUNTRY_CD, IP_COUNTRY_CD As String
      Dim sSqlSub, sSqlInsert, sSqlInsertLog As String
      Dim TRAN_SEQ As Integer
      
      On Error GoTo ErrDB
      Conn.BeginTrans

      'Dim rs, rsSub1, rsSub2 As New ADODB.Recordset
      Dim rs, rsSub1, rsSub2 As Object
      Set rs = CreateObject("ADODB.Recordset")
      Set rsSub1 = CreateObject("ADODB.Recordset")
      Set rsSub2 = CreateObject("ADODB.Recordset")
      
      Dim sDate As String

      sDate = iDate
      tStart_dt = ""
      '-------------------------------------------------------------------- Find holiday -----------------------------------------------------------------
      sSlqHol = " SELECT DATE(DAYS(MAX(DATE))+1) AS T_START_DT From HOLIDAY WHERE HOLIDAY_FLAG ='N' "
      sSlqHol = sSlqHol & " AND DATE < '" & sDate & "'    WITH UR; "
      Set rs = Conn.Execute(sSlqHol)
      tStart_dt = Format(Trim(rs("T_START_DT")), "YYYY") & "-" & Format(Trim(rs("T_START_DT")), "MM") & "-" & Format(Trim(rs("T_START_DT")), "DD")
      Set rs = Nothing
      '-------------------------------------------------------------------- Find SEQ-----------------------------------------------------------------
            sSqlSEQ = " SELECT COALESCE(MAX(TRAN_SEQ), 0)  AS TRAN_SEQ "
            sSqlSEQ = sSqlSEQ & " FROM DS_PTX WHERE DATA_SET_DATE =  '" & sDate & "' AND DATA_SYSTEM_CD = 'HIPO'   WITH UR; "
            
            Set rs = Conn.Execute(sSqlSEQ)
           If Not rs.EOF Then
                  TRAN_SEQ = Trim(rs("TRAN_SEQ")) + 1
            Else
                  TRAN_SEQ = 1
            End If
            Set rs = Nothing

'-------------------------------------------------------------------- Select Record -----------------------------------------------------------------
      sSqlSelect = "  SELECT  BOND_ID, CURRENCY   , SUM(CHG_WAV) AS CHG_WAV  FROM TMS_FI_WAV"
      sSqlSelect = sSqlSelect & "  WHERE TRANS_DATE BETWEEN  '" & tStart_dt & "' AND '" & sDate & "'   "
      'sSqlSelect = sSqlSelect & "  WHERE SETTLE_DATE BETWEEN  '" & tStart_dt & "' AND '" & sDate & "'   "
      sSqlSelect = sSqlSelect & "  AND TYPE  = 'FA' AND CURRENCY <> 'THB'"
      sSqlSelect = sSqlSelect & "  GROUP BY BOND_ID,CURRENCY    WITH UR;  "
      
      Set rs = Conn.Execute(sSqlSelect)
      Do While Not rs.EOF
      
            flag_insert = ""
            IP_NAME = ""
            '------------------------------------------------------------------
            DEBT_ARR_TYPE_CD = ""
            ISIN_CODE = ""
            DEBT_INS_NAME = ""
            ISSUER_NAME = ""
            COUNTRY_CD_ISS = ""
            ISSUE_DTE = "null"
            MATURITY_DTE = "null"
            COUPON_RATE = "null"
            INT_COUNTRY_CD = ""
            '------------------------------------------------------------------
            
             '---------Select Group 1 ISIN_CODE,DEBT_INS_NAME,ISSUE_DTE,MATURITY_DTE,COUPON_RATE
                              sSqlSub = "  SELECT  SEC_FEATURE.ISIN_CODE AS ISIN_CODE ,  SEC_FEATURE.INST_DESC  AS DEBT_INS_NAME, SEC_FEATURE.ISSUE_DATE AS ISSUE_DTE  "
                              sSqlSub = sSqlSub & "    ,SEC_FEATURE.MAT_DATE AS MATURITY_DTE ,SEC_FEATURE.COUPON_RATE AS COUPON_RATE  "
                              sSqlSub = sSqlSub & "    FROM ESL_SEC_FEATURE AS SEC_FEATURE  WHERE SEC_FEATURE.SECURITY_ID ='" & Trim(rs("BOND_ID")) & "'   WITH UR;   "
                               
                                                            
                              Set rsSub1 = Conn.Execute(sSqlSub)
                              If Not rsSub1.EOF Then
                                          'ISIN_CODE = Trim(rsSub1("ISIN_CODE"))
                                          If Not IsNull(rsSub1("ISIN_CODE")) Then
                                                   ISIN_CODE = Replace(Trim(rsSub1("ISIN_CODE")), "&", " AND ")
                                          End If
                                          'DEBT_INS_NAME = Trim(rsSub1("DEBT_INS_NAME"))
                                          If Not IsNull(rsSub1("DEBT_INS_NAME")) Then
                                              DEBT_INS_NAME = Replace(Trim(rsSub1("DEBT_INS_NAME")), "&", " AND ")
                                          End If
                                          If IsDate(Trim(rsSub1("ISSUE_DTE"))) = True Then
                                                ISSUE_DTE = "'" & Format(Trim(rsSub1("ISSUE_DTE")), "YYYY") & "-" & Format(Trim(rsSub1("ISSUE_DTE")), "MM") & "-" & Format(Trim(rsSub1("ISSUE_DTE")), "DD") & "'"
                                          End If
                                          If IsDate(Trim(rsSub1("MATURITY_DTE"))) = True Then
                                                MATURITY_DTE = "'" & Format(Trim(rsSub1("MATURITY_DTE")), "YYYY") & "-" & Format(Trim(rsSub1("MATURITY_DTE")), "MM") & "-" & Format(Trim(rsSub1("MATURITY_DTE")), "DD") & "'"
                                          End If
                                          COUPON_RATE = Trim(rsSub1("COUPON_RATE"))

                              End If
                              Set rsSub1 = Nothing
                              
            '---------Select Group 2 ISSUER_NAME
                              sSqlSub = "  SELECT SYENTITY.LONG_NAME  AS ISSUER_NAME  FROM ESL_SEC_FEATURE AS SEC_FEATURE "
                              sSqlSub = sSqlSub & "  INNER JOIN ESL_SYENTITY AS SYENTITY       ON     SEC_FEATURE.ENTITY_CODE = SYENTITY.ENTITY_CODE "
                              sSqlSub = sSqlSub & "   WHERE SEC_FEATURE.SECURITY_ID ='" & Trim(rs("BOND_ID")) & "'   WITH UR;  "
                              Set rsSub1 = Conn.Execute(sSqlSub)
                              If Not rsSub1.EOF Then
                                    If Not rsSub1.EOF And Not IsNull(rsSub1("ISSUER_NAME")) Then
                                        'ISSUER_NAME = Trim(rsSub1("ISSUER_NAME"))
                                        ISSUER_NAME = Replace(Trim(rsSub1("ISSUER_NAME")), "&", " AND ")
                                    End If
                              End If
                              Set rsSub1 = Nothing
                              
            '---------Select Group 3 COUNTRY_CD_ISS
                  sSqlSub = "   SELECT MAP_CODE.MAP_CD2 AS COUNTRY_CD_ISS     FROM ESL_SEC_FEATURE AS SEC_FEATURE"
                  sSqlSub = sSqlSub & "   INNER JOIN ESL_SYENTITY AS SYENTITY       ON     SEC_FEATURE.ENTITY_CODE = SYENTITY.ENTITY_CODE"
                  sSqlSub = sSqlSub & "   INNER JOIN MAP_CODE AS MAP_CODE       ON     SYENTITY.DOMI_CODE =  MAP_CODE.MAP_CD1       AND    MAP_CODE.MAP_TABLE_CD = 'MAP029' "
                  sSqlSub = sSqlSub & "   WHERE SEC_FEATURE.SECURITY_ID ='" & Trim(rs("BOND_ID")) & "'   WITH UR;  "
                              Set rsSub1 = Conn.Execute(sSqlSub)
                              If Not rsSub1.EOF Then
                                    COUNTRY_CD_ISS = Trim(rsSub1("COUNTRY_CD_ISS"))
                              End If
                              Set rsSub1 = Nothing
                              
            '---------Select Group 4 INT_COUNTRY_CD
                  sSqlSub = "   SELECT MAP_CODE3.MAP_CD2 AS INT_COUNTRY_CD  FROM ESL_SEC_FEATURE AS SEC_FEATURE"
                  sSqlSub = sSqlSub & "   INNER JOIN MAP_CODE AS MAP_CODE3      ON     SEC_FEATURE.CURRENCY =  MAP_CODE3.MAP_CD1"
                  sSqlSub = sSqlSub & "   AND    MAP_CODE3.MAP_TABLE_CD = 'MAP030'"
                  sSqlSub = sSqlSub & "   WHERE SEC_FEATURE.SECURITY_ID ='" & Trim(rs("BOND_ID")) & "'   WITH UR;  "
                              
                              Set rsSub1 = Conn.Execute(sSqlSub)
                              If Not rsSub1.EOF Then
                                    INT_COUNTRY_CD = Trim(rsSub1("INT_COUNTRY_CD"))
                              End If
                              Set rsSub1 = Nothing
                              
            '---------Select Group 5 DEBT_ARR_TYPE_CD
                  sSqlSub = "  SELECT MAP_CODE1.MAP_CD2 AS DEBT_ARR_TYPE_CD  FROM ESL_SEC_FEATURE AS SEC_FEATURE"
                  sSqlSub = sSqlSub & "   INNER JOIN MAP_CODE AS MAP_CODE1   ON     SEC_FEATURE.BOND_TYPE = MAP_CODE1.MAP_CD1      AND    MAP_CODE1.MAP_TABLE_CD = 'MAP028'"
                  sSqlSub = sSqlSub & "   WHERE SEC_FEATURE.SECURITY_ID ='" & Trim(rs("BOND_ID")) & "'   WITH UR;  "
                              Set rsSub1 = Conn.Execute(sSqlSub)
                              If Not rsSub1.EOF Then
                                    DEBT_ARR_TYPE_CD = Trim(rsSub1("DEBT_ARR_TYPE_CD"))
                              End If
                              Set rsSub1 = Nothing


            
            
            sSqlInsert = " INSERT INTO DS_PTX ( ARR_TYPE_CD, ARR_NUMBER, TRAN_DTE, TRAN_TYPE_CD, ITEM_TYPE_CD, TRAN_SEQ, REC_PAY_ITEM_DSC, DATA_SET_DATE, PROVIDER_BR_NO, IBF_IND_CD, PAY_METHOD_CD, CURR_CD, TRAN_AMT_CURR, DEBT_ARR_TYPE_CD, ISIN_CODE, DEBT_INS_NAME, ISSUER_NAME, COUNTRY_CD_ISS, ISSUE_DTE, MATURITY_DTE, ORG_TRM, ORG_TRM_U, COUPON_RATE, INT_COUNTRY_CD, UNIT_OF_TRAN, SELL_SEC_AMT, PURCHASE_DTE, OS_AMT, LU_DTE, LU_TIME, USER_LU, DATA_SYSTEM_CD, FLAG_ON_OFF, PART_ID, IP_NAME, IP_COUNTRY_CD) "
            sSqlInsertLog = " INSERT INTO DS_PTX_LOG ( ARR_TYPE_CD, ARR_NUMBER, TRAN_DTE, TRAN_TYPE_CD, ITEM_TYPE_CD, TRAN_SEQ, ACTION_TIMESTAMP, ACTION_FLAG, REC_PAY_ITEM_DSC, DATA_SET_DATE, PROVIDER_BR_NO, IBF_IND_CD, PAY_METHOD_CD, CURR_CD, TRAN_AMT_CURR, DEBT_ARR_TYPE_CD, ISIN_CODE, DEBT_INS_NAME, ISSUER_NAME, COUNTRY_CD_ISS, ISSUE_DTE, MATURITY_DTE, ORG_TRM, ORG_TRM_U, COUPON_RATE, INT_COUNTRY_CD, UNIT_OF_TRAN, SELL_SEC_AMT, PURCHASE_DTE, OS_AMT, LU_DTE, LU_TIME, USER_LU, DATA_SYSTEM_CD, FLAG_ON_OFF, PART_ID, IP_NAME, IP_COUNTRY_CD) "
            
            
'
'             sSqlSub = " SELECT  MAP_CODE1.MAP_CD2 AS DEBT_ARR_TYPE_CD , SEC_FEATURE.ISIN_CODE AS ISIN_CODE ,  SEC_FEATURE.INST_TYPE AS DEBT_INS_NAME ,  SYENTITY.ENTITY_20CODE AS ISSUER_NAME , MAP_CODE2.MAP_CD2 AS COUNTRY_CD_ISS , SEC_FEATURE.ISSUE_DATE AS ISSUE_DTE ,SEC_FEATURE.MAT_DATE AS MATURITY_DTE ,SEC_FEATURE.COUPON_RATE AS COUPON_RATE , MAP_CODE3.MAP_CD2 AS INT_COUNTRY_CD ,  "
'             sSqlSub = sSqlSub & " SEC_FEATURE.*, SYENTITY.* ,MAP_CODE1.*,MAP_CODE2.*,MAP_CODE3.*  , MONTH(SEC_FEATURE.MAT_DATE -  SEC_FEATURE.ISSUE_DATE ) AS DIFF_MONTH ,YEAR(SEC_FEATURE.MAT_DATE -  SEC_FEATURE.ISSUE_DATE ) AS DIFF_YEAR  FROM ESL_SEC_FEATURE AS SEC_FEATURE "
'            sSqlSub = sSqlSub & " INNER JOIN ESL_SYENTITY AS SYENTITY ON SEC_FEATURE.ENTITY_CODE = SYENTITY.ENTITY_CODE"
'            sSqlSub = sSqlSub & " INNER JOIN MAP_CODE AS MAP_CODE1 ON SEC_FEATURE.BOND_TYPE = MAP_CODE1.MAP_CD1 AND MAP_CODE1.MAP_TABLE_CD = 'MAP028'"
'            sSqlSub = sSqlSub & " INNER JOIN MAP_CODE AS MAP_CODE2 ON SYENTITY.DOMI_CODE =  MAP_CODE2.MAP_CD1 AND MAP_CODE2.MAP_TABLE_CD = 'MAP029'"
'            sSqlSub = sSqlSub & " INNER JOIN MAP_CODE AS MAP_CODE3 ON SEC_FEATURE.CURRENCY =  MAP_CODE3.MAP_CD1 AND MAP_CODE3.MAP_TABLE_CD = 'MAP030'"
'            sSqlSub = sSqlSub & " WHERE SEC_FEATURE.SECURITY_ID ='" & Trim(rs("BOND_ID")) & "'   WITH UR;  "
'
'            Set rsSub1 = Conn.Execute(sSqlSub)
'
'            If Not rsSub1.EOF Then
'                  DEBT_ARR_TYPE_CD = Trim(rsSub1("DEBT_ARR_TYPE_CD"))
'                  ISIN_CODE = Trim(rsSub1("ISIN_CODE"))
'                  DEBT_INS_NAME = Trim(rsSub1("DEBT_INS_NAME"))
'                  ISSUER_NAME = Trim(rsSub1("ISSUER_NAME"))
'                  COUNTRY_CD_ISS = Trim(rsSub1("COUNTRY_CD_ISS"))
'                  If Trim(rsSub1("ISSUE_DTE")) <> "" Then
'                        ISSUE_DTE = "'" & Trim(rsSub1("ISSUE_DTE")) & "'"
'                  End If
'                  If Trim(rsSub1("MATURITY_DTE")) <> "" Then
'                        MATURITY_DTE = "'" & Trim(rsSub1("MATURITY_DTE")) & "'"
'                  End If
'                  COUPON_RATE = Trim(rsSub1("COUPON_RATE"))
'                  INT_COUNTRY_CD = Trim(rsSub1("INT_COUNTRY_CD"))
'            End If
            IP_NAME = ""
            IP_COUNTRY_CD = ""
                             
             
                   sSqlSub = " SELECT * FROM ESL_SEC_FEATURE AS SEC_FEATURE"
                   sSqlSub = sSqlSub & " INNER JOIN  ESL_SYENTITY AS SYENTITY ON SEC_FEATURE.ENTITY_CODE = SYENTITY.ENTITY_CODE"
                   sSqlSub = sSqlSub & "  INNER JOIN MAP_CODE AS MAP_CODE ON SYENTITY.DOMI_CODE =  MAP_CODE.MAP_CD1 AND MAP_CODE.MAP_TABLE_CD = 'MAP029' "
                   sSqlSub = sSqlSub & " WHERE SEC_FEATURE.SECURITY_ID ='" & Trim(rs("BOND_ID")) & "'   WITH UR; "

                   Set rsSub2 = Conn.Execute(sSqlSub)
                  If Not rsSub2.EOF Then
                        If Not IsNull(Trim(rsSub2("ENTITY_20CODE"))) Then
                              IP_NAME = Trim(rsSub2("ENTITY_20CODE"))
                              IP_NAME = Replace(IP_NAME, "&", " AND ")
                        End If
                      IP_COUNTRY_CD = Trim(rsSub2("MAP_CD2"))
                  Else
                      IP_NAME = ""
                      IP_COUNTRY_CD = ""
                  End If
             'Set rsSub2 = Conn.Execute(sSqlSub)
             'txt = InStr(1, rsSub2("COMMON.REMARK"), "/CF")
             'txt2 = Mid("txt", 1, txt)
            
            '110316 Ton : เพิ่ม Process การแปลงข้อมูลที่อาจมีค่า '
            ISIN_CODE = repQuote(ISIN_CODE)
            DEBT_INS_NAME = repQuote(DEBT_INS_NAME)
            ISSUER_NAME = repQuote(ISSUER_NAME)
            IP_NAME = repQuote(IP_NAME)
            '*****************************************************

            If Trim(rs("CHG_WAV")) > "0" Then
                  
                  sSqlInsert = sSqlInsert & "   VALUES ('999999' ,'9999999999' ,'" & sDate & "' ,'270001' ,'268017' ," & TRAN_SEQ & " ,'DISCOUNT' ,'" & sDate & "'  "
                  sSqlInsert = sSqlInsert & "   ,'','' ,'' ,'" & Trim(rs("CURRENCY")) & "'  ," & Abs(rs("CHG_WAV")) & " ,'' ,'" & ISIN_CODE & "'"
                  sSqlInsert = sSqlInsert & "   , '" & DEBT_INS_NAME & "' , '" & ISSUER_NAME & "' , '" & COUNTRY_CD_ISS & "'  ,null , null , null ,'', null , '' , null , 0 , null "
                  sSqlInsert = sSqlInsert & "    ,0 ,Current Date ,Current Time ,'PLDMS' ,'HIPO' ,'0' ,'' ,'" & IP_NAME & "' ,'" & IP_COUNTRY_CD & "' ) "
                  
                  sSqlInsertLog = sSqlInsertLog & "   VALUES ('999999' ,'9999999999' ,'" & sDate & "' ,'270001' ,'268017' ," & TRAN_SEQ & " , CURRENT TIMESTAMP, 'I' ,'DISCOUNT' ,'" & sDate & "'  "
                  sSqlInsertLog = sSqlInsertLog & "   ,'','' ,'' ,'" & Trim(rs("CURRENCY")) & "'  ," & Abs(rs("CHG_WAV")) & " ,'' ,'" & ISIN_CODE & "'"
                  sSqlInsertLog = sSqlInsertLog & "   , '" & DEBT_INS_NAME & "' , '" & ISSUER_NAME & "' , '" & COUNTRY_CD_ISS & "'  ,null , null , null ,'', null , '' , null , 0 , null "
                  sSqlInsertLog = sSqlInsertLog & "   ,0 ,Current Date ,Current Time ,'PLDMS' ,'HIPO' ,'0' ,'' ,'" & IP_NAME & "' ,'" & IP_COUNTRY_CD & "') "
                  
                  Conn.Execute (sSqlInsert)
                  Conn.Execute (sSqlInsertLog)
                  TRAN_SEQ = TRAN_SEQ + 1
                  count_insert = count_insert + 1
                  
                                        
            ElseIf Trim(rs("CHG_WAV")) < "0" Then
                  
                  sSqlInsert = sSqlInsert & "   VALUES ('999999' ,'9999999999' ,'" & sDate & "' ,'270002' ,'268017' ," & TRAN_SEQ & " ,'PREMIUM' ,'" & sDate & "'  "
                  sSqlInsert = sSqlInsert & "   ,'','' ,'' ,'" & Trim(rs("CURRENCY")) & "'  ," & Abs(rs("CHG_WAV")) & " ,'' ,'" & ISIN_CODE & "'"
                  sSqlInsert = sSqlInsert & "   , '" & DEBT_INS_NAME & "' , '" & ISSUER_NAME & "' , '" & COUNTRY_CD_ISS & "'  ,null , null , null ,'', null , '' , null , 0 , null "
                  sSqlInsert = sSqlInsert & "    ,0 ,Current Date ,Current Time ,'PLDMS' ,'HIPO' ,'0' ,'' ,'" & IP_NAME & "' ,'" & IP_COUNTRY_CD & "' ) "
                  
                  sSqlInsertLog = sSqlInsertLog & "   VALUES ('999999' ,'9999999999' ,'" & sDate & "' ,'270002' ,'268017' ," & TRAN_SEQ & " , CURRENT TIMESTAMP, 'I' ,'PREMIUM' ,'" & sDate & "'  "
                  sSqlInsertLog = sSqlInsertLog & "   ,'','' ,'' ,'" & Trim(rs("CURRENCY")) & "'  ," & Abs(rs("CHG_WAV")) & " ,'' ,'" & ISIN_CODE & "'"
                  sSqlInsertLog = sSqlInsertLog & "   , '" & DEBT_INS_NAME & "' , '" & ISSUER_NAME & "' , '" & COUNTRY_CD_ISS & "'  ,null , null , null ,'', null , '' , null , 0 , null "
                  sSqlInsertLog = sSqlInsertLog & "   ,0 ,Current Date ,Current Time ,'PLDMS' ,'HIPO' ,'0' ,'' ,'" & IP_NAME & "' ,'" & IP_COUNTRY_CD & "') "
                  Conn.Execute (sSqlInsert)
                  Conn.Execute (sSqlInsertLog)
                   'count_insert = count_insert + 1
                  TRAN_SEQ = TRAN_SEQ + 1
                  count_insert = count_insert + 1
                  
            End If

            Set rsSub1 = Nothing
            Set rsSub2 = Nothing
            rs.MoveNext
      Loop
      Set rs = Nothing

ErrDB:
      If Err.Number <> 0 Then
            Conn.RollbackTrans
            PTX_FI_21 = Err.Description
      Else
            Conn.CommitTrans
            PTX_FI_21 = ""
      End If
End Function

Private Function PTX_FI_22(iDate) As String
    '110316 Ton : เพิ่ม Process การแปลงข้อมูลที่อาจมีค่า '
      On Error GoTo ErrDB
      Conn.BeginTrans
      
      'Dim rs, rsSub1, rsSub2 As New ADODB.Recordset
      Dim rs, rsSub1, rsSub2 As Object
      Set rs = CreateObject("ADODB.Recordset")
      Set rsSub1 = CreateObject("ADODB.Recordset")
      Set rsSub2 = CreateObject("ADODB.Recordset")
      
      Dim sSlqHol, sSqlSEQ, sSqlSelect, sSqlSub, sHeadSqlInsert, sHeadSqlInsertLog, sSqlInsert, sSqlInsertLog As String
      Dim tStart_dt, sDate, flag_insert, sSplit, sORG_TRM, endTxt As String
      Dim TRAN_SEQ As Integer
      Dim IP_NAME, DEBT_ARR_TYPE_CD, ISIN_CODE, DEBT_INS_NAME, ISSUER_NAME, COUNTRY_CD_ISS As String
      Dim ISSUE_DTE, MATURITY_DTE, ORG_TRM, ORG_TRM_U, COUPON_RATE, INT_COUNTRY_CD, IP_COUNTRY_CD As String
      
      sDate = iDate
      tStart_dt = ""
      '-------------------------------------------------------------------- Find holiday -----------------------------------------------------------------
      sSlqHol = " SELECT DATE(DAYS(MAX(DATE))+1) AS T_START_DT From HOLIDAY WHERE HOLIDAY_FLAG ='N' "
      sSlqHol = sSlqHol & " AND DATE < '" & sDate & "'   WITH UR; "
      Set rs = Conn.Execute(sSlqHol)
      tStart_dt = Format(Trim(rs("T_START_DT")), "YYYY") & "-" & Format(Trim(rs("T_START_DT")), "MM") & "-" & Format(Trim(rs("T_START_DT")), "DD")
      Set rs = Nothing
      '-------------------------------------------------------------------- Find SEQ-----------------------------------------------------------------
            sSqlSEQ = " SELECT COALESCE(MAX(TRAN_SEQ), 0)  AS TRAN_SEQ "
            sSqlSEQ = sSqlSEQ & " FROM DS_PTX WHERE DATA_SET_DATE =  '" & sDate & "' AND DATA_SYSTEM_CD = 'HIPO'   WITH UR; "
            
            Set rs = Conn.Execute(sSqlSEQ)
           If Not rs.EOF Then
                  TRAN_SEQ = Trim(rs("TRAN_SEQ")) + 1
            Else
                  TRAN_SEQ = 1
            End If
            Set rs = Nothing
      '-------------------------------------------------------------------- Select Record -----------------------------------------------------------------

      sSqlSelect = "  SELECT  BOND_ID ,TRANS_NUM ,TRANS_DATE ,CPTY , CURRENCY ,CHG_WAV ,PROFIT ,LOSS   FROM TMS_FI_WAV "
      'sSqlSelect = sSqlSelect & "   WHERE  TRANS_DATE  BETWEEN '" & tStart_dt & "' AND '" & sDate & "'      AND  TYPE  = 'M' AND      CURRENCY <> 'THB'     WITH UR; "
      sSqlSelect = sSqlSelect & "   WHERE  SETTLE_DATE  BETWEEN '" & tStart_dt & "' AND '" & sDate & "'      AND  TYPE  = 'M' AND      CURRENCY <> 'THB'     WITH UR; "
      Set rs = Conn.Execute(sSqlSelect)
      'TRAN_SEQ = 0
      Do While Not rs.EOF
      
            flag_insert = ""
            IP_NAME = ""
            '------------------------------------------------------------------
            DEBT_ARR_TYPE_CD = ""
            ISIN_CODE = ""
            DEBT_INS_NAME = ""
            ISSUER_NAME = ""
            COUNTRY_CD_ISS = ""
            ISSUE_DTE = "null"
            MATURITY_DTE = "null"
            ORG_TRM = "null"
            ORG_TRM_U = ""
            COUPON_RATE = "null"
            INT_COUNTRY_CD = ""
            '------------------------------------------------------------------
            '---------Select Group 1 ISIN_CODE,DEBT_INS_NAME,ISSUE_DTE,MATURITY_DTE,COUPON_RATE
                              sSqlSub = "  SELECT  SEC_FEATURE.ISIN_CODE AS ISIN_CODE ,  SEC_FEATURE.INST_DESC AS DEBT_INS_NAME, SEC_FEATURE.ISSUE_DATE AS ISSUE_DTE  "
                              sSqlSub = sSqlSub & "    ,SEC_FEATURE.MAT_DATE AS MATURITY_DTE ,SEC_FEATURE.COUPON_RATE AS COUPON_RATE, "
                              sSqlSub = sSqlSub & "     MONTH(SEC_FEATURE.MAT_DATE -  SEC_FEATURE.ISSUE_DATE ) AS DIFF_MONTH ,YEAR(SEC_FEATURE.MAT_DATE -  SEC_FEATURE.ISSUE_DATE ) AS DIFF_YEAR  "
                              sSqlSub = sSqlSub & "    FROM ESL_SEC_FEATURE AS SEC_FEATURE  WHERE SEC_FEATURE.SECURITY_ID ='" & Trim(rs("BOND_ID")) & "'   WITH UR;   "
                               
                                                            
                              Set rsSub1 = Conn.Execute(sSqlSub)
                              If Not rsSub1.EOF Then
                                          'ISIN_CODE = Trim(rsSub1("ISIN_CODE"))
                                          If Not IsNull(rsSub1("ISIN_CODE")) Then
                                              ISIN_CODE = Replace(Trim(rsSub1("ISIN_CODE")), "&", " AND ")
                                          End If
                                          'DEBT_INS_NAME = Trim(rsSub1("DEBT_INS_NAME"))
                                          If Not IsNull(rsSub1("DEBT_INS_NAME")) Then
                                              DEBT_INS_NAME = Replace(Trim(rsSub1("DEBT_INS_NAME")), "&", " AND ")
                                          End If
                                          If IsDate(Trim(rsSub1("ISSUE_DTE"))) = True Then
                                                ISSUE_DTE = "'" & Format(Trim(rsSub1("ISSUE_DTE")), "YYYY") & "-" & Format(Trim(rsSub1("ISSUE_DTE")), "MM") & "-" & Format(Trim(rsSub1("ISSUE_DTE")), "DD") & "'"
                                          End If
                                          If IsDate(Trim(rsSub1("MATURITY_DTE"))) = True Then
                                                MATURITY_DTE = "'" & Format(Trim(rsSub1("MATURITY_DTE")), "YYYY") & "-" & Format(Trim(rsSub1("MATURITY_DTE")), "MM") & "-" & Format(Trim(rsSub1("MATURITY_DTE")), "DD") & "'"
                                          End If
                                          COUPON_RATE = Trim(rsSub1("COUPON_RATE"))
                                          
                                          'MATURITY_DTE , ISSUE_DTE
                                          sORG_TRM = Cal_ORG_TRM(Trim(rsSub1("MATURITY_DTE")), Trim(rsSub1("ISSUE_DTE")))
                                          sSplit = Split(sORG_TRM, ";")
                                          ORG_TRM = sSplit(0)
                                          ORG_TRM_U = sSplit(1)
                                          
'                                         If Trim(rsSub1("DIFF_YEAR")) >= "1" Then
'                                                      ORG_TRM = Trim(rsSub1("DIFF_YEAR"))
'                                                      ORG_TRM_U = "Y"
'                                          Else
'                                                      ORG_TRM = Trim(rsSub1("DIFF_MONTH"))
'                                                      ORG_TRM_U = "M"
'                                          End If
                              End If
                              Set rsSub1 = Nothing
                              
            '---------Select Group 2 ISSUER_NAME
                              sSqlSub = "  SELECT SYENTITY.LONG_NAME  AS ISSUER_NAME  FROM ESL_SEC_FEATURE AS SEC_FEATURE "
                              sSqlSub = sSqlSub & "  INNER JOIN ESL_SYENTITY AS SYENTITY       ON     SEC_FEATURE.ENTITY_CODE = SYENTITY.ENTITY_CODE "
                              sSqlSub = sSqlSub & "   WHERE SEC_FEATURE.SECURITY_ID ='" & Trim(rs("BOND_ID")) & "'   WITH UR;  "
                              Set rsSub1 = Conn.Execute(sSqlSub)
                              If Not rsSub1.EOF Then
                                    If Not rsSub1.EOF And Not IsNull(rsSub1("ISSUER_NAME")) Then
                                        'ISSUER_NAME = Trim(rsSub1("ISSUER_NAME"))
                                        ISSUER_NAME = Replace(Trim(rsSub1("ISSUER_NAME")), "&", " AND ")
                                    End If
                              End If
                              Set rsSub1 = Nothing
                              
            '---------Select Group 3 COUNTRY_CD_ISS
                  sSqlSub = "   SELECT MAP_CODE.MAP_CD2 AS COUNTRY_CD_ISS     FROM ESL_SEC_FEATURE AS SEC_FEATURE"
                  sSqlSub = sSqlSub & "   INNER JOIN ESL_SYENTITY AS SYENTITY       ON     SEC_FEATURE.ENTITY_CODE = SYENTITY.ENTITY_CODE"
                  sSqlSub = sSqlSub & "   INNER JOIN MAP_CODE AS MAP_CODE       ON     SYENTITY.DOMI_CODE =  MAP_CODE.MAP_CD1       AND    MAP_CODE.MAP_TABLE_CD = 'MAP029' "
                  sSqlSub = sSqlSub & "   WHERE SEC_FEATURE.SECURITY_ID ='" & Trim(rs("BOND_ID")) & "'   WITH UR;  "
                              Set rsSub1 = Conn.Execute(sSqlSub)
                              If Not rsSub1.EOF Then
                                    COUNTRY_CD_ISS = Trim(rsSub1("COUNTRY_CD_ISS"))
                              End If
                              Set rsSub1 = Nothing
                              
            '---------Select Group 4 INT_COUNTRY_CD
                  sSqlSub = "   SELECT MAP_CODE3.MAP_CD2 AS INT_COUNTRY_CD  FROM ESL_SEC_FEATURE AS SEC_FEATURE"
                  sSqlSub = sSqlSub & "   INNER JOIN MAP_CODE AS MAP_CODE3      ON     SEC_FEATURE.CURRENCY =  MAP_CODE3.MAP_CD1"
                  sSqlSub = sSqlSub & "   AND    MAP_CODE3.MAP_TABLE_CD = 'MAP030'"
                  sSqlSub = sSqlSub & "   WHERE SEC_FEATURE.SECURITY_ID ='" & Trim(rs("BOND_ID")) & "'   WITH UR;  "
                              
                              Set rsSub1 = Conn.Execute(sSqlSub)
                              If Not rsSub1.EOF Then
                                    INT_COUNTRY_CD = Trim(rsSub1("INT_COUNTRY_CD"))
                              End If
                              Set rsSub1 = Nothing
                              
            '---------Select Group 5 DEBT_ARR_TYPE_CD
                  sSqlSub = "  SELECT MAP_CODE1.MAP_CD2 AS DEBT_ARR_TYPE_CD  FROM ESL_SEC_FEATURE AS SEC_FEATURE"
                  sSqlSub = sSqlSub & "   INNER JOIN MAP_CODE AS MAP_CODE1   ON     SEC_FEATURE.BOND_TYPE = MAP_CODE1.MAP_CD1      AND    MAP_CODE1.MAP_TABLE_CD = 'MAP028'"
                  sSqlSub = sSqlSub & "   WHERE SEC_FEATURE.SECURITY_ID ='" & Trim(rs("BOND_ID")) & "'   WITH UR;  "
                              Set rsSub1 = Conn.Execute(sSqlSub)
                              If Not rsSub1.EOF Then
                                    DEBT_ARR_TYPE_CD = Trim(rsSub1("DEBT_ARR_TYPE_CD"))
                              End If
                              Set rsSub1 = Nothing


            
            sHeadSqlInsert = " INSERT INTO DS_PTX ( ARR_TYPE_CD, ARR_NUMBER, TRAN_DTE, TRAN_TYPE_CD, ITEM_TYPE_CD, TRAN_SEQ, REC_PAY_ITEM_DSC, DATA_SET_DATE, PROVIDER_BR_NO, IBF_IND_CD, PAY_METHOD_CD, CURR_CD, TRAN_AMT_CURR, DEBT_ARR_TYPE_CD, ISIN_CODE, DEBT_INS_NAME, ISSUER_NAME, COUNTRY_CD_ISS, ISSUE_DTE, MATURITY_DTE, ORG_TRM, ORG_TRM_U, COUPON_RATE, INT_COUNTRY_CD, UNIT_OF_TRAN, SELL_SEC_AMT, PURCHASE_DTE, OS_AMT, LU_DTE, LU_TIME, USER_LU, DATA_SYSTEM_CD, FLAG_ON_OFF, PART_ID, IP_NAME, IP_COUNTRY_CD) "
            sHeadSqlInsertLog = " INSERT INTO DS_PTX_LOG ( ARR_TYPE_CD, ARR_NUMBER, TRAN_DTE, TRAN_TYPE_CD, ITEM_TYPE_CD, TRAN_SEQ, ACTION_TIMESTAMP, ACTION_FLAG, REC_PAY_ITEM_DSC, DATA_SET_DATE, PROVIDER_BR_NO, IBF_IND_CD, PAY_METHOD_CD, CURR_CD, TRAN_AMT_CURR, DEBT_ARR_TYPE_CD, ISIN_CODE, DEBT_INS_NAME, ISSUER_NAME, COUNTRY_CD_ISS, ISSUE_DTE, MATURITY_DTE, ORG_TRM, ORG_TRM_U, COUPON_RATE, INT_COUNTRY_CD, UNIT_OF_TRAN, SELL_SEC_AMT, PURCHASE_DTE, OS_AMT, LU_DTE, LU_TIME, USER_LU, DATA_SYSTEM_CD, FLAG_ON_OFF, PART_ID, IP_NAME, IP_COUNTRY_CD) "
            
            
'
'             sSqlSub = " SELECT  MAP_CODE1.MAP_CD2 AS DEBT_ARR_TYPE_CD , SEC_FEATURE.ISIN_CODE AS ISIN_CODE ,  SEC_FEATURE.INST_TYPE AS DEBT_INS_NAME ,  SYENTITY.ENTITY_20CODE AS ISSUER_NAME , MAP_CODE2.MAP_CD2 AS COUNTRY_CD_ISS , SEC_FEATURE.ISSUE_DATE AS ISSUE_DTE ,SEC_FEATURE.MAT_DATE AS MATURITY_DTE ,SEC_FEATURE.COUPON_RATE AS COUPON_RATE , MAP_CODE3.MAP_CD2 AS INT_COUNTRY_CD ,  "
'             sSqlSub = sSqlSub & " SEC_FEATURE.*, SYENTITY.* ,MAP_CODE1.*,MAP_CODE2.*,MAP_CODE3.*  , MONTH(SEC_FEATURE.MAT_DATE -  SEC_FEATURE.ISSUE_DATE ) AS DIFF_MONTH ,YEAR(SEC_FEATURE.MAT_DATE -  SEC_FEATURE.ISSUE_DATE ) AS DIFF_YEAR  FROM ESL_SEC_FEATURE AS SEC_FEATURE "
'            sSqlSub = sSqlSub & " INNER JOIN ESL_SYENTITY AS SYENTITY ON SEC_FEATURE.ENTITY_CODE = SYENTITY.ENTITY_CODE"
'            sSqlSub = sSqlSub & " INNER JOIN MAP_CODE AS MAP_CODE1 ON SEC_FEATURE.BOND_TYPE = MAP_CODE1.MAP_CD1 AND MAP_CODE1.MAP_TABLE_CD = 'MAP028'"
'            sSqlSub = sSqlSub & " INNER JOIN MAP_CODE AS MAP_CODE2 ON SYENTITY.DOMI_CODE =  MAP_CODE2.MAP_CD1 AND MAP_CODE2.MAP_TABLE_CD = 'MAP029'"
'            sSqlSub = sSqlSub & " INNER JOIN MAP_CODE AS MAP_CODE3 ON SEC_FEATURE.CURRENCY =  MAP_CODE3.MAP_CD1 AND MAP_CODE3.MAP_TABLE_CD = 'MAP030'"
'            sSqlSub = sSqlSub & " WHERE SEC_FEATURE.SECURITY_ID ='" & Trim(rs("BOND_ID")) & "'   WITH UR;  "
'
'            Set rsSub1 = Conn.Execute(sSqlSub)
'            If Not rsSub1.EOF Then
'                        If Trim(rsSub1("DIFF_YEAR")) >= "1" Then
'                                    ORG_TRM = Trim(rsSub1("DIFF_YEAR"))
'                                    ORG_TRM_U = "Y"
'                        Else
'                                    ORG_TRM = Trim(rsSub1("DIFF_MONTH"))
'                                    ORG_TRM_U = "M"
'                        End If
'
'                        DEBT_ARR_TYPE_CD = Trim(rsSub1("DEBT_ARR_TYPE_CD"))
'                        ISIN_CODE = Trim(rsSub1("ISIN_CODE"))
'                        DEBT_INS_NAME = Trim(rsSub1("DEBT_INS_NAME"))
'                        ISSUER_NAME = Trim(rsSub1("ISSUER_NAME"))
'                        COUNTRY_CD_ISS = Trim(rsSub1("COUNTRY_CD_ISS"))
'                        If Trim(rsSub1("ISSUE_DTE")) <> "" Then
'                              ISSUE_DTE = "'" & Trim(rsSub1("ISSUE_DTE")) & "'"
'                        End If
'                        If Trim(rsSub1("MATURITY_DTE")) <> "" Then
'                              MATURITY_DTE = "'" & Trim(rsSub1("MATURITY_DTE")) & "'"
'                        End If
'                        COUPON_RATE = Trim(rsSub1("COUPON_RATE"))
'                        INT_COUNTRY_CD = Trim(rsSub1("INT_COUNTRY_CD"))
'            End If

             If Trim(rs("TRANS_NUM")) <> "" Then
                                    
                  sSqlSub = " SELECT * FROM ESL_COMMON AS COMMON "
                  sSqlSub = sSqlSub & " INNER JOIN  ESL_SYENTITY AS SYENTITY ON COMMON.CNTR_PARTY = SYENTITY.ENTITY_CODE"
                  sSqlSub = sSqlSub & "  INNER JOIN MAP_CODE AS MAP_CODE ON SYENTITY.DOMI_CODE =  MAP_CODE.MAP_CD1 AND MAP_CODE.MAP_TABLE_CD = 'MAP029' "
                  'sSqlSub = sSqlSub & " WHERE COMMON.TRANS_NUM ='" & Trim(rs("TRANS_NUM")) & "'  WITH UR;  "
                  'SORN
                  sSqlSub = sSqlSub & " WHERE COMMON.TRANS_NUM ='" & Trim(rs("TRANS_NUM")) & "' and  COMMON.AS_OF_DT =  '" & sDate & "'  WITH UR;  "
                  
                  Set rsSub2 = Conn.Execute(sSqlSub)
                  If Not rsSub2.EOF Then
                        If Not IsNull(Trim(rsSub2("REMARK"))) Then
                                    endTxt = InStr(1, rsSub2("REMARK"), "/CF")
                                     If endTxt <> "0" Then
                                           IP_NAME = Mid(rsSub2("REMARK"), 1, endTxt)
                                           IP_NAME = Replace(IP_NAME, "&", " AND ")
'                                     Else
'                                           IP_NAME = rsSub2("ENTITY_20CODE")
'                                           IP_NAME = Replace(IP_NAME, "&", " AND ")
                                     End If
                        End If
                              IP_COUNTRY_CD = Trim(rsSub2("MAP_CD2"))
                              '20110422 yos add IP_NAME
                              If Not IsNull(Trim(rsSub2("ENTITY_20CODE"))) Then
                                   IP_NAME = Trim(rsSub2("ENTITY_20CODE"))
                                   IP_NAME = Replace(IP_NAME, "&", " AND ")
                              End If
                  Else
                        IP_NAME = ""
                        IP_COUNTRY_CD = ""
                  End If
                                                
            Else
                  sSqlSub = " SELECT * FROM ESL_SEC_FEATURE AS SEC_FEATURE"
                  sSqlSub = sSqlSub & " INNER JOIN  ESL_SYENTITY AS SYENTITY ON SEC_FEATURE.ENTITY_CODE = SYENTITY.ENTITY_CODE"
                  sSqlSub = sSqlSub & "  INNER JOIN MAP_CODE AS MAP_CODE ON SYENTITY.DOMI_CODE =  MAP_CODE.MAP_CD1 AND MAP_CODE.MAP_TABLE_CD = 'MAP029' "
                  sSqlSub = sSqlSub & " WHERE SEC_FEATURE.SECURITY_ID ='" & Trim(rs("BOND_ID")) & "'   WITH UR;  "
                  
                  Set rsSub2 = Conn.Execute(sSqlSub)
                  If Not rsSub2.EOF Then
                        If Not IsNull(Trim(rsSub2("ENTITY_20CODE"))) Then
                              IP_NAME = rsSub2("ENTITY_20CODE")
                              IP_NAME = Replace(IP_NAME, "&", " AND ")
                        End If
                        IP_COUNTRY_CD = Trim(rsSub2("MAP_CD2"))
                  Else
                        IP_NAME = ""
                        IP_COUNTRY_CD = ""
                  End If
                  
            End If
             'Set rsSub2 = Conn.Execute(sSqlSub)
             'txt = InStr(1, rsSub2("COMMON.REMARK"), "/CF")
             'txt2 = Mid("txt", 1, txt)
             
            '110316 Ton : เพิ่ม Process การแปลงข้อมูลที่อาจมีค่า '
            ISIN_CODE = repQuote(ISIN_CODE)
            DEBT_INS_NAME = repQuote(DEBT_INS_NAME)
            ISSUER_NAME = repQuote(ISSUER_NAME)
            IP_NAME = repQuote(IP_NAME)
            '*****************************************************
            
            If Trim(rs("CHG_WAV")) <> "0" Then
                        sSqlInsert = sHeadSqlInsert & "  VALUES ( '999999' ,'9999999999' ,'" & sDate & "' ,'270003' ,'268027' , " & TRAN_SEQ & "  ,'' ,'" & sDate & "' ,'' ,'' ,'234005'"
                        sSqlInsert = sSqlInsert & "  ,  '" & Trim(rs("CURRENCY")) & "' ,  " & Abs(rs("CHG_WAV")) & " , '" & DEBT_ARR_TYPE_CD & "' , '" & ISIN_CODE & "' , '" & DEBT_INS_NAME & "'  "
                        sSqlInsert = sSqlInsert & "  , '" & ISSUER_NAME & "'  , '" & COUNTRY_CD_ISS & "'  , " & ISSUE_DTE & " , " & MATURITY_DTE & "  , " & ORG_TRM & "  , '" & ORG_TRM_U & "' "
                        sSqlInsert = sSqlInsert & "  ,  " & COUPON_RATE & "  ,'" & INT_COUNTRY_CD & "' , NULL , 0 , NULL , 0  , CURRENT DATE , CURRENT TIME "
                        sSqlInsert = sSqlInsert & "  , 'PLDMS'  , 'HIPO'  , '0' , '' , '" & IP_NAME & "' , '" & IP_COUNTRY_CD & "');  "
                        
                        sSqlInsertLog = sHeadSqlInsertLog & "  VALUES ( '999999' ,'9999999999' ,'" & sDate & "' ,'270003' ,'268027' , " & TRAN_SEQ & "   ,CURRENT TIMESTAMP, 'I'    ,'' ,'" & sDate & "' ,'' ,'' ,'234005'"
                        sSqlInsertLog = sSqlInsertLog & "  ,  '" & Trim(rs("CURRENCY")) & "' ,  " & Abs(rs("CHG_WAV")) & " , '" & DEBT_ARR_TYPE_CD & "' , '" & ISIN_CODE & "' , '" & DEBT_INS_NAME & "'  "
                        sSqlInsertLog = sSqlInsertLog & "  , '" & ISSUER_NAME & "'  , '" & COUNTRY_CD_ISS & "'  , " & ISSUE_DTE & " , " & MATURITY_DTE & "  , " & ORG_TRM & "  , '" & ORG_TRM_U & "' "
                        sSqlInsertLog = sSqlInsertLog & "  ,  " & COUPON_RATE & "  ,'" & INT_COUNTRY_CD & "' , NULL , 0 , NULL , 0  , CURRENT DATE , CURRENT TIME "
                        sSqlInsertLog = sSqlInsertLog & "  , 'PLDMS'  , 'HIPO'  , '0' , '' , '" & IP_NAME & "' , '" & IP_COUNTRY_CD & "');  "

                        Conn.Execute (sSqlInsert)
                        Conn.Execute (sSqlInsertLog)
                        count_insert = count_insert + 1
                        
                        TRAN_SEQ = TRAN_SEQ + 1
                        
          End If
            
          If Trim(rs("PROFIT")) <> "0" Then
                        
                        
                        sSqlInsert = sHeadSqlInsert & "     VALUES ( '999999' ,'9999999999' ,'" & sDate & "' ,'270001' ,'268017' ,  " & TRAN_SEQ & "  ,'GAIN' , '" & sDate & "' ,'' ,'' ,''  "
                        sSqlInsert = sSqlInsert & "   , '" & Trim(rs("CURRENCY")) & "' ," & Abs(rs("PROFIT")) & " ,'' , '" & ISIN_CODE & "'  "
                        sSqlInsert = sSqlInsert & "    ,  '" & DEBT_INS_NAME & "' , '" & ISSUER_NAME & "' , '" & COUNTRY_CD_ISS & "'   "
                        sSqlInsert = sSqlInsert & "    , NULL,NULL,NULL,'',NULL,'',NULL,0,NULL,0  ,Current Date ,Current Time ,'PLDMS' ,'HIPO' ,'0' ,'' ,'" & IP_NAME & "' ,'" & IP_COUNTRY_CD & "' ) "
                        
                        
                        sSqlInsertLog = sHeadSqlInsertLog & "     VALUES ( '999999' ,'9999999999' ,'" & sDate & "' ,'270001' ,'268017' ,  " & TRAN_SEQ & "  ,CURRENT TIMESTAMP, 'I'   ,'GAIN' , '" & sDate & "' ,'' ,'' ,''  "
                        sSqlInsertLog = sSqlInsertLog & "   , '" & Trim(rs("CURRENCY")) & "' ," & Abs(rs("PROFIT")) & " ,'' , '" & ISIN_CODE & "'  "
                        sSqlInsertLog = sSqlInsertLog & "    ,  '" & DEBT_INS_NAME & "' , '" & ISSUER_NAME & "' , '" & COUNTRY_CD_ISS & "'   "
                        sSqlInsertLog = sSqlInsertLog & "    , NULL,NULL,NULL,'',NULL,'',NULL,0,NULL,0  ,Current Date ,Current Time ,'PLDMS' ,'HIPO' ,'0' ,'' ,'" & IP_NAME & "' ,'" & IP_COUNTRY_CD & "' ) "

                        Conn.Execute (sSqlInsert)
                        Conn.Execute (sSqlInsertLog)
                        count_insert = count_insert + 1
                        
                        TRAN_SEQ = TRAN_SEQ + 1

          End If

          If Trim(rs("LOSS")) <> "0" Then
                        
                        sSqlInsert = sHeadSqlInsert & "     VALUES ( '999999' ,'9999999999' ,'" & sDate & "' ,'270002' ,'268017' ,  " & TRAN_SEQ & "  ,'LOSS' , '" & sDate & "' ,'' ,'' ,''  "
                        sSqlInsert = sSqlInsert & "   , '" & Trim(rs("CURRENCY")) & "' ," & Abs(rs("LOSS")) & " ,'' , '" & ISIN_CODE & "'  "
                        sSqlInsert = sSqlInsert & "    ,  '" & DEBT_INS_NAME & "' , '" & ISSUER_NAME & "' , '" & COUNTRY_CD_ISS & "'   "
                        sSqlInsert = sSqlInsert & "    , NULL,NULL,NULL,'',NULL,'',NULL,0,NULL,0  ,Current Date ,Current Time ,'PLDMS' ,'HIPO' ,'0' ,'' ,'" & IP_NAME & "' ,'" & IP_COUNTRY_CD & "' ) "
                        
                        
                        sSqlInsertLog = sHeadSqlInsertLog & "     VALUES ( '999999' ,'9999999999' ,'" & sDate & "' ,'270002' ,'268017' ,  " & TRAN_SEQ & "  ,CURRENT TIMESTAMP, 'I'   ,'LOSS' , '" & sDate & "' ,'' ,'' ,''  "
                        sSqlInsertLog = sSqlInsertLog & "   , '" & Trim(rs("CURRENCY")) & "' ," & Abs(rs("LOSS")) & " ,'' , '" & ISIN_CODE & "'  "
                        sSqlInsertLog = sSqlInsertLog & "    ,  '" & DEBT_INS_NAME & "' , '" & ISSUER_NAME & "' , '" & COUNTRY_CD_ISS & "'   "
                        sSqlInsertLog = sSqlInsertLog & "    , NULL,NULL,NULL,'',NULL,'',NULL,0,NULL,0  ,Current Date ,Current Time ,'PLDMS' ,'HIPO' ,'0' ,'' ,'" & IP_NAME & "' ,'" & IP_COUNTRY_CD & "' ) "

                        Conn.Execute (sSqlInsert)
                        Conn.Execute (sSqlInsertLog)
                        TRAN_SEQ = TRAN_SEQ + 1
                        count_insert = count_insert + 1

          End If

            Set rsSub1 = Nothing
            Set rsSub2 = Nothing
            rs.MoveNext
      
      Loop
      Set rs = Nothing

ErrDB:
      If Err.Number <> 0 Then
            Conn.RollbackTrans
            PTX_FI_22 = Err.Description
            
      Else
            Conn.CommitTrans
            PTX_FI_22 = ""
      End If
End Function
Private Function PTX_FI_COMMON(iDate) As String
      'sheet 3-PTX-FI(2)
      '20110215 change main query
      On Error GoTo ErrDB
      Conn.BeginTrans
      
      'Dim rs, rsSub1, rsSub2 As New ADODB.Recordset
      Dim rs, rsSub1, rsSub2 As Object
      Set rs = CreateObject("ADODB.Recordset")
      Set rsSub1 = CreateObject("ADODB.Recordset")
      Set rsSub2 = CreateObject("ADODB.Recordset")
      Dim sData_set_date, sDate As String
      Dim sSqlSEQ, sSqlSelect, sSqlInsert, sSqlInsertLog, sSqlSub As String
      Dim flag_insert As String
      Dim IP_NAME, DEBT_ARR_TYPE_CD, ISIN_CODE, DEBT_INS_NAME, ISSUER_NAME, COUNTRY_CD_ISS, ISSUE_DTE As String
      Dim MATURITY_DTE, ORG_TRM, ORG_TRM_U, COUPON_RATE, INT_COUNTRY_CD As String
      Dim TRAN_SEQ As Integer
      
      sDate = iDate
      '-------------------------------------------------------------------- Find SEQ-----------------------------------------------------------------
            sSqlSEQ = " SELECT COALESCE(MAX(TRAN_SEQ), 0)  AS TRAN_SEQ "
            sSqlSEQ = sSqlSEQ & " FROM DS_PTX WHERE DATA_SET_DATE =  '" & sDate & "' AND DATA_SYSTEM_CD = 'HIPO'    WITH UR; "
            
            Set rs = Conn.Execute(sSqlSEQ)
           If Not rs.EOF Then
                  TRAN_SEQ = Trim(rs("TRAN_SEQ")) + 1
            Else
                  TRAN_SEQ = 1
            End If
            Set rs = Nothing
            
      '-------------------------------------------------------------------- Select Record -----------------------------------------------------------------
      sSqlSelect = "  SELECT COMMON.*,TRANS_NUM ,TRADE_DATE , CNTR_PARTY ,CONT_FAMT ,SECURITY_ID ,PURCHASE_INT ,PURCHASE_INT_DR_CR_IND  "
      sSqlSelect = sSqlSelect & "   FROM ESL_COMMON AS COMMON  "
      'sSqlSelect = sSqlSelect & "     WHERE  PROD_GROUP = 'FI'      AND    CONT_CUR  <> 'THB'     AND   GL_PROD_CD <> '0181'    AND    TRANS_Status = 1       AND    TRADE_DATE = '" & sDate & "'   AND    AS_OF_DT =  '" & sDate & "' and  PURCHASE_INT <> 0 WITH UR;  "
      sSqlSelect = sSqlSelect & "     WHERE  PROD_GROUP = 'FI'      AND    CONT_CUR  <> 'THB'     AND   GL_PROD_CD <> '0181'    AND    TRANS_Status = 1       AND    settle_date = '" & sDate & "'   AND    AS_OF_DT =  '" & sDate & "' and  PURCHASE_INT <> 0 WITH UR;  "
      Set rs = Conn.Execute(sSqlSelect)
      Do While Not rs.EOF
            'sData_set_date = Format(rs("TRADE_DATE"), "YYYY-MM-DD")
            sData_set_date = Format(rs("SETTLE_DATE"), "YYYY-MM-DD")
            flag_insert = ""
            IP_NAME = ""
            '------------------------------------------------------------------
            DEBT_ARR_TYPE_CD = ""
            ISIN_CODE = ""
            DEBT_INS_NAME = ""
            ISSUER_NAME = ""
            COUNTRY_CD_ISS = ""
            ISSUE_DTE = "null"
            MATURITY_DTE = "null"
            ORG_TRM = "null"
            ORG_TRM_U = ""
            COUPON_RATE = "null"
            INT_COUNTRY_CD = ""
            '------------------------------------------------------------------
            
            
            sSqlInsert = " INSERT INTO DS_PTX ( ARR_TYPE_CD, ARR_NUMBER, TRAN_DTE, TRAN_TYPE_CD, ITEM_TYPE_CD, TRAN_SEQ, REC_PAY_ITEM_DSC, DATA_SET_DATE, PROVIDER_BR_NO, IBF_IND_CD, PAY_METHOD_CD, CURR_CD, TRAN_AMT_CURR, DEBT_ARR_TYPE_CD, ISIN_CODE, DEBT_INS_NAME, ISSUER_NAME, COUNTRY_CD_ISS, ISSUE_DTE, MATURITY_DTE, ORG_TRM, ORG_TRM_U, COUPON_RATE, INT_COUNTRY_CD, UNIT_OF_TRAN, SELL_SEC_AMT, PURCHASE_DTE, OS_AMT, LU_DTE, LU_TIME, USER_LU, DATA_SYSTEM_CD, FLAG_ON_OFF, PART_ID, IP_NAME, IP_COUNTRY_CD) "
            sSqlInsertLog = " INSERT INTO DS_PTX_LOG ( ARR_TYPE_CD, ARR_NUMBER, TRAN_DTE, TRAN_TYPE_CD, ITEM_TYPE_CD, TRAN_SEQ, ACTION_TIMESTAMP, ACTION_FLAG, REC_PAY_ITEM_DSC, DATA_SET_DATE, PROVIDER_BR_NO, IBF_IND_CD, PAY_METHOD_CD, CURR_CD, TRAN_AMT_CURR, DEBT_ARR_TYPE_CD, ISIN_CODE, DEBT_INS_NAME, ISSUER_NAME, COUNTRY_CD_ISS, ISSUE_DTE, MATURITY_DTE, ORG_TRM, ORG_TRM_U, COUPON_RATE, INT_COUNTRY_CD, UNIT_OF_TRAN, SELL_SEC_AMT, PURCHASE_DTE, OS_AMT, LU_DTE, LU_TIME, USER_LU, DATA_SYSTEM_CD, FLAG_ON_OFF, PART_ID, IP_NAME, IP_COUNTRY_CD) "
            

                  '---------Select Group 1 ISIN_CODE,DEBT_INS_NAME,
                              sSqlSub = "  SELECT  SEC_FEATURE.ISIN_CODE AS ISIN_CODE ,  SEC_FEATURE.INST_DESC AS DEBT_INS_NAME  "
                              sSqlSub = sSqlSub & "    FROM ESL_SEC_FEATURE AS SEC_FEATURE  WHERE SEC_FEATURE.SECURITY_ID ='" & Trim(rs("SECURITY_ID")) & "'   WITH UR;   "
                                          
                              Set rsSub1 = Conn.Execute(sSqlSub)
                              If Not rsSub1.EOF Then

                                          'ISIN_CODE = Trim(rsSub1("ISIN_CODE"))
                                          If Not IsNull(rsSub1("ISIN_CODE")) Then
                                              ISIN_CODE = Replace(Trim(rsSub1("ISIN_CODE")), "&", " AND ")
                                          End If
                                          'DEBT_INS_NAME = Trim(rsSub1("DEBT_INS_NAME"))
                                          If Not IsNull(rsSub1("DEBT_INS_NAME")) Then
                                              DEBT_INS_NAME = Replace(Trim(rsSub1("DEBT_INS_NAME")), "&", " AND ")
                                          End If
                              End If
                              Set rsSub1 = Nothing

 
            '---------Select Group 2 ISSUER_NAME
                              sSqlSub = "  SELECT SYENTITY.LONG_NAME  AS ISSUER_NAME  FROM ESL_SEC_FEATURE AS SEC_FEATURE "
                              sSqlSub = sSqlSub & "  INNER JOIN ESL_SYENTITY AS SYENTITY       ON     SEC_FEATURE.ENTITY_CODE = SYENTITY.ENTITY_CODE "
                              sSqlSub = sSqlSub & "   WHERE SEC_FEATURE.SECURITY_ID ='" & Trim(rs("SECURITY_ID")) & "'   WITH UR;  "
                              Set rsSub1 = Conn.Execute(sSqlSub)
                              If Not rsSub1.EOF Then
                                    If Not rsSub1.EOF And Not IsNull(rsSub1("ISSUER_NAME")) Then
                                        'ISSUER_NAME = Trim(rsSub1("ISSUER_NAME"))
                                        ISSUER_NAME = Replace(Trim(rsSub1("ISSUER_NAME")), "&", " AND ")
                                    End If
                              End If
                              Set rsSub1 = Nothing
                              
            '---------Select Group 3 COUNTRY_CD_ISS
                  sSqlSub = "   SELECT MAP_CODE.MAP_CD2 AS COUNTRY_CD_ISS     FROM ESL_SEC_FEATURE AS SEC_FEATURE"
                  sSqlSub = sSqlSub & "   INNER JOIN ESL_SYENTITY AS SYENTITY       ON     SEC_FEATURE.ENTITY_CODE = SYENTITY.ENTITY_CODE"
                  sSqlSub = sSqlSub & "   INNER JOIN MAP_CODE AS MAP_CODE       ON     SYENTITY.DOMI_CODE =  MAP_CODE.MAP_CD1       AND    MAP_CODE.MAP_TABLE_CD = 'MAP029' "
                  sSqlSub = sSqlSub & "   WHERE SEC_FEATURE.SECURITY_ID ='" & Trim(rs("SECURITY_ID")) & "'   WITH UR;  "
                              Set rsSub1 = Conn.Execute(sSqlSub)
                              If Not rsSub1.EOF Then
                                    COUNTRY_CD_ISS = Trim(rsSub1("COUNTRY_CD_ISS"))
                              End If
                              Set rsSub1 = Nothing
                              
            '110316 Ton : เพิ่ม Process การแปลงข้อมูลที่อาจมีค่า '
            ISIN_CODE = repQuote(ISIN_CODE)
            DEBT_INS_NAME = repQuote(DEBT_INS_NAME)
            ISSUER_NAME = repQuote(ISSUER_NAME)
            IP_NAME = repQuote(IP_NAME)
            '*****************************************************
                        
            'Generate รายการดอกเบี้ยจ่าย
             'If Trim(rs("PURCHASE_INT_DR_CR_IND")) = "D" Then
             If Trim(rs("BUY_SELL")) = "1" Then

                        sSqlInsert = sSqlInsert & "   VALUES ('999999' ,'" & rs("TRANS_NUM") & "' ,'" & sData_set_date & "' ,'270002' ,'268006' ,  " & TRAN_SEQ & "  ,'' ,'" & sData_set_date & "' ,'' ,'' ,'' "
                        sSqlInsert = sSqlInsert & "  ,'" & rs("CONT_CUR") & "', " & Abs(rs("PURCHASE_INT")) & ",''  , '" & ISIN_CODE & "'   "
                        sSqlInsert = sSqlInsert & "  , '" & DEBT_INS_NAME & "' , '" & ISSUER_NAME & "' , '" & COUNTRY_CD_ISS & "' ,NULL ,NULL ,NULL ,'' ,NULL ,'' ,NULL ,0 ,NULL ,0  "
                        sSqlInsert = sSqlInsert & "  ,Current Date ,Current Time ,'PLDMS' ,'HIPO' ,'0' ,'' ,'' ,''  )"
                        
                        sSqlInsertLog = sSqlInsertLog & "   VALUES ('999999' ,'" & rs("TRANS_NUM") & "' ,'" & sData_set_date & "' ,'270002' ,'268006' ,  " & TRAN_SEQ & "  ,CURRENT TIMESTAMP, 'I'  , '' ,'" & sData_set_date & "' ,'' ,'' ,'' "
                        sSqlInsertLog = sSqlInsertLog & "  ,'" & rs("CONT_CUR") & "', " & Abs(rs("PURCHASE_INT")) & ",''  , '" & ISIN_CODE & "'   "
                        sSqlInsertLog = sSqlInsertLog & "  , '" & DEBT_INS_NAME & "' , '" & ISSUER_NAME & "' , '" & COUNTRY_CD_ISS & "' ,NULL ,NULL ,NULL ,'' ,NULL ,'' ,NULL ,0 ,NULL ,0  "
                        sSqlInsertLog = sSqlInsertLog & "  ,Current Date ,Current Time ,'PLDMS' ,'HIPO' ,'0' ,'' ,'' ,''  )"

                        Conn.Execute (sSqlInsert)
                        Conn.Execute (sSqlInsertLog)

           ElseIf Trim(rs("BUY_SELL")) = "2" Then
           'ElseIf Trim(rs("PURCHASE_INT_DR_CR_IND")) = "C" Then
            'Generate รายการดอกเบี้ยรับ
                        sSqlInsert = sSqlInsert & "   VALUES ('999999' ,'" & rs("TRANS_NUM") & "' ,'" & sData_set_date & "' ,'270001' ,'268006' ,  " & TRAN_SEQ & "  ,'' ,'" & sData_set_date & "' ,'' ,'' ,'' "
                        sSqlInsert = sSqlInsert & "  ,'" & rs("CONT_CUR") & "', " & Abs(rs("PURCHASE_INT")) & ",''  , '" & ISIN_CODE & "'   "
                        sSqlInsert = sSqlInsert & "  , '" & DEBT_INS_NAME & "' , '" & ISSUER_NAME & "' , '" & COUNTRY_CD_ISS & "' ,NULL ,NULL ,NULL ,'' ,NULL ,'' ,NULL ,0 ,NULL ,0  "
                        sSqlInsert = sSqlInsert & "  ,Current Date ,Current Time ,'PLDMS' ,'HIPO' ,'0' ,'' ,'' ,''  )"
                        
                        sSqlInsertLog = sSqlInsertLog & "   VALUES ('999999' ,'" & rs("TRANS_NUM") & "' ,'" & sData_set_date & "' ,'270001' ,'268006' ,  " & TRAN_SEQ & "  ,CURRENT TIMESTAMP, 'I'  , '' ,'" & sData_set_date & "' ,'' ,'' ,'' "
                        sSqlInsertLog = sSqlInsertLog & "  ,'" & rs("CONT_CUR") & "', " & Abs(rs("PURCHASE_INT")) & ",''  , '" & ISIN_CODE & "'   "
                        sSqlInsertLog = sSqlInsertLog & "  , '" & DEBT_INS_NAME & "' , '" & ISSUER_NAME & "' , '" & COUNTRY_CD_ISS & "' ,NULL ,NULL ,NULL ,'' ,NULL ,'' ,NULL ,0 ,NULL ,0  "
                        sSqlInsertLog = sSqlInsertLog & "  ,Current Date ,Current Time ,'PLDMS' ,'HIPO' ,'0' ,'' ,'' ,''  )"
                        
                        Conn.Execute (sSqlInsert)
                        Conn.Execute (sSqlInsertLog)
                        
                        
           End If
           TRAN_SEQ = TRAN_SEQ + 1
           count_insert = count_insert + 1
           
            Set rsSub1 = Nothing
            Set rsSub2 = Nothing
            rs.MoveNext
      
      Loop
      Set rs = Nothing

ErrDB:
      If Err.Number <> 0 Then
            Conn.RollbackTrans
            PTX_FI_COMMON = Err.Description
      Else
            Conn.CommitTrans
            PTX_FI_COMMON = ""
      End If
End Function


Private Function PTX_FI_BACK_DT(iDate) As String
    '110316 Ton : เพิ่ม Process การแปลงข้อมูลที่อาจมีค่า '
      On Error GoTo ErrDB
      Conn.BeginTrans

      'Dim rs, rsSub1, rsSub2, rsBackdate As New ADODB.Recordset
      Dim rs, rsSub1, rsSub2, rsBackdate As Object
      Set rs = CreateObject("ADODB.Recordset")
      Set rsSub1 = CreateObject("ADODB.Recordset")
      Set rsSub2 = CreateObject("ADODB.Recordset")
      Set rsBackdate = CreateObject("ADODB.Recordset")
      
      Dim sSqlSEQ, sSqlSelect, sSqlSub, sHeadSqlInsert, sHeadSqlInsertLog, sSqlInsert, sSqlInsertLog, sSqlBackdate As String
      Dim sSplit, flag_insert, dateInsert, endTxt, BackDate, BackDateInsert As String
      Dim TRAN_SEQ As Integer
      Dim IP_NAME, DEBT_ARR_TYPE_CD, ISIN_CODE, DEBT_INS_NAME, ISSUER_NAME, COUNTRY_CD_ISS, ISSUE_DTE As String
      Dim MATURITY_DTE, sORG_TRM, ORG_TRM, ORG_TRM_U, COUPON_RATE, INT_COUNTRY_CD, IP_COUNTRY_CD As String
      Dim sDate As String

      sDate = iDate
       sSqlBackdate = "  SELECT MAX(DATE) AS BACK_DT FROM HOLIDAY  WHERE HOLIDAY_FLAG = 'N'  AND DATE < '" & sDate & "'    WITH UR; "
      Set rsBackdate = Conn.Execute(sSqlBackdate)
      BackDate = rsBackdate("BACK_DT")
      BackDateInsert = Format(Trim(rsBackdate("BACK_DT")), "YYYY") & "-" & Format(Trim(rsBackdate("BACK_DT")), "MM") & "-" & Format(Trim(rsBackdate("BACK_DT")), "DD")
      
      '-------------------------------------------------------------------- Find SEQ-----------------------------------------------------------------
            sSqlSEQ = " SELECT COALESCE(MAX(TRAN_SEQ), 0)  AS TRAN_SEQ "
            'sSqlSEQ = sSqlSEQ & " FROM DS_PTX WHERE DATA_SET_DATE = ( SELECT MAX(DATE) AS BACK_DT FROM HOLIDAY  WHERE HOLIDAY_FLAG = 'N'  AND DATE < '" & sDate & "'  ) AND DATA_SYSTEM_CD = 'HIPO'    WITH UR; "
            sSqlSEQ = sSqlSEQ & " FROM DS_PTX WHERE DATA_SET_DATE =  '" & Format(BackDate, "YYYY-MM-DD") & "'   AND DATA_SYSTEM_CD = 'HIPO'    WITH UR; "
            
            Set rs = Conn.Execute(sSqlSEQ)
           If Not rs.EOF Then
                  TRAN_SEQ = Trim(rs("TRAN_SEQ")) + 1
            Else
                  TRAN_SEQ = 1
            End If
            Set rs = Nothing

      '-------------------------------------------------------------------- Select Record -----------------------------------------------------------------
      ' ปกติ รายการ back date จะมี B 1 record และ Ac 1 record ใน Trans Date เดียวกัน
    '  sSqlSelect = "  SELECT FI_WAV_B.BOND_ID , FI_WAV_B.TRANS_DATE , FI_WAV_B.TRANS_NUM  ,FI_WAV_B.CURRENCY "
    '  sSqlSelect = sSqlSelect & "  , FI_WAV_AC.CHG_WAV , FI_WAV_AC.PROFIT , FI_WAV_AC.LOSS"
    '  sSqlSelect = sSqlSelect & "  FROM   TMS_FI_WAV  AS FI_WAV_B INNER JOIN TMS_FI_WAV  AS FI_WAV_AC"
    '  sSqlSelect = sSqlSelect & "  ON FI_WAV_B.BOND_ID = FI_WAV_AC.BOND_ID   AND   FI_WAV_B.TRANS_DATE = FI_WAV_AC.TRANS_DATE"
    '  sSqlSelect = sSqlSelect & "  AND     FI_WAV_B.TYPE ='B'    AND     FI_WAV_AC.TYPE ='AC'"
    '  sSqlSelect = sSqlSelect & "  WHERE  FI_WAV_B.TRANS_DATE =( SELECT MAX(DATE) AS BACK_DT FROM HOLIDAY  WHERE HOLIDAY_FLAG = 'N'"
    '  sSqlSelect = sSqlSelect & "  AND DATE < '" & sDate & "'  ) WITH   UR;"
  
     ' ปกติ รายการ back date จะมี B 1 record และ Ac 1 record ใน Trans Date เดียวกัน กรณีที่มีมากกว่า 1 AC มากกว่า 1 รายการ ให้เอารายการที่มี LINE_NO มากสุด
     ' แก้ Query ว่าเอาวันที่ Sdate มา where แทน Back date
     ' may
 '   sSqlSelect = " SELECT    FI_WAV_B.BOND_ID ,    FI_WAV_B.TRANS_DATE    , FI_WAV_B.TRANS_NUM    ,FI_WAV_B.CURRENCY "
 '   sSqlSelect = sSqlSelect & " , FI_WAV_AC.CHG_WAV , FI_WAV_AC.PROFIT , FI_WAV_AC.LOSS"
 '   sSqlSelect = sSqlSelect & "   FROM     TMS_FI_WAV  AS FI_WAV_B "
 '   sSqlSelect = sSqlSelect & "   Inner Join"
 '   sSqlSelect = sSqlSelect & " (SELECT BOND_ID ,  CHG_WAV,PROFIT,LOSS,TRANS_DATE From TMS_FI_WAV"
 '   sSqlSelect = sSqlSelect & "  WHERE TRANS_DATE =   '" & sDate & "'  AND  TYPE = 'AC' AND LINE_NO in"
 '   sSqlSelect = sSqlSelect & "   (  SELECT MAX(LINE_NO)  From TMS_FI_WAV   WHERE TRANS_DATE =   '" & sDate & "'   AND  TYPE = 'AC'    )  ) AS  FI_WAV_AC "
 '   sSqlSelect = sSqlSelect & " ON FI_WAV_AC.BOND_ID = FI_WAV_B.BOND_ID  AND FI_WAV_AC.TRANS_DATE = FI_WAV_B.TRANS_DATE "
 '   sSqlSelect = sSqlSelect & "   WHERE FI_WAV_B.TRANS_DATE =  '" & sDate & "'  AND  FI_WAV_B.TYPE = 'B' "
'may2
sSqlSelect = " SELECT    FI_WAV_B.BOND_ID ,    FI_WAV_B.TRANS_DATE    , FI_WAV_B.TRANS_NUM    ,FI_WAV_B.CURRENCY"
sSqlSelect = sSqlSelect & "  , FI_WAV_AC.CHG_WAV , FI_WAV_AC.PROFIT , FI_WAV_AC.LOSS FROM     TMS_FI_WAV  AS FI_WAV_B "
sSqlSelect = sSqlSelect & "   Inner Join ( SELECT   A.BOND_ID ,  A.CHG_WAV,A.PROFIT,A.LOSS,A.TRANS_DATE From     TMS_FI_WAV    AS A "
sSqlSelect = sSqlSelect & "   Inner Join ( SELECT   BOND_ID , MAX(LINE_NO) From TMS_FI_WAV WHERE    TRANS_DATE =   '" & sDate & "'   AND   TYPE = 'AC'  GROUP BY BOND_ID  ) B "
sSqlSelect = sSqlSelect & "   ON A.BOND_ID = B.BOND_ID  WHERE   A.TRANS_DATE =   '" & sDate & "' AND   A.TYPE = 'AC'  ) AS  FI_WAV_AC "
sSqlSelect = sSqlSelect & "   ON FI_WAV_AC.BOND_ID = FI_WAV_B.BOND_ID  AND FI_WAV_AC.TRANS_DATE = FI_WAV_B.TRANS_DATE"
sSqlSelect = sSqlSelect & "   WHERE FI_WAV_B.TRANS_DATE =  '" & sDate & "'  AND  FI_WAV_B.TYPE = 'B' "

    Set rs = Conn.Execute(sSqlSelect)
   
      If Not rs.EOF Then
      Do While Not rs.EOF
                  flag_insert = ""
            IP_NAME = ""
            '------------------------------------------------------------------
            DEBT_ARR_TYPE_CD = ""
            ISIN_CODE = ""
            DEBT_INS_NAME = ""
            ISSUER_NAME = ""
            COUNTRY_CD_ISS = ""
            ISSUE_DTE = "null"
            MATURITY_DTE = "null"
            sORG_TRM = ""
            sSplit = ""
            ORG_TRM = "null"
            ORG_TRM_U = ""
            COUPON_RATE = "null"
            INT_COUNTRY_CD = ""
            '------------------------------------------------------------------
            '---------Select Group 1 ISIN_CODE,DEBT_INS_NAME,ISSUE_DTE,MATURITY_DTE,COUPON_RATE
                              sSqlSub = "  SELECT  SEC_FEATURE.ISIN_CODE AS ISIN_CODE ,  SEC_FEATURE.INST_DESC AS DEBT_INS_NAME, SEC_FEATURE.ISSUE_DATE AS ISSUE_DTE  "
                              sSqlSub = sSqlSub & "    ,SEC_FEATURE.MAT_DATE AS MATURITY_DTE ,SEC_FEATURE.COUPON_RATE AS COUPON_RATE, "
                              sSqlSub = sSqlSub & "     MONTH(SEC_FEATURE.MAT_DATE -  SEC_FEATURE.ISSUE_DATE ) AS DIFF_MONTH ,YEAR(SEC_FEATURE.MAT_DATE -  SEC_FEATURE.ISSUE_DATE ) AS DIFF_YEAR  "
                              sSqlSub = sSqlSub & "    FROM ESL_SEC_FEATURE AS SEC_FEATURE  WHERE SEC_FEATURE.SECURITY_ID ='" & Trim(rs("BOND_ID")) & "'   WITH UR;   "
                               
                                                            
                              Set rsSub1 = Conn.Execute(sSqlSub)
                              If Not rsSub1.EOF Then
                                          'MATURITY_DTE , ISSUE_DTE
                                          sORG_TRM = Cal_ORG_TRM(Trim(rsSub1("MATURITY_DTE")), Trim(rsSub1("ISSUE_DTE")))
                                          sSplit = Split(sORG_TRM, ";")
                                          ORG_TRM = sSplit(0)
                                          ORG_TRM_U = sSplit(1)
                                          'ISIN_CODE = Trim(rsSub1("ISIN_CODE"))
                                          If Not IsNull(rsSub1("ISIN_CODE")) Then
                                              ISIN_CODE = Replace(Trim(rsSub1("ISIN_CODE")), "&", " AND ")
                                          End If
                                          'DEBT_INS_NAME = Trim(rsSub1("DEBT_INS_NAME"))
                                          If Not IsNull(rsSub1("DEBT_INS_NAME")) Then
                                              DEBT_INS_NAME = Replace(Trim(rsSub1("DEBT_INS_NAME")), "&", " AND ")
                                          End If
                                          If IsDate(Trim(rsSub1("ISSUE_DTE"))) = True Then
                                                ISSUE_DTE = "'" & Format(Trim(rsSub1("ISSUE_DTE")), "YYYY") & "-" & Format(Trim(rsSub1("ISSUE_DTE")), "MM") & "-" & Format(Trim(rsSub1("ISSUE_DTE")), "DD") & "'"
                                          End If
                                          If IsDate(Trim(rsSub1("MATURITY_DTE"))) = True Then
                                                MATURITY_DTE = "'" & Format(Trim(rsSub1("MATURITY_DTE")), "YYYY") & "-" & Format(Trim(rsSub1("MATURITY_DTE")), "MM") & "-" & Format(Trim(rsSub1("MATURITY_DTE")), "DD") & "'"
                                          End If
                                          COUPON_RATE = Trim(rsSub1("COUPON_RATE"))
'                                         If Trim(rsSub1("DIFF_YEAR")) >= "1" Then
'                                                      ORG_TRM = Trim(rsSub1("DIFF_YEAR"))
'                                                      ORG_TRM_U = "Y"
'                                          Else
'                                                      ORG_TRM = Trim(rsSub1("DIFF_MONTH"))
'                                                      ORG_TRM_U = "M"
'                                          End If
                              End If
                              Set rsSub1 = Nothing
                              
            '---------Select Group 2 ISSUER_NAME
                              sSqlSub = "  SELECT SYENTITY.LONG_NAME  AS ISSUER_NAME  FROM ESL_SEC_FEATURE AS SEC_FEATURE "
                              sSqlSub = sSqlSub & "  INNER JOIN ESL_SYENTITY AS SYENTITY       ON     SEC_FEATURE.ENTITY_CODE = SYENTITY.ENTITY_CODE "
                              sSqlSub = sSqlSub & "   WHERE SEC_FEATURE.SECURITY_ID ='" & Trim(rs("BOND_ID")) & "'   WITH UR;  "
                              Set rsSub1 = Conn.Execute(sSqlSub)
                              If Not rsSub1.EOF Then
                                    If Not rsSub1.EOF And Not IsNull(rsSub1("ISSUER_NAME")) Then
                                        'ISSUER_NAME = Trim(rsSub1("ISSUER_NAME"))
                                        ISSUER_NAME = Replace(Trim(rsSub1("ISSUER_NAME")), "&", " AND ")
                                    End If
                              End If
                              Set rsSub1 = Nothing
                              
            '---------Select Group 3 COUNTRY_CD_ISS
                  sSqlSub = "   SELECT MAP_CODE.MAP_CD2 AS COUNTRY_CD_ISS     FROM ESL_SEC_FEATURE AS SEC_FEATURE"
                  sSqlSub = sSqlSub & "   INNER JOIN ESL_SYENTITY AS SYENTITY       ON     SEC_FEATURE.ENTITY_CODE = SYENTITY.ENTITY_CODE"
                  sSqlSub = sSqlSub & "   INNER JOIN MAP_CODE AS MAP_CODE       ON     SYENTITY.DOMI_CODE =  MAP_CODE.MAP_CD1       AND    MAP_CODE.MAP_TABLE_CD = 'MAP029' "
                  sSqlSub = sSqlSub & "   WHERE SEC_FEATURE.SECURITY_ID ='" & Trim(rs("BOND_ID")) & "'   WITH UR;  "
                              Set rsSub1 = Conn.Execute(sSqlSub)
                              If Not rsSub1.EOF Then
                                    COUNTRY_CD_ISS = Trim(rsSub1("COUNTRY_CD_ISS"))
                              End If
                              Set rsSub1 = Nothing
                              
            '---------Select Group 4 INT_COUNTRY_CD
                  sSqlSub = "   SELECT MAP_CODE3.MAP_CD2 AS INT_COUNTRY_CD  FROM ESL_SEC_FEATURE AS SEC_FEATURE"
                  sSqlSub = sSqlSub & "   INNER JOIN MAP_CODE AS MAP_CODE3      ON     SEC_FEATURE.CURRENCY =  MAP_CODE3.MAP_CD1"
                  sSqlSub = sSqlSub & "   AND    MAP_CODE3.MAP_TABLE_CD = 'MAP030'"
                  sSqlSub = sSqlSub & "   WHERE SEC_FEATURE.SECURITY_ID ='" & Trim(rs("BOND_ID")) & "'   WITH UR;  "
                              
                              Set rsSub1 = Conn.Execute(sSqlSub)
                              If Not rsSub1.EOF Then
                                    INT_COUNTRY_CD = Trim(rsSub1("INT_COUNTRY_CD"))
                              End If
                              Set rsSub1 = Nothing
                              
            '---------Select Group 5 DEBT_ARR_TYPE_CD
                  sSqlSub = "  SELECT MAP_CODE1.MAP_CD2 AS DEBT_ARR_TYPE_CD  FROM ESL_SEC_FEATURE AS SEC_FEATURE"
                  sSqlSub = sSqlSub & "   INNER JOIN MAP_CODE AS MAP_CODE1   ON     SEC_FEATURE.BOND_TYPE = MAP_CODE1.MAP_CD1      AND    MAP_CODE1.MAP_TABLE_CD = 'MAP028'"
                  sSqlSub = sSqlSub & "   WHERE SEC_FEATURE.SECURITY_ID ='" & Trim(rs("BOND_ID")) & "'   WITH UR;  "
                              Set rsSub1 = Conn.Execute(sSqlSub)
                              If Not rsSub1.EOF Then
                                    DEBT_ARR_TYPE_CD = Trim(rsSub1("DEBT_ARR_TYPE_CD"))
                              End If
                              Set rsSub1 = Nothing


            'DATE
            dateInsert = Format(Trim(rs("TRANS_DATE")), "YYYY") & "-" & Format(Trim(rs("TRANS_DATE")), "MM") & "-" & Format(Trim(rs("TRANS_DATE")), "DD")

            'TRAN_SEQ = TRAN_SEQ + 1
            sHeadSqlInsert = " INSERT INTO DS_PTX ( ARR_TYPE_CD, ARR_NUMBER, TRAN_DTE, TRAN_TYPE_CD, ITEM_TYPE_CD, TRAN_SEQ, REC_PAY_ITEM_DSC, DATA_SET_DATE, PROVIDER_BR_NO, IBF_IND_CD, PAY_METHOD_CD, CURR_CD, TRAN_AMT_CURR, DEBT_ARR_TYPE_CD, ISIN_CODE, DEBT_INS_NAME, ISSUER_NAME, COUNTRY_CD_ISS, ISSUE_DTE, MATURITY_DTE, ORG_TRM, ORG_TRM_U, COUPON_RATE, INT_COUNTRY_CD, UNIT_OF_TRAN, SELL_SEC_AMT, PURCHASE_DTE, OS_AMT, LU_DTE, LU_TIME, USER_LU, DATA_SYSTEM_CD, FLAG_ON_OFF, PART_ID, IP_NAME, IP_COUNTRY_CD) "
            sHeadSqlInsertLog = " INSERT INTO DS_PTX_LOG ( ARR_TYPE_CD, ARR_NUMBER, TRAN_DTE, TRAN_TYPE_CD, ITEM_TYPE_CD, TRAN_SEQ, ACTION_TIMESTAMP, ACTION_FLAG, REC_PAY_ITEM_DSC, DATA_SET_DATE, PROVIDER_BR_NO, IBF_IND_CD, PAY_METHOD_CD, CURR_CD, TRAN_AMT_CURR, DEBT_ARR_TYPE_CD, ISIN_CODE, DEBT_INS_NAME, ISSUER_NAME, COUNTRY_CD_ISS, ISSUE_DTE, MATURITY_DTE, ORG_TRM, ORG_TRM_U, COUPON_RATE, INT_COUNTRY_CD, UNIT_OF_TRAN, SELL_SEC_AMT, PURCHASE_DTE, OS_AMT, LU_DTE, LU_TIME, USER_LU, DATA_SYSTEM_CD, FLAG_ON_OFF, PART_ID, IP_NAME, IP_COUNTRY_CD) "
            
'                                    sSqlSub = " SELECT  MAP_CODE1.MAP_CD2 AS DEBT_ARR_TYPE_CD , SEC_FEATURE.ISIN_CODE AS ISIN_CODE ,  SEC_FEATURE.INST_TYPE AS DEBT_INS_NAME ,  SYENTITY.ENTITY_20CODE AS ISSUER_NAME , MAP_CODE2.MAP_CD2 AS COUNTRY_CD_ISS , SEC_FEATURE.ISSUE_DATE AS ISSUE_DTE ,SEC_FEATURE.MAT_DATE AS MATURITY_DTE ,SEC_FEATURE.COUPON_RATE AS COUPON_RATE , MAP_CODE3.MAP_CD2 AS INT_COUNTRY_CD ,  "
'                                    sSqlSub = sSqlSub & " SEC_FEATURE.*, SYENTITY.* ,MAP_CODE1.*,MAP_CODE2.*,MAP_CODE3.*  , MONTH(SEC_FEATURE.MAT_DATE -  SEC_FEATURE.ISSUE_DATE ) AS DIFF_MONTH ,YEAR(SEC_FEATURE.MAT_DATE -  SEC_FEATURE.ISSUE_DATE ) AS DIFF_YEAR  FROM ESL_SEC_FEATURE AS SEC_FEATURE "
'                                    sSqlSub = sSqlSub & " INNER JOIN ESL_SYENTITY AS SYENTITY ON SEC_FEATURE.ENTITY_CODE = SYENTITY.ENTITY_CODE"
'                                    sSqlSub = sSqlSub & " INNER JOIN MAP_CODE AS MAP_CODE1 ON SEC_FEATURE.BOND_TYPE = MAP_CODE1.MAP_CD1 AND MAP_CODE1.MAP_TABLE_CD = 'MAP028'"
'                                    sSqlSub = sSqlSub & " INNER JOIN MAP_CODE AS MAP_CODE2 ON SYENTITY.DOMI_CODE =  MAP_CODE2.MAP_CD1 AND MAP_CODE2.MAP_TABLE_CD = 'MAP029'"
'                                    sSqlSub = sSqlSub & " INNER JOIN MAP_CODE AS MAP_CODE3 ON SEC_FEATURE.CURRENCY =  MAP_CODE3.MAP_CD1 AND MAP_CODE3.MAP_TABLE_CD = 'MAP030'"
'                                    sSqlSub = sSqlSub & " WHERE SEC_FEATURE.SECURITY_ID ='" & Trim(rs("BOND_ID")) & "'   WITH UR;  "
'
'                                    Set rsSub1 = Conn.Execute(sSqlSub)
'
'                                    If Not rsSub1.EOF Then
'                                                DEBT_ARR_TYPE_CD = Trim(rsSub1("DEBT_ARR_TYPE_CD"))
'                                                ISIN_CODE = Trim(rsSub1("ISIN_CODE"))
'                                                DEBT_INS_NAME = Trim(rsSub1("DEBT_INS_NAME"))
'                                                ISSUER_NAME = Trim(rsSub1("ISSUER_NAME"))
'                                                COUNTRY_CD_ISS = Trim(rsSub1("COUNTRY_CD_ISS"))
'                                                If Trim(rsSub1("ISSUE_DTE")) <> "" Then
'                                                      ISSUE_DTE = "'" & Trim(rsSub1("ISSUE_DTE")) & "'"
'                                                End If
'                                                If Trim(rsSub1("MATURITY_DTE")) <> "" Then
'                                                      MATURITY_DTE = "'" & Trim(rsSub1("MATURITY_DTE")) & "'"
'                                                End If
'                                                 If Trim(rsSub1("DIFF_YEAR")) >= "1" Then
'                                                            ORG_TRM = Trim(rsSub1("DIFF_YEAR"))
'                                                            ORG_TRM_U = "Y"
'                                                Else
'                                                            ORG_TRM = Trim(rsSub1("DIFF_MONTH"))
'                                                            ORG_TRM_U = "M"
'                                                End If
'                                                COUPON_RATE = Trim(rsSub1("COUPON_RATE"))
'                                                INT_COUNTRY_CD = Trim(rsSub1("INT_COUNTRY_CD"))
'                                    End If
                              
                                    If Trim(rs("TRANS_NUM")) <> "" Then
                                    
                                          sSqlSub = " SELECT * FROM ESL_COMMON AS COMMON "
                                          sSqlSub = sSqlSub & " INNER JOIN  ESL_SYENTITY AS SYENTITY ON COMMON.CNTR_PARTY = SYENTITY.ENTITY_CODE"
                                          sSqlSub = sSqlSub & "  INNER JOIN MAP_CODE AS MAP_CODE ON SYENTITY.DOMI_CODE =  MAP_CODE.MAP_CD1 AND MAP_CODE.MAP_TABLE_CD = 'MAP029' "
                                          'sSqlSub = sSqlSub & " WHERE COMMON.TRANS_NUM ='" & Trim(rs("TRANS_NUM")) & "'   WITH UR; "
                                          'SORN
                                         sSqlSub = sSqlSub & " WHERE COMMON.TRANS_NUM ='" & Trim(rs("TRANS_NUM")) & "' and   COMMON.AS_OF_DT =  '" & sDate & "'   WITH UR; "
                                          
                                          Set rsSub2 = Conn.Execute(sSqlSub)
                                          If Not rsSub2.EOF Then
                                                If Not IsNull(Trim(rsSub2("REMARK"))) Then
                                                     endTxt = InStr(1, Trim(rsSub2("REMARK")), "/CF")
                                                      If endTxt <> "0" Then
                                                            IP_NAME = Mid(Trim(rsSub2("REMARK")), 1, endTxt)
                                                            IP_NAME = Replace(IP_NAME, "&", " AND ")
                                                            
'                                                      Else
'                                                            IP_NAME = Trim(rsSub2("ENTITY_20CODE"))
'                                                            IP_NAME = Replace(IP_NAME, "&", " AND ")
                                                      End If
                                                End If
                                                      'sorn
                                                      IP_COUNTRY_CD = Trim(rsSub2("MAP_CD2"))
                                                      '20110422 yos add IP_NAME
                                                       If Not IsNull(Trim(rsSub2("ENTITY_20CODE"))) Then
                                                            IP_NAME = Trim(rsSub2("ENTITY_20CODE"))
                                                            IP_NAME = Replace(IP_NAME, "&", " AND ")
                                                       End If
                                          Else
                                                IP_NAME = ""
                                                'sorn
                                                IP_COUNTRY_CD = ""
                                          End If
                                                                        
                                    Else
                                          sSqlSub = " SELECT * FROM ESL_SEC_FEATURE AS SEC_FEATURE"
                                          sSqlSub = sSqlSub & " INNER JOIN  ESL_SYENTITY AS SYENTITY ON SEC_FEATURE.ENTITY_CODE = SYENTITY.ENTITY_CODE"
                                          sSqlSub = sSqlSub & "  INNER JOIN MAP_CODE AS MAP_CODE ON SYENTITY.DOMI_CODE =  MAP_CODE.MAP_CD1 AND MAP_CODE.MAP_TABLE_CD = 'MAP029' "
                                          sSqlSub = sSqlSub & " WHERE SEC_FEATURE.SECURITY_ID ='" & Trim(rs("BOND_ID")) & "'   WITH UR; "
                                          
                                          Set rsSub2 = Conn.Execute(sSqlSub)
                                          If Not rsSub2.EOF Then
                                                If Not IsNull(Trim(rsSub2("ENTITY_20CODE"))) Then
                                                      IP_NAME = Trim(rsSub2("ENTITY_20CODE"))
                                                      IP_NAME = Replace(IP_NAME, "&", " AND ")
                                                End If
                                                IP_COUNTRY_CD = Trim(rsSub2("MAP_CD2"))
                                          Else
                                                IP_NAME = ""
                                                IP_COUNTRY_CD = ""
                                          End If
                                          
                                    End If
    
            '110316 Ton : เพิ่ม Process การแปลงข้อมูลที่อาจมีค่า '
            ISIN_CODE = repQuote(ISIN_CODE)
            DEBT_INS_NAME = repQuote(DEBT_INS_NAME)
            ISSUER_NAME = repQuote(ISSUER_NAME)
            IP_NAME = repQuote(IP_NAME)
            '*****************************************************
            If Trim(rs("CHG_WAV")) <> "0" Then
                  If Trim(rs("CHG_WAV")) < "0" Then
                        'CHG_WAV < 0
                        ' Change Insert by dataInsert to BackDateInsert
                        sSqlInsert = sHeadSqlInsert & "  VALUES('999999' , '" & Trim(rs("TRANS_NUM")) & "' , '" & BackDateInsert & "' , '270003' , '268027' ,  " & TRAN_SEQ & "   "
                        sSqlInsert = sSqlInsert & "  , '' , '" & BackDateInsert & "' , '' , '' , '234005' , '" & Trim(rs("CURRENCY")) & "' ,  " & Abs(rs("CHG_WAV")) & " "
                        sSqlInsert = sSqlInsert & "  , '" & DEBT_ARR_TYPE_CD & "' , '" & ISIN_CODE & "' , '" & DEBT_INS_NAME & "' , '" & ISSUER_NAME & "'  "
                        sSqlInsert = sSqlInsert & "  , '" & COUNTRY_CD_ISS & "'  , " & ISSUE_DTE & " , " & MATURITY_DTE & "  , " & ORG_TRM & "  , '" & ORG_TRM_U & "' "
                        sSqlInsert = sSqlInsert & "  ,  " & COUPON_RATE & "  ,'" & INT_COUNTRY_CD & "' , NULL , 0 , NULL , 0  , CURRENT DATE , CURRENT TIME "
                        sSqlInsert = sSqlInsert & "  , 'PLDMS'  , 'HIPO'  , '0' , '' , '" & IP_NAME & "' , '" & IP_COUNTRY_CD & "');  "
                        
                        
                        ' Change Insert by dataInsert to BackDateInsert
                         sSqlInsertLog = sHeadSqlInsertLog & " VALUES( '999999' , '" & Trim(rs("TRANS_NUM")) & "' , '" & BackDateInsert & "' , '270003' , '268027' ,  " & TRAN_SEQ & "  , CURRENT TIMESTAMP, 'I'  "
                         sSqlInsertLog = sSqlInsertLog & "  , '' , '" & BackDateInsert & "' , '' , '' , '234005' , '" & Trim(rs("CURRENCY")) & "' ,  " & Abs(rs("CHG_WAV")) & "  "
                         sSqlInsertLog = sSqlInsertLog & "  , '" & DEBT_ARR_TYPE_CD & "' , '" & ISIN_CODE & "' , '" & DEBT_INS_NAME & "' , '" & ISSUER_NAME & "'  "
                         sSqlInsertLog = sSqlInsertLog & "  , '" & COUNTRY_CD_ISS & "'  , " & ISSUE_DTE & " , " & MATURITY_DTE & "  , " & ORG_TRM & "  , '" & ORG_TRM_U & "' "
                         sSqlInsertLog = sSqlInsertLog & "  ,  " & COUPON_RATE & "  ,'" & INT_COUNTRY_CD & "' , NULL , 0 , NULL , 0  , CURRENT DATE , CURRENT TIME "
                         sSqlInsertLog = sSqlInsertLog & "  , 'PLDMS'  , 'HIPO'  , '0' , '' , '" & IP_NAME & "' , '" & IP_COUNTRY_CD & "');  "

                        Conn.Execute (sSqlInsert)
                        Conn.Execute (sSqlInsertLog)

                        
                  Else
                        'CHG_WAV > 0
                        ' Change Insert by dataInsert to BackDateInsert
                        sSqlInsert = sHeadSqlInsert & " VALUES ('999999', '" & Trim(rs("TRANS_NUM")) & "', '" & BackDateInsert & "', '270003', '268022', " & TRAN_SEQ & " "
                        sSqlInsert = sSqlInsert & " , '', '" & BackDateInsert & "', '', '', '234005', '" & Trim(rs("CURRENCY")) & "', " & Abs(rs("CHG_WAV")) & ""
                        sSqlInsert = sSqlInsert & " , '" & DEBT_ARR_TYPE_CD & "', '" & ISIN_CODE & "', '" & DEBT_INS_NAME & "', '" & ISSUER_NAME & "' "
                        sSqlInsert = sSqlInsert & " , '" & COUNTRY_CD_ISS & "' , " & ISSUE_DTE & ", " & MATURITY_DTE & " "
                        sSqlInsert = sSqlInsert & " ," & ORG_TRM & ", '" & ORG_TRM_U & "', " & COUPON_RATE & " ,'" & INT_COUNTRY_CD & "', NULL , 0, NULL "
                        sSqlInsert = sSqlInsert & " , 0 , current date, current time, 'PLDMS' , 'HIPO' , '0', '',  '" & IP_NAME & "' , '" & IP_COUNTRY_CD & "'); "
                        ' Change Insert by dataInsert to BackDateInsert
                        sSqlInsertLog = sHeadSqlInsertLog & " VALUES ('999999', '" & Trim(rs("TRANS_NUM")) & "', '" & BackDateInsert & "', '270003', '268022', " & TRAN_SEQ & " , CURRENT TIMESTAMP, 'I'  "
                        sSqlInsertLog = sSqlInsertLog & " , '', '" & BackDateInsert & "', '', '', '234005', '" & Trim(rs("CURRENCY")) & "', " & Abs(rs("CHG_WAV")) & ""
                        sSqlInsertLog = sSqlInsertLog & " , '" & DEBT_ARR_TYPE_CD & "', '" & ISIN_CODE & "', '" & DEBT_INS_NAME & "', '" & ISSUER_NAME & "' "
                        sSqlInsertLog = sSqlInsert & " , '" & COUNTRY_CD_ISS & "' , " & ISSUE_DTE & ", " & MATURITY_DTE & " "
                        sSqlInsertLog = sSqlInsertLog & " ," & ORG_TRM & ", '" & ORG_TRM_U & "', " & COUPON_RATE & " ,'" & INT_COUNTRY_CD & "', NULL, 0, NULL  "
                        sSqlInsertLog = sSqlInsertLog & " , 0 , current date, current time, 'PLDMS' , 'HIPO' , '0', '',  '" & IP_NAME & "' , '" & IP_COUNTRY_CD & "'); "
                        
                        Conn.Execute (sSqlInsert)
                        Conn.Execute (sSqlInsertLog)
                  End If
                  TRAN_SEQ = TRAN_SEQ + 1
                  count_insert = count_insert + 1
            End If
            
            If Trim(rs("PROFIT")) <> "0" Then
                        'PROFIT
                        ' Change Insert by dataInsert to BackDateInsert
                        sSqlInsert = sHeadSqlInsert & " VALUES ('999999', '" & Trim(rs("TRANS_NUM")) & "', '" & BackDateInsert & "', '270001', '268017', " & TRAN_SEQ & "  "
                        sSqlInsert = sSqlInsert & " , 'GAIN', '" & BackDateInsert & "', '', '', '', '" & Trim(rs("CURRENCY")) & "', " & Abs(rs("PROFIT")) & ""
                        sSqlInsert = sSqlInsert & " ,'', '" & ISIN_CODE & "', '" & DEBT_INS_NAME & "', '" & ISSUER_NAME & "' "
                        sSqlInsert = sSqlInsert & " , '" & COUNTRY_CD_ISS & "' , NULL ,NULL "
                        sSqlInsert = sSqlInsert & " ,NULL , '',NULL  ,''   "
                        sSqlInsert = sSqlInsert & " , NULL , 0 , NULL , 0 , current date, current time, 'PLDMS' , 'HIPO' , '0', '',  '" & IP_NAME & "' , '" & IP_COUNTRY_CD & "'); "
                         ' Change Insert by dataInsert to BackDateInsert
                        sSqlInsertLog = sHeadSqlInsertLog & " VALUES ('999999', '" & Trim(rs("TRANS_NUM")) & "', '" & BackDateInsert & "', '270001', '268017', " & TRAN_SEQ & "  , CURRENT TIMESTAMP, 'I'  "
                        sSqlInsertLog = sSqlInsertLog & " , 'GAIN', '" & BackDateInsert & "', '', '', '', '" & Trim(rs("CURRENCY")) & "', " & Abs(rs("PROFIT")) & ""
                        sSqlInsertLog = sSqlInsertLog & " ,'', '" & ISIN_CODE & "', '" & DEBT_INS_NAME & "', '" & ISSUER_NAME & "' "
                        sSqlInsertLog = sSqlInsertLog & " , '" & COUNTRY_CD_ISS & "' , NULL ,NULL "
                        sSqlInsertLog = sSqlInsertLog & " ,NULL , '',NULL  ,''  "
                        sSqlInsertLog = sSqlInsertLog & " , NULL , 0 , NULL , 0 , current date, current time, 'PLDMS' , 'HIPO' , '0', '',  '" & IP_NAME & "' , '" & IP_COUNTRY_CD & "'); "
                        
                        Conn.Execute (sSqlInsert)
                        Conn.Execute (sSqlInsertLog)
                        'SORN
                         count_insert = count_insert + 1
                        TRAN_SEQ = TRAN_SEQ + 1
            End If
            
             If Trim(rs("LOSS")) <> "0" Then
                        'LOSS
                        ' Change Insert by dataInsert to BackDateInsert
                        sSqlInsert = sHeadSqlInsert & " VALUES ('999999', '" & Trim(rs("TRANS_NUM")) & "', '" & BackDateInsert & "', '270002', '268017', " & TRAN_SEQ & " "
                        sSqlInsert = sSqlInsert & " , 'LOSS', '" & BackDateInsert & "', '', '', '', '" & Trim(rs("CURRENCY")) & "', " & Abs(rs("LOSS")) & ""
                        sSqlInsert = sSqlInsert & " ,'', '" & ISIN_CODE & "', '" & DEBT_INS_NAME & "', '" & ISSUER_NAME & "' "
                        sSqlInsert = sSqlInsert & " , '" & COUNTRY_CD_ISS & "' , NULL ,NULL "
                        sSqlInsert = sSqlInsert & " ,NULL , '',NULL  ,'' "
                        sSqlInsert = sSqlInsert & " , NULL , 0 , NULL , 0 , current date, current time, 'PLDMS' , 'HIPO' , '0', '',  '" & IP_NAME & "', '" & IP_COUNTRY_CD & "'); "
                        ' Change Insert by dataInsert to BackDateInsert
                        sSqlInsertLog = sHeadSqlInsertLog & " VALUES ('999999', '" & Trim(rs("TRANS_NUM")) & "', '" & BackDateInsert & "', '270002', '268017', " & TRAN_SEQ & "  , CURRENT TIMESTAMP, 'I'  "
                        sSqlInsertLog = sSqlInsertLog & " , 'LOSS', '" & BackDateInsert & "', '', '', '', '" & Trim(rs("CURRENCY")) & "', " & Abs(rs("LOSS")) & ""
                        sSqlInsertLog = sSqlInsertLog & " ,'', '" & ISIN_CODE & "', '" & DEBT_INS_NAME & "', '" & ISSUER_NAME & "' "
                        sSqlInsertLog = sSqlInsertLog & " , '" & COUNTRY_CD_ISS & "' , NULL ,NULL "
                        sSqlInsertLog = sSqlInsertLog & " ,NULL , '',NULL  ,'' "
                        sSqlInsertLog = sSqlInsertLog & " , NULL , 0 , NULL , 0 , current date, current time, 'PLDMS' , 'HIPO' , '0', '', '" & IP_NAME & "'  , '" & IP_COUNTRY_CD & "'); "
                        
                        Conn.Execute (sSqlInsert)
                        Conn.Execute (sSqlInsertLog)
                        'SORN
                         count_insert = count_insert + 1
                        TRAN_SEQ = TRAN_SEQ + 1
            End If
            Set rsSub1 = Nothing
            Set rsSub2 = Nothing
            rs.MoveNext
      Loop
      End If
      Set rs = Nothing
ErrDB:
      If Err.Number <> 0 Then
            Conn.RollbackTrans
            PTX_FI_BACK_DT = Err.Description
      Else
            Conn.CommitTrans
            PTX_FI_BACK_DT = ""
      End If
End Function

Private Function Cal_ORG_TRM(sMATURITY_DTE, sISSUE_DTE) As String
      Dim sORG, sORG_U As String
      Dim dMonth, mMonth As Double
      If Not (IsDate(sMATURITY_DTE) And IsDate(sISSUE_DTE)) Then
            Cal_ORG_TRM = ""
            Exit Function
      End If
    
    'หาจำนวนเดือน
    dMonth = DateDiff("D", sISSUE_DTE, sMATURITY_DTE)
    mMonth = DateDiff("M", sISSUE_DTE, sMATURITY_DTE)
    
        'หาจำนวนวันที่เหลือ กรณีบวกเดือนที่หามาได้แล้วยังน้อยกว่า Enddate ถือว่าเหลือจำนวนวันทีเกินเดือน ให้เพิ่มค่าเดือนอีก 1
    If mMonth >= 12 Then
            sORG = mMonth / 12
            If Fix(sORG) < (mMonth / 12) Then
                  sORG = sORG + 1
            End If
            sORG = sORG & ";Y"
    Else
            sORG = mMonth
            If DateAdd("M", mMonth, sISSUE_DTE) < DateValue(sMATURITY_DTE) Then
                  sORG = mMonth + 1
              End If
              sORG = sORG & ";M"
    End If
    Cal_ORG_TRM = sORG
    
End Function

Private Function ConnectDB(ConfigPath) As Boolean
    On Error GoTo ErrFunc
    
    Dim iFile
    'Dim ObjEnCode As New SCrypt.clsRijndael
    Dim ObjEnCode As Object
    Dim StrOut1 As String
    Dim StrOut2 As String
    Dim StrOut3 As String
    Dim P1 As String
    Dim P2 As String
    Dim P3 As String
    
    Set Conn = CreateObject("ADODB.Connection")
    Set ObjEnCode = CreateObject("SCrypt.clsRijndael")
    
    iFile = FreeFile
    Open ConfigPath For Input Access Read As #iFile
    
    Input #iFile, StrOut1
    P1 = ObjEnCode.DecryptString(StrOut1, "DMS-BOT", True) ' DSN
    
    Input #iFile, StrOut2
    P2 = ObjEnCode.DecryptString(StrOut2, "DMS-BOT", True) ' user
    
    Input #iFile, StrOut3
    P3 = ObjEnCode.DecryptString(StrOut3, "DMS-BOT", True) ' Password
    
    If Conn.State = 1 Then
        Conn.Close
    End If
    
    Conn.CommandTimeout = 0
    Conn.ConnectionTimeout = 0
    Conn.Open P1, P2, P3
            
ErrFunc:
    If Err.Number <> 0 Then
        ConnectDB = False
    Else
        ConnectDB = True
    End If
End Function

Private Function writeTextFile(sData As String, createNew As Boolean)
    Dim fs As New FileSystemObject
    Dim ts As TextStream
    Dim outputFileName As String
    Dim openTextMode As Integer
    
    outputFileName = App.Path & "\INSERT_PTX_FI.txt"
    If createNew Then
        openTextMode = ForWriting
    Else
        openTextMode = ForAppending
    End If
    
    Set ts = fs.OpenTextFile(outputFileName, openTextMode, True)
    ts.Write sData
    ts.Close
    
    Set fs = Nothing
End Function

Private Function repQuote(sData) As String
    If IsNull(sData) Then
        repQuote = ""
    Else
        repQuote = Replace(sData, "'", "''")
    End If
End Function

Private Function PTX_FI_DAILY(iDate) As String
      Err.Clear
      On Error GoTo ErrDB
      Conn.BeginTrans

      'Dim Conn As New ADODB.Connection
      'Dim rs, rsSub1, rsSub2 As New ADODB.Recordset
      Dim rs, rsSub1, rsSub2 As Object
      Set rs = CreateObject("ADODB.Recordset")
      Set rsSub1 = CreateObject("ADODB.Recordset")
      Set rsSub2 = CreateObject("ADODB.Recordset")
      
      Dim sDate, dateInsert, flag_insert As String
      Dim sSqlSEQ, sSqlSelect, sSqlSub, sSqlInsert, sSqlInsertLog As String
      Dim TRAN_SEQ As Integer
      Dim ISIN_CODE, DEBT_INS_NAME, ISSUER_NAME, COUNTRY_CD_ISS, TRAN_TYPE_CD As String
      Dim INST_DESC, LONG_NAME, MAP_CD2 As String
      
      sDate = iDate
                  '-------------------------------------------------------------------- Find SEQ-----------------------------------------------------------------
                         sSqlSEQ = " SELECT COALESCE(MAX(TRAN_SEQ), 0)  AS TRAN_SEQ "
                         sSqlSEQ = sSqlSEQ & " FROM DS_PTX WHERE DATA_SET_DATE =  '" & sDate & "' AND DATA_SYSTEM_CD = 'HIPO'    WITH UR; "

                         Set rs = Conn.Execute(sSqlSEQ)
                        If Not rs.EOF Then
                               TRAN_SEQ = Trim(rs("TRAN_SEQ")) + 1
                         Else
                               TRAN_SEQ = 0
                         End If
                         Set rs = Nothing
                        '-------------------------------------------------------------------- Select Record -----------------------------------------------------------------
           '             sSqlSelect = "   SELECT  * FROM ESL_DAILY_PL "
           '             sSqlSelect = sSqlSelect & "    WHERE PL_CUR <> 'THB' AND PL_CATEGORY_TYPE = 7  AND PL_METRIC_TYPE = 24   AND TRANS_DT = '" & sDate & "'     WITH UR;  "
           '             Set rs = Conn.Execute(sSqlSelect)
                        sSqlSelect = " Select S.* , C.SECURITY_ID AS SECURITY_ID FROM ESL_SETTLEMENT AS S INNER JOIN ESL_COMMON AS C ON  "
                        sSqlSelect = sSqlSelect & " S.TRANS_NUM = C.TRANS_NUM  And C.AS_OF_DT = '" & sDate & "'  And C.PROD_GROUP = 'FI' And C.GL_PROD_CD <> '0181'  "
                        sSqlSelect = sSqlSelect & " Where S.TRANS_DT =  '" & sDate & "'  and  S.Payment_CUR <> 'THB'  AND S.Payment_Type = 501 With UR ;"
                        Set rs = Conn.Execute(sSqlSelect)
                        
                        If Not rs.EOF Then
                        Do While Not rs.EOF

                              flag_insert = ""
                              '------------------------------------------------------------------
                              ISIN_CODE = ""
                              DEBT_INS_NAME = ""
                              ISSUER_NAME = ""
                              COUNTRY_CD_ISS = ""

                              dateInsert = Format(Trim(rs("TRANS_DT")), "YYYY") & "-" & Format(Trim(rs("TRANS_DT")), "MM") & "-" & Format(Trim(rs("TRANS_DT")), "DD")
                              If Trim(rs("payment_flag")) = "P" Then
                                    TRAN_TYPE_CD = "270002"
                               Else
                                    TRAN_TYPE_CD = "270001"
                               End If
                               
                              
                              '------------------------------------------------------------------


     
                            '  sSqlSub = "       SELECT SEC_FEATURE.INST_DESC AS INST_DESC ,  SEC_FEATURE.INST_TYPE AS INST_TYPE  , SEC_FEATURE.ISIN_CODE  AS ISIN_CODE  FROM ESL_COMMON AS COMMON "
                            '  sSqlSub = sSqlSub & " INNER JOIN ESL_SEC_FEATURE AS SEC_FEATURE "
                            '  sSqlSub = sSqlSub & " ON COMMON.SECURITY_ID = SEC_FEATURE.SECURITY_ID  "
                            '  sSqlSub = sSqlSub & "  WHERE COMMON.TRANS_NUM = '" & Trim(rs("TRANS_NUM")) & "'   WITH UR;  "
                           '  Set rsSub1 = Conn.Execute(sSqlSub)
                           sSqlSub = " SELECT SF.INST_DESC  ,  SF.INST_TYPE   , SF.ISIN_CODE , SF.ENTITY_CODE , E.LONG_NAME,M.MAP_CD2 "
                           sSqlSub = sSqlSub & " From ESL_SEC_FEATURE  AS SF LEFT JOIN ESL_SYENTITY AS E  "
                           sSqlSub = sSqlSub & " On SF.ENTITY_CODE = E.ENTITY_CODE"
                           sSqlSub = sSqlSub & " Left Join MAP_CODE  AS M On E.DOMI_CODE = M.MAP_CD1 And  M.MAP_TABLE_CD = 'MAP029' "
                           sSqlSub = sSqlSub & " Where  SF.SECURITY_ID  = '" & Trim(rs("SECURITY_ID")) & "'  "
                           Set rsSub1 = Conn.Execute(sSqlSub)
                           
                              If Not rsSub1.EOF Then
                                     'DEBT_INS_NAME = Replace(Trim(rsSub1("INST_DESC")), "&", " AND ")
                                    ' ISIN_CODE = Replace(Trim(rsSub1("ISIN_CODE")), "&", " AND ")
                                     If IsNull(rsSub1("INST_DESC")) = True Or Trim(rsSub1("INST_DESC")) = "" Then
                                                 INST_DESC = ""
                                     Else
                                                 INST_DESC = Trim(rsSub1("INST_DESC"))
                                                 DEBT_INS_NAME = Replace(INST_DESC, "&", " AND ")
                                     End If
                                       If IsNull(rsSub1("ISIN_CODE")) = True Or Trim(rsSub1("ISIN_CODE")) = "" Then
                                                 ISIN_CODE = ""
                                     Else
                                                 ISIN_CODE = Trim(rsSub1("ISIN_CODE"))
                                                 ISIN_CODE = Replace(ISIN_CODE, "&", " AND ")
                                     End If
                                     If IsNull(rsSub1("LONG_NAME")) = True Or Trim(rsSub1("LONG_NAME")) = "" Then
                                                 ISSUER_NAME = ""
                                     Else
                                                 ISSUER_NAME = Trim(rsSub1("LONG_NAME"))
                                                 ISSUER_NAME = Replace(ISSUER_NAME, "&", " AND ")
                                     End If
                                     
                                          If IsNull(rsSub1("MAP_CD2")) = True Or Trim(rsSub1("MAP_CD2")) = "" Then
                                                 COUNTRY_CD_ISS = ""
                                     Else
                                                 COUNTRY_CD_ISS = Trim(rsSub1("MAP_CD2"))
                                                 COUNTRY_CD_ISS = Replace(COUNTRY_CD_ISS, "&", " AND ")
                                     End If
                                     
                                
                               End If
                                          
                           
                                          
                                          
                                          sSqlInsert = " INSERT INTO DS_PTX ( ARR_TYPE_CD, ARR_NUMBER, TRAN_DTE, TRAN_TYPE_CD, ITEM_TYPE_CD, TRAN_SEQ, REC_PAY_ITEM_DSC, DATA_SET_DATE, PROVIDER_BR_NO, IBF_IND_CD, PAY_METHOD_CD, CURR_CD, TRAN_AMT_CURR, DEBT_ARR_TYPE_CD, ISIN_CODE, DEBT_INS_NAME, ISSUER_NAME, COUNTRY_CD_ISS, ISSUE_DTE, MATURITY_DTE, ORG_TRM, ORG_TRM_U, COUPON_RATE, INT_COUNTRY_CD, UNIT_OF_TRAN, SELL_SEC_AMT, PURCHASE_DTE, OS_AMT, LU_DTE, LU_TIME, USER_LU, DATA_SYSTEM_CD, FLAG_ON_OFF, PART_ID, IP_NAME, IP_COUNTRY_CD) "
                                         sSqlInsert = sSqlInsert & "   VALUES ( '999999' ,'" & Trim(rs("TRANS_NUM")) & "' ,'" & dateInsert & "' ,'" & TRAN_TYPE_CD & "' ,'268006' ,  " & TRAN_SEQ & "  ,'' ,'" & dateInsert & "' ,''  ,'' ,''  "
                                          sSqlInsert = sSqlInsert & "  ,'" & Trim(rs("PAYMENT_CUR")) & "' , " & Abs(rs("INT_AMT")) & " ,'' , '" & ISIN_CODE & "' , '" & DEBT_INS_NAME & "' , '" & ISSUER_NAME & "', '" & COUNTRY_CD_ISS & "' ,  NULL ,NULL ,NULL ,'' ,NULL ,'' ,NULL ,0, NULL ,0       "
                                          sSqlInsert = sSqlInsert & "  ,Current Date ,Current Time ,'PLDMS' ,'HIPO' ,'0' ,'' ,'' ,'' ) "

                                          sSqlInsertLog = " INSERT INTO DS_PTX_LOG ( ARR_TYPE_CD, ARR_NUMBER, TRAN_DTE, TRAN_TYPE_CD, ITEM_TYPE_CD, TRAN_SEQ, ACTION_TIMESTAMP, ACTION_FLAG, REC_PAY_ITEM_DSC, DATA_SET_DATE, PROVIDER_BR_NO, IBF_IND_CD, PAY_METHOD_CD, CURR_CD, TRAN_AMT_CURR, DEBT_ARR_TYPE_CD, ISIN_CODE, DEBT_INS_NAME, ISSUER_NAME, COUNTRY_CD_ISS, ISSUE_DTE, MATURITY_DTE, ORG_TRM, ORG_TRM_U, COUPON_RATE, INT_COUNTRY_CD, UNIT_OF_TRAN, SELL_SEC_AMT, PURCHASE_DTE, OS_AMT, LU_DTE, LU_TIME, USER_LU, DATA_SYSTEM_CD, FLAG_ON_OFF, PART_ID, IP_NAME, IP_COUNTRY_CD) "
                                          sSqlInsertLog = sSqlInsertLog & "   VALUES ( '999999' ,'" & Trim(rs("TRANS_NUM")) & "' ,'" & dateInsert & "' ,'" & TRAN_TYPE_CD & "' ,'268006' ,  " & TRAN_SEQ & "  ,CURRENT TIMESTAMP, 'I'   ,'' ,'" & dateInsert & "' ,''  ,'' ,''  "
                                          sSqlInsertLog = sSqlInsertLog & "  ,'" & Trim(rs("PAYMENT_CUR")) & "' , " & Abs(rs("INT_AMT")) & " ,'' , '" & ISIN_CODE & "' , '" & DEBT_INS_NAME & "' , '" & ISSUER_NAME & "', '" & COUNTRY_CD_ISS & "' ,  NULL ,NULL ,NULL ,'' ,NULL ,'' ,NULL ,0, NULL ,0       "
                                          sSqlInsertLog = sSqlInsertLog & "   ,Current Date ,Current Time ,'PLDMS' ,'HIPO' ,'0' ,'' ,'' ,'' ) "
                                          Conn.Execute (sSqlInsert)
                                          Conn.Execute (sSqlInsertLog)
                                          count_insert = count_insert + 1
                                          TRAN_SEQ = TRAN_SEQ + 1
                              Set rsSub1 = Nothing
                              Set rsSub2 = Nothing
                              rs.MoveNext

                        Loop
                        End If
                        Set rs = Nothing
ErrDB:
      If Err.Number <> 0 Then
            Conn.RollbackTrans
            PTX_FI_DAILY = Err.Description
      Else
            Conn.CommitTrans
            PTX_FI_DAILY = ""
      End If
End Function
