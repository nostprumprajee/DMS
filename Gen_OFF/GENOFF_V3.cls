VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GENOFF_V3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function GenXMLOFF(GenDate) As String

Dim Conn As New ADODB.Connection
Dim xmlDOMTmp As New DOMDocument30
Dim rs As New ADODB.Recordset
Dim RsLogUpdate As New ADODB.Recordset
Dim RsLog As New ADODB.Recordset
Dim sql, SqlLog, SqlUpdateLog As String
Dim strXML As String
Dim PathXML As String
Dim Sumamount As Variant
Dim fso As New FileSystemObject
Dim objFil As File
Dim objTs As TextStream
Dim iFile As Integer
Dim ObjEnCode As New SCrypt.clsRijndael
Dim StrOut1 As String
Dim StrOut2 As String
Dim StrOut3 As String
Dim P1 As String
Dim P2 As String
Dim P3 As String

Sumamount = 0
On Error GoTo ErrHdl
    Err.Clear
   
fso.CreateTextFile ("C:\DMS_DATA\TEMP_OFF_TCB_" & GenDate & ".txt")
 Set objFil = fso.GetFile("C:\DMS_DATA\TEMP_OFF_TCB_" & GenDate & ".txt")
 Set objTs = objFil.OpenAsTextStream(ForWriting)

   iFile = FreeFile
Open "c:\Config.ini" For Input Access Read As #iFile
Input #iFile, StrOut1
Input #iFile, StrOut2
Input #iFile, StrOut3

P1 = ObjEnCode.DecryptString(StrOut1, "DMS-BOT", True) ' DSN
P2 = ObjEnCode.DecryptString(StrOut2, "DMS-BOT", True) ' user
P3 = ObjEnCode.DecryptString(StrOut3, "DMS-BOT", True) ' Password

      
    Conn.open P1, P2, P3

    
    sql = "SELECT * FROM DS_OFF  WHERE DATA_SET_DATE = '" & GenDate & "' " & _
    "Order by TRAN_SEQ "
    
    SqlLog = "insert into Gen_XMl_log values ('" & GenDate & "','DS_OFF',Current Timestamp,Current Timestamp,0,0)"

    With rs
            If .State = adStateOpen Then .Close
           .ActiveConnection = Conn
           .CursorType = adOpenForwardOnly
           .CursorLocation = adUseClient
           .open sql
    If .EOF Then
           
        GenXMLOFF = "Data Not Found"
    Exit Function
  Else
      Set RsLog = Conn.Execute(SqlLog)
            
            
        'Write XML Header
            strXML = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & " standalone=" & Chr(34) & "yes" & Chr(34) & "?>" & vbCrLf
            strXML = strXML & "<DS_OFF name=" & Chr(34) & "Other Fee Rate of Foreign Currency Deposit" & Chr(34) & " type=" & Chr(34) & "No Fixed CL no sub-repeating" & Chr(34) & " schemaVersion=" & Chr(34) & "TCB 3.0" & Chr(34) & " > " & vbCrLf
            strXML = strXML & "<CommonHeader>" & vbCrLf
            strXML = strXML & "<OrganizationId>014</OrganizationId>" & vbCrLf
            strXML = strXML & "<DataSetDate>" & Format(GenDate, "yyyy-mm-dd") & "</DataSetDate>" & vbCrLf
            strXML = strXML & "</CommonHeader>" & vbCrLf
            strXML = strXML & "<DS_OFF_Header />" & vbCrLf
            strXML = strXML & "<DS_OFF_Content>" & vbCrLf

        Do While rs.EOF = False
            
            strXML = strXML & "<ContentRecordGroup>" & vbCrLf
                        
            ' 1. ArrangementType  018039  018040  018042
            strXML = strXML & "<ArrangementType>" & Trim(rs("ARR_TYPE_CD")) & "</ArrangementType>" & vbCrLf
              
            ' 2. CurrencyId ต้องไม่เท่ากับ THB
            strXML = strXML & "<CurrencyId>" & Trim(rs("CURR_CD")) & "</CurrencyId>" & vbCrLf

            ' 3. MinimumOpenAccountAmount
            If (Trim(rs("MIN_OPEN_AC_AMT")) = "" Or IsNull(Trim(rs("MIN_OPEN_AC_AMT"))) = True) Then
                   strXML = strXML & "<MinimumOpenAccountAmount/>" & vbCrLf
            Else
                    strXML = strXML & "<MinimumOpenAccountAmount>" & Trim(rs("MIN_OPEN_AC_AMT")) & "</MinimumOpenAccountAmount>" & vbCrLf
            End If
            
            ' 4. MinimumBalanceForMaintenanceFee
            If (Trim(rs("MIN_BAL_MT_FEE")) = "" Or IsNull(Trim(rs("MIN_BAL_MT_FEE"))) = True) Then
                   strXML = strXML & "<MinimumBalanceForMaintenanceFee/>" & vbCrLf
            Else
                   strXML = strXML & "<MinimumBalanceForMaintenanceFee >" & Trim(rs("MIN_BAL_MT_FEE")) & "</MinimumBalanceForMaintenanceFee >" & vbCrLf
            End If
            
            ' 5. MaintenanceFee
            If (Trim(rs("MT_FEE")) = "" Or IsNull(Trim(rs("MT_FEE"))) = True) Then
                strXML = strXML & "<MaintenanceFee/>" & vbCrLf
            Else
                strXML = strXML & "<MaintenanceFee>" & Trim(rs("MT_FEE")) & "</MaintenanceFee>" & vbCrLf
            End If
                  
            ' 6. InactiveTerm
            If (Trim(rs("INACT_TERM")) = "" Or IsNull(Trim(rs("INACT_TERM"))) = True) Then
                strXML = strXML & "<InactiveTerm/>" & vbCrLf
            Else
                strXML = strXML & "<InactiveTerm>" & Trim(rs("INACT_TERM")) & "</InactiveTerm>" & vbCrLf
            End If
            
            ' 7. InactiveTermUnit
            If (Trim(rs("INACT_TERM_U")) = "" Or IsNull(Trim(rs("INACT_TERM_U"))) = True) Then
                strXML = strXML & "<InactiveTermUnit/>" & vbCrLf
            Else
                strXML = strXML & "<InactiveTermUnit>" & Trim(rs("INACT_TERM_U")) & "</InactiveTermUnit>" & vbCrLf
            End If
             
            ' 8. InactiveFee
            If (Trim(rs("INACT_FEE")) = "" Or IsNull(Trim(rs("INACT_FEE"))) = True) Then
                strXML = strXML & "<InactiveFee/>" & vbCrLf
            Else
                strXML = strXML & "<InactiveFee>" & Trim(rs("INACT_FEE")) & "</InactiveFee>" & vbCrLf
            End If
                    
            ' 9. EffectiveDate
            If (Trim(rs("EFFECTIVE_DATE")) = "" Or IsNull(Trim(rs("EFFECTIVE_DATE"))) = True) Then
                strXML = strXML & "<EffectiveDate/>" & vbCrLf
            Else
                strXML = strXML & "<EffectiveDate>" & Trim(Format(rs("EFFECTIVE_DATE"), "yyyy-mm-dd")) & "</EffectiveDate>" & vbCrLf
            End If

            ' 10. EndDate
            If (Trim(rs("END_DATE")) = "" Or IsNull(Trim(rs("END_DATE"))) = True) Then
                strXML = strXML & "<EndDate/>" & vbCrLf
            Else
                strXML = strXML & "<EndDate>" & Trim(Format(rs("END_DATE"), "yyyy-mm-dd")) & "</EndDate>" & vbCrLf
            End If
 
            objTs.Write (strXML)
            strXML = ""
            rs.MoveNext
            
                    
                    
                strXML = strXML & "</ContentRecordGroup>" & vbCrLf
           If rs.EOF Then Exit Do
        Loop
   
                strXML = strXML & "</DS_OFF_Content>" & vbCrLf
                strXML = strXML & "</DS_OFF>" & vbCrLf
                    
objTs.Write (strXML)
strXML = ""
Set objTs = fso.OpenTextFile("C:\DMS_DATA\TEMP_OFF_TCB_" & GenDate & ".txt", ForReading, False)
 Tem = objTs.ReadAll
xmlDOMTmp.loadXML (Tem)
 
 PathXML = "C:\DMS_DATA\DS_OFF_TCB_" & GenDate & ".xml"
 xmlDOMTmp.save PathXML
 SqlUpdateLog = "Update gen_xml_log Set ds_stop = current timestamp," & _
                "ds_record =" & rs.RecordCount & ",ds_amount = " & Sumamount & " where data_set_date ='" & GenDate & "' and ds_abbr_name ='DS_OFF' and ds_start =" & _
                "(select max(ds_start) from gen_xml_log where data_set_date ='" & GenDate & "' and ds_abbr_name = 'DS_OFF')"
                
    
 Set RsLogUpdate = Conn.Execute(SqlUpdateLog)
 
            End If
    End With
    
rs.Close
Conn.Close
Set rs = Nothing

Set RsLog = Nothing
Set RsLogUpdate = Nothing
Set Conn = Nothing
    GenXMLOFF = "0"        ' 0 = Success
Exit Function
ErrHdl:
      GenXMLOFF = "Gen_OFF Error" & Err.Number & Err.Description
End Function


