VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GENFRF_V5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function GenXMLFRF(GenDate) As String

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
   
fso.CreateTextFile ("C:\DMS_DATA\TEMP_FRF_TCB_" & GenDate & ".txt")
 Set objFil = fso.GetFile("C:\DMS_DATA\TEMP_FRF_TCB_" & GenDate & ".txt")
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

    
    sql = "SELECT * FROM DS_FRF  WHERE DATA_SET_DATE = '" & GenDate & "' " & _
    "Order by TRAN_SEQ "
    
    SqlLog = "insert into Gen_XMl_log values ('" & GenDate & "','DS_FRF',Current Timestamp,Current Timestamp,0,0)"

    With rs
            If .State = adStateOpen Then .Close
           .ActiveConnection = Conn
           .CursorType = adOpenForwardOnly
           .CursorLocation = adUseClient
           .open sql
    If .EOF Then
           
        GenXMLFRF = "Data Not Found"
    Exit Function
  Else
      Set RsLog = Conn.Execute(SqlLog)
            
            
        'Write XML Header
            strXML = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & " standalone=" & Chr(34) & "yes" & Chr(34) & "?>" & vbCrLf
            strXML = strXML & "<DS_FRF name=" & Chr(34) & "Fee Rate of Foreign Currency Deposit" & Chr(34) & " type=" & Chr(34) & "No Fixed CL no sub-repeating" & Chr(34) & " schemaVersion=" & Chr(34) & "TCB 5.0" & Chr(34) & " > " & vbCrLf
            strXML = strXML & "<CommonHeader>" & vbCrLf
            strXML = strXML & "<OrganizationId>014</OrganizationId>" & vbCrLf
            strXML = strXML & "<DataSetDate>" & Format(GenDate, "yyyy-mm-dd") & "</DataSetDate>" & vbCrLf
            strXML = strXML & "</CommonHeader>" & vbCrLf
            strXML = strXML & "<DS_FRF_Header />" & vbCrLf
            strXML = strXML & "<DS_FRF_Content>" & vbCrLf

        Do While rs.EOF = False
            
            strXML = strXML & "<ContentRecordGroup>" & vbCrLf
                        
            ' 1. LoanDepositTransactionType  ( 184007 = Deposit ,184010 = Withdraw)
            strXML = strXML & "<LoanDepositTransactionType>" & Trim(rs("TRAN_TYPE_CD")) & "</LoanDepositTransactionType>" & vbCrLf
              
            ' 2. FCDCurrencyId
            strXML = strXML & "<FcdCurrencyId>" & Trim(rs("CURR_CD")) & "</FcdCurrencyId>" & vbCrLf

            ' 3. PaymentMethod  (234001,234002,234003,234007)
            strXML = strXML & "<PaymentMethod>" & Trim(rs("PAY_METHOD_CD")) & "</PaymentMethod>" & vbCrLf
      
            ' 4. BrOrBcFlag  (Null , 0,1 )
            If (Trim(rs("BR_BC_FLAG")) = "" Or IsNull(Trim(rs("BR_BC_FLAG"))) = True) Then
                   strXML = strXML & "<BrOrBcFlag/>" & vbCrLf
            Else
                   strXML = strXML & "<BrOrBcFlag >" & Trim(rs("BR_BC_FLAG")) & "</BrOrBcFlag >" & vbCrLf
            End If
            
            ' 5. CurrencyofCommissioninlieu
            If (Trim(rs("BR_BC_FLAG")) = "" Or IsNull(Trim(rs("BR_BC_FLAG"))) = True) Then
                   strXML = strXML & "<CurrencyOfCommissionInLieu/>" & vbCrLf
            Else
                   strXML = strXML & "<CurrencyOfCommissionInLieu >" & Trim(rs("BR_BC_FLAG")) & "</CurrencyOfCommissionInLieu >" & vbCrLf
            End If
            
            ' 6. CommissionInLieuRate
            If (Trim(rs("COMM_DESC")) = "" Or IsNull(Trim(rs("COMM_DESC"))) = True) Then
                strXML = strXML & "<CommissionInLieuRate/>" & vbCrLf
            Else
                strXML = strXML & "<CommissionInLieuRate >" & Trim(rs("COMM_DESC")) & "</CommissionInLieuRate >" & vbCrLf
            End If
                     
            ' 7. MinCommissionInLieu
            If (Trim(rs("MIN_COMM_DESC")) = "" Or IsNull(Trim(rs("MIN_COMM_DESC"))) = True) Then
                strXML = strXML & "<MinCommissionInLieu/>" & vbCrLf
            Else
                strXML = strXML & "<MinCommissionInLieu  >" & Trim(rs("MIN_COMM_DESC")) & "</MinCommissionInLieu>" & vbCrLf
            End If
            
            ' 7. MaxCommissionInLieu
            'If (Trim(rs("MAX_COMM_DESC")) = "" Or IsNull(Trim(rs("MAX_COMM_DESC"))) = True) Then
             '   strXML = strXML & "<MaxCommissionInLieu />" & vbCrLf
            'Else
             '   strXML = strXML & "<MaxCommissionInLieu >" & Trim(rs("MAX_COMM_DESC")) & "</MaxCommissionInLieu>" & vbCrLf
            'End If
            
            ' 8. CurrencyOfOtherFee
            If (Trim(rs("OTH_DESC")) = "" Or IsNull(Trim(rs("OTH_DESC"))) = True) Then
                strXML = strXML & "<CurrencyOfOtherFee />" & vbCrLf
            Else
                strXML = strXML & "<CurrencyOfOtherFee >" & Trim(rs("OTH_DESC")) & "</CurrencyOfOtherFee>" & vbCrLf
            End If
            
             ' 9. OtherFeeAmount
            If (Trim(rs("OTH_DESC")) = "" Or IsNull(Trim(rs("OTH_DESC"))) = True) Then
                strXML = strXML & "<OtherFeeAmount />" & vbCrLf
            Else
                strXML = strXML & "<OtherFeeAmount >" & Trim(rs("OTH_DESC")) & "</OtherFeeAmount>" & vbCrLf
            End If
             
            ' 10. RemarkofFee
            If (Trim(rs("OTH_DESC")) = "" Or IsNull(Trim(rs("OTH_DESC"))) = True) Then
                strXML = strXML & "<RemarkofFee />" & vbCrLf
            Else
                strXML = strXML & "<RemarkofFee >" & Trim(rs("OTH_DESC")) & "</RemarkofFee>" & vbCrLf
            End If
                    
            ' 11. EffectiveDate
            If (Trim(rs("EFFECTIVE_DATE")) = "" Or IsNull(Trim(rs("EFFECTIVE_DATE"))) = True) Then
                strXML = strXML & "<EffectiveDate />" & vbCrLf
            Else
                strXML = strXML & "<EffectiveDate>" & Trim(Format(rs("EFFECTIVE_DATE"), "yyyy-mm-dd")) & "</EffectiveDate>" & vbCrLf
            End If

            ' 12. EndDate
            If (Trim(rs("END_DATE")) = "" Or IsNull(Trim(rs("END_DATE"))) = True) Then
                strXML = strXML & "<EndDate />" & vbCrLf
            Else
                strXML = strXML & "<EndDate >" & Trim(Format(rs("END_DATE"), "yyyy-mm-dd")) & "</EndDate>" & vbCrLf
            End If
            
               ' 13. WebsiteLink
            If (Trim(rs("END_DATE")) = "" Or IsNull(Trim(rs("END_DATE"))) = True) Then
                strXML = strXML & "<Websitelink />" & vbCrLf
            Else
                strXML = strXML & "<WebsiteLink >" & Trim(Format(rs("END_DATE"), "yyyy-mm-dd")) & "</WebsiteLink>" & vbCrLf
            End If
 
            objTs.Write (strXML)
            strXML = ""
            rs.MoveNext
            
                    
                    
                strXML = strXML & "</ContentRecordGroup>" & vbCrLf
           If rs.EOF Then Exit Do
        Loop
   
                strXML = strXML & "</DS_FRF_Content>" & vbCrLf
                strXML = strXML & "</DS_FRF>" & vbCrLf
                    
objTs.Write (strXML)
strXML = ""
Set objTs = fso.OpenTextFile("C:\DMS_DATA\TEMP_FRF_TCB_" & GenDate & ".txt", ForReading, False)
 Tem = objTs.ReadAll
xmlDOMTmp.loadXML (Tem)
 
 PathXML = "C:\DMS_DATA\DS_FRF_TCB_" & GenDate & ".xml"
 xmlDOMTmp.save PathXML
 SqlUpdateLog = "Update gen_xml_log Set ds_stop = current timestamp," & _
                "ds_record =" & rs.RecordCount & ",ds_amount = " & Sumamount & " where data_set_date ='" & GenDate & "' and ds_abbr_name ='DS_FRF' and ds_start =" & _
                "(select max(ds_start) from gen_xml_log where data_set_date ='" & GenDate & "' and ds_abbr_name = 'DS_FRF')"
                
    
 Set RsLogUpdate = Conn.Execute(SqlUpdateLog)
 
            End If
    End With
    
rs.Close
Conn.Close
Set rs = Nothing

Set RsLog = Nothing
Set RsLogUpdate = Nothing
Set Conn = Nothing
    GenXMLFRF = "0"        ' 0 = Success
Exit Function
ErrHdl:
      GenXMLFRF = "Gen_FRF Error" & Err.Number & Err.Description
End Function






