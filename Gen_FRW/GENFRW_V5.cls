VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GENFRW_V5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function GenXMLFRW(GenDate) As String

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
   
fso.CreateTextFile ("C:\DMS_DATA\TEMP_FRW_TCB_" & GenDate & ".txt")
 Set objFil = fso.GetFile("C:\DMS_DATA\TEMP_FRW_TCB_" & GenDate & ".txt")
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

    
    sql = "SELECT * FROM DS_FRW  WHERE DATA_SET_DATE = '" & GenDate & "' " & _
    "Order by TRAN_SEQ "
    
    SqlLog = "insert into Gen_XMl_log values ('" & GenDate & "','DS_FRW',Current Timestamp,Current Timestamp,0,0)"

    With rs
            If .State = adStateOpen Then .Close
           .ActiveConnection = Conn
           .CursorType = adOpenForwardOnly
           .CursorLocation = adUseClient
           .open sql
    If .EOF Then
           
        GenXMLFRW = "Data Not Found"
    Exit Function
  Else
      Set RsLog = Conn.Execute(SqlLog)
            
            
        'Write XML Header
            strXML = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & " standalone=" & Chr(34) & "yes" & Chr(34) & "?>" & vbCrLf
            strXML = strXML & "<DS_FRW name=" & Chr(34) & "Fee Rate of SWIFT" & Chr(34) & " type=" & Chr(34) & "No Fixed CL no sub-repeating" & Chr(34) & " schemaVersion=" & Chr(34) & "TCB 5.0" & Chr(34) & " > " & vbCrLf
            strXML = strXML & "<CommonHeader>" & vbCrLf
            strXML = strXML & "<OrganizationId>014</OrganizationId>" & vbCrLf
            strXML = strXML & "<DataSetDate>" & Format(GenDate, "yyyy-mm-dd") & "</DataSetDate>" & vbCrLf
            strXML = strXML & "</CommonHeader>" & vbCrLf
            strXML = strXML & "<DS_FRW_Header />" & vbCrLf
            strXML = strXML & "<DS_FRW_Content>" & vbCrLf

        Do While rs.EOF = False
            
            strXML = strXML & "<ContentRecordGroup>" & vbCrLf
                        
            ' 1. FCD CurrencyId
            strXML = strXML & "<FcdCurrencyId>" & Trim(rs("CURR_CD")) & "</FcdCurrencyId>" & vbCrLf

             ' 2. Currency of Fee
             If (Trim(rs("COMM_DEP")) = "" Or IsNull(Trim(rs("COMM_DEP"))) = True) Then
                    strXML = strXML & "<CurrencyOfFee/>" & vbCrLf
            Else
                    strXML = strXML & "<CurrencyOfFee>" & Trim(rs("COMM_DEP")) & "</CurrencyOfFee>" & vbCrLf
            End If

            ' 3. CommissionInLieuRateForDeposit
            If (Trim(rs("COMM_DEP")) = "" Or IsNull(Trim(rs("COMM_DEP"))) = True) Then
                    strXML = strXML & "<CommissionInLieuRateForDeposit/>" & vbCrLf
            Else
                    strXML = strXML & "<CommissionInLieuRateForDeposit>" & Trim(rs("COMM_DEP")) & "</CommissionInLieuRateForDeposit>" & vbCrLf
            End If
            
            ' 4. MinCommissionInLieuForDeposit
            If (Trim(rs("MIN_COMM_DEP")) = "" Or IsNull(Trim(rs("MIN_COMM_DEP"))) = True) Then
                   strXML = strXML & "<MinCommissionInLieuForDeposit/>" & vbCrLf
            Else
                   strXML = strXML & "<MinCommissionInLieuForDeposit>" & Trim(rs("MIN_COMM_DEP")) & "</MinCommissionInLieuForDeposit>" & vbCrLf
            End If
            
            ' 5. MaxCommissionInLieuForDeposit
            If (Trim(rs("MAX_COMM_DEP")) = "" Or IsNull(Trim(rs("MAX_COMM_DEP"))) = True) Then
                strXML = strXML & "<MaxCommissionInLieuForDeposit/>" & vbCrLf
            Else
                strXML = strXML & "<MaxCommissionInLieuForDeposit>" & Trim(rs("MAX_COMM_DEP")) & "</MaxCommissionInLieuForDeposit>" & vbCrLf
            End If
                     
            ' 6. InwardsTransferFeeForDeposit
            If (Trim(rs("INW_FEE_DEP")) = "" Or IsNull(Trim(rs("INW_FEE_DEP"))) = True) Then
                strXML = strXML & "<InwardsTransferFeeForDeposit/>" & vbCrLf
            Else
                strXML = strXML & "<InwardsTransferFeeForDeposit>" & Trim(rs("INW_FEE_DEP")) & "</InwardsTransferFeeForDeposit>" & vbCrLf
            End If
            
            ' 7. MinInwardsTransferFeeForDeposit
            If (Trim(rs("MIN_INW_FEE_DEP")) = "" Or IsNull(Trim(rs("MIN_INW_FEE_DEP"))) = True) Then
                strXML = strXML & "<MinInwardsTransferFeeForDeposit />" & vbCrLf
            Else
                strXML = strXML & "<MinInwardsTransferFeeForDeposit >" & Trim(rs("MIN_INW_FEE_DEP")) & "</MinInwardsTransferFeeForDeposit>" & vbCrLf
            End If
             
            ' 8. MaxInwardsTransferFeeForDeposit
            If (Trim(rs("MAX_INW_FEE_DEP")) = "" Or IsNull(Trim(rs("MAX_INW_FEE_DEP"))) = True) Then
                strXML = strXML & "<MaxInwardsTransferFeeForDeposit />" & vbCrLf
            Else
                strXML = strXML & "<MaxInwardsTransferFeeForDeposit >" & Trim(rs("MAX_INW_FEE_DEP")) & "</MaxInwardsTransferFeeForDeposit>" & vbCrLf
            End If
            
            ' 9. OtherFeeForDeposit
            If (Trim(rs("OTH_FEE_DEP")) = "" Or IsNull(Trim(rs("OTH_FEE_DEP"))) = True) Then
                strXML = strXML & "<OtherFeeForDeposit />" & vbCrLf
            Else
                strXML = strXML & "<OtherFeeForDeposit >" & Trim(rs("OTH_FEE_DEP")) & "</OtherFeeForDeposit>" & vbCrLf
            End If
            
            ' 10. RemarksOfFeeForDeposit
            If (Trim(rs("OTH_FEE_DEP")) = "" Or IsNull(Trim(rs("OTH_FEE_DEP"))) = True) Then
                strXML = strXML & "<RemarksOfFeeForDeposit />" & vbCrLf
            Else
                strXML = strXML & "<RemarksOfFeeForDeposit >" & Trim(rs("OTH_FEE_DEP")) & "</RemarksOfFeeForDeposit>" & vbCrLf
            End If
            
            ' 11. CommissionInLieuRateForWithdrawal
            If (Trim(rs("COMM_WD")) = "" Or IsNull(Trim(rs("COMM_WD"))) = True) Then
                strXML = strXML & "<CommissionInLieuRateForWithdrawal />" & vbCrLf
            Else
                strXML = strXML & "<CommissionInLieuRateForWithdrawal >" & Trim(rs("COMM_WD")) & "</CommissionInLieuRateForWithdrawal>" & vbCrLf
            End If
            
            ' 12. MinCommissionInLieuForWithdrawal
            If (Trim(rs("MIN_COMM_WD")) = "" Or IsNull(Trim(rs("MIN_COMM_WD"))) = True) Then
                strXML = strXML & "<MinCommissionInLieuForWithdrawal />" & vbCrLf
            Else
                strXML = strXML & "<MinCommissionInLieuForWithdrawal >" & Trim(rs("MIN_COMM_WD")) & "</MinCommissionInLieuForWithdrawal>" & vbCrLf
            End If

            ' 13. MaxCommissionInLieuForWithdrawal
            If (Trim(rs("MAX_COMM_WD")) = "" Or IsNull(Trim(rs("MAX_COMM_WD"))) = True) Then
                strXML = strXML & "<MaxCommissionInLieuForWithdrawal />" & vbCrLf
            Else
                strXML = strXML & "<MaxCommissionInLieuForWithdrawal >" & Trim(rs("MAX_COMM_WD")) & "</MaxCommissionInLieuForWithdrawal>" & vbCrLf
            End If

            ' 14. WithdrawalFeeForOtherTransfer
            'If (Trim(rs("WD_FEE_OTH")) = "" Or IsNull(Trim(rs("WD_FEE_OTH"))) = True) Then
             '   strXML = strXML & "<WithdrawalFeeForOtherTransfer />" & vbCrLf
            'Else
             '   strXML = strXML & "<WithdrawalFeeForOtherTransfer >" & Trim(rs("WD_FEE_OTH")) & "</WithdrawalFeeForOtherTransfer>" & vbCrLf
            'End If
            
             ' 14. CurrencyOfTransferFeeForWithdrawal
            If (Trim(rs("EFFECTIVE_DATE")) = "" Or IsNull(Trim(rs("EFFECTIVE_DATE"))) = True) Then
                strXML = strXML & "<CurrencyOfTransferFeeForWithdrawal />" & vbCrLf
            Else
                strXML = strXML & "<CurrencyOfTransferFeeForWithdrawal>" & Trim(Format(rs("EFFECTIVE_DATE"), "yyyy-mm-dd")) & "</CurrencyOfTransferFeeForWithdrawal>" & vbCrLf
            End If
            
             ' 15.TransferFeeForWithdrawal
            If (Trim(rs("EFFECTIVE_DATE")) = "" Or IsNull(Trim(rs("EFFECTIVE_DATE"))) = True) Then
                strXML = strXML & "<TransferFeeForWithdrawal />" & vbCrLf
            Else
                strXML = strXML & "<TransferFeeForWithdrawal>" & Trim(Format(rs("EFFECTIVE_DATE"), "yyyy-mm-dd")) & "</TransferFeeForWithdrawal>" & vbCrLf
            End If
            
            ' 16. RemarkOfFeeForWithdrawal
            If (Trim(rs("EFFECTIVE_DATE")) = "" Or IsNull(Trim(rs("EFFECTIVE_DATE"))) = True) Then
                strXML = strXML & "<RemarkOfFeeForWithdrawal />" & vbCrLf
            Else
                strXML = strXML & "<RemarkOfFeeForWithdrawal>" & Trim(Format(rs("EFFECTIVE_DATE"), "yyyy-mm-dd")) & "</RemarkOfFeeForWithdrawal>" & vbCrLf
            End If
            
             ' 17. EffectiveDate
            If (Trim(rs("EFFECTIVE_DATE")) = "" Or IsNull(Trim(rs("EFFECTIVE_DATE"))) = True) Then
                strXML = strXML & "<EffectiveDate />" & vbCrLf
            Else
                strXML = strXML & "<EffectiveDate>" & Trim(Format(rs("EFFECTIVE_DATE"), "yyyy-mm-dd")) & "</EffectiveDate>" & vbCrLf
            End If
            
            ' 18. EndDate
            If (Trim(rs("END_DATE")) = "" Or IsNull(Trim(rs("END_DATE"))) = True) Then
                strXML = strXML & "<EndDate />" & vbCrLf
            Else
                strXML = strXML & "<EndDate >" & Trim(Format(rs("END_DATE"), "yyyy-mm-dd")) & "</EndDate>" & vbCrLf
            End If
            
            ' 19. WebsiteLinkForDeposit
            If (Trim(rs("EFFECTIVE_DATE")) = "" Or IsNull(Trim(rs("EFFECTIVE_DATE"))) = True) Then
                strXML = strXML & "<WebsiteLinkForDeposit />" & vbCrLf
            Else
                strXML = strXML & "<WebsiteLinkForDeposit>" & Trim(Format(rs("EFFECTIVE_DATE"), "yyyy-mm-dd")) & "</WebsiteLinkForDeposit>" & vbCrLf
            End If
            
            
            ' 20. WebsiteLinkForWithdrawal
            If (Trim(rs("EFFECTIVE_DATE")) = "" Or IsNull(Trim(rs("EFFECTIVE_DATE"))) = True) Then
                strXML = strXML & "<WebsiteLinkForWithdrawal />" & vbCrLf
            Else
                strXML = strXML & "<WebsiteLinkForWithdrawal>" & Trim(Format(rs("EFFECTIVE_DATE"), "yyyy-mm-dd")) & "</WebsiteLinkForWithdrawal>" & vbCrLf
            End If

 
            objTs.Write (strXML)
            strXML = ""
            rs.MoveNext
            
                    
                    
                strXML = strXML & "</ContentRecordGroup>" & vbCrLf
           If rs.EOF Then Exit Do
        Loop
   
                strXML = strXML & "</DS_FRW_Content>" & vbCrLf
                strXML = strXML & "</DS_FRW>" & vbCrLf
                    
objTs.Write (strXML)
strXML = ""
Set objTs = fso.OpenTextFile("C:\DMS_DATA\TEMP_FRW_TCB_" & GenDate & ".txt", ForReading, False)
 Tem = objTs.ReadAll
xmlDOMTmp.loadXML (Tem)
 
 PathXML = "C:\DMS_DATA\DS_FRW_TCB_" & GenDate & ".xml"
 xmlDOMTmp.save PathXML
 SqlUpdateLog = "Update gen_xml_log Set ds_stop = current timestamp," & _
                "ds_record =" & rs.RecordCount & ",ds_amount = " & Sumamount & " where data_set_date ='" & GenDate & "' and ds_abbr_name ='DS_FRW' and ds_start =" & _
                "(select max(ds_start) from gen_xml_log where data_set_date ='" & GenDate & "' and ds_abbr_name = 'DS_FRW')"
                
    
 Set RsLogUpdate = Conn.Execute(SqlUpdateLog)
 
            End If
    End With
    
rs.Close
Conn.Close
Set rs = Nothing

Set RsLog = Nothing
Set RsLogUpdate = Nothing
Set Conn = Nothing
    GenXMLFRW = "0"        ' 0 = Success
Exit Function
ErrHdl:
      GenXMLFRW = "Gen_FRW Error" & Err.Number & Err.Description
End Function






