VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GENFRF_V1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function GenXMLCOS(GenDate) As String

Dim Conn As New ADODB.Connection
Dim xmlDOMTmp As New DOMDocument30
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
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
Dim OldFiReportingID, FiReportingID As String
Dim OldCreditRiskMethod, CreditRiskMethod As String
Dim OldArrTypeCd, ArrTypeCd As String
Dim Test As String

Sumamount = 0
On Error GoTo ErrHdl
    Err.Clear
   
fso.CreateTextFile ("C:\DMS_DATA\TEMP_COS_TCB_" & GenDate & ".txt")
 Set objFil = fso.GetFile("C:\DMS_DATA\TEMP_COS_TCB_" & GenDate & ".txt")
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

    
    sql = "SELECT * FROM DS_COS  WHERE DATA_SET_DATE = '" & GenDate & "' " & _
    "Order by FI_REPORTING_ID,CREDIT_RISK_METHOD,ARR_TYPE_CD "
    
    SqlLog = "insert into Gen_XMl_log values ('" & GenDate & "','DS_COS',Current Timestamp,Current Timestamp,0,0)"

    With rs
            If .State = adStateOpen Then .Close
           .ActiveConnection = Conn
           .CursorType = adOpenForwardOnly
           .CursorLocation = adUseClient
           .open sql
           
    If .EOF Then
        GenXMLCOS = "Data Not Found"
    Exit Function
  Else
      Set RsLog = Conn.Execute(SqlLog)
            
            
        'Write XML Header
            strXML = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & " ?>" & vbCrLf
            strXML = strXML & "<DS_COS xmlns:xsi=" & Chr(34) & "http://www.w3.org/2001/XMLSchema-instance" & Chr(34) & " name=" & Chr(34) & "Contingent Summary" & Chr(34) & " type=" & Chr(34) & "No Fixed CL with sub-repeating" & Chr(34) & " schemaVersion=" & Chr(34) & "TCB 3.0" & Chr(34) & " > " & vbCrLf
            strXML = strXML & "<CommonHeader>" & vbCrLf
            strXML = strXML & "<OrganizationId>014</OrganizationId>" & vbCrLf
            strXML = strXML & "<DataSetDate>" & Format(GenDate, "yyyy-mm-dd") & "</DataSetDate>" & vbCrLf
            strXML = strXML & "</CommonHeader>" & vbCrLf
            strXML = strXML & "<DS_COS_Header />" & vbCrLf
            strXML = strXML & "<DS_COS_Content>" & vbCrLf
            
            OldFiReportingID = ""
            OldCreditRiskMethod = ""
            OldArrTypeCd = ""
            
            Do While rs.EOF = False
            
            strXML = strXML & "<ContentRecordGroup>" & vbCrLf
            
                    FiReportingID = rs("FI_REPORTING_ID")
                    ArrTypeCd = rs("ARR_TYPE_CD")
                    
                    If FiReportingID <> OldFiReportingID Then
                            OldFiReportingID = FiReportingID
                        
                    ' FI_REPORTING_ID  ( 116002 ,116003)
                            strXML = strXML & "<FiReportingGroupId>" & rs("FI_REPORTING_ID") & "</FiReportingGroupId>" & vbCrLf
                    End If
                      
                    ' CREDIT_RISK_METHOD  (447001 to 447004)
                        
                            strXML = strXML & "<CreditRiskMethod>" & rs("CREDIT_RISK_METHOD") & "</CreditRiskMethod>" & vbCrLf

                 
                    ' NET_PL_FRM_MARK_TO_MKT
                            strXML = strXML & "<NetProfitAndLossFromMarkToMarket>" & rs("NET_PL_FRM_MARK_TO_MKT") & "</NetProfitAndLossFromMarkToMarket>" & vbCrLf
                 
                 

                 strXML = strXML & "<DataGroupInfo>" & vbCrLf
                 
                 
                 ' ----------------------------------------------------------------------- Loop DataGroupDetails  -----------------------------------------------------------------------
   Do
                 ' DataGroupDetails
                 strXML = strXML & "<DataGroupDetails>" & vbCrLf
                             
                             '  CUSTOMER GROUP (Null , 0,1 )
                             If (Trim(rs("CUSTOMER_GROUP")) = "" Or IsNull(Trim(rs("CUSTOMER_GROUP"))) = True) Then
                                    strXML = strXML & "<CustomerGroup />" & vbCrLf
                             Else
                                    strXML = strXML & "<CustomerGroup>" & rs("CUSTOMER_GROUP") & "</CustomerGroup>" & vbCrLf
                             End If
                             
                             If ArrTypeCd <> OldArrTypeCd Then
                                OldArrTypeCd = ArrTypeCd
                             End If
                             
                             ' CONTINGENT TYPE (ARR_TYPE_CD)
                                    strXML = strXML & "<ContingentType>" & rs("ARR_TYPE_CD") & "</ContingentType>" & vbCrLf
                            
                            
                             '  EXPOSURE METHOD
                             If (Trim(rs("EXPOSURE_METHOD")) = "" Or IsNull(Trim(rs("EXPOSURE_METHOD"))) = True) Then
                                    strXML = strXML & "<ExposureMethod />" & vbCrLf
                             Else
                                    strXML = strXML & "<ExposureMethod>" & rs("EXPOSURE_METHOD") & "</ExposureMethod>" & vbCrLf
                             End If
                             
                            '  BOOKING TYPE
                             If (Trim(rs("BOOKING_TYPE")) = "" Or IsNull(Trim(rs("BOOKING_TYPE"))) = True) Then
                                    strXML = strXML & "<BookingType  />" & vbCrLf
                             Else
                                    strXML = strXML & "<BookingType >" & rs("BOOKING_TYPE") & "</BookingType >" & vbCrLf
                             End If
                                  
                            '  NETTING  TYPE
                             If (Trim(rs("NETTING_TYPE")) = "" Or IsNull(Trim(rs("NETTING_TYPE"))) = True) Then
                                    strXML = strXML & "<NettingType  />" & vbCrLf
                             Else
                                    strXML = strXML & "<NettingType >" & rs("NETTING_TYPE") & "</NettingType >" & vbCrLf
                             End If
                             
                             ' SPECIFIC PROVISION AMT
                             strXML = strXML & "<SpecificProvision>" & rs("SPECIFIC_PROVISION_AMT") & "</SpecificProvision>" & vbCrLf
                             
                             ' NET CR_EQUIVALENT AMT
                             strXML = strXML & "<NetCreditEquivalentAmount>" & rs("NET_CR_EQUIVALENT_AMT") & "</NetCreditEquivalentAmount>" & vbCrLf
                             
                             '  ADD_PO_FUTURE_CR_EXP_AMT
                             If (Trim(rs("ADD_PO_FUTURE_CR_EXP_AMT")) = "" Or IsNull(Trim(rs("ADD_PO_FUTURE_CR_EXP_AMT"))) = True) Then
                                    strXML = strXML & "<AddOnOrPotentialFutureCreditExposure  />" & vbCrLf
                             Else
                                    strXML = strXML & "<AddOnOrPotentialFutureCreditExposure >" & rs("ADD_PO_FUTURE_CR_EXP_AMT") & "</AddOnOrPotentialFutureCreditExposure >" & vbCrLf
                             End If
                             
                              ' RWA_OS_AMT
                             strXML = strXML & "<RiskWeightedAssetOutstandingAmount>" & rs("RWA_OS_AMT") & "</RiskWeightedAssetOutstandingAmount>" & vbCrLf
                             
                             '  P_FRM_MARK_TO_MKT
                             If (Trim(rs("P_FRM_MARK_TO_MKT")) = "" Or IsNull(Trim(rs("P_FRM_MARK_TO_MKT"))) = True) Then
                                    strXML = strXML & "<ProfitFromMarkToMarket  />" & vbCrLf
                             Else
                                    strXML = strXML & "<ProfitFromMarkToMarket >" & rs("P_FRM_MARK_TO_MKT") & "</ProfitFromMarkToMarket >" & vbCrLf
                             End If
                             
                             '  L_FRM_MARK_TO_MKT
                             If (Trim(rs("L_FRM_MARK_TO_MKT")) = "" Or IsNull(Trim(rs("L_FRM_MARK_TO_MKT"))) = True) Then
                                    strXML = strXML & "<LossFromMarkToMarket  />" & vbCrLf
                             Else
                                    strXML = strXML & "<LossFromMarkToMarket >" & rs("L_FRM_MARK_TO_MKT") & "</LossFromMarkToMarket >" & vbCrLf
                             End If
                             
                             
                            strXML = strXML & "<BusinessGroupInfo>" & vbCrLf
                               
                            ' ----------------------------------------------------------------------- Loop BusinessGroupDetails  -----------------------------------------------------------------------
                                sql2 = "select  *  from DS_COS  where  DATA_SET_DATE = '" & GenDate & "'  " & _
                                            "and FI_REPORTING_ID =  '" & FiReportingID & "' and CREDIT_RISK_METHOD = '" & rs("CREDIT_RISK_METHOD") & "' " & _
                                            "and ARR_TYPE_CD =  '" & rs("ARR_TYPE_CD") & "'  " & _
                                            "Order by FI_REPORTING_ID,CREDIT_RISK_METHOD,ARR_TYPE_CD"


                                Set rs2 = Conn.Execute(sql2)
                                If rs2.EOF = False And (ArrTypeCd = rs("ARR_TYPE_CD")) Then

                                            Do
                                                    strXML = strXML & "<BusinessGroupDetails>" & vbCrLf
                                                    
                                                    ' CCF_RATE
                                                    strXML = strXML & "<CcfRate>" & rs2("CCF_RATE") & "</CcfRate>" & vbCrLf
                                                    
                                                     ' RWA_RATE
                                                    strXML = strXML & "<RwaRate>" & rs2("RWA_RATE") & "</RwaRate>" & vbCrLf
                                                    
                                                    ' IP_TYPE_CD
                                                    strXML = strXML & "<InvolvedPartyType>" & rs2("IP_TYPE_CD") & "</InvolvedPartyType>" & vbCrLf
                                                    
                                                    ' CURRENCY_FLAG
                                                    strXML = strXML & "<ArrangementCurrencyFlag>" & rs2("CURRENCY_FLAG") & "</ArrangementCurrencyFlag>" & vbCrLf
                                                    
                                                    ' NOTIONAL_PRC_AMT
                                                    strXML = strXML & "<NotionalPrincipleAmount>" & rs2("NOTIONAL_PRC_AMT") & "</NotionalPrincipleAmount>" & vbCrLf
                                                    strXML = strXML & "</BusinessGroupDetails>" & vbCrLf
                                            
                                                    rs2.MoveNext
                                                    rs.MoveNext
                                                    If rs2.EOF Then Exit Do
                                                    ArrTypeCd = rs2("ARR_TYPE_CD")
                                            Loop While ArrTypeCd = OldArrTypeCd
                                            
                                    End If


                                    
                                ' -------------------------------------------------------------------- End Loop BusinessGroupDetails  -------------------------------------------------------------------
                                
                                strXML = strXML & "</BusinessGroupInfo>" & vbCrLf
                                strXML = strXML & "</DataGroupDetails>" & vbCrLf
                                ' ---------------------------------------------------------------------- End Loop DataGroupDetails  -----------------------------------------------------------------------
                
            objTs.Write (strXML)
            strXML = ""
             '.MoveNext
            
                    If .EOF Then Exit Do
                    
                    FiReportingID = rs("FI_REPORTING_ID")
                    CreditRiskMethod = rs("CREDIT_RISK_METHOD")
                    ArrTypeCd = rs("ARR_TYPE_CD")
              Loop While FiReportingID = OldFiReportingID
            

               
                strXML = strXML & "</DataGroupInfo>" & vbCrLf
                strXML = strXML & "</ContentRecordGroup>" & vbCrLf
              
            Loop
   
                strXML = strXML & "</DS_COS_Content>" & vbCrLf
                strXML = strXML & "</DS_COS>" & vbCrLf
                    
objTs.Write (strXML)
strXML = ""
Set objTs = fso.OpenTextFile("C:\DMS_DATA\TEMP_COS_TCB_" & GenDate & ".txt", ForReading, False)
 Tem = objTs.ReadAll
xmlDOMTmp.loadXML (Tem)
 
 PathXML = "C:\DMS_DATA\DS_COS_TCB_" & GenDate & ".xml"
 xmlDOMTmp.save PathXML
 SqlUpdateLog = "Update gen_xml_log Set ds_stop = current timestamp," & _
                "ds_record =" & rs.RecordCount & ",ds_amount = " & Sumamount & " where data_set_date ='" & GenDate & "' and ds_abbr_name ='DS_COS' and ds_start =" & _
                "(select max(ds_start) from gen_xml_log where data_set_date ='" & GenDate & "' and ds_abbr_name = 'DS_COS')"
                
    
 Set RsLogUpdate = Conn.Execute(SqlUpdateLog)
 
            End If
    End With
    
rs2.Close
Set rs2 = Nothing
rs.Close
Conn.Close
Set rs = Nothing

Set RsLog = Nothing
Set RsLogUpdate = Nothing
Set Conn = Nothing
    GenXMLCOS = "0"        ' 0 = Success
Exit Function
ErrHdl:
      GenXMLCOS = "Gen_COS Error" & Err.Number & Err.Description
End Function


