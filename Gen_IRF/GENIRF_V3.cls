VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GENIRF_V3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function GenXMLIRF(GenDate) As String

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
   
fso.CreateTextFile ("C:\DMS_DATA\TEMP_IRF_TCB_" & GenDate & ".txt")
 Set objFil = fso.GetFile("C:\DMS_DATA\TEMP_IRF_TCB_" & GenDate & ".txt")
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

    
    sql = "SELECT * FROM DS_IRF  WHERE DATA_SET_DATE = '" & GenDate & "' " & _
    "Order by TRAN_SEQ "
    
    SqlLog = "insert into Gen_XMl_log values ('" & GenDate & "','DS_IRF',Current Timestamp,Current Timestamp,0,0)"

    With rs
            If .State = adStateOpen Then .Close
           .ActiveConnection = Conn
           .CursorType = adOpenForwardOnly
           .CursorLocation = adUseClient
           .open sql
    If .EOF Then
           
        GenXMLIRF = "Data Not Found"
    Exit Function
  Else
      Set RsLog = Conn.Execute(SqlLog)
            
            
        'Write XML Header
            strXML = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "utf-8" & Chr(34) & "?>" & vbCrLf
            strXML = strXML & "<DS_IRF xmlns:xsi=" & Chr(34) & "http://www.w3.org/2001/XMLSchema-instance" & Chr(34) & "  name=" & Chr(34) & "Interest Rate of Foreign Currency Deposit" & Chr(34) & " type=" & Chr(34) & "No Fixed CL with sub-repeating" & Chr(34) & " schemaVersion=" & Chr(34) & "TCB 3.0" & Chr(34) & " > " & vbCrLf
            strXML = strXML & "<CommonHeader>" & vbCrLf
            strXML = strXML & "<OrganizationId>014</OrganizationId>" & vbCrLf
            strXML = strXML & "<DataSetDate>" & Format(GenDate, "yyyy-mm-dd") & "</DataSetDate>" & vbCrLf
            strXML = strXML & "</CommonHeader>" & vbCrLf
            strXML = strXML & "<DS_IRF_Header />" & vbCrLf
            strXML = strXML & "<DS_IRF_Content>" & vbCrLf

        Do While rs.EOF = False
            
            strXML = strXML & "<ContentRecordGroup>" & vbCrLf
                        
            ' ArrangementType  (1)  018040 - Deposit from Customer \ Demand Deposit \ Saving / At Call (2) 018042 - Deposit from Customer  \ Term Deposit \ Fixed Deposit / Contractual Saving Deposit
            strXML = strXML & "<ArrangementType>" & Trim(rs("ARR_TYPE_CD")) & "</ArrangementType>" & vbCrLf
              
            ' InvolvedPartyType (1) 176001 - ผู้มีถิ่นฐานใน ปท. (2) 176003 - ธุรกิจ \นิติบุคคล (3) 176067 - ผู้มีถิ่นฐานในต่าง ปท. - คน (4) 176068 - ผู้มีถิ่นฐานในต่าง ปท. - ธุรกิจ
            strXML = strXML & "<InvolvedPartyType>" & Trim(rs("IP_TYPE_CD")) & "</InvolvedPartyType>" & vbCrLf

            ' CurrencyId
            strXML = strXML & "<CurrencyId>" & Trim(rs("CURR_CD")) & "</CurrencyId>" & vbCrLf
      
            ' DepositTerm
            If (Trim(rs("DEP_TERM")) = "" Or IsNull(Trim(rs("DEP_TERM"))) = True) Then
                   strXML = strXML & "<DepositTerm/>" & vbCrLf
            Else
                   strXML = strXML & "<DepositTerm>" & Trim(rs("DEP_TERM")) & "</DepositTerm>" & vbCrLf
            End If
            
            ' DepositTermUnit
            If (Trim(rs("DEP_TERM_UNIT")) = "" Or IsNull(Trim(rs("DEP_TERM_UNIT"))) = True) Then
                strXML = strXML & "<DepositTermUnit/>" & vbCrLf
            Else
                strXML = strXML & "<DepositTermUnit >" & Trim(rs("DEP_TERM_UNIT")) & "</DepositTermUnit>" & vbCrLf
            End If
                     
            ' ------------------------
            strXML = strXML & "<DataGroupInfo>" & vbCrLf
           strXML = strXML & "<DataGroupDetails>" & vbCrLf
            ' BalanceTierAmount
            If (Trim(rs("BAL_TIER_AMOUNT")) = "" Or IsNull(Trim(rs("BAL_TIER_AMOUNT"))) = True) Then
                strXML = strXML & "<BalanceTierAmount/>" & vbCrLf
            Else
                strXML = strXML & "<BalanceTierAmount  >" & Trim(rs("BAL_TIER_AMOUNT")) & "</BalanceTierAmount>" & vbCrLf
            End If
            
            ' InterestRate
            If (Trim(rs("INTEREST_RATE")) = "" Or IsNull(Trim(rs("INTEREST_RATE"))) = True) Then
                strXML = strXML & "<InterestRate />" & vbCrLf
            Else
                strXML = strXML & "<InterestRate >" & Trim(rs("INTEREST_RATE")) & "</InterestRate>" & vbCrLf
            End If
            ' ------------------------
           strXML = strXML & "</DataGroupDetails>" & vbCrLf
            strXML = strXML & "</DataGroupInfo>" & vbCrLf
             
            ' EffectiveDate
            If (Trim(rs("EFFECTIVE_DT")) = "" Or IsNull(Trim(rs("EFFECTIVE_DT"))) = True) Then
                strXML = strXML & "<EffectiveDate />" & vbCrLf
            Else
                ' strXML = strXML & "<EffectiveDate>" & Trim(rs("EFFECTIVE_DT")) & "</EffectiveDate>" & vbCrLf
                strXML = strXML & "<EffectiveDate>" & Format((rs("EFFECTIVE_DT").Value), "yyyy-mm-dd") & "</EffectiveDate>" & vbCrLf
            End If

            ' EndDate
            If (Trim(rs("END_DT")) = "" Or IsNull(Trim(rs("END_DT"))) = True) Then
                strXML = strXML & "<EndDate />" & vbCrLf
            Else
                ' strXML = strXML & "<EndDate >" & Trim(rs("END_DT")) & "</EndDate>" & vbCrLf
                strXML = strXML & "<EndDate >" & Format((rs("END_DT").Value), "yyyy-mm-dd") & "</EndDate>" & vbCrLf
            End If
            
            objTs.Write (strXML)
            strXML = ""
            rs.MoveNext
            
                    
                    
                strXML = strXML & "</ContentRecordGroup>" & vbCrLf
           If rs.EOF Then Exit Do
        Loop
   
                strXML = strXML & "</DS_IRF_Content>" & vbCrLf
                strXML = strXML & "</DS_IRF>" & vbCrLf
                    
objTs.Write (strXML)
strXML = ""
Set objTs = fso.OpenTextFile("C:\DMS_DATA\TEMP_IRF_TCB_" & GenDate & ".txt", ForReading, False)
 Tem = objTs.ReadAll
xmlDOMTmp.loadXML (Tem)
 
 PathXML = "C:\DMS_DATA\DS_IRF_TCB_" & GenDate & ".xml"
 xmlDOMTmp.save PathXML
 SqlUpdateLog = "Update gen_xml_log Set ds_stop = current timestamp," & _
                "ds_record =" & rs.RecordCount & ",ds_amount = " & Sumamount & " where data_set_date ='" & GenDate & "' and ds_abbr_name ='DS_IRF' and ds_start =" & _
                "(select max(ds_start) from gen_xml_log where data_set_date ='" & GenDate & "' and ds_abbr_name = 'DS_IRF')"
                
    
 Set RsLogUpdate = Conn.Execute(SqlUpdateLog)
 
            End If
    End With
    
rs.Close
Conn.Close
Set rs = Nothing

Set RsLog = Nothing
Set RsLogUpdate = Nothing
Set Conn = Nothing
    GenXMLIRF = "0"        ' 0 = Success
Exit Function
ErrHdl:
      GenXMLIRF = "Gen_IRF Error" & Err.Number & Err.Description
End Function




