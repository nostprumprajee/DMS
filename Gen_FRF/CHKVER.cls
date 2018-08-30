VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHKVER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public DateE As Date
Public Conn As ADODB.Connection
Public Type UTDVer
    UTD_XVer As String
    UTD_CVer As String
    UTD_Error As String
End Type

Private Function Returndate(x As Date) As String
Dim y As String
  y = Year(x) & "-" & Month(x) & "-" & Day(x)
 Returndate = Format(y, "yyyy-mm-dd")
End Function

Private Function CloseDB() As String
        Conn.Close
        Set Conn = Nothing
End Function

Private Function ConnectDB() As String
Dim iFile As Integer
Dim ObjEnCode As New SCrypt.clsRijndael
Dim StrOut1 As String
Dim StrOut2 As String
Dim StrOut3 As String
Dim P1 As String
Dim P2 As String
Dim P3 As String
Dim Path_Con As String
On Error GoTo Err_Handler
         iFile = FreeFile
        Path_Con = "C:\config.ini"
        Open Path_Con For Input Access Read As #iFile
        Input #iFile, StrOut1
        P1 = ObjEnCode.DecryptString(StrOut1, "DMS-BOT", True) ' DSN
        Input #iFile, StrOut2
        P2 = ObjEnCode.DecryptString(StrOut2, "DMS-BOT", True) ' user
        Input #iFile, StrOut3
        P3 = ObjEnCode.DecryptString(StrOut3, "DMS-BOT", True) ' Password

        Set Conn = New ADODB.Connection
        Conn.Open P1, P2, P3
        ConnectDB = "1"
        Exit Function

Err_Handler:
        ConnectDB = "0"
        Exit Function

End Function
Public Function CHK_Web_Ver(Ds_Date As String) As String
Dim temp, sql As String
Dim mVer As UTDVer
On Error GoTo Err_Handler
Err.Clear
      mVer = CHK_Version("CLASSI", Ds_Date)
    If mVer.UTD_Error = "" Then
            CHK_Web_Ver = mVer.UTD_CVer
    Else
            CHK_Web_Ver = mVer.UTD_Error
    End If
    Exit Function
    
Err_Handler:
    CHK_Web_Ver = "CHK_Web_Ver CHKVer Error " & Err.Number & Err.Description
     Exit Function
End Function

Public Function CHK_Version(Data_set, Ds_Date As String) As UTDVer
Dim mVer As UTDVer
Dim temp, sql As String
Dim rs, rs1 As ADODB.Recordset
On Error GoTo Err_Handler
Err.Clear
    
    temp = ConnectDB()
    If temp = 0 Then
           mVer.UTD_XVer = ""
           mVer.UTD_CVer = ""
           mVer.UTD_Error = "CHK_Version CHKVer Error : Can not Connect to Database "
           CHK_Version = mVer
           Exit Function
    End If
    mVer.UTD_XVer = ""
    mVer.UTD_CVer = ""
    mVer.UTD_Error = ""
        If Data_set <> "CLASSI" Then
                sql = "select * from ref_version where data_set = '" & Data_set & "' and effective_dt <= '" & Ds_Date & "' and end_dt >= '" & Ds_Date & "' order by effective_dt,end_dt"
                Set rs = Conn.Execute(sql)
                If (rs.EOF Or rs.BOF) Then
                        mVer.UTD_XVer = ""
                Else
                        mVer.UTD_XVer = Trim(rs("Version"))
                End If
        End If
        sql = "select * from ref_version where data_set = 'CLASSI' and effective_dt <= '" & Ds_Date & "' and end_dt >= '" & Ds_Date & "' order by effective_dt,end_dt"
        Set rs = Conn.Execute(sql)
        If (rs.EOF Or rs.BOF) Then
                mVer.UTD_CVer = ""
                mVer.UTD_Error = "CHKVer Not Found"
                CHK_Version = mVer
        Else
                mVer.UTD_CVer = Trim(rs("Version"))
                mVer.UTD_Error = ""
                CHK_Version = mVer
        End If
    temp = CloseDB()
    Exit Function

Err_Handler:
    mVer.UTD_XVer = ""
    mVer.UTD_CVer = ""
    mVer.UTD_Error = "CHK_Version CHKVer Error " & Err.Number & Err.Description
    CHK_Version = mVer
 Exit Function
    
End Function

Function Path_Config() As String
    Dim strFDPathInfo, strFDPhysicalPath, ConfigPath
    Dim tmpFolder, tmpPath
    Dim objFDfso, objfdfile
    Dim j, l
    
    strFDPhysicalPath = App.Path
    'MsgBox ("strFDPhysicalPath :" & strFDPhysicalPath)
     l = InStrRev(strFDPhysicalPath, "\", Len(strFDPhysicalPath))
    tmpFolder = Mid(strFDPhysicalPath, 1, l)
    'MsgBox ("tmpFolder:" & tmpFolder)
    ConfigPath = tmpFolder & "SCB-DMS-BOT\config.ini"
    'MsgBox ("configPath:" & ConfigPath)
    Path_Config = ConfigPath
End Function

