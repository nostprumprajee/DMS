VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GENFRW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function GenXMLFRW(GenDate As String) As String
    Dim chkver As New DMSCHKVER.chkver
    Dim mVer As UTDVer
    Dim Ds_Date As String
    Dim DMSFRW As Object
    
    mVer = chkver.CHK_Version("DS_FRW", CStr(GenDate))
    If mVer.UTD_Error = "" Then
        Select Case Trim(mVer.UTD_XVer)
            Case "3.0" ' Effective 2009-01-01
                Set DMSFRW = CreateObject("DMSGENFRW.GENFRW_V3")
                GenXMLFRW = DMSFRW.GenXMLFRW(GenDate)
                Set DMSFRW = Nothing
            Case "4.0" ' Effective 2009-01-01
                Set DMSFRW = CreateObject("DMSGENFRW.GENFRW_V4")
                GenXMLFRW = DMSFRW.GenXMLFRW(GenDate)
                Set DMSFRW = Nothing
            Case "5.0" ' Effective 2018-10-01
                Set DMSFRW = CreateObject("DMSGENFRW.GENFRW_V5")
                GenXMLFRW = DMSFRW.GenXMLFRW(GenDate)
                Set DMSFRW = Nothing
            Case Else
                GenXMLFRW = "Version not found. Date = " & GenDate
        End Select
    End If
    Set chkver = Nothing
End Function

