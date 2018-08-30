VERSION 5.00
Begin VB.Form TEST_FRW 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Gen DS_FRW"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1560
      TabIndex        =   0
      Top             =   1185
      Width           =   1365
   End
End
Attribute VB_Name = "TEST_FRW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim OBJ As Object
    Set OBJ = CreateObject("DMSGENFRW.GENFRW")
    MsgBox OBJ.GenXMLFRW("2012-10-29")
End Sub

