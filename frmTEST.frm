VERSION 5.00
Object = "*\ADataAccessCTL.vbp"
Begin VB.Form frmTEST 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin DataAccessCTL.DATA_ACCESS xDATA1 
      Left            =   60
      Top             =   60
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin VB.TextBox txtTEST 
      Height          =   285
      Left            =   1260
      TabIndex        =   0
      Top             =   420
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "You must fill in the properties of the ActiveData control to use an existing database on your machine."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   1620
      Width           =   4275
   End
End
Attribute VB_Name = "frmTEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    xDATA1.Connect xSQL
    txtTEST.Text = xDATA1.xRecordset.Fields.Item(0)
    MsgBox xDATA1.Instructions, vbInformation, "INSTRUCTIONS"

End Sub
