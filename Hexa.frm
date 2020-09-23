VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ascii 2 Hexa"
   ClientHeight    =   4155
   ClientLeft      =   4230
   ClientTop       =   3885
   ClientWidth     =   5670
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy to clipboard"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   5415
   End
   Begin VB.TextBox txtHexa 
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1920
      Width           =   5415
   End
   Begin VB.TextBox txtAscii 
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCopy_Click()

    txtHexa.SetFocus
    SendKeys "+(^{END})"
    SendKeys "^(C)"
    SendKeys "{END}"
    
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub txtAscii_Change()

txtHexa = ""

    For i = 1 To Len(txtAscii)
        txtHexa = txtHexa & Hex(Asc(Mid(txtAscii, i, 1)))
    Next

End Sub
