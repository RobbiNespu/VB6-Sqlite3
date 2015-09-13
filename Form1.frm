VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "VB6+ SqLite"
   ClientHeight    =   2805
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3765
   LinkTopic       =   "Form1"
   ScaleHeight     =   2805
   ScaleWidth      =   3765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Add"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   360
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Next"
      Height          =   315
      Left            =   2280
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Masuk Data Baru"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Data dalam Database"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------
' Basic VB6 dengan sqlite
' Instruction: copy sqlite.dll ke folder system32
' Author: ApoNie [http://GeeKzLife.Net]
'-------------------------------------------------

Private Sub Form_Load()

    Call connectDb
    Call getData ' get data
    Text1.Text = row(i, 1) ' data pada column 1 (name) dan row i
    Text2.Text = row(i, 0) ' data pada column 0 (ic) dan row i
    
End Sub
Private Sub Form_unload(Response As Integer)

    Call closeDB

End Sub
Private Sub Command1_Click()

    If (i < numrows) Then 'untuk pastikan papar dalam range yang diselect sahaja
        i = i + 1
    Else
        i = 1
    End If

    Text1.Text = row(i, 1) ' data pada column 1 (name) dan row i
    Text2.Text = row(i, 0) ' data pada column 0 (ic) dan row i

End Sub

Private Sub Command2_Click()

    Dim crows As Variant ' current rows (prive variable)
    
    query = "insert into users (nama,ic) VALUES ('" + Text4.Text + "','" + Text3.Text + "')"
    crows = sqlite_get_table(DBz, query, minfo) ' query database
    
    If (minfo = "") Then
        MsgBox "Data berjaya di masukkan"
    Else
        MsgBox "Error: minfo"
    End If
    
    Call getData ' update latest data
    
End Sub
