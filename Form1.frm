VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "IMG in MDB - By Zani"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   3840
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLoad 
      Caption         =   "LoadPic"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   3960
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   2280
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Prox 
      Caption         =   "Next >"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   3960
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "< Prev"
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   3960
      Width           =   735
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
   Begin VB.Frame Frame 
      Caption         =   "Picture"
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   3615
      Begin VB.Image img 
         Height          =   3015
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Top             =   3960
      Width           =   735
   End
   Begin VB.CommandButton cmdPut 
      Caption         =   "Save"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "Pic Title . :"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Con As ADODB.Connection
Public Rs As ADODB.Recordset
Public ID As Long

'Pic vars


Private Sub cmdLoad_Click()
CD.Filter = "JPGs...|*.jpg"
CD.ShowOpen
img.Picture = LoadPicture(CD.FileName)
End Sub

Private Sub cmdNew_Click()
ID = 0
txtName.Text = ""
img.Picture = Nothing
End Sub

Private Sub cmdPut_Click()
Rs.Close
Rs.Open "SELECT * FROM Clientes WHERE ID = " & ID, Con, 1, 2
If Rs.EOF Then Rs.AddNew
Rs!Nome = txtName.Text

'Put da pic
File2Field CD.FileName, Rs!Foto
' pick is in db

Rs.Update
ID = Rs!ID
Rs.Close
Rs.Open "SELECT * FROM Clientes", Con, 1, 2
End Sub

Private Sub Command1_Click()
If Rs.BOF Then Exit Sub
Rs.MovePrevious
LoadData
End Sub

Private Sub Form_Load()
Set Con = New ADODB.Connection
Con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\data.mdb;Persist Security Info=False"
Set Rs = New ADODB.Recordset
Rs.Open "SELECT * FROM Clientes", Con, 1, 2
LoadData
End Sub
Sub LoadData()
If Rs.EOF = False And Rs.BOF = False Then
ID = Rs!ID
txtName.Text = IIf(IsNull(Rs!Nome), "", Rs!Nome)
Field2File App.Path & "\TempPic", Rs!Foto
img.Picture = LoadPicture(App.Path & "\TempPic")
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Dir(App.Path & "\TempPic") <> "" Then Kill App.Path & "\TempPic"
End Sub

Private Sub Prox_Click()
If Rs.EOF Then Exit Sub
Rs.MoveNext
LoadData
End Sub
