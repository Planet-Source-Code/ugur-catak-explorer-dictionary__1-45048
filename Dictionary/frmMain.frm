VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IE Integration [Imported Information From IE Page]"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2865
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   2865
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      DataField       =   "a2"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1935
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\bdd.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ugur"
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      DataField       =   "a1"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
'Open "ogrenci.dat" For Random As #1 Len = Len(alanlar)
'kayitno = LOF(1) / Len(alanlar)
'If kayitno = 0 Then
'kayitno = 1
'End If
'KayitSayisi = DosUz / DesUz
'kayitno = KayitSayisi + 1


End Sub
