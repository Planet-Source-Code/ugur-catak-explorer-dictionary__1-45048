VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IEExtension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'http://www.basic.tr.cx

Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SETTEXT = &HC


Private mvarURL As Variant 'local copy
Private mvarPageURL As Variant 'local copy
Private mvarSelText As Variant 'local copy
Private mvarLinkType As Variant 'local copy


Public Property Let LinkType(ByVal vData As Variant)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.LinkType = 5
    mvarLinkType = vData
End Property

Public Property Set LinkType(ByVal vData As Variant)
    'used when assigning an Object to the property, on the left side of a Set statement.
    'Syntax: Set x.LinkType = Form1
    Set mvarLinkType = vData
End Property

Public Property Get LinkType() As Variant
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.LinkType
    If IsObject(mvarLinkType) Then
        Set LinkType = mvarLinkType
    Else
        LinkType = mvarLinkType
    End If
End Property

Public Property Let SelText(ByVal vData As Variant)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.SelText = 5
    mvarSelText = vData
End Property

Public Property Set SelText(ByVal vData As Variant)
    'used when assigning an Object to the property, on the left side of a Set statement.
    'Syntax: Set x.SelText = Form1
    Set mvarSelText = vData
End Property

Public Property Get SelText() As Variant
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.SelText
    If IsObject(mvarSelText) Then
        Set SelText = mvarSelText
    Else
        SelText = mvarSelText
    End If
End Property

Public Property Let PageURL(ByVal vData As Variant)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.PageURL = 5
    mvarPageURL = vData
End Property

Public Property Set PageURL(ByVal vData As Variant)
    'used when assigning an Object to the property, on the left side of a Set statement.
    'Syntax: Set x.PageURL = Form1
    Set mvarPageURL = vData
End Property

Public Property Get PageURL() As Variant
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.PageURL
    If IsObject(mvarPageURL) Then
        Set PageURL = mvarPageURL
    Else
        PageURL = mvarPageURL
    End If
End Property

Public Property Let URL(ByVal vData As Variant)
    'used when assigning a value to the property, on the left side of an assignment.
    'Syntax: X.URL = 5
    mvarURL = vData
End Property

Public Property Set URL(ByVal vData As Variant)
    'used when assigning an Object to the property, on the left side of a Set statement.
    'Syntax: Set x.URL = Form1
    Set mvarURL = vData
End Property

Public Property Get URL() As Variant
    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.URL
    If IsObject(mvarURL) Then
        Set URL = mvarURL
    Else
        URL = mvarURL
    End If
End Property





Public Function ShowInfo()
Dim kayitsayisi As Long
Dim uzunluk As Integer
Dim lTarget_hWnd As Long
Dim degis As String
Dim secilimetin As String
 secilimetin = mvarSelText
 degis = mvarSelText
 StrSQL = "SELECT * FROM ugur"
 StrSQL = StrSQL & " WHERE a1='" & degis & "'"

frmMain.Data1.RecordSource = StrSQL
frmMain.Data1.Refresh
 
If frmMain.Text2 = "" Then
  StrSQL = "SELECT * FROM ugur"
  StrSQL = StrSQL & " WHERE a2='" & degis & "'"
  frmMain.Data1.RecordSource = StrSQL
  frmMain.Data1.Refresh
    If frmMain.Text1.Text = "" Then
      degis = secilimetin & " : " & " None!"
      GoTo labelim
    End If
    degis = secilimetin & " : " & frmMain.Text1.Text
    GoTo labelim
End If
    
  degis = secilimetin & " : " & frmMain.Text2.Text
    
labelim:
    lTarget_hWnd = GetActiveWindow
    SendMessage lTarget_hWnd, WM_SETTEXT, 0, ByVal degis
If degis = secilimetin & " : " & " None!" Then GoTo son
 Close #1
 Open "c:\hafiza.dat" For Random As #1 Len = Len(alanlar)
 kayitsayisi = LOF(1) / Len(alanlar)
 If kayitsayisi >= 10 Then
 Close #1
 Kill "C:\hafiza.dat"
 GoTo son
 End If
 kayitno = kayitsayisi + 1
 alanlar.ad = secilimetin & "?"
 alanlar.soyad = degis
 Put #1, kayitno, alanlar
 Close #1
son:
End Function

