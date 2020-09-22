VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Zip Code to GPS Co-Ordinates"
   ClientHeight    =   7290
   ClientLeft      =   4320
   ClientTop       =   3510
   ClientWidth     =   6525
   Icon            =   "LatLong.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   6525
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   375
      Left            =   5520
      TabIndex        =   9
      Top             =   6840
      Width           =   975
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   6840
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3975
      Left            =   0
      TabIndex        =   7
      Top             =   2760
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   7011
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ZipCode"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "City(STATE)"
         Object.Width           =   4233
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Latitude"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Longitude  "
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.ListBox List1 
      Height          =   4350
      Left            =   120
      TabIndex        =   6
      Top             =   9000
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Get All"
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   6840
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   2175
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "LatLong.frx":1042
      Top             =   570
      Width           =   6495
   End
   Begin VB.TextBox txtLong 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox txtLatt 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   2760
      Top             =   7800
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1080
      MaxLength       =   5
      TabIndex        =   1
      Text            =   "84047"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuViewExtend 
         Caption         =   "View Extended"
      End
      Begin VB.Menu mnuViewMin 
         Caption         =   "View Minimal"
      End
   End
   Begin VB.Menu MnuGet 
      Caption         =   "Get"
      Begin VB.Menu mnuGetIndividual 
         Caption         =   "Get Individual"
      End
      Begin VB.Menu MnuGetList 
         Caption         =   "Get List"
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "Settings"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim startscrape As Integer
Dim strHTML, inetsource, parsedstring As String
Dim Stopnow As Boolean


Private Sub cmdStop_Click()
Stopnow = True
End Sub

Private Sub Command1_Click()
        strHTML = Inet1.OpenURL("http://www.census.gov/cgi-bin/gazetteer?city=&state=&zip=" & Text1.Text)
        inetsource = LCase(strHTML)
'parse throught text to find the word "location: "
        startscrape = InStr(1, inetsource, "location: ")
'take the starting point of " Location: " and add 10 characters to get the starting
'point for the latitude, go for 25 characters to end with longtitude
        parsedstring = Mid(inetsource, (startscrape + 10), 25)
        latlong = Split(parsedstring, ",")
        txtLatt = latlong(0)
        txtLong = latlong(1)
End Sub

Private Sub Command2_Click()
getlatlong
Stopnow = False
End Sub

Public Function getlatlong()

ziparray = Split(Text2.Text, Chr(44))

For i = 0 To UBound(ziparray)
'code to stop the loop... kinda
If Stopnow = False Then

ProgressBar1.Value = (((i + 1) * 100) / (UBound(ziparray) + 1))
Me.Caption = "Zip Code to GPS Co-Ordinates " & Left((((i + 1) * 100) / (UBound(ziparray) + 1)), 3) & "%"

    Text1.Text = ziparray(i)
        strHTML = Inet1.OpenURL("http://www.census.gov/cgi-bin/gazetteer?city=&state=&zip=" & ziparray(i))
    Do While Inet1.StillExecuting = True
        DoEvents
    Loop
    
        inetsource = LCase(strHTML)
        
'parse throught text to find the word "po name: "
        startnamescrape = InStr(1, inetsource, "po name: ")
        endnamescrape = InStr(1, inetsource, ")<br>")
        
'parse throught text to find the word "location: "
        startscrape = InStr(1, inetsource, "location: ")
        endscrape = InStr(1, inetsource, " w<br>")

'take the starting point of " Location: " and add 10 characters to get the starting
'point for the latitude, go for 25 characters to end with longtitude
'if i cant find information mark parsedstring as UNKNOWN
If startscrape <= 1 Then
parsedstring = "UNINCORPORATED, No USGS Info"
Else
        Cityname = Mid(inetsource, (startnamescrape + 9), (endnamescrape - (startnamescrape + 9) + 1))
        parsedstring = Mid(inetsource, (startscrape + 10), (endscrape - (startscrape + 10) + 2))
        On Error Resume Next
       latlong = Split(parsedstring, ",")
       txtLatt = latlong(0)
       txtLong = latlong(1)
End If

'add the zip, lat and long to a list box
    List1.AddItem (Text1.Text & "," & UCase(Cityname) & "," & txtLatt.Text & "," & txtLong.Text)
    AddItem ListView1, Text1.Text, UCase(Cityname), txtLatt.Text, txtLong.Text

Else
'do nothing if else

End If

Next i

End Function

Public Sub SaveListBox(TheList As ListBox, Directory As String)
    Dim SaveList As Long
    On Error Resume Next
    Open Directory$ For Output As #1

    For SaveList& = 0 To TheList.ListCount - 1
        Print #1, TheList.List(SaveList&)
    Next SaveList&
    Close #1
End Sub

'add items to listview specify listview, items csv format.
Sub AddItem(lst As ListView, ParamArray Items())
    Set a = lst.ListItems.Add(, , Items(0))
    For i = 1 To UBound(Items())
        a.SubItems(i) = Items(i)
    Next i
End Sub

'Drop menu Code start here ------------------------------------------
Private Sub mnuSave_Click()
'Example: Call LoadListBox(list1, "C:\Temp\MyList.dat")
Call SaveListBox(List1, App.Path & "\ziplist.txt")
End Sub
