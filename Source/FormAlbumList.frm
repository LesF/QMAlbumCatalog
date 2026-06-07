VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormAlbumList 
   Caption         =   "QM Album Catalog"
   ClientHeight    =   5550
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   7665
   Icon            =   "FormAlbumList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   5175
      Width           =   7665
      _ExtentX        =   13520
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "Not Connected"
            TextSave        =   "Not Connected"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10425
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LVCatalog 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   8493
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "artist"
         Text            =   "Artist"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "album"
         Text            =   "Album"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "mediatype"
         Text            =   "Media Type"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu menuFile 
      Caption         =   "File"
      Begin VB.Menu menuFileSettings 
         Caption         =   "Settings"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu menuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu menuSession 
      Caption         =   "Session"
      Begin VB.Menu menuSessionConnect 
         Caption         =   "Connect"
      End
      Begin VB.Menu menuSessionDisconnect 
         Caption         =   "Disconnect"
      End
   End
   Begin VB.Menu menuAlbum 
      Caption         =   "Album"
      Begin VB.Menu menuAlbumEdit 
         Caption         =   "Edit Selected"
      End
      Begin VB.Menu menuAlbumNew 
         Caption         =   "New Album"
      End
   End
End
Attribute VB_Name = "FormAlbumList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-----
Dim isQMConnected As Boolean
Private m_LastX As Single
Private m_LastY As Single
'-----

Private Sub Form_Load()
    ConfigFolder = App.title
    isQMConnected = False
    SetConnectState False
    RetrieveConfig
    
    If Len(SVC_Host) > 0 And Len(SVC_User) > 0 Then
        'We have retrieved saved settings, try to connect
        If ConnectQM() Then RetrieveCatalog
    End If
    If Not isQMConnected Then StatusBar1.Panels.Item(2).Text = "Configure settings then select Session/Connect"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    QMDisconnectAll
End Sub

Private Sub Form_Resize()
    '-----
    Dim pad As Integer, listH As Integer, listW As Integer
    Dim clientW As Integer, clientH As Integer
    '-----
    If FormAlbumList.WindowState = vbMinimized Then Exit Sub
    clientW = FormAlbumList.ScaleWidth
    clientH = FormAlbumList.ScaleHeight - StatusBar1.Height
    pad = LVCatalog.Left
    If clientW < (pad * 5) Then Exit Sub
    If clientH < (pad * 10) Then Exit Sub
    '-----
    listW = clientW - (pad * 2)
    If listW < (pad * 6) Then listW = LVCatalog.Width
    listH = clientH - (LVCatalog.Top + pad)
    If listH < (pad * 4) Then listH = LVCatalog.Height
    LVCatalog.Move LVCatalog.Left, LVCatalog.Top, listW, listH
End Sub

Private Sub menuAlbumEdit_Click()
    Dim albumId As String
    Dim dlg As FormEditAlbum
    If Not LVCatalog.SelectedItem Is Nothing Then
        albumId = LVCatalog.SelectedItem.Key
        Debug.Print "Edit '" & albumId & "'"
        Set dlg = New FormEditAlbum
        dlg.albumId = albumId
        dlg.Show vbModal, Me
        RetrieveCatalog
    End If
End Sub

Private Sub menuAlbumNew_Click()
    FormEditAlbum.albumId = "_new_"
    FormEditAlbum.Show vbModal, Me
    RetrieveCatalog
End Sub

Private Sub menuFileExit_Click()
    QMDisconnectAll
    Unload Me
End Sub

Private Sub menuFileSettings_Click()
    ShowConfig
End Sub

Private Sub menuSessionConnect_Click()
    If ConnectQM() Then RetrieveCatalog
End Sub

Private Sub menuSessionDisconnect_Click()
    QMDisconnectAll
    isQMConnected = False
    StatusBar1.Panels.Item(1).Text = "Not Connected"
    SetConnectState False
End Sub

Public Sub ShowConfig()
    FormServiceConfig.Show vbModal, Me
    SVC_Host = IniRead("Connection", "Host", "undefined")
    SVC_Port = CLng(IniRead("Connection", "Port", "4243"))
    SVC_User = IniRead("Connection", "User", "")
    SVC_Pass = IniRead("Connection", "Password", "")
    SVC_Account = IniRead("Connection", "Account", "")
End Sub

Public Sub RetrieveConfig()
    'Retrieve config settings.  If not found pop up the settings dialog.
    SVC_Host = IniRead("Connection", "Host", "undefined")
    If Len(SVC_Host) = 0 Or SVC_Host = "undefined" Then
        ShowConfig
    Else
        SVC_Port = CLng(IniRead("Connection", "Port", "4243"))
        SVC_User = IniRead("Connection", "User", "")
        SVC_Pass = IniRead("Connection", "Password", "")
        SVC_Account = IniRead("Connection", "Account", "")
    End If
End Sub

Private Function ConnectQM() As Boolean
    '-----
    Dim gotConnect As Boolean
    '-----
    On Error GoTo ConnectError
    QMDisconnectAll
    isQMConnected = False
    StatusBar1.Panels.Item(1).Text = "Connecting"
    StatusBar1.Panels.Item(2).Text = ""
    LVCatalog.Enabled = False
    gotConnect = QMConnect(SVC_Host, SVC_Port, SVC_User, SVC_Pass, SVC_Account)
    If gotConnect Then
        If QMConnected() <> 0 Then
            StatusBar1.Panels.Item(1).Text = "Connected"
            isQMConnected = True
            SetConnectState True
            ConnectQM = True
        Else
            StatusBar1.Panels.Item(1).Text = "Connected: false"
        End If
    Else
        StatusBar1.Panels.Item(1).Text = "Connect failed"
    End If
    Exit Function
ConnectError:
    'might fit something useful into the status panel?
    StatusBar1.Panels.Item(1).Text = "Connect failed"
    StatusBar1.Panels.Item(2).Text = "Connection error: " & Err.Description
End Function

Private Function SetConnectState(isConnected As Boolean)
    LVCatalog.Enabled = isConnected
    menuAlbum.Enabled = isConnected
    If Not isConnected Then LVCatalog.ListItems.Clear
End Function

Private Sub RetrieveCatalog()
    '-----
    Dim albumsFile As Integer, artistsFile As Integer
    Dim albumId As String, rec As String
    Dim artistId As String, artistName As String, artistRec As String
    Dim title As String, mediaType As String
    Dim errNo As Integer
    Dim Item As ListItem
    '-----
    StatusBar1.Panels.Item(2).Text = "Retrieving catalog"
    On Error GoTo ErrorHandler
    '-----
    LVCatalog.ListItems.Clear
    
    albumsFile = QMOpen("ALBUMS")
    If albumsFile = 0 Then
        StatusBar1.Panels.Item(2).Text = "Error: Could not open ALBUMS table"
        Exit Sub
    End If
    
    artistsFile = QMOpen("ARTISTS")
    
    QMSelect albumsFile, 0
    
    Do
        albumId = QMReadNext(0, errNo)
        If errNo = SV_ELSE Then Exit Do    ' no more records in select list
        
        rec = QMRead(albumsFile, albumId, errNo)
        If errNo <> SV_OK Then GoTo NextAlbum
        
        ' Schema: Field 1 = ARTIST.ID, Field 2 = TITLE, Field 7 = MEDIA.TYPE
        artistId = QMExtract(rec, 1, 0, 0)
        title = QMExtract(rec, 2, 0, 0)
        mediaType = QMExtract(rec, 7, 0, 0)
        
        ' Look up artist name from ARTISTS table (Field 1 = NAME)
        artistName = artistId  ' fallback: show ID if lookup fails
        If artistsFile > 0 And Len(artistId) > 0 Then
            artistRec = QMRead(artistsFile, artistId, errNo)
            If errNo = SV_OK Then
                artistName = QMExtract(artistRec, 1, 0, 0)
            End If
        End If
        
        ' ARTIST.ID is multivalued - show first value only
        If InStr(artistName, Chr(253)) > 0 Then
            artistName = Left(artistName, InStr(artistName, Chr(253)) - 1)
        End If
        
        'Add(index, key, text)
        Set Item = LVCatalog.ListItems.Add(, albumId, artistName)
        Item.SubItems(1) = title
        Item.SubItems(2) = mediaType

NextAlbum:
    Loop
    
    If artistsFile > 0 Then QMClose artistsFile
    QMClose albumsFile
    Call AutoSizeListViewColumns(LVCatalog, False)
    Call AutoSizeListViewColumnsFillLast(LVCatalog)
    '-----
    StatusBar1.Panels.Item(2).Text = "Catalog items: " & LVCatalog.ListItems.Count
    Exit Sub
ErrorHandler:
    StatusBar1.Panels.Item(2).Text = "Error: " & Err.Description
End Sub

Private Sub LVCatalog_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    m_LastX = x
    m_LastY = y
End Sub

Private Sub LVCatalog_DblClick()
    Dim itm As ListItem

    Set itm = LVCatalog.HitTest(m_LastX, m_LastY)
    If itm Is Nothing Then
        ' User double-clicked empty space — ignore
        Exit Sub
    End If
    ' Otherwise, they clicked a real row
    Call menuAlbumEdit_Click
End Sub


