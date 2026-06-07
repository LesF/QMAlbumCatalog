VERSION 5.00
Begin VB.Form FormEditAlbum 
   Caption         =   "Edit Album"
   ClientHeight    =   7095
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6165
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   6165
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox TextArtist 
      Height          =   615
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Top             =   2520
      Width           =   5895
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "cancel"
      Height          =   375
      Left            =   4920
      TabIndex        =   14
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   3600
      TabIndex        =   13
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   12
      Top             =   6600
      Width           =   1095
   End
   Begin VB.TextBox TextTrackLength 
      Height          =   735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   5640
      Width           =   5895
   End
   Begin VB.TextBox TextTrackList 
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   3720
      Width           =   5895
   End
   Begin VB.ComboBox ComboMedia 
      Height          =   315
      ItemData        =   "FormEditAlbum.frx":0000
      Left            =   1320
      List            =   "FormEditAlbum.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox TextYear 
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox TextTitle 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   720
      Width           =   4575
   End
   Begin VB.TextBox TextID 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Artist(s) / band name"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   2160
      Width           =   1470
   End
   Begin VB.Label Label6 
      Caption         =   "Track times (semi-colon delimited)"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   5280
      Width           =   4575
   End
   Begin VB.Label Label5 
      Caption         =   "Track names (semi-colon delimited)"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3360
      Width           =   3615
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Media Type"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Release Year"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Title"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Album ID"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   645
   End
End
Attribute VB_Name = "FormEditAlbum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----
Option Explicit
'-----
Private albumsFile As Integer
'-----

Public albumId As String

Private Sub Form_Load()
    ' Dialog mode; load record using pre-populated album.id?
    
    albumsFile = QMOpen("ALBUMS")
    If albumsFile = 0 Then
        ' StatusBar1.Panels.Item(2).Text = "Error: Could not open ALBUMS table"
        Exit Sub
    End If

    If Len(albumId) = 0 Or albumId = "_new_" Then
        'Anything to prepare for a new record?
        ComboMedia.ListIndex = 1
    Else
        LoadAlbum
    End If
End Sub

Private Sub CmdCancel_Click()
    'Nothing to do
    Unload Me
End Sub

Private Sub CmdDelete_Click()
    'TODO delete Album record
    Unload Me
End Sub

Private Sub CmdSave_Click()
    'TODO create or update Albu
    Unload Me
End Sub

Private Sub Form_Resize()
    '-----
    Dim pad As Integer, clientW As Integer, clientH As Integer
    Dim minW As Integer, minH As Integer, controlW As Integer, controlH As Integer, controlTop As Integer
    '-----
    ' I hate non-resizable dialogs so here goes...
    If Me.WindowState = vbMinimized Then Exit Sub
    clientW = Me.ScaleWidth
    clientH = Me.ScaleHeight
    pad = TextTrackList.Left
    minW = (CmdCancel.Width * 3) + (pad * 4)
    minH = TextTrackLength.Top + (CmdCancel.Height * 3)
    '-----
    controlW = clientW - (TextTitle.Left + pad)
    If controlW > pad Then TextTitle.Width = controlW
    controlW = clientW - (TextTrackList.Left + pad)
    If controlW > pad Then
        TextArtist.Width = controlW
        TextTrackList.Width = controlW
        TextTrackLength.Width = controlW
    End If
    
    If clientW > minW Then
        CmdCancel.Left = clientW - (CmdCancel.Width + pad)
        CmdDelete.Left = CmdCancel.Left - (CmdDelete.Width + pad)
        CmdSave.Left = CmdDelete.Left - (CmdSave.Width + pad)
    End If
    
    controlH = clientH - (TextTrackLength.Top + CmdSave.Height + (pad * 2))
    If controlH < pad Then controlH = pad
    TextTrackLength.Height = controlH
    
    controlTop = clientH - (CmdSave.Height + pad)
    If controlTop > (TextTrackLength.Top + TextTrackLength.Height) Then
        CmdSave.Top = controlTop
        CmdDelete.Top = controlTop
        CmdCancel.Top = controlTop
    End If
End Sub

Private Sub LoadAlbum()
    '-----
    Dim rec As String
    Dim errNo As Integer
    Dim artistId As String, artistName As String 'artistRec As String
    Dim title As String, mediaType As String, albumYear As String
    Dim trackNames As String, trackTimes As String, genres As String
    Dim mediaCount As Integer, mediaIdx As Integer
    '-----
    rec = QMRead(albumsFile, albumId, errNo)
    If errNo <> SV_OK Then
        
    Else
        ' Schema: Field 1 = ARTIST.ID, Field 2 = TITLE, Field 3 = Year, Field 4 = Track list, Field 5 = Track lengths, Field 6 = Genres, Field 7 = MEDIA.TYPE
        artistId = QMExtract(rec, 1, 0, 0)
        title = QMExtract(rec, 2, 0, 0)
        albumYear = QMExtract(rec, 3, 0, 0)
        trackNames = QMExtract(rec, 4, 0, 0)
        trackTimes = QMExtract(rec, 5, 0, 0)
        genres = QMExtract(rec, 6, 0, 0)
        mediaType = QMExtract(rec, 7, 0, 0)
        'TODO match mediaType to a dropdown list item
        
        'TODO Look up artist name from ARTISTS table (Field 1 = NAME)
'        artistName = artistId  ' fallback: show ID if lookup fails
'        If artistsFile > 0 And Len(artistId) > 0 Then
'            artistRec = QMRead(artistsFile, artistId, errNo)
'            If errNo = SV_OK Then
'                artistName = QMExtract(artistRec, 1, 0, 0)
'            End If
'        End If
        '-----
        TextID.Text = albumId
        TextID.Enabled = False
        TextTitle.Text = title
        TextYear.Text = albumYear
        TextArtist.Text = Replace(artistId, Chr(253), "; ")
        TextTrackList.Text = Replace(trackNames, Chr(253), "; ")
        TextTrackLength.Text = Replace(trackTimes, Chr(253), "; ")
        
        If Len(mediaType) > 0 Then
            ComboMedia.ListIndex = 4   'Other?
            mediaCount = ComboMedia.ListCount - 1
            For mediaIdx = 0 To mediaCount
                If LCase(ComboMedia.ItemData(mediaIdx)) = LCase(mediaType) Then
                    ComboMedia.ListIndex = mediaIdx
                    Exit For
                End If
            Next mediaIdx
        Else
            ComboMedia.ListIndex = 0
        End If
    End If
End Sub
