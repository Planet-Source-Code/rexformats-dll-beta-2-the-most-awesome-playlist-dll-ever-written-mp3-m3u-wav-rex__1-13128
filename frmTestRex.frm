VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "REX Format - Beta 2"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   9255
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReadAll 
      Caption         =   "Read all records"
      Height          =   375
      Left            =   7320
      TabIndex        =   12
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton cmdWriteMp3ToRex 
      Caption         =   "Write MP3 to Rex"
      Height          =   375
      Left            =   7320
      TabIndex        =   11
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton cmdMP3 
      Caption         =   "Read a MP3 file"
      Height          =   375
      Left            =   7320
      TabIndex        =   10
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton cmdOpenRexFile 
      Caption         =   "Open Rex File"
      Height          =   375
      Left            =   7320
      TabIndex        =   8
      Top             =   600
      Width           =   1815
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3495
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   6165
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Result"
         Object.Width           =   9701
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "m3u Time"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdM3U 
      Caption         =   "&Read a m3u file"
      Height          =   375
      Left            =   7320
      TabIndex        =   4
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find Data"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7320
      TabIndex        =   3
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton cmdWriteSampleData 
      Caption         =   "&Write Sample Data"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7320
      TabIndex        =   2
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      Height          =   375
      Left            =   7320
      TabIndex        =   1
      Top             =   4800
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   1080
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCreateRex 
      Caption         =   "&Create Rex File"
      Height          =   375
      Left            =   7320
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   $"frmTestRex.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   4080
      Width           =   7095
   End
   Begin VB.Label Label2 
      Caption         =   "RexFormats.DLL supports m3u files, try it."
      Height          =   495
      Left            =   7320
      TabIndex        =   7
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "RexFormats Driver 1.0 - Written by Sveinn R. Sigurdss (MrHippo) (C) 2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   7095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private r As New SoundFormats ' First, let's declare the RexFormat DLL
Private rexFile As String ' A static variable for the rex file location



' Create a rex PlayList file
Private Sub cmdCreateRex_Click()
        ' Now, let's create the RexFile
        cd.Filter = "Rex File Format (*.rex)|*.rex"
        cd.DefaultExt = "Untitled.rex|*.rex"
        cd.ShowSave
        r.FileName = cd.FileName
        rexFile = cd.FileName
        r.rex.CreateRexFile
        r.rex.Initialize
        
        cmdWriteSampleData.Enabled = True
        cmdFind.Enabled = True
End Sub



' Now, find data from the file we've previously created
Private Sub cmdFind_Click()
    Dim i As Long

    ' This clears all possible search criteries
    r.rex.NewSearch
    ' Simply set desired criteria
    r.rex.Find.Album = "Great"
    r.rex.Find.Artist = "Var"
    ' Find the result
    r.rex.Search
    
    If r.rex.RecordCount <> -1 Then
        ListView1.ListItems.Clear
        For i = 0 To lResults
            ' This will read all properties properties from the
            ' selected row into the record object
            r.rex.ReadRecord (i)
            Call DIsplayRow
        Next i
    End If
End Sub


' Read a Winamp m3u file
Private Sub cmdM3U_Click()
    Dim Count As Long
    Dim x As Long
        
    cd.Filter = "Winamp m3u File (*.m3u)|*.m3u"
    cd.DefaultExt = "*.m3u"
    cd.ShowOpen
    r.FileName = cd.FileName

    r.m3u.Readm3u (r.FileName)
    Count = r.m3u.Count
   
    ' Dim FilePath$, tmpString$, i%, FindComma%
    ListView1.ListItems.Clear
    For x = 0 To r.m3u.Count - 1
        ListView1.ListItems.Add , , r.m3u.FileTitle(x)
        ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , r.m3u.FileSeconds(x) & " seconds"
    Next x
End Sub



' Read in a mp3 file
Private Sub cmdMP3_Click()
    cd.DialogTitle = "Open Existing mp3 File"
    cd.Filter = "MP3 Music format (*.mp3)|*.mp3"
    cd.DefaultExt = "*.mp3"
    cd.ShowOpen
    If cd.FileName <> "" Then
        r.FileName = cd.FileName
        ' Now let's read in all the tags
        r.mp3.ReadMP3Info
        r.mp3.ReadHeader
        
        Call ReadMP3Tags
        
        cmdWriteSampleData.Enabled = True
        cmdFind.Enabled = True
    End If
End Sub



' Open an existing rex format file
Private Sub cmdOpenRexFile_Click()
    cd.DialogTitle = "Open Existing Rex File"
    cd.Filter = "Rex File Format (*.rex)|*.rex"
    cd.DefaultExt = "*.rex"
    cd.ShowOpen
    
    If cd.FileName <> "" Then
        r.FileName = cd.FileName
        r.rex.Initialize
        rexFile = r.FileName ' I'm storing the location of the file as static
                             ' this comes in handy when we are writing other
                             ' file formats to rex.
        cmdWriteSampleData.Enabled = True
        cmdFind.Enabled = True
    End If
End Sub



Private Sub cmdReadAll_Click()
    Dim i As Long
    Call r.rex.LoadAllRecords
    
    With r.rex
        If .RecordCount > 0 Then
            ListView1.ListItems.Clear
            For i = 0 To .RecordCount - 1
                r.rex.ReadRecord (i)
                ListView1.ListItems.Add , , "Title := " & r.rex.record.Title
            Next i
        End If
    End With
            
End Sub

' Write a sample Data to the file
Private Sub cmdWriteSampleData_Click()

        ' You can also set Global preferences in the file
        ' by using the Preferences attributes
        ' Note :
        '       - EnablePreviews enables the dll to cut
        '         a 5 - 15 second clip from the file
        '         and store it in the rexFile.
        '       - AllowDuplicates True/False (Boolean)
        '         If you do not want the same song to
        '         be written twice to the database
        '         set the property to false
        r.rex.Preferences.Security_PasswordProtected = False
        r.rex.Preferences.Security_Password = "anna79"
        r.rex.Preferences.Security_Username = "svenni76"
        r.rex.Preferences.EnablePreviews = True
        r.rex.Preferences.AllowDuplicates = False
        r.rex.Preferences.Author_Address = "Sunnuflöt 42"
        r.rex.Preferences.Author_Name = "Sveinn R. Sigurðsson"
        r.rex.Preferences.Author_Gender = Mr
        r.rex.Preferences.Author_Country = ICELAND

        ' Now, let's write a single record to the new file
        ' Notice that the filename & path fields must be separated
        r.rex.record.Album = "Johny Be Good"
        r.rex.record.Title = "Hello"
        r.rex.record.FileName = "TestMusic.mp3"
        r.rex.record.Path = "C:\"
        r.rex.record.ArtistWebsite = "www.svenni.com"
        r.rex.record.Genre = Pop
        r.rex.record.Rate = r75
        r.rex.record.Artist = "Sveinn R. Sigurðsson"
        ' Save the song to the file
        r.rex.Add
        
        ' Write another song to the file
        ' Notice that now we are writing to another field
        r.rex.record.Album = "Great Expectations"
        r.rex.record.Title = "Mono - Live in Mono"
        r.rex.record.FileName = "mono - live in mono.mp3"
        r.rex.record.Path = "C:\music"
        r.rex.record.Comments = "No Comments at all"
        r.rex.record.Genre = Pop
        r.rex.record.Producer = "Sveinn R. Sigurdsson"
        r.rex.record.Artist = "Various"
        r.rex.record.Arrangement = "Alan Silvestri"
        r.rex.record.ArtistWebsite = "www.greatexpectations.com"
        r.rex.record.AttributesLastAccessed = "24.12.1999"
        r.rex.record.AttributesLastModified = "23.12.1999"
        r.rex.record.AttributesReadOnly = "False"
        r.rex.record.AudioSiteUrl = "www.amazone.com"
        r.rex.record.CDLabel = "Great Expectations - the CD"
        ' Save the song to the file
        r.rex.Add
        
End Sub



Public Sub DIsplayRow()
    With ListView1
        .ListItems.Add , , "Album := " & r.rex.record.Album
        .ListItems.Add , , "Arrangement := " & r.rex.record.Arrangement
        .ListItems.Add , , "Artist := " & r.rex.record.Artist
        .ListItems.Add , , "ArtistWebsite := " & r.rex.record.ArtistWebsite
        .ListItems.Add , , "LastAccessed := " & r.rex.record.AttributesLastAccessed
        .ListItems.Add , , "LastModified := " & r.rex.record.AttributesLastModified
        .ListItems.Add , , "ReadOnly := " & r.rex.record.AttributesReadOnly
        .ListItems.Add , , "AudioSiteURL := " & r.rex.record.AudioSiteUrl
        .ListItems.Add , , "ArtistBiography := " & r.rex.record.Biography
        .ListItems.Add , , "BuyCDUrl := " & r.rex.record.BuyCDUrl
        .ListItems.Add , , "CD Friendly Name := " & r.rex.record.CDLabel
        .ListItems.Add , , "Comments := " & r.rex.record.Comments
        .ListItems.Add , , "Conductor := " & r.rex.record.Conductor
        .ListItems.Add , , "AlbumImage := " & r.rex.record.CoverImage
        .ListItems.Add , , "Distribution := " & r.rex.record.Distribution
        .ListItems.Add , , "Engineer := " & r.rex.record.Engineer
        .ListItems.Add , , "FanSiteUrl := " & r.rex.record.FanSite
        .ListItems.Add , , "Filename := " & r.rex.record.FileName
        .ListItems.Add , , "FileSize := " & r.rex.record.FileSize
        .ListItems.Add , , "Genre := " & r.rex.record.Genre
        .ListItems.Add , , "LastPlayed := " & r.rex.record.LastPlayedDate
        .ListItems.Add , , "Length := " & r.rex.record.Length
        .ListItems.Add , , "Lyrics := " & r.rex.record.Lyrics
        .ListItems.Add , , "Mood := " & r.rex.record.Mood
        .ListItems.Add , , "Notes provided by := " & r.rex.record.MusicNotesProvidedBy
        .ListItems.Add , , "Notes := " & r.rex.record.Notes
        .ListItems.Add , , "Original := " & r.rex.record.Original
        .ListItems.Add , , "Path := " & r.rex.record.Path
        .ListItems.Add , , "Preferences := " & r.rex.record.Preference
        .ListItems.Add , , "Preview := " & r.rex.record.Preview
        .ListItems.Add , , "Producer := " & r.rex.record.Producer
        .ListItems.Add , , "Rate := " & r.rex.record.Rate
        .ListItems.Add , , "RecordedAt := " & r.rex.record.RecordedAt
        .ListItems.Add , , "Situation := " & r.rex.record.Situation
        .ListItems.Add , , "Bitrate := " & r.rex.record.SongBitrate
        .ListItems.Add , , "Copyright := " & r.rex.record.SongCopyright
        .ListItems.Add , , "CRC := " & r.rex.record.SongCRC
        .ListItems.Add , , "Duration := " & r.rex.record.SongDuration
        .ListItems.Add , , "Frequency := " & r.rex.record.SongFrequency
        .ListItems.Add , , "Layer := " & r.rex.record.SongLayer
        .ListItems.Add , , "Mode := " & r.rex.record.SongMode
        .ListItems.Add , , "Padding := " & r.rex.record.SongPadding
        .ListItems.Add , , "Private := " & r.rex.record.SongPrivate
        .ListItems.Add , , "Version := " & r.rex.record.SongVersion
        .ListItems.Add , , "Volume := " & r.rex.record.SongVolume
        .ListItems.Add , , "Studio := " & r.rex.record.Studio
        .ListItems.Add , , "Tempo := " & r.rex.record.Tempo
        .ListItems.Add , , "Title := " & r.rex.record.Title
        .ListItems.Add , , "TrackNumber := " & r.rex.record.TrackNumber
        .ListItems.Add , , "year := " & r.rex.record.Year
   End With
End Sub



Private Sub ReadMP3Tags()
    With ListView1
        .ListItems.Clear
        .ListItems.Add , , "Album := " & r.mp3.Album
        .ListItems.Add , , "Artist := " & r.mp3.Artist
        .ListItems.Add , , "Bitrate := " & r.mp3.BitRate
        .ListItems.Add , , "ChannelMode := " & r.mp3.ChannelMode
        .ListItems.Add , , "Comment := " & r.mp3.Comment
        .ListItems.Add , , "Copyright := " & r.mp3.Copyright
        .ListItems.Add , , "CRC Present :=" & r.mp3.CRCPresent
        .ListItems.Add , , "Emphasis := " & r.mp3.Emphasis
        .ListItems.Add , , "File Attributes := " & r.mp3.FileAttributes
        .ListItems.Add , , "FileName :=" & r.mp3.FileName
        .ListItems.Add , , "Framelength := " & r.mp3.FrameLength
        .ListItems.Add , , "FullName := " & r.mp3.FullName
        .ListItems.Add , , "Genre := " & r.mp3.Genre
        .ListItems.Add , , "Layer := " & r.mp3.Layer
        .ListItems.Add , , "Mode Extension := " & r.mp3.ModeExtension
        .ListItems.Add , , "MPEG Version := " & r.mp3.MPEGVersion
        .ListItems.Add , , "Original := " & r.mp3.Original
        .ListItems.Add , , "Padding := " & r.mp3.Padding
        .ListItems.Add , , "Path := " & r.mp3.Path
        .ListItems.Add , , "PlayTime := " & r.mp3.PlayTime
        .ListItems.Add , , "PrivateBits := " & r.mp3.PrivateBit
        .ListItems.Add , , "SampleRate :=" & r.mp3.SampleRate
        .ListItems.Add , , "TagPresent := " & r.mp3.TagPresent
        .ListItems.Add , , "Title := " & r.mp3.Title
        .ListItems.Add , , "TotalFrames := " & r.mp3.TotalFrames
        .ListItems.Add , , "ValidHeader := " & r.mp3.ValidHeader
        .ListItems.Add , , "Year := " & r.mp3.Year
    End With
End Sub



' This demonstrates how to use the RexFormats.DLL to validate information
' and write them to another format using only the DLL
Private Sub cmdWriteMp3ToRex_Click()
    
    ' Let's start by opening a mp3 file
    cd.DialogTitle = "Open Existing MP3 File"
    cd.Filter = "MP3 sound Format (*.mp3)|*.mp3"
    cd.DefaultExt = "*.mp3"
    cd.ShowOpen
    
    If cd.FileName <> "" Then
        r.FileName = cd.FileName
        ' Initalize the mp3 engine
        r.mp3.ReadHeader
        r.mp3.ReadMP3Info
        ' Now I retrieve the stored value of the rexFilename
        r.FileName = rexFile
        ' Initialize the rex engine
        r.rex.Initialize
        ' Now let's read in the information
    Else
        Exit Sub
    End If
    
    With r.rex
        ListView1.ListItems.Clear
        
        .record.ClearFields ' Clear all data fram memory
        .record.Album = r.mp3.Album
        .record.Artist = r.mp3.Artist
        .record.SongBitrate = r.mp3.BitRate
        .record.Comments = r.mp3.Comment
        .record.SongCopyright = r.mp3.Copyright
        .record.SongCRC = r.mp3.CRCPresent
        .record.SongEmpasis = r.mp3.Emphasis
        .record.AttributesReadOnly = r.mp3.FileAttributes
        .record.FileName = r.mp3.FileName
        .record.Genre = r.mp3.Genre
        .record.SongLayer = r.mp3.Layer
        .record.SongMode = r.mp3.ModeExtension
        .record.Original = r.mp3.Original
        .record.SongPadding = r.mp3.Padding
        .record.Path = r.mp3.Path
        .record.SongDuration = r.mp3.PlayTime
        .record.SongPrivate = r.mp3.PrivateBit
        .record.Title = r.mp3.Title
        .record.Year = r.mp3.Year
        ' Of course you can also add tags to other fields that the mp3 file
        ' format does not store. F.ex. the user enters the url for the
        ' artists website, then you could add this also by simply writing
        .record.ArtistWebsite = "www.johnlennon.com"
        .Add ' Save this data to rex
    End With
End Sub









