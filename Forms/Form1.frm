VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Reverb Tests"
   ClientHeight    =   3375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox TxFilename 
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   2760
      Width           =   6975
   End
   Begin VB.CommandButton BtnSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   2760
      TabIndex        =   7
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton BtnPlayDry 
      Caption         =   "Dry"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton BtnStop 
      Caption         =   "Stop"
      Height          =   495
      Left            =   1440
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton BtnPlayFreeVerb 
      Caption         =   "FreeVerb"
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   840
      Width           =   2655
   End
   Begin VB.ComboBox CbGVerb 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   480
      Width           =   2655
   End
   Begin VB.CommandButton BtnPlayGVerb 
      Caption         =   "GVerb"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "FreeVerb by Jezar at Dreampoint (2000)"
      Height          =   255
      Left            =   3360
      TabIndex        =   4
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "GVerb by Juhana Sadeharju (1999)"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private SndDry() As Single 'das trockene Signal, so wie es aus der Datei gelesen wird
Private Snd()    As Single 'das verhallte Signal nach dem zufügen von Hall durch einen der beiden Reverbs

Private DataOffset As Long
Private FileLength As Long
Private nChannels  As Integer
Private rev As ty_gverb
Private frv As revmodel
Private Declare Function PlaySoundData Lib "winmm.dll" Alias "PlaySoundA" (lpData As Any, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Private Const SND_ASYNC  As Long = &H1&
Private Const SND_MEMORY As Long = &H4&
'Private Const SND_LOOP   As Long = &H8&

Private DefaultFileName As String
Private FileName As String

'Links:
'======
'http://stackoverflow.com/questions/5318989/reverb-algorithm
'http://www.soundonsound.com/sos/Oct01/articles/advancedreverb1.asp
'http://www.earlevel.com/main/1997/01/19/a-bit-about-reverb/
'https://ccrma.stanford.edu/~jos/pasp/
'http://freeverb3vst.osdn.jp/downloads.shtml
'https://github.com/swh/lv2/tree/master/gverb
'http://wiki.audacityteam.org/wiki/GVerb

Private Sub Form_Load()
    
    DefaultFileName = App.Path & "\Resources\TestWav_mit_Hall_"
    'DefaultFileName = "C:\TestDir\TestWav_mit_Hall_"
    CbGVerb.AddItem "The Quick Fix"
    CbGVerb.AddItem "Nice hall effect"
    CbGVerb.AddItem "Singing in the Sewer"
    CbGVerb.AddItem "Last row of the church"
    CbGVerb.AddItem "Electric guitar and electric bass"
    CbGVerb.ListIndex = 0
    TxFilename = DefaultFileName
    
    File_Read SndDry, App.Path & "\" & "Reverb_Test17.wav"
    
    'um es möglichst einfach zu halten werden die Parameter vorweg hier bestimmt
    DataOffset = 88 / 4
    nChannels = 2
    Snd = SndDry

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Player_Stop
    
End Sub

'0 dB = 1.0 and -6.02 dB is 0.5.
Private Function DecibelsToAmplitude(ByVal decibels As Double) As Double
    
    DecibelsToAmplitude = 10# ^ (decibels / 20) * 4

End Function

Private Sub CbGVerb_Click()
    Dim SPS As Long: SPS = 44100
                'Samplerate,
                '    maxroomsize,
                '           roomsize,
                '               reverbtime,
                '                     damping,
                '                          spread,
                '                              inputbandwith,
                '                                   earlylevel, taillevel
'verb = gverb_new(48000, 300.0f, 50.0f, 7.0f, 0.5f, 15.0f, 0.5f, 0.5f, 0.5f)

    'ein paar Presets vom Audacity-Wiki
    Select Case CbGVerb.ListIndex
    Case 0
        'The Quick Fix
        '=============
        'Roomsize:                40   m²
        'Reverb time:              4   s
        'Damping:                  0.9
        'Input bandwidth:          0.75
        'Dry signal level:         0   dB
        'Early reflection level: -22   dB
        'Tail level:             -28   dB
        rev = gverb_new(SPS, 40#, 40#, 4, 0.9, 15, 0.75, DecibelsToAmplitude(-22), DecibelsToAmplitude(-28))
        'rev = gverb_new(SPS, 40#, 40#, 4, 0.9, 15, 0.75, -22, -28)
    Case 1
        'Bright, small hall
        '==================
        'Roomsize:                50   m²
        'Reverb time:              1.5 s
        'Damping:                  0.1
        'Input bandwidth:          0.75
        'Dry signal level:        -1.5 dB
        'Early reflection level: -10   dB
        'Tail level:             -20   dB
        rev = gverb_new(SPS, 50#, 50#, 15, 0.1, 15, 0.75, DecibelsToAmplitude(-10), DecibelsToAmplitude(-20))
        'rev = gverb_new(SPS, 50#, 50#, 15, 0.1, 15, 0.75, -10, -20)
    Case 2
        'Nice hall effect
        '================
        'Roomsize:                40   m²
        'Reverb time:             20   s
        'Damping:                  0.50
        'Input bandwidth:          0.75
        'Dry signal level:         0   dB
        'Early reflection level: -10   dB
        'Tail level:             -30   dB
        rev = gverb_new(SPS, 40#, 40#, 20, 0.5, 15, 0.75, DecibelsToAmplitude(-10), DecibelsToAmplitude(-30))
        'rev = gverb_new(SPS, 40#, 40#, 20, 0.5, 15, 0.75, -10, -30)
    Case 3
        'Singing in the Sewer
        '====================
        'Roomsize:                 6 m²
        'Reverb time:             15 s
        'Damping:                  0.9
        'Input bandwidth:          0.1
        'Dry signal level:       -10 dB
        'Early reflection level: -10 dB
        'Tail level:             -10 dB
        rev = gverb_new(SPS, 6#, 6#, 15, 0.9, 15, 0.1, DecibelsToAmplitude(-10), DecibelsToAmplitude(-10))
        'rev = gverb_new(SPS, 6#, 6#, 15, 0.9, 15, 0.1, -10, -10)
    Case 4
        'Last row of the church
        '======================
        'Roomsize:               200   m²
        'Reverb time:              9   s
        'Damping:                  0.7
        'Input bandwidth:          0.8
        'Dry signal level:       -20   dB
        'Early reflection level: -15   dB
        'Tail level:              -8   dB
        rev = gverb_new(SPS, 200#, 200#, 9, 0.7, 0.8, 0.8, DecibelsToAmplitude(-15), DecibelsToAmplitude(-8))
        'rev = gverb_new(SPS, 200#, 200#, 9, 0.7, 0.8, 0.8, -15, -8)
    Case 5
        'Electric guitar and electric bass
        '=================================
        'Roomsize:                 1   m²
        'Reverb time:           one beat ¶
        'Damping:                  1
        'Input bandwidth:          0.7
        'Dry signal level:         0   dB
        'Early reflection level: -15   dB
        'Tail level:               0   dB
        rev = gverb_new(SPS, 1, 1, 1#, 1#, 1#, 0.7, DecibelsToAmplitude(-15), DecibelsToAmplitude(0))
        'rev = gverb_new(SPS, 1, 1, 1#, 1#, 1#, 0.7, -15, 0)
        'rev = gverb_new(SPS, 50, 50, 2.2, 1, 10, 0.7, -15, -0.1)
    End Select
End Sub

Private Sub BtnPlayGVerb_Click()
    FileName = "GVerb.wav"
    
    Dim i_l As Long, i_r As Long
    
    Dim in_l As Single
    Dim in_r As Single
    Dim ret_l As Single
    Dim ret_r As Single
    
    Dim stp As Long: stp = nChannels
    
    For i_l = DataOffset To (FileLength / 4) - stp Step stp
        
        i_r = i_l + 1
        in_l = SndDry(i_l)
        in_r = SndDry(i_r)
        
        MGVerb.gverb_do rev, (in_l + in_r) / 2, ret_l, ret_r
        
        Snd(i_l) = (in_l + ret_l) / 2
        Snd(i_r) = (in_l + ret_r) / 2
        
    Next
    
    Player_Stop
    Player_Play Snd()
    'Debug.Print MaxAmplitude(Snd)
End Sub

Private Sub BtnPlayFreeVerb_Click()
    FileName = "FreeVerb.wav"
    
    frv = MFreeVerb.New_revmodel

'scalewet        As Single = 3

'initialwet      As Single = 1 / scalewet '0.3333
'initialroom     As Single = 0.5
'initialdry      As Single = 0
'initialdamp     As Single = 0.5
'initialwidth    As Single = 1
'initialmode     As Single = 0
    
    'MFreeVerb.Freeverb_setParameters frv, aWet:=0.5, aRoomsize:=2, aDry:=1, aDamp:=0.1, aWidth:=1, aMode:=0# '50
    
    Dim i_l As Long, i_r As Long
    
    Dim in_l As Single
    Dim in_r As Single
    Dim ret_l As Single
    Dim ret_r As Single
    
    Dim stp As Long: stp = nChannels
    
    For i_l = DataOffset To (FileLength / 4) - stp Step stp
        
        i_r = i_l + 1
        
        in_l = SndDry(i_l)
        in_r = SndDry(i_r)
        
        MFreeVerb.revmodel_processmix frv, in_l, in_r, ret_l, ret_r
        
        Snd(i_l) = ret_l
        Snd(i_r) = ret_r
        
    Next
    
    Player_Stop
    Player_Play Snd()
    'Debug.Print MaxAmplitude(Snd)
    
End Sub

Private Function MaxAmplitude(amp() As Single) As Single
    
    MaxAmplitude = MaxSngArr(amp)

End Function


Private Sub BtnPlayDry_Click()
    
    Player_Play SndDry()

End Sub
Private Sub BtnStop_Click()
    
    Player_Stop

End Sub
Private Sub BtnSave_Click()
    
    File_Save Snd, DefaultFileName & FileName
    
End Sub

Private Sub TxFilename_Change()
    
    DefaultFileName = TxFilename

End Sub

'##########'    File Functions    '##########'
Private Sub File_Read(file() As Single, ByVal FileName As String)
    
Try: On Error GoTo Finally

    Dim FNr As Integer: FNr = FreeFile
    Open FileName For Binary Access Read As FNr
    FileLength = LOF(FNr)
    ReDim file(0 To (FileLength / 4) - 1)
    Get FNr, , file
        
Finally:
    Close FNr
    
    If Err Then MsgBox Err & " " & Err.Description: Err.Clear

End Sub

Private Sub File_Save(file() As Single, ByVal FileName As String)
    
Try: On Error GoTo Finally

    Dim FNr As Integer: FNr = FreeFile
    Open FileName For Binary Access Write As FNr
    Put FNr, , file
    
Finally:
    Close FNr
    
    If Err Then MsgBox Err & " " & Err.Description: Err.Clear

End Sub

'##########'    Player Functions    '##########'
Sub Player_Stop()

    PlaySoundData ByVal 0, 0, 0

End Sub

Sub Player_Play(wav() As Single)
    
    PlaySoundData wav(0), 0, SND_ASYNC Or SND_MEMORY

End Sub

