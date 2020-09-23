VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "WordList by Robert Rayment"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10305
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FindWords.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   687
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   720
      TabIndex        =   0
      Top             =   60
      Width           =   10095
      Begin VB.CommandButton Command1 
         Caption         =   "&GO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   9420
         TabIndex        =   9
         Top             =   315
         Width           =   555
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   3
         Left            =   5100
         TabIndex        =   4
         Text            =   "length"
         Top             =   540
         Width           =   1275
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   3480
         TabIndex        =   3
         Text            =   "ending"
         Top             =   540
         Width           =   1275
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   1860
         TabIndex        =   2
         Text            =   "containing"
         Top             =   540
         Width           =   1275
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Text            =   "start"
         Top             =   540
         Width           =   1275
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"FindWords.frx":0E42
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   795
         Left            =   6540
         TabIndex        =   10
         Top             =   180
         Width           =   2715
      End
      Begin VB.Label Label1 
         Caption         =   "of length"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   5160
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "ending with"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3540
         TabIndex        =   7
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "containing"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1980
         TabIndex        =   6
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "starting with"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   300
         TabIndex        =   5
         Top             =   240
         Width           =   1275
      End
   End
   Begin VB.Menu brk2 
      Caption         =   "&Dictionary"
      Begin VB.Menu ProperNames 
         Caption         =   "&Proper Names"
      End
      Begin VB.Menu UKACD 
         Caption         =   "&UKACD"
      End
      Begin VB.Menu brk1 
         Caption         =   "-"
      End
      Begin VB.Menu ExitProg 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FindWords by Robert Rayment

'Dictionaries
'UKACD lower case extracted & proper names extracted
'Both modified to remove phrases, accented letters,
'apostrophes etc.
'Source:-
'The UKACD dictionary is copyrighted but freeware
'see UKACD.txt in the Apps folder

Option Base 1
Dim WordData() As Byte
Dim ALB&()  'Alphabet start binary positions
Dim FileIn$, FSize&
Dim NumOfWords&, maxlen, WordCount, wpntr&
Dim serstart$, ser$, serend$, sertyp, wordlen$, lenlim, lenop


Private Sub Form_Load()
Show
Form1.Caption = "FindWords by Robert Rayment  " + Str$(Now)
Refresh
MousePointer = 11

'Start dictionary "ACDLC0A.txt"

UKACD_Click

'FileIn$ = "ACDLC0A.txt"

'OpenDictionary 'FileIn$

MousePointer = 0

End Sub


Private Sub Command1_Click()
'GO
'Global NumOfWords&, maxlen, WordCount, wpntr&
'Global serstart$, ser$, serend$, lenlim, sertyp
' text  0          1     2        3
'sertyp 2          4     8
'2, 4, 8
'6(2+4), 10(2+8), 12(4+8), 14(2+4+8)

Form1.Cls
Refresh

SubWordCount = 0

Form1.CurrentY = 80
Form1.MousePointer = 11
sertyp = 0
Firstletter$ = ""

If serstart$ <> "" Then
   serstart$ = LCase(serstart$)
   If FileIn$ = "ProperNames.txt" Then
      a$ = UCase$(Left$(serstart$, 1))
      Mid$(serstart$, 1, 1) = a$
   End If
   sertyp = 2
   Firstletter$ = Left$(serstart$, 1)
End If
If ser$ <> "" Then sertyp = sertyp + 4: ser$ = LCase(ser$)
If serend$ <> "" Then sertyp = sertyp + 8: serend$ = LCase(serend$):
lenwordbits = Len(serstart$) + Len(ser$) + Len(serend$)

Find_lenlim_lenop

If lenwordbits = 0 And lenlim = 0 Then Form1.MousePointer = 0: Beep: Exit Sub
If lenwordbits = 0 And lenlim > 0 Then sertyp = 16

'Set search to start letter
If wpntr& = 1 Then
   If Left(serstart$, 1) <> "" Then
      a$ = Left(serstart$, 1)
      n = Asc(a$) - 96
      If n >= 1 And n <= 26 Then wpntr& = ALB&(n)
   End If
End If

Do
   Word$ = ""
   For i& = wpntr& To wpntr& + 50
      If i& >= FSize& Then Exit Do
      If WordData(i&) = &HA Then p2& = i&: Exit For
   Next i&
   For j& = wpntr& To p2& - 1
      Word$ = Word$ + Chr$(WordData(j&))
   Next j&
   
   If Firstletter$ <> "" And Left$(Word$, 1) > Firstletter$ Then Exit Do
   lenword = Len(Word$)
   
   wpntr& = p2& + 1  'For next word
   
   Form1.CurrentX = 50
   If SubWordCount >= 22 Then Form1.CurrentX = 50 + 170
   If SubWordCount >= 44 Then Form1.CurrentX = 50 + 170 + 170
   If SubWordCount >= 66 Then Form1.CurrentX = 50 + 170 + 170 + 170
   
   If lenword < lenwordbits And sertyp <> 16 Then GoTo nextloop
   
   matchtest = False
   Select Case sertyp
   Case 2 'Words starting with serstart$
      m1 = InStr(1, Word$, serstart$)
      If m1 = 1 Then matchtest = True
   Case 4   'Words containing ser$
      m1 = InStr(1, Word$, ser$)
      If m1 <> 0 Then matchtest = True
   Case 6   'Words starting with serstart$ and containing ser$
      m1 = InStr(1, Word$, serstart$)
      If m1 = 1 Then
         m2 = InStr(m1 + 1, Word$, ser$)
         If m2 <> 0 Then matchtest = True
      End If
   Case 8   'Words ending in serend$
      m1 = InStr(1, Word$, serend$)
      If m1 <> 0 And m1 = lenword - Len(serend$) + 1 Then matchtest = True
   Case 10   'Words starting with serstart$ and ending in serend$
      m1 = InStr(1, Word$, serstart$)
      If m1 = 1 Then
         m2 = InStr(m1 + 1, Word$, serend$)
         If m2 <> 0 And m2 = lenword - Len(serend$) + 1 Then matchtest = True
      End If
   Case 12   'Words containing ser$ and ending in serend$
      m1 = InStr(1, Word$, ser$)
      If m1 <> 0 Then
         m2 = InStr(m1 + 1, Word$, serend$)
         If m2 <> 0 And m2 = lenword - Len(serend$) + 1 Then matchtest = True
      End If
   Case 14   'Words starting with serstart$, containing ser$ and ending in serend$
      m1 = InStr(1, Word$, serstart$)
      If m1 = 1 Then
         m2 = InStr(m1 + 1, Word$, ser$)
         If m2 <> 0 Then
            m3 = InStr(m2 + 1, Word$, serend$)
            If m3 <> 0 And m3 = lenword - Len(serend$) + 1 Then matchtest = True
         End If
      End If
   Case 16  'all words of length=lenlim
      matchtest = True
   End Select
   
   If matchtest Then
      match = False
      Select Case lenop
      Case 0: If lenword = lenlim Or lenlim = 0 Then match = True
      Case 1: If lenword < lenlim Then match = True
      Case 2: If lenword <= lenlim Then match = True
      Case 3: If lenword > lenlim Then match = True
      Case 4: If lenword >= lenlim Then match = True
      End Select
      If match Then
         WordCount = WordCount + 1: SubWordCount = SubWordCount + 1: Form1.Print Word$
         If SubWordCount > 87 Then Exit Do
         If SubWordCount = 22 Then Form1.CurrentY = 80
         If SubWordCount = 44 Then Form1.CurrentY = 80
         If SubWordCount = 66 Then Form1.CurrentY = 80
      End If
   End If

nextloop:
   If wpntr& >= FSize& Then Exit Do
Loop

Form1.MousePointer = 0
Beep
Form1.CurrentX = 50
Form1.CurrentY = 525
If SubWordCount = 88 Then
   Form1.Print WordCount; " words found.  GO for more"
Else
   Form1.Print WordCount; " words found"
End If
Form1.Refresh

If SubWordCount = 0 Then
   WordCount = 0
   wpntr& = 1
End If

End Sub

Private Sub Find_lenlim_lenop()
'Global wordlen$,lenlim,lenop
lenlim = 0
lenop = 0   '0=   1<   2<=  3>  4>=
If wordlen$ <> "" Then
    lef1$ = Left$(wordlen$, 1)
    If lef1$ >= "1" And lef1$ <= "9" Then 'Number only
        lenlim = Val(wordlen$)
        Exit Sub
    End If
    lef2$ = Mid$(wordlen$, 2, 1)
    Select Case lef1$
    Case "="
        lenlim = Val(Mid$(wordlen$, 2))
    Case "<"
        If lef2$ <> "=" Then
            lenlim = Val(Mid$(wordlen$, 2))
            lenop = 1
        Else
            lenlim = Val(Mid$(wordlen$, 3))
            lenop = 2
        End If
    Case ">"
        If lef2$ <> "=" Then
            lenlim = Val(Mid$(wordlen$, 2))
            lenop = 3
        Else
            lenlim = Val(Mid$(wordlen$, 3))
            lenop = 4
        End If
    End Select

End If
End Sub




Private Sub Text1_Change(Index As Integer)
'Comes here as each character pressed
Select Case Index
Case 0: serstart$ = Text1(Index).Text
Case 1: ser$ = Text1(Index).Text
Case 2: serend$ = Text1(Index).Text
Case 3: wordlen$ = Text1(Index).Text
End Select

WordCount = 0
wpntr& = 1

End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 27 Then Exit Sub
If KeyAscii = 13 Then
   KeyAscii = 0
   Select Case Index
   Case 0: serstart$ = Trim$(Text1(0).Text)
   Case 1: ser$ = Trim$(Text1(1).Text)
   Case 2: serend$ = Trim$(Text1(2).Text)
   Case 3: a$ = Trim$(Text1(Index).Text)
      lenlim = 0
      lenop = 0 '0=   1<   2<=
      If a$ <> "" Then
        op$ = Left(a$, 1)
        If op$ = "<" Then
           lenop = 1
           a$ = Mid$(a$, 2)
           If op$ = "=" Then
              lenop = 2
              a$ = Mid$(a$, 2)
            End If
        ElseIf op$ = "=" Then
           lenop = 0
           a$ = Mid$(a$, 2)
        End If
        lenlim = Val(a$)
      End If
   End Select
   
   WordCount = 0
   wpntr& = 1

End If

End Sub

Private Sub UKACD_Click()
Frame1.Caption = "Find words from 173099 words in UKACD's Dictionary"

'UKACD Lower Case, words terminated by 0Ah
FileIn$ = "ACDLC0A.txt"

OpenDictionary 'FileIn$

End Sub


Private Sub ProperNames_Click()
Frame1.Caption = "Find words from 24256 Proper Names"

FileIn$ = "ProperNames.txt"

OpenDictionary 'FileIn$

End Sub

Private Sub OpenDictionary()
'Global WordData() As Byte
'Global ALB&()  'Alphabet start binary positions
'Global FileIn$, FSize&
'Global NumOfWords&, maxlen, WordCount, wpntr&
Form1.Cls
Refresh

NumOfWords& = 0&
maxlen = 0
Open FileIn$ For Binary As #1
FSize& = LOF(1)
ReDim WordData(FSize&)
Get #1, , WordData
Close
'NB Words are separated by &HA
'Set ALB&() pointer to the start of each alphabetic group
ReDim ALB&(1 To 26)
ALB&(1) = 1 'Pointer to 'a'
ab = 97    'a
ai = 1
ab = 98    'b
ai = 2
For n& = 1 To FSize& - 1
    If WordData(n&) = &HA Then 'separator after end of word
       If WordData(n& + 1) = ab Then
          ALB&(ai) = n& + 1
          ab = ab + 1
          ai = ai + 1
       End If
    End If
Next n&

'Global serstart$, ser$, serend$, lenlim, sertyp
For i = 0 To 3
Text1(i) = ""
Next i
'Global NumOfWords&, maxlen, WordCount, wpntr&
WordCount = 0
wpntr& = 1
serstart$ = ""
ser$ = ""
serend$ = ""
lenlim = 0
lenop = 0 '0=   1<   2<=

Text1(0) = serstart$
Text1(1) = ser$
Text1(2) = serend$
Text1(3) = ""
sertyp = 0
DoEvents

End Sub

Private Sub ExitProg_Click()
Unload Me
End
End Sub

