VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Listbox with highlight bar, scrollbar, and different fonts"
   ClientHeight    =   5385
   ClientLeft      =   1965
   ClientTop       =   1560
   ClientWidth     =   5970
   Height          =   5790
   Left            =   1905
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   5970
   Top             =   1215
   Width           =   6090
   Begin VB.CommandButton Command5 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Add"
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete highlighted "
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Populate"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   4080
      Width           =   1815
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3855
      Left            =   4920
      Min             =   1
      TabIndex        =   3
      Top             =   120
      Value           =   1
      Width           =   255
   End
   Begin VB.Label LblAge 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   15
      Left            =   3600
      TabIndex        =   38
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label LblAge 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   3600
      TabIndex        =   37
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label LblAge 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   3600
      TabIndex        =   36
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label LblAge 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   3600
      TabIndex        =   35
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label LblAge 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   3600
      TabIndex        =   34
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label LblAge 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   3600
      TabIndex        =   33
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label LblAge 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   3600
      TabIndex        =   32
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label LblAge 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   3600
      TabIndex        =   31
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label LblAge 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   3600
      TabIndex        =   30
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label LblAge 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   3600
      TabIndex        =   29
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label LblAge 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   3600
      TabIndex        =   28
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label LblAge 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   3600
      TabIndex        =   27
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label LblAge 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   3600
      TabIndex        =   26
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label LblAge 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   3600
      TabIndex        =   25
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label LblAge 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   24
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label LblName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   15
      Left            =   720
      TabIndex        =   23
      Top             =   3720
      Width           =   2895
   End
   Begin VB.Label LblName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   720
      TabIndex        =   22
      Top             =   3480
      Width           =   2895
   End
   Begin VB.Label LblName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   720
      TabIndex        =   21
      Top             =   3240
      Width           =   2895
   End
   Begin VB.Label LblName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   720
      TabIndex        =   20
      Top             =   3000
      Width           =   2895
   End
   Begin VB.Label LblName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   720
      TabIndex        =   19
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Label LblName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   720
      TabIndex        =   18
      Top             =   2520
      Width           =   2895
   End
   Begin VB.Label LblName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   720
      TabIndex        =   17
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Label LblName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   720
      TabIndex        =   16
      Top             =   2040
      Width           =   2895
   End
   Begin VB.Label LblName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   720
      TabIndex        =   15
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label LblName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   720
      TabIndex        =   14
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Label LblName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   720
      TabIndex        =   13
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label LblName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   720
      TabIndex        =   12
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label LblName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   720
      TabIndex        =   11
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label LblName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   10
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label LblName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   9
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808080&
      Caption         =   "Age:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808080&
      Caption         =   "Name:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3900
      Left            =   690
      TabIndex        =   0
      Top             =   90
      Width           =   4515
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit
'Visual Basic Listbox with highlight and scrollbar. It allows multiple fonts, and different backcolors and forcolors concurrently. This demo uses a mini database to populate the listbox.
'The listbox is an array of labels (can be text boxes)
'this is made to appear as a SINGLE custom control, there are no ocx's used.
'it also demonstrates a mini database engine build in Visual Basic and its recordsets are used to populate the listbox.

'hope you find this useful.



'2000 Jason Bennison  icq#48408885  Jasonbennison@hotmail.com

Dim J As Integer 'loop control
Dim Indexes(1 To 2) As Integer 'used by the highlighter bar

Private Sub Command1_Click()
'fill the array and populate the listbox

PopulateArray

'display the data on the screen
ShowDetails




End Sub


Private Sub Command2_Click()
Close #Filenum 'close the database file

Kill "C:\My Documents\MyDB.dat" 'destroy the file on the hard drive

'now open a new one with the same name

Filenum = FreeFile
RecordLen = Len(MyDB)
Open "c:\my documents\MyDB.dat" For Random As Filenum Len = RecordLen
Currentrecord = 1 'make sure this isnt zero
LastRecord = 1 'there is no data in the file yet - you just cleared it!

'clear our listbox
ShowDetails
End Sub

Private Sub Command3_Click()
Dim Response

'get the clicked recordset from the hard drive...
Get #Filenum, ((VScroll1.Value - 1) + Indexes(1)), MyDB

'You want the record the user clicked...
'get the current vaule of the scrollbar, and subtract its MIN value;
'and add the index number of the label control that was clicked;
'get that record from the hard drive,  and put it in a yesno Msgbox

Response = MsgBox("Delete " & Trim(MyDB.Names) & " from the list? - Confirm", vbYesNo + vbQuestion, "Listbox Demo")
If Response = vbNo Then Exit Sub

'now delete that record from the recordset..  and move the rest up one place.. I think you can handle that.
MsgBox "Put your delete code here..."
'You could put the delete code in the labels DoubleClick Event


End Sub

Private Sub Command4_Click()

'first; create a new blank record
LastRecord = LastRecord + 1

'get a name from the user
Dim Response As String * 42
Response = InputBox("Enter a name:")


'now put it onto our database file MyDB.dat
MyDB.Names = Response
Put #Filenum, LastRecord, MyDB

'Get an age from the user
Dim Response2
Response2 = InputBox("Enter an age")

'make sure the user hasnt entered anything stupid!
If Not IsNumeric(Response2) Then MsgBox "not a number!", vbExclamation: End

'Make sure the value is an integer
If Response2 >= 32768 Or Response2 <= -32768 Then MsgBox "Not an Integer!", vbExclamation: End
'(you will need error handlers in a real project, terminating the program is not acceptable)


'now put the age into the MyBD.dat file on the hard drive
MyDB.Ages = Response2
Put #Filenum, LastRecord, MyDB

'all done! - now display it with the last record at the bottom of the listbox
VScroll1.Max = LastRecord - 14  'number of lines in out listbox less the MIN value of the scrollbar.
VScroll1.Value = VScroll1.Max
ShowDetails
End Sub

Private Sub Command5_Click()
End
End Sub

Private Sub Form_Load()





'Now open a Database file on the hard drive.. and call it "MyDB.dat"
Filenum = FreeFile  ' dont use #1 - ever! you cant multi-session it
RecordLen = Len(MyDB)
Open "c:\my documents\MyDB.dat" For Random As Filenum Len = RecordLen
Currentrecord = 1 'make sure this isnt zero
LastRecord = 30 'because there are 30 records/entries being filled - see below

'....and if the file has not been used before; load it with some stuff...
Get #Filenum, 1, MyDB
If Asc(MyDB.Names) = 0 Or Asc(MyDB.Names) = 32 Then
    PopulateArray
Else
    'find out how many records exist in the file
    For J = 1 To 32767 'no more than an integer' or use do loop method instead
        Get #Filenum, J, MyDB
            If Asc(MyDB.Names) = 0 Or Asc(MyDB.Names) = 32 Then
            LastRecord = J - 1
            Exit For
        End If
    Next
End If


'there is no 0 in the label control array
Indexes(1) = 1
Indexes(2) = 1

'now its all on the hard drive in a file called c:\my documents\MyDB.dat, lets display them
ShowDetails
End Sub



Public Function ShowDetails()
'make this a function:  we dont know where in the project it will be called from yet.
Dim CNT As Integer
Dim NumberofLines As Integer

NumberofLines = 15 'lines of text in our listbox

'we need to find out how many records we have on the hard drive.
'if there are NumberofLines more more,  then we need to use a scrollbar, so the user can scroll thru the listed records
'the value of lastrecord can give us this.

If LastRecord <= NumberofLines Then  'NumberofLines is the number of lines in your home-made listbox
    VScroll1.Enabled = False 'if less, we dont need the scrollbar
    VScroll1.Value = 1 'show the data from the first record
Else
    VScroll1.Enabled = True
    VScroll1.Min = 1
    VScroll1.Max = (LastRecord - (NumberofLines - 1))
End If

'get from the hard drive, the record numbers starting from the current value of the scroillbar and the following fifteen (fifteen lines in the listbox, you can add more if you like)
'and put them into the array of labels

For J = VScroll1.Value To (VScroll1.Value + (NumberofLines - 1))
    CNT = CNT + 1 'this will be the index of lblname and lblage
    Get #Filenum, J, MyDB
    LblName(CNT).Caption = MyDB.Names
    LblAge(CNT).Caption = MyDB.Ages
Next





End Function

Private Sub Label1_Click()
'This label is for cosmetic reasons:  It gives the impression that the entire listbox is a SINGLE control, when in fact, it is an array of labels with a scrollbar. To allow editing of cations, use Textboxes, but this will slow dow the scrolling effect of the listbox.
End Sub

Private Sub LblAge_Click(Index As Integer)
If LblAge(Index).Caption = Trim("") Then Exit Sub
MsgBox "You selected : " & Trim(LblName(Index).Caption) & "   age: " & Trim(LblAge(Index).Caption)
End Sub

Private Sub LblAge_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If LblAge(Index).Caption = "" Then Exit Sub
If Indexes(1) = Index Then Exit Sub
Indexes(2) = Indexes(1)
Indexes(1) = Index
Highlight
End Sub


Private Sub LblName_Click(Index As Integer)
If LblName(Index).Caption = Trim("") Then Exit Sub
MsgBox "You selected : " & Trim(LblName(Index).Caption) & "   age: " & Trim(LblAge(Index).Caption)
End Sub

Private Sub LblName_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If LblName(Index).Caption = "" Then Exit Sub
If Indexes(1) = Index Then Exit Sub
Indexes(2) = Indexes(1)
Indexes(1) = Index
Highlight
End Sub


Private Sub VScroll1_Change()
ShowDetails
End Sub


Private Sub VScroll1_Scroll()
ShowDetails
End Sub



Public Sub PopulateArray()
'If calling this from another program make this into a function
MyDB.Names = "Jason Bennison"
MyDB.Ages = 38
Put #Filenum, 1, MyDB
MyDB.Names = "June Ackland"
MyDB.Ages = 19
Put #Filenum, 2, MyDB
MyDB.Names = "Jack Meadows"
MyDB.Ages = 33
Put #Filenum, 3, MyDB
MyDB.Names = "Jim Carver"
MyDB.Ages = 45
Put #Filenum, 4, MyDB
MyDB.Names = "Chris Deakin"
MyDB.Ages = 17
Put #Filenum, 5, MyDB
MyDB.Names = "John Boulton"
MyDB.Ages = 9
Put #Filenum, 6, MyDB
MyDB.Names = "Don Beech"
MyDB.Ages = 14
Put #Filenum, 7, MyDB
MyDB.Names = "Matt Boyden"
MyDB.Ages = 23
Put #Filenum, 8, MyDB
MyDB.Names = "Reg Hollis"
MyDB.Ages = 39
Put #Filenum, 9, MyDB
MyDB.Names = "Bob Cryer"
MyDB.Ages = 65
Put #Filenum, 10, MyDB
MyDB.Names = "Dave Quinnan"
MyDB.Ages = 90
Put #Filenum, 11, MyDB
MyDB.Names = "Vicky Hagan"
MyDB.Ages = 26
Put #Filenum, 12, MyDB
MyDB.Names = "Polly Page"
MyDB.Ages = 14
Put #Filenum, 13, MyDB
MyDB.Names = "Charles Brownlow"
MyDB.Ages = 16
Put #Filenum, 14, MyDB
MyDB.Names = "Derek Conway"
MyDB.Ages = 77
Put #Filenum, 15, MyDB
MyDB.Names = "Tony Stamp"
MyDB.Ages = 73
Put #Filenum, 16, MyDB
MyDB.Names = "Duncan Lennox"
MyDB.Ages = 67
Put #Filenum, 17, MyDB
MyDB.Names = "Claire Stanton"
MyDB.Ages = 53
Put #Filenum, 18, MyDB
MyDB.Names = "Geoff Daly"
MyDB.Ages = 58
Put #Filenum, 19, MyDB
MyDB.Names = "Kerry Holmes"
MyDB.Ages = 16
Put #Filenum, 20, MyDB
MyDB.Names = "Eddie Santini"
MyDB.Ages = 77
Put #Filenum, 21, MyDB
MyDB.Names = "Norika Datta"
MyDB.Ages = 73
Put #Filenum, 22, MyDB
MyDB.Names = "Danny Glase"
MyDB.Ages = 67
Put #Filenum, 23, MyDB
MyDB.Names = "Cathy Marshall"
MyDB.Ages = 53
Put #Filenum, 24, MyDB
MyDB.Names = "John Maitland"
MyDB.Ages = 58
Put #Filenum, 25, MyDB
MyDB.Names = "Alec Peters"
MyDB.Ages = 16
Put #Filenum, 26, MyDB
MyDB.Names = "Donna Harris"
MyDB.Ages = 67
Put #Filenum, 27, MyDB
MyDB.Names = "Ron Smollett"
MyDB.Ages = 53
Put #Filenum, 28, MyDB
MyDB.Names = "George Garfield"
MyDB.Ages = 58
Put #Filenum, 29, MyDB
MyDB.Names = "Mike Dashwood"
MyDB.Ages = 16
Put #Filenum, 30, MyDB
'Names Courtesy of the TV series: 'The Bill' Bosun House (aka SunHill Nick), Windsor Avenue, Merton, London SW15 3RT
LastRecord = 30 ' this is how many records we have just created
End Sub

Public Sub Highlight()
'Indexes(1) is the value of the label Index that PREVIOUSLY had focus (highlighted)
'Indexes(2) is the value of the label Index then NOW has focus (highlighted)


LblName(Indexes(2)).BackColor = RGB(255, 255, 255) 'white
LblName(Indexes(2)).ForeColor = RGB(0, 0, 0) 'black

LblName(Indexes(1)).BackColor = RGB(0, 130, 0) 'green - set the color or choice here
LblName(Indexes(1)).ForeColor = RGB(255, 255, 255) '  white forcolor

LblAge(Indexes(2)).BackColor = RGB(255, 255, 255)
LblAge(Indexes(2)).ForeColor = RGB(0, 0, 0)

LblAge(Indexes(1)).BackColor = RGB(0, 130, 0) 'green - set the color or choice here
LblAge(Indexes(1)).ForeColor = RGB(255, 255, 255) ' with white forcolor



End Sub
