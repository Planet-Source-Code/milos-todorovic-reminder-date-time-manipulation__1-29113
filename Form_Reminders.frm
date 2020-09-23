VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form_Reminders 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reminders"
   ClientHeight    =   4320
   ClientLeft      =   645
   ClientTop       =   1230
   ClientWidth     =   6840
   Icon            =   "Form_Reminders.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   6840
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo_Period 
      Height          =   315
      ItemData        =   "Form_Reminders.frx":038A
      Left            =   1260
      List            =   "Form_Reminders.frx":03A0
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   3840
      Width           =   4215
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6210
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Reminders.frx":03D0
            Key             =   "Account"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Reminders.frx":076A
            Key             =   "Lead"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Reminders.frx":0B04
            Key             =   "Directory"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command_Snooze 
      Caption         =   "&Snooze"
      Default         =   -1  'True
      Height          =   345
      Left            =   5610
      TabIndex        =   4
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command_Dismiss 
      Caption         =   "&Dismiss"
      Enabled         =   0   'False
      Height          =   345
      Left            =   5610
      TabIndex        =   3
      Top             =   3210
      Width           =   1095
   End
   Begin VB.CommandButton Command_Open 
      Caption         =   "&Open Item"
      Enabled         =   0   'False
      Height          =   345
      Left            =   4380
      TabIndex        =   2
      Top             =   3210
      Width           =   1095
   End
   Begin VB.CommandButton Command_DismissAll 
      Caption         =   "Dismiss &All"
      Enabled         =   0   'False
      Height          =   345
      Left            =   120
      TabIndex        =   1
      Top             =   3210
      Width           =   1095
   End
   Begin VB.ComboBox Combo_Numbers 
      Height          =   315
      ItemData        =   "Form_Reminders.frx":0E9E
      Left            =   120
      List            =   "Form_Reminders.frx":0EFF
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   3840
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListView_Reminders 
      Height          =   1980
      Left            =   120
      TabIndex        =   6
      Top             =   1170
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   3493
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Type"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Follow-up Flag"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Telephone"
         Object.Width           =   2293
      EndProperty
   End
   Begin VB.Label Label_Name 
      AutoSize        =   -1  'True
      Caption         =   "[Name] ([Type])"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   420
      TabIndex        =   9
      Top             =   180
      Width           =   1395
   End
   Begin VB.Label Label_DateTime 
      AutoSize        =   -1  'True
      Caption         =   "Start time: Date at Time"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   420
      TabIndex        =   8
      Top             =   570
      Width           =   1725
   End
   Begin VB.Label Label_Flag 
      AutoSize        =   -1  'True
      Caption         =   "Subject: [Value]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   420
      TabIndex        =   7
      Top             =   840
      Width           =   1155
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   120
      Picture         =   "Form_Reminders.frx":0F76
      Top             =   150
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Click Snooze to be remainded again in:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   3630
      Width           =   2760
   End
End
Attribute VB_Name = "Form_Reminders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lvListItems As ListItem

Private Sub Form_Load()

On Error Resume Next
    'Set default values
    Combo_Numbers.Text = "5"
    Combo_Period.Text = "Minutes"
    'Add some sample values to the list
    ListView_Reminders.ListItems.Clear
    Set lvListItems = ListView_Reminders.ListItems.Add(, , "Account", , "Account")
        lvListItems.SubItems(1) = "Send E-mail"
        lvListItems.SubItems(2) = "George Bush"
        lvListItems.SubItems(3) = "(800) 888-8888"
    Set lvListItems = ListView_Reminders.ListItems.Add(, , "Leads", , "Lead")
        lvListItems.SubItems(1) = "Send Letter"
        lvListItems.SubItems(2) = "Mickey Mouse"
        lvListItems.SubItems(3) = "(800) 777-7777"
    Set lvListItems = ListView_Reminders.ListItems.Add(, , "Directory", , "Directory")
        lvListItems.SubItems(1) = "Call back"
        lvListItems.SubItems(2) = "Jackie Brown"
        lvListItems.SubItems(3) = "(800) 666-6666"
    Set lvListItems = ListView_Reminders.ListItems.Add(, , "Directory", , "Directory")
        lvListItems.SubItems(1) = "Send E-mail"
        lvListItems.SubItems(2) = "Bruce Lee"
        lvListItems.SubItems(3) = "(800) 555-5555"
    Set lvListItems = ListView_Reminders.ListItems.Add(, , "Account", , "Account")
        lvListItems.SubItems(1) = "Send E-mail"
        lvListItems.SubItems(2) = "George Washington"
        lvListItems.SubItems(3) = "(800) 444-4444"
    Set lvListItems = ListView_Reminders.ListItems.Add(, , "Account", , "Account")
        lvListItems.SubItems(1) = "Send E-mail"
        lvListItems.SubItems(2) = "George Bush"
        lvListItems.SubItems(3) = "(800) 333-3333"
    'Display labels accourding to the list's current selection
    RefreshLabels
    
End Sub

Private Sub Command_Snooze_Click()

Dim NewDate, NewTime As String
    'If the change in date does not occure the date will be set to todays date
    NewDate = Date 'Set value to today
    'If the change in time does not occure the time will be set to current time
    NewTime = Time 'Set value to now
    Select Case Combo_Period.Text 'What are we adding?
        Case "Minutes"
            NewTime = GetFutureTime(Time, Combo_Numbers.Text, "Minutes")
            'GetFutureTime function will set DateUpdateNeeded to True is day change occured
            If DateUpdateNeeded = True Then NewDate = GetFutureDate(Date, 1, "Days") 'Increase date by one day
        Case "Hours"
            NewTime = GetFutureTime(Time, Combo_Numbers.Text, "Hours")
            'GetFutureTime function will set DateUpdateNeeded to True is day change occured
            If DateUpdateNeeded = True Then NewDate = GetFutureDate(Date, 1, "Days") 'Increase date by one day
        Case "Days"
            NewDate = GetFutureDate(Date, Combo_Numbers.Text, "Days")
        Case "Weeks"
            NewDate = GetFutureDate(Date, Combo_Numbers.Text, "Weeks")
        Case "Months"
            NewDate = GetFutureDate(Date, Combo_Numbers.Text, "Months")
        Case "Years"
            NewDate = GetFutureDate(Date, Combo_Numbers.Text, "Days")
    End Select
    'Display new value
    MsgBox "Snoozed until: " & NewDate & " at " & NewTime, vbInformation + vbOKOnly, "Information"

End Sub

Private Sub ListView_Reminders_ItemClick(ByVal Item As MSComctlLib.ListItem)
    RefreshLabels
End Sub

Private Sub Combo_Period_Click()
    'This controls maximum values for different units
    Select Case Combo_Period.Text
        Case "Minutes"
            Combo_Numbers.Clear
            Combo_Numbers.AddItem "5"
            Combo_Numbers.AddItem "10"
            Combo_Numbers.AddItem "15"
            Combo_Numbers.AddItem "20"
            Combo_Numbers.AddItem "25"
            Combo_Numbers.AddItem "30"
            Combo_Numbers.AddItem "35"
            Combo_Numbers.AddItem "40"
            Combo_Numbers.AddItem "45"
            Combo_Numbers.AddItem "50"
            Combo_Numbers.AddItem "55"
            Combo_Numbers.AddItem "60"
            Combo_Numbers.Text = "5"
        Case "Hours"
            Combo_Numbers.Clear
            Combo_Numbers.AddItem "1"
            Combo_Numbers.AddItem "2"
            Combo_Numbers.AddItem "3"
            Combo_Numbers.AddItem "4"
            Combo_Numbers.AddItem "5"
            Combo_Numbers.AddItem "6"
            Combo_Numbers.AddItem "7"
            Combo_Numbers.AddItem "8"
            Combo_Numbers.AddItem "9"
            Combo_Numbers.AddItem "10"
            Combo_Numbers.AddItem "11"
            Combo_Numbers.AddItem "12"
            Combo_Numbers.AddItem "13"
            Combo_Numbers.AddItem "14"
            Combo_Numbers.AddItem "15"
            Combo_Numbers.AddItem "16"
            Combo_Numbers.AddItem "17"
            Combo_Numbers.AddItem "18"
            Combo_Numbers.AddItem "19"
            Combo_Numbers.AddItem "20"
            Combo_Numbers.AddItem "21"
            Combo_Numbers.AddItem "22"
            Combo_Numbers.AddItem "23"
            Combo_Numbers.AddItem "24"
            Combo_Numbers.Text = "1"
        Case "Days"
            Combo_Numbers.Clear
            Combo_Numbers.AddItem "1"
            Combo_Numbers.AddItem "2"
            Combo_Numbers.AddItem "3"
            Combo_Numbers.AddItem "4"
            Combo_Numbers.AddItem "5"
            Combo_Numbers.AddItem "6"
            Combo_Numbers.AddItem "7"
            Combo_Numbers.AddItem "8"
            Combo_Numbers.AddItem "9"
            Combo_Numbers.AddItem "10"
            Combo_Numbers.AddItem "11"
            Combo_Numbers.AddItem "12"
            Combo_Numbers.AddItem "13"
            Combo_Numbers.AddItem "14"
            Combo_Numbers.AddItem "15"
            Combo_Numbers.AddItem "16"
            Combo_Numbers.AddItem "17"
            Combo_Numbers.AddItem "18"
            Combo_Numbers.AddItem "19"
            Combo_Numbers.AddItem "20"
            Combo_Numbers.AddItem "21"
            Combo_Numbers.AddItem "22"
            Combo_Numbers.AddItem "23"
            Combo_Numbers.AddItem "24"
            Combo_Numbers.AddItem "25"
            Combo_Numbers.AddItem "26"
            Combo_Numbers.AddItem "27"
            Combo_Numbers.AddItem "28"
            Combo_Numbers.AddItem "29"
            Combo_Numbers.AddItem "30"
            Combo_Numbers.AddItem "31"
            Combo_Numbers.Text = "1"
        Case "Weeks"
            Combo_Numbers.Clear
            Combo_Numbers.AddItem "1"
            Combo_Numbers.AddItem "2"
            Combo_Numbers.AddItem "3"
            Combo_Numbers.AddItem "4"
            Combo_Numbers.Text = "1"
        Case "Months"
            Combo_Numbers.Clear
            Combo_Numbers.AddItem "1"
            Combo_Numbers.AddItem "2"
            Combo_Numbers.AddItem "3"
            Combo_Numbers.AddItem "4"
            Combo_Numbers.AddItem "5"
            Combo_Numbers.AddItem "6"
            Combo_Numbers.AddItem "7"
            Combo_Numbers.AddItem "8"
            Combo_Numbers.AddItem "9"
            Combo_Numbers.AddItem "10"
            Combo_Numbers.AddItem "11"
            Combo_Numbers.AddItem "12"
            Combo_Numbers.Text = "1"
        Case "Years"
            Combo_Numbers.Clear
            Combo_Numbers.AddItem "1"
            Combo_Numbers.AddItem "2"
            Combo_Numbers.AddItem "3"
            Combo_Numbers.AddItem "4"
            Combo_Numbers.AddItem "5"
            Combo_Numbers.Text = "1"
    End Select
    
End Sub

Private Sub RefreshLabels()

On Error Resume Next
    Label_Name.Caption = ListView_Reminders.SelectedItem.ListSubItems(2).Text & " (" & ListView_Reminders.SelectedItem.Text & ")"
    Label_DateTime.Caption = "Start time: " & Format(Date, "Long Date") & " at " & Format(Time, "Short Time")
    Label_Flag.Caption = "Subject: " & ListView_Reminders.SelectedItem.ListSubItems(1).Text

End Sub
