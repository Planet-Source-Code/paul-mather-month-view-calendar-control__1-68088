VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calendar Sample App"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   7080
   StartUpPosition =   2  'CenterScreen
   Begin Calendar.ctlCalendar ctlCalendar1 
      Height          =   2325
      Left            =   120
      TabIndex        =   49
      Top             =   120
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   4101
      ShowLastMonthButton=   -1  'True
      ShowNextMonthButton=   -1  'True
      ShowLastMonthDays=   -1  'True
      ShowNextMonthDays=   -1  'True
      ShowTodayLabel  =   -1  'True
      ColorBackgroundHeader=   12582912
      ColorForegroundHeader=   16777215
      ColorSelectedBack=   16711680
      ColorSelectedFore=   16777215
      ColorToday      =   255
      ColorDayColumn  =   8388608
      ColorAlarms     =   0
      ColorBackground =   -2147483643
      ColorForeground =   0
      ColorButtons    =   -2147483633
      ColorLastNextMonthDayColor=   8421504
      ColorLine       =   -2147483640
      ColorWeekNumber =   8421504
      WeekStartsWith  =   1
      ShowSelected    =   -1  'True
      ShowToolTipText =   -1  'True
      ShowWeekNumbers =   0   'False
      ShowWeekNumberLeft=   -1  'True
      AllowRightClick =   -1  'True
      UseAlarms       =   -1  'True
      ShowShortDays   =   0   'False
      BeginProperty FontHeader {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDay {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontToday {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontColumn {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdFont 
      Caption         =   "Day Font"
      Height          =   375
      Index           =   1
      Left            =   5640
      TabIndex        =   32
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmdFont 
      Caption         =   "Header Font"
      Height          =   375
      Index           =   0
      Left            =   5640
      TabIndex        =   29
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton cmdFont 
      Caption         =   "Column Font"
      Height          =   375
      Index           =   3
      Left            =   4320
      TabIndex        =   30
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmdFont 
      Caption         =   "Today Font"
      Height          =   375
      Index           =   2
      Left            =   4320
      TabIndex        =   28
      Top             =   3720
      Width           =   1335
   End
   Begin VB.ComboBox cboWeekDay 
      Height          =   315
      Left            =   6000
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   3315
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Event Options"
      Height          =   1575
      Left            =   2160
      TabIndex        =   33
      Top             =   2520
      Width           =   1935
      Begin VB.CheckBox chkOptions 
         Caption         =   "Allow Right Click"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Use Alarms"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Use Tool Tips"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Create Alarms"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1455
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Calendar Enabled"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Show Options"
      Height          =   2535
      Left            =   120
      TabIndex        =   31
      Top             =   2520
      Width           =   1960
      Begin VB.CheckBox chkOptions 
         Caption         =   "Week Numbers Left"
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Short Day Name"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   8
         Top             =   2160
         Width           =   1700
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Week Numbers"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Last Month Button"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Next Month Button"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Last Month Days"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Next Month Days"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Today Label"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Selected"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   7
         Top             =   1920
         Value           =   1  'Checked
         Width           =   1455
      End
   End
   Begin VB.Label lblColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   6000
      TabIndex        =   14
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblColorName 
      Alignment       =   1  'Right Justify
      Caption         =   "Background Color"
      Height          =   255
      Index           =   0
      Left            =   4320
      TabIndex        =   48
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   6000
      TabIndex        =   15
      Top             =   360
      Width           =   975
   End
   Begin VB.Label lblColorName 
      Alignment       =   1  'Right Justify
      Caption         =   "Foreground Color"
      Height          =   255
      Index           =   1
      Left            =   4320
      TabIndex        =   47
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label lblColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   2
      Left            =   6000
      TabIndex        =   16
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblColorName 
      Alignment       =   1  'Right Justify
      Caption         =   "Header Background"
      Height          =   255
      Index           =   2
      Left            =   4320
      TabIndex        =   46
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label lblColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   3
      Left            =   6000
      TabIndex        =   17
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblColorName 
      Alignment       =   1  'Right Justify
      Caption         =   "Header Foreground"
      Height          =   255
      Index           =   3
      Left            =   4320
      TabIndex        =   45
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lblColorName 
      Alignment       =   1  'Right Justify
      Caption         =   "Line Color"
      Height          =   255
      Index           =   4
      Left            =   4320
      TabIndex        =   44
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label lblColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   4
      Left            =   6000
      TabIndex        =   18
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lblColorName 
      Alignment       =   1  'Right Justify
      Caption         =   "Selected Background"
      Height          =   255
      Index           =   5
      Left            =   4320
      TabIndex        =   43
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label lblColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   5
      Left            =   6000
      TabIndex        =   19
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblColorName 
      Alignment       =   1  'Right Justify
      Caption         =   "Selected Foreground"
      Height          =   255
      Index           =   6
      Left            =   4320
      TabIndex        =   42
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lblColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   6
      Left            =   6000
      TabIndex        =   20
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblColorName 
      Alignment       =   1  'Right Justify
      Caption         =   "Button Color"
      Height          =   255
      Index           =   7
      Left            =   4320
      TabIndex        =   41
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label lblColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   7
      Left            =   6000
      TabIndex        =   21
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lblColorName 
      Alignment       =   1  'Right Justify
      Caption         =   "Today Circle Color"
      Height          =   255
      Index           =   8
      Left            =   4320
      TabIndex        =   40
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label lblColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   8
      Left            =   6000
      TabIndex        =   22
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblColorName 
      Alignment       =   1  'Right Justify
      Caption         =   "Alarm Bold Color"
      Height          =   255
      Index           =   9
      Left            =   4320
      TabIndex        =   39
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label lblColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   9
      Left            =   6000
      TabIndex        =   23
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lblColorName 
      Alignment       =   1  'Right Justify
      Caption         =   "Day Heading Color"
      Height          =   255
      Index           =   10
      Left            =   4320
      TabIndex        =   38
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label lblColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   10
      Left            =   6000
      TabIndex        =   24
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label lblColorName 
      Alignment       =   1  'Right Justify
      Caption         =   "Next/Last Month Days"
      Height          =   255
      Index           =   11
      Left            =   4080
      TabIndex        =   37
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label lblColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   11
      Left            =   6000
      TabIndex        =   25
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label lblColorName 
      Alignment       =   1  'Right Justify
      Caption         =   "Week Starts With"
      Height          =   255
      Index           =   12
      Left            =   4080
      TabIndex        =   36
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label lblColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   12
      Left            =   6000
      TabIndex        =   26
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label lblColorName 
      Alignment       =   1  'Right Justify
      Caption         =   "Week Number Color"
      Height          =   255
      Index           =   13
      Left            =   4080
      TabIndex        =   35
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label lblEvent 
      Alignment       =   1  'Right Justify
      Caption         =   "Last Event : None"
      Height          =   255
      Left            =   3240
      TabIndex        =   34
      Top             =   4680
      Width           =   3735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboWeekDay_Change()
    ctlCalendar1.weekStartsWith = cboWeekDay.ListIndex + 1
End Sub

Private Sub cboWeekDay_Click()
    ctlCalendar1.weekStartsWith = cboWeekDay.ListIndex + 1
End Sub

Private Sub chkOptions_Click(Index As Integer)
    With ctlCalendar1
        Select Case Index
            Case 0
                .ShowWeekNumbers = IIf(chkOptions(Index).value = vbChecked, True, False)
            Case 1
                .ShowLastMonthButton = IIf(chkOptions(Index).value = vbChecked, True, False)
            Case 2
                .ShowNextMonthButton = IIf(chkOptions(Index).value = vbChecked, True, False)
            Case 3
                .AllowRightClick = IIf(chkOptions(Index).value = vbChecked, True, False)
            Case 4
                .ShowLastMonthDays = IIf(chkOptions(Index).value = vbChecked, True, False)
            Case 5
                .ShowNextMonthDays = IIf(chkOptions(Index).value = vbChecked, True, False)
            Case 6
                .ShowTodayLabel = IIf(chkOptions(Index).value = vbChecked, True, False)
            Case 7
                .ShowSelected = IIf(chkOptions(Index).value = vbChecked, True, False)
            Case 8
                .UseAlarms = IIf(chkOptions(Index).value = vbChecked, True, False)
            Case 9
                .ShowToolTipText = IIf(chkOptions(Index).value = vbChecked, True, False)
            Case 10
                Dim cAlarms As New cAlarmGroup
                
                If chkOptions(10).value = vbChecked Then
                    With cAlarms
                        .add DateAdd("d", -4, Now), ccPopup, "", "4 days ago - Special Event"
                        .add DateAdd("d", -1, Now), ccPopup, "", "Yesterday - Special Event"
                        .add DateAdd("n", 10, Now), ccPopup, "", "Today - Special Event"
                        .add DateAdd("h", 1, Now), ccPopup, "", "Today - Other Event"
                        .add DateAdd("d", 1, Now), ccPopup, "", "Tomorrow - Special Event"
                        .add DateAdd("d", 6, Now), ccPopup, "", "6 days from Now - Special Event"
                    End With
                    ctlCalendar1.ShowDate Date, cAlarms
                Else
                    ctlCalendar1.SetAlarms cAlarms
                End If
                
            Case 11
                .Enabled = IIf(chkOptions(Index).value = vbChecked, True, False)
            Case 12
                .ShowShortDays = IIf(chkOptions(Index).value = vbChecked, True, False)
            Case 13
                .ShowWeekNumberLeft = IIf(chkOptions(Index).value = vbChecked, True, False)
        End Select
    End With
End Sub

Private Sub cmdAbout_Click()
    ctlCalendar1.About
End Sub

Private Sub cmdFont_Click(Index As Integer)
    Dim oFont As SelectedFont
    Dim oFontObj As StdFont
    
    With ctlCalendar1
        Select Case Index
            Case 0
                Set oFontObj = .FontHeader
            Case 1
                Set oFontObj = .FontDay
            Case 2
                Set oFontObj = .FontToday
            Case 3
                Set oFontObj = .FontColumn
        End Select
        
        oFont = ShowFont(Me.hWnd, oFontObj, True)
        If oFont.bCanceled = False Then
            oFontObj.Bold = oFont.bBold
            oFontObj.Name = oFont.sSelectedFont
            oFontObj.Italic = oFont.bItalic
            oFontObj.Strikethrough = oFont.bStrikeOut
            oFontObj.Underline = oFont.bUnderline
            oFontObj.Size = oFont.nSize
            Set cmdFont(Index).Font = oFontObj
        End If
    End With
End Sub

Private Sub ctlCalendar1_AddNewAlarm(inputDate As Date)
    lblEvent.Caption = "Add New Alarm Event : " & Format(inputDate, "mm/dd/yyyy")
End Sub

Private Sub ctlCalendar1_AlarmSelected(UID As Double)
    lblEvent.Caption = "Alarm Selected : UID = " & UID
End Sub

Private Sub ctlCalendar1_DateClicked(inputDate As Date)
    lblEvent.Caption = "Date Clicked : " & Format(inputDate, "mm/dd/yyyy")
    SetSelected
End Sub

Private Sub ctlCalendar1_DateDblClicked(inputDate As Date)
    lblEvent.Caption = "Date Double Clicked : " & Format(inputDate, "mm/dd/yyyy")
End Sub

Private Sub ctlCalendar1_LastButtonClicked(inputDate As Date)
    lblEvent.Caption = "Last Month Button Clicked :" & Format(inputDate, "mm/dd/yyyy")
End Sub

Private Sub ctlCalendar1_MonthChanged(inputDate As Date)
    lblEvent.Caption = "Month Changed : " & Format(inputDate, "mm/dd/yyyy")
    SetSelected
End Sub

Private Sub ctlCalendar1_MonthHeadingClicked(inputDate As Date)
    lblEvent.Caption = "Month Heading Clicked : " & Format(inputDate, "mm/dd/yyyy")
    SetSelected
End Sub

Private Sub ctlCalendar1_MonthHeadingDblClicked(inputDate As Date)
    lblEvent.Caption = "Month Heading Double Clicked : " & Format(inputDate, "mm/dd/yyyy")
End Sub

Private Sub ctlCalendar1_NextButtonClicked(inputDate As Date)
    lblEvent.Caption = "Next Month Button Clicked : " & Format(inputDate, "mm/dd/yyyy")
End Sub

Private Sub ctlCalendar1_TodayClicked(inputDate As Date)
    lblEvent.Caption = "Today Clicked : " & Format(inputDate, "mm/dd/yyyy")
End Sub

Private Sub ctlCalendar1_WeekHeadingClicked(weekday As VbDayOfWeek)
    lblEvent.Caption = "Week Heading Clicked : " & WeekdayName(weekday)
End Sub

Private Sub ctlCalendar1_WeekHeadingDblClicked(weekday As VbDayOfWeek)
    lblEvent.Caption = "Week Heading Double Clicked : " & WeekdayName(weekday)
End Sub

Private Sub ctlCalendar1_WeekNumberClicked(weekNumber As Integer)
    lblEvent.Caption = "Week Number Clicked : " & weekNumber
End Sub

Private Sub ctlCalendar1_WeekNumberDblClicked(weekNumber As Integer)
    lblEvent.Caption = "Week Number Double Clicked : " & weekNumber
End Sub

Private Sub SetSelected()
    chkOptions(7).value = IIf(ctlCalendar1.ShowSelected = True, vbChecked, vbUnchecked)
End Sub
Private Sub Form_Load()
    With ctlCalendar1
        .AllowRightClick = True
        .UseAlarms = True
    
        lblColors(0).BackColor = .ColorBackground
        lblColors(1).BackColor = .ColorForeground
        lblColors(2).BackColor = .ColorBackgroundHeader
        lblColors(3).BackColor = .ColorForegroundHeader
        lblColors(4).BackColor = .ColorLine
        lblColors(5).BackColor = .ColorSelectedBack
        lblColors(6).BackColor = .ColorSelectedFore
        lblColors(7).BackColor = .ColorButtons
        lblColors(8).BackColor = .ColorToday
        lblColors(9).BackColor = .ColorAlarms
        lblColors(10).BackColor = .ColorDayColumn
        lblColors(11).BackColor = .ColorLastNextMonthDayColor
        lblColors(12).BackColor = .ColorWeekNumber
        
        Set cmdFont(0).Font = ctlCalendar1.FontHeader
        Set cmdFont(1).Font = ctlCalendar1.FontDay
        Set cmdFont(2).Font = ctlCalendar1.FontToday
        Set cmdFont(3).Font = ctlCalendar1.FontColumn
    End With
    
    With cboWeekDay
        .Clear
        .AddItem WeekdayName(vbSunday, True)
        .AddItem WeekdayName(vbMonday, True)
        .AddItem WeekdayName(vbTuesday, True)
        .AddItem WeekdayName(vbWednesday, True)
        .AddItem WeekdayName(vbThursday, True)
        .AddItem WeekdayName(vbFriday, True)
        .AddItem WeekdayName(vbSaturday, True)
        .ListIndex = 0
    End With
    
End Sub

Private Sub lblColors_Click(Index As Integer)
Dim oColor As SelectedColor
    
    oColor = ShowColor(Me.hWnd, True)
    If oColor.bCanceled = False Then
        lblColors(Index).BackColor = oColor.oSelectedColor
        
        With ctlCalendar1
            Select Case Index
                Case 0
                    .ColorBackground = oColor.oSelectedColor
                Case 1
                    .ColorForeground = oColor.oSelectedColor
                Case 2
                    .ColorBackgroundHeader = oColor.oSelectedColor
                Case 3
                    .ColorForegroundHeader = oColor.oSelectedColor
                Case 4
                    .ColorLine = oColor.oSelectedColor
                Case 5
                    .ColorSelectedBack = oColor.oSelectedColor
                Case 6
                    .ColorSelectedFore = oColor.oSelectedColor
                Case 7
                    .ColorButtons = oColor.oSelectedColor
                Case 8
                    .ColorToday = oColor.oSelectedColor
                Case 9
                    .ColorAlarms = oColor.oSelectedColor
                Case 10
                    .ColorDayColumn = oColor.oSelectedColor
                Case 11
                    .ColorLastNextMonthDayColor = oColor.oSelectedColor
                Case 12
                    .ColorWeekNumber = oColor.oSelectedColor
            End Select
        End With
    End If
End Sub
