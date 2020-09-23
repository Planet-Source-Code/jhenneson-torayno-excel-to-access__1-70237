VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmTransafer 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Transfer"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   10680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtsheet 
      Height          =   285
      Left            =   8520
      TabIndex        =   19
      Top             =   120
      Width           =   2055
   End
   Begin VB.PictureBox P1 
      Height          =   735
      Left            =   3120
      ScaleHeight     =   675
      ScaleWidth      =   5235
      TabIndex        =   15
      Top             =   3000
      Visible         =   0   'False
      Width           =   5295
      Begin VB.Label lbl1 
         Caption         =   "Loading please wait . . . . . . . . ."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   5175
      End
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load XLS"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txt1 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   6015
   End
   Begin VB.PictureBox Picture1 
      Enabled         =   0   'False
      Height          =   1695
      Left            =   120
      ScaleHeight     =   1635
      ScaleWidth      =   5115
      TabIndex        =   4
      Top             =   840
      Width           =   5175
      Begin VB.ComboBox cboTable 
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   4935
      End
      Begin VB.CommandButton cmdMDB 
         Caption         =   "MDB DESTINATION"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   2175
      End
      Begin VB.TextBox txtmdb 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   600
         Width           =   4935
      End
      Begin VB.Label Label2 
         Caption         =   "Choose destination table:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   3015
      End
   End
   Begin VB.Frame frme1 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   360
      Width           =   3855
      Begin VB.OptionButton opt1 
         Caption         =   "Manage Fields"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   13
         Top             =   120
         Width           =   1935
      End
      Begin VB.OptionButton opt1 
         Caption         =   "Access Database"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   14
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.CheckBox chk1 
      Caption         =   "Check All"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   6360
      Width           =   1815
   End
   Begin VB.PictureBox Picture2 
      Enabled         =   0   'False
      Height          =   1695
      Left            =   5400
      ScaleHeight     =   1635
      ScaleWidth      =   5115
      TabIndex        =   5
      Top             =   840
      Width           =   5175
      Begin MSComctlLib.ListView lst2 
         Height          =   1335
         Left            =   65
         TabIndex        =   17
         Top             =   240
         Width           =   4990
         _ExtentX        =   8811
         _ExtentY        =   2355
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fields"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "DataType"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "Check field to be excluded"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   0
         Width           =   3375
      End
   End
   Begin MSComctlLib.ListView lst1 
      Height          =   3615
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   6376
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   360
      Top             =   6960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdTrans 
      Caption         =   "Transfer"
      Enabled         =   0   'False
      Height          =   495
      Left            =   8160
      TabIndex        =   0
      Top             =   6360
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "SHEET NAME"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7440
      TabIndex        =   8
      Top             =   180
      Width           =   1095
   End
End
Attribute VB_Name = "frmTransafer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private strqry As String
Private db As New ADODB.Connection
Private XLS As Excel.Application
Private WBOOK As Excel.Workbook
Private WSHEET As Excel.Worksheet
Private RNG As Excel.Range
Private strDT(20) As String

Private Sub chk1_Click()
    Dim int1 As Integer
    If chk1.Value = 1 Then
        For int1 = 1 To lst1.ListItems.Count
            lst1.ListItems(int1).Checked = True
        Next int1
        lst1.Refresh
    Else
        For int1 = 1 To lst1.ListItems.Count
            lst1.ListItems(int1).Checked = False
        Next int1
        lst1.Refresh
    End If
End Sub

Private Sub cmdLoad_Click()
    'open a dialog to search for xls file
    With CD1
        .DialogTitle = "OPEN XLS"
        .FileName = ""
        'filter the file to be searched
        .Filter = "Excel Files (*.xls)" + Chr$(124) + "*.xls" + Chr$(124)
        .ShowOpen
    End With
    P1.Visible = True
    If CD1.FileName = "" Or IsNull(CD1.FileName) Then
        Exit Sub
    End If
    txt1.Text = CD1.FileName
    
    'Create a new instance of Excel
    Set XLS = CreateObject("Excel.Application")
    'Open XLS file                  or txt1.text
    Set WBOOK = XLS.Workbooks.Open(CD1.FileName)
        
    'close XLS file w/o saving
    WBOOK.Close False
    'quit excel
    XLS.Quit
    P1.Visible = False
    txtsheet.SetFocus
End Sub

Private Sub LoadXLS()
    Dim y As ListItem, ctr%
    If CD1.FileName <> "" Then
    
        Set XLS = CreateObject("Excel.Application")
        Set WBOOK = XLS.Workbooks.Open(txt1)
        'Set the WSHEET variable to the selected worksheet
        On Error GoTo err
        err.Description = "Worksheet not found!"
        Set WSHEET = WBOOK.Worksheets(txtsheet.Text)
        'Get the used range of the current worksheet
        Set RNG = WSHEET.UsedRange
        'load the no. of excel columns to counter
        Counter = RNG.Columns.Count
        
        lst1.ListItems.Clear
        lst1.ColumnHeaders.Clear
        
        Dim no%, no1%, cntr As Boolean, str2 As String, x As ListItem
        no1% = 1
        
        Do While Not IsNull(WSHEET.Cells(no1%, no% + 1).Value) And WSHEET.Cells(no1%, no% + 1).Value <> ""
            For no% = 0 To Counter - 1
                If Not IsNull(WSHEET.Cells(1, no% + 1).Value) And WSHEET.Cells(1, no% + 1).Value <> "" Then
GoBack:
                    str2 = InputBox("Please input the data type of " & WSHEET.Cells(1, no% + 1).Value & Chr(13) & Chr(13) & "Choose from the following:" & Chr(13) & Chr(13) & "str = for string" & Chr(13) & "int = for integer" & Chr(13) & "dbl = for double" & Chr(13) & "date = for Date")
                    If UCase(str2) <> "STR" And UCase(str2) <> "DBL" And UCase(str2) <> "INT" And UCase(str2) <> "DATE" Then
                        MsgBox "Invalid Input!", vbExclamation
                        GoTo GoBack
                    End If
                    Set x = lst2.ListItems.Add(, , WSHEET.Cells(1, no% + 1).Value)
                    x.SubItems(1) = UCase(str2)
                    lst1.ColumnHeaders.Add , , WSHEET.Cells(1, no% + 1).Value
                Else
                    GoTo nxtrow
                End If
            Next no%
            no% = 1
nxtrow:
            no1% = no1% + 1
            cntr = True
        Loop
        
        If cntr = False Then
            MsgBox "The sheet is empty!", vbExclamation
            cntr = True
            chk1.Enabled = False
            frme1.Enabled = False
            Exit Sub
        Else
            chk1.Enabled = True
            frme1.Enabled = True
        End If
        GoTo nxtpart
        'loads data of XLS file to the grid
        Do Until IsNull(WSHEET.Cells(no1%, no%).Value) Or WSHEET.Cells(no1%, no%).Value = ""
nxtpart:
            If Not IsNull(WSHEET.Cells(no1%, no%).Value) And WSHEET.Cells(no1%, no%).Value <> "" Then
                Set y = lst1.ListItems.Add(, , WSHEET.Cells(no1%, no%).Value)
                Dim intctr As Integer
                For intctr = no% + 1 To Counter
                    y.SubItems(intctr - 1) = WSHEET.Cells(no1%, intctr)
                Next intctr
            End If
            no1% = no1% + 1
        Loop
        'close XLS file w/o saving
        lst1.Refresh
        WBOOK.Close False
        'quit excel
        XLS.Quit
        
    End If
    Exit Sub
err:
    WBOOK.Close False
    XLS.Quit
    MsgBox err.Description, vbExclamation, "Search Sheet"
End Sub
Private Sub cmdMDB_Click()
    Dim db1 As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim iPos As Integer
    'open a dialog to search for xls file
    With CD1
        .DialogTitle = "OPEN XLS"
        .FileName = ""
        'filter the file to be searched
        .Filter = "Microsoft Database Files (*.mdb)" + Chr$(124) + "*.mdb" + Chr$(124)
        .ShowOpen
    End With
    txtmdb.Text = CD1.FileName
    db1.ConnectionString = "Provider=Microsoft.Jet.oledb.4.0;Data Source = " & txtmdb.Text
    db1.Open
    
    Set rs = db1.OpenSchema(adSchemaTables)
    Do Until rs.EOF
        iPos = InStr(1, rs!TABLE_NAME, "MSys")
        If iPos = 0 Then
            cboTable.AddItem rs("TABLE_NAME")
        End If
        rs.MoveNext
    Loop
    cboTable.Text = cboTable.List(0)
    db1.Close
    cmdTrans.Enabled = True
End Sub

Function DataType(lstno As Integer) As String
    DataType = lst2.ListItems(lstno).SubItems(1)
    Select Case (DataType)
    Case "STR": DataType = "'"
    Case "INT": DataType = ""
    Case "DATE": DataType = "#"
    Case "DBL": DataType = ""
    End Select
End Function

Private Sub cmdTrans_Click()
    Dim rs As New ADODB.Recordset
    On Error GoTo err
    If MsgBox("Warning! Make sure that the table you choose has " & Chr(13) & "the same fields from the excel file!", vbYesNo + vbQuestion, "Confirmation") = vbYes Then
        Call Connection
        lbl1.Caption = "Transfering please wait . . . . "
        P1.Visible = True
        
        Dim int1%, int2%, int3%, int4%
        Dim str1(20) As String
        Dim str2 As String
        Dim str3 As String
        Dim bool As Boolean, bool2 As Boolean
        
        int2% = lst1.ColumnHeaders.Count
        int1% = 1
        int4% = 2
        
        For int3% = 1 To lst1.ListItems.Count - 1
            If lst1.ListItems(int3%).Checked = False Then
                GoTo lastNext
            End If
           ' DataType (int3%)
            str1(int3%) = DataType(1) & Trim$(lst1.ListItems(int3%).Text) & DataType(1) & ","
            If int4% <= lst1.ColumnHeaders.Count Then
                str2 = lst1.ColumnHeaders(1).Text & ","
            End If
            int4% = 2
            Do While Not int1% = int2%
                str3 = str1(int3%)
                str1(int3%) = str1(int3%) & DataType(int1% + 1) & Trim$(lst1.ListItems(int3%).SubItems(int1%)) & DataType(int1% + 1) & ","
                
                If bool = False Then
                    If lst2.ListItems(int4%).Checked = False Then
                        str2 = str2 & lst1.ColumnHeaders(int4%).Text & ","
                    End If
                End If
                
                If int4% <= lst1.ColumnHeaders.Count Then
                    If lst2.ListItems(int4%).Checked = True Then
                        str1(int3%) = str3
                    End If
                    If int4% = lst1.ColumnHeaders.Count Then
                        bool = True
                    End If
                End If
                
                int4% = int4% + 1
                
                int1% = int1% + 1
            Loop
            str1(int3%) = Trim$(str1(int3%))
            str1(int3%) = Left$(str1(int3%), Len(str1(int3%)) - 1)
            
            If bool2 = False Then
                If int4% = lst1.ColumnHeaders.Count + 1 Then
                    str2 = Trim$(str2)
                    str2 = Left$(str2, Len(str2) - 1)
                    int4% = 5
                End If
                bool2 = True
            End If
lastNext:
            int1% = 1
            'int4% = 2
        Next int3%
        
        For int1% = 1 To lst1.ListItems.Count - 1
            If lst1.ListItems(int1%).Checked = False Then GoTo NxtNext
            strqry = "Select * from " & cboTable.Text & " where " & lst1.ColumnHeaders(1).Text & " = " & "'" & Trim$(lst1.ListItems(int1%).Text) & "'"
            Set rs = db.Execute(strqry)
            If rs.EOF Then
                strqry = "insert into " & cboTable.Text & "(" & str2 & ")" & " Values (" & str1(int1%)
                strqry = strqry & ")"
                db.Execute strqry
            End If
NxtNext:
        Next int1%
    End If
    lbl1.Caption = "Loading please wait . . . . . . . "
    P1.Visible = False
    bool = False
    bool2 = False
    int4% = 2
    MsgBox "Transfering Complete!", vbInformation
    Exit Sub
err:
End Sub

Private Sub Connection()
    If db.State = 1 Then db.Close
    db.CursorLocation = adUseClient
    db.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & txtmdb.Text
    db.Open
End Sub

Private Sub opt1_Click(Index As Integer)
    If Index = 0 Then
        Picture1.Enabled = True
        Picture2.Enabled = False
    Else
        Picture2.Enabled = True
        Picture1.Enabled = False
    End If
End Sub

Private Sub txtsheet_KeyPress(KeyAscii As Integer)
    If txtsheet.Text <> "" And Not IsNull(txtsheet.Text) Then
        If KeyAscii = 13 Then
            P1.Visible = True
            LoadXLS
            P1.Visible = False
            opt1(0).Value = True
        End If
    End If
End Sub
