VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form DTSPackages 
   Caption         =   "DTS Packages Test"
   ClientHeight    =   4905
   ClientLeft      =   5175
   ClientTop       =   3645
   ClientWidth     =   7200
   Icon            =   "DTSPackages.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   7200
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox ssdbDTSPkgs 
      Height          =   3960
      ItemData        =   "DTSPackages.frx":030A
      Left            =   90
      List            =   "DTSPackages.frx":030C
      MultiSelect     =   1  'Simple
      Sorted          =   -1  'True
      TabIndex        =   15
      Top             =   465
      Width           =   4365
   End
   Begin VB.ComboBox cboServers 
      Enabled         =   0   'False
      Height          =   315
      Left            =   4620
      TabIndex        =   14
      Top             =   3000
      Width           =   2430
   End
   Begin MSComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   13
      Top             =   4560
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12277
         EndProperty
      EndProperty
   End
   Begin VB.Frame frmDSN 
      Caption         =   "DSN or DSN Less?"
      Height          =   690
      Left            =   4620
      TabIndex        =   8
      Top             =   360
      Width           =   2475
      Begin VB.OptionButton DSNLess 
         Height          =   270
         Left            =   1200
         TabIndex        =   10
         Top             =   285
         Width           =   285
      End
      Begin VB.OptionButton DSN 
         Height          =   270
         Left            =   150
         TabIndex        =   9
         Top             =   285
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "DSN-Less"
         Height          =   195
         Left            =   1485
         TabIndex        =   12
         Top             =   315
         Width           =   720
      End
      Begin VB.Label lblDSN 
         AutoSize        =   -1  'True
         Caption         =   "DSN"
         Height          =   195
         Left            =   435
         TabIndex        =   11
         Top             =   315
         Width           =   345
      End
   End
   Begin VB.TextBox txtPwd 
      Enabled         =   0   'False
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   4620
      PasswordChar    =   "*"
      TabIndex        =   5
      ToolTipText     =   "Password for Login to Server"
      Top             =   4095
      Width           =   1950
   End
   Begin VB.TextBox txtLogin 
      Enabled         =   0   'False
      Height          =   330
      Left            =   4620
      TabIndex        =   4
      ToolTipText     =   "Login ID for Server"
      Top             =   3525
      Width           =   1950
   End
   Begin VB.CommandButton cmdRunDTSPkg 
      Caption         =   "Run DTS Pkg"
      Enabled         =   0   'False
      Height          =   420
      Left            =   5895
      TabIndex        =   1
      Top             =   1200
      Width           =   1290
   End
   Begin VB.CommandButton cmdGetDTSList 
      Caption         =   "List DTS Pkgs"
      Enabled         =   0   'False
      Height          =   420
      Left            =   4620
      TabIndex        =   0
      Top             =   1200
      Width           =   1290
   End
   Begin DTSList.EBDSNCombo txtDSNName 
      Height          =   315
      Left            =   4620
      TabIndex        =   17
      Top             =   2325
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      Enabled         =   0   'False
      Enabled         =   0   'False
   End
   Begin VB.Label lblRider 
      Caption         =   "You can List the DTS Packages from a DSN, but you must fill in the Server Name, Login ID and Password to run a chosen Package."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   4500
      TabIndex        =   18
      Top             =   1650
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.Label lblDTSPackagesList 
      AutoSize        =   -1  'True
      Caption         =   "DTS Packages will be shown below:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   105
      TabIndex        =   16
      Top             =   225
      Width           =   2955
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Enabled         =   0   'False
      Height          =   195
      Left            =   4620
      TabIndex        =   7
      Top             =   3870
      Width           =   735
   End
   Begin VB.Label lblID 
      AutoSize        =   -1  'True
      Caption         =   "Login ID:"
      Enabled         =   0   'False
      Height          =   195
      Left            =   4620
      TabIndex        =   6
      Top             =   3315
      Width           =   645
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      DrawMode        =   1  'Blackness
      X1              =   4620
      X2              =   7140
      Y1              =   2730
      Y2              =   2730
   End
   Begin VB.Label lblDTSSvrNAme 
      AutoSize        =   -1  'True
      Caption         =   "DTS Server Name:"
      Enabled         =   0   'False
      Height          =   195
      Left            =   4620
      TabIndex        =   3
      Top             =   2775
      Width           =   1350
   End
   Begin VB.Label lblDSNNameMSDB 
      AutoSize        =   -1  'True
      Caption         =   "DSN Name to msdb database:"
      Enabled         =   0   'False
      Height          =   195
      Left            =   4620
      TabIndex        =   2
      Top             =   2085
      Width           =   2160
   End
   Begin VB.Menu m_File 
      Caption         =   "&File"
      Begin VB.Menu m_About 
         Caption         =   "About"
      End
      Begin VB.Menu m_Spacer 
         Caption         =   "-"
      End
      Begin VB.Menu m_Exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "DTSPackages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This object is set up to retrieve the Server listing for SQL
Private objApplication As New SQLDMO.Application
Private Sub cmdGetDTSList_Click()

    Dim cnn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim sSQL As String
    
    On Error GoTo Err_Trap
    If DSN.Value = True Then
        If Len(txtDSNName.DSN) = 0 Then
            MsgBox "Choose a DSN!", , "Nothing Chosen"
            Exit Sub
        End If
    ElseIf DSNLess.Value = True Then
        If Len(cboServers.Text) = 0 Or Len(txtLogin.Text) = 0 Then
            MsgBox "Choose a Server and Login", , "Nothing Chosen"
            Exit Sub
        End If
    End If
    'Here we select the name of the packages from the msdb database sysdtspackages table
    sSQL = "SELECT DISTINCT name FROM sysdtspackages"
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    sb.SimpleText = "Running......."
    If DSN.Value = True Then
        With cnn
            .ConnectionString = "DATA SOURCE=" & txtDSNName.DSN
            .CursorLocation = adUseClient
            .Open
            'process the stored procedure command with no records to return
            Set rst = .Execute(sSQL)
        End With
    Else
        With cnn
            .ConnectionString = "driver={SQL Server};server=" & cboServers.Text & ";uid=" & txtLogin.Text & ";pwd=" & txtPwd.Text & ";database=msdb"
            .ConnectionTimeout = 30
            .Open
            Set rst = .Execute(sSQL)
        End With
    End If
    ssdbDTSPkgs.Clear
'**********************************
    'Load the listbox
    While Not rst.EOF
        With ssdbDTSPkgs
            .AddItem rst("name").Value
        End With
        rst.MoveNext
    Wend
'**********************************
    If DSN.Value = True Then
        lblDTSPackagesList.Caption = "DTS Packages for DSN: " & txtDSNName.DSN
    ElseIf DSNLess.Value = True Then
        lblDTSPackagesList.Caption = "DTS Packages on Server: " & cboServers.Text
    End If
    DSNLess.Value = True
    Call DSNLess_Click
    
    Set rst = Nothing
    Set cnn = Nothing
    sb.SimpleText = "Finished!"
Exit_Sub:
    Exit Sub
    
Err_Trap:
    MsgBox "You possibly entered a Server or DSN" & vbCrLf & "that doesn't have DTS Packages or a name that doesn't exist" & vbCrLf & vbCrLf _
    & "Check your values and try again!", vbCritical, "Problem with entries"

End Sub

Private Sub cmdRunDTSPkg_Click()
    Dim myDTS As DTS.Package
    Dim intSel As Long
    Dim X As Long
    intSel = ssdbDTSPkgs.SelCount
    
    If Len(cboServers.Text) = 0 Or Len(txtLogin.Text) = 0 Then
        MsgBox "Choose a Server and Login", , "Nothing Chosen"
        Exit Sub
    End If
    
    For X = 1 To intSel
        Set myDTS = New DTS.Package
        sb.SimpleText = "Running......."
        myDTS.LoadFromSQLServer cboServers.Text, txtLogin.Text, txtPwd.Text, 256, , , , ssdbDTSPkgs.List(X)
    
        ' Set ExecuteInMainThread for ALL Steps
        Dim i As Integer
        For i = 1 To myDTS.Steps.Count
            myDTS.Steps(i).ExecuteInMainThread = True
        Next i
        'execute all step in the DTS Package
        myDTS.Execute
        sb.Font.Bold = True
        sb.SimpleText = "Finished!"
    Next
    myDTS.UnInitialize
    Set myDTS = Nothing
    frmDSN.Enabled = True
    
End Sub

Private Sub cmdRunDTSPkg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    txtDSNName.Visible = False
    lblRider.Visible = True
    
End Sub

Private Sub DSN_Click()
    lblDSNNameMSDB.Enabled = True
    cmdGetDTSList.Enabled = True
    txtDSNName.Enabled = True
    lblDTSSvrNAme.Enabled = False
    lblID.Enabled = False
    cboServers.Enabled = False
    txtLogin.Enabled = False
    txtPwd.Enabled = False
    lblPassword.Enabled = False
End Sub

Private Sub DSNLess_Click()
    lblDSNNameMSDB.Enabled = False
    txtDSNName.Enabled = False
    cmdGetDTSList.Enabled = True
    lblDTSSvrNAme.Enabled = True
    lblID.Enabled = True
    lblPassword.Enabled = True
    cboServers.Enabled = True
    txtLogin.Enabled = True
    txtPwd.Enabled = True
End Sub

Private Sub Form_Load()
    'Get a list of Servers from my Server group using the SQLDMO
    'This code is compliments of Jeremy van Dijk who has listed the
    'SQL Scripter Tool on www.planet-source-code.com
    Dim objServerGroup As SQLDMO.ServerGroup
    Dim objRegisteredServer As SQLDMO.RegisteredServer
    Dim i As Integer, j As Integer
    
    For Each objServerGroup In objApplication.ServerGroups
        For Each objRegisteredServer In objServerGroup.RegisteredServers
            cboServers.AddItem objRegisteredServer.Name
            cboServers.ItemData(cboServers.NewIndex) = CStr(objRegisteredServer.UseTrustedConnection)
        Next objRegisteredServer
    Next objServerGroup
    
    txtDSNName.DriverFilter = "SQL Server"
    
    DSN.Value = False
    DSNLess.Value = False
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    txtDSNName.Visible = True
    lblRider.Visible = False

End Sub

Private Sub m_About_Click()
    Me.Hide
    Load DTSAbout
    DTSAbout.Show
End Sub

Private Sub m_Exit_Click()
    Unload Me
End Sub

Private Sub ssdbDTSPkgs_Click()
    If ssdbDTSPkgs.ListCount > 0 Then
        cmdRunDTSPkg.Enabled = True
    End If
End Sub
