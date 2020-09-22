VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form DTSPackages 
   Caption         =   "DTS Packages Test"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   16
      Top             =   4005
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12383
         EndProperty
      EndProperty
   End
   Begin VB.Frame frmDSN 
      Caption         =   "DSN or DSN Less?"
      Height          =   690
      Left            =   4620
      TabIndex        =   11
      Top             =   735
      Width           =   2475
      Begin VB.OptionButton DSNLess 
         Height          =   270
         Left            =   1200
         TabIndex        =   13
         Top             =   285
         Width           =   285
      End
      Begin VB.OptionButton DSN 
         Height          =   270
         Left            =   150
         TabIndex        =   12
         Top             =   285
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "DSN-Less"
         Height          =   195
         Left            =   1485
         TabIndex        =   15
         Top             =   315
         Width           =   720
      End
      Begin VB.Label lblDSN 
         AutoSize        =   -1  'True
         Caption         =   "DSN"
         Height          =   195
         Left            =   435
         TabIndex        =   14
         Top             =   315
         Width           =   345
      End
   End
   Begin VB.TextBox txtPwd 
      Enabled         =   0   'False
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   4770
      PasswordChar    =   "*"
      TabIndex        =   8
      ToolTipText     =   "Password for Login to Server"
      Top             =   3495
      Width           =   1950
   End
   Begin VB.TextBox txtLogin 
      Enabled         =   0   'False
      Height          =   330
      Left            =   4770
      TabIndex        =   7
      ToolTipText     =   "Login ID for Server"
      Top             =   2925
      Width           =   1950
   End
   Begin VB.TextBox txtServerName 
      Enabled         =   0   'False
      Height          =   330
      Left            =   4770
      TabIndex        =   5
      ToolTipText     =   "Server Name to Login To"
      Top             =   2370
      Width           =   1950
   End
   Begin VB.TextBox txtDSNName 
      Enabled         =   0   'False
      Height          =   330
      Left            =   4770
      TabIndex        =   3
      Top             =   1710
      Width           =   1950
   End
   Begin VB.CommandButton cmdRunDTSPkg 
      Caption         =   "Run DTS Pkg"
      Enabled         =   0   'False
      Height          =   420
      Left            =   5895
      TabIndex        =   2
      Top             =   240
      Width           =   1290
   End
   Begin VB.CommandButton cmdGetDTSList 
      Caption         =   "List DTS Pkgs"
      Height          =   420
      Left            =   4575
      TabIndex        =   1
      Top             =   240
      Width           =   1290
   End
   Begin SSDataWidgets_B.SSDBGrid ssdbDTSPkgs 
      Height          =   3735
      Left            =   105
      TabIndex        =   0
      Top             =   240
      Width           =   4425
      _Version        =   196617
      DataMode        =   2
      RecordSelectors =   0   'False
      AllowAddNew     =   -1  'True
      AllowUpdate     =   0   'False
      MultiLine       =   0   'False
      RowSelectionStyle=   1
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeRow   =   1
      SelectByCell    =   -1  'True
      ForeColorEven   =   0
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns(0).Width=   7514
      Columns(0).Caption=   "DTS Package to Run"
      Columns(0).Name =   "DTSPkgRun"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).HasForeColor=   -1  'True
      Columns(0).HasBackColor=   -1  'True
      Columns(0).BackColor=   16777215
      _ExtentX        =   7805
      _ExtentY        =   6588
      _StockProps     =   79
      Caption         =   "DTS Packages Available for Execution"
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Enabled         =   0   'False
      Height          =   195
      Left            =   4755
      TabIndex        =   10
      Top             =   3270
      Width           =   735
   End
   Begin VB.Label lblID 
      AutoSize        =   -1  'True
      Caption         =   "Login ID:"
      Enabled         =   0   'False
      Height          =   195
      Left            =   4770
      TabIndex        =   9
      Top             =   2715
      Width           =   645
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      DrawMode        =   1  'Blackness
      X1              =   4650
      X2              =   6855
      Y1              =   2130
      Y2              =   2130
   End
   Begin VB.Label lblDTSSvrNAme 
      AutoSize        =   -1  'True
      Caption         =   "DTS Server Name:"
      Enabled         =   0   'False
      Height          =   195
      Left            =   4770
      TabIndex        =   6
      Top             =   2175
      Width           =   1350
   End
   Begin VB.Label lblDSNNameMSDB 
      AutoSize        =   -1  'True
      Caption         =   "DSN Name to msdb database:"
      Enabled         =   0   'False
      Height          =   195
      Left            =   4770
      TabIndex        =   4
      Top             =   1485
      Width           =   2160
   End
End
Attribute VB_Name = "DTSPackages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdGetDTSList_Click()

    Dim cnn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim sSQL As String
    
    On Error GoTo Err_Trap
    
    sSQL = "SELECT DISTINCT name FROM sysdtspackages"
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    sb.SimpleText = "Running......."
    If DSN.Value = True Then
        With cnn
            .ConnectionString = "DATA SOURCE=" & txtDSNName.Text
            .CursorLocation = adUseClient
            .Open
            'process the stored procedure command with no records to return
            Set rst = .Execute(sSQL)
        End With
    Else
        With cnn
            .ConnectionString = "driver={SQL Server};server=" & txtServerName.Text & ";uid=" & txtLogin.Text & ";pwd=" & txtPwd.Text & ";database=msdb"
            .ConnectionTimeout = 30
            .Open
            Set rst = .Execute(sSQL)
        End With
    End If
    ssdbDTSPkgs.RemoveAll
    
    While Not rst.EOF
        With ssdbDTSPkgs
            .AddNew
            .Columns(0).Text = rst("name").Value
            .Update
        End With
        rst.MoveNext
    Wend

    frmDSN.Enabled = False
    cmdRunDTSPkg.Enabled = True
    Call DSNLess_Click
    
    Set rst = Nothing
    Set cnn = Nothing
    sb.SimpleText = "Finished!"
Exit_Sub:
    Exit Sub
    
Err_Trap:
    MsgBox "You possibly entered a Server or DSN name that doesn't exist" & vbCrLf & vbCrLf _
    & "Check your values and try again!", vbCritical, "Problem with entries"

End Sub

Private Sub cmdRunDTSPkg_Click()
    Dim oPKG As New DTS.Package

    sb.SimpleText = "Running......."
    oPKG.LoadFromSQLServer txtServerName.Text, txtLogin.Text, txtPwd.Text, 256, , , , ssdbDTSPkgs.Columns(0).Text

    ' Set ExecuteInMainThread for ALL Steps
    Dim i As Integer
    For i = 1 To oPKG.Steps.Count
        oPKG.Steps(i).ExecuteInMainThread = True
    Next i
  
    oPKG.Execute
    
    sb.Font.Bold = True
    sb.SimpleText = "Finished!"
    
    oPKG.UnInitialize
    Set oPKG = Nothing

End Sub

Private Sub DSN_Click()
    lblDSNNameMSDB.Enabled = True
    txtDSNName.Enabled = True
    lblDTSSvrNAme.Enabled = False
    lblID.Enabled = False
    txtServerName.Enabled = False
    txtLogin.Enabled = False
    txtPwd.Enabled = False
    lblPassword.Enabled = False
End Sub

Private Sub DSNLess_Click()
    lblDSNNameMSDB.Enabled = False
    txtDSNName.Enabled = False
    lblDTSSvrNAme.Enabled = True
    lblID.Enabled = True
    lblPassword.Enabled = True
    txtServerName.Enabled = True
    txtLogin.Enabled = True
    txtPwd.Enabled = True
End Sub
