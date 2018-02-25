VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmTipSummary 
   Caption         =   "Form1"
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   9120
   StartUpPosition =   3  'Windows Default
   Begin MSDataListLib.DataList lstCat 
      Bindings        =   "frmTipSummary.frx":0000
      Height          =   3960
      Left            =   60
      TabIndex        =   3
      Top             =   540
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   6985
      _Version        =   393216
      ListField       =   "Name"
      BoundColumn     =   "Name"
   End
   Begin RichTextLib.RichTextBox rtbTip 
      Height          =   4155
      Left            =   3960
      TabIndex        =   0
      Top             =   480
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   7329
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmTipSummary.frx":0017
   End
   Begin MSAdodcLib.Adodc datCatRS 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   4920
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "PROVIDER=MSDASQL;dsn=TIP;uid=;pwd=;"
      OLEDBString     =   "PROVIDER=MSDASQL;dsn=TIP;uid=;pwd=;"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select Serial, Name, Enabled from Category WHERE ENABLED <> 0"
      Caption         =   " "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc datTipRS 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "PROVIDER=MSDASQL;dsn=TIP;uid=;pwd=;"
      OLEDBString     =   "PROVIDER=MSDASQL;dsn=TIP;uid=;pwd=;"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from Tip"
      Caption         =   " "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label lblTip 
      Caption         =   "Tip"
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   180
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Category"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   1755
   End
End
Attribute VB_Name = "frmTipSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Tips() As Tip


Private Sub Form_Load()
    Dim i As Integer
    Dim randNum As Integer
    randNum = Date2Num(Date)
    i = datCatRS.Recordset.RecordCount
    ReDim Tips(i)
    datCatRS.Recordset.MoveFirst
    i = 0
    While Not datCatRS.Recordset.EOF
        datTipRS.RecordSource = "SELECT * FROM TIP WHERE Category_Serial = " & datCatRS.Recordset!Serial
        datTipRS.Refresh
        If (datTipRS.Recordset.RecordCount > 0) Then
            datTipRS.Recordset.Move (randNum Mod (datTipRS.Recordset.RecordCount - 1))
            If Not IsNull(datTipRS.Recordset!Text) Then
                Tips(i).Text = datTipRS.Recordset!Text
            End If
            If Not IsNull(datTipRS.Recordset!Name) Then Tips(i).Name = datTipRS.Recordset!Name
        End If
        'MsgBox Tips(i).Name
        i = i + 1
        datCatRS.Recordset.MoveNext
    Wend
End Sub

Private Sub List1_Click()
End Sub

Private Sub lstCat_Click()
    If lstCat.SelectedItem <> -1 Then
        rtbTip.TextRTF = Tips(lstCat.SelectedItem - 1).Text
        lblTip.Caption = "Tip: " & Tips(lstCat.SelectedItem - 1).Name
    End If
End Sub
