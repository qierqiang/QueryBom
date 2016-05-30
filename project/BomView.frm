VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form BomView 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "查看BOM - 仅查看"
   ClientHeight    =   8295
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14430
   Icon            =   "BomView.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   14430
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text_7 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   12360
      TabIndex        =   47
      Top             =   7440
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      CausesValidation=   0   'False
      Height          =   3135
      Left            =   120
      TabIndex        =   44
      Top             =   4200
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   5530
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text_6 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   8760
      TabIndex        =   43
      Top             =   7890
      Width           =   2000
   End
   Begin VB.TextBox Text_3 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   5040
      TabIndex        =   40
      Top             =   7440
      Width           =   2000
   End
   Begin VB.TextBox Text_2 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   1500
      TabIndex        =   38
      Top             =   7890
      Width           =   2000
   End
   Begin VB.TextBox Text_1 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   1500
      TabIndex        =   37
      Top             =   7440
      Width           =   2000
   End
   Begin VB.TextBox Text_5 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   8760
      TabIndex        =   36
      Top             =   7440
      Width           =   2000
   End
   Begin VB.TextBox Text_4 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   5040
      TabIndex        =   34
      Top             =   7890
      Width           =   2000
   End
   Begin VB.TextBox Text16 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   1500
      ScrollBars      =   2  'Vertical
      TabIndex        =   31
      Top             =   3360
      Width           =   12795
   End
   Begin VB.TextBox Text15 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   12240
      TabIndex        =   29
      Top             =   2422
      Width           =   2000
   End
   Begin VB.TextBox Text14 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   12240
      TabIndex        =   27
      Top             =   1972
      Width           =   2000
   End
   Begin VB.TextBox Text13 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   12240
      TabIndex        =   25
      Top             =   1522
      Width           =   2000
   End
   Begin VB.TextBox Text12 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   8580
      TabIndex        =   23
      Top             =   2422
      Width           =   2000
   End
   Begin VB.TextBox Text11 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   8580
      TabIndex        =   21
      Top             =   1972
      Width           =   2000
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   5040
      TabIndex        =   18
      Top             =   1522
      Width           =   2000
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   2700
      TabIndex        =   17
      Top             =   2872
      Width           =   795
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   1500
      TabIndex        =   13
      Top             =   2422
      Width           =   2000
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   1500
      TabIndex        =   12
      Top             =   1972
      Width           =   2000
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   1500
      TabIndex        =   11
      Top             =   1522
      Width           =   2000
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   1500
      TabIndex        =   10
      Top             =   1072
      Width           =   2000
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   8580
      TabIndex        =   8
      Top             =   1522
      Width           =   2000
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   5040
      TabIndex        =   6
      Top             =   2872
      Width           =   2000
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   5040
      TabIndex        =   5
      Top             =   2422
      Width           =   2000
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   5040
      TabIndex        =   4
      Top             =   1972
      Width           =   2000
   End
   Begin VB.Label lChecked 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "审核"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   12600
      TabIndex        =   48
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label24 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "最后更新日期："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10920
      TabIndex        =   46
      Top             =   7455
      Width           =   1500
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BOM单"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6720
      TabIndex        =   45
      Top             =   360
      Width           =   960
   End
   Begin VB.Label Label22 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "PDM导入时间："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7200
      TabIndex        =   42
      Top             =   7905
      Width           =   1500
   End
   Begin VB.Label Label21 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "建立日期："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3780
      TabIndex        =   41
      Top             =   7455
      Width           =   1200
   End
   Begin VB.Label Label20 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "审核日期："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3780
      TabIndex        =   39
      Top             =   7905
      Width           =   1200
   End
   Begin VB.Label Label19 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "最后更新人员："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7200
      TabIndex        =   35
      Top             =   7455
      Width           =   1500
   End
   Begin VB.Label Label18 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "审核人员："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   33
      Top             =   7905
      Width           =   1200
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "建立人员："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   32
      Top             =   7455
      Width           =   1200
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "备     注："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   3375
      Width           =   1200
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "成品率(%)："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10920
      TabIndex        =   28
      Top             =   2430
      Width           =   1200
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "物料属性："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10920
      TabIndex        =   26
      Top             =   1980
      Width           =   1200
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "图    号："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10920
      TabIndex        =   24
      Top             =   1530
      Width           =   1200
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "数     量："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7320
      TabIndex        =   22
      Top             =   2430
      Width           =   1200
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "规     格："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7320
      TabIndex        =   20
      Top             =   1980
      Width           =   1200
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "版    本："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3780
      TabIndex        =   19
      Top             =   1530
      Width           =   1200
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "跳    层："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3780
      TabIndex        =   16
      Top             =   2880
      Width           =   1200
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "工艺路线："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3780
      TabIndex        =   15
      Top             =   2430
      Width           =   1200
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "物料名称："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3780
      TabIndex        =   14
      Top             =   1980
      Width           =   1200
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "是否特性配置来源物料："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2880
      Width           =   2400
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "状     态："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7320
      TabIndex        =   7
      Top             =   1530
      Width           =   1200
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "单     位："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2430
      Width           =   1200
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "物料代码："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1980
      Width           =   1200
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "BOM单组别："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1200
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "BOM单编号："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1530
      Width           =   1200
   End
End
Attribute VB_Name = "BomView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const HWND_TOPMOST = -1
Const SWP_SHOWWINDOW = &H40
Const iniFileName = "c:\colWidth.ini"

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Dim m_billInfo As Recordset
Dim m_entryInfo As Recordset

Public Property Set BillInfo(ByRef value As Recordset)
    Set m_billInfo = value
End Property

Public Property Set EntryInfo(ByRef value As Recordset)
    Set m_entryInfo = value
End Property

Private Sub Form_load()
    On Error GoTo ex
    
    '保持最前
    SetWindowPos Me.hwnd, HWND_TOPMOST, Me.CurrentX, Me.CurrentY, 980, 590, SWP_SHOWWINDOW
    
    '设置只读及清空
    Dim ctrl As Control
    For Each ctrl In Me
        If LCase(TypeName(ctrl)) = "textbox" Then ctrl.Locked = True
    Next ctrl
    
    '显示字段
    FillForm
    
    '审核状态
    If m_billInfo.Fields("FStatus") = 0 Then lChecked.Visible = False
    
    '加载列宽
    LoadColWidths
    Exit Sub
ex:
    MsgBox Err.Description, vbOKOnly, "金蝶提示"
    Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveColWidths
End Sub

Private Sub FillForm()
    
    Text1.Text = NullCheck(m_billInfo.Fields("FParentIDName"))
    Text2.Text = NullCheck(m_billInfo.Fields("FBOMNumber"))
    Text3.Text = NullCheck(m_billInfo.Fields("FItemIDNumber"))
    Text4.Text = NullCheck(m_billInfo.Fields("FUnitIDNumber"))
    If m_billInfo.Fields("FBomType") = 4 Then Text5.Text = "是" Else Text5.Text = "否"
    Text6.Text = NullCheck(m_billInfo.Fields("FVersion"))
    Text7.Text = NullCheck(m_billInfo.Fields("FItemIDName"))
    Text8.Text = NullCheck(m_billInfo.Fields("FRoutingIDName"))
    Text9.Text = NullCheck(m_billInfo.Fields("FBOMSkipName"))
    Text10.Text = NullCheck(m_billInfo.Fields("FUseStatusName"))
    Text11.Text = NullCheck(m_billInfo.Fields("FModel"))
    Text12.Text = VBA.Round(VBA.CDec(m_billInfo.Fields("FAuxQty")), 3)
    Text13.Text = NullCheck(m_billInfo.Fields("FChartNumber"))
    Text14.Text = NullCheck(m_billInfo.Fields("FErpClsID"))
    Text15.Text = VBA.Round(VBA.CDec(m_billInfo.Fields("FYield")))
    Text16.Text = NullCheck(m_billInfo.Fields("FNote"))
    
    Text_1.Text = NullCheck(m_billInfo.Fields("FCheckIDName"))
    Text_2.Text = NullCheck(m_billInfo.Fields("FCheckerIDName"))
    Text_3.Text = NullCheck(m_billInfo.Fields("FCheckDate"))
    Text_4.Text = NullCheck(m_billInfo.Fields("FAudDate"))
    Text_5.Text = NullCheck(m_billInfo.Fields("FOperatorIDName"))
    Text_6.Text = NullCheck(m_billInfo.Fields("FPDMImportDate"))
    Text_7.Text = NullCheck(m_billInfo.Fields("FEnterTime"))
    
    Set DataGrid1.DataSource = m_entryInfo
    DataGrid1.Columns(7).Alignment = dbgRight
    DataGrid1.Columns(8).Alignment = dbgRight
End Sub

Private Function NullCheck(ByVal value As Variant) As String
    If IsNull(value) Then
        NullCheck = ""
    Else
        NullCheck = VBA.CStr(value)
    End If
End Function

Private Sub LoadColWidths()
    Dim i, w, nSize As Long
    Dim s, nStr As String
    
    On Error Resume Next
    nSize = 255: nStr = String(nSize, 0)
    
    For i = 0 To 27
        GetPrivateProfileString "BomViewColWidth", CStr(i), "1000", nStr, nSize, iniFileName
        w = CLng(nStr)
        DataGrid1.Columns(i).Width = w
    Next i
End Sub

Private Sub SaveColWidths()
    Dim i As Integer
    
    On Error Resume Next
    
    For i = 0 To 27
        WritePrivateProfileString "BomViewColWidth", CStr(i), CStr(DataGrid1.Columns(i).Width), iniFileName
    Next i
End Sub






























