VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QueryBom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This is Industry Bill Interface Class, made by K3BOSPLUGINSWIZAED"
Option Explicit
Option Base 0

Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Private strConn As String   '数据库连接字串
Private UserId As Long      '当前用户内码
Private UserName As String  '当前用户名称


'K3Connection示例：
'ConnectString={Persist Security Info=True;Provider=SQLOLEDB.1;User ID=sa;Password=123456;Data Source=10.211.55.21;Initial Catalog=AIS20091105191238};UserName=administrator;UserID=16394;DBMS Name=Microsoft SQL Server;DBMS Version=2000/2005;SubID=xs;AcctType=gy;Setuptype=Industry;Language=chs;IP=10.211.55.19;K3Version=KUE;MachineName=SNOKYEA68F;UUID=8EFACFB9-192C-4EDD-9AFE-610EEF9705E8
'数据库连接示例
'Persist Security Info=True;Provider=SQLOLEDB.1;User ID=sa;Password=123456;Data Source=10.211.55.21;Initial Catalog=AIS20091105191238

Public Sub MainFunction(ByVal sKey As String, oList As Object, ByRef bCancel As Boolean)
    If sKey = "btnQueryBom" Then
        If oList.ListID = 32 Then
            On Error GoTo ex
            
            '  --  1  --  查询BOM（bomid）
            Dim interid As Long
            Dim entryid As Long
            Dim sql As String
            Dim rs As ADODB.Recordset
            Dim bomid As Long
            
            interid = oList.ListSelectBillinfo(1)("FInterID")
            entryid = oList.ListSelectBillinfo(1)("FEntryID")
            
            Prepare
            sql = "SELECT FInterID FROM ICBOM WHERE FUseStatus=1072 AND FItemID=(SELECT FItemID FROM SEOrderEntry WHERE FInterID=" & interid & " AND FEntryID=" & entryid & ")"
            Set rs = GetData(sql)
            bomid = -1
            If Not rs.EOF Then bomid = rs.Fields(0)
            Set rs = Nothing
            
            '  --  2  --  显示BOM
            If bomid <> -1 Then
                ShowBom bomid
            Else
                MsgBox "该产品没有在使用的BOM", vbOKOnly, "金蝶提示"
            End If
        Else
            MsgBox "该功能只在销售出库序时簿中可用", vbOKOnly, "金蝶提示"
        End If
    End If
    Exit Sub
ex:
    MsgBox Err.Description
End Sub

Private Sub ShowBom(ByVal bomid)
    Dim objBill As Object
    Dim dlg As Object
    Dim LocalCnStr As String
    Dim sSubID As String
    Dim sSubName As String
    Dim lModel As Long
    Dim lModelDetail As Long
    
    On Error GoTo ex
    
    Set objBill = CreateObject("K3BillsEx.Bills")
    Set dlg = CreateObject("CSystemDlg.Sys")
    LocalCnStr = dlg.LocalCnn
    Set dlg = Nothing
    
    With objBill
        .LocalCnn = LocalCnStr
        .SystemName = sSubName
        .SetOpt UserId, UserName
        If Not .SaveVect(1).Lookup("sDsn") Then .SaveVect(1)("sDsn") = strConn
        .ListRecordset = SetBillRec(bomid, 50)
        .ListRSFieldVect = SetBillVect
        .NewBillTransType = "50"

        .Show 900, 2
    End With
    
    Set objBill = Nothing
    Exit Sub
ex:
    MsgBox Err.Description
    Set objBill = Nothing
End Sub

Private Function SetBillRec(ByVal aInterID As Long, ByVal aTranType As Long) As ADODB.Recordset
    Dim tRec As ADODB.Recordset
    Set tRec = New ADODB.Recordset
    tRec.Fields.Append "FInterID", adInteger
    tRec.Fields.Append "FTranType", adInteger
    tRec.Open
    tRec.AddNew
    tRec!FInterID = aInterID
    tRec!FTranType = aTranType
    tRec.Update
    Set SetBillRec = tRec
End Function

Private Function SetBillVect() As KFO.Vector
    Dim tVect As KFO.Vector
    Set tVect = New KFO.Vector
    Dim tDict As KFO.Dictionary
    Set tDict = New KFO.Dictionary
    tDict("FColName") = "FInterID"
    tDict("FISPrimary") = 1
    tVect.Add tDict
    
    Set tDict = New KFO.Dictionary
    tDict("FColName") = "FTranType"
    tDict("FISPrimary") = 3
    tVect.Add tDict
    
    Set tDict = Nothing
    Set SetBillVect = tVect
End Function

'从当前进程中获取K3连接字串
Private Function GetK3Conn() As String
    Dim lProc As Long
    Dim spmMgr As Object
    Dim cnn As String
    
    lProc = GetCurrentProcessId()
    Set spmMgr = CreateObject("PropsMgr.ShareProps")
    GetK3Conn = spmMgr.GetProperty(lProc, "PropsString")
    Set spmMgr = Nothing
     'delete===============
    If lProc = 3460 Then
        GetK3Conn = "ConnectString={Persist Security Info=True;Provider=SQLOLEDB.1;User ID=sa;Password=123456;Data Source=10.211.55.21;Initial Catalog=AIS20091105191238};UserName=administrator;UserID=16394;DBMS Name=Microsoft SQL Server;DBMS Version=2000/2005;SubID=xs;AcctType=gy;Setuptype=Industry;Language=chs;IP=10.211.55.19;K3Version=KUE;MachineName=SNOKYEA68F;UUID=8EFACFB9-192C-4EDD-9AFE-610EEF9705E8"
    End If
End Function

'从K3连接字串中获取数据库连接字串
Private Function GetConn(ByVal k3Conn As String) As String
    Dim cnn As String
    cnn = k3Conn
    cnn = Left(cnn, InStr(1, cnn, "}") - 1)
    cnn = Mid(cnn, 16)
    GetConn = cnn
End Function

'准备，获取连接字串，用户ID等
Private Sub Prepare()
    If strConn = "" Then
        Dim k3cnn As String
        Dim tmp As String
        
        k3cnn = GetK3Conn
        strConn = GetConn(k3cnn)
        tmp = k3cnn
        tmp = Mid(tmp, InStr(1, tmp, "UserID=") + 7)
        tmp = Left(tmp, InStr(1, tmp, ";") - 1)
        UserId = VBA.CStr(tmp)
        tmp = k3cnn
        tmp = Mid(tmp, InStr(1, tmp, "UserName=") + 9)
        tmp = Left(tmp, InStr(1, tmp, ";") - 1)
        UserName = tmp
    End If
End Sub

'执行SQL，返回recordset
Private Function GetData(ByVal strSql As String) As ADODB.Recordset
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    On Error GoTo ex
    
    Set cn = New ADODB.Connection
    cn.ConnectionString = strConn
    cn.Open
    Set rs = cn.Execute(strSql)
    Set GetData = rs
    
    'cn.Close
ex:
    Set cn = Nothing
End Function

































