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
            Dim rs As Object 'ADODB.Recordset
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
                showbom bomid
            Else
                MsgBox "该产品没有在使用的BOM", vbOKOnly, "金蝶提示"
            End If
        Else
            MsgBox "该功能只在销售出库序时簿中可用", vbOKOnly, "金蝶提示"
        End If
    End If
    Exit Sub
ex:
    MsgBox Err.Description, vbOKOnly, "金蝶提示"
End Sub

Private Sub showbom(ByVal bomid)
    Dim sql As String
    Dim rsBill, rsEntry As Object ' ADODB.Recordset
    
    sql = "SELECT ICBom.FInterID, ICBom.FBOMNumber, ICBom.FVersion, ICBom.FUseStatus, t002.FID as  FUseStatusNumber, t002.FName as  FUseStatusName, ICBom.FItemID, t003.FNumber as  FItemIDNumber, t003.FName as  FItemIDName, t003.FQtyDecimal, " & _
        "t003.FPriceDecimal, t003.FUnitGroupID, ICBom.FUnitID, t004.FNumber as  FUnitIDNumber, t004.FName as  FUnitIDName, t003.FName as FName, t003.FModel as FModel,(SELECT FName FROM t_SubMessage WHERE t003.FErpClsID=FInterID) as FErpClsID, " & _
        "ICBom.FAuxPropID, t008.FNumber as  FAuxPropIDNumber, t008.FName as  FAuxPropIDName, ICBom.FAuxQty, ICBom.FYield, ICBom.FNote, ICBom.FCheckID, t012.FDescription as  FCheckIDNumber, t012.FName as  FCheckIDName, ICBom.FCheckDate, " & _
        "ICBom.FOperatorID, t014.FDescription as  FOperatorIDNumber, t014.FName as  FOperatorIDName, ICBom.FEntertime, t003.FChartNumber as FChartNumber, t003.FFixLeadTime as FLeadTime, ICBom.FRoutingID, t018.FBillNO as  FRoutingIDNumber, " & _
        "t018.FRoutingName as  FRoutingIDName, ICBom.FBomType, ICBom.FCustID, ICBom.FParentID, t021.FNumber as  FParentIDNumber, t021.FName as  FParentIDName, ICBom.FCheckerID, t022.FDescription as  FCheckerIDNumber, t022.FName as " & _
        "FCheckerIDName, ICBom.FAudDate, ICBom.FImpMode, t003.FAuxClassID as FAuxPropCls, ICBom.FPDMImportDate, t003.FIsCharSourceItem as FIsCharSourceItem, ICBom.FBOMSkip, t028.FID as  FBOMSkipNumber, t028.FName as  FBOMSkipName, ICBom.FBrNo, " & _
        "ICBom.FTranType , ICBom.FCancellation, ICBom.FStatus " & _
        " FROM ICBom left join t_SubMessage  t002 on ICBom.FUseStatus= t002.FInterID  AND t002.FInterID<>0  join t_ICItem  t003 on ICBom.FItemID= t003.FItemID  AND t003.FItemID<>0  left join t_MeasureUnit  t004 on ICBom.FUnitID= t004.FItemID " & _
        "AND t004.FItemID<>0  left join t_AuxItem  t008 on ICBom.FAuxPropID= t008.FItemID  AND t008.FItemID<>0  left join t_User  t012 on ICBom.FCheckID= t012.FUserID  AND t012.FUserID<>0  left join t_User  t014 on ICBom.FOperatorID= " & _
        "t014.FUserID  AND t014.FUserID<>0  left join t_Routing  t018 on ICBom.FRoutingID= t018.FInterID  AND t018.FInterID<>0  left join IcBomGroup  t021 on ICBom.FParentID= t021.FInterID  AND t021.FInterID<>0  left join t_User  t022 on " & _
        "ICBom.FCheckerID= t022.FUserID  AND t022.FUserID<>0  left join t_SubMessage  t028 on ICBom.FBOMSkip= t028.FInterID  AND t028.FInterID<>0 " & _
        " WHERE ICBom.FInterID =" & bomid
    Set rsBill = GetData(sql)
    sql = "SELECT FEntryID AS '顺序号',FItemIDNumber AS '物料代码',FItemIDName AS '物料名称',FModel AS '规格型号',FErpCls AS '物料属性',FUnitIDName AS '单位',FMaterielTypeName AS '子项类型',FAuxQty AS '用量',FScrap AS '损耗率(%)',FPositionNo AS '位置号',FItemSize AS '坯料尺寸',FItemSuite AS '坯料数',FOperSN AS '工序号',FOperIDName AS '工序',FMachinePos AS '工位',FOffSetDay AS '提前期偏置',FBackFlushName AS '是否倒冲',CASE FIsKeyItem WHEN 0 THEN '否' ELSE '是' END AS '关键件',FUseState AS '使用状态',CASE FForbitUse WHEN 0 THEN '否' ELSE '是' END AS '是否禁用',FStockIDName AS '发料仓库',FSPIDName AS '发料仓位',FNote AS '备注',FNote1 AS '备注1',FNote2 AS '备注2',FNote3 AS '备注3',FPDMImportDate AS 'PDM导入时间','N' AS '存在替代关系' FROM (" & _
        "SELECT ICBomChild.FInterID,ICBomChild.FEntryID, ICBomChild.FItemID, t002.FQtyDecimal, t002.FPriceDecimal, t002.FUnitGroupID, t002.FNumber as  FItemIDNumber, t002.FName as  FItemIDName, t002.FName as FItemName, " & _
        "t002.FModel as FModel,(SELECT FName FROM t_SubMessage WHERE t002.FErpClsID=FInterID) as FErpCls, ICBomChild.FHasChar, ICBomChild.FAuxPropID, t007.FNumber as  FAuxPropIDNumber, t007.FName as  FAuxPropIDName, ICBomChild.FUnitID, " & _
        "t008.FNumber as  FUnitIDNumber, t008.FName as  FUnitIDName, t009.FName as FBaseUnit, ICBomChild.FMaterielType, t010.FID as  FMaterielTypeNumber, t010.FName as  FMaterielTypeName, ICBomChild.FMarshalType, t011.FID as " & _
        "FMarshalTypeNumber, t011.FName as  FMarshalTypeName, ICBomChild.FQty, ICBomChild.FAuxQty, ICBomChild.FBeginDay, ICBomChild.FEndDay, ICBomChild.FPercent, ICBomChild.FScrap, ICBomChild.FPositionNo, ICBomChild.FItemSize, " & _
        "ICBomChild.FItemSuite, ICBomChild.FOperSN, ICBomChild.FOperSN as  FOperSNNumber, ICBomChild.FOperSN as  FOperSNName, ICBomChild.FOperID, t022.FID as  FOperIDNumber, t022.FName as  FOperIDName, ICBomChild.FMachinePos, " & _
        "ICBomChild.FOffSetDay, ICBomChild.FBackFlush, t025.FID as  FBackFlushNumber, t025.FName as  FBackFlushName, t002.FIsKeyItem as FIsKeyItem,(SELECT FName FROM t_SubMessage WHERE t002.FUseState=FInterID) as FUseState, t002.FDeleted as " & _
        "FForbitUse, ICBomChild.FStockID, t029.FNumber as  FStockIDNumber, t029.FName as  FStockIDName, ICBomChild.FSPID, t030.FNumber as  FSPIDNumber, t030.FName as  FSPIDName, ICBomChild.FNote, ICBomChild.FNote1, ICBomChild.FNote2, " & _
        "ICBomChild.FNote3, t002.FErpClsID as FErpClsID, t002.FAuxClassID as FAuxPropCls, ICBomChild.FPDMImportDate, ICBomChild.FDetailID,ICBomChild.FEntryID AS FOldEntryID,  ICBomChild.FBrNo " & _
        " FROM ICBomChild join t_ICItem  t002 on t002.FItemID = ICBomChild.FItemID AND t002.FItemID <>0  left join t_AuxItem  t007 on t007.FItemID = ICBomChild.FAuxPropID AND t007.FItemID <>0  left join t_MeasureUnit  t008 on t008.FItemID = " & _
        "ICBomChild.FUnitID AND t008.FItemID <>0  left join t_MeasureUnit t009 on  t009.FItemID=t002.FUnitID  AND t009.FItemID<>0  left join t_SubMessage  t010 on t010.FInterID = ICBomChild.FMaterielType AND t010.FInterID <>0  left join " & _
        "t_SubMessage  t011 on t011.FInterID = ICBomChild.FMarshalType AND t011.FInterID <>0  left join t_SubMessage  t022 on t022.FInterID = ICBomChild.FOperID AND t022.FInterID <>0  left join t_SubMessage  t025 on t025.FInterID = " & _
        "ICBomChild.FBackFlush AND t025.FInterID <>0  left join t_Stock  t029 on t029.FItemID = ICBomChild.FStockID AND t029.FItemID <>0  left join t_StockPlace  t030 on t030.FSPID = ICBomChild.FSPID AND t030.FSPID <>0 " & _
        " WHERE ICBomChild.FInterID = " & bomid & ") as datasource ORDER BY FEntryID"
    Set rsEntry = GetData(sql)
    
    Dim frm As New BomView
    Set frm.BillInfo = rsBill
    Set frm.EntryInfo = rsEntry
    frm.Show
    Set rsBill = Nothing
    Set rsEntry = Nothing
    Set frm = Nothing
End Sub

'Private Sub ShowBom(ByVal bomid)
'    Dim frm As New BomView
'    frm.Show
'    Set frm = Nothing
'    Exit Sub
'
'    Dim objBill As Object
'    Dim dlg As Object
'    Dim LocalCnStr As String
'    Dim sSubID As String
'    Dim sSubName As String
'    Dim lModel As Long
'    Dim lModelDetail As Long
'
'    On Error GoTo ex
'
'    Set objBill = CreateObject("K3BillsEx.Bills")
'    Set dlg = CreateObject("CSystemDlg.Sys")
'    LocalCnStr = dlg.LocalCnn
'    Set dlg = Nothing
'
'    sSubName = "Plan"
'
'    With objBill
'        .LocalCnn = LocalCnStr
'        .SystemName = sSubName
'        .SetOpt UserId, UserName
'        If Not .SaveVect(1).Lookup("sDsn") Then .SaveVect(1)("sDsn") = strConn
'        .ListRecordset = SetBillRec(bomid, 50)
'        .ListRSFieldVect = SetBillVect
'        .NewBillTransType = "50"
'
'        .Show 900, 2
'    End With
'
'    Set objBill = Nothing
'    Exit Sub
'ex:
'    MsgBox Err.Description
'    Set objBill = Nothing
'End Sub
'
'Private Function SetBillRec(ByVal aInterID As Long, ByVal aTranType As Long) As ADODB.Recordset
'    Dim tRec As ADODB.Recordset
'    Set tRec = New ADODB.Recordset
'    tRec.Fields.Append "FInterID", adInteger
'    tRec.Fields.Append "FTranType", adInteger
'    tRec.Open
'    tRec.AddNew
'    tRec!FInterID = aInterID
'    tRec!FTranType = aTranType
'    tRec.Update
'    Set SetBillRec = tRec
'End Function
'
'Private Function SetBillVect() As KFO.Vector
'    Dim tVect As KFO.Vector
'    Set tVect = New KFO.Vector
'    Dim tDict As KFO.Dictionary
'    Set tDict = New KFO.Dictionary
'    tDict("FColName") = "FInterID"
'    tDict("FISPrimary") = 1
'    tVect.Add tDict
'
'    Set tDict = New KFO.Dictionary
'    tDict("FColName") = "FTranType"
'    tDict("FISPrimary") = 3
'    tVect.Add tDict
'
'    Set tDict = Nothing
'    Set SetBillVect = tVect
'End Function

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
'    If lProc = 2576 Then
'        GetK3Conn = "ConnectString={Persist Security Info=True;Provider=SQLOLEDB.1;User ID=sa;Password=123456;Data Source=10.211.55.21;Initial Catalog=AIS20091105191238};UserName=administrator;UserID=16394;DBMS Name=Microsoft SQL Server;DBMS Version=2000/2005;SubID=xs;AcctType=gy;Setuptype=Industry;Language=chs;IP=10.211.55.19;K3Version=KUE;MachineName=SNOKYEA68F;UUID=8EFACFB9-192C-4EDD-9AFE-610EEF9705E8"
'    End If
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
Private Function GetData(ByVal strSql As String) As Object ' ADODB.Recordset
    Dim cn As Object 'New ADODB.Connection
    Dim rs As Object 'New ADODB.Recordset
    
    Set cn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    cn.ConnectionString = strConn
    cn.CursorLocation = 3
    cn.Open
    rs.Open strSql, cn, 2, 2
    Set GetData = rs
    Set rs = Nothing
    Set cn = Nothing
End Function
