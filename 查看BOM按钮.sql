--(1) 增加元数据按钮

--FToolID = 70001我开发的插件按钮从70001开始，避免冲突
DELETE t_MenuToolBar WHERE FToolID = 70001
INSERT INTO t_MenuToolBar (FToolID,FName,FCaption,FCaption_CHT,FCaption_EN,FImageName,FToolTip,FToolTip_CHT,FToolTip_EN,FControlType,FVisible,FEnable,FChecked,FShortCut,FCBList,FCBList_CHT,FCBList_EN,FCBStyle,FCBWidth,FIndex,FToolCaption,FToolCaption_CHT,FToolCaption_EN)
VALUES (70001,'btnQueryBom','查看BOM','查看BOM','Show BOM','43','查看BOM','查看BOM','Show BOM',0,1,1,0,0,'','','',0,0,0,'查看BOM','查看BOM','Show BOM')


--(2) 把注册的按钮添加到工具栏里

--FBandID=53是固定值，表示按钮放置的容器ID
--销售订单列表界面FID: 32, FMenuID: 100
DELETE t_BandToolMapping WHERE FBandID=48 AND FToolID = 70001 AND FID = 100 
INSERT INTO t_BandToolMapping (FID,FBandID,FToolID,FSubBandID,FIndex,FComName,FBeginGroup)
VALUES (100,48,70001,0,62,'|SEOrderBOMQuery.QueryBom',1) 

--(3) 在销售订单序时薄显示按钮


--在销售订单序时薄显示按钮(如果里面有"|V",则只能在后面加菜单项) 
UPDATE ICListTemplate 
SET FLogicStr=FLogicStr + Case When Right(FLogicStr,1)='|' then 'V:btnQueryBom' else '|V:btnQueryBom' end 
WHERE FID = 32 AND FLogicStr NOT LIKE '%btnQueryBom%' 