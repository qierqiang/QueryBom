-- =============================================
-- Author:      lovethesea@qq.com
-- Create date: 2016-05-26
-- Description: BOM表(ICBom)表的FUseStatus改变时更新销售订单分录中的'有无BOM'的值
-- =============================================
CREATE TRIGGER tg_notifyhasbom 
   ON ICBOM 
   FOR UPDATE
AS 
BEGIN
    SET NOCOUNT ON;

	--BOM新增时不能修改使用状态，所以只在使用状态更新时才需要执行
	IF ((SELECT FUseStatus FROM DELETED)<>(SELECT FUseStatus FROM INSERTED))
	BEGIN
		DECLARE @itemid INT
		DECLARE @bomid INT
		SELECT @itemid=FItemID FROM INSERTED
		SELECT @bomid=FInterID FROM ICBOM WHERE FItemID=@itemID AND FUseStatus=1072
	    IF (@bomid IS NULL)
		BEGIN
			UPDATE SEOrderEntry SET FEntrySelfS0175=29814 WHERE FItemID=@itemid --无
		END 
		ELSE
		BEGIN
			UPDATE SEOrderEntry SET FEntrySelfS0175=29813 WHERE FItemID=@itemid --有
		END
	END

    SET NOCOUNT OFF;
END