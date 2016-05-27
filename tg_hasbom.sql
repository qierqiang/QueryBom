-- =============================================
-- Author:		lovethesea@qq.com
-- Create date: 2016-05-25
-- Description: 销售订单分录表(SEOrderEntry)添加记录或表的FItemID字段改变时更新'有无BOM'的值
-- =============================================
CREATE TRIGGER tg_hasbom 
   ON dbo.SEOrderEntry 
   AFTER INSERT,UPDATE
AS 
BEGIN
	SET NOCOUNT ON;

	DECLARE @detailid INT
	DECLARE @itemid INT
	DECLARE @bomid INT
	SELECT @detailid=FDetailID, @itemid=FItemID FROM INSERTED
	SELECT @bomid=FInterID FROM ICBOM WHERE FItemID=@itemID AND FUseStatus=1072
	IF (@bomid IS NULL)
	BEGIN
		UPDATE SEOrderEntry SET FEntrySelfS0175=29814 WHERE FDetailID=@detailid --无
	END	
	ELSE
	BEGIN
		UPDATE SEOrderEntry SET FEntrySelfS0175=29813 WHERE FDetailID=@detailid --有
	END

	SET NOCOUNT OFF;
END