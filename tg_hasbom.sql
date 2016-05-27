-- =============================================
-- Author:		lovethesea@qq.com
-- Create date: 2016-05-25
-- Description: ���۶�����¼��(SEOrderEntry)��Ӽ�¼����FItemID�ֶθı�ʱ����'����BOM'��ֵ
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
		UPDATE SEOrderEntry SET FEntrySelfS0175=29814 WHERE FDetailID=@detailid --��
	END	
	ELSE
	BEGIN
		UPDATE SEOrderEntry SET FEntrySelfS0175=29813 WHERE FDetailID=@detailid --��
	END

	SET NOCOUNT OFF;
END