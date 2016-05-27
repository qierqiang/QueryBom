-- =============================================
-- Author:      lovethesea@qq.com
-- Create date: 2016-05-26
-- Description: BOM��(ICBom)���FUseStatus�ı�ʱ�������۶�����¼�е�'����BOM'��ֵ
-- =============================================
CREATE TRIGGER tg_notifyhasbom 
   ON ICBOM 
   FOR UPDATE
AS 
BEGIN
    SET NOCOUNT ON;

	--BOM����ʱ�����޸�ʹ��״̬������ֻ��ʹ��״̬����ʱ����Ҫִ��
	IF ((SELECT FUseStatus FROM DELETED)<>(SELECT FUseStatus FROM INSERTED))
	BEGIN
		DECLARE @itemid INT
		DECLARE @bomid INT
		SELECT @itemid=FItemID FROM INSERTED
		SELECT @bomid=FInterID FROM ICBOM WHERE FItemID=@itemID AND FUseStatus=1072
	    IF (@bomid IS NULL)
		BEGIN
			UPDATE SEOrderEntry SET FEntrySelfS0175=29814 WHERE FItemID=@itemid --��
		END 
		ELSE
		BEGIN
			UPDATE SEOrderEntry SET FEntrySelfS0175=29813 WHERE FItemID=@itemid --��
		END
	END

    SET NOCOUNT OFF;
END