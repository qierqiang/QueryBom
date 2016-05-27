--FItemID=29813, FName='有'	
--FItemID=29814, FName='无'
UPDATE SEOrderEntry SET FEntrySelfS0175=29814--默认无
UPDATE t1 SET t1.FEntrySelfS0175=29813
FROM SEOrderEntry t1
JOIN ICBOM t2 ON t2.FItemID=t1.FItemID
WHERE t2.FUseStatus=1072