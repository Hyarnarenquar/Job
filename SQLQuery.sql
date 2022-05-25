use WheelSheet

DECLARE @Xml xml;
SELECT @Xml = (SELECT * 
FROM OPENROWSET (BULK 'D:\Data.xml', SINGLE_BLOB) as [xml]);

INSERT INTO WagonSheet
select t.c.value('./TrainNumber[1]', 'int')                          as TrainNumber
		  ,t.c.value('./TrainIndexCombined[1]', 'nvarchar(200)')		    as TrainIndexCombined
		  ,t.c.value('./ToStationName[1]', 'nvarchar(200)')	        as ToStationName
		  ,t.c.value('./FromStationName[1]', 'nvarchar(200)')                  as FromStationName
		  ,t.c.value('./LastStationName[1]', 'nvarchar(200)')	as LastStationName
	from @Xml.nodes('Root/row') T(c)


INSERT INTO WagonList
select t.c.value('./PositionInTrain[1]', 'int')	as PositionInTrain
		  ,t.c.value('./CarNumber[1]', 'int')	as CarNumber
		  ,t.c.value('./InvoiceNum[1]', 'nvarchar(200)')	as InvoiceNum
		  ,t.c.value('./LastOperationName[1]', 'nvarchar(300)')	as LastOperationName
		  ,t.c.value('./WhenLastOperation[1]', 'nvarchar(datetime)')	as WhenLastOperation
		  ,t.c.value('./FreightEtsngName[1]', 'nvarchar(200)')	as FreightEtsngName
		  ,t.c.value('./FreightTotalWeightKg[1]', 'nvarchar(int)')	as FreightTotalWeightKg
	from @Xml.nodes('Root/row') T(c)

