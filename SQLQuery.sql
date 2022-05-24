use WheelSheet

DECLARE @Xml xml;
SELECT @Xml = (SELECT * 
FROM OPENROWSET (BULK 'D:\Data.xml', SINGLE_BLOB) as [xml]);

INSERT INTO WagonSheet
select t.c.value('./TrainNumber[1]', 'nvarchar(150)')                          as TrainNumber
		  ,t.c.value('./TrainIndexCombined[1]', 'nvarchar(150)')		    as TrainIndexCombined
		  ,t.c.value('./ToStationName[1]', 'nvarchar(150)')	        as ToStationName
		  ,t.c.value('./FromStationName[1]', 'nvarchar(150)')                  as FromStationName
		  ,t.c.value('./LastStationName[1]', 'nvarchar(150)')	as LastStationName
	from @Xml.nodes('Root/row') T(c)


INSERT INTO WagonList
select t.c.value('./PositionInTrain[1]', 'nvarchar(150)')	as PositionInTrain
		  ,t.c.value('./CarNumber[1]', 'nvarchar(150)')	as CarNumber
		  ,t.c.value('./InvoiceNum[1]', 'nvarchar(150)')	as InvoiceNum
		  ,t.c.value('./LastOperationName[1]', 'nvarchar(150)')	as LastOperationName
		  ,t.c.value('./WhenLastOperation[1]', 'nvarchar(150)')	as WhenLastOperation
		  ,t.c.value('./FreightEtsngName[1]', 'nvarchar(150)')	as FreightEtsngName
		  ,t.c.value('./FreightTotalWeightKg[1]', 'nvarchar(150)')	as FreightTotalWeightKg
	from @Xml.nodes('Root/row') T(c)

