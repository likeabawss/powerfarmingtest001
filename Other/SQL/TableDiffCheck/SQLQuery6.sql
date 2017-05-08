INSERT INTO [dbo].[TableDiffCheckHeader] ([XMLFile],[StartDateTime],[EndDateTime],[Run])
	VALUES('test',getdate(),NULL,IsnUll((SELECT Max([Run])+1 FROM [dbo].[TableDiffCheckHeader] where XMLFile = 'test' group by XMLFile),0))


	--SELECT top 1 [Run]+1 FROM [dbo].[TableDiffCheckHeader] where XMLFile = 'test' order by Run desc