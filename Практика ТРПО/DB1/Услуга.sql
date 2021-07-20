CREATE TABLE [dbo].[Услуга]
(
	[ID_Услуги] INT NOT NULL PRIMARY KEY IDENTITY,
	[Наименование] NCHAR(15) NOT NULL,
	[Стоимость] money NOT NULL,
	[Время предоставления] time NULL
)
