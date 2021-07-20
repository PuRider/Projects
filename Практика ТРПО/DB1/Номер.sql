CREATE TABLE [dbo].[Номер]
(
	[ID_Номера] INT NOT NULL PRIMARY KEY IDENTITY,
	[Этаж] INT NOT NULL, 
    [Количество комнат] INT NOT NULL, 
    [Класс] NCHAR(10) NOT NULL,
	[Стоимость] MONEY NOT null, 
    [Примечание] TEXT NULL,
)
