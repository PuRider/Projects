CREATE TABLE [dbo].[Клиент]
(
	[ID_Клиент] INT NOT NULL PRIMARY KEY IDENTITY, 
    [Фамилия] NCHAR(15) NOT NULL,
	[Имя] NCHAR(15) NOT NULL,
	[Отчество] NCHAR(15) NOT NULL, 
    [Контактные данные] NCHAR(15) NULL, 
    [Паспортные данные] NCHAR(15) NOT NULL
)
