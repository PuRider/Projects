CREATE TABLE [dbo].[Администратор]
(
	[ID_Администратора] INT NOT NULL PRIMARY KEY IDENTITY, 
    [Фамилия] NCHAR(15) NOT NULL, 
    [Имя] NCHAR(15) NOT NULL,
    [Отчество] NCHAR(15) NOT NULL,
    [Дата Рождения] DATE NOT NULL, 
    [Контактные данные] NCHAR(10) NULL

)
