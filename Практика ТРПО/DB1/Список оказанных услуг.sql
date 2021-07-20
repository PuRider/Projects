CREATE TABLE [dbo].[Список оказанных услуг]
(
	[ID_Списка] INT NOT NULL PRIMARY KEY IDENTITY,
	[ID_Услуги] INT NOT NULL
	CONSTRAINT FK_UB FOREIGN KEY (ID_Услуги) REFERENCES Услуга(ID_Услуги),
	[ID_Брони] INT NOT NULL,
	CONSTRAINT FK_BU FOREIGN KEY (ID_Брони) REFERENCES Бронь(ID_Брони),
)
