CREATE TABLE [dbo].[Бронь]
(
	[ID_Брони] INT NOT NULL PRIMARY KEY IDENTITY, 
    [ID_Администратора] int NOT NULL
	CONSTRAINT FK_AB FOREIGN KEY (ID_Администратора) REFERENCES Администратор(ID_Администратора) ,
	[ID_Клиент] int NOT NULL
	CONSTRAINT FK_CB FOREIGN KEY (ID_Клиент) REFERENCES Клиент(ID_Клиент),
	[ID_Номер] int NOT NULL
	CONSTRAINT FK_NB FOREIGN KEY (ID_Номер) REFERENCES Номер(ID_Номера),


	
)
