/*
   domingo, 11 de marzo de 201210:34:43 a.m.
   Usuario: 
   Servidor: JUAN\SQL2005
   Base de datos: EasyWayBiomtrics
   Aplicación: 
*/

/* Para evitar posibles problemas de pérdida de datos, debe revisar esta secuencia de comandos detalladamente antes de ejecutarla fuera del contexto del diseñador de base de datos.*/
BEGIN TRANSACTION
SET QUOTED_IDENTIFIER ON
SET ARITHABORT ON
SET NUMERIC_ROUNDABORT OFF
SET CONCAT_NULL_YIELDS_NULL ON
SET ANSI_NULLS ON
SET ANSI_PADDING ON
SET ANSI_WARNINGS ON
COMMIT
BEGIN TRANSACTION
GO
CREATE TABLE dbo.T_Checkinout
	(
	Logid numeric(18, 0) NOT NULL IDENTITY (1, 1),
	Userid nvarchar(MAX) NULL,
	CheckTime datetime NULL,
	CheckType nvarchar(MAX) NULL,
	Sensorid nvarchar(MAX) NULL,
	Checked bit NULL,
	WorkType numeric(18, 0) NULL,
	AttFlag numeric(18, 0) NULL
	)  ON [PRIMARY]
	 TEXTIMAGE_ON [PRIMARY]
GO
COMMIT
