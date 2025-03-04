/*
   jueves, 23 de abril de 202011:09:24 a.m.
   Usuario: 
   Servidor: JUANBERMUDEZ-PC\SQL2014
   Base de datos: CROSCHECK
   Aplicación: 
*/

/* Para evitar posibles problemas de pérdida de datos, debe revisar este script detalladamente antes de ejecutarlo fuera del contexto del diseñador de base de datos.*/
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
ALTER TABLE dbo.Schedule ADD
	AutoClass nvarchar(50) NULL
GO
ALTER TABLE dbo.Schedule ADD CONSTRAINT
	DF_Schedule_AutoClass DEFAULT 0 FOR AutoClass
GO
ALTER TABLE dbo.Schedule SET (LOCK_ESCALATION = TABLE)
GO
COMMIT
