-- CREAMOS UN TRIGGER DDL LOGON

ALTER TRIGGER [CONNECTION_LIMIT_TRIGGER]
ON ALL SERVER
FOR LOGON
AS
SET NOCOUNT ON
--CON ESTE TRIGGER QUE EL LOGIN DEMOAPP SOLO PUEDA SER USADO DESDE NUESTRA APP
BEGIN
	DECLARE @login_time NVARCHAR(128)


	IF  (SUSER_SNAME() ='SISMA_APP' OR SUSER_SNAME()='sa' OR SUSER_SNAME() LIKE '%SQLSERVERAGENT%' ) OR ( APP_NAME()='Visual Basic' OR APP_NAME() LIKE '%Management%') 
		BEGIN
			SET @login_time = GETDATE()
		END
	ELSE
		BEGIN
			ROLLBACK
			INSERT INTO TRACER.dbo.LogOn VALUES(ERROR_MESSAGE(),SUSER_SNAME(), APP_NAME())
		END
END
GO

GO

SET QUOTED_IDENTIFIER OFF
GO

ENABLE TRIGGER [CONNECTION_LIMIT_TRIGGER] ON ALL SERVER
GO

--DISABLE TRIGGER [CONNECTION_LIMIT_TRIGGER] ON ALL SERVER
--GO

--DROP TRIGGER [CONNECTION_LIMIT_TRIGGER]
--GO
