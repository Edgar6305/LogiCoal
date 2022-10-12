DECLARE @Cliente as Varchar(15)='320'
DECLARE @Usuario AS Varchar(10)='sisma'

DECLARE @Fecha AS Varchar(10)
DECLARE @Placa AS Varchar(7)
DECLARE @Tiquete AS int
DECLARE @PesoLleno AS Float
DECLARE @PesoVacio AS Float
DECLARE @PesoNeto AS Float
DECLARE @HoraLlegada AS Varchar(8)
DECLARE @HoraSalida AS Varchar(8)


DECLARE @sql  nvarchar(4000)
DECLARE @Driver sysname
DECLARE @Cadena sysname
DECLARE @Sentencia sysname

SET DATEFORMAT DMY

/* EXTRAEMOS DRIVER, CADENA Y SENTENCIA DEL ARCHIVO DE MOVIMIENTO, HAY 1 POR CADA CLIENTE QUE ENVIA DATOS*/
SELECT @Driver=Driver, @Cadena=Cadena, @Sentencia=Sentencia  FROM Movimientos Where Cliente=@Cliente

SET @sql='DECLARE Movimiento CURSOR FOR' 
SET @sql=@sql + ' SELECT CONVERT(char(10), F1, 103) AS FechaLlegada, F2, F3, F8, F9, F10, CONVERT(char(8), F12, 108) AS HoraEntrada, CONVERT(char(8), F13, 108) AS HoraSalida '
SET @sql=@sql + ' FROM OPENROWSET('+@Driver+', '+@Cadena+', '+@Sentencia+')'

--Print @sql

BEGIN TRY  
	BEGIN TRAN
		EXEC sp_executesql @sql
		--DECLARE Movimiento CURSOR FOR  
		--SELECT CONVERT(char(10), F1, 103) AS FechaLlegada, F2, F3, F8, F9, F10, CONVERT(char(8), F12, 108) AS HoraEntrada, CONVERT(char(8), F13, 108) AS HoraSalida
		--FROM OPENROWSET('Microsoft.ACE.OLEDB.16.0','Excel 12.0;HDR=NO;Database=C:\FRONTIER\Movimientos\2022\GESELCA\Prueba-Bascula.xlsx','Select * From [Hoja1$]')
		OPEN Movimiento
		FETCH NEXT FROM Movimiento
		INTO  @Fecha, @Placa, @Tiquete, @PesoLleno, @PesoVacio, @PesoNeto, @HoraLlegada, @HoraSalida
		WHILE @@FETCH_STATUS = 0  
			BEGIN 
				IF @Tiquete IS NOT NULL
					BEGIN
						INSERT INTO MovimientosDetalle (Cliente, Tiquete, Placa, FechaLlegada, FechaSalida, PesoLleno, PesoVacio, Usuario, Fecha) VALUES (@Cliente, @Tiquete, REPLACE(@Placa,'-',''), CONCAT(@Fecha,' ',@HoraLlegada), CONCAT(@Fecha,' ',@HoraSalida), @PesoLleno, @PesoVacio, @Usuario, Getdate())
					END
				FETCH NEXT FROM Movimiento
				INTO  @Fecha, @Placa, @Tiquete, @PesoLleno, @PesoVacio, @PesoNeto, @HoraLlegada, @HoraSalida
			END

			CLOSE Movimiento 
			DEALLOCATE Movimiento

		COMMIT TRAN
		SELECT 'OK'
	END TRY  		
BEGIN CATCH  
	ROLLBACK TRAN
	CLOSE Movimiento 
	DEALLOCATE Movimiento
	SELECT ERROR_MESSAGE()
END CATCH 