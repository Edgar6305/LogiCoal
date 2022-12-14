USE [TRACER]
GO
/****** Object:  StoredProcedure [dbo].[PA_CierreOrdenTrituracion]    Script Date: 10/01/2022 03:19:56 p.m. ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[PA_CargarLotePilaDetalle] 
@IdTrituracion int

AS
SET NOCOUNT ON;
DECLARE @Lote as Int
DECLARE @Muestra as Int
DECLARE @Pila as Int
DECLARE @Por as float
DECLARE @CantidadLote as float
DECLARE @Cantidad as float

BEGIN
	BEGIN TRY  

		SELECT @Muestra=IdMuestra FROM Calidad Where TransaccionOrigen='LT' AND Numero=@Lote AND TipoMuestra=2
		SELECT @Lote=Idlote FROM TrituracionDetalle Where IdTrituracion=@IdTrituracion
		SELECT @CantidadLote=Cantidad FROM Lotes Where IdLote=@Lote
		BEGIN TRAN
			--Se recorre la order de TrituracionDetalle para revisar las diferentes pilas afectadas
			DECLARE Cursor_Pilas CURSOR FOR  
			SELECT IdPilaDestino,Porcentaje FROM TrituracionDetalle Where IdTrituracion=@IdTrituracion
			OPEN Cursor_Pilas 
			FETCH NEXT FROM Cursor_Pilas INTO @Pila,@Por
			WHILE @@FETCH_STATUS = 0  
			   BEGIN  
				  SET @Cantidad=@CantidadLote*@Por	
				  INSERT INTO PilasDetalle
				  VALUES(@Pila,'LT',@Lote,@Cantidad,@Muestra,0,0,0,0,0,0,0,0,0,0,0)

				  FETCH NEXT FROM Employee_Cursor INTO @Pila,@Por
			   END
			CLOSE Cursor_Pilas 
			DEALLOCATE Cursor_Pilas
		COMMIT TRAN
		Select 'OK'
	END TRY  		
	BEGIN CATCH  
		ROLLBACK TRAN
		Select ERROR_MESSAGE()
	END CATCH 
END
