USE [TRACER]
GO
/****** Object:  StoredProcedure [dbo].[PA_CierreOrdenTrituracion]    Script Date: 11/01/2022 10:22:04 a.m. ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[PA_CierreMuestraCalidad] 
@IdMuestra int,
@Usuario Varchar(10)

AS
SET NOCOUNT ON;
DECLARE @H as Float
DECLARE @S as Float
DECLARE @BTU as Float
DECLARE @CNZ as Float
DECLARE @CF as Float
DECLARE @V as Float
DECLARE @Dry_S as Float
DECLARE @Dry_BTU as Float
DECLARE @Dry_CNZ as Float
DECLARE @Dry_CF as Float
DECLARE @Dry_V as Float

DECLARE @H1 as Float
DECLARE @S1 as Float
DECLARE @BTU1 as Float
DECLARE @CNZ1 as Float
DECLARE @CF1 as Float
DECLARE @V1 as Float
DECLARE @Dry_S1 as Float
DECLARE @Dry_BTU1 as Float
DECLARE @Dry_CNZ1 as Float
DECLARE @Dry_CF1 as Float
DECLARE @Dry_V1 as Float

DECLARE @T_H as Float
DECLARE @T_S as Float
DECLARE @T_BTU as Float
DECLARE @T_CNZ as Float
DECLARE @T_CF as Float
DECLARE @T_V as Float
DECLARE @T_Dry_S as Float
DECLARE @T_Dry_BTU as Float
DECLARE @T_Dry_CNZ as Float
DECLARE @T_Dry_CF as Float
DECLARE @T_Dry_V as Float

DECLARE @IdPila as int
DECLARE @UltimoIdMuestra int
DECLARE @SaldoPila float
DECLARE @Cantidad float
DECLARE @CantidadMuestra float
BEGIN
		UPDATE Calidad 
		SET Estado='AC', UsuarioEntrega=@Usuario, FechaEntrega=Getdate() 
		WHERE IdMuestra=@IdMuestra
		
		SELECT @CantidadMuestra=Cantidad, @H=H,@S=S,@BTU=BTU,@CNZ=CNZ,@CF=CF,@V=V,@Dry_S=Dry_S,@Dry_BTU=Dry_BTU,@Dry_CNZ=Dry_CNZ,@Dry_CF=Dry_CF,@Dry_V=Dry_V 
		FROM Calidad 
		WHERE IdMuestra=@IdMuestra

		SELECT @IdPila=IdPila FROM PilasDetalle WHERE IdMuestra=@IdMuestra

		--Se obtiene la cantidad Inicial de la PILADETALLE
		SELECT @SaldoPila=Cantidad FROM PilasDetalle WHERE IdMuestra=0

		--Revisa la secuensialidad de las Muestra Grabada en PilasDetalle estado IN Buscando que coincida con el IDMuestra Entregado
		SELECT @UltimoIdMuestra=MIN(IdMuestra) FROM PilasDetalle WHERE IdPila=@IdPila AND Estado='IN'

		IF @UltimoIdMuestra=@IdMuestra
			BEGIN TRY  
				BEGIN TRAN
					DECLARE PilasCalidad CURSOR FOR  
					SELECT Cantidad, H, S, BTU, CNZ, CF, V, Dry_S, Dry_BTU, Dry_CNZ, Dry_CF, Dry_V FROM PilasDetalle WHERE IdPila=@IdPila AND Estado='AC'
					OPEN PilasCalidad
					FETCH NEXT FROM Cursor_Pilas 
					INTO @Cantidad, @H1, @S1, @BTU1, @CNZ1, @CF1, @V1, @Dry_S1, @Dry_BTU1, @Dry_CNZ1, @Dry_CF1, @Dry_V1 
					WHILE @@FETCH_STATUS = 0  
					   BEGIN  
						  SET @SaldoPila=@SaldoPila+@Cantidad
						  SET @T_H = @T_H+ (@H1*@Cantidad)
						  SET @T_S = @T_S+ (@S1*@Cantidad) 
						  SET @T_BTU = @T_BTU+ (@BTU1*@Cantidad)
						  SET @T_CNZ = @T_CNZ+ (@CNZ1*@Cantidad)
						  SET @T_CF = @T_CF+ (@CF1*@Cantidad)
						  SET @T_V = @T_V + (@V1*@Cantidad)
						  SET @T_Dry_S = @T_Dry_S+ (@Dry_S1*@Cantidad)
						  SET @T_Dry_BTU = @T_Dry_BTU + (@Dry_BTU1*@Cantidad)
						  SET @T_Dry_CNZ = @T_Dry_CNZ+ (@Dry_CNZ1*@Cantidad)
						  SET @T_Dry_CF = @T_Dry_CF+ (@Dry_CF1*@Cantidad)
						  SET @T_Dry_V = @T_Dry_V + (@Dry_V1*@Cantidad)

						  FETCH NEXT FROM Employee_Cursor 
						  INTO @Cantidad, @H1, @S1, @BTU1, @CNZ1, @CF1, @V1, @Dry_S1, @Dry_BTU1, @Dry_CNZ1, @Dry_CF1, @Dry_V1 
					   END

					SET @SaldoPila = @SaldoPila+@CantidadMuestra
					SET @T_H       = (@T_H+ (@H*@Cantidad)) / @SaldoPila
					SET @T_S       = (@T_S+ (@S*@Cantidad)) / @SaldoPila 
					SET @T_BTU     = (@T_BTU+ (@BTU*@Cantidad)) / @SaldoPila
					SET @T_CNZ     = (@T_CNZ+ (@CNZ*@Cantidad)) / @SaldoPila
					SET @T_CF      = (@T_CF+ (@CF*@Cantidad)) / @SaldoPila
					SET @T_V       = (@T_V + (@V*@Cantidad)) / @SaldoPila
					SET @T_Dry_S   = (@T_Dry_S+ (@Dry_S*@Cantidad)) / @SaldoPila
					SET @T_Dry_BTU = (@T_Dry_BTU + (@Dry_BTU*@Cantidad)) / @SaldoPila
					SET @T_Dry_CNZ = (@T_Dry_CNZ+ (@Dry_CNZ*@Cantidad)) / @SaldoPila
					SET @T_Dry_CF  = (@T_Dry_CF+ (@Dry_CF*@Cantidad)) / @SaldoPila
					SET @T_Dry_V   = (@T_Dry_V + (@Dry_V*@Cantidad)) / @SaldoPila

					-- actualiza PilasDetalle con la Calidad del IdMuestra
					UPDATE PilasDetalle 
					SET H=@T_H,S=@T_S,BTU=@T_BTU,CNZ=@T_CNZ,CF=@T_CF,V=@T_V,Dry_S=@T_Dry_S,Dry_BTU=@T_Dry_BTU,Dry_CNZ=@T_Dry_CNZ,Dry_CF=@T_Dry_CF,Dry_V=@T_Dry_V 
					WHERE IdMuestra=@IdMuestra

					CLOSE Cursor_Pilas 
					DEALLOCATE Cursor_Pilas

				COMMIT TRAN
				SELECT 'OK'
			END TRY  		
			BEGIN CATCH  
				ROLLBACK TRAN
				SELECT ERROR_MESSAGE()
			END CATCH 
		ELSE 
		BEGIN
			SELECT 'El IdMuestra No ' + @IdMuestra + ' NO es el consecutivo en la Pila, por favor revise'
		END
END
