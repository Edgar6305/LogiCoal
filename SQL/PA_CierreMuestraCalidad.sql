USE [TRACER]
GO
/****** Object:  StoredProcedure [dbo].[PA_CierreMuestraCalidad]    Script Date: 07/07/2022 18:18:27 p.m. ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
ALTER PROCEDURE [dbo].[PA_CierreMuestraCalidad] 
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

DECLARE @H1 as Float =0
DECLARE @S1 as Float =0
DECLARE @BTU1 as Float =0
DECLARE @CNZ1 as Float =0
DECLARE @CF1 as Float =0
DECLARE @V1 as Float =0
DECLARE @Dry_S1 as Float =0
DECLARE @Dry_BTU1 as Float =0
DECLARE @Dry_CNZ1 as Float =0
DECLARE @Dry_CF1 as Float =0
DECLARE @Dry_V1 as Float =0

DECLARE @T_H as Float =0
DECLARE @T_S as Float =0
DECLARE @T_BTU as Float =0
DECLARE @T_CNZ as Float =0
DECLARE @T_CF as Float =0
DECLARE @T_V as Float =0
DECLARE @T_Dry_S as Float =0
DECLARE @T_Dry_BTU as Float =0
DECLARE @T_Dry_CNZ as Float =0
DECLARE @T_Dry_CF as Float =0
DECLARE @T_Dry_V as Float =0

DECLARE @IdPila as int =0
DECLARE @UltimoIdMuestra int =0
DECLARE @Cantidad float =0
DECLARE @CantidadMuestra float =0
DECLARE @Origen Varchar(2) 
DECLARE @Numero int
DECLARE @FechaOrigen Datetime

DECLARE @SaldoPila float =0
DECLARE @SaldoInicial float =0

BEGIN
	SELECT @CantidadMuestra=Cantidad, @Origen=TransaccionOrigen, @Numero=Numero, @H=H,@S=S,@BTU=BTU,@CNZ=CNZ,@CF=CF,@V=V,@Dry_S=Dry_S,@Dry_BTU=Dry_BTU,@Dry_CNZ=Dry_CNZ,@Dry_CF=Dry_CF,@Dry_V=Dry_V 
	FROM Calidad 
	WHERE IdMuestra=@IdMuestra

	SELECT @IdPila=IdPila FROM PilasDetalle WHERE IdMuestra=@IdMuestra AND Transaccion=@Origen AND Numero=@Numero

	--Se obtiene la cantidad Inicial de la PILADETALLE con la transaccion Origen
	SELECT @FechaOrigen=FechaCierre From Lotes Where IdLote=@Numero
	SELECT @SaldoInicial=dbo.FS_CantidadPilas(@IdPila, @FechaOrigen)  

	SET @SaldoPila=@SaldoInicial+@CantidadMuestra

	BEGIN TRY  
		BEGIN TRAN
			--DECLARE PilasCalidad CURSOR FOR  
			--SELECT Cantidad, H, S, BTU, CNZ, CF, V, Dry_S, Dry_BTU, Dry_CNZ, Dry_CF, Dry_V FROM PilasDetalle WHERE IdPila=@IdPila AND Estado='AC' /*==> Esto es La ULTIMA Procesada*/
			--OPEN PilasCalidad
			--FETCH NEXT FROM PilasCalidad
			--INTO @Cantidad, @H1, @S1, @BTU1, @CNZ1, @CF1, @V1, @Dry_S1, @Dry_BTU1, @Dry_CNZ1, @Dry_CF1, @Dry_V1 
			--WHILE @@FETCH_STATUS = 0  
			--	BEGIN  
			--		SET @SaldoPila=@SaldoPila+@Cantidad
			--		SET @T_H = @T_H+ (@H1*@Cantidad)
			--		SET @T_S = @T_S+ (@S1*@Cantidad) 
			--		SET @T_BTU = @T_BTU+ (@BTU1*@Cantidad)
			--		SET @T_CNZ = @T_CNZ+ (@CNZ1*@Cantidad)
			--		SET @T_CF = @T_CF+ (@CF1*@Cantidad)
			--		SET @T_V = @T_V + (@V1*@Cantidad)
			--		SET @T_Dry_S = @T_Dry_S+ (@Dry_S1*@Cantidad)
			--		SET @T_Dry_BTU = @T_Dry_BTU + (@Dry_BTU1*@Cantidad)
			--		SET @T_Dry_CNZ = @T_Dry_CNZ+ (@Dry_CNZ1*@Cantidad)
			--		SET @T_Dry_CF = @T_Dry_CF+ (@Dry_CF1*@Cantidad)
			--		SET @T_Dry_V = @T_Dry_V + (@Dry_V1*@Cantidad)

			--		FETCH NEXT FROM PilasCalidad 
			--		INTO @Cantidad, @H1, @S1, @BTU1, @CNZ1, @CF1, @V1, @Dry_S1, @Dry_BTU1, @Dry_CNZ1, @Dry_CF1, @Dry_V1 
			--	END

			SELECT   TOP 1 @H1=H,@S1=S,@BTU1=BTU,@CNZ1=CNZ,@CF1=CF,@V1=V,@Dry_S1=Dry_S,@Dry_BTU1=Dry_BTU,@Dry_CNZ1=Dry_CNZ,@Dry_CF1=Dry_CF,@Dry_V1=Dry_V 
			FROM     PilasDetalle
			WHERE    IdPila=@IdPila AND Estado='AC'
			ORDER BY IdPilaDetalle DESC

			SET @T_H       = (@H1*@SaldoInicial)+(@H*@CantidadMuestra) / @SaldoPila
			SET @T_S       = (@S1*@SaldoInicial)+(@S*@CantidadMuestra) / @SaldoPila
			SET @T_BTU     = (@BTU1*@SaldoInicial)+(@BTU*@CantidadMuestra) / @SaldoPila
			SET @T_CNZ     = (@CNZ1*@SaldoInicial)+(@CNZ*@CantidadMuestra) / @SaldoPila
			SET @T_CF      = (@CF1*@SaldoInicial)+(@CF*@CantidadMuestra) / @SaldoPila
			SET @T_V       = (@V1*@SaldoInicial)+(@V*@CantidadMuestra) / @SaldoPila
			SET @T_Dry_S   = (@Dry_S1*@SaldoInicial)+(@Dry_S*@CantidadMuestra) / @SaldoPila
			SET @T_Dry_BTU = (@Dry_BTU1*@SaldoInicial)+(@Dry_BTU*@CantidadMuestra) / @SaldoPila
			SET @T_Dry_CNZ = (@Dry_CNZ1*@SaldoInicial)+(@Dry_CNZ*@CantidadMuestra) / @SaldoPila
			SET @T_Dry_CF  = (@Dry_CF1*@SaldoInicial)+(@Dry_CF*@CantidadMuestra) / @SaldoPila
			SET @T_Dry_V   = (@Dry_V1*@SaldoInicial)+(@Dry_V*@CantidadMuestra) / @SaldoPila

			-- Actualiza PilasDetalle y Calidad con la Calidad del IdMuestra
			UPDATE PilasDetalle 
			SET H=@T_H,S=@T_S,BTU=@T_BTU,CNZ=@T_CNZ,CF=@T_CF,V=@T_V,Dry_S=@T_Dry_S,Dry_BTU=@T_Dry_BTU,Dry_CNZ=@T_Dry_CNZ,Dry_CF=@T_Dry_CF,Dry_V=@T_Dry_V, Estado='AC' 
			WHERE IdMuestra=@IdMuestra AND Transaccion=@Origen AND Numero=@Numero

			UPDATE Calidad 
			SET Estado='AC', UsuarioEntrega=@Usuario, FechaEntrega=Getdate() 
			WHERE IdMuestra=@IdMuestra

			--CLOSE PilasCalidad 
			--DEALLOCATE PilasCalidad

		COMMIT TRAN
		SELECT 'OK'
	END TRY  		
	BEGIN CATCH  
		ROLLBACK TRAN
		SELECT ERROR_MESSAGE()
	END CATCH 
END
