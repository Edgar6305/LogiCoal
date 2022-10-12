Select * From FT_DB_RecepcionHora('02/05/2022 06:00:00', '02/05/2022 18:00:00', 'LT' ) 

--Numero de Volquetas Hora Recepcion
Set Dateformat DMY
DECLARE @FI2 as Datetime='02/05/2022 06:00:00'
DECLARE @FF2 as Datetime='02/05/2022 18:00:00'
DECLARE @TT AS varchar(2)='TR'  -->Transaccion a Monitorear

SELECT   Count(*) AS Cantidad,
(SELECT  Count(*) FROM Bascula WHERE (TransaccionOrigen = @TT) AND (Estado = 'AC') AND (IdMaterial = 1) AND (FechaLleno>=@FI2 AND FechaLleno<=@FF2) AND DATEPART(hh,FechaLleno)=6) AS H6,
(SELECT  Count(*) FROM Bascula WHERE (TransaccionOrigen = @TT) AND (Estado = 'AC') AND (IdMaterial = 1) AND (FechaLleno>=@FI2 AND FechaLleno<=@FF2) AND DATEPART(hh,FechaLleno)=7) AS H7,
(SELECT  Count(*) FROM Bascula WHERE (TransaccionOrigen = @TT) AND (Estado = 'AC') AND (IdMaterial = 1) AND (FechaLleno>=@FI2 AND FechaLleno<=@FF2) AND DATEPART(hh,FechaLleno)=8) AS H8,
(SELECT  Count(*) FROM Bascula WHERE (TransaccionOrigen = @TT) AND (Estado = 'AC') AND (IdMaterial = 1) AND (FechaLleno>=@FI2 AND FechaLleno<=@FF2) AND DATEPART(hh,FechaLleno)=9) AS H9,
(SELECT  Count(*) FROM Bascula WHERE (TransaccionOrigen = @TT) AND (Estado = 'AC') AND (IdMaterial = 1) AND (FechaLleno>=@FI2 AND FechaLleno<=@FF2) AND DATEPART(hh,FechaLleno)=10) AS H10,
(SELECT  Count(*) FROM Bascula WHERE (TransaccionOrigen = @TT) AND (Estado = 'AC') AND (IdMaterial = 1) AND (FechaLleno>=@FI2 AND FechaLleno<=@FF2) AND DATEPART(hh,FechaLleno)=11) AS H11,
(SELECT  Count(*) FROM Bascula WHERE (TransaccionOrigen = @TT) AND (Estado = 'AC') AND (IdMaterial = 1) AND (FechaLleno>=@FI2 AND FechaLleno<=@FF2) AND DATEPART(hh,FechaLleno)=12) AS H12,
(SELECT  Count(*) FROM Bascula WHERE (TransaccionOrigen = @TT) AND (Estado = 'AC') AND (IdMaterial = 1) AND (FechaLleno>=@FI2 AND FechaLleno<=@FF2) AND DATEPART(hh,FechaLleno)=13) AS H13,
(SELECT  Count(*) FROM Bascula WHERE (TransaccionOrigen = @TT) AND (Estado = 'AC') AND (IdMaterial = 1) AND (FechaLleno>=@FI2 AND FechaLleno<=@FF2) AND DATEPART(hh,FechaLleno)=14) AS H14,
(SELECT  Count(*) FROM Bascula WHERE (TransaccionOrigen = @TT) AND (Estado = 'AC') AND (IdMaterial = 1) AND (FechaLleno>=@FI2 AND FechaLleno<=@FF2) AND DATEPART(hh,FechaLleno)=15) AS H15,
(SELECT  Count(*) FROM Bascula WHERE (TransaccionOrigen = @TT) AND (Estado = 'AC') AND (IdMaterial = 1) AND (FechaLleno>=@FI2 AND FechaLleno<=@FF2) AND DATEPART(hh,FechaLleno)=16) AS H16,
(SELECT  Count(*) FROM Bascula WHERE (TransaccionOrigen = @TT) AND (Estado = 'AC') AND (IdMaterial = 1) AND (FechaLleno>=@FI2 AND FechaLleno<=@FF2) AND DATEPART(hh,FechaLleno)=17) AS H17

FROM     Bascula
WHERE   (TransaccionOrigen = @TT) AND (Estado ='AC') AND (IdMaterial = 1) AND (FechaLleno>=@FI2 AND FechaLleno<=@FF2)
