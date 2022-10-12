Set DateFormat DMY 
SELECT    Trituracion.IdTrituracion, Trituradoras.Descripcion, Trituracion.FechaInicio, Trituracion.FechaCierre, PilasFisicas.Descripcion AS DesPilaOrigen, 
		  DATEDIFF(HOUR, Trituracion.FechaInicio, Trituracion.FechaCierre) AS HorasOrden, ROUND(Trituracion.HorasEfectivas,0) HorasEfectivas, 

		  (SELECT SUM(DATEDIFF(HOUR,FechaInicio, FechaFin)) FROM   TrituradoraParos WHERE IdTrituracion=113) AS HorasParos1,Trituracion.Cantidad AS CantidadTotal,

		  (SELECT  Top 1 TD.Porcentaje FROM  TrituracionDetalle AS TD WHERE TD.IdTrituracion =Trituracion.IdTrituracion) AS Porcentaje1,
		  (SELECT  Top 1 TD.Porcentaje FROM  TrituracionDetalle AS TD WHERE TD.IdTrituracion =Trituracion.IdTrituracion)*Trituracion.Cantidad/100 AS CantProcesada1,
		  (SELECT  Top 1 PF.Descripcion FROM PilasFisicas AS PF INNER JOIN  Pilas ON PF.IdPilaFisica = Pilas.IdPilaFisica INNER JOIN  TrituracionDetalle AS TD ON Pilas.IdPila = TD.PilaDestino WHERE   TD.IdTrituracion =Trituracion.IdTrituracion) AS TipoDes1,

		  (SELECT  Top 1 TD.Porcentaje FROM  TrituracionDetalle AS TD WHERE TD.IdTrituracion =Trituracion.IdTrituracion Order By TD.IdTrituracionDetalle DESC) AS Porcentaje2,
		  (SELECT  Top 1 TD.Porcentaje FROM  TrituracionDetalle AS TD WHERE TD.IdTrituracion =Trituracion.IdTrituracion Order By TD.IdTrituracionDetalle DESC)*Trituracion.Cantidad/100 AS CantProcesada2,
		  (SELECT  Top 1 PF.Descripcion FROM PilasFisicas AS PF INNER JOIN  Pilas ON PF.IdPilaFisica = Pilas.IdPilaFisica INNER JOIN  TrituracionDetalle AS TD ON Pilas.IdPila = TD.PilaDestino 
		   WHERE   TD.IdTrituracion =Trituracion.IdTrituracion ORDER BY TD.IdTrituracionDetalle DESC) AS TipoDes2
		  		  
FROM      PilasFisicas INNER JOIN
          Pilas ON PilasFisicas.IdPilaFisica = Pilas.IdPilaFisica INNER JOIN
          Trituracion INNER JOIN
          Trituradoras ON Trituracion.IdTrituradora = Trituradoras.IdTrituradora ON Pilas.IdPila = Trituracion.PilaOrigen
WHERE    (Trituradoras.ProduccionHora > 0) AND (Trituracion.FechaInicio >= '28/09/2022 06:00') AND (Trituracion.FechaCierre <= '29/09/2022 06:00')
ORDER BY Trituracion.IdTrituracion