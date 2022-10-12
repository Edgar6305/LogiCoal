SET DATEFORMAT DMY
SELECT        vMovimientosDS.IdTiquete, vMovimientosDS.Tercero, vMovimientosDS.FechaLleno, vMovimientosDS.PesoNeto, MovimientosDetalle.Cliente, MovimientosDetalle.FechaLlegada, 
                         MovimientosDetalle.PesoLleno - MovimientosDetalle.PesoVacio AS NetoLlegada, CASE WHEN MovimientosDetalle.Cliente IS NULL THEN DATEDIFF(hh, vMovimientosDS.FechaLleno, Getdate()) / 24 ELSE DATEDIFF(hh, 
                         vMovimientosDS.FechaLleno, MovimientosDetalle.FechaLlegada) / 24 END AS DiasRecorrido, CASE WHEN MovimientosDetalle.PesoLleno IS NULL 
                         THEN vMovimientosDS.PesoNeto ELSE vMovimientosDS.PesoNeto - (MovimientosDetalle.PesoLleno - MovimientosDetalle.PesoVacio) END AS Diferencia, Terceros.Descripcion
FROM            Terceros INNER JOIN
                         vMovimientosDS ON Terceros.Identificacion = vMovimientosDS.Tercero LEFT OUTER JOIN
                         MovimientosDetalle ON vMovimientosDS.IdTiquete = MovimientosDetalle.Tiquete
WHERE        (vMovimientosDS.FechaLleno >= '23/05/2022') AND (vMovimientosDS.Tercero = '315')



--Select Distinct vMovimientosDS.Tercero, Terceros.Descripcion
--FROM            Terceros INNER JOIN
--                         vMovimientosDS ON Terceros.Identificacion = vMovimientosDS.Tercero LEFT OUTER JOIN
--                         MovimientosDetalle ON vMovimientosDS.IdTiquete = MovimientosDetalle.Tiquete
--WHERE        (vMovimientosDS.FechaLleno >= '23/05/2022') 