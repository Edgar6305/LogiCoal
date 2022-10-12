UPDATE Bascula 
SET FechaVacio=TP.FechaVacio,
	FechaLleno=TP.FechaLleno
FROM Temp as TP 
WHERE Bascula.IdTiquete=TP.Tiquete


--UPDATE
--     Tabla
--SET
--     Tabla.col1 = otra_tabla.col1,
--     Tabla.col2 = otra_tabla.col2 
--FROM
--     Tabla
--INNER JOIN     
--     otra_tabla
--ON     
--     Tabla.id = otra_tabla.id 
--WHERE EXISTS(SELECT Tabla.Col1, Tabla.Col2 EXCEPT SELECT otra_tabla.Col1, otra_tabla.Col2))