DECLARE @sql char(150)
DECLARE @Driver sysname
DECLARE @Cadena sysname
DECLARE @Sentencia sysname

SET @Driver= ''''+'Microsoft.Jet.OLEDB.4.0'+''''
SET @Cadena=''''+'Excel 8.0; Database=\\F18SRV04\F18TMP\GARANTIAS.xls'+''''
SET @Sentencia=''''+'SELECT * FROM [GARANTIAS$B1:B2]'+''''

SET @sql='SELECT * FROM OPENROWSET('+@Driver+', '+@Cadena+', '+@Sentencia+')'


print @sql

execute (@sql)