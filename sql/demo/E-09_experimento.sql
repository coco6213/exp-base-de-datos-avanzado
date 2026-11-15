-- E-09 | Resultados del experimento (10 corridas en caliente por consulta)
USE LabExperimento;
SELECT Escenario, ConsultaID, CAST(AVG(TiempoMs) AS DECIMAL(10,1)) TiempoProm_ms, CAST(STDEV(TiempoMs) AS DECIMAL(10,1)) DesvEst_ms,
       AVG(Lecturas) LecturasProm, MAX(Filas) Filas
FROM dbo.Resultado WHERE TipoCorrida='CALIENTE' GROUP BY Escenario, ConsultaID ORDER BY ConsultaID, Escenario;
SELECT ConsultaID, COUNT(DISTINCT ChecksumRes) ResultadosDistintos FROM dbo.Resultado GROUP BY ConsultaID;
