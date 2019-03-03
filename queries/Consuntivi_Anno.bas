PARAMETERS anno Long;
SELECT Risorsa.Nome, Progetto.codSIPROS, Task.codSIPROS AS codSIPROSTask, [Progetto].[descProgetto] & ' - ' & [Task].[descTask] AS [desc], Utilizzo.mese, [pct]/100 AS PctUso
FROM Risorsa INNER JOIN ((Progetto INNER JOIN Task ON Progetto.[codSIPROS] = Task.[codSIPROSProg]) INNER JOIN Utilizzo ON Task.[ID] = Utilizzo.[IDTask]) ON Risorsa.[ID] = Utilizzo.[IDRisorsa]
WHERE (((Utilizzo.mese)>=[anno]*100));

