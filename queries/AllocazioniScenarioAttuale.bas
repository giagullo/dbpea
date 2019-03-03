PARAMETERS scenario Text ( 255 );
SELECT Progetto.codSIPROS, [Progetto].[descProgetto] & ' - ' & [Task].[descTask] AS [desc], Task.codSIPROS AS codSIPROSTask, Risorsa.Nome, Allocazioni.mese, [pct]/100 AS PctUso
FROM Progetto INNER JOIN (Risorsa INNER JOIN (Task INNER JOIN (PianoTask INNER JOIN Allocazioni ON (PianoTask.scenario = Allocazioni.scenario) AND (PianoTask.IDTask = Allocazioni.IDTask)) ON Task.ID = PianoTask.IDTask) ON Risorsa.ID = Allocazioni.IDRisorsa) ON Progetto.codSIPROS = Task.codSIPROSProg
WHERE (((PianoTask.scenario)=[scenario]))
ORDER BY Progetto.codSIPROS;

