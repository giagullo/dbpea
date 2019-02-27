INSERT INTO Allocazioni ( IDRisorsa, IDTask, Scenario, Mese, Pct )
SELECT Risorsa.ID AS IDRisorsa, Task.ID AS IDTask, "4Q2018" AS Scenario, 201800+[NumMese] AS Mese, [Pct]*100 AS Allocato
FROM (((wk_Alloc4q2018 INNER JOIN wk_RisCogn ON wk_Alloc4q2018.Risorsa = wk_RisCogn.Cognome) INNER JOIN Risorsa ON wk_RisCogn.Risorsa = Risorsa.Nome) INNER JOIN wk_TaskCorr ON wk_Alloc4q2018.Iniziativa = wk_TaskCorr.NomeTaskInAllocXls) INNER JOIN Task ON wk_TaskCorr.codSiprosTask = Task.codSIPROS;

