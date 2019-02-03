** Query qry_TabelloneAllocazioniXScenario
PARAMETERS scenario Text ( 255 );
SELECT Progetto.codSIPROS AS [Codice progetto], Task.codSIPROS AS [Codice task], Progetto.descProgetto, Task.descTask, PianoTask.dtInizio, PianoTask.dtFine, Risorsa.Nome, Allocazioni.mese, Allocazioni.pct, Task.codPortfolio
FROM Progetto INNER JOIN (Risorsa INNER JOIN (Task INNER JOIN (PianoTask INNER JOIN Allocazioni ON (PianoTask.scenario = Allocazioni.scenario) AND (PianoTask.IDTask = Allocazioni.IDTask)) ON Task.ID = PianoTask.IDTask) ON Risorsa.ID = Allocazioni.IDRisorsa) ON Progetto.codSIPROS = Task.codSIPROSProg
WHERE (((Allocazioni.scenario)=[scenario]))
ORDER BY Progetto.codSIPROS, Task.codSIPROS, Risorsa.Nome, Allocazioni.mese;
