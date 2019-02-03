** Query qry_ScenarioDiPiano
PARAMETERS scenario Text ( 255 );
SELECT Progetto.codSIPROS, Progetto.descProgetto, Task.codSIPROS AS codTask, Task.descTask, PianoTask.dtInizio, PianoTask.dtFine, Task.codPortfolio
FROM (Progetto INNER JOIN Task ON Progetto.codSIPROS = Task.codSIPROSProg) INNER JOIN PianoTask ON Task.ID = PianoTask.IDTask
WHERE (((PianoTask.scenario)=[scenario]))
ORDER BY Progetto.codSIPROS, Task.codSIPROS;
