SELECT Progetto.codSIPROS, Progetto.descProgetto, Task.codSIPROS, Task.descTask, PianoTask.IDTask, PianoTask.scenario, PianoTask.dtInizio, PianoTask.dtFine
FROM (Progetto INNER JOIN Task ON Progetto.codSIPROS = Task.codSIPROSProg) INNER JOIN PianoTask ON Task.ID = PianoTask.IDTask
WHERE (((PianoTask.scenario)="1Q2019"));

