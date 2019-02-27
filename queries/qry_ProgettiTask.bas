SELECT Progetto.codSIPROS, Progetto.descProgetto, Task.codSIPROS, Task.descTask
FROM Progetto INNER JOIN Task ON Progetto.codSIPROS = Task.codIniziativa
ORDER BY Progetto.codSIPROS, Task.codSIPROS;

