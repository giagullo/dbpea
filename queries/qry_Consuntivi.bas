SELECT DISTINCTROW Progetto.codSIPROS, Progetto.descProgetto, Task.descTask, Risorsa.Nome, Sum([pct]/100) AS [Mesi impegnati]
FROM Risorsa INNER JOIN ((Progetto INNER JOIN Task ON Progetto.[codSIPROS] = Task.[codSIPROSProg]) INNER JOIN Utilizzo ON Task.[ID] = Utilizzo.[IDTask]) ON Risorsa.[ID] = Utilizzo.[IDRisorsa]
GROUP BY Progetto.codSIPROS, Progetto.descProgetto, Task.descTask, Risorsa.Nome;

