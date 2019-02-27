SELECT Utilizzo.IDTask, Utilizzo.IDRisorsa, Utilizzo.mese, Utilizzo.pct
FROM Utilizzo LEFT JOIN Task ON Utilizzo.[IDTask] = Task.[ID]
WHERE (((Task.ID) Is Null));

