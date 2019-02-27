SELECT Utilizzo.IDTask, Utilizzo.IDRisorsa, Utilizzo.mese, Utilizzo.pct
FROM Utilizzo LEFT JOIN Risorsa ON Utilizzo.[IDRisorsa] = Risorsa.[ID]
WHERE (((Risorsa.ID) Is Null));

