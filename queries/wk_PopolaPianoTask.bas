INSERT INTO PianoTask ( IDTask, dtInizio, dtFine, scenario )
SELECT Foglio1.IDTask, Foglio1.dtIni, Foglio1.dtFin, "4Q2018" AS Espr1
FROM Foglio1;

