SELECT A.codSIPROS, A.desc, A.codSIPROSTask,  A.Nome, A.mese, A.PctUso AS Prev, B.PctUso AS Cons
FROM qry_AllocazioniScenarioAttuale  A INNER JOIN qry_Consuntivi_anno B ON (A.mese = B.mese) AND (A.Nome = B.Nome) AND (A.codSIPROSTask = B.codSIPROSTask)
UNION
SELECT A.codSIPROS, A.desc, A.codSIPROSTask,  A.Nome, A.mese, A.PctUso AS Prev, B.PctUso As Cons 
FROM qry_AllocazioniScenarioAttuale  A LEFT JOIN qry_Consuntivi_anno B ON (A.mese = B.mese) AND (A.Nome = B.Nome) AND (A.codSIPROSTask = B.codSIPROSTask) 
    WHERE B.codSIPROSTask IS NULL
UNION SELECT B.codSIPROS, B.desc, B.codSIPROSTask, B.Nome, B.mese, A.PctUso AS Prev , B.PctUso As Cons
FROM qry_AllocazioniScenarioAttuale  A RIGHT  JOIN qry_Consuntivi_anno B ON (A.mese = B.mese) AND (A.Nome = B.Nome) AND (A.codSIPROSTask = B.codSIPROSTask) 
    WHERE A.codSIPROSTask IS NULL;

