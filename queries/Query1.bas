SELECT A.key1, A.key2, A.dato, B.dato
FROM Tabella1 A INNER JOIN Tabella2 B ON (A.key2 = B.key2) AND (A.key1 = B.key1)
UNION
SELECT A.key1, A.key2, A.dato, Nz(B.dato, 'zero')
FROM Tabella1 A LEFT JOIN Tabella2 B ON (A.key1 = B.key1) AND (A.key2 = B.key2) WHERE B.key1 IS NULL
UNION SELECT B.key1, B.key2, nz(A.dato,'zero') ,B.dato
FROM Tabella1 A RIGHT  JOIN Tabella2 B ON (A.key1 = B.key1) AND (A.key2 = B.key2)  WHERE  A.key1 IS NULL;

