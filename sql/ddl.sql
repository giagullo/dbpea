-- Create table scenario
CREATE TABLE scenario (codScenario char(10) primary key ,
			descScenario varchar(255) not null);
ALTER TABLE PianoTask
ADD  FOREIGN KEY (scenario) 
REFERENCES Scenario(codScenario);

CREATE TABLE LogErrori (
	ID timestamp PRIMARY KEY NOT NULL,
	nomPgm varchar(255) NOT NULL,
	numRiga INT ,
	msg varchar(255) );

-- Added 3/3 ---
CREATE TABLE tblErrorLog (
	errorLogId SERIAL PRIMARY KEY,
	errorNo INT NOT NULL, 
	errorMessage varchar(255),
	errorDateTime timestamp,
	errorProc varchar(255),
	WindowsUsername varchar (255));
	
-- ALTER 4/3 
CREATE TABLE CodPPPM (
	codStatoPPPM char(3) NOT NULL PRIMARY KEY,
	descStatoPPPM varchar(255) NOT NULL);
ALTER TABLE Progetto ADD (
	codStatoPPPM char(3) NOT NULL DEFAULT 'NOP');

ALTER TABLE Progetto ADD FOREIGN KEY (codStatoPPPM) 
     REFERENCES CodPPPM(codStatoPPPM);
--- COMMITTED
ALTER TABLE Task (codSIPROS varchar(24) check non newline);
--- COMMITTED