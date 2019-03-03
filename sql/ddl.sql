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
	
	

