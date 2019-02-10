-- Create table scenario
CREATE TABLE scenario (codScenario char(10) primary key ,
			descScenario varchar(255) not null);
ALTER TABLE PianoTask
ADD  FOREIGN KEY (scenario) 
REFERENCES Scenario(codScenario);

