--Create Dummy PLDA Import Header for unit testing
CREATE TABLE [PLDA IMPORT DETAIL] (
	[Code] VARCHAR(21) NOT NULL,
	[Header] INT  NOT NULL,
	[Detail] INT  NOT NULL,
	[L1] VARCHAR(10) NULL,
	[L2] VARCHAR(4) NULL,
	[L3] VARCHAR(4) NULL,
	[L4] VARCHAR(4) NULL,
	[L5] VARCHAR(4) NULL,
	CONSTRAINT pkPLDAImportDetail PRIMARY KEY(Code, Header, Detail)
);

--Populate with Test Data
INSERT INTO [PLDA IMPORT HEADER] VALUES  ('000000993949532508849', 1, 'IM', 'Z', 'P945304849540810005246', '20081206', 'VP442020');
INSERT INTO [PLDA IMPORT HEADER] VALUES  ('000000237661242485047', 1, 'IM', 'Y', 'P945304849540810005248', '20081207', 'VP442023');
INSERT INTO [PLDA IMPORT HEADER] VALUES  ('000000958309650421142', 1, 'IM', 'Z', 'P945304849540810005250', '20081208', 'VP442026');
INSERT INTO [PLDA IMPORT HEADER] VALUES  ('000000372680902481079', 1, 'IM', 'Y', 'P945304849540810005255', '20081209', 'VP442029');
INSERT INTO [PLDA IMPORT HEADER] VALUES  ('000000757792890071868', 1, 'IM', 'Z', 'P945304849540810005286', '20081210', 'VP442037');
