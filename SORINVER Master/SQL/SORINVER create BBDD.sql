IF object_id('Sorinver..sysfiles') IS NULL BEGIN
   DECLARE @dir_datos VARCHAR(255)
   exec master.dbo.xp_regread 'HKEY_LOCAL_MACHINE',
           'SOFTWARE\Microsoft\MSSQLServer\Setup',
           'SQLPath',
           @param = @dir_datos OUTPUT
   Exec ('CREATE DATABASE Sorinver ON ( NAME = Sorinver, FILENAME = ''' + @dir_datos  + '\DATA\Sorinver.mdf'', SIZE = 5, MAXSIZE = UNLIMITED, FILEGROWTH = 5)')
END

-- Creamos la tabla de Acciones de empresas 
CREATE TABLE Acciones
(
  Id_Acciones  INT IDENTITY NOT NULL,
  Nombre       VARCHAR (50) NOT NULL,
  TickerYahoo  VARCHAR (20),
  CONSTRAINT PK_Acciones PRIMARY KEY (Id_Acciones)
)
GO

-- Insertamos las empresas del IBEX 35 con el ticker utilizado en Yahoo Finance
INSERT INTO Acciones (Nombre, TickerYahoo) VALUES ('ABERTIS INFRASTRUCTURAS', 'ABE.MC')
INSERT INTO Acciones (Nombre, TickerYahoo) VALUES ('ACCIONA', 'ANA.MC')
INSERT INTO Acciones (Nombre, TickerYahoo) VALUES ('ACERINOX', 'ACX.MC')
INSERT INTO Acciones (Nombre, TickerYahoo) VALUES ('ACERLORMITTAL', 'MTS.MC')
INSERT INTO Acciones (Nombre, TickerYahoo) VALUES ('ACS', 'ACS.MC')
INSERT INTO Acciones (Nombre, TickerYahoo) VALUES ('AENA', 'AENA.MC')
INSERT INTO Acciones (Nombre, TickerYahoo) VALUES ('AMADEUS IT GROUP', 'AMS.MC')
INSERT INTO Acciones (Nombre, TickerYahoo) VALUES ('BANCO DE SABADELL', 'SAB.MC')
INSERT INTO Acciones (Nombre, TickerYahoo) VALUES ('BANCO POPULAR ESPAÑOL', 'POP.MC')
INSERT INTO Acciones (Nombre, TickerYahoo) VALUES ('BANCO SANTANDER', 'SAN.MC')
INSERT INTO Acciones (Nombre, TickerYahoo) VALUES ('BANKIA', 'BKIA.MC')
INSERT INTO Acciones (Nombre, TickerYahoo) VALUES ('BANKINTER', 'BKT.MC')
INSERT INTO Acciones (Nombre, TickerYahoo) VALUES ('BBVA', 'BBVA.MC')
INSERT INTO Acciones (Nombre, TickerYahoo) VALUES ('CAIXABANK', 'CABK.MC')
INSERT INTO Acciones (Nombre, TickerYahoo) VALUES ('CELLNEX TELECOM', 'CLNX.MC')
INSERT INTO Acciones (Nombre, TickerYahoo) VALUES ('DIA', 'DIA.MC')
INSERT INTO Acciones (Nombre, TickerYahoo) VALUES ('ENAGAS', 'ENG.MC')
INSERT INTO Acciones (Nombre, TickerYahoo) VALUES ('ENDESA', 'ELE.MC')
INSERT INTO Acciones (Nombre, TickerYahoo) VALUES ('FERROVIAL', 'FER.MC')
INSERT INTO Acciones (Nombre, TickerYahoo) VALUES ('GAMESA', 'GAM.MC')
INSERT INTO Acciones (Nombre, TickerYahoo) VALUES ('GAS NATURAL', 'GAS.MC')
INSERT INTO Acciones (Nombre, TickerYahoo) VALUES ('GRIFOLS', 'GRF.MC')
INSERT INTO Acciones (Nombre, TickerYahoo) VALUES ('IAG', 'IAG.MC')
INSERT INTO Acciones (Nombre, TickerYahoo) VALUES ('IBERDROLA', 'IBE.MC')
INSERT INTO Acciones (Nombre, TickerYahoo) VALUES ('INDITEX', 'ITX.MC')
INSERT INTO Acciones (Nombre, TickerYahoo) VALUES ('INDRA SISTEMAS', 'IDR.MC')
INSERT INTO Acciones (Nombre, TickerYahoo) VALUES ('MAPFRE', 'MAP.MC')
INSERT INTO Acciones (Nombre, TickerYahoo) VALUES ('MEDIASET', 'TL5.MC')
INSERT INTO Acciones (Nombre, TickerYahoo) VALUES ('MELIA HOTELS', 'MEL.MC')
INSERT INTO Acciones (Nombre, TickerYahoo) VALUES ('MERLIN PROPERTIES', 'MRL.MC')
INSERT INTO Acciones (Nombre, TickerYahoo) VALUES ('RED ELECTRICA CORPORACION', 'REE.MC')
INSERT INTO Acciones (Nombre, TickerYahoo) VALUES ('REPSOL', 'REP.MC')
INSERT INTO Acciones (Nombre, TickerYahoo) VALUES ('TECNICAS REUNIDAS', 'TRE.MC')
INSERT INTO Acciones (Nombre, TickerYahoo) VALUES ('TELEFONICA', 'TEF.MC')
INSERT INTO Acciones (Nombre, TickerYahoo) VALUES ('VISCOFAN', 'VIS.MC')

CREATE TABLE Mercados
(
  Id_Mercados      INT IDENTITY NOT NULL,
  Nombre           VARCHAR (50) NOT NULL,
  TickerYahoo      VARCHAR (10),
  Zona             VARCHAR (20),
  ControlHis       BIT DEFAULT (0),
  ControlCot       BIT DEFAULT (0),
  ControlValorHis  BIT DEFAULT (0),
  ControlValorCot  BIT DEFAULT (0),
  CONSTRAINT PK_Mercados PRIMARY KEY (Id_Mercados)
)
GO

INSERT INTO Mercados (Nombre, TickerYahoo, Zona, ControlHis, ControlCot, ControlValorHis, ControlValorCot)
VALUES ('Ibex35', 'IBEX', 'España', '1', '0', '1', '0')

CREATE TABLE AccionesMercados
(
  Id_Acciones  INT NOT NULL,
  Id_Mercados  INT NOT NULL
)
GO

INSERT INTO AccionesMercados (Id_Acciones, Id_Mercados)
SELECT Id_Acciones, Id_Mercados
FROM Acciones, Mercados
WHERE Mercados.Id_Mercados = '1'

CREATE TABLE Acciones_Indicadores
(
  Id_Acciones  INT NOT NULL,
  K            DECIMAL (16, 4),
  Kn1          DECIMAL (16, 4),
  Kn2          DECIMAL (16, 4),
  Kn3          DECIMAL (16, 4),
  Kn4          DECIMAL (16, 4),
  D            DECIMAL (16, 4),
  Dn1          DECIMAL (16, 4),
  Dn2          DECIMAL (16, 4),
  Dn3          DECIMAL (16, 4),
  Dn4          DECIMAL (16, 4),
  DS           DECIMAL (16, 4),
  DSn1         DECIMAL (16, 4),
  DSn2         DECIMAL (16, 4),
  DSn3         DECIMAL (16, 4),
  DSn4         DECIMAL (16, 4),
  DSS          DECIMAL (16, 4),
  DSSn1        DECIMAL (16, 4),
  DSSn2        DECIMAL (16, 4),
  DSSn3        DECIMAL (16, 4),
  DSSn4        DECIMAL (16, 4),
  RSI          DECIMAL (16, 4),
  RSIn1        DECIMAL (16, 4),
  RSIn2        DECIMAL (16, 4),
  RSIn3        DECIMAL (16, 4),
  RSIn4        DECIMAL (16, 4),
  FechaDatos   DATETIME,
  ActFecha     DATETIME,
  CONSTRAINT PK_Acciones_Indicadores PRIMARY KEY (Id_Acciones)
)
GO

CREATE TABLE AccionesATecnico
(
  Id_Acciones          INT NOT NULL,
  SignoMM20            VARCHAR (1),
  SignoMM50            VARCHAR (1),
  SignoMM200           VARCHAR (1),
  MM20                 DECIMAL (16, 4),
  MM20n1               DECIMAL (16, 4),
  MM20n2               DECIMAL (16, 4),
  MM20n3               DECIMAL (16, 4),
  MM20n4               DECIMAL (16, 4),
  MM50                 DECIMAL (16, 4),
  MM50n1               DECIMAL (16, 4),
  MM50n2               DECIMAL (16, 4),
  MM50n3               DECIMAL (16, 4),
  MM50n4               DECIMAL (16, 4),
  MM200                DECIMAL (16, 4),
  MM200n1              DECIMAL (16, 4),
  MM200n2              DECIMAL (16, 4),
  MM200n3              DECIMAL (16, 4),
  MM200n4              DECIMAL (16, 4),
  VolM20               DECIMAL (16, 4),
  VolMin20             DECIMAL (16, 4),
  VolMax20             DECIMAL (16, 4),
  VolM50               DECIMAL (16, 4),
  VolMin50             DECIMAL (16, 4),
  VolMax50             DECIMAL (16, 4),
  VolM200              DECIMAL (16, 4),
  VolMin200            DECIMAL (16, 4),
  VolMax200            DECIMAL (16, 4),
  VelaM20              DECIMAL (16, 4),
  VelaMin20            DECIMAL (16, 4),
  VelaMax20            DECIMAL (16, 4),
  VelaM50              DECIMAL (16, 4),
  VelaMin50            DECIMAL (16, 4),
  VelaMax50            DECIMAL (16, 4),
  VelaM200             DECIMAL (16, 4),
  VelaMin200           DECIMAL (16, 4),
  VelaMax200           DECIMAL (16, 4),
  Soporte20            DECIMAL (16, 4),
  FechaSoporte20       DATETIME,
  Resistencia20        DECIMAL (16, 4),
  FechaResistencia20   DATETIME,
  Soporte50            DECIMAL (16, 4),
  FechaSoporte50       DATETIME,
  Resistencia50        DECIMAL (16, 4),
  FechaResistencia50   DATETIME,
  Soporte200           DECIMAL (16, 4),
  FechaSoporte200      DATETIME,
  Resistencia200       DECIMAL (16, 4),
  FechaResistencia200  DATETIME,
  Fecha1TALargo        DATETIME,
  Valor1TALargo        DECIMAL (16, 4),
  Fecha2TALargo        DATETIME,
  Valor2TALargo        DECIMAL (16, 4),
  PorTALargo           DECIMAL (16, 4),
  PorAcumuladoTALargo  DECIMAL (16, 4),
  DiasTALargo          DECIMAL (16, 4),
  TipoTALargo          VARCHAR (20),
  Fecha1TAMedio        DATETIME,
  Valor1TAMedio        DECIMAL (16, 4),
  Fecha2TAMedio        DATETIME,
  Valor2TAMedio        DECIMAL (16, 4),
  PorTAMedio           DECIMAL (16, 4),
  PorAcumuladoTAMedio  DECIMAL (16, 4),
  DiasTAMedio          DECIMAL (16, 4),
  TipoTAMedio          VARCHAR (20),
  Fecha1TACorto        DATETIME,
  Valor1TACorto        DECIMAL (16, 4),
  Fecha2TACorto        DATETIME,
  Valor2TACorto        DECIMAL (16, 4),
  PorTACorto           DECIMAL (16, 4),
  PorAcumuladoTACorto  DECIMAL (16, 4),
  DiasTACorto          DECIMAL (16, 4),
  TipoTACorto          VARCHAR (20),
  Fecha1TBLargo        DATETIME,
  Valor1TBLargo        DECIMAL (16, 4),
  Fecha2TBLargo        DATETIME,
  Valor2TBLargo        DECIMAL (16, 4),
  PorTBLargo           DECIMAL (16, 4),
  PorAcumuladoTBLargo  DECIMAL (16, 4),
  DiasTBLargo          DECIMAL (16, 4),
  TipoTBLargo          VARCHAR (20),
  Fecha1TBMedio        DATETIME,
  Valor1TBMedio        DECIMAL (16, 4),
  Fecha2TBMedio        DATETIME,
  Valor2TBMedio        DECIMAL (16, 4),
  PorTBMedio           DECIMAL (16, 4),
  PorAcumuladoTBMedio  DECIMAL (16, 4),
  DiasTBMedio          DECIMAL (16, 4),
  TipoTBMedio          VARCHAR (20),
  Fecha1TBCorto        DATETIME,
  Valor1TBCorto        DECIMAL (16, 4),
  Fecha2TBCorto        DATETIME,
  Valor2TBCorto        DECIMAL (16, 4),
  PorTBCorto           DECIMAL (16, 4),
  PorAcumuladoTBCorto  DECIMAL (16, 4),
  DiasTBCorto          DECIMAL (16, 4),
  TipoTBCorto          VARCHAR (20),
  FechaDatos           DATETIME,
  ActFecha             DATETIME,
  CONSTRAINT PK_AccionesATecnico PRIMARY KEY (Id_Acciones)
)
GO

CREATE TABLE AccionesCotizaciones
(
  Id_Acciones  INT NOT NULL,
  Fecha        DATETIME NOT NULL,
  Apertura     DECIMAL (16, 4) NOT NULL,
  Cierre       DECIMAL (16, 4) NOT NULL,
  Maximo       DECIMAL (16, 4) NOT NULL,
  Minimo       DECIMAL (16, 4) NOT NULL,
  Volumen      DECIMAL (16, 4) NOT NULL,
  Origen       VARCHAR (1),
  ActFecha     DATETIME
)
GO



CREATE TABLE Mercados_Indicadores
(
  Id_Mercados  INT NOT NULL,
  K            DECIMAL (16, 4) DEFAULT (0),
  Kn1          DECIMAL (16, 4) DEFAULT (0),
  Kn2          DECIMAL (16, 4) DEFAULT (0),
  Kn3          DECIMAL (16, 4) DEFAULT (0),
  Kn4          DECIMAL (16, 4) DEFAULT (0),
  D            DECIMAL (16, 4) DEFAULT (0),
  Dn1          DECIMAL (16, 4) DEFAULT (0),
  Dn2          DECIMAL (16, 4) DEFAULT (0),
  Dn3          DECIMAL (16, 4) DEFAULT (0),
  Dn4          DECIMAL (16, 4) DEFAULT (0),
  DS           DECIMAL (16, 4) DEFAULT (0),
  DSn1         DECIMAL (16, 4) DEFAULT (0),
  DSn2         DECIMAL (16, 4) DEFAULT (0),
  DSn3         DECIMAL (16, 4) DEFAULT (0),
  DSn4         DECIMAL (16, 4) DEFAULT (0),
  DSS          DECIMAL (16, 4) DEFAULT (0),
  DSSn1        DECIMAL (16, 4) DEFAULT (0),
  DSSn2        DECIMAL (16, 4) DEFAULT (0),
  DSSn3        DECIMAL (16, 4) DEFAULT (0),
  DSSn4        DECIMAL (16, 4) DEFAULT (0),
  RSI          DECIMAL (16, 4) DEFAULT (0),
  RSIn1        DECIMAL (16, 4) DEFAULT (0),
  RSIn2        DECIMAL (16, 4) DEFAULT (0),
  RSIn3        DECIMAL (16, 4) DEFAULT (0),
  RSIn4        DECIMAL (16, 4) DEFAULT (0),
  FechaDatos   DATETIME DEFAULT ('1/1/1900'),
  ActFecha     DATETIME DEFAULT ('1/1/1900'),
  CONSTRAINT PK_Mercados_Indicadores PRIMARY KEY (Id_Mercados)
)
GO

CREATE TABLE MercadosATecnico
(
  Id_Mercados          INT NOT NULL,
  SignoMM20            VARCHAR (1) DEFAULT ('='),
  SignoMM50            VARCHAR (1) DEFAULT ('='),
  SignoMM200           VARCHAR (1) DEFAULT ('='),
  MM20                 DECIMAL (16, 4) DEFAULT (0),
  MM20n1               DECIMAL (16, 4) DEFAULT (0),
  MM20n2               DECIMAL (16, 4) DEFAULT (0),
  MM20n3               DECIMAL (16, 4) DEFAULT (0),
  MM20n4               DECIMAL (16, 4) DEFAULT (0),
  MM50                 DECIMAL (16, 4) DEFAULT (0),
  MM50n1               DECIMAL (16, 4) DEFAULT (0),
  MM50n2               DECIMAL (16, 4) DEFAULT (0),
  MM50n3               DECIMAL (16, 4) DEFAULT (0),
  MM50n4               DECIMAL (16, 4) DEFAULT (0),
  MM200                DECIMAL (16, 4) DEFAULT (0),
  MM200n1              DECIMAL (16, 4) DEFAULT (0),
  MM200n2              DECIMAL (16, 4) DEFAULT (0),
  MM200n3              DECIMAL (16, 4) DEFAULT (0),
  MM200n4              DECIMAL (16, 4) DEFAULT (0),
  VolM20               DECIMAL (16, 4) DEFAULT (0),
  VolMin20             DECIMAL (16, 4) DEFAULT (0),
  VolMax20             DECIMAL (16, 4) DEFAULT (0),
  VolM50               DECIMAL (16, 4) DEFAULT (0),
  VolMin50             DECIMAL (16, 4) DEFAULT (0),
  VolMax50             DECIMAL (16, 4) DEFAULT (0),
  VolM200              DECIMAL (16, 4) DEFAULT (0),
  VolMin200            DECIMAL (16, 4) DEFAULT (0),
  VolMax200            DECIMAL (16, 4) DEFAULT (0),
  VelaM20              DECIMAL (16, 4) DEFAULT (0),
  VelaMin20            DECIMAL (16, 4) DEFAULT (0),
  VelaMax20            DECIMAL (16, 4) DEFAULT (0),
  VelaM50              DECIMAL (16, 4) DEFAULT (0),
  VelaMin50            DECIMAL (16, 4) DEFAULT (0),
  VelaMax50            DECIMAL (16, 4) DEFAULT (0),
  VelaM200             DECIMAL (16, 4) DEFAULT (0),
  VelaMin200           DECIMAL (16, 4) DEFAULT (0),
  VelaMax200           DECIMAL (16, 4) DEFAULT (0),
  Soporte20            DECIMAL (16, 4) DEFAULT (0),
  FechaSoporte20       DATETIME DEFAULT ('1/1/1900'),
  Resistencia20        DECIMAL (16, 4) DEFAULT (0),
  FechaResistencia20   DATETIME DEFAULT ('1/1/1900'),
  Soporte50            DECIMAL (16, 4) DEFAULT (0),
  FechaSoporte50       DATETIME DEFAULT ('1/1/1900'),
  Resistencia50        DECIMAL (16, 4) DEFAULT (0),
  FechaResistencia50   DATETIME DEFAULT (0),
  Soporte200           DECIMAL (16, 4) DEFAULT (0),
  FechaSoporte200      DATETIME DEFAULT ('1/1/1900'),
  Resistencia200       DECIMAL (16, 4) DEFAULT (0),
  FechaResistencia200  DATETIME DEFAULT ('1/1/1900'),
  Fecha1TALargo        DATETIME DEFAULT ('1/1/1900'),
  Valor1TALargo        DECIMAL (16, 4) DEFAULT (0),
  Fecha2TALargo        DATETIME DEFAULT ('1/1/1900'),
  Valor2TALargo        DECIMAL (16, 4) DEFAULT (0),
  PorTALargo           DECIMAL (16, 4) DEFAULT (0),
  PorAcumuladoTALargo  DECIMAL (16, 4),
  DiasTALargo          DECIMAL (16, 4) DEFAULT (0),
  TipoTALargo          VARCHAR (20),
  Fecha1TAMedio        DATETIME DEFAULT ('1/1/1900'),
  Valor1TAMedio        DECIMAL (16, 4) DEFAULT (0),
  Fecha2TAMedio        DATETIME DEFAULT ('1/1/1900'),
  Valor2TAMedio        DECIMAL (16, 4) DEFAULT (0),
  PorTAMedio           DECIMAL (16, 4) DEFAULT (0),
  PorAcumuladoTAMedio  DECIMAL (16, 4) DEFAULT (0),
  DiasTAMedio          DECIMAL (16, 4) DEFAULT (0),
  TipoTAMedio          VARCHAR (20),
  Fecha1TACorto        DATETIME DEFAULT ('1/1/1900'),
  Valor1TACorto        DECIMAL (16, 4) DEFAULT (0),
  Fecha2TACorto        DATETIME DEFAULT ('1/1/1900'),
  Valor2TACorto        DECIMAL (16, 4) DEFAULT (0),
  PorTACorto           DECIMAL (16, 4) DEFAULT (0),
  PorAcumuladoTACorto  DECIMAL (16, 4) DEFAULT (0),
  DiasTACorto          DECIMAL (16, 4) DEFAULT (0),
  TipoTACorto          VARCHAR (20),
  Fecha1TBLargo        DATETIME DEFAULT ('1/1/1900'),
  Valor1TBLargo        DECIMAL (16, 4) DEFAULT (0),
  Fecha2TBLargo        DATETIME DEFAULT ('1/1/1900'),
  Valor2TBLargo        DECIMAL (16, 4) DEFAULT (0),
  PorTBLargo           DECIMAL (16, 4) DEFAULT (0),
  PorAcumuladoTBLargo  DECIMAL (16, 4) DEFAULT (0),
  DiasTBLargo          DECIMAL (16, 4) DEFAULT (0),
  TipoTBLargo          VARCHAR (20),
  Fecha1TBMedio        DATETIME DEFAULT ('1/1/1900'),
  Valor1TBMedio        DECIMAL (16, 4) DEFAULT (0),
  Fecha2TBMedio        DATETIME DEFAULT ('1/1/1900'),
  Valor2TBMedio        DECIMAL (16, 4) DEFAULT (0),
  PorTBMedio           DECIMAL (16, 4) DEFAULT (0),
  PorAcumuladoTBMedio  DECIMAL (16, 4) DEFAULT (0),
  DiasTBMedio          DECIMAL (16, 4) DEFAULT (0),
  TipoTBMedio          VARCHAR (20),
  Fecha1TBCorto        DATETIME DEFAULT ('1/1/1900'),
  Valor1TBCorto        DECIMAL (16, 4) DEFAULT (0),
  Fecha2TBCorto        DATETIME DEFAULT ('1/1/1900'),
  Valor2TBCorto        DECIMAL (16, 4) DEFAULT (0),
  PorTBCorto           DECIMAL (16, 4) DEFAULT (0),
  PorAcumuladoTBCorto  DECIMAL (16, 4) DEFAULT (0),
  DiasTBCorto          DECIMAL (16, 4) DEFAULT (0),
  TipoTBCorto          VARCHAR (20),
  FechaDatos           DATETIME DEFAULT ('1/1/1900'),
  ActFecha             DATETIME DEFAULT ('1/1/1900'),
  CONSTRAINT PK_MercadosATecnico PRIMARY KEY (Id_Mercados)
)
GO

CREATE TABLE MercadosCotizaciones
(
  Id_Mercados  INT NOT NULL,
  Fecha        DATETIME NOT NULL,
  Apertura     DECIMAL (16, 4) NOT NULL,
  Cierre       DECIMAL (16, 4) NOT NULL,
  Maximo       DECIMAL (16, 4) NOT NULL,
  Minimo       DECIMAL (16, 4) NOT NULL,
  Volumen      DECIMAL (16, 4) NOT NULL,
  Origen       VARCHAR (1),
  ActFecha     DATETIME
)
GO


CREATE TABLE Defectos
(
  Id_Defectos	INT NOT NULL ,
  Lane_Periodo INT,
  Lane_K INT,
  Lane_D INT,
  Lane_DS INT,
  Lane_DSS INT,
  Lane_SobreCompra INT,
  Lane_SobreVenta INT,
  Lane_AvisoClasica_Lento BIT,
  Lane_AvisoClasica_Rapido BIT,
  Lane_AvisoSZona_Lento BIT,
  Lane_AvisoSZona_Rapido BIT,
  Lane_AvisoPopCorn_Lento BIT,
  Lane_AvisoPopCorn_Rapido BIT,
  RSI_Periodo INT,
  RSI_SobreCompra INT,
  RSI_SobreVenta INT,
  RSI_AvisoSalidaZona INT,
  RSI_AvisoFailureSwing INT,
  RSI_AvisoDivergencia INT,
  ActFecha	DATETIME,
  CONSTRAINT PK_Defectos PRIMARY KEY (Id_Defectos)
)
GO

INSERT INTO Defectos (Id_Defectos, Lane_Periodo, Lane_K, Lane_D, Lane_DS, Lane_DSS, Lane_SobreCompra, Lane_SobreVenta, Lane_AvisoClasica_Lento, Lane_AvisoClasica_Rapido, Lane_AvisoSZona_Lento, Lane_AvisoSZona_Rapido, Lane_AvisoPopCorn_Lento, Lane_AvisoPopCorn_Rapido, RSI_Periodo, RSI_SobreCompra, RSI_SobreVenta, RSI_AvisoSalidaZona, RSI_AvisoFailureSwing, RSI_AvisoDivergencia)
VALUES ('1', '14', '3', '3', '3', '3', '80', '20', '1', '1', '0', '0', '0', '0', '14', '70', '30', '1', '1', '1')

CREATE TABLE Velas
(
  Id_Velas     INT IDENTITY NOT NULL,
  Valor        INT NOT NULL,
  DesCorta     VARCHAR (30),
  DesLarga     VARCHAR (100),
  Comentarios  VARCHAR (200),
  Aviso        BIT DEFAULT (0),
  CONSTRAINT PK_Velas PRIMARY KEY (Id_Velas)
)
GO

INSERT INTO Velas (Valor, DesCorta, DesLarga, Aviso) VALUES ('1', 'Doji mas', 'Doji mas', '0')
INSERT INTO Velas (Valor, DesCorta, DesLarga, Aviso) VALUES ('2', 'Doji cruz', 'Doji cruz', '0')
INSERT INTO Velas (Valor, DesCorta, DesLarga, Aviso) VALUES ('3', 'Doji cruz invertida', 'Doji cruz invertida', '0')
INSERT INTO Velas (Valor, DesCorta, DesLarga, Aviso) VALUES ('4', 'Doji dragon volador', 'Doji dragon volador T', '0')
INSERT INTO Velas (Valor, DesCorta, DesLarga, Aviso) VALUES ('5', 'Doji piedra funeraria', 'Doji piedra funeraria T invertida', '0')
INSERT INTO Velas (Valor, DesCorta, DesLarga, Aviso) VALUES ('11', 'Martillo/Colgado', 'Martillo/Colgado', '0')
INSERT INTO Velas (Valor, DesCorta, DesLarga, Aviso) VALUES ('12', 'Estrella fugaz', 'Estrella fugaz', '0')
INSERT INTO Velas (Valor, DesCorta, DesLarga, Aviso) VALUES ('13', 'Peonza', 'Peonza', '0')
INSERT INTO Velas (Valor, DesCorta, DesLarga, Aviso) VALUES ('21', 'Marubozu medio', 'Marubozu medio', '0')
INSERT INTO Velas (Valor, DesCorta, DesLarga, Aviso) VALUES ('22', 'Cuerpo medio', 'Cuerpo medio', '0')
INSERT INTO Velas (Valor, DesCorta, DesLarga, Aviso) VALUES ('31', 'Marubozu grande', 'Marubozu grande', '0')
INSERT INTO Velas (Valor, DesCorta, DesLarga, Aviso) VALUES ('32', 'Cuerpo grande', 'Cuerpo grande', '0')


