/*
 Navicat Premium Data Transfer

 Source Server         : railway
 Source Server Type    : SQL Server
 Source Server Version : 15004033
 Source Host           : localhost:1433
 Source Catalog        : railway
 Source Schema         : dbo

 Target Server Type    : SQL Server
 Target Server Version : 15004033
 File Encoding         : 65001

 Date: 09/06/2020 19:18:10
*/


-- ----------------------------
-- Table structure for CAR_CENSUS_LISTS
-- ----------------------------
IF EXISTS (SELECT * FROM sys.all_objects WHERE object_id = OBJECT_ID(N'[dbo].[CAR_CENSUS_LISTS]') AND type IN ('U'))
	DROP TABLE [dbo].[CAR_CENSUS_LISTS]
GO

CREATE TABLE [dbo].[CAR_CENSUS_LISTS] (
  [LOCATION_ESR] int  NOT NULL,
  [LIST_NO] int  NOT NULL,
  [LINE_NO] int  NOT NULL,
  [CAR_NO] int  NOT NULL,
  [CAR_TYPE] varchar(6) COLLATE SQL_Latin1_General_CP1_CI_AS  NOT NULL,
  [CAR_LOCATION] int  NOT NULL,
  [IS_LOADED] int  NOT NULL,
  [IS_WORKING] int  NOT NULL,
  [OWNER] varchar(255) COLLATE SQL_Latin1_General_CP1_CI_AS  NOT NULL,
  [ADM_CODE] int  NOT NULL,
  [NON_WORKING_STATE] int  NOT NULL,
  [BUILT_YEAR] int  NOT NULL
)
GO

ALTER TABLE [dbo].[CAR_CENSUS_LISTS] SET (LOCK_ESCALATION = TABLE)
GO


-- ----------------------------
-- Table structure for STATIONS
-- ----------------------------
IF EXISTS (SELECT * FROM sys.all_objects WHERE object_id = OBJECT_ID(N'[dbo].[STATIONS]') AND type IN ('U'))
	DROP TABLE [dbo].[STATIONS]
GO

CREATE TABLE [dbo].[STATIONS] (
  [ESR] int  IDENTITY(1,1) NOT NULL,
  [NAME] varchar(255) COLLATE SQL_Latin1_General_CP1_CI_AS  NOT NULL
)
GO

ALTER TABLE [dbo].[STATIONS] SET (LOCK_ESCALATION = TABLE)
GO


-- ----------------------------
-- Primary Key structure for table STATIONS
-- ----------------------------
ALTER TABLE [dbo].[STATIONS] ADD CONSTRAINT [PK__STATIONS__C19007C815BC6537] PRIMARY KEY CLUSTERED ([ESR])
WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)  
ON [PRIMARY]
GO


-- ----------------------------
-- Foreign Keys structure for table CAR_CENSUS_LISTS
-- ----------------------------
ALTER TABLE [dbo].[CAR_CENSUS_LISTS] ADD CONSTRAINT [FKСAR_CENSUS818939] FOREIGN KEY ([LOCATION_ESR]) REFERENCES [dbo].[STATIONS] ([ESR]) ON DELETE NO ACTION ON UPDATE NO ACTION
GO

