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

 Date: 09/06/2020 20:00:36
*/


-- ----------------------------
-- Table structure for CAR_CENSUS_LISTS
-- ----------------------------
IF EXISTS (SELECT *
FROM sys.all_objects
WHERE object_id = OBJECT_ID(N'[dbo].[CAR_CENSUS_LISTS]') AND type IN ('U'))
	DROP TABLE [dbo].[CAR_CENSUS_LISTS]
GO

CREATE TABLE [dbo].[CAR_CENSUS_LISTS]
(
  [LOCATION_ESR] int NOT NULL,
  [LIST_NO] int NOT NULL,
  [LINE_NO] int NOT NULL,
  [CAR_NO] int NOT NULL,
  [CAR_TYPE] varchar(6) NOT NULL,
  [CAR_LOCATION] varchar(16) NOT NULL,
  [IS_LOADED] int NOT NULL,
  [IS_WORKING] int NOT NULL,
  [OWNER] varchar(255) NOT NULL,
  [ADM_CODE] int NOT NULL,
  [NON_WORKING_STATE] int NOT NULL,
  [BUILT_YEAR] int NOT NULL
)
GO

ALTER TABLE [dbo].[CAR_CENSUS_LISTS] SET (LOCK_ESCALATION = TABLE)
GO

-- ----------------------------
-- Table structure for STATIONS
-- ----------------------------
IF EXISTS (SELECT *
FROM sys.all_objects
WHERE object_id = OBJECT_ID(N'[dbo].[STATIONS]') AND type IN ('U'))
	DROP TABLE [dbo].[STATIONS]
GO

CREATE TABLE [dbo].[STATIONS]
(
  [ESR] int NOT NULL,
  [NAME] varchar(255) NOT NULL
)
GO

ALTER TABLE [dbo].[STATIONS] SET (LOCK_ESCALATION = TABLE)
GO

