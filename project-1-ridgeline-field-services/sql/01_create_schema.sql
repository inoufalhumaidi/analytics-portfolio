/*
================================================================================
Project 1 — Ridgeline Field Services: Operational Efficiency
Script:  01_create_schema.sql
Purpose: Creates the RidgelineFieldOps database and its star schema
         (dimension + fact tables) for HVAC/electrical field dispatch
         operations.

DATA DISCLOSURE: Ridgeline Field Services is a fictional company. All data
loaded into this schema is synthetically generated (see 02_generate_
synthetic_data.sql). No confidential, proprietary, or production data of
any real organization is used or represented here.
================================================================================
*/

IF DB_ID('RidgelineFieldOps') IS NULL
BEGIN
    CREATE DATABASE RidgelineFieldOps;
END
GO

USE RidgelineFieldOps;
GO

-- Drop tables if re-running (fact first, then dims) --------------------------
IF OBJECT_ID('dbo.Fact_ServiceJobs', 'U') IS NOT NULL DROP TABLE dbo.Fact_ServiceJobs;
IF OBJECT_ID('dbo.Dim_Technician', 'U') IS NOT NULL DROP TABLE dbo.Dim_Technician;
IF OBJECT_ID('dbo.Dim_Client', 'U') IS NOT NULL DROP TABLE dbo.Dim_Client;
IF OBJECT_ID('dbo.Dim_ServiceType', 'U') IS NOT NULL DROP TABLE dbo.Dim_ServiceType;
IF OBJECT_ID('dbo.Dim_Region', 'U') IS NOT NULL DROP TABLE dbo.Dim_Region;
IF OBJECT_ID('dbo.Dim_Date', 'U') IS NOT NULL DROP TABLE dbo.Dim_Date;
GO

-- =============================================================================
-- DIMENSION: Dim_Date
-- =============================================================================
CREATE TABLE dbo.Dim_Date (
    DateKey         INT         NOT NULL PRIMARY KEY,   -- YYYYMMDD
    [Date]          DATE        NOT NULL,
    [Year]          SMALLINT    NOT NULL,
    [Quarter]       TINYINT     NOT NULL,
    [Month]         TINYINT     NOT NULL,
    MonthName       VARCHAR(10) NOT NULL,
    [Week]          TINYINT     NOT NULL,
    DayOfWeek       TINYINT     NOT NULL,               -- 1=Sunday ... 7=Saturday
    DayName         VARCHAR(10) NOT NULL,
    IsWeekend       BIT         NOT NULL
);
GO

-- =============================================================================
-- DIMENSION: Dim_Region
-- =============================================================================
CREATE TABLE dbo.Dim_Region (
    RegionKey       INT IDENTITY(1,1) PRIMARY KEY,
    RegionCode      VARCHAR(10) NOT NULL UNIQUE,
    RegionName      VARCHAR(50) NOT NULL,
    TechnicianTarget SMALLINT   NOT NULL   -- authorized headcount, used for utilization targets
);
GO

-- =============================================================================
-- DIMENSION: Dim_Technician
-- =============================================================================
CREATE TABLE dbo.Dim_Technician (
    TechnicianKey       INT IDENTITY(1,1) PRIMARY KEY,
    TechnicianID        VARCHAR(10) NOT NULL UNIQUE,
    TechnicianName      VARCHAR(100) NOT NULL,
    RegionKey           INT NOT NULL REFERENCES dbo.Dim_Region(RegionKey),
    SkillLevel          VARCHAR(20) NOT NULL CHECK (SkillLevel IN ('Apprentice','Journeyman','Master')),
    HireDate            DATE NOT NULL,
    EmploymentStatus    VARCHAR(10) NOT NULL CHECK (EmploymentStatus IN ('Active','Inactive')),
    DailyCapacityHours  DECIMAL(4,1) NOT NULL DEFAULT 8.0 CHECK (DailyCapacityHours > 0 AND DailyCapacityHours <= 16)
);
GO

-- =============================================================================
-- DIMENSION: Dim_Client
-- =============================================================================
CREATE TABLE dbo.Dim_Client (
    ClientKey       INT IDENTITY(1,1) PRIMARY KEY,
    ClientID        VARCHAR(10) NOT NULL UNIQUE,
    ClientName      VARCHAR(100) NOT NULL,
    ClientType      VARCHAR(20) NOT NULL CHECK (ClientType IN ('Residential','Commercial')),
    RegionKey       INT NOT NULL REFERENCES dbo.Dim_Region(RegionKey),
    ContractTier    VARCHAR(20) NOT NULL CHECK (ContractTier IN ('Standard','Priority','Premium'))
);
GO

-- =============================================================================
-- DIMENSION: Dim_ServiceType
-- =============================================================================
CREATE TABLE dbo.Dim_ServiceType (
    ServiceTypeKey      INT IDENTITY(1,1) PRIMARY KEY,
    ServiceTypeID       VARCHAR(10) NOT NULL UNIQUE,
    ServiceCategory     VARCHAR(20) NOT NULL CHECK (ServiceCategory IN ('HVAC','Electrical')),
    ServiceName         VARCHAR(60) NOT NULL,
    StandardDurationMin SMALLINT NOT NULL CHECK (StandardDurationMin > 0),
    SLAHours            DECIMAL(5,1) NOT NULL CHECK (SLAHours > 0),
    BaseRevenue         DECIMAL(9,2) NOT NULL CHECK (BaseRevenue >= 0),
    BaseCost            DECIMAL(9,2) NOT NULL CHECK (BaseCost >= 0)
);
GO

-- =============================================================================
-- FACT: Fact_ServiceJobs  (grain: one row per dispatched job)
-- =============================================================================
CREATE TABLE dbo.Fact_ServiceJobs (
    JobKey              BIGINT IDENTITY(1,1) PRIMARY KEY,
    JobID               VARCHAR(15) NOT NULL UNIQUE,
    DateKey             INT NOT NULL REFERENCES dbo.Dim_Date(DateKey),
    TechnicianKey       INT NOT NULL REFERENCES dbo.Dim_Technician(TechnicianKey),
    ClientKey           INT NOT NULL REFERENCES dbo.Dim_Client(ClientKey),
    ServiceTypeKey      INT NOT NULL REFERENCES dbo.Dim_ServiceType(ServiceTypeKey),
    RegionKey           INT NOT NULL REFERENCES dbo.Dim_Region(RegionKey),

    ScheduledStart      DATETIME2(0) NOT NULL,
    ActualStart         DATETIME2(0) NULL,
    ActualEnd           DATETIME2(0) NULL,

    TravelTimeMin       SMALLINT NULL CHECK (TravelTimeMin IS NULL OR TravelTimeMin >= 0),
    JobDurationMin      SMALLINT NULL CHECK (JobDurationMin IS NULL OR JobDurationMin >= 0),

    FirstTimeFix        BIT NULL,       -- NULL only valid when JobStatus <> 'Completed'
    CallbackFlag        BIT NOT NULL DEFAULT 0,
    SLAMet              BIT NULL,       -- NULL only valid when JobStatus <> 'Completed'

    JobCost             DECIMAL(9,2) NULL CHECK (JobCost IS NULL OR JobCost >= 0),
    JobRevenue          DECIMAL(9,2) NULL CHECK (JobRevenue IS NULL OR JobRevenue >= 0),

    JobStatus           VARCHAR(15) NOT NULL CHECK (JobStatus IN ('Completed','Cancelled','Rescheduled')),

    LoadTimestamp       DATETIME2(0) NOT NULL DEFAULT SYSDATETIME()
);
GO

CREATE INDEX IX_FactServiceJobs_DateKey        ON dbo.Fact_ServiceJobs(DateKey);
CREATE INDEX IX_FactServiceJobs_TechnicianKey  ON dbo.Fact_ServiceJobs(TechnicianKey);
CREATE INDEX IX_FactServiceJobs_RegionKey      ON dbo.Fact_ServiceJobs(RegionKey);
CREATE INDEX IX_FactServiceJobs_ServiceTypeKey ON dbo.Fact_ServiceJobs(ServiceTypeKey);
GO

-- =============================================================================
-- REFERENCE: Ref_KPITargets
-- Business-defined thresholds that drive the priority action queue, the
-- Excel "operational control" dashboard, and the Power BI conditional
-- formatting. Centralizing them here means changing a target changes every
-- downstream view/report consistently — this is the single source of truth
-- for "what counts as a problem."
-- =============================================================================
IF OBJECT_ID('dbo.Ref_KPITargets', 'U') IS NOT NULL DROP TABLE dbo.Ref_KPITargets;
GO
CREATE TABLE dbo.Ref_KPITargets (
    MetricName      VARCHAR(40) NOT NULL PRIMARY KEY,
    TargetValue     DECIMAL(6,2) NOT NULL,     -- expressed as the same unit the KPI is reported in (percent or minutes)
    WarningValue    DECIMAL(6,2) NOT NULL,     -- breach point that triggers the priority action queue
    Direction       VARCHAR(15) NOT NULL CHECK (Direction IN ('HigherBetter','LowerBetter')),
    Description     VARCHAR(200) NOT NULL
);
GO

INSERT INTO dbo.Ref_KPITargets (MetricName, TargetValue, WarningValue, Direction, Description) VALUES
('UtilizationPct',      75.00, 60.00, 'HigherBetter', 'Target technician utilization (billable+travel minutes / available capacity)'),
('UtilizationCeilingPct', 90.00, 95.00, 'LowerBetter', 'Above this, technicians are at burnout/overcapacity risk'),
('FirstTimeFixPct',     90.00, 85.00, 'HigherBetter', 'Target first-time-fix rate'),
('SLACompliancePct',    95.00, 90.00, 'HigherBetter', 'Target arrival-window SLA compliance'),
('CallbackPct',          8.00, 12.00, 'LowerBetter',  'Target callback (rework) rate');
GO

PRINT 'Schema created successfully: 5 dimensions + 1 fact table + 1 reference (KPI targets) table.';
GO
