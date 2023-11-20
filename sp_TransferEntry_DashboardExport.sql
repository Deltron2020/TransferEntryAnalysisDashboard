  
CREATE PROCEDURE [dbo].[sp_TransferEntry_DashboardExport]    
AS    
BEGIN    
  
DROP TABLE IF EXISTS ##temp_transfers;  
  
DECLARE @YearID SMALLINT = (SELECT YearID FROM xrYearColor WHERE IsCurrentFlag = 1 GROUP BY YearID);    
  
  
;WITH TransferEntries AS  
(  
SELECT   
 TransferID   
 ,MIN(VersionID) [minVersionID]  
FROM   
 dbo.Transfers t  
GROUP BY  
 TransferID  
)  
  
  
, TransferEntryData AS  
(  
SELECT  
 [TransferID]  
 ,[VersionID]  
 ,[YearID]  
 ,[CreateUser]  
 ,CAST([CreateDate] AS DATE) [CreateDate]  
 ,YEAR([CreateDate]) [Year]  
 ,MONTH([CreateDate]) [Month]  
 ,CONCAT(DATENAME(MONTH,[CreateDate]),'-',YEAR([CreateDate])) [Month-Year] --changed this to datename for month  
 ,CASE WHEN MONTH([CreateDate]) IN (1,2,3) THEN 1  
   WHEN MONTH([CreateDate]) IN (4,5,6) THEN 2  
   WHEN MONTH([CreateDate]) IN (7,8,9) THEN 3  
   WHEN MONTH([CreateDate]) IN (10,11,12) THEN 4  
   ELSE 'NA' END [Quarter]  
 FROM   
  dbo.Transfers t  
)  
  
  
, TransferSalesData AS  
(  
SELECT   
  [TransferID]  
  ,[VersionID]  
  ,CAST([SaleDate] AS DATE) [SaleDate]  
  ,MONTH([SaleDate]) [SaleDateMonth]  
  ,YEAR([SaleDate]) [SaleDateYear]  
  ,CONCAT(MONTH([SaleDate]), '-', YEAR([SaleDate])) [SD_Month-Year]  
  ,CAST([RecordedDate] AS DATE) [RecordedDate]  
  ,[SalePrice]  
  ,[DocumentStamps]  
  ,[Book]  
  ,[Page]  
  ,[SalesCertificate]  
  ,[LegalReference]  
  ,IIF(NULLIF(xrD.xrDeedID,'') IS NULL, '', CONCAT(LTRIM(RTRIM(xrD.Deed)),' - ',LTRIM(RTRIM(xrD.ShortDescription)))) [InstrumentType]  
  ,IIF(NULLIF(t.xrSaleAdjustmentID,'') IS NULL, '', CONCAT(LTRIM(RTRIM(xrSA.SaleAdjustment)),' - ',LTRIM(RTRIM(xrSA.ShortDescription)))) [Tenancy]  
FROM  
 Transfers t  
JOIN  
 (SELECT xrSaleAdjustmentID, SaleAdjustment, ShortDescription FROM dbo.GetxrSaleAdjustmentTable(1,@YearID,1)) xrSA ON xrSA.xrSaleAdjustmentID = t.xrSaleAdjustmentID  
JOIN  
 (SELECT xrDeedID, Deed, ShortDescription FROM dbo.GetxrDeedTable(1,@YearID,1)) xrD ON xrD.xrDeedID = t.xrDeedID  
)  
  
  
, PropertyTransferEntries AS  
(  
 SELECT  
  TransferID  
  ,MIN(VersionID) [minPTVersionID]  
FROM   
 dbo.PropertyTransfers pt  
GROUP BY  
 TransferID  
)  
  
  
, PropertyTransferData AS  
(  
SELECT  
 PropertyTransferID  
 ,VersionID  
 ,YearID  
 ,PropertyID  
 ,TransferID [ptTransferID]  
 ,xrLU.ShortDescription [LUC@Sale]  
 ,IIF(NULLIF(xrSV.xrSalesValidityID,'') IS NULL, '', CONCAT(LTRIM(RTRIM(xrSV.SalesValidity)),' - ',LTRIM(RTRIM(xrSV.ShortDescription)))) [QualCode]  
 ,IIF(NULLIF(pt.xrVerificationID,'') IS NULL,'',CONCAT(LTRIM(RTRIM(xrV.Verification)),' - ',LTRIM(RTRIM(xrV.ShortDescription)))) [ReviewedBy]  
 ,SoldasVacantFlag  
 ,RetainCapFlag  
 ,CreateUser  
 ,CAST([CreateDate] AS DATE) [CreateDate]  
FROM  
 PropertyTransfers pt  
JOIN  
 (SELECT xrLandUseID, LandUse, ShortDescription FROM dbo.GetxrLandUseTable(1,@YearID,1)) xrLU ON xrLU.xrLandUseID = pt.xrSaleLandUseID  
JOIN  
 (SELECT xrSalesValidityID, SalesValidity, ShortDescription FROM dbo.GetxrSalesValidityTable(1,@YearID,1)) xrSV ON xrSV.xrSalesValidityID = pt.xrSalesValidityID  
JOIN  
 (SELECT xrVerificationID, Verification, ShortDescription FROM dbo.GetxrVerificationTable(1,@YearID,1)) xrV ON xrV.xrVerificationID = pt.xrVerificationID  
)  
  
SELECT  
 te.TransferID  
 ,te.minVersionID  
 ,ted.YearID  
 ,ted.CreateUser  
 ,ted.CreateDate  
 ,ted.[Year]  
 ,ted.[Month]  
 ,ted.[Month-Year]  
 ,ted.Quarter  
 ,tsd.SaleDate  
 ,tsd.RecordedDate  
 ,tsd.SalePrice  
 ,tsd.DocumentStamps  
 ,tsd.Book  
 ,tsd.Page  
 ,tsd.SalesCertificate  
 ,tsd.LegalReference  
 ,tsd.InstrumentType  
 ,tsd.Tenancy  
 ,pt.PropertyTransferID  
 ,pt.minPTVersionID  
 ,pt.PropertyID  
 ,pt.LUC@Sale  
 ,pt.QualCode  
 ,pt.ReviewedBy  
 ,pt.SoldasVacantFlag  
 ,pt.RetainCapFlag  
 ,IIF(ted.CreateUser LIKE ('PA\%'), 'Manual Entry', 'Just Appraised') [EntryType]  
 ,tsd.SaleDateMonth  
 ,tsd.SaleDateYear  
 ,tsd.[SD_Month-Year]  
  
INTO ##temp_transfers  
FROM  
 TransferEntries te  
JOIN  
 TransferEntryData ted ON ted.TransferID = te.TransferID AND ted.VersionID = te.minVersionID  
JOIN  
 TransferSalesData tsd ON tsd.TransferID = te.TransferID AND tsd.VersionID = te.minVersionID  
JOIN  
 (  
 SELECT  
  *  
 FROM  
  PropertyTransferEntries pte  
 JOIN  
  PropertyTransferData ptd ON ptd.[ptTransferID] = pte.TransferID AND ptd.VersionID = pte.minPTVersionID  
 )  
  pt ON pt.TransferID = te.TransferID  
  
WHERE   
 1=1  
AND  
 ted.CreateDate >= CAST('01-01-2021' AS DATE) -- Only transfers entered starting in 2021  
AND  
 ted.CreateUser <> 'apro' -- eliminates 3191 transfers bulk entered by Patriot on 1/26/21, 1/27/21, & 3/8/21  
  
/* ====================================== */  
  
 EXEC dbo.ext_ExportDataToCsv @dbName = N'tempdb',          -- nvarchar(100)    
         @includeHeaders = 1, -- bit    
         @filePath = N'\\filepath\Dashboard_Data',        -- nvarchar(512)    
         @tableName = N'##temp_transfers',       -- nvarchar(100)    
         @reportName = N'transfer_entry_data.csv',      -- nvarchar(100)    
         @delimiter = N'|'        -- nvarchar(4)    
    
    
 DECLARE @excelColumns TABLE (Number SMALLINT, Letter VARCHAR(4));    
 INSERT INTO @excelColumns ( Number, Letter )    
 VALUES    
  (1,'A'),    
  (2,'B'),    
  (3,'C'),    
  (4,'D'),    
  (5,'E'),    
  (6,'F'),    
  (7,'G'),    
  (8,'H'),    
  (9,'I'),    
  (10,'J'),    
  (11,'K'),    
  (12,'L'),    
  (13,'M'),    
  (14,'N'),    
  (15,'O'),    
  (16,'P'),    
  (17,'Q'),    
  (18,'R'),    
  (19,'S'),    
  (20,'T'),    
  (21,'U'),    
  (22,'V'),    
  (23,'W'),    
  (24,'X'),    
  (25,'Y'),    
  (26,'Z'),    
  (27,'AA'),    
  (28,'AB'),    
  (29,'AC'),    
  (30,'AD'),    
  (31,'AE'),    
  (32,'AF'),    
  (33,'AG'),    
  (34,'AH'),    
  (35,'AI'),    
  (36,'AJ'),    
  (37,'AK'),    
  (38,'AL'),    
  (39,'AM'),    
  (40,'AN'),    
  (41,'AO'),    
  (42,'AP'),    
  (43,'AQ'),    
  (44,'AR'),    
  (45,'AS'),    
  (46,'AT'),    
  (47,'AU'),    
  (48,'AV'),    
  (49,'AW'),    
  (50,'AX'),    
  (51,'AY'),    
  (52,'AZ');    
    
 --SELECT * FROM @excelColumns    
    
 DECLARE @columnLetter VARCHAR(4) = (SELECT Letter FROM @excelColumns JOIN (SELECT COUNT(COLUMN_NAME) [c] FROM tempdb.INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '##temp_transfers') t ON t.c = [@excelColumns].Number);    
    
 DECLARE @recordCount SMALLINT = (SELECT COUNT(*) + 1 FROM ##temp_transfers);    
    
 EXEC [Assess50].[dbo].CSVtoXLSXwTable @fullCsvPath = '\\filepath\Dashboard_Data\transfer_entry_data.csv',  -- varchar(512)    
               @fullXlsxPath = '\\filepath\Dashboard_Data\transfer_entry_data.xlsx', -- varchar(512)    
               @rowCount = @recordCount,      -- int    
               @colCharacter = @columnLetter  -- varchar(4)    
    
    
 DROP TABLE IF EXISTS ##temp_transfers;    
  
END