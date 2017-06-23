CREATE PROCEDURE [dbo].[AB_OS]
	-- Add the parameters for the stored procedure here
	@FromDt Datetime,@ToDate Datetime,@LOS Varchar(100)
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;
--use PWC_DEV
Declare @CompnyCode Varchar(100);
Declare @CompnyName Varchar(100);

--declare @FromDt Datetime;
--declare @ToDate Datetime;
--set @FromDt='2015-01-01 00:00'
--set @ToDate='2015-03-10 00:00'


declare @Month int;
declare @Year int;
set @Month= Month(@ToDate);

Set @Year=Year(@ToDate);
Declare @dt as Varchar(30);
set @Dt = (Select cast(@Year as varchar) + '-' + cast(@Month as varchar)  + '-' + '01')

Declare @LastMon as Varchar(30)
Declare @LastYr as Varchar(30)
Declare @PriorYr as Varchar(30)

Declare @STMONTH AS VARCHAR(30)
Declare @STYEAR AS VARCHAR(30)

SET @STMONTH = (SELECT Month(DateAdd(month, 6, Convert(date,  @Dt ))));
SET @STYEAR = (SELECT Year(DateAdd(month, 6, Convert(date,  @Dt ))));

set @LastMon= (SELECT Month(DateAdd(month, -1, Convert(date,  @Dt ))));
Set @LastYr=(SELECT Year(DateAdd(month, -1, Convert(date,  @Dt ))));


Declare @STLASTMON AS VARCHAR(30)
Declare @STLASTYR AS VARCHAR(30)

SET @STLASTMON = (SELECT Month(DateAdd(month, 5, Convert(date,  @Dt ))));
SET @STLASTYR = (SELECT Year(DateAdd(month, 5, Convert(date,  @Dt ))));

Set @PriorYr=(SELECT Year(DateAdd(YEAR, -1, Convert(date,  @Dt ))));

Declare @STPRIORYR AS VARCHAR(30)
SET @STPRIORYR =(SELECT Year(DateAdd(YEAR, 5, Convert(date,  @Dt ))));

Declare @STPRTIORYR AS VARCHAR(30)
SET @STPRTIORYR =(SELECT Year(DateAdd(YEAR, 5, Convert(date,  @Dt ))));


Create  Table #OPERRES (Rownum int,Bold Varchar(1),CompanyCode  Varchar(50),CompanyName  Varchar(100),
Level1   Varchar(100),Level2  Varchar(100),Level3 Varchar(100),Actual [numeric](19, 6) NULL,Budget [numeric](19, 6) NULL,
PriorYr [numeric](19, 6) NULL,YTDActual [numeric](19, 6) NULL,YTDBudget [numeric](19, 6) NULL,YTDPriorYr [numeric](19, 6) NULL);

BEGIN ---[AAAS]-ST 

Set @CompnyCode=(SELECT Top 1 CompnyName FROM [AAAS].[dbo].OADM with (Nolock));
Set @CompnyName=(SELECT Top 1 CompnyName FROM [AAAS].[dbo].OADM with (Nolock)) ;

Insert into #OPERRES Values(1 , 'N', @CompnyCode,@CompnyName,'Engagement Margin','Engagement Revenue','Market Scale',0,0,0,0,0,0)
Insert into #OPERRES Values(10, 'N', @CompnyCode,@CompnyName,'Engagement Margin','Engagement Revenue','Planned variance',0,0,0,0,0,0)
Insert into #OPERRES Values(20, 'N', @CompnyCode,@CompnyName,'Engagement Margin','Engagement Revenue','Actual variance (WIP system)',0,0,0,0,0,0)
Insert into #OPERRES Values(30, 'N', @CompnyCode,@CompnyName,'Engagement Margin','Engagement Revenue','Manual adjustments (G/L only)',0,0,0,0,0,0)
Insert into #OPERRES Values(40, 'N', @CompnyCode,@CompnyName,'Engagement Margin','Engagement Revenue','Disbursements charged',0,0,0,0,0,0)

Insert into #OPERRES Values(50, 'N', @CompnyCode,@CompnyName,'Engagement Margin','Engagement Revenue','Time based revenue',0,0,0,0,0,0)
Insert into #OPERRES Values(60, 'Y', @CompnyCode,@CompnyName,'Engagement Margin','Engagement Revenue','Other revenue',0,0,0,0,0,0)
Insert into #OPERRES Values(70, 'N', @CompnyCode,@CompnyName,'Engagement Margin','Engagement Revenue','A/R provisions and write-offs',0,0,0,0,0,0)
Insert into #OPERRES Values(80, 'N', @CompnyCode,@CompnyName,'Engagement Margin','Engagement Revenue','Territory revenue from external clients',0,0,0,0,0,0)
Insert into #OPERRES Values(81, 'N', @CompnyCode,@CompnyName,'Engagement Margin','Engagement Revenue','Less: Disbursements costs (3rd party)',0,0,0,0,0,0)
Insert into #OPERRES Values(82, 'N', @CompnyCode,@CompnyName,'Engagement Margin','Engagement Revenue','Less: Inter territory costs',0,0,0,0,0,0)
Insert into #OPERRES Values(83, 'Y', @CompnyCode,@CompnyName,'Engagement Margin','Engagement Revenue','Net Revenue from external clients',0,0,0,0,0,0)


Insert into #OPERRES Values(90, 'N', @CompnyCode,@CompnyName,'Cost of Sales','','ResourceCosts',0,0,0,0,0,0)
Insert into #OPERRES Values(100, 'N', @CompnyCode,@CompnyName,'Cost of Sales','','Client proposal / relationship costs',0,0,0,0,0,0)
Insert into #OPERRES Values(110, 'Y', @CompnyCode,@CompnyName,'Cost of Sales','','Engagement margin.',0,0,0,0,0,0)


Insert into #OPERRES Values(120, 'N', @CompnyCode,@CompnyName,'Resource Margin','','Income from chargeable hours - external',0,0,0,0,0,0)
Insert into #OPERRES Values(130, 'N', @CompnyCode,@CompnyName,'Resource Margin','','Income from chargeable hours - internal',0,0,0,0,0,0)
Insert into #OPERRES Values(140, 'Y', @CompnyCode,@CompnyName,'Resource Margin','','Income from resource utilisation',0,0,0,0,0,0)
Insert into #OPERRES Values(150, 'Y', @CompnyCode,@CompnyName,'Resource Margin','','Net fee income before B/L Mark up',0,0,0,0,0,0)
Insert into #OPERRES Values(160, 'Y', @CompnyCode,@CompnyName,'Resource Margin','','Borrow / Loan mark up',0,0,0,0,0,0)
Insert into #OPERRES Values(170, 'Y', @CompnyCode,@CompnyName,'Resource Margin','','Net fee income after B/L Mark up',0,0,0,0,0,0)

Insert into #OPERRES Values(180, 'N', @CompnyCode,@CompnyName, 'Direct Costs','People related costs','Client service staff costs',0,0,0,0,0,0)
Insert into #OPERRES Values(181, 'N', @CompnyCode,@CompnyName, 'Direct Costs','People related costs','Charges from SDCs',0,0,0,0,0,0)
Insert into #OPERRES Values(190, 'N', @CompnyCode,@CompnyName, 'Direct Costs','People related costs','Practice support staff costs',0,0,0,0,0,0)
Insert into #OPERRES Values(200, 'N', @CompnyCode,@CompnyName, 'Direct Costs','People related costs','Continuing education',0,0,0,0,0,0)
Insert into #OPERRES Values(210, 'N', @CompnyCode,@CompnyName, 'Direct Costs','People related costs','Other personnel costs',0,0,0,0,0,0)
Insert into #OPERRES Values(220, 'N', @CompnyCode,@CompnyName, 'Direct Costs','People related costs','Recruiting',0,0,0,0,0,0)
Insert into #OPERRES Values(230, 'N', @CompnyCode,@CompnyName, 'Direct Costs','People related costs','All Other',0,0,0,0,0,0)

Insert into #OPERRES Values(240, 'N', @CompnyCode,@CompnyName, 'Occupancy and Infrastructure','','Occupancy',0,0,0,0,0,0)
Insert into #OPERRES Values(250, 'N', @CompnyCode,@CompnyName, 'Occupancy and Infrastructure','','Communications',0,0,0,0,0,0)
Insert into #OPERRES Values(260, 'N', @CompnyCode,@CompnyName, 'Occupancy and Infrastructure','','Technology',0,0,0,0,0,0)
Insert into #OPERRES Values(270, 'N', @CompnyCode,@CompnyName, 'Occupancy and Infrastructure','','All Other',0,0,0,0,0,0)

--Insert into #OPERRES Values(280, 'N', @CompnyCode,@CompnyName,  'General (excluding interest and taxes)','','Occupancy',0,0,0,0,0,0)
Insert into #OPERRES Values(290, 'N', @CompnyCode,@CompnyName,  'General (excluding interest and taxes)','','Travel and subsistence costs',0,0,0,0,0,0)
Insert into #OPERRES Values(300, 'N', @CompnyCode,@CompnyName,  'General (excluding interest and taxes)','','Marketing and business development',0,0,0,0,0,0)
Insert into #OPERRES Values(310, 'N', @CompnyCode,@CompnyName,  'General (excluding interest and taxes)','','Practice protection',0,0,0,0,0,0)
Insert into #OPERRES Values(320, 'N', @CompnyCode,@CompnyName,  'General (excluding interest and taxes)','','Imputed interest',0,0,0,0,0,0)
Insert into #OPERRES Values(330, 'N', @CompnyCode,@CompnyName,  'General (excluding interest and taxes)','','Goodwill amortisation',0,0,0,0,0,0)
Insert into #OPERRES Values(340, 'N', @CompnyCode,@CompnyName,  'General (excluding interest and taxes)','','Other income and expenses',0,0,0,0,0,0)
Insert into #OPERRES Values(350, 'N', @CompnyCode,@CompnyName,  'General (excluding interest and taxes)','','Internal projects debit at resource cost',0,0,0,0,0,0)
Insert into #OPERRES Values(360, 'N', @CompnyCode,@CompnyName,  'General (excluding interest and taxes)','','Total direct costs',0,0,0,0,0,0)

Insert into #OPERRES Values(370, 'Y', @CompnyCode,@CompnyName,'Resource margin.','','Resource margin',0,0,0,0,0,0)
Insert into #OPERRES Values(380, 'Y', @CompnyCode,@CompnyName,'Controllable margin.','','Controllable margin',0,0,0,0,0,0)

Insert into #OPERRES Values(390, 'Y', @CompnyCode,@CompnyName,'Key Metrics','','',0,0,0,0,0,0)

Insert into #OPERRES Values(400, 'N', @CompnyCode,@CompnyName,'Revenue','','Net Revenue per host hour ($)',0,0,0,0,0,0)

Insert into #OPERRES Values(410, 'N', @CompnyCode,@CompnyName,'Profitability','','Time realisation %',0,0,0,0,0,0)
Insert into #OPERRES Values(420, 'N', @CompnyCode,@CompnyName,'Profitability','','Engagement Margin %',0,0,0,0,0,0)
Insert into #OPERRES Values(421, 'N', @CompnyCode,@CompnyName,'Profitability','','Gross Margin %',0,0,0,0,0,0)
Insert into #OPERRES Values(430, 'N', @CompnyCode,@CompnyName,'Profitability','','Controllable Margin %',0,0,0,0,0,0)
Insert into #OPERRES Values(431, 'N', @CompnyCode,@CompnyName,'Profitability','','Gross Margin',0,0,0,0,0,0)

Insert into #OPERRES Values(440, 'N', @CompnyCode,@CompnyName,'Cost Management','','Direct Costs per Capita (per Total CS) (¥)',0,0,0,0,0,0)
Insert into #OPERRES Values(450, 'N', @CompnyCode,@CompnyName,'Cost Management','','Direct Costs as % of CS Staff Costs',0,0,0,0,0,0)
Insert into #OPERRES Values(451, 'N', @CompnyCode,@CompnyName,'Cost Management','','Overhead per Chargeable hour',0,0,0,0,0,0)

Insert into #OPERRES Values(460, 'N', @CompnyCode,@CompnyName,'Resource Management','','Client service staff costs per hour ($)',0,0,0,0,0,0)
Insert into #OPERRES Values(470, 'N', @CompnyCode,@CompnyName,'Resource Management','','Resource Margin %',0,0,0,0,0,0)

Insert into #OPERRES Values(480, 'N', @CompnyCode,@CompnyName,'Resource Utilisation','','Partner Utilisation %',0,0,0,0,0,0)
Insert into #OPERRES Values(490, 'N', @CompnyCode,@CompnyName,'Resource Utilisation','','Client Service Staff Utilisation %',0,0,0,0,0,0)
Insert into #OPERRES Values(500, 'N', @CompnyCode,@CompnyName,'Resource Utilisation','','Total Resource Utilisation %',0,0,0,0,0,0)
Insert into #OPERRES Values(501, 'N', @CompnyCode,@CompnyName,'Resource Utilisation','','Partner hours per head',0,0,0,0,0,0)
Insert into #OPERRES Values(502, 'N', @CompnyCode,@CompnyName,'Resource Utilisation','','Client service staff hours per head',0,0,0,0,0,0)
Insert into #OPERRES Values(503, 'N', @CompnyCode,@CompnyName,'Resource Utilisation','','Total resource hours per head',0,0,0,0,0,0)
Insert into #OPERRES Values(510, 'N', @CompnyCode,@CompnyName,'Resource Utilisation','','Client Service Leverage',0,0,0,0,0,0)
---verified............
Insert into #OPERRES Values(520, 'N', @CompnyCode,@CompnyName,'Average headcount','','Client service partners',0,0,0,0,0,0)
Insert into #OPERRES Values(530, 'N', @CompnyCode,@CompnyName,'Average headcount','','Practice support partners',0,0,0,0,0,0)
Insert into #OPERRES Values(540, 'N', @CompnyCode,@CompnyName,'Average headcount','','Client service staff',0,0,0,0,0,0)
Insert into #OPERRES Values(550, 'N', @CompnyCode,@CompnyName,'Average headcount','','Practice support staff',0,0,0,0,0,0)
Insert into #OPERRES Values(551, 'N', @CompnyCode,@CompnyName,'Average headcount','','Total headcount',0,0,0,0,0,0)
Insert into #OPERRES Values(552, 'N', @CompnyCode,@CompnyName,'Average headcount','','Contractors',0,0,0,0,0,0)
Insert into #OPERRES Values(553, 'N', @CompnyCode,@CompnyName,'Average headcount','','Total headcount including contractors',0,0,0,0,0,0)

Insert into #OPERRES Values(560, 'N', @CompnyCode,@CompnyName,'Closing headcount','','Client service partners',0,0,0,0,0,0)
Insert into #OPERRES Values(570, 'N', @CompnyCode,@CompnyName,'Closing headcount','','Practice support partners',0,0,0,0,0,0)
Insert into #OPERRES Values(580, 'N', @CompnyCode,@CompnyName,'Closing headcount','','Client service staff',0,0,0,0,0,0)
Insert into #OPERRES Values(590, 'N', @CompnyCode,@CompnyName,'Closing headcount','','Practice support staff',0,0,0,0,0,0)
Insert into #OPERRES Values(592, 'N', @CompnyCode,@CompnyName,'Closing headcount','','Total headcount',0,0,0,0,0,0)
Insert into #OPERRES Values(594, 'N', @CompnyCode,@CompnyName,'Closing headcount','','Contractors',0,0,0,0,0,0)
Insert into #OPERRES Values(596, 'N', @CompnyCode,@CompnyName,'Closing headcount','','Total headcount including contractors',0,0,0,0,0,0)

Insert into #OPERRES Values(600, 'N', @CompnyCode,@CompnyName,'Balance Sheet','','Work in progress',0,0,0,0,0,0)
Insert into #OPERRES Values(610, 'N', @CompnyCode,@CompnyName,'Balance Sheet','','Provisions',0,0,0,0,0,0)
Insert into #OPERRES Values(620, 'N', @CompnyCode,@CompnyName,'Balance Sheet','','Billed accounts receivable',0,0,0,0,0,0)
Insert into #OPERRES Values(630, 'N', @CompnyCode,@CompnyName,'Balance Sheet','','Reserve for billed',0,0,0,0,0,0)
Insert into #OPERRES Values(632, 'N', @CompnyCode,@CompnyName,'Balance Sheet','','Total',0,0,0,0,0,0)

Insert into #OPERRES Values(640, 'N', @CompnyCode,@CompnyName,'Balance Sheet','','Days Outstanding - WIP',0,0,0,0,0,0)
Insert into #OPERRES Values(650, 'N', @CompnyCode,@CompnyName,'Balance Sheet','','Days Outstanding - Billed accounts',0,0,0,0,0,0)
Insert into #OPERRES Values(660, 'N', @CompnyCode,@CompnyName,'Balance Sheet','','Collections',0,0,0,0,0,0)

Insert into #OPERRES Values(670, 'N', @CompnyCode,@CompnyName,'Other Metrics','','',0,0,0,0,0,0)

Insert into #OPERRES Values(680, 'N', @CompnyCode,@CompnyName,'Revenue.','','Net Revenue per partner ($ ''000)',0,0,0,0,0,0)

Insert into #OPERRES Values(690, 'N', @CompnyCode,@CompnyName,'Profitability.','','Resource cost per hour ($)',0,0,0,0,0,0)

Insert into #OPERRES Values(700, 'N', @CompnyCode,@CompnyName,'Cost Management.','','Client Service Costs per Capita (per Total CS) ($ ''000)',0,0,0,0,0,0)

Insert into #OPERRES Values(710, 'N', @CompnyCode,@CompnyName,'Host Chargeable Hours','','Client service partners',0,0,0,0,0,0)
Insert into #OPERRES Values(720, 'N', @CompnyCode,@CompnyName,'Host Chargeable Hours','','Client service staff',0,0,0,0,0,0)
Insert into #OPERRES Values(722, 'N', @CompnyCode,@CompnyName,'Host Chargeable Hours','','Total client service',0,0,0,0,0,0)

Insert into #OPERRES Values(730, 'N', @CompnyCode,@CompnyName,'Host Chargeable Hours','','Client service partners',0,0,0,0,0,0)
Insert into #OPERRES Values(740, 'N', @CompnyCode,@CompnyName,'Host Chargeable Hours','','Client service staff',0,0,0,0,0,0)
Insert into #OPERRES Values(750, 'N', @CompnyCode,@CompnyName,'Host Chargeable Hours','','Client service partners hours with credit',0,0,0,0,0,0)
Insert into #OPERRES Values(760, 'N', @CompnyCode,@CompnyName,'Host Chargeable Hours','','Client service staff hours with credit',0,0,0,0,0,0)


Insert into #OPERRES Values(770, 'N', @CompnyCode,@CompnyName,'Home Available Hours','','Client service partners',0,0,0,0,0,0)
Insert into #OPERRES Values(780, 'N', @CompnyCode,@CompnyName,'Home Available Hours','','Client service staff',0,0,0,0,0,0)
Insert into #OPERRES Values(782, 'N', @CompnyCode,@CompnyName,'Home Available Hours','','Total client service',0,0,0,0,0,0)

Insert into #OPERRES Values(789, 'N', @CompnyCode,@CompnyName,'Memo','','',0,0,0,0,0,0)

Insert into #OPERRES Values(790, 'N', @CompnyCode,@CompnyName,'Net Revenue','','Final billings',0,0,0,0,0,0)
Insert into #OPERRES Values(800, 'N', @CompnyCode,@CompnyName,'Net Revenue','','Net WIP movement',0,0,0,0,0,0)
Insert into #OPERRES Values(810, 'N', @CompnyCode,@CompnyName,'Net Revenue','','Loans to other OUs',0,0,0,0,0,0)
Insert into #OPERRES Values(820, 'N', @CompnyCode,@CompnyName,'Net Revenue','','Borrows from other OUs',0,0,0,0,0,0)
Insert into #OPERRES Values(830, 'N', @CompnyCode,@CompnyName,'Net Revenue','','Subtotal',0,0,0,0,0,0)
Insert into #OPERRES Values(840, 'N', @CompnyCode,@CompnyName,'Net Revenue','','Bad debts recovered',0,0,0,0,0,0)
Insert into #OPERRES Values(850, 'N', @CompnyCode,@CompnyName,'Net Revenue','','Discounts & Bad debts w/o',0,0,0,0,0,0)
Insert into #OPERRES Values(860, 'N', @CompnyCode,@CompnyName,'Net Revenue','','Specific bad debts provision',0,0,0,0,0,0)
Insert into #OPERRES Values(870, 'N', @CompnyCode,@CompnyName,'Net Revenue','','General bad debts  provision',0,0,0,0,0,0)
Insert into #OPERRES Values(880, 'N', @CompnyCode,@CompnyName,'Net Revenue','','Provision against WIP- Specific',0,0,0,0,0,0)
Insert into #OPERRES Values(890, 'N', @CompnyCode,@CompnyName,'Net Revenue','','Provision against WIP- General',0,0,0,0,0,0)
Insert into #OPERRES Values(900, 'N', @CompnyCode,@CompnyName,'Net Revenue','','Disbursements write-off (or write-up)',0,0,0,0,0,0)
Insert into #OPERRES Values(910, 'N', @CompnyCode,@CompnyName,'Net Revenue','','Other direct costs',0,0,0,0,0,0)
Insert into #OPERRES Values(920, 'N', @CompnyCode,@CompnyName,'Net Revenue','','Net fee income',0,0,0,0,0,0)
Insert into #OPERRES Values(930, 'N', @CompnyCode,@CompnyName,'Net Revenue','','Other income',0,0,0,0,0,0)
Insert into #OPERRES Values(940, 'N', @CompnyCode,@CompnyName,'Net Revenue','','Less: Borrow / Loan markup',0,0,0,0,0,0)
Insert into #OPERRES Values(950, 'Y', @CompnyCode,@CompnyName,'Net Revenue','','Net fee income before Borrow / Loan mark up',0,0,0,0,0,0)


BEGIN ------- Start Actual AAAS ---
update #OPERRES set Actual= Actual+ (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) )  from [AAAS].[dbo].JDT1 T1 with(nolock)
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate) 
where 
Account between '61121100' and '61121600' and
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=1

update #OPERRES set Actual=Actual+ (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock) 
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate) AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61121700' and '61124900' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=20

update #OPERRES set Actual=Actual +(
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE ='ST804001' and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=40



update #OPERRES set Actual=Actual + (isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (1,10,20,40) ),0)) where CompanyCode=@CompnyCode and Rownum=50

update #OPERRES set Actual=Actual+ (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock) 
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate) AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61124200' and '61124600' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=70

update #OPERRES set Actual=Actual + (isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (50,60,70) ),0)) where CompanyCode=@CompnyCode and Rownum=80

update #OPERRES set Actual=Actual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST804002','ST805010') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=81

update #OPERRES set Actual=Actual + (
isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (80,81) ),0)
) where CompanyCode=@CompnyCode and Rownum=83


update #OPERRES set Actual=Actual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST805012','ST805013') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=90



update #OPERRES set Actual=Actual + (
isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (83) ),0)-isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (90) ),0)-isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (100) ),0)
) where CompanyCode=@CompnyCode and Rownum=110


update #OPERRES set Actual=Actual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST805015') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=120


update #OPERRES set Actual=Actual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST805016') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=130


update #OPERRES set Actual=Actual +(
isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (130,120) ),0)
) where CompanyCode=@CompnyCode and Rownum=140

update #OPERRES set Actual=Actual +(
isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (110,140) ),0)
) where CompanyCode=@CompnyCode and Rownum=150

update #OPERRES set Actual=Actual + (
(isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate) AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61121900' and '61122000' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@Year ),0)) + 
(Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST805012','ST805015') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0))
) 
where CompanyCode=@CompnyCode and Rownum=160


update #OPERRES set Actual=Actual + (
isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (150,160) ),0)
) where CompanyCode=@CompnyCode and Rownum=170


update #OPERRES set Actual=Actual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate) AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '71111100' and '71121300' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=180

update #OPERRES set Actual=Actual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '71131100' and '71131800' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=190


update #OPERRES set Actual=Actual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '71141100' and '71142120' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=200

update #OPERRES set Actual=Actual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '71149100' and '71191100' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=210


update #OPERRES set Actual=Actual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '71201100' and '71211140' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=220



update #OPERRES set Actual=Actual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '71221100' and '71221120' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=230


update #OPERRES set Actual=Actual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '72111100' and '72121400' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=240

update #OPERRES set Actual=Actual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '71181410' and '72151300' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=250

update #OPERRES set Actual=Actual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '72131100' and '72141300' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=260

update #OPERRES set Actual=Actual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '73112000' and '73111102' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=290

update #OPERRES set Actual=Actual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock) 
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
 where 
Account between '73131100' and '73132100' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=300

update #OPERRES set Actual=Actual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '73141100' and '73141100' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=310

update #OPERRES set Actual=Actual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '73151100' and '73151100' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=320

update #OPERRES set Actual=Actual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '73161100' and '81121100' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=340	

update #OPERRES set Actual=Actual + (
isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 180
and  350),0) 
) where CompanyCode=@CompnyCode and Rownum=360


update #OPERRES set Actual=Actual + (
isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (140,360)),0) 
) where CompanyCode=@CompnyCode and Rownum=370


update #OPERRES set Actual=Actual + (
isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (170)),0) -isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (360)),0) 
) where CompanyCode=@CompnyCode and Rownum=380

update #OPERRES set Actual=Actual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST806100,ST806101') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=520

update #OPERRES set Actual=Actual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST806300','ST806301') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=530

update #OPERRES set Actual=Actual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST806102','ST806103','ST806104','ST806105','ST806106','ST806107','ST806108','ST806109','ST806110','ST806111','ST806112','ST806113','ST806114','ST806115','ST806116','ST806117','ST806118','ST806119','ST806120') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=540

update #OPERRES set Actual=Actual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST806302','ST806303','ST806304','ST806305','ST806306','ST806307','ST806308','ST806309','ST806310','ST806311','ST806312','ST806313','ST806314','ST806315','ST806316','ST806317','ST806318','ST806319','ST806320') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=550

update #OPERRES set Actual=Actual + (isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (520,530,540,550) ),0)) where CompanyCode=@CompnyCode and Rownum=551


update #OPERRES set Actual=Actual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST806100,ST806101') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=560

update #OPERRES set Actual=Actual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST806300','ST806301') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=570

update #OPERRES set Actual=Actual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST806102','ST806103','ST806104','ST806105','ST806106','ST806107','ST806108','ST806109','ST806110','ST806111','ST806112','ST806113','ST806114','ST806115','ST806116','ST806117','ST806118','ST806119','ST806120') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=580

update #OPERRES set Actual=Actual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST806302','ST806303','ST806304','ST806305','ST806306','ST806307','ST806308','ST806309','ST806310','ST806311','ST806312','ST806313','ST806314','ST806315','ST806316','ST806317','ST806318','ST806319','ST806320') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=590

update #OPERRES set Actual=Actual + (isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (560,570,580,590) ),0)) where CompanyCode=@CompnyCode and Rownum=592

update #OPERRES set Actual=Actual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST808016','ST804018','ST804010') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=600

update #OPERRES set Actual=Actual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST804010','ST804015','ST804017') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=610

update #OPERRES set Actual=Actual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST804020','ST804021','ST804022','ST804023') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=620

update #OPERRES set Actual=Actual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST804026','ST804027','ST804028') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=630

update #OPERRES set Actual=Actual + (isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (600) ),0)-isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (610,620,630) ),0)) where CompanyCode=@CompnyCode and Rownum=632

update #OPERRES set Actual=Actual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST803001') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=660


update #OPERRES set Actual=Actual + (
case when isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (520,530) ),0)=0 then 0 
else (isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (83) ),0)/isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (520,530) ),0)) end
) 
where CompanyCode=@CompnyCode and Rownum=680	

update #OPERRES set Actual=Actual + (
case when isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (520,530) ),0)=0 then 0 
else (isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (180) ),0) /isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (520,530) ),0)) end
) 
where CompanyCode=@CompnyCode and Rownum=700

update #OPERRES set Actual=Actual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST801070','ST801075') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=710

update #OPERRES set Actual=Actual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST801071','ST801073','ST801076','ST801078') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=720

update #OPERRES set Actual=Actual + (isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (710,720) ),0)) where CompanyCode=@CompnyCode and Rownum=722

update #OPERRES set Actual=Actual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST801020','ST801025') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=730

update #OPERRES set Actual=Actual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST801021','ST801023','ST801026','ST801028') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=740

update #OPERRES set Actual=Actual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST800001') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=770

update #OPERRES set Actual=Actual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST800002') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=780

update #OPERRES set Actual=Actual + (isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (770,780) ),0)) where CompanyCode=@CompnyCode and Rownum=782

update #OPERRES set Actual=Actual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST804005') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=790


update #OPERRES set Actual=Actual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST804006') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=800

update #OPERRES set Actual=Actual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST805005') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0) + isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61122000' and '61122000' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=810


update #OPERRES set Actual=Actual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST805001') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0) + isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61121900' and '61121900' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=820


update #OPERRES set Actual=Actual + (
isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (790,800,810,820) ),0)
) where CompanyCode=@CompnyCode and Rownum=830

update #OPERRES set Actual=Actual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61124600' and '61124600' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=840


update #OPERRES set Actual=Actual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61124500' and '61124500' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=850

update #OPERRES set Actual=Actual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61124200' and '61124200' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=860

update #OPERRES set Actual=Actual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61124300' and '61124400' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=870

update #OPERRES set Actual=Actual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61124100' and '61124100' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=880

update #OPERRES set Actual=Actual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61123800' and '61123900' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=890

update #OPERRES set Actual=Actual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61123000' and '61123100' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=900

update #OPERRES set Actual=Actual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61122700' and '61124900' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=910

update #OPERRES set Actual=Actual + (
isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 830 and 910 ),0)
) where CompanyCode=@CompnyCode and Rownum=920

update #OPERRES set Actual=Actual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61121900' and '61122000' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=940


update #OPERRES set Actual=Actual + (isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (920,930) ),0)-isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (940) ),0) ) where CompanyCode=@CompnyCode and Rownum=950


----** 400
update #OPERRES set Actual=Actual + (
CASE WHEN isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 782 and 782 ),0)=0 then 0 else
(isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 83 and 83 ),0)/isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 782 and 782 ),0)) *1000
end 
) where CompanyCode=@CompnyCode and Rownum=400

update #OPERRES set Actual=Actual + (
CASE WHEN isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 1 and 1 ),0)=0 then 0 else
(isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 50 and 50 ),0)-isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 40 and 40 ),0))*100/ isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 1 and 1 ),0)
end 
) where CompanyCode=@CompnyCode and Rownum=410

update #OPERRES set Actual=Actual + (
CASE WHEN isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 83 and 83 ),0)=0 then 0 else
isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 110 and 110 ),0)*100/ isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 83 and 83 ),0)
end 
) where CompanyCode=@CompnyCode and Rownum=420


update #OPERRES set Actual=Actual + (
CASE WHEN isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 83 and 83 ),0)=0 then 0 else
isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 380 and 380 ),0)*100/ isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 83 and 83 ),0)
end 
) where CompanyCode=@CompnyCode and Rownum=430


update #OPERRES set Actual=Actual + (
CASE WHEN isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (520,540) ),0)=0 then 0 else
isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 380 and 380 ),0)*1000/ isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (520,540)),0)
end 
) where CompanyCode=@CompnyCode and Rownum=440



update #OPERRES set Actual=Actual + (
CASE WHEN isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (180) ),0)=0 then 0 else
isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 380 and 380 ),0)*100/ isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (180)),0)
end 
) where CompanyCode=@CompnyCode and Rownum=450



update #OPERRES set Actual=Actual + (
CASE WHEN isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (740,760) ),0)=0 then 0 else
isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (180)),0)*1000/ isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (740,760)),0)
end 
) where CompanyCode=@CompnyCode and Rownum=460


update #OPERRES set Actual=Actual + (
CASE WHEN isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (140) ),0)=0 then 0 else
isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (370)),0)*100/ isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (140)),0)
end 
) where CompanyCode=@CompnyCode and Rownum=470

update #OPERRES set Actual=Actual + (
CASE WHEN isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (770) ),0)=0 then 0 else
isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (730)),0)*100/ isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (770)),0)
end 
) where CompanyCode=@CompnyCode and Rownum=480

update #OPERRES set Actual=Actual + (
CASE WHEN isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (780) ),0)=0 then 0 else
isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (740)),0)*100/ isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (780)),0)
end 
) where CompanyCode=@CompnyCode and Rownum=490

update #OPERRES set Actual=Actual + (
CASE WHEN isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (782) ),0)=0 then 0 else
isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (730,740)),0)*100/ isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (782)),0)
end 
) where CompanyCode=@CompnyCode and Rownum=500

update #OPERRES set Actual=Actual + (
CASE WHEN isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (520) ),0)=0 then 0 else
isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (540)),0)*100/ isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (520)),0)
end 
) where CompanyCode=@CompnyCode and Rownum=510

update #OPERRES set Actual=Actual + (
CASE WHEN isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (722) ),0)=0 then 0 else
isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (90)),0)*100/ isnull((select sum(Actual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (722)),0)
end 
) where CompanyCode=@CompnyCode and Rownum=690


END ------- END Actual AAAS ---

BEGIN ------- Start PriorYr AAAS ---
update #OPERRES set PriorYr= PriorYr+ (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) )  from [AAAS].[dbo].JDT1 T1 with(nolock)
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)
where 
Account between '61121100' and '61121600' and
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=1

update #OPERRES set PriorYr=PriorYr+ (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock) 
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61121700' and '61124900' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=30

update #OPERRES set PriorYr=PriorYr +(
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE ='ST804001' and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STPRTIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=40



update #OPERRES set PriorYr=PriorYr + (isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (1,20,30,40) ),0)) where CompanyCode=@CompnyCode and Rownum=50

update #OPERRES set PriorYr=PriorYr+ (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock) 
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61124200' and '61124600' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=70

update #OPERRES set PriorYr=PriorYr + (isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (50,60,70) ),0)) where CompanyCode=@CompnyCode and Rownum=80

update #OPERRES set PriorYr=PriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST804002','ST805010') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STPRTIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=81

update #OPERRES set PriorYr=PriorYr + (
isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (80,81) ),0)
) where CompanyCode=@CompnyCode and Rownum=83


update #OPERRES set PriorYr=PriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST805012','ST805013') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STPRTIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=90



update #OPERRES set PriorYr=PriorYr + (
isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (83) ),0)-isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (90) ),0)-isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (100) ),0)
) where CompanyCode=@CompnyCode and Rownum=110


update #OPERRES set PriorYr=PriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST805015') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STPRTIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=120


update #OPERRES set PriorYr=PriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST805016') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STPRTIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=130


update #OPERRES set PriorYr=PriorYr +(
isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (130,120) ),0)
) where CompanyCode=@CompnyCode and Rownum=140

update #OPERRES set PriorYr=PriorYr +(
isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (110,140) ),0)
) where CompanyCode=@CompnyCode and Rownum=150

update #OPERRES set PriorYr=PriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61121900' and '61122000' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=160


update #OPERRES set PriorYr=PriorYr + (
isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (150,160) ),0)
) where CompanyCode=@CompnyCode and Rownum=170


update #OPERRES set PriorYr=PriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '71111100' and '71121300' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=180

update #OPERRES set PriorYr=PriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '71131100' and '71131800' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=190


update #OPERRES set PriorYr=PriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '71141100' and '71142120' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=200

update #OPERRES set PriorYr=PriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '71149100' and '71191100' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=210


update #OPERRES set PriorYr=PriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '71201100' and '71211140' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=220



update #OPERRES set PriorYr=PriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '71221100' and '71221120' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=230


update #OPERRES set PriorYr=PriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '72111100' and '72121400' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=240

update #OPERRES set PriorYr=PriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '71181410' and '72151300' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=250

update #OPERRES set PriorYr=PriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '72131100' and '72141300' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=260

update #OPERRES set PriorYr=PriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '73112000' and '73111102' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=290

update #OPERRES set PriorYr=PriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock) 
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
 where 
Account between '73131100' and '73132100' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=300

update #OPERRES set PriorYr=PriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '73141100' and '73141100' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=310

update #OPERRES set PriorYr=PriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '73151100' and '73151100' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=320

update #OPERRES set PriorYr=PriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '73161100' and '81121100' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=340	

update #OPERRES set PriorYr=PriorYr + (
isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 180
and  350),0) 
) where CompanyCode=@CompnyCode and Rownum=360


update #OPERRES set PriorYr=PriorYr + (
isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (140,360)),0) 
) where CompanyCode=@CompnyCode and Rownum=370


update #OPERRES set PriorYr=PriorYr + (
isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (170)),0) -isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (360)),0) 
) where CompanyCode=@CompnyCode and Rownum=380

update #OPERRES set PriorYr=PriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST806100,ST806101') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STPRTIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=520

update #OPERRES set PriorYr=PriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST806300','ST806301') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STPRTIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=530

update #OPERRES set PriorYr=PriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST806102','ST806103','ST806104','ST806105','ST806106','ST806107','ST806108','ST806109','ST806110','ST806111','ST806112','ST806113','ST806114','ST806115','ST806116','ST806117','ST806118','ST806119','ST806120') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STPRTIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=540

update #OPERRES set PriorYr=PriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST806302','ST806303','ST806304','ST806305','ST806306','ST806307','ST806308','ST806309','ST806310','ST806311','ST806312','ST806313','ST806314','ST806315','ST806316','ST806317','ST806318','ST806319','ST806320') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STPRTIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=550

update #OPERRES set PriorYr=PriorYr + (isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (520,530,540,550) ),0)) where CompanyCode=@CompnyCode and Rownum=551


update #OPERRES set PriorYr=PriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST806100,ST806101') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STPRTIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=560

update #OPERRES set PriorYr=PriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST806300','ST806301') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STPRTIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=570

update #OPERRES set PriorYr=PriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST806102','ST806103','ST806104','ST806105','ST806106','ST806107','ST806108','ST806109','ST806110','ST806111','ST806112','ST806113','ST806114','ST806115','ST806116','ST806117','ST806118','ST806119','ST806120') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STPRTIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=580

update #OPERRES set PriorYr=PriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST806302','ST806303','ST806304','ST806305','ST806306','ST806307','ST806308','ST806309','ST806310','ST806311','ST806312','ST806313','ST806314','ST806315','ST806316','ST806317','ST806318','ST806319','ST806320') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STPRTIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=590

update #OPERRES set PriorYr=PriorYr + (isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (560,570,580,590) ),0)) where CompanyCode=@CompnyCode and Rownum=592

update #OPERRES set PriorYr=PriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST804016','ST804018','ST804010') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STPRTIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=600

update #OPERRES set PriorYr=PriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST804010','ST804015','ST804017') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STPRTIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=610

update #OPERRES set PriorYr=PriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST804020','ST804021','ST804022','ST804023') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STPRTIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=620

update #OPERRES set PriorYr=PriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST804026','ST804027','ST804028') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STPRTIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=630

update #OPERRES set PriorYr=PriorYr + (isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (600) ),0)-isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (610,620,630) ),0)) where CompanyCode=@CompnyCode and Rownum=632

update #OPERRES set PriorYr=PriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST803001') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STPRTIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=660


update #OPERRES set PriorYr=PriorYr + (
case when isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (520,530) ),0)=0 then 0 
else (1/isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (520,530) ),0)) end
) 
where CompanyCode=@CompnyCode and Rownum=680

update #OPERRES set PriorYr=PriorYr + (
case when isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (520,540) ),0)=0 then 0 
else (isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (180) ),0) /isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (520,534) ),0)) end
) 
where CompanyCode=@CompnyCode and Rownum=700

update #OPERRES set PriorYr=PriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST801070','ST801075') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STPRTIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=710

update #OPERRES set PriorYr=PriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST801071','ST801073','ST801076','ST801078') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STPRTIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=720

update #OPERRES set PriorYr=PriorYr + (isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (710,720) ),0)) where CompanyCode=@CompnyCode and Rownum=722

update #OPERRES set PriorYr=PriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST801020','ST801025') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STPRTIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=730

update #OPERRES set PriorYr=PriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST801021','ST801023','ST801026','ST801028') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STPRTIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=740

update #OPERRES set PriorYr=PriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST800001') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STPRTIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=770

update #OPERRES set PriorYr=PriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST800002') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STPRTIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=780

update #OPERRES set PriorYr=PriorYr + (isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (770,780) ),0)) where CompanyCode=@CompnyCode and Rownum=782

update #OPERRES set PriorYr=PriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST804005') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STPRTIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=790


update #OPERRES set PriorYr=PriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST804006') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STPRTIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=800

update #OPERRES set PriorYr=PriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST805005') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STPRTIORYR),0) + isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61122000' and '61122000' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=810


update #OPERRES set PriorYr=PriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST805001') and 
Cast(Right(U_AB_PERIOD,2) as Int)=@STMONTH and Left(U_AB_PERIOD,4)=@STPRTIORYR),0) + isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61121900' and '61121900' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=820


update #OPERRES set PriorYr=PriorYr + (
isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (790,800,810,820) ),0)
) where CompanyCode=@CompnyCode and Rownum=830

update #OPERRES set PriorYr=PriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61124600' and '61124600' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=840


update #OPERRES set PriorYr=PriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61124500' and '61124500' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=850

update #OPERRES set PriorYr=PriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61124200' and '61124200' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=860

update #OPERRES set PriorYr=PriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61124300' and '61124400' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=870

update #OPERRES set PriorYr=PriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61124100' and '61124100' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=880

update #OPERRES set PriorYr=PriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61123800' and '61123900' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=890

update #OPERRES set PriorYr=PriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61123000' and '61123100' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=900

update #OPERRES set PriorYr=PriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61122700' and '61124900' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=910

update #OPERRES set PriorYr=PriorYr + (
isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 830 and 910 ),0)
) where CompanyCode=@CompnyCode and Rownum=920

update #OPERRES set PriorYr=PriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61121900' and '61122000' and 
OO2.PrcCode = @LOS and Month(RefDate)=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=940


update #OPERRES set PriorYr=PriorYr + (isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (920,930) ),0)-isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (940) ),0) ) where CompanyCode=@CompnyCode and Rownum=950


----** 400
update #OPERRES set PriorYr=PriorYr + (
CASE WHEN isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 782 and 782 ),0)=0 then 0 else
(1/isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 782 and 782 ),0)) *1000
end 
) where CompanyCode=@CompnyCode and Rownum=400

update #OPERRES set PriorYr=PriorYr + (
CASE WHEN isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 1 and 1 ),0)=0 then 0 else
(isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 50 and 50 ),0)-isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 40 and 40 ),0))*100/ isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 1 and 1 ),0)
end 
) where CompanyCode=@CompnyCode and Rownum=410

update #OPERRES set PriorYr=PriorYr + (
CASE WHEN isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 83 and 83 ),0)=0 then 0 else
isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 110 and 110 ),0)*100/ isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 83 and 83 ),0)
end 
) where CompanyCode=@CompnyCode and Rownum=420


update #OPERRES set PriorYr=PriorYr + (
CASE WHEN isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 83 and 83 ),0)=0 then 0 else
isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 80 and 80 ),0)*100/ isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 83 and 83 ),0)
end 
) where CompanyCode=@CompnyCode and Rownum=430


update #OPERRES set PriorYr=PriorYr + (
CASE WHEN isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (520,540) ),0)=0 then 0 else
isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 380 and 380 ),0)*1000/ isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (520,540)),0)
end 
) where CompanyCode=@CompnyCode and Rownum=440



update #OPERRES set PriorYr=PriorYr + (
CASE WHEN isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (180) ),0)=0 then 0 else
isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 380 and 380 ),0)*100/ isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (180)),0)
end 
) where CompanyCode=@CompnyCode and Rownum=450



update #OPERRES set PriorYr=PriorYr + (
CASE WHEN isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (740,760) ),0)=0 then 0 else
isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (180)),0)*1000/ isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (740,760)),0)
end 
) where CompanyCode=@CompnyCode and Rownum=460


update #OPERRES set PriorYr=PriorYr + (
CASE WHEN isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (140) ),0)=0 then 0 else
isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (370)),0)*100/ isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (140)),0)
end 
) where CompanyCode=@CompnyCode and Rownum=470

update #OPERRES set PriorYr=PriorYr + (
CASE WHEN isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (770) ),0)=0 then 0 else
isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (730)),0)*100/ isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (770)),0)
end 
) where CompanyCode=@CompnyCode and Rownum=480

update #OPERRES set PriorYr=PriorYr + (
CASE WHEN isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (780) ),0)=0 then 0 else
isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (740)),0)*100/ isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (780)),0)
end 
) where CompanyCode=@CompnyCode and Rownum=490

update #OPERRES set PriorYr=PriorYr + (
CASE WHEN isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (730,740) ),0)=0 then 0 else
isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (782)),0)*100/ isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (730,740)),0)
end 
) where CompanyCode=@CompnyCode and Rownum=500

update #OPERRES set PriorYr=PriorYr + (
CASE WHEN isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (520) ),0)=0 then 0 else
isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (540)),0)*100/ isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (520)),0)
end 
) where CompanyCode=@CompnyCode and Rownum=510

update #OPERRES set PriorYr=PriorYr + (
CASE WHEN isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (722) ),0)=0 then 0 else
isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (90)),0)*100/ isnull((select sum(PriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (722)),0)
end 
) where CompanyCode=@CompnyCode and Rownum=690


END ------- END PriorYr AAAS ---


BEGIN ------- Start YTDActual AAAS ---
update #OPERRES set YTDActual= YTDActual+ (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) )  from [AAAS].[dbo].JDT1 T1 with(nolock)
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)
where 
Account between '61121100' and '61121600' and
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=1

update #OPERRES set YTDActual=YTDActual+ (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock) 
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61121700' and '61124900' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=30

update #OPERRES set YTDActual=YTDActual +(
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE ='ST804001' and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=40



update #OPERRES set YTDActual=YTDActual + (isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (1,20,30,40) ),0)) where CompanyCode=@CompnyCode and Rownum=50

update #OPERRES set YTDActual=YTDActual+ (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock) 
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61124200' and '61124600' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=70

update #OPERRES set YTDActual=YTDActual + (isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (50,60,70) ),0)) where CompanyCode=@CompnyCode and Rownum=80

update #OPERRES set YTDActual=YTDActual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST804002','ST805010') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=81

update #OPERRES set YTDActual=YTDActual + (
isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (80,81) ),0)
) where CompanyCode=@CompnyCode and Rownum=83


update #OPERRES set YTDActual=YTDActual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST805012','ST805013') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=90



update #OPERRES set YTDActual=YTDActual + (
isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (83) ),0)-isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (90) ),0)-isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (100) ),0)
) where CompanyCode=@CompnyCode and Rownum=110


update #OPERRES set YTDActual=YTDActual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST805015') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=120


update #OPERRES set YTDActual=YTDActual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST805016') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=130


update #OPERRES set YTDActual=YTDActual +(
isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (130,120) ),0)
) where CompanyCode=@CompnyCode and Rownum=140

update #OPERRES set YTDActual=YTDActual +(
isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (110,140) ),0)
) where CompanyCode=@CompnyCode and Rownum=150

update #OPERRES set YTDActual=YTDActual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61121900' and '61122000' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=160


update #OPERRES set YTDActual=YTDActual + (
isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (150,160) ),0)
) where CompanyCode=@CompnyCode and Rownum=170


update #OPERRES set YTDActual=YTDActual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '71111100' and '71121300' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=180

update #OPERRES set YTDActual=YTDActual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '71131100' and '71131800' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=190


update #OPERRES set YTDActual=YTDActual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '71141100' and '71142120' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=200

update #OPERRES set YTDActual=YTDActual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '71149100' and '71191100' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=210


update #OPERRES set YTDActual=YTDActual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '71201100' and '71211140' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=220



update #OPERRES set YTDActual=YTDActual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '71221100' and '71221120' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=230


update #OPERRES set YTDActual=YTDActual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '72111100' and '72121400' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=240

update #OPERRES set YTDActual=YTDActual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '71181410' and '72151300' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=250

update #OPERRES set YTDActual=YTDActual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '72131100' and '72141300' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=260

update #OPERRES set YTDActual=YTDActual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '73112000' and '73111102' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=290

update #OPERRES set YTDActual=YTDActual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock) 
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
 where 
Account between '73131100' and '73132100' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=300

update #OPERRES set YTDActual=YTDActual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '73141100' and '73141100' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=310

update #OPERRES set YTDActual=YTDActual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '73151100' and '73151100' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=320

update #OPERRES set YTDActual=YTDActual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '73161100' and '81121100' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=340	

update #OPERRES set YTDActual=YTDActual + (
isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 180
and  350),0) 
) where CompanyCode=@CompnyCode and Rownum=360


update #OPERRES set YTDActual=YTDActual + (
isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (140,360)),0) 
) where CompanyCode=@CompnyCode and Rownum=370


update #OPERRES set YTDActual=YTDActual + (
isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (170)),0) -isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (360)),0) 
) where CompanyCode=@CompnyCode and Rownum=380

update #OPERRES set YTDActual=YTDActual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST806100,ST806101') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=520

update #OPERRES set YTDActual=YTDActual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST806300','ST806301') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=530

update #OPERRES set YTDActual=YTDActual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST806102','ST806103','ST806104','ST806105','ST806106','ST806107','ST806108','ST806109','ST806110','ST806111','ST806112','ST806113','ST806114','ST806115','ST806116','ST806117','ST806118','ST806119','ST806120') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=540

update #OPERRES set YTDActual=YTDActual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST806302','ST806303','ST806304','ST806305','ST806306','ST806307','ST806308','ST806309','ST806310','ST806311','ST806312','ST806313','ST806314','ST806315','ST806316','ST806317','ST806318','ST806319','ST806320') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=550

update #OPERRES set YTDActual=YTDActual + (isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (520,530,540,550) ),0)) where CompanyCode=@CompnyCode and Rownum=551


update #OPERRES set YTDActual=YTDActual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST806100,ST806101') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=560

update #OPERRES set YTDActual=YTDActual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST806300','ST806301') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=570

update #OPERRES set YTDActual=YTDActual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST806102','ST806103','ST806104','ST806105','ST806106','ST806107','ST806108','ST806109','ST806110','ST806111','ST806112','ST806113','ST806114','ST806115','ST806116','ST806117','ST806118','ST806119','ST806120') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=580

update #OPERRES set YTDActual=YTDActual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST806302','ST806303','ST806304','ST806305','ST806306','ST806307','ST806308','ST806309','ST806310','ST806311','ST806312','ST806313','ST806314','ST806315','ST806316','ST806317','ST806318','ST806319','ST806320') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=590

update #OPERRES set YTDActual=YTDActual + (isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (560,570,580,590) ),0)) where CompanyCode=@CompnyCode and Rownum=592

update #OPERRES set YTDActual=YTDActual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST804016','ST804018','ST804010') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=600

update #OPERRES set YTDActual=YTDActual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST804010','ST804015','ST804017') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=610

update #OPERRES set YTDActual=YTDActual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST804020','ST804021','ST804022','ST804023') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=620

update #OPERRES set YTDActual=YTDActual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST804026','ST804027','ST804028') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=630

update #OPERRES set YTDActual=YTDActual + (isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (600) ),0)-isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (610,620,630) ),0)) where CompanyCode=@CompnyCode and Rownum=632

update #OPERRES set YTDActual=YTDActual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST803001') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=660


update #OPERRES set YTDActual=YTDActual + (
case when isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (520,530) ),0)=0 then 0 
else (1/isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (520,530) ),0)) end
) 
where CompanyCode=@CompnyCode and Rownum=680

update #OPERRES set YTDActual=YTDActual + (
case when isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (520,540) ),0)=0 then 0 
else (isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (180) ),0) /isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (520,534) ),0)) end
) 
where CompanyCode=@CompnyCode and Rownum=700

update #OPERRES set YTDActual=YTDActual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST801070','ST801075') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=710

update #OPERRES set YTDActual=YTDActual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST801071','ST801073','ST801076','ST801078') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=720

update #OPERRES set YTDActual=YTDActual + (isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (710,720) ),0)) where CompanyCode=@CompnyCode and Rownum=722

update #OPERRES set YTDActual=YTDActual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST801020','ST801025') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=730

update #OPERRES set YTDActual=YTDActual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST801021','ST801023','ST801026','ST801028') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=740

update #OPERRES set YTDActual=YTDActual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST800001') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=770

update #OPERRES set YTDActual=YTDActual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST800002') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=780

update #OPERRES set YTDActual=YTDActual + (isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (770,780) ),0)) where CompanyCode=@CompnyCode and Rownum=782

update #OPERRES set YTDActual=YTDActual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST804005') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=790


update #OPERRES set YTDActual=YTDActual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST804006') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0)
) where CompanyCode=@CompnyCode and Rownum=800

update #OPERRES set YTDActual=YTDActual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST805005') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0) + isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61122000' and '61122000' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=810


update #OPERRES set YTDActual=YTDActual + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST805001') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STYEAR),0) + isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61121900' and '61121900' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=820


update #OPERRES set YTDActual=YTDActual + (
isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (790,800,810,820) ),0)
) where CompanyCode=@CompnyCode and Rownum=830

update #OPERRES set YTDActual=YTDActual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61124600' and '61124600' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=840


update #OPERRES set YTDActual=YTDActual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61124500' and '61124500' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=850

update #OPERRES set YTDActual=YTDActual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61124200' and '61124200' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=860

update #OPERRES set YTDActual=YTDActual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61124300' and '61124400' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=870

update #OPERRES set YTDActual=YTDActual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61124100' and '61124100' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=880

update #OPERRES set YTDActual=YTDActual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61123800' and '61123900' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=890

update #OPERRES set YTDActual=YTDActual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61123000' and '61123100' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=900

update #OPERRES set YTDActual=YTDActual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61122700' and '61124900' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=910

update #OPERRES set YTDActual=YTDActual + (
isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 830 and 910 ),0)
) where CompanyCode=@CompnyCode and Rownum=920

update #OPERRES set YTDActual=YTDActual + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61121900' and '61122000' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@Year ),0)
) where CompanyCode=@CompnyCode and Rownum=940


update #OPERRES set YTDActual=YTDActual + (isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (920,930) ),0)-isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (940) ),0) ) where CompanyCode=@CompnyCode and Rownum=950


----** 400
update #OPERRES set YTDActual=YTDActual + (
CASE WHEN isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 782 and 782 ),0)=0 then 0 else
(1/isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 782 and 782 ),0)) *1000
end 
) where CompanyCode=@CompnyCode and Rownum=400

update #OPERRES set YTDActual=YTDActual + (
CASE WHEN isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 1 and 1 ),0)=0 then 0 else
(isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 50 and 50 ),0)-isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 40 and 40 ),0))*100/ isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 1 and 1 ),0)
end 
) where CompanyCode=@CompnyCode and Rownum=410

update #OPERRES set YTDActual=YTDActual + (
CASE WHEN isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 83 and 83 ),0)=0 then 0 else
isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 110 and 110 ),0)*100/ isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 83 and 83 ),0)
end 
) where CompanyCode=@CompnyCode and Rownum=420


update #OPERRES set YTDActual=YTDActual + (
CASE WHEN isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 83 and 83 ),0)=0 then 0 else
isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 80 and 80 ),0)*100/ isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 83 and 83 ),0)
end 
) where CompanyCode=@CompnyCode and Rownum=430


update #OPERRES set YTDActual=YTDActual + (
CASE WHEN isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (520,540) ),0)=0 then 0 else
isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 380 and 380 ),0)*1000/ isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (520,540)),0)
end 
) where CompanyCode=@CompnyCode and Rownum=440



update #OPERRES set YTDActual=YTDActual + (
CASE WHEN isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (180) ),0)=0 then 0 else
isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 380 and 380 ),0)*100/ isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (180)),0)
end 
) where CompanyCode=@CompnyCode and Rownum=450



update #OPERRES set YTDActual=YTDActual + (
CASE WHEN isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (740,760) ),0)=0 then 0 else
isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (180)),0)*1000/ isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (740,760)),0)
end 
) where CompanyCode=@CompnyCode and Rownum=460


update #OPERRES set YTDActual=YTDActual + (
CASE WHEN isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (140) ),0)=0 then 0 else
isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (370)),0)*100/ isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (140)),0)
end 
) where CompanyCode=@CompnyCode and Rownum=470

update #OPERRES set YTDActual=YTDActual + (
CASE WHEN isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (770) ),0)=0 then 0 else
isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (730)),0)*100/ isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (770)),0)
end 
) where CompanyCode=@CompnyCode and Rownum=480

update #OPERRES set YTDActual=YTDActual + (
CASE WHEN isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (780) ),0)=0 then 0 else
isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (740)),0)*100/ isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (780)),0)
end 
) where CompanyCode=@CompnyCode and Rownum=490

update #OPERRES set YTDActual=YTDActual + (
CASE WHEN isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (730,740) ),0)=0 then 0 else
isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (782)),0)*100/ isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (730,740)),0)
end 
) where CompanyCode=@CompnyCode and Rownum=500

update #OPERRES set YTDActual=YTDActual + (
CASE WHEN isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (520) ),0)=0 then 0 else
isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (540)),0)*100/ isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (520)),0)
end 
) where CompanyCode=@CompnyCode and Rownum=510

update #OPERRES set YTDActual=YTDActual + (
CASE WHEN isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (722) ),0)=0 then 0 else
isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (90)),0)*100/ isnull((select sum(YTDActual) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (722)),0)
end 
) where CompanyCode=@CompnyCode and Rownum=690


END ------- END YTDActual AAAS ---


BEGIN ------- Start YTDPriorYr AAAS ---
update #OPERRES set YTDPriorYr= YTDPriorYr+ (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) )  from [AAAS].[dbo].JDT1 T1 with(nolock)
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)
where 
Account between '61121100' and '61121600' and
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=1

update #OPERRES set YTDPriorYr=YTDPriorYr+ (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock) 
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61121700' and '61124900' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=30

update #OPERRES set YTDPriorYr=YTDPriorYr +(
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE ='ST804001' and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STPRIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=40



update #OPERRES set YTDPriorYr=YTDPriorYr + (isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (1,20,30,40) ),0)) where CompanyCode=@CompnyCode and Rownum=50

update #OPERRES set YTDPriorYr=YTDPriorYr+ (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock) 
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61124200' and '61124600' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=70

update #OPERRES set YTDPriorYr=YTDPriorYr + (isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (50,60,70) ),0)) where CompanyCode=@CompnyCode and Rownum=80

update #OPERRES set YTDPriorYr=YTDPriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST804002','ST805010') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STPRIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=81

update #OPERRES set YTDPriorYr=YTDPriorYr + (
isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (80,81) ),0)
) where CompanyCode=@CompnyCode and Rownum=83


update #OPERRES set YTDPriorYr=YTDPriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST805012','ST805013') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STPRIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=90



update #OPERRES set YTDPriorYr=YTDPriorYr + (
isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (83) ),0)-isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (90) ),0)-isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (100) ),0)
) where CompanyCode=@CompnyCode and Rownum=110


update #OPERRES set YTDPriorYr=YTDPriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST805015') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STPRIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=120


update #OPERRES set YTDPriorYr=YTDPriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST805016') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STPRIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=130


update #OPERRES set YTDPriorYr=YTDPriorYr +(
isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (130,120) ),0)
) where CompanyCode=@CompnyCode and Rownum=140

update #OPERRES set YTDPriorYr=YTDPriorYr +(
isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (110,140) ),0)
) where CompanyCode=@CompnyCode and Rownum=150

update #OPERRES set YTDPriorYr=YTDPriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61121900' and '61122000' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=160


update #OPERRES set YTDPriorYr=YTDPriorYr + (
isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (150,160) ),0)
) where CompanyCode=@CompnyCode and Rownum=170


update #OPERRES set YTDPriorYr=YTDPriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '71111100' and '71121300' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=180

update #OPERRES set YTDPriorYr=YTDPriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '71131100' and '71131800' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=190


update #OPERRES set YTDPriorYr=YTDPriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '71141100' and '71142120' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=200

update #OPERRES set YTDPriorYr=YTDPriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '71149100' and '71191100' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=210


update #OPERRES set YTDPriorYr=YTDPriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '71201100' and '71211140' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=220



update #OPERRES set YTDPriorYr=YTDPriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '71221100' and '71221120' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=230


update #OPERRES set YTDPriorYr=YTDPriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '72111100' and '72121400' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=240

update #OPERRES set YTDPriorYr=YTDPriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '71181410' and '72151300' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=250

update #OPERRES set YTDPriorYr=YTDPriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '72131100' and '72141300' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=260

update #OPERRES set YTDPriorYr=YTDPriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '73112000' and '73111102' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=290

update #OPERRES set YTDPriorYr=YTDPriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock) 
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
 where 
Account between '73131100' and '73132100' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=300

update #OPERRES set YTDPriorYr=YTDPriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '73141100' and '73141100' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=310

update #OPERRES set YTDPriorYr=YTDPriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '73151100' and '73151100' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=320

update #OPERRES set YTDPriorYr=YTDPriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '73161100' and '81121100' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=340	

update #OPERRES set YTDPriorYr=YTDPriorYr + (
isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 180
and  350),0) 
) where CompanyCode=@CompnyCode and Rownum=360


update #OPERRES set YTDPriorYr=YTDPriorYr + (
isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (140,360)),0) 
) where CompanyCode=@CompnyCode and Rownum=370


update #OPERRES set YTDPriorYr=YTDPriorYr + (
isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (170)),0) -isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (360)),0) 
) where CompanyCode=@CompnyCode and Rownum=380

update #OPERRES set YTDPriorYr=YTDPriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST806100,ST806101') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STPRIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=520

update #OPERRES set YTDPriorYr=YTDPriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST806300','ST806301') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STPRIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=530

update #OPERRES set YTDPriorYr=YTDPriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST806102','ST806103','ST806104','ST806105','ST806106','ST806107','ST806108','ST806109','ST806110','ST806111','ST806112','ST806113','ST806114','ST806115','ST806116','ST806117','ST806118','ST806119','ST806120') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STPRIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=540

update #OPERRES set YTDPriorYr=YTDPriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST806302','ST806303','ST806304','ST806305','ST806306','ST806307','ST806308','ST806309','ST806310','ST806311','ST806312','ST806313','ST806314','ST806315','ST806316','ST806317','ST806318','ST806319','ST806320') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STPRIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=550

update #OPERRES set YTDPriorYr=YTDPriorYr + (isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (520,530,540,550) ),0)) where CompanyCode=@CompnyCode and Rownum=551


update #OPERRES set YTDPriorYr=YTDPriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST806100,ST806101') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STPRIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=560

update #OPERRES set YTDPriorYr=YTDPriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST806300','ST806301') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STPRIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=570

update #OPERRES set YTDPriorYr=YTDPriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST806102','ST806103','ST806104','ST806105','ST806106','ST806107','ST806108','ST806109','ST806110','ST806111','ST806112','ST806113','ST806114','ST806115','ST806116','ST806117','ST806118','ST806119','ST806120') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STPRIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=580

update #OPERRES set YTDPriorYr=YTDPriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST806302','ST806303','ST806304','ST806305','ST806306','ST806307','ST806308','ST806309','ST806310','ST806311','ST806312','ST806313','ST806314','ST806315','ST806316','ST806317','ST806318','ST806319','ST806320') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STPRIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=590

update #OPERRES set YTDPriorYr=YTDPriorYr + (isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (560,570,580,590) ),0)) where CompanyCode=@CompnyCode and Rownum=592

update #OPERRES set YTDPriorYr=YTDPriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST804016','ST804018','ST804010') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STPRIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=600

update #OPERRES set YTDPriorYr=YTDPriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST804010','ST804015','ST804017') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STPRIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=610

update #OPERRES set YTDPriorYr=YTDPriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST804020','ST804021','ST804022','ST804023') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STPRIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=620

update #OPERRES set YTDPriorYr=YTDPriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST804026','ST804027','ST804028') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STPRIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=630

update #OPERRES set YTDPriorYr=YTDPriorYr + (isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (600) ),0)-isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (610,620,630) ),0)) where CompanyCode=@CompnyCode and Rownum=632

update #OPERRES set YTDPriorYr=YTDPriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST803001') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STPRIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=660


update #OPERRES set YTDPriorYr=YTDPriorYr + (
case when isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (520,530) ),0)=0 then 0 
else (1/isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (520,530) ),0)) end
) 
where CompanyCode=@CompnyCode and Rownum=680

update #OPERRES set YTDPriorYr=YTDPriorYr + (
case when isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (520,540) ),0)=0 then 0 
else (isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (180) ),0) /isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (520,534) ),0)) end
) 
where CompanyCode=@CompnyCode and Rownum=700

update #OPERRES set YTDPriorYr=YTDPriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST801070','ST801075') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STPRIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=710

update #OPERRES set YTDPriorYr=YTDPriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST801071','ST801073','ST801076','ST801078') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STPRIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=720

update #OPERRES set YTDPriorYr=YTDPriorYr + (isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (710,720) ),0)) where CompanyCode=@CompnyCode and Rownum=722

update #OPERRES set YTDPriorYr=YTDPriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST801020','ST801025') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STPRIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=730

update #OPERRES set YTDPriorYr=YTDPriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST801021','ST801023','ST801026','ST801028') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STPRIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=740

update #OPERRES set YTDPriorYr=YTDPriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST800001') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STPRIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=770

update #OPERRES set YTDPriorYr=YTDPriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST800002') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STPRIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=780

update #OPERRES set YTDPriorYr=YTDPriorYr + (isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (770,780) ),0)) where CompanyCode=@CompnyCode and Rownum=782

update #OPERRES set YTDPriorYr=YTDPriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST804005') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STPRIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=790


update #OPERRES set YTDPriorYr=YTDPriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST804006') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STPRIORYR),0)
) where CompanyCode=@CompnyCode and Rownum=800

update #OPERRES set YTDPriorYr=YTDPriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST805005') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STPRIORYR),0) + isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61122000' and '61122000' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=810


update #OPERRES set YTDPriorYr=YTDPriorYr + (
Isnull((select sum(isnull(U_AB_AMOUNT,0)) from 
[AAAS].[dbo].[@AB_STATITISTICSDATA]  with(nolock)
where U_AB_OPER_UNIT=@LOS  and U_AB_GLCODE  in ('ST805001') and 
Cast(Right(U_AB_PERIOD,2) as Int)<=@STMONTH and Left(U_AB_PERIOD,4)=@STPRIORYR),0) + isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61121900' and '61121900' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=820


update #OPERRES set YTDPriorYr=YTDPriorYr + (
isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (790,800,810,820) ),0)
) where CompanyCode=@CompnyCode and Rownum=830

update #OPERRES set YTDPriorYr=YTDPriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61124600' and '61124600' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=840


update #OPERRES set YTDPriorYr=YTDPriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61124500' and '61124500' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=850

update #OPERRES set YTDPriorYr=YTDPriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61124200' and '61124200' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=860

update #OPERRES set YTDPriorYr=YTDPriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61124300' and '61124400' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=870

update #OPERRES set YTDPriorYr=YTDPriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61124100' and '61124100' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=880

update #OPERRES set YTDPriorYr=YTDPriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61123800' and '61123900' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=890

update #OPERRES set YTDPriorYr=YTDPriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61123000' and '61123100' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=900

update #OPERRES set YTDPriorYr=YTDPriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61122700' and '61124900' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=910

update #OPERRES set YTDPriorYr=YTDPriorYr + (
isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 830 and 910 ),0)
) where CompanyCode=@CompnyCode and Rownum=920

update #OPERRES set YTDPriorYr=YTDPriorYr + (
isnull((select Sum((isnull(Debit,0)-isnull(Credit,0))* (OO2.PrcAmount/OO2.OcrTotal) ) from [AAAS].[dbo].JDT1 T1 with(nolock)  
LEFT JOIN [AAAS].[dbo].OOCR OO1  ON OO1.[OcrCode] = T1.OcrCode3 
LEFT JOIN [AAAS].[dbo].OCR1 OO2 ON OO1.[OcrCode] = OO2.[OcrCode] AND (OO2.ValidFrom <= T1.RefDate and Isnull(OO2.ValidTo,'9999-12-31') >= T1.RefDate)   
where 
Account between '61121900' and '61122000' and 
OO2.PrcCode = @LOS and Month(RefDate)<=@Month and Year(RefDate)=@PriorYr ),0)
) where CompanyCode=@CompnyCode and Rownum=940


update #OPERRES set YTDPriorYr=YTDPriorYr + (isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (920,930) ),0)-isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (940) ),0) ) where CompanyCode=@CompnyCode and Rownum=950


----** 400
update #OPERRES set YTDPriorYr=YTDPriorYr + (
CASE WHEN isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 782 and 782 ),0)=0 then 0 else
(1/isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 782 and 782 ),0)) *1000
end 
) where CompanyCode=@CompnyCode and Rownum=400

update #OPERRES set YTDPriorYr=YTDPriorYr + (
CASE WHEN isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 1 and 1 ),0)=0 then 0 else
(isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 50 and 50 ),0)-isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 40 and 40 ),0))*100/ isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 1 and 1 ),0)
end 
) where CompanyCode=@CompnyCode and Rownum=410

update #OPERRES set YTDPriorYr=YTDPriorYr + (
CASE WHEN isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 83 and 83 ),0)=0 then 0 else
isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 110 and 110 ),0)*100/ isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 83 and 83 ),0)
end 
) where CompanyCode=@CompnyCode and Rownum=420


update #OPERRES set YTDPriorYr=YTDPriorYr + (
CASE WHEN isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 83 and 83 ),0)=0 then 0 else
isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 80 and 80 ),0)*100/ isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 83 and 83 ),0)
end 
) where CompanyCode=@CompnyCode and Rownum=430


update #OPERRES set YTDPriorYr=YTDPriorYr + (
CASE WHEN isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (520,540) ),0)=0 then 0 else
isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 380 and 380 ),0)*1000/ isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (520,540)),0)
end 
) where CompanyCode=@CompnyCode and Rownum=440



update #OPERRES set YTDPriorYr=YTDPriorYr + (
CASE WHEN isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (180) ),0)=0 then 0 else
isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum between 380 and 380 ),0)*100/ isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (180)),0)
end 
) where CompanyCode=@CompnyCode and Rownum=450



update #OPERRES set YTDPriorYr=YTDPriorYr + (
CASE WHEN isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (740,760) ),0)=0 then 0 else
isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (180)),0)*1000/ isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (740,760)),0)
end 
) where CompanyCode=@CompnyCode and Rownum=460


update #OPERRES set YTDPriorYr=YTDPriorYr + (
CASE WHEN isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (140) ),0)=0 then 0 else
isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (370)),0)*100/ isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (140)),0)
end 
) where CompanyCode=@CompnyCode and Rownum=470

update #OPERRES set YTDPriorYr=YTDPriorYr + (
CASE WHEN isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (770) ),0)=0 then 0 else
isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (730)),0)*100/ isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (770)),0)
end 
) where CompanyCode=@CompnyCode and Rownum=480

update #OPERRES set YTDPriorYr=YTDPriorYr + (
CASE WHEN isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (780) ),0)=0 then 0 else
isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (740)),0)*100/ isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (780)),0)
end 
) where CompanyCode=@CompnyCode and Rownum=490

update #OPERRES set YTDPriorYr=YTDPriorYr + (
CASE WHEN isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (730,740) ),0)=0 then 0 else
isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (782)),0)*100/ isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (730,740)),0)
end 
) where CompanyCode=@CompnyCode and Rownum=500

update #OPERRES set YTDPriorYr=YTDPriorYr + (
CASE WHEN isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (520) ),0)=0 then 0 else
isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (540)),0)*100/ isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (520)),0)
end 
) where CompanyCode=@CompnyCode and Rownum=510

update #OPERRES set YTDPriorYr=YTDPriorYr + (
CASE WHEN isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (722) ),0)=0 then 0 else
isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (90)),0)*100/ isnull((select sum(YTDPriorYr) from #OPERRES where CompanyCode=@CompnyCode and Rownum in (722)),0)
end 
) where CompanyCode=@CompnyCode and Rownum=690


END ------- END YTDPriorYr AAAS ---

END---[AAAS]-END
---




Select 
Rownum,Bold,CompanyCode,CompanyName,Level1,Level2,Level3,case when Actual=0 then 0 else Actual/1000 end As Actual,
case when Budget=0 then 0 else Budget/1000 end As Budget,
case when PriorYr=0 then 0 else PriorYr/1000 end As PriorYr,
case when YTDActual=0 then 0 else YTDActual/1000 end As YTDActual,
case when YTDBudget=0 then 0 else YTDBudget/1000 end As YTDBudget,
case when YTDPriorYr=0 then 0 else YTDPriorYr/1000 end As YTDPriorYr

from #OPERRES where  Rownum >=0 order by Rownum 

drop Table  #OPERRES;

SET NOCOUNT OFF;
END
