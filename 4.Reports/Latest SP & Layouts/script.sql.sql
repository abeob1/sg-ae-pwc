USE [PWCL]
GO
/****** Object:  StoredProcedure [dbo].[@AE_SP002_InsertintoBudgetTable]    Script Date: 22/08/2017 11:41:16 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO




CREATE procedure [dbo].[@AE_SP002_InsertintoBudgetTable]

@BudgetType as varchar(30),
@BudgetName as varchar(200),
@BudgetPeriod as varchar(30),
@OUCode as varchar(30),
@BUCode as varchar(30),
@ProjectCode as varchar(30),
@BudgetAmount as decimal(19,3),
@Account as Varchar(35),
@Month1 as decimal(19,3),
@Month2 as decimal(19,3),
@Month3 as decimal(19,3),
@Month4 as decimal(19,3),
@Month5 as decimal(19,3),
@Month6 as decimal(19,3),
@Month7 as decimal(19,3),
@Month8 as decimal(19,3),
@Month9 as decimal(19,3),
@Month10 as decimal(19,3),
@Month11 as decimal(19,3),
@Month12 as decimal(19,3),
@sAmount as decimal(19,3)

as
begin
Declare @Docnum as integer
Declare @DocEntry as integer
--U_BalAmount

if @BudgetType = 'OU' 
begin
    select @Docnum = isnull(max(cast(DocEntry as integer)),0) + 1 from PWCL.. [@AB_OUBUDGET]
   insert into PWCL.. [@AB_OUBUDGET] ([DocEntry],[DocNum],[CreateDate], [Object]  ,[U_BudName],[U_Period],[U_Account],[U_Division],[U_BudAmount],[U_OUCode],[U_Month1]
           ,[U_Month2],[U_Month3],[U_Month4],[U_Month5],[U_Month6],[U_Month7],[U_Month8],[U_Month9],[U_Month10],[U_Month11],[U_Month12])
		   Values (@Docnum,@Docnum,getdate() , 'OU_Budget', @BudgetName,@BudgetPeriod, @Account, 'Equally',@BudgetAmount,@OUCode, @Month1,@Month2,@Month3,@Month4,@Month5,
		   @Month6,@Month7,@Month8,@Month9,@Month10,@Month11,@Month12)

   select @DocEntry = isnull(cast(DocEntry as integer),0) from PWCL.. [@AB_CONSOLBUDGET] where U_BudName = @BudgetName and U_Period = @BudgetPeriod and U_Account = @Account and U_OUCode = @OUCode 

   if @DocEntry is null
     begin
       select @Docnum = isnull(max(cast(DocEntry as integer)),0) + 1 from PWCL.. [@AB_CONSOLBUDGET]

       insert into PWCL.. [@AB_CONSOLBUDGET] ([DocEntry],[DocNum], [Object] ,[CreateDate],[U_BudName],[U_Period],[U_Account],[U_BudAmount],[U_OUCode],[U_BalAmount])
            Values(@Docnum,@Docnum,'CONSOLBUDGET', GETDATE(),@BudgetName,@BudgetPeriod, @Account, @BudgetAmount,@OUCode,@BudgetAmount )
     end
   else
     begin
       update PWCL.. [@AB_CONSOLBUDGET] set U_BudAmount += @BudgetAmount , U_BalAmount += @BudgetAmount where DocEntry = @DocEntry
     end

end



if @BudgetType = 'IF'
begin
   select @Docnum = isnull(max(cast(DocEntry as integer)),0) + 1 from PWCL.. [@AB_PROJECTBUDGET]
 if isnull(@BUCode,'')  = ''
 begin
    select @DocEntry = isnull(cast(DocEntry as integer),0) from PWCL.. [@AB_PROJECTBUDGET] where U_BudName = @BudgetName and U_Period = @BudgetPeriod and U_Account = @Account and U_PrjCode = @ProjectCode 
	if @DocEntry is null
         begin
            insert into PWCL.. [@AB_PROJECTBUDGET] ([DocEntry],[DocNum],[CreateDate], [Object]  ,[U_BudName],[U_Period],[U_Account],[U_Division],[U_BudAmount],[U_PrjCode],[U_BUCode],[U_BalAmount] )
		   Values (@Docnum,@Docnum,getdate(), 'PR_Budget' ,@BudgetName,@BudgetPeriod, @Account, 'Equally',@BudgetAmount,@ProjectCode, @BUCode, @BudgetAmount )
        end
      else
        begin
	        update PWCL.. [@AB_PROJECTBUDGET] set U_BudAmount += @BudgetAmount ,  U_BalAmount += @BudgetAmount  where DocEntry = @DocEntry
       end
 end
 else
 begin
   
   select @DocEntry = isnull(cast(DocEntry as integer),0) from PWCL.. [@AB_PROJECTBUDGET] where U_BudName = @BudgetName and U_Period = @BudgetPeriod and U_Account = @Account and U_BUCode = @BUCode
   if @DocEntry is null
         begin
            insert into PWCL.. [@AB_PROJECTBUDGET] ([DocEntry],[DocNum],[CreateDate], [Object]  ,[U_BudName],[U_Period],[U_Account],[U_Division],[U_BudAmount],[U_PrjCode],[U_BUCode],[U_BalAmount] )
		   Values (@Docnum,@Docnum,getdate(), 'PR_Budget' ,@BudgetName,@BudgetPeriod, @Account, 'Equally',@BudgetAmount,@ProjectCode, @BUCode, @BudgetAmount )
        end
      else
        begin
	        update PWCL.. [@AB_PROJECTBUDGET] set U_BudAmount += @BudgetAmount ,  U_BalAmount += @BudgetAmount  where DocEntry = @DocEntry
       end
  
 end

  
/*
   if isnull(@BUCode,'')  = ''
    begin
	    select @DocEntry = isnull(cast(DocEntry as integer),0) from PWCL.. [@AB_CONSOLBUDGET] where U_BudName = @BudgetName and U_Period = @BudgetPeriod and U_Account = @Account and U_PrjCode = @ProjectCode 
   
       if @DocEntry is null
         begin
           select @Docnum = isnull(max(cast(DocEntry as integer)),0) + 1 from PWCL.. [@AB_CONSOLBUDGET]

          insert into PWCL.. [@AB_CONSOLBUDGET] ([DocEntry],[DocNum],[CreateDate],[Object] ,[U_BudName],[U_Period],[U_Account],[U_BudAmount],[U_PrjCode] ,[U_PrjAmount],[U_BalAmount]   )
            Values(@Docnum,@Docnum,GETDATE(),'CONSOLBUDGET',@BudgetName,@BudgetPeriod, @Account, @BudgetAmount,@ProjectCode , @BudgetAmount, @BudgetAmount   )
        end
      else
        begin
	        update PWCL.. [@AB_CONSOLBUDGET] set U_BudAmount += @BudgetAmount, U_PrjCode  = @ProjectCode , U_PrjAmount  += @BudgetAmount, U_BalAmount += @BudgetAmount  where DocEntry = @DocEntry
       end
	end
	else
	Begin
	   select @DocEntry = isnull(cast(DocEntry as integer),0) from PWCL.. [@AB_CONSOLBUDGET] where U_BudName = @BudgetName and U_Period = @BudgetPeriod and U_Account = @Account and U_BUCode = @BUCode  
   
       if @DocEntry is null
         begin
           select @Docnum = isnull(max(cast(DocEntry as integer)),0) + 1 from PWCL.. [@AB_CONSOLBUDGET]

          insert into PWCL.. [@AB_CONSOLBUDGET] ([DocEntry],[DocNum],[CreateDate],[Object] ,[U_BudName],[U_Period],[U_Account],[U_BudAmount],[U_PrjCode] ,[U_PrjAmount],[U_BalAmount],[U_BUCode]    )
            Values(@Docnum,@Docnum,GETDATE(),'CONSOLBUDGET',@BudgetName,@BudgetPeriod, @Account, @BudgetAmount,@ProjectCode , @BudgetAmount, @BudgetAmount, @BUCode    )
        end
      else
        begin
	        update PWCL.. [@AB_CONSOLBUDGET] set U_BudAmount += @BudgetAmount, U_BuAmount  += @BudgetAmount, U_BalAmount += @BudgetAmount  where DocEntry = @DocEntry
       end

	end
	*/
  
end

end

GO
/****** Object:  StoredProcedure [dbo].[AB_OS]    Script Date: 22/08/2017 11:41:16 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>

-- exec  [dbo].[AB_OS] '2015-03-03','2015-03-31','52000'
-- =============================================
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


GO
/****** Object:  StoredProcedure [dbo].[AE_APPROVALGRID]    Script Date: 22/08/2017 11:41:16 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[AE_APPROVALGRID](@HOLDINGDB NVARCHAR(100),@APPTYPE NVARCHAR(100),@APPCODE NVARCHAR(100),@APPROVER NVARCHAR(100),@CREATOR NVARCHAR(100),@DOCUMENT NVARCHAR(10),@DRAFTNO NVARCHAR(10),@VENDOR NVARCHAR(100),@FROMDATE DATE,@TODATE DATE,@ENTITY NVARCHAR(100))
AS
BEGIN
--AE_APPROVALGRID 'ALL','ADM3','Goh Li Ling','','ALL','','','2016-01-01','2016-11-10',''

--DECLARE @APPTYPE NVARCHAR(100),@APPCODE NVARCHAR(100),@APPROVER NVARCHAR(100),@CREATOR NVARCHAR(100),@DOCUMENT NVARCHAR(10),@DRAFTNO NVARCHAR(10),@VENDOR NVARCHAR(100),@FROMDATE DATE,@TODATE DATE,@ENTITY NVARCHAR(100)
--DECLARE @HOLDINGDB NVARCHAR(100)

--SET @HOLDINGDB = 'PWCL'
--SET @APPROVER = 'Marcus HC Lam'
--SET @DOCUMENT = 'ALL'
--SET @ENTITY = ''
--SET @DRAFTNO = ''
--SET @CREATOR = '' 
--SET @VENDOR = ''
--SET @FROMDATE = '2015-01-01'
--SET @TODATE = '2016-12-31'
--SET @APPTYPE = 'ALL'
--SET @APPCODE = 'ADM5'

DECLARE @SQL NVARCHAR(MAX) ,@SQL1 NVARCHAR(MAX)

CREATE TABLE #ALLENTITYLIST(Name NVARCHAR(MAX))

CREATE TABLE #TABLE(ENTITY NVARCHAR(50),[DOCUMENT TYPE] NVARCHAR(100),[SELECT] NVARCHAR(2),[DRAFT NO] INTEGER,[DOCUMENT NO] NVARCHAR(50),
				    [CREATOR NAME] NVARCHAR(100),[POSTING DATE] DATE,[VENDOR NAME] NVARCHAR(100),[AMOUNT(BEFORE GST)] NUMERIC(18,3),[APPROVAL GRID] NVARCHAR(100),
					[APPROVALCODE] NVARCHAR(100),[USERCODE] NVARCHAR(100),[STATUS] NVARCHAR(100),REMARKS NVARCHAR(255),APPROVER NVARCHAR(100), 
					[REMARKS BY APPROVER] NVARCHAR(254),APPSEQ INTEGER,SEQ INTEGER,STEPCODE INTEGER,ACODE NVARCHAR(240),Astatus NVARCHAR(10))


CREATE TABLE #TABLEF(ENTITY NVARCHAR(50),[DOCUMENT TYPE] NVARCHAR(100),[SELECT] NVARCHAR(2),[DRAFT NO] INTEGER,[DOCUMENT NO] NVARCHAR(50),
				    [CREATOR NAME] NVARCHAR(100),[POSTING DATE] DATE,[VENDOR NAME] NVARCHAR(100),[AMOUNT(BEFORE GST)] NUMERIC(18,3),[APPROVAL GRID] NVARCHAR(100),
					[APPROVALCODE] NVARCHAR(100),[USERCODE] NVARCHAR(100),[STATUS] NVARCHAR(100),REMARKS NVARCHAR(255),APPROVER NVARCHAR(100), 
					[REMARKS BY APPROVER] NVARCHAR(254),APPSEQ INTEGER,SEQ INTEGER,STEPCODE INTEGER,ACODE NVARCHAR(240),Astatus NVARCHAR(10))

CREATE TABLE #BEFOREFINAL_UPDATED(ENTITY NVARCHAR(50),[ENTITY NAME] NVARCHAR(240),[DOCUMENT TYPE] NVARCHAR(100),[SELECT] NVARCHAR(2),[DRAFT NO] INTEGER,[DOCUMENT NO] NVARCHAR(50),
				    [CREATOR NAME] NVARCHAR(100),[POSTING DATE] DATE,[VENDOR NAME] NVARCHAR(100),[AMOUNT(BEFORE GST)] NUMERIC(18,3),[APPROVED BY] NVARCHAR(100),[APPROVAL GRID] NVARCHAR(100),
					[APPROVALCODE] NVARCHAR(100),[USERCODE] NVARCHAR(100),[APPROVAL TYPE] NVARCHAR(100),[STATUS] NVARCHAR(100),REMARKS NVARCHAR(255),
					[REASON FOR NOT APPROVING] NVARCHAR(254),SEQ INTEGER)

CREATE TABLE #FINAL(ENTITY NVARCHAR(50),[ENTITY NAME] NVARCHAR(240),[DOCUMENT TYPE] NVARCHAR(100),[SELECT] NVARCHAR(2),[DRAFT NO] INTEGER,[DOCUMENT NO] NVARCHAR(50),
				    [CREATOR NAME] NVARCHAR(100),[POSTING DATE] DATE,[VENDOR NAME] NVARCHAR(100),[AMOUNT(BEFORE GST)] NUMERIC(18,3),[APPROVED BY] NVARCHAR(100),[APPROVAL GRID] NVARCHAR(100),
					[APPROVALCODE] NVARCHAR(100),[USERCODE] NVARCHAR(100),[APPROVAL TYPE] NVARCHAR(100),[STATUS] NVARCHAR(100),REMARKS NVARCHAR(255),
					[REASON FOR NOT APPROVING] NVARCHAR(254),SEQ INTEGER)

CREATE TABLE #TMPCOMPANY(Code NVARCHAR(50),Name NVARCHAR(100),U_AB_COMPANYNAME NVARCHAR(100))


IF ISNULL(@ENTITY, '') <> ''
BEGIN
SET @SQL = 'SELECT A.DocEntry, CASE WHEN A.ObjType = ''22'' THEN ''PURCHASE ORDER'' WHEN A.ObjType = ''1470000113'' THEN ''PURCHASE REQUEST'' END [Document Type],
A.ObjType, A.DocEntry [Draft No],F.SeriesName + '' - '' + cast(A.DocNum as nvarchar(30)) [Document No],
B.U_NAME [Creator Name],A.DocDate [Posting Date],CardName [Vendor Name],A.Doctotal -A.VatSum [DocTotal] , A.U_AB_APPROVALCODE , c.WddCode ,E.U_NAME ,E.USERID ,
 E.USER_CODE , ''Pending'' [Status] , A.Comments [Remarks] , D.Status [AStatus] , 
 --ROW_NUMBER() OVER(partition by c.WddCode order by c.WddCode) AS AppSeq,
 ROW_NUMBER() OVER(partition by c.WddCode order by H.SortId) AS AppSeq,
 cast( cast(c.WddCode as nvarchar(20)) + cast(D.[StepCode] as nvarchar(20)) as integer) [StepCode],H.SortId
INTO #TmpInitialise
FROM ' + @Entity + '..ODRF A
INNER JOIN ' + @Entity + '..OUSR B ON B.USERID = A.UserSign
LEFT JOIN ' + @Entity + '..OWDD C ON A.DocEntry = C.DocEntry 
JOIN ' + @Entity + '..WDD1 D ON C.WddCode = D.WddCode 
JOIN ' + @Entity + '..WTM2 H ON H.WtmCode = C.WtmCode AND D.StepCode = H.WstCode
JOIN ' + @Entity + '..OUSR E ON E.USERID = D.UserID 
JOIN ' + @Entity + '..NNM1 F ON F.Series = A.Series 
LEFT OUTER JOIN ' + @HOLDINGDB + ' ..[@AB_APPROVALMATRIX] G ON G.U_ApprGridCode  =  A.U_AB_APPROVALCODE 		
WHERE DocStatus = ''O''
AND A.DocDate >= '''+ CAST(@FROMDATE AS VARCHAR(50))+''' AND A.DocDate <= '''+ CAST(@TODATE AS VARCHAR(50))+'''
AND A.ObjType IN  (22,1470000113) '

SET @SQL = @SQL + 'SELECT ''' + @ENTITY + ''' [Entity], [Document Type],'''' [Select], [Draft No],[Document No],[Creator Name],[Posting Date],[Vendor Name],DocTotal ,U_AB_APPROVALCODE,WddCode,U_NAME,
USERID,USER_CODE , Status , Remarks ,[AStatus],AppSeq,StepCode,SortId
into #TmpApproval
FROM #TmpInitialise
WHERE ObjType = (CASE WHEN ''' + @DOCUMENT + ''' = ''ALL'' THEN ObjType WHEN ''' + @DOCUMENT + ''' = ''PR'' THEN ''1470000113'' WHEN ''' + @DOCUMENT + ''' = ''PO'' THEN ''22'' END) 
AND [Creator Name] = (CASE WHEN ISNULL(''' + @CREATOR + ''','''') = '''' THEN  [Creator Name] ELSE ''' + @CREATOR + ''' END)
AND DocEntry = (CASE WHEN ISNULL(''' + @DRAFTNO + ''','''') = '''' THEN DocEntry ELSE ''' + @DRAFTNO + ''' END)
AND ISNULL([Vendor Name],'''') = (CASE WHEN ISNULL(''' + @VENDOR + ''','''') = '''' THEN ISNULL([Vendor Name],'''') ELSE ''' + @VENDOR + ''' END)

--select *  , dense_rank() OVER (PARTITION BY [draft no] ORDER BY stepcode ) [seq] into #TmpApprovalF 
select [Entity], [Document Type], [Select], [Draft No],[Document No],[Creator Name],[Posting Date],[Vendor Name],DocTotal ,U_AB_APPROVALCODE,WddCode,U_NAME,
USERID,USER_CODE , Status , Remarks ,[AStatus],AppSeq,StepCode,SortId [seq] into #TmpApprovalF
from #TmpApproval 
group by [Entity],[Select], [Document Type],  [Draft No],[Document No],[Creator Name],[Posting Date],[Vendor Name],DocTotal ,
U_AB_APPROVALCODE,WddCode,U_NAME,USERID,USER_CODE , Status , Remarks ,[AStatus],AppSeq,StepCode,SortId
order by [Draft No] , AppSeq 

select * into #TmpApprovalF1 from #TmpApprovalF
 
INSERT INTO #TABLE
select Entity,[Document Type],[Select],[Draft No],[Document No],[Creator Name],[Posting Date],[Vendor Name],DocTotal,U_AB_APPROVALCODE,
WddCode,U_NAME,[Status],[Remarks],
(select top 1 U_Name from #TmpApprovalF1 T0 where T0.[Seq] = T1.[Seq] -1 and T0.[Draft No] = T1.[Draft No]  and T0.AStatus = ''Y'' group by U_Name) as [Approvedby] 
,'''' [ApproverRemarks],AppSeq,Seq,StepCode,USER_CODE, AStatus
from #TmpApprovalF T1 -- where AStatus = ''W'' AND U_NAME =  ''' + @APPROVER + '''
order by [Draft No] , AppSeq 

select [DRAFT NO],seq, stepcode , case when AStatus in (''Y'',''N'') then ''1'' else ''0'' end [RC], AStatus into #TmpCR from #table 

delete from #TmpCR where stepcode in (select stepcode from #TmpCR where RC  = ''1'' )
'
set @SQL1 = '  INSERT INTO #TABLEF
select x.*  from(
select * from #TABLE WHERE SEQ=1 AND [DRAFT NO] NOT IN (select [DRAFT NO] from #TABLE WHERE SEQ=1 AND AStatus= ''Y'') UNION ALL
select * from #TABLE WHERE SEQ=2 AND [DRAFT NO] IN (select [DRAFT NO] from #TABLE WHERE SEQ=1 AND AStatus=''Y'') UNION ALL
select * from #TABLE WHERE SEQ=3 AND [DRAFT NO] IN (select [DRAFT NO] from #TABLE WHERE SEQ=2 AND AStatus=''Y'') UNION ALL
select * from #TABLE WHERE SEQ=4 AND [DRAFT NO] IN (select [DRAFT NO] from #TABLE WHERE SEQ=3 AND AStatus=''Y'') UNION ALL
select * from #TABLE WHERE SEQ=5 AND [DRAFT NO] IN (select [DRAFT NO] from #TABLE WHERE SEQ=4 AND AStatus=''Y'') UNION ALL
select * from #TABLE WHERE SEQ=6 AND [DRAFT NO] IN (select [DRAFT NO] from #TABLE WHERE SEQ=5 AND AStatus=''Y'') UNION ALL
select * from #TABLE WHERE SEQ=7 AND [DRAFT NO] IN (select [DRAFT NO] from #TABLE WHERE SEQ=6 AND AStatus=''Y'') UNION ALL
select * from #TABLE WHERE SEQ=8 AND [DRAFT NO] IN (select [DRAFT NO] from #TABLE WHERE SEQ=7 AND AStatus=''Y'') UNION ALL
select * from #TABLE WHERE SEQ=9 AND [DRAFT NO] IN (select [DRAFT NO] from #TABLE WHERE SEQ=8 AND AStatus=''Y'')) x --join #TmpCR y on x.stepcode = y.stepcode
 WHERE   X.ACODE = ''' + @APPCODE +  ''' AND X.STEPCODE IN ( SELECT STEPCODE FROM #TMPCR    ) order by x.[Draft No],x.[AppSeq]

delete from #TABLEF where StepCode in (select T0.StepCode from #TABLEF T0 where T0.AStatus in (''Y'',''N''))

Drop table #TmpApprovalF1
Drop table #TmpApprovalF
Drop table #TmpApproval
Drop table #TmpInitialise 
drop table #TmpCR
'
--SELECT @SQL + @SQL1

EXEC(@SQL + @SQL1)
END
ELSE IF ISNULL(@ENTITY,'') = ''
BEGIN

DECLARE @ENTITYNAME NVARCHAR(MAX)

EXEC(' INSERT INTO #ALLENTITYLIST
SELECT A.Name FROM ' + @HOLDINGDB + ' ..[@AB_COMPANYDATA] A
INNER JOIN [SBO-COMMON].dbo.SRGC B ON B.dbName = A.Name
')

DECLARE C1 CURSOR FOR
SELECT Name FROM #ALLENTITYLIST

OPEN C1;
FETCH NEXT FROM C1 INTO @ENTITYNAME
WHILE @@FETCH_STATUS = 0
BEGIN

SET @SQL = 'SELECT A.DocEntry, CASE WHEN A.ObjType = ''22'' THEN ''PURCHASE ORDER'' WHEN A.ObjType = ''1470000113'' THEN ''PURCHASE REQUEST'' END [Document Type],
A.ObjType, A.DocEntry [Draft No],F.SeriesName + '' - '' + cast(A.DocNum as nvarchar(30)) [Document No],
B.U_NAME [Creator Name],A.DocDate [Posting Date],CardName [Vendor Name],A.Doctotal -A.VatSum [DocTotal] , A.U_AB_APPROVALCODE , c.WddCode ,E.U_NAME ,E.USERID ,
 E.USER_CODE , ''Pending'' [Status] , A.Comments [Remarks] , D.Status [AStatus] , 
 --ROW_NUMBER() OVER(partition by c.WddCode order by c.WddCode) AS AppSeq,
 ROW_NUMBER() OVER(partition by c.WddCode order by H.SortId) AS AppSeq,
 cast( cast(c.WddCode as nvarchar(20)) + cast(D.[StepCode] as nvarchar(20)) as integer) [StepCode],H.SortId
INTO #TmpInitialise
FROM ' + @ENTITYNAME + '..ODRF A
INNER JOIN ' + @ENTITYNAME + '..OUSR B ON B.USERID = A.UserSign
LEFT JOIN ' + @ENTITYNAME + '..OWDD C ON A.DocEntry = C.DocEntry 
JOIN ' + @ENTITYNAME + '..WDD1 D ON C.WddCode = D.WddCode 
JOIN ' + @ENTITYNAME + '..WTM2 H ON H.WtmCode = C.WtmCode AND D.StepCode = H.WstCode
JOIN ' + @ENTITYNAME + '..OUSR E ON E.USERID = D.UserID 
JOIN ' + @ENTITYNAME + '..NNM1 F ON F.Series = A.Series 
LEFT OUTER JOIN ' + @HOLDINGDB + ' ..[@AB_APPROVALMATRIX] G ON G.U_ApprGridCode  =  A.U_AB_APPROVALCODE 
WHERE DocStatus = ''O''
AND A.DocDate >= '''+ CAST(@FROMDATE AS VARCHAR(50))+''' AND A.DocDate <= '''+ CAST(@TODATE AS VARCHAR(50))+'''
AND A.ObjType IN  (22,1470000113) 

SELECT ''' + @ENTITYNAME + ''' [Entity], [Document Type],'''' [Select], [Draft No],[Document No],[Creator Name],[Posting Date],[Vendor Name],DocTotal ,U_AB_APPROVALCODE,WddCode,U_NAME,
USERID,USER_CODE , Status , Remarks ,[AStatus],AppSeq,StepCode,SortId
into #TmpApproval
FROM #TmpInitialise
WHERE ObjType = (CASE WHEN ''' + @DOCUMENT + ''' = ''ALL'' THEN ObjType WHEN ''' + @DOCUMENT + ''' = ''PR'' THEN ''1470000113'' WHEN ''' + @DOCUMENT + ''' = ''PO'' THEN ''22'' END) 
AND [Creator Name] = (CASE WHEN ISNULL(''' + @CREATOR + ''','''') = '''' THEN  [Creator Name] ELSE ''' + @CREATOR + ''' END)
AND DocEntry = (CASE WHEN ISNULL(''' + @DRAFTNO + ''','''') = '''' THEN DocEntry ELSE ''' + @DRAFTNO + ''' END)
AND ISNULL([Vendor Name],'''') = (CASE WHEN ISNULL(''' + @VENDOR + ''','''') = '''' THEN ISNULL([Vendor Name],'''') ELSE ''' + @VENDOR + ''' END)

--select *  , dense_rank() OVER (PARTITION BY [draft no] ORDER BY stepcode ) [seq] into #TmpApprovalF 
select [Entity], [Document Type], [Select], [Draft No],[Document No],[Creator Name],[Posting Date],[Vendor Name],DocTotal ,U_AB_APPROVALCODE,WddCode,U_NAME,
USERID,USER_CODE , Status , Remarks ,[AStatus],AppSeq,StepCode,SortId [seq] into #TmpApprovalF
from #TmpApproval 
group by [Entity],[Select], [Document Type],  [Draft No],[Document No],[Creator Name],[Posting Date],[Vendor Name],DocTotal ,
U_AB_APPROVALCODE,WddCode,U_NAME,USERID,USER_CODE , Status , Remarks ,[AStatus],AppSeq,StepCode,SortId
order by [Draft No] , AppSeq 

select * into #TmpApprovalF1 from #TmpApprovalF'
set @SQL1 = '  INSERT INTO #TABLE
select Entity,[Document Type],[Select],[Draft No],[Document No],[Creator Name],[Posting Date],[Vendor Name],DocTotal,U_AB_APPROVALCODE,
WddCode,U_NAME,[Status],[Remarks],
(select top 1 U_Name from #TmpApprovalF1 T0 where T0.[Seq] = T1.[Seq] -1 and T0.[Draft No] = T1.[Draft No]  and T0.AStatus = ''Y'' group by U_Name) as [Approvedby] 
,'''' [ApproverRemarks],AppSeq,Seq,StepCode,USER_CODE, AStatus
from #TmpApprovalF T1 -- where AStatus = ''W'' AND U_NAME =  ''' + @APPROVER + '''
order by [Draft No] , AppSeq 
select [DRAFT NO],seq, stepcode , case when AStatus in (''Y'',''N'') then ''1'' else ''0'' end [RC], AStatus into #TmpCR from #table 
delete from #TmpCR where stepcode in (select stepcode from #TmpCR where RC  = ''1'' )
INSERT INTO #TABLEF
select x.*  from(
select * from #TABLE WHERE SEQ=1 AND [DRAFT NO] NOT IN (select [DRAFT NO] from #TABLE WHERE SEQ=1 AND AStatus= ''Y'') UNION ALL
select * from #TABLE WHERE SEQ=2 AND [DRAFT NO] IN (select [DRAFT NO] from #TABLE WHERE SEQ=1 AND AStatus=''Y'') UNION ALL
select * from #TABLE WHERE SEQ=3 AND [DRAFT NO] IN (select [DRAFT NO] from #TABLE WHERE SEQ=2 AND AStatus=''Y'') UNION ALL
select * from #TABLE WHERE SEQ=4 AND [DRAFT NO] IN (select [DRAFT NO] from #TABLE WHERE SEQ=3 AND AStatus=''Y'') UNION ALL
select * from #TABLE WHERE SEQ=5 AND [DRAFT NO] IN (select [DRAFT NO] from #TABLE WHERE SEQ=4 AND AStatus=''Y'') UNION ALL
select * from #TABLE WHERE SEQ=6 AND [DRAFT NO] IN (select [DRAFT NO] from #TABLE WHERE SEQ=5 AND AStatus=''Y'') UNION ALL
select * from #TABLE WHERE SEQ=7 AND [DRAFT NO] IN (select [DRAFT NO] from #TABLE WHERE SEQ=6 AND AStatus=''Y'') UNION ALL
select * from #TABLE WHERE SEQ=8 AND [DRAFT NO] IN (select [DRAFT NO] from #TABLE WHERE SEQ=7 AND AStatus=''Y'') UNION ALL
select * from #TABLE WHERE SEQ=9 AND [DRAFT NO] IN (select [DRAFT NO] from #TABLE WHERE SEQ=8 AND AStatus=''Y'')) x --join #TmpCR y on x.stepcode = y.stepcode
 WHERE   X.ACODE = ''' + @APPCODE +  ''' AND X.STEPCODE IN ( SELECT STEPCODE FROM #TMPCR    ) order by x.[Draft No],x.[AppSeq]
 delete from #TABLEF where StepCode in (select T0.StepCode from #TABLEF T0 where T0.AStatus in (''Y'',''N''))
Drop table #TmpApprovalF1
Drop table #TmpApprovalF
Drop table #TmpApproval
Drop table #TmpInitialise 
 DELETE FROM #TABLE
drop table #TmpCR
'

EXEC (@SQL + @sql1)

FETCH NEXT FROM C1 INTO @ENTITYNAME
END
CLOSE C1;
DEALLOCATE C1;

END

EXEC(' INSERT INTO #TMPCOMPANY
SELECT Code,Name,U_AB_COMPANYNAME FROM '+ @HOLDINGDB +' ..[@AB_COMPANYDATA] ')

SELECT  ENTITY, U_AB_COMPANYNAME [ENTITY NAME],[DOCUMENT TYPE],[SELECT],[DRAFT NO],[DOCUMENT NO],[CREATOR NAME],[POSTING DATE],[VENDOR NAME],[AMOUNT(BEFORE GST)],
case when isnull(APPROVER,'') = '' then 'You Are The First Approver' else APPROVER end [APPROVED BY],
[APPROVAL GRID],APPROVALCODE,USERCODE,'' [APPROVAL TYPE], [STATUS], REMARKS,[REMARKS BY APPROVER] [REASON FOR NOT APPROVING],SEQ 
INTO #BEFOREFINAL
FROM #TABLEF 
LEFT OUTER JOIN #TMPCOMPANY ON #TABLEF.ENTITY COLLATE DATABASE_DEFAULT = #TMPCOMPANY.NAME COLLATE DATABASE_DEFAULT
ORDER BY ENTITY,[DRAFT NO],APPSEQ 

DECLARE @SEQUENCE NVARCHAR(MAX),@BEFDRAFTNO NVARCHAR(MAX),@BEFENTITIY NVARCHAR(MAX),@WDDCODE NVARCHAR(MAX),@LINEAPPROVALTYPE NVARCHAR(MAX)
DECLARE @APPRGRIDCODE NVARCHAR(25)

DECLARE C2 CURSOR FOR
SELECT DISTINCT SEQ,ENTITY,[DRAFT NO],APPROVALCODE,[APPROVAL GRID] FROM #BEFOREFINAL 

OPEN C2;
FETCH NEXT FROM C2 INTO @SEQUENCE,@BEFENTITIY,@BEFDRAFTNO,@WDDCODE,@APPRGRIDCODE
WHILE @@FETCH_STATUS = 0
BEGIN

SET @SQL = 'DECLARE @COLS NVARCHAR(MAX), @COUNT_MAIN INTEGER,@COUNT_BACKUP INTEGER,@USRAPPRLTYPE NVARCHAR(MAX)
			SELECT  @COLS = COALESCE(@COLS + '','' + B.U_NAME + '''','''' + B.U_NAME + '''' ) 
			FROM '+ @BEFENTITIY +'..WDD1 A 
			INNER JOIN '+ @BEFENTITIY +'..OUSR B ON B.USERID = A.UserID 
			WHERE A.WddCode IN(SELECT WddCode FROM '+ @BEFENTITIY +'..OWDD WHERE DocEntry = '+ @BEFDRAFTNO +') AND A.Status = ''Y'' 
			AND A.WddCode = ''' + @WDDCODE + '''

			IF ''' + @APPTYPE + ''' = ''MAIN''
			BEGIN 
				SET @USRAPPRLTYPE = ''MAIN APPROVER''
			END
			ELSE IF ''' + @APPTYPE + ''' = ''BACKUP''
			BEGIN
				SET @USRAPPRLTYPE = ''BACKUP APPROVER''
			END
			ELSE
				IF ' + @SEQUENCE + ' = 1
				BEGIN
					SELECT @COUNT_MAIN = COUNT(*) FROM '+ @HOLDINGDB +' ..[@AB_APPROVALMATRIX] WHERE isnull(U_Appr1,'''') = ''' + @APPCODE + ''' AND U_ApprGridCode = ''' + @APPRGRIDCODE + '''
					SELECT @COUNT_BACKUP = COUNT(*) FROM '+ @HOLDINGDB +' ..[@AB_APPROVALMATRIX] WHERE isnull(U_Appr1B,'''') = ''' + @APPCODE + ''' AND U_ApprGridCode = ''' + @APPRGRIDCODE  + '''
					IF @COUNT_MAIN >= 1	AND @COUNT_BACKUP = 0
					BEGIN
						SET @USRAPPRLTYPE = ''MAIN APPROVER''
					END
					ELSE IF @COUNT_MAIN = 0	AND @COUNT_BACKUP >= 1
					BEGIN
						SET @USRAPPRLTYPE = ''BACKUP APPROVER''
					END
					ELSE
					BEGIN
						SET @USRAPPRLTYPE = ''MAIN APPROVER''
					END
				END
				ELSE IF ' + @SEQUENCE + ' = 2
				BEGIN
					SELECT @COUNT_MAIN = COUNT(*) FROM '+ @HOLDINGDB +' ..[@AB_APPROVALMATRIX] WHERE isnull(U_Appr2,'''') = ''' + @APPCODE + ''' AND U_ApprGridCode = ''' + @APPRGRIDCODE + '''
					SELECT @COUNT_BACKUP = COUNT(*) FROM '+ @HOLDINGDB +' ..[@AB_APPROVALMATRIX] WHERE isnull(U_Appr2B,'''') = ''' + @APPCODE + ''' AND U_ApprGridCode = ''' + @APPRGRIDCODE  + '''
					IF @COUNT_MAIN >= 1	AND @COUNT_BACKUP = 0
					BEGIN
						SET @USRAPPRLTYPE = ''MAIN APPROVER''
					END
					ELSE IF @COUNT_MAIN = 0	AND @COUNT_BACKUP >= 1
					BEGIN
						SET @USRAPPRLTYPE = ''BACKUP APPROVER''
					END
					ELSE
					BEGIN
						SET @USRAPPRLTYPE = ''MAIN APPROVER''
					END
				END
				ELSE IF ' + @SEQUENCE + ' = 3
				BEGIN
					SELECT @COUNT_MAIN = COUNT(*) FROM '+ @HOLDINGDB +' ..[@AB_APPROVALMATRIX] WHERE isnull(U_Appr3,'''') = ''' + @APPCODE + ''' AND U_ApprGridCode = ''' + @APPRGRIDCODE + '''
					SELECT @COUNT_BACKUP = COUNT(*) FROM '+ @HOLDINGDB +' ..[@AB_APPROVALMATRIX] WHERE isnull(U_Appr3B,'''') = ''' + @APPCODE + ''' AND U_ApprGridCode = ''' + @APPRGRIDCODE  + '''
					IF @COUNT_MAIN >= 1	AND @COUNT_BACKUP = 0
					BEGIN
						SET @USRAPPRLTYPE = ''MAIN APPROVER''
					END
					ELSE IF @COUNT_MAIN = 0	AND @COUNT_BACKUP >= 1
					BEGIN
						SET @USRAPPRLTYPE = ''BACKUP APPROVER''
					END
					ELSE
					BEGIN
						SET @USRAPPRLTYPE = ''MAIN APPROVER''
					END
				END

							
			--UPDATE #BEFOREFINAL SET [APPROVED BY] = (CASE WHEN ISNULL(@COLS,'''') <> '''' THEN ISNULL(@COLS,'''') ELSE ''You Are The First Approver'' END),
			-- [APPROVAL TYPE] = @USRAPPRLTYPE
			--WHERE [DRAFT NO] = ''' + @BEFDRAFTNO + ''' AND ENTITY = ''' + @BEFENTITIY + ''' 
			
			SELECT A.ENTITY, A.[ENTITY NAME],A.[DOCUMENT TYPE],A.[SELECT],A.[DRAFT NO],A.[DOCUMENT NO],A.[CREATOR NAME],A.[POSTING DATE],A.[VENDOR NAME],A.[AMOUNT(BEFORE GST)],
			(CASE WHEN ISNULL(@COLS,'''') <> '''' THEN ISNULL(@COLS,'''') ELSE ''You Are The First Approver'' END) [APPROVED BY],A.[APPROVAL GRID],A.APPROVALCODE,A.USERCODE,
			@USRAPPRLTYPE [APPROVAL TYPE], A.[STATUS], A.REMARKS,A.[REASON FOR NOT APPROVING],SEQ
			INTO #BEFOREFINAL_UPDATED
			FROM #BEFOREFINAL A

 INSERT INTO #FINAL 
SELECT A.ENTITY, A.[ENTITY NAME],A.[DOCUMENT TYPE],A.[SELECT],A.[DRAFT NO],A.[DOCUMENT NO],A.[CREATOR NAME],A.[POSTING DATE],A.[VENDOR NAME],A.[AMOUNT(BEFORE GST)],
[APPROVED BY],A.[APPROVAL GRID],A.APPROVALCODE,A.USERCODE, [APPROVAL TYPE], A.[STATUS], A.REMARKS,A.[REASON FOR NOT APPROVING],SEQ
FROM #BEFOREFINAL_UPDATED A
LEFT OUTER JOIN ' + @HOLDINGDB +' ..[@AB_APPROVALMATRIX] B ON B.U_ApprGridCode COLLATE DATABASE_DEFAULT =  A.[APPROVAL GRID] COLLATE DATABASE_DEFAULT 
WHERE A.[DRAFT NO] = ' + @BEFDRAFTNO +' '

IF @APPTYPE = 'MAIN'
BEGIN
	SET @LINEAPPROVALTYPE = 'MAIN APPROVER'

	IF @SEQUENCE = 1
	BEGIN
		SET @SQL = @SQL + ' and (isnull(U_Appr1,'''') = '''+ @APPCODE +''')'
	END
	ELSE IF @SEQUENCE = 2
	BEGIN
		SET @SQL = @SQL + ' AND (isnull(U_Appr2,'''') = '''+ @APPCODE +''' )'
	END
	ELSE IF @SEQUENCE >= 3
	BEGIN
		SET @SQL = @SQL + ' and (isnull(U_Appr3,'''') = '''+ @APPCODE +''')'
	END
END
ELSE IF @APPTYPE = 'BACKUP'
BEGIN
	SET @LINEAPPROVALTYPE = 'BACKUP APPROVER'
	IF @SEQUENCE = 1
	BEGIN
		SET @SQL = @SQL + ' and (isnull(U_Appr1B,'''') = '''+ @APPCODE +''' )'
	END
	ELSE IF @SEQUENCE = 2
	BEGIN
		SET @SQL = @SQL + ' and (isnull(U_Appr2B,'''') = '''+ @APPCODE +''')'
	END
	ELSE IF @SEQUENCE >= 3
	BEGIN
		SET @SQL = @SQL + ' and (isnull(U_Appr3B,'''') = '''+ @APPCODE +''')'
	END
END
ELSE
BEGIN  
	SET @SQL = @SQL + ' and (isnull(U_Appr1,'''') = '''+ @APPCODE +''' or isnull(U_Appr2,'''') = '''+ @APPCODE +''' or isnull(U_Appr3,'''') = '''+ @APPCODE +''' or isnull(U_Appr1B,'''') = '''+ @APPCODE +''' or isnull(U_Appr2B,'''') = '''+ @APPCODE +''' or isnull(U_Appr3B,'''') = '''+ @APPCODE +''')'
END

EXEC (@SQL)

FETCH NEXT FROM C2 INTO @SEQUENCE,@BEFENTITIY,@BEFDRAFTNO,@WDDCODE,@APPRGRIDCODE
END;
CLOSE C2;
DEALLOCATE C2;

SELECT [APPROVAL TYPE],ENTITY, [ENTITY NAME],[DOCUMENT TYPE],[SELECT],[DRAFT NO],[DOCUMENT NO],[POSTING DATE],[VENDOR NAME],[AMOUNT(BEFORE GST)],[APPROVED BY],
[APPROVAL GRID],APPROVALCODE,USERCODE,[STATUS], REMARKS,[CREATOR NAME], [REASON FOR NOT APPROVING]
 FROM #FINAL

DROP TABLE #ALLENTITYLIST
DROP TABLE #TABLEF
DROP TABLE #TABLE
DROP TABLE #TMPCOMPANY
DROP TABLE #BEFOREFINAL
DROP TABLE #BEFOREFINAL_UPDATED
DROP TABLE #FINAL

END
GO
/****** Object:  StoredProcedure [dbo].[AE_SP_Annual Internal PO]    Script Date: 22/08/2017 11:41:16 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[AE_SP_Annual Internal PO]
----@DateF as Datetime,
----@DateT as Datetime
--@DocKey numeric(18,0)
@dOCkEY int

----[dbo].[AE_SP_Annual Internal PO] '300000002'
as


SELECT 
T0.[DocNum] as 'Annual PO No.',
T0.DocTotal as 'PO DocAmount',
T1.[Currency] as 'PO Curr',
T1.[ItemCode] as 'PO ItemCode', 
T1.Dscription as 'PO ItemDescription',
T1.[Quantity] as 'PO Qty',  
T1.[LineTotal] as 'PO LineAmount',
T1.[VatSum] as 'PO GST', 
T1.[GTotal] as 'PO LineGrossAmount', 
T1.[VatSumFrgn] as 'PO GSTFC', 
T1.[TotalFrgn] as 'PO Amt FC', 
T1.[GTotalFC] as'PO GrossAmountFC', 
T1.[OpenSum], 
T1.[OpenSumFC], 
T1.[Quantity], 
T0.[CardCode], 
T0.[CardName], 
T0.[SlpCode],
T2.[ItemCode] as 'GRN ItemCode', 
T2.[Dscription] as 'GRN ItemDescription',
T2.[Quantity] AS 'GRN Qty', 
T2.[LineTotal] as 'GRN LineAmt', 
T2.[TotalFrgn] as 'GRN LineAmt FC', 
T2.[VatSum] as 'GRN GST', 
T2.[VatSumFrgn] as 'GRN GST FC', 
T2.[GTotal] as 'GRN LineGross Amt', 
T2.[GTotalFC] as 'GRN LineGross Amt FC', 
T3.[DocNum] as 'GRN No',
T3.[NumAtCard] as 'GRN ref no', 
T3.[DocTotal] as 'GRN DocTotal', 
T3.[DocTotalFC] as 'GRN DocTotalFC', 
T3.[UserSign] as 'GRN CreatorCode',
T11.U_NAME as 'GRN CreatorName',
T7.UpdateDate as'DateApproved',
T8.U_NAME as 'ApproverName',
T9.[DocNum] as 'AP Inv No', 
T9.[DocDate] as 'AP DocDate', 
T9.[DocDueDate] as 'AP DocDueDate', 
T9.[TaxDate] as 'APTaxDate', 
T9.[NumAtCard] as 'AP RefNum',
T10.SlpName as 'PurchasingUnit', 
( dbo.AE_FN003_GetApprover(T0.DocEntry,1)) as 'Approver1',
( dbo.AE_FN003_GetApprover(T0.DocEntry,2)) as 'Approver2',
( dbo.AE_FN003_GetApprover(T0.DocEntry,3)) as 'Approver3',



T5.CompnyName,T5.Phone1,T5.Phone2,T5.GlblLocNum,T5.Fax, T5.FreeZoneNo, T5.TaxIdNum,T5.LogoImage,
T5.CompnyAddr, T5.BlockF,T5.StreetF,T5.Country
 
FROM OPOR T0 
INNER JOIN POR1 T1 ON T0.[DocEntry] = T1.[DocEntry] 
LEFT OUTER JOIN PDN1 T2 ON T2.BaseEntry = T1.DocEntry and T2.ItemCode = T1.ItemCode 
INNER JOIN OPDN T3 ON T2.[DocEntry] = T3.[DocEntry]  and T3.CANCELED <> 'Y'and T3.InvntSttus <> 'O'
LEFT OUTER JOIN PCH1 T4 ON T4.BaseEntry = T2.DocEntry and T4.ItemCode = T2.ItemCode
LEFT  JOIN OPCH T9 ON T4.[DocEntry] = T9.[DocEntry] 
LEFT OUTER JOIN OSLP T10 ON T0.SlpCode = T10.SlpCode 
Inner JOIN OUSR T11 ON T3.UserSign = T11.USERID 


join 
(
 select top(1) isnull(T0.PrintHeadr,T0.CompnyName) CompnyName,T0.Phone1,T0.Phone2,T1.GlblLocNum, T0.Fax, T0.FreeZoneNo, T0.TaxIdNum,T1.ZipCode,
  T0.CompnyAddr,T1.BlockF,T1.StreetF,T3.LogoImage,
  T2.Name Country
 from OADM T0 with(nolock) 
 join ADM1 T1 with(nolock) on 1=1
 join OCRY T2 with(nolock) on T2.Code=T1.Country
  join OADP T3 with(nolock) on 1=1
) T5 on 1=1
inner join OWDD T6 on t0.docentry =  t6.docentry and T6.[ObjType]  = 22
inner join  WDD1 T7 on  T6.[WddCode]  =  T7.[WddCode] and T7.[Status]  = 'y' and T6.ObjType = '22'
INNER JOIN OUSR T8 ON T7.UserID = T8.USERID 




WHERE T0.[DocNum]  = @dOCkEY

GO
/****** Object:  StoredProcedure [dbo].[AE_SP_Emaillog_Statusupdate]    Script Date: 22/08/2017 11:41:16 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE Procedure [dbo].[AE_SP_Emaillog_Statusupdate]

as

begin

DECLARE @getEmaillog CURSOR
DECLARE @Sno Varchar(100)
DECLARE @Objecttype Varchar(100)
DECLARE @Draftkey Varchar(100)
DECLARE @sUser Varchar(1000)
DECLARE @Seq integer
DECLARE @rtotalcount integer

DECLARE @sEntity Varchar(1000)

Declare @SQL varchar(max)

---------------------- Fetching the information from Email log table
select (select max(cast(TT.Seq as integer)) from [AB_EmailStatus] TT where TT.DraftKey = T0.DraftKey ) as Rcount , 
T0.Sno , T0.ObjectType , T0.DraftKey , T0.sUser ,T0.Seq , T0.Status ,  T1.U_AB_COMCODE  [Entity]   into #Tmp_Emailstatus from [AB_EmailStatus] T0
join [@AB_COMPANYDATA] T1 on T0.Entity = T1.U_AB_COMPANYNAME 
--delete [#Tmp_Emailstatus] where Rcount = Seq 
--------------------------------------------------------------------
--Select Sno , ObjectType , DraftKey , '%' + sUser + '%' , Seq, Rcount   from [#Tmp_Emailstatus] where status = 'Pending'
--------------------   Cursor Declartion to identify the Draft no is approved for this user
SET @getEmaillog = CURSOR FOR
Select   Sno , ObjectType , DraftKey , '%' + sUser + '%' , Seq, Rcount , Entity  from [#Tmp_Emailstatus] where status = 'Pending'
OPEN @getEmaillog
FETCH NEXT
FROM @getEmaillog INTO @Sno, @Objecttype, @Draftkey, @sUser, @Seq, @rtotalcount , @sEntity
WHILE @@FETCH_STATUS = 0
BEGIN



--select @Sno, @Objecttype, @Draftkey, @sUser, @Seq, @rtotalcount , @sEntity
-------------------  Getting information from sap table with respective draft key
set @SQL = '
Declare @PLevelUser Varchar(1000)
Declare @rcount integer
SELECT TT1.[StepCode], TT0.[DocEntry], TT3.ObjType, Usercode = SUBSTRING((
    SELECT ''/'' + cast(T2.[USER_CODE]    as varchar) 
    FROM ' + @sEntity + ' ..OWDD T0  join ' + @sEntity + ' ..WDD1 T1 on T0.WddCode = T1.WddCode
	join ' + @sEntity + ' ..ousr T2 on T2.USERID = T1.UserID 
    join ' + @sEntity + ' ..odrf T3 on T3.DocEntry = T0.DocEntry
	WHERE T0.[DocEntry] = TT0.[DocEntry] and T1.[StepCode] = TT1.[StepCode]
    for XML PATH ('''')), 1,10000) + ''/'',
	Status = SUBSTRING((
    SELECT ''/'' + cast(T1.[Status]     as varchar) 
    FROM ' + @sEntity + ' ..OWDD T0  join WDD1 T1 on T0.WddCode = T1.WddCode
	join ' + @sEntity + ' ..ousr T2 on T2.USERID = T1.UserID 
    join ' + @sEntity + ' ..odrf T3 on T3.DocEntry = T0.DocEntry
	WHERE T0.[DocEntry] = TT0.[DocEntry] and T1.[StepCode] = TT1.[StepCode]
    for XML PATH ('''')), 1,10000) + ''/'' INTO #Tmp_OWDD
	FROM ' + @sEntity + ' ..OWDD TT0  join ' + @sEntity + ' ..WDD1 TT1 on TT0.WddCode = TT1.WddCode 
    join ' + @sEntity + ' ..ousr TT2 on TT2.USERID = TT1.UserID 
    join ' + @sEntity + ' ..odrf TT3 on TT3.DocEntry = TT0.DocEntry
    WHERE TT0.[DocEntry] = '''+ @Draftkey +''' and  TT3.ObjType = '''+ @Objecttype +'''
GROUP BY TT1.[StepCode],TT0.[DocEntry],TT3.ObjType
select @PLevelUser  = suser from [AB_EmailStatus]  where DraftKey  = ''' + @Draftkey + ''' and ObjectType  = '''+ @Objecttype +''' and Sno = '+ @sno +'  -1 and Entity = '''+  @sEntity +'''

 select @rcount = count(*) from #Tmp_OWDD where DocEntry = '''+ @Draftkey +''' and ObjType = '''+ @Objecttype +''' and Usercode like  ''% ISNULL(@PLevelUser,0) %''  and Status like ''%Y%''

if @rcount > 0
   begin
  	 update [AB_EmailStatus] set Status = ''Open'' where [Sno] = '''+ @Sno +''' and status = ''Pending''
   end
Drop table [#Tmp_OWDD]'
exec (@SQL )

FETCH NEXT
FROM @getEmaillog INTO @Sno, @Objecttype, @Draftkey, @sUser,  @Seq,  @rtotalcount,  @sEntity
END
CLOSE @getEmaillog
DEALLOCATE @getEmaillog
--------------------------------------------------------------------------------------------
Drop table [#Tmp_Emailstatus] 
end
GO
/****** Object:  StoredProcedure [dbo].[AE_SP0006_Journal Entry  Report]    Script Date: 22/08/2017 11:41:16 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[AE_SP0006_Journal Entry  Report]
----@DateF as Datetime,
----@DateT as Datetime
@DocKey numeric(18,0)

----AE_SP0006_Journal Voucher Report 7
as



SELECT 
 T0.[BatchNum], 
 T0.[Number],
 T0.[LocTotal],  
 T0.[UserSign],
 T0.[Memo], 
 T0.[TransId],
 T1.[Line_ID], 
 T1.[Account], 
 T2.AcctName,
 T1.[ShortName], 
 T1.[Debit], 
 T1.[Credit], 
 T1.[FCDebit], 
 T1.[FCCredit],
 T1.FCCurrency, 
 T1.[LineMemo], 
 T1.[Ref1], 
 T1.[Ref2], 
 T1.[Ref3Line], 
 T1.[RefDate], 
 T1.[Ref2Date],
 T1.[TaxDate],  
 T1.[TransCode], 
 T1.[ProfitCode], 
 T1.[BatchNum], 
 T1.[FinncPriod], 
 T1.[VatRate], 
 T1.[VatAmount], 
 T1.[GrossValue], 
 T1.[TaxCode], 
 T1.[OcrCode2], 
 T1.[OcrCode3], 
 T1.[OcrCode4], 
 T1.[OcrCode5], 
 T1.[TotalVat], 
 T1.[U_AB_PARTNER],T4.[Name] as 'Partner Name',
 T1.[U_AB_OUName],
 T1.VatGroup,
 T1.Project,
 T3.CompnyName,T3.Phone1,T3.Phone2,T3.GlblLocNum,T3.Fax, T3.FreeZoneNo, T3.TaxIdNum,T3.LogoImage,
T3.CompnyAddr, T3.BlockF,T3.StreetF,T3.Country

FROM OJDT T0  
INNER JOIN JDT1 T1 ON T0.[TransId] = T1.[TransId] 
LEFT JOIN OACT T2 ON T1.[Account] = T2.AcctCode
join 
(
 select top(1) isnull(T0.PrintHeadr,T0.CompnyName) CompnyName,T0.Phone1,T0.Phone2,T1.GlblLocNum, T0.Fax, T0.FreeZoneNo, T0.TaxIdNum,T1.ZipCode,
  T0.CompnyAddr,T1.BlockF,T1.StreetF,T3.LogoImage,
  T2.Name Country
 from OADM T0 with(nolock) 
 join ADM1 T1 with(nolock) on 1=1
 join OCRY T2 with(nolock) on T2.Code=T1.Country
  join OADP T3 with(nolock) on 1=1
) T3 on 1=1
LEFT JOIN [dbo].[@AB_PARTNER]  T4 ON T1.[U_AB_PARTNER] = T4.Code


WHERE T0.[TransId]= @DocKey 
GO
/****** Object:  StoredProcedure [dbo].[AE_SP0006_Journal Entry  Report_MultipleDocs]    Script Date: 22/08/2017 11:41:16 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[AE_SP0006_Journal Entry  Report_MultipleDocs]
@DateF as Datetime,
@DateT as Datetime,
@DocKey numeric(18,0),
@Creator varchar(100)

----AE_SP0006_Journal Voucher Report 7
as



SELECT 
 T0.[BatchNum], 
 T0.[Number],
 T0.[LocTotal],  
 T0.[UserSign],
 T0.[Memo], 
 T0.[TransId],
 T0.UserSign,
 T5.U_Name 'as Creator', 
 T1.[Line_ID], 
 T1.[Account], 
 T2.AcctName,
 T1.[ShortName], 
 T1.[Debit], 
 T1.[Credit], 
 T1.[FCDebit], 
 T1.[FCCredit],
 T1.FCCurrency, 
 T1.[LineMemo], 
 T1.[Ref1], 
 T1.[Ref2], 
 T1.[Ref3Line], 
 T1.[RefDate], 
 T1.[Ref2Date],
 T1.[TaxDate],  
 T1.[TransCode], 
 T1.[ProfitCode], 
 T1.[BatchNum], 
 T1.[FinncPriod], 
 T1.[VatRate], 
 T1.[VatAmount], 
 T1.[GrossValue], 
 T1.[TaxCode], 
 T1.[OcrCode2], 
 T1.[OcrCode3], 
 T1.[OcrCode4], 
 T1.[OcrCode5], 
 T1.[TotalVat], 
 T1.[U_AB_PARTNER],T4.[Name] as 'Partner Name',
 T1.[U_AB_OUName],
 T1.VatGroup,
 T1.Project,
 T3.CompnyName,T3.Phone1,T3.Phone2,T3.GlblLocNum,T3.Fax, T3.FreeZoneNo, T3.TaxIdNum,T3.LogoImage,
T3.CompnyAddr, T3.BlockF,T3.StreetF,T3.Country

FROM OJDT T0  
INNER JOIN JDT1 T1 ON T0.[TransId] = T1.[TransId] 
LEFT JOIN OACT T2 ON T1.[Account] = T2.AcctCode
join 
(
 select top(1) isnull(T0.PrintHeadr,T0.CompnyName) CompnyName,T0.Phone1,T0.Phone2,T1.GlblLocNum, T0.Fax, T0.FreeZoneNo, T0.TaxIdNum,T1.ZipCode,
  T0.CompnyAddr,T1.BlockF,T1.StreetF,T3.LogoImage,
  T2.Name Country
 from OADM T0 with(nolock) 
 join ADM1 T1 with(nolock) on 1=1
 join OCRY T2 with(nolock) on T2.Code=T1.Country
  join OADP T3 with(nolock) on 1=1
) T3 on 1=1
LEFT JOIN [dbo].[@AB_PARTNER]  T4 ON T1.[U_AB_PARTNER] = T4.Code
LEFT JOIN OUSR T5 on T0.UserSign = T5.USER_CODE

----WHERE T0.[TransId]= @DocKey 
GO
/****** Object:  StoredProcedure [dbo].[AE_SP0006_Journal Voucher Report]    Script Date: 22/08/2017 11:41:16 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[AE_SP0006_Journal Voucher Report]
----@DateF as Datetime,
----@DateT as Datetime
@DocKey numeric(18,0)

----AE_SP0006_Journal Voucher Report 7
as



SELECT 
 T0.[BatchNum], 
 T0.[LocTotal],  
 T0.[UserSign],
 T5.[Memo], 
 T0.[Remarks], 
 T1.[TransId], 
 T1.[Line_ID], 
 T1.[Account], 
 T2.AcctName,
 T1.[ShortName], 
 T1.[Debit], 
 T1.[Credit], 
 T1.[FCDebit], 
 T1.[FCCredit],
 t1.FCCurrency, 
 T1.[LineMemo], 
 T1.[Ref1], 
 T1.[Ref2], 
 T1.[Ref3Line], 
 T1.[RefDate], 
 T1.[Ref2Date],
 T1.[TaxDate],  
 T1.[TransCode], 
 T1.[ProfitCode], 
 T1.[BatchNum], 
 T1.[FinncPriod], 
 T1.[VatRate], 
 T1.[VatAmount], 
 T1.[GrossValue], 
 T1.[TaxCode], 
 T1.[OcrCode2], 
 T1.[OcrCode3], 
 T1.[OcrCode4], 
 T1.[OcrCode5], 
 T1.[TotalVat], 
 T1.[U_AB_PARTNER],T4.[Name] as 'Partner Name',
 T1.[U_AB_OUName],
 T1.VatGroup,
 T1.Project,
 T3.CompnyName,T3.Phone1,T3.Phone2,T3.GlblLocNum,T3.Fax, T3.FreeZoneNo, T3.TaxIdNum,T3.LogoImage,
T3.CompnyAddr, T3.BlockF,T3.StreetF,T3.Country

FROM OBTD T0  
INNER JOIN BTF1 T1 ON T0.[BatchNum] = T1.[BatchNum] 
INNER JOIN OBTF T5 ON T1.[BatchNum] = T5.[BatchNum] 
LEFT JOIN OACT T2 ON T1.[Account] = T2.AcctCode
join 
(
 select top(1) isnull(T0.PrintHeadr,T0.CompnyName) CompnyName,T0.Phone1,T0.Phone2,T1.GlblLocNum, T0.Fax, T0.FreeZoneNo, T0.TaxIdNum,T1.ZipCode,
  T0.CompnyAddr,T1.BlockF,T1.StreetF,T3.LogoImage,
  T2.Name Country
 from OADM T0 with(nolock) 
 join ADM1 T1 with(nolock) on 1=1
 join OCRY T2 with(nolock) on T2.Code=T1.Country
  join OADP T3 with(nolock) on 1=1
) T3 on 1=1
LEFT JOIN [dbo].[@AB_PARTNER]  T4 ON T1.[U_AB_PARTNER] = T4.Code

WHERE T0.[BatchNum]= @DocKey 
GO
/****** Object:  StoredProcedure [dbo].[AE_SP001_TextFileGeneration]    Script Date: 22/08/2017 11:41:16 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO






--[AE_SP001_TextFileGeneration] '20150430','20150430','52000','ZZZZZ','SVSL'



CREATE procedure [dbo].[AE_SP001_TextFileGeneration]

@DateFrom varchar(20),

@DateTo varchar(20),

@OUCodeFrom varchar(20),

@OUCodeTo varchar (20),

@Entity varchar (30)

as

begin



Declare @SQLString varchar(max)

declare @year integer

Declare @monthName varchar(20)

Declare @month integer

Declare @Day varchar(5)

Declare @Period varchar(20)





select top(1) T0.Name into #period from OFPR T0

delete #period 

---update on 07052015

--------------------- Get Period from the SAP Period Table based on the Date

set @SQLString = '

SELECT isnull(RIGHT(T0.[Code],4) + ''0'' + LEFT (T0.[Code],2),''0'') [Name]

           FROM ' + @Entity + '.. OFPR T0 WHERE month(T0.[F_RefDate]) = month('''+ @DateFrom + ''')  and  year(T0.[F_RefDate]) = year(''' + @DateFrom + ''' ) '



insert into #period 

execute (@SQLString )



select @Period = Name from #period 



-------------------- Table structure creation



 select top(1) jdt1.Account [AcctCode] , jdt1.Account  [RefDate] , jdt1.OcrCode3 [OU Code], 

 jdt1.Account [Entity], jdt1.Account [DC], jdt1.Credit [Amount] , jdt1.Account [Cat] into #level1 from JDT1

 

 delete  #level1



/*

set @year = LEFT(@DateTo ,4)

set @Day = RIGHT(@DateTo ,2)

set @month = left(RIGHT(@DateTo,4),2)



select @monthName = DateName( month , DateAdd( month , @month , 0 ) - 1 )

select @Day = day(dateadd(day, -1, dateadd(month, 1, dateadd(day, 1 - day(@DateTo), @DateTo))))



select @Period = code from [@AB_IPOWERPERIOD] where Name = @monthName + ' ' + @Day



set @Period =  cast(@year as varchar) + replicate('0', 3 - LEN(@period)) + cast(@period as varchar)





set @SQLString = 'select * from (

SELECT T0.AcctCode , ''' +  @Period + ''' [RefDate]   ,isnull(T1.OcrCode3,'''') [OU Code] , ''' + @Entity + ''' [Entity],

case when T0.GroupMask = 1 then ''D''

when T0.GroupMask = 2 then ''C''

when T0.GroupMask = 3 then ''C'' end [DC], abs(sum(T1.Debit - T1.Credit)) [Amount]   

FROM ' + @Entity + '.. OACT T0 INNER JOIN ' + @Entity + '.. JDT1 T1 on T0.AcctCode  = T1.Account  INNER JOIN ' + @Entity + '.. OJDT T2 

ON T1.TransId = T2.TransId

where T2.RefDate <= ''' + @DateTo + ''' and T0.GroupMask  in (1,2,3)

--and T1.OcrCode3 >= ''' + @OUCodeFrom + ''' and T1.OcrCode3 <= ''' + @OUCodeTo + '''  

group by T0.AcctCode , T1.OcrCode3  ,T0.GroupMask

*/



------------------------  Query for getting the Balance sheet 

set @SQLString = 'select *   from (

--------------------- Balance Sheet 

SELECT T0.AcctCode , ''' +  @Period + ''' [RefDate]   ,isnull(T1.OcrCode3,'''') [OU Code] , ''' + @Entity + ''' [Entity],

case when sum(T1.Debit - T1.Credit) >= 0 then ''D'' else ''C'' end [DC], abs(sum(T1.Debit - T1.Credit)) [Amount] , ''JE'' [Cat]

FROM ' + @Entity + '.. OACT T0 INNER JOIN ' + @Entity + '.. JDT1 T1 on T0.AcctCode  = T1.Account  INNER JOIN ' + @Entity + '.. OJDT T2 

ON T1.TransId = T2.TransId

where T2.RefDate <= ''' + @DateTo + ''' and T0.GroupMask  in (1,2,3)

and T1.OcrCode3 >= ''' + @OUCodeFrom + ''' and T1.OcrCode3 <= ''' + @OUCodeTo + '''  

group by T0.AcctCode , T1.OcrCode3  ,T0.GroupMask

union all

-------------------- Profit and loss

SELECT T0.AcctCode , ''' +  @Period + ''' [Period]   ,isnull(T1.OcrCode3,'''') [OU Code] , ''' + @Entity + ''' [Entity],

case when sum(T1.Debit - T1.Credit) >= 0 then ''D''

when sum(T1.Debit - T1.Credit) < 0 then ''C''

 end [DC], abs(sum(T1.Debit - T1.Credit)) [Amount] , ''JE'' [Cat]

FROM ' + @Entity + '.. OACT T0 INNER JOIN ' + @Entity + '.. JDT1 T1 on T0.AcctCode  = T1.Account  INNER JOIN ' + @Entity + '.. OJDT T2 

ON T1.TransId = T2.TransId

where T2.RefDate >= ''' + @DateFrom + ''' and T2.RefDate <= ''' + @DateTo + ''' and T0.GroupMask not in (1,2,3)

and T1.OcrCode3 >= ''' + @OUCodeFrom + ''' and T1.OcrCode3 <= ''' + @OUCodeTo + ''' 

group by T0.AcctCode , T1.OcrCode3  ,T0.GroupMask 

union all

------------------- Import Statistics

select T0.U_AB_GLCODE , T0.U_AB_PERIOD ,isnull(T0.U_AB_OPER_UNIT,'''') [OU Code] , ''' + @Entity + ''' [Entity], T0.U_AB_DEBIT_CREDIT , 

sum(T0.U_AB_AMOUNT) [Amount] , ''IM'' [Cat]

from ' + @Entity + '..[@AB_STATITISTICSDATA] T0

where T0.U_AB_PERIOD = ''' +  @Period + '''

and T0.U_AB_OPER_UNIT >= ''' + @OUCodeFrom + ''' and T0.U_AB_OPER_UNIT <= ''' + @OUCodeTo + '''

group by T0.U_AB_GLCODE , T0.U_AB_PERIOD ,T0.U_AB_OPER_UNIT, T0.U_AB_DEBIT_CREDIT ) tmp

order by tmp.AcctCode'



insert into #level1 

    execute (@SQLString)

    

 ---------------  Segregating Journal memos and Statistics data
 
 

select * into #memojournals from #level1 T0 where T0.Cat = 'JE'

select * into #importstatistics from #level1 T0 where T0.Cat = 'IM'



---------------  including the distribution rules in the journal memos data

delete #level1 



--if @Entity = 'IAS7'

-- begin

--  set @SQLString = '

--   select ''"'' + T0.AcctCode + ''"'', ''"'' + T0.RefDate + ''"'', ''"'' + T2.PrcCode + ''"'' [OU Code] , ''"'' + T0.Entity + ''"'' , ''"'' + T0.DC + ''"'', (T2.PrcAmount/T1.OcrTotal ) * T0.Amount [Amount], '''' [Cat]

--   from #memojournals T0 left outer Join ' + @Entity + '.. OOCR T1 ON T1.OcrCode = T0.[OU Code] left outer Join ' + @Entity + '.. OCR1 T2 

--    on T1.OcrCode = T2.OcrCode where T2.ValidFrom <= ''' + @DateFrom + ''' and (T2.ValidTo >= ''' + @DateTo + ''' or T2.ValidTo is null )'

-- end

--else

-- begin

--  set @SQLString = '

--     select ''"'' + T0.AcctCode + ''"'', ''"'' + T0.RefDate + ''"'', ''"'' + T2.PrcCode + ''"'' [OU Code] , ''"'' + T0.Entity + ''"'' , ''"'' + T0.DC + ''"'', (T2.PrcAmount/T1.OcrTotal ) * T0.Amount [Amount], '''' [Cat]

--      from #memojournals T0 left outer Join ' + @Entity + '.. OOCR T1 ON T1.OcrCode = T0.[OU Code] left outer Join ' + @Entity + '.. OCR1 T2 

--       on T1.OcrCode = T2.OcrCode '

-- end



 set @SQLString = '

   select ''"'' + isnull(T0.AcctCode,'''') + ''"'' [AcctCode], ''"'' + isnull(T0.RefDate,'''') + ''"'' [RefDate], ''"'' + isnull(T2.PrcCode,'''') + ''"'' [OU Code] , ''"'' + isnull(T0.Entity,'''') + ''"'' [Entity] , ''"'' + isnull(T0.DC,'''') + ''"'' [DC], isnull( (T2.PrcAmount/T1.OcrTotal ) * T0.Amount,0.00) [Amount], '''' [Cat]
   into #level1
   from #memojournals T0 left outer Join ' + @Entity + '.. OOCR T1 ON T1.OcrCode = T0.[OU Code] left outer Join ' + @Entity + '.. OCR1 T2 

    on T1.OcrCode = T2.OcrCode where T2.ValidFrom <= ''' + @DateFrom + ''' and (T2.ValidTo >= ''' + @DateTo + ''' or T2.ValidTo is null )
	
	select T0.AcctCode , T0.RefDate , T0.[OU Code] , T0.Entity , T0.DC , T0.Amount from #level1 T0 where T0.Amount > 0

union all

select ''"'' + isnull(T0.AcctCode,'''') + ''"'' [AcctCode], ''"'' + isnull(T0.RefDate,'''') + ''"'' [RefDate], ''"'' + isnull(T0.[OU Code],'''') + ''"'' [OU Code] , ''"'' + isnull(T0.Entity,'''') + ''"'' [Entity] , ''"'' + isnull(T0.DC,'''') + ''"'' [DC] , isnull(T0.Amount,0.00)  [Amount] from #importstatistics T0

where T0.Amount > 0

order by AcctCode
	
	
	'

	
 print @SQLString

 

--insert into #level1 

execute (@SQLString)

 

----------------  Final Output

--select T0.AcctCode , T0.RefDate , T0.[OU Code] , T0.Entity , T0.DC , T0.Amount from #level1 T0 where T0.Amount > 0

--union all

--select '"' + T0.AcctCode + '"' , '"' + T0.RefDate + '"' , '"' + T0.[OU Code] + '"' , '"' + T0.Entity + '"' , '"' + T0.DC + '"' , T0.Amount from #importstatistics T0

--where T0.Amount > 0

--order by AcctCode



--select * from #level1

--print @SQLString



end

GO
/****** Object:  StoredProcedure [dbo].[AE_SP001_TextFileGeneration_001]    Script Date: 22/08/2017 11:41:16 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

--[AE_SP001_TextFileGeneration] '20150430','20150430','52000','ZZZZZ','PWCL'

CREATE procedure [dbo].[AE_SP001_TextFileGeneration_001]
@DateFrom varchar(20),
@DateTo varchar(20),
@OUCodeFrom varchar(20),
@OUCodeTo varchar (20),
@Entity varchar (30)
as
begin

Declare @SQLString varchar(max)
declare @year integer
Declare @monthName varchar(20)
Declare @month integer
Declare @Day varchar(5)
Declare @Period varchar(20)


select top(1) T0.Name into #period from OFPR T0
delete #period 
---update on 07052015
--------------------- Get Period from the SAP Period Table based on the Date
set @SQLString = '
SELECT isnull(RIGHT(T0.[Code],4) + ''0'' + LEFT (T0.[Code],2),''0'') [Name]
           FROM ' + @Entity + '.. OFPR T0 WHERE month(T0.[F_RefDate]) = month('''+ @DateFrom + ''')  and  year(T0.[F_RefDate]) = year(''' + @DateFrom + ''' ) '

insert into #period 
execute (@SQLString )

select @Period = Name from #period 

-------------------- Table structure creation

 select top(1) jdt1.Account [AcctCode] , jdt1.Account  [RefDate] , jdt1.OcrCode3 [OU Code], 
 jdt1.Account [Entity], jdt1.Account [DC], jdt1.Credit [Amount] , jdt1.Account [Cat] into #level1 from JDT1
 
 delete  #level1

/*
set @year = LEFT(@DateTo ,4)
set @Day = RIGHT(@DateTo ,2)
set @month = left(RIGHT(@DateTo,4),2)

select @monthName = DateName( month , DateAdd( month , @month , 0 ) - 1 )
select @Day = day(dateadd(day, -1, dateadd(month, 1, dateadd(day, 1 - day(@DateTo), @DateTo))))

select @Period = code from [@AB_IPOWERPERIOD] where Name = @monthName + ' ' + @Day

set @Period =  cast(@year as varchar) + replicate('0', 3 - LEN(@period)) + cast(@period as varchar)


set @SQLString = 'select * from (
SELECT T0.AcctCode , ''' +  @Period + ''' [RefDate]   ,isnull(T1.OcrCode3,'''') [OU Code] , ''' + @Entity + ''' [Entity],
case when T0.GroupMask = 1 then ''D''
when T0.GroupMask = 2 then ''C''
when T0.GroupMask = 3 then ''C'' end [DC], abs(sum(T1.Debit - T1.Credit)) [Amount]   
FROM ' + @Entity + '.. OACT T0 INNER JOIN ' + @Entity + '.. JDT1 T1 on T0.AcctCode  = T1.Account  INNER JOIN ' + @Entity + '.. OJDT T2 
ON T1.TransId = T2.TransId
where T2.RefDate <= ''' + @DateTo + ''' and T0.GroupMask  in (1,2,3)
--and T1.OcrCode3 >= ''' + @OUCodeFrom + ''' and T1.OcrCode3 <= ''' + @OUCodeTo + '''  
group by T0.AcctCode , T1.OcrCode3  ,T0.GroupMask
*/

------------------------  Query for getting the Balance sheet & PL Accounts along with Satistics data
set @SQLString = 'select *   from (
--------------------- Balance Sheet 
SELECT T0.AcctCode , ''' +  @Period + ''' [RefDate]   ,isnull(T1.OcrCode3,'''') [OU Code] , ''' + @Entity + ''' [Entity],
case when sum(T1.Debit - T1.Credit) >= 0 then ''D'' else ''C'' end [DC], abs(sum(T1.Debit - T1.Credit)) [Amount] , ''JE'' [Cat]
FROM ' + @Entity + '.. OACT T0 INNER JOIN ' + @Entity + '.. JDT1 T1 on T0.AcctCode  = T1.Account  INNER JOIN ' + @Entity + '.. OJDT T2 
ON T1.TransId = T2.TransId
where T2.RefDate <= ''' + @DateTo + ''' and T0.GroupMask  in (1,2,3)
and T1.OcrCode3 >= ''' + @OUCodeFrom + ''' and T1.OcrCode3 <= ''' + @OUCodeTo + '''  
group by T0.AcctCode , T1.OcrCode3  ,T0.GroupMask
union all
-------------------- Profit and loss
SELECT T0.AcctCode , ''' +  @Period + ''' [Period]   ,isnull(T1.OcrCode3,'''') [OU Code] , ''' + @Entity + ''' [Entity],
case when sum(T1.Debit - T1.Credit) >= 0 then ''D''
when sum(T1.Debit - T1.Credit) < 0 then ''C''
 end [DC], abs(sum(T1.Debit - T1.Credit)) [Amount] , ''JE'' [Cat]
FROM ' + @Entity + '.. OACT T0 INNER JOIN ' + @Entity + '.. JDT1 T1 on T0.AcctCode  = T1.Account  INNER JOIN ' + @Entity + '.. OJDT T2 
ON T1.TransId = T2.TransId
where T2.RefDate >= ''' + @DateFrom + ''' and T2.RefDate <= ''' + @DateTo + ''' and T0.GroupMask not in (1,2,3)
and T1.OcrCode3 >= ''' + @OUCodeFrom + ''' and T1.OcrCode3 <= ''' + @OUCodeTo + ''' 
group by T0.AcctCode , T1.OcrCode3  ,T0.GroupMask 
union all
------------------- Import Statistics
select T0.U_AB_GLCODE , T0.U_AB_PERIOD ,isnull(T0.U_AB_OPER_UNIT,'''') [OU Code] , ''' + @Entity + ''' [Entity], T0.U_AB_DEBIT_CREDIT , 
sum(T0.U_AB_AMOUNT) [Amount] , ''IM'' [Cat]
from ' + @Entity + '..[@AB_STATITISTICSDATA] T0
where T0.U_AB_PERIOD = ''' +  @Period + '''
and T0.U_AB_OPER_UNIT >= ''' + @OUCodeFrom + ''' and T0.U_AB_OPER_UNIT <= ''' + @OUCodeTo + '''
group by T0.U_AB_GLCODE , T0.U_AB_PERIOD ,T0.U_AB_OPER_UNIT, T0.U_AB_DEBIT_CREDIT ) tmp
order by tmp.AcctCode'

insert into #level1 
    execute (@SQLString)
    
 ---------------  Segregating Journal memos and Statistics data
 
select * into #memojournals from #level1 T0 where T0.Cat = 'JE'
select * into #importstatistics from #level1 T0 where T0.Cat = 'IM'

---------------  including the distribution rules in the journal memos data
delete #level1 

set @SQLString = '
select T0.AcctCode , T0.RefDate , T2.PrcCode [OU Code] , T0.Entity , T0.DC , (T2.PrcAmount/T1.OcrTotal ) * T0.Amount [Amount], '''' [Cat]
from #memojournals T0 left outer Join ' + @Entity + '.. OOCR T1 ON T1.OcrCode = T0.[OU Code] left outer Join ' + @Entity + '.. OCR1 T2 
on T1.OcrCode = T2.OcrCode '

insert into #level1 
execute (@SQLString)

----------------  Final Output
select T0.AcctCode , T0.RefDate , T0.[OU Code] , T0.Entity , T0.DC , T0.Amount from #level1 T0
union all
select T0.AcctCode , T0.RefDate , T0.[OU Code] , T0.Entity , T0.DC , T0.Amount from #importstatistics T0
order by AcctCode

--select * from #level1
--print @SQLString

end
GO
/****** Object:  StoredProcedure [dbo].[AE_SP001_TextFileGeneration_09112015]    Script Date: 22/08/2017 11:41:16 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


--[AE_SP001_TextFileGeneration] '20150430','20150430','52000','ZZZZZ','SVSL'

CREATE procedure [dbo].[AE_SP001_TextFileGeneration_09112015]
@DateFrom varchar(20),
@DateTo varchar(20),
@OUCodeFrom varchar(20),
@OUCodeTo varchar (20),
@Entity varchar (30)
as
begin

Declare @SQLString varchar(max)
declare @year integer
Declare @monthName varchar(20)
Declare @month integer
Declare @Day varchar(5)
Declare @Period varchar(20)


select top(1) T0.Name into #period from OFPR T0
delete #period 
---update on 07052015
--------------------- Get Period from the SAP Period Table based on the Date
set @SQLString = '
SELECT isnull(RIGHT(T0.[Code],4) + ''0'' + LEFT (T0.[Code],2),''0'') [Name]
           FROM ' + @Entity + '.. OFPR T0 WHERE month(T0.[F_RefDate]) = month('''+ @DateFrom + ''')  and  year(T0.[F_RefDate]) = year(''' + @DateFrom + ''' ) '

insert into #period 
execute (@SQLString )

select @Period = Name from #period 

-------------------- Table structure creation

 select top(1) jdt1.Account [AcctCode] , jdt1.Account  [RefDate] , jdt1.OcrCode3 [OU Code], 
 jdt1.Account [Entity], jdt1.Account [DC], jdt1.Credit [Amount] , jdt1.Account [Cat] into #level1 from JDT1
 
 delete  #level1

/*
set @year = LEFT(@DateTo ,4)
set @Day = RIGHT(@DateTo ,2)
set @month = left(RIGHT(@DateTo,4),2)

select @monthName = DateName( month , DateAdd( month , @month , 0 ) - 1 )
select @Day = day(dateadd(day, -1, dateadd(month, 1, dateadd(day, 1 - day(@DateTo), @DateTo))))

select @Period = code from [@AB_IPOWERPERIOD] where Name = @monthName + ' ' + @Day

set @Period =  cast(@year as varchar) + replicate('0', 3 - LEN(@period)) + cast(@period as varchar)


set @SQLString = 'select * from (
SELECT T0.AcctCode , ''' +  @Period + ''' [RefDate]   ,isnull(T1.OcrCode3,'''') [OU Code] , ''' + @Entity + ''' [Entity],
case when T0.GroupMask = 1 then ''D''
when T0.GroupMask = 2 then ''C''
when T0.GroupMask = 3 then ''C'' end [DC], abs(sum(T1.Debit - T1.Credit)) [Amount]   
FROM ' + @Entity + '.. OACT T0 INNER JOIN ' + @Entity + '.. JDT1 T1 on T0.AcctCode  = T1.Account  INNER JOIN ' + @Entity + '.. OJDT T2 
ON T1.TransId = T2.TransId
where T2.RefDate <= ''' + @DateTo + ''' and T0.GroupMask  in (1,2,3)
--and T1.OcrCode3 >= ''' + @OUCodeFrom + ''' and T1.OcrCode3 <= ''' + @OUCodeTo + '''  
group by T0.AcctCode , T1.OcrCode3  ,T0.GroupMask
*/

------------------------  Query for getting the Balance sheet 
set @SQLString = 'select *   from (
--------------------- Balance Sheet 
SELECT T0.AcctCode , ''' +  @Period + ''' [RefDate]   ,isnull(T1.OcrCode3,'''') [OU Code] , ''' + @Entity + ''' [Entity],
case when sum(T1.Debit - T1.Credit) >= 0 then ''D'' else ''C'' end [DC], abs(sum(T1.Debit - T1.Credit)) [Amount] , ''JE'' [Cat]
FROM ' + @Entity + '.. OACT T0 INNER JOIN ' + @Entity + '.. JDT1 T1 on T0.AcctCode  = T1.Account  INNER JOIN ' + @Entity + '.. OJDT T2 
ON T1.TransId = T2.TransId
where T2.RefDate <= ''' + @DateTo + ''' and T0.GroupMask  in (1,2,3)
and T1.OcrCode3 >= ''' + @OUCodeFrom + ''' and T1.OcrCode3 <= ''' + @OUCodeTo + '''  
group by T0.AcctCode , T1.OcrCode3  ,T0.GroupMask
union all
-------------------- Profit and loss
SELECT T0.AcctCode , ''' +  @Period + ''' [Period]   ,isnull(T1.OcrCode3,'''') [OU Code] , ''' + @Entity + ''' [Entity],
case when sum(T1.Debit - T1.Credit) >= 0 then ''D''
when sum(T1.Debit - T1.Credit) < 0 then ''C''
 end [DC], abs(sum(T1.Debit - T1.Credit)) [Amount] , ''JE'' [Cat]
FROM ' + @Entity + '.. OACT T0 INNER JOIN ' + @Entity + '.. JDT1 T1 on T0.AcctCode  = T1.Account  INNER JOIN ' + @Entity + '.. OJDT T2 
ON T1.TransId = T2.TransId
where T2.RefDate >= ''' + @DateFrom + ''' and T2.RefDate <= ''' + @DateTo + ''' and T0.GroupMask not in (1,2,3)
and T1.OcrCode3 >= ''' + @OUCodeFrom + ''' and T1.OcrCode3 <= ''' + @OUCodeTo + ''' 
group by T0.AcctCode , T1.OcrCode3  ,T0.GroupMask 
union all
------------------- Import Statistics
select T0.U_AB_GLCODE , T0.U_AB_PERIOD ,isnull(T0.U_AB_OPER_UNIT,'''') [OU Code] , ''' + @Entity + ''' [Entity], T0.U_AB_DEBIT_CREDIT , 
sum(T0.U_AB_AMOUNT) [Amount] , ''IM'' [Cat]
from ' + @Entity + '..[@AB_STATITISTICSDATA] T0
where T0.U_AB_PERIOD = ''' +  @Period + '''
and T0.U_AB_OPER_UNIT >= ''' + @OUCodeFrom + ''' and T0.U_AB_OPER_UNIT <= ''' + @OUCodeTo + '''
group by T0.U_AB_GLCODE , T0.U_AB_PERIOD ,T0.U_AB_OPER_UNIT, T0.U_AB_DEBIT_CREDIT ) tmp
order by tmp.AcctCode'

insert into #level1 
    execute (@SQLString)
    
 ---------------  Segregating Journal memos and Statistics data
 
select * into #memojournals from #level1 T0 where T0.Cat = 'JE'
select * into #importstatistics from #level1 T0 where T0.Cat = 'IM'

---------------  including the distribution rules in the journal memos data
delete #level1 

--if @Entity = 'IAS7'
-- begin
--  set @SQLString = '
--   select ''"'' + T0.AcctCode + ''"'', ''"'' + T0.RefDate + ''"'', ''"'' + T2.PrcCode + ''"'' [OU Code] , ''"'' + T0.Entity + ''"'' , ''"'' + T0.DC + ''"'', (T2.PrcAmount/T1.OcrTotal ) * T0.Amount [Amount], '''' [Cat]
--   from #memojournals T0 left outer Join ' + @Entity + '.. OOCR T1 ON T1.OcrCode = T0.[OU Code] left outer Join ' + @Entity + '.. OCR1 T2 
--    on T1.OcrCode = T2.OcrCode where T2.ValidFrom <= ''' + @DateFrom + ''' and (T2.ValidTo >= ''' + @DateTo + ''' or T2.ValidTo is null )'
-- end
--else
-- begin
--  set @SQLString = '
--     select ''"'' + T0.AcctCode + ''"'', ''"'' + T0.RefDate + ''"'', ''"'' + T2.PrcCode + ''"'' [OU Code] , ''"'' + T0.Entity + ''"'' , ''"'' + T0.DC + ''"'', (T2.PrcAmount/T1.OcrTotal ) * T0.Amount [Amount], '''' [Cat]
--      from #memojournals T0 left outer Join ' + @Entity + '.. OOCR T1 ON T1.OcrCode = T0.[OU Code] left outer Join ' + @Entity + '.. OCR1 T2 
--       on T1.OcrCode = T2.OcrCode '
-- end

 set @SQLString = '
   select ''"'' + T0.AcctCode + ''"'', ''"'' + T0.RefDate + ''"'', ''"'' + T2.PrcCode + ''"'' [OU Code] , ''"'' + T0.Entity + ''"'' , ''"'' + T0.DC + ''"'', (T2.PrcAmount/T1.OcrTotal ) * T0.Amount [Amount], '''' [Cat]
   from #memojournals T0 left outer Join ' + @Entity + '.. OOCR T1 ON T1.OcrCode = T0.[OU Code] left outer Join ' + @Entity + '.. OCR1 T2 
    on T1.OcrCode = T2.OcrCode where T2.ValidFrom <= ''' + @DateFrom + ''' and (T2.ValidTo >= ''' + @DateTo + ''' or T2.ValidTo is null )'


 print @SQLString

insert into #level1 
execute (@SQLString)

----------------  Final Output
select T0.AcctCode , T0.RefDate , T0.[OU Code] , T0.Entity , T0.DC , T0.Amount from #level1 T0 where T0.Amount > 0
union all
select '"' + T0.AcctCode + '"' , '"' + T0.RefDate + '"' , '"' + T0.[OU Code] + '"' , '"' + T0.Entity + '"' , '"' + T0.DC + '"' , T0.Amount from #importstatistics T0
where T0.Amount > 0
order by AcctCode

--select * from #level1
--print @SQLString

end
GO
/****** Object:  StoredProcedure [dbo].[AE_SP001_TextFileGeneration_14102015]    Script Date: 22/08/2017 11:41:16 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

--[AE_SP001_TextFileGeneration] '20150430','20150430','52000','ZZZZZ','SVSL'

CREATE procedure [dbo].[AE_SP001_TextFileGeneration_14102015]
@DateFrom varchar(20),
@DateTo varchar(20),
@OUCodeFrom varchar(20),
@OUCodeTo varchar (20),
@Entity varchar (30)
as
begin

Declare @SQLString varchar(max)
declare @year integer
Declare @monthName varchar(20)
Declare @month integer
Declare @Day varchar(5)
Declare @Period varchar(20)


select top(1) T0.Name into #period from OFPR T0
delete #period 
---update on 07052015
--------------------- Get Period from the SAP Period Table based on the Date
set @SQLString = '
SELECT isnull(RIGHT(T0.[Code],4) + ''0'' + LEFT (T0.[Code],2),''0'') [Name]
           FROM ' + @Entity + '.. OFPR T0 WHERE month(T0.[F_RefDate]) = month('''+ @DateFrom + ''')  and  year(T0.[F_RefDate]) = year(''' + @DateFrom + ''' ) '

insert into #period 
execute (@SQLString )

select @Period = Name from #period 

-------------------- Table structure creation

 select top(1) jdt1.Account [AcctCode] , jdt1.Account  [RefDate] , jdt1.OcrCode3 [OU Code], 
 jdt1.Account [Entity], jdt1.Account [DC], jdt1.Credit [Amount] , jdt1.Account [Cat] into #level1 from JDT1
 
 delete  #level1

/*
set @year = LEFT(@DateTo ,4)
set @Day = RIGHT(@DateTo ,2)
set @month = left(RIGHT(@DateTo,4),2)

select @monthName = DateName( month , DateAdd( month , @month , 0 ) - 1 )
select @Day = day(dateadd(day, -1, dateadd(month, 1, dateadd(day, 1 - day(@DateTo), @DateTo))))

select @Period = code from [@AB_IPOWERPERIOD] where Name = @monthName + ' ' + @Day

set @Period =  cast(@year as varchar) + replicate('0', 3 - LEN(@period)) + cast(@period as varchar)


set @SQLString = 'select * from (
SELECT T0.AcctCode , ''' +  @Period + ''' [RefDate]   ,isnull(T1.OcrCode3,'''') [OU Code] , ''' + @Entity + ''' [Entity],
case when T0.GroupMask = 1 then ''D''
when T0.GroupMask = 2 then ''C''
when T0.GroupMask = 3 then ''C'' end [DC], abs(sum(T1.Debit - T1.Credit)) [Amount]   
FROM ' + @Entity + '.. OACT T0 INNER JOIN ' + @Entity + '.. JDT1 T1 on T0.AcctCode  = T1.Account  INNER JOIN ' + @Entity + '.. OJDT T2 
ON T1.TransId = T2.TransId
where T2.RefDate <= ''' + @DateTo + ''' and T0.GroupMask  in (1,2,3)
--and T1.OcrCode3 >= ''' + @OUCodeFrom + ''' and T1.OcrCode3 <= ''' + @OUCodeTo + '''  
group by T0.AcctCode , T1.OcrCode3  ,T0.GroupMask
*/

------------------------  Query for getting the Balance sheet & PL Accounts along with Satistics data
set @SQLString = 'select *   from (
--------------------- Balance Sheet 
SELECT T0.AcctCode , ''' +  @Period + ''' [RefDate]   ,isnull(T1.OcrCode3,'''') [OU Code] , ''' + @Entity + ''' [Entity],
case when sum(T1.Debit - T1.Credit) >= 0 then ''D'' else ''C'' end [DC], abs(sum(T1.Debit - T1.Credit)) [Amount] , ''JE'' [Cat]
FROM ' + @Entity + '.. OACT T0 INNER JOIN ' + @Entity + '.. JDT1 T1 on T0.AcctCode  = T1.Account  INNER JOIN ' + @Entity + '.. OJDT T2 
ON T1.TransId = T2.TransId
where T2.RefDate <= ''' + @DateTo + ''' and T0.GroupMask  in (1,2,3)
and T1.OcrCode3 >= ''' + @OUCodeFrom + ''' and T1.OcrCode3 <= ''' + @OUCodeTo + '''  
group by T0.AcctCode , T1.OcrCode3  ,T0.GroupMask
union all
-------------------- Profit and loss
SELECT T0.AcctCode , ''' +  @Period + ''' [Period]   ,isnull(T1.OcrCode3,'''') [OU Code] , ''' + @Entity + ''' [Entity],
case when sum(T1.Debit - T1.Credit) >= 0 then ''D''
when sum(T1.Debit - T1.Credit) < 0 then ''C''
 end [DC], abs(sum(T1.Debit - T1.Credit)) [Amount] , ''JE'' [Cat]
FROM ' + @Entity + '.. OACT T0 INNER JOIN ' + @Entity + '.. JDT1 T1 on T0.AcctCode  = T1.Account  INNER JOIN ' + @Entity + '.. OJDT T2 
ON T1.TransId = T2.TransId
where T2.RefDate >= ''' + @DateFrom + ''' and T2.RefDate <= ''' + @DateTo + ''' and T0.GroupMask not in (1,2,3)
and T1.OcrCode3 >= ''' + @OUCodeFrom + ''' and T1.OcrCode3 <= ''' + @OUCodeTo + ''' 
group by T0.AcctCode , T1.OcrCode3  ,T0.GroupMask 
union all
------------------- Import Statistics
select T0.U_AB_GLCODE , T0.U_AB_PERIOD ,isnull(T0.U_AB_OPER_UNIT,'''') [OU Code] , ''' + @Entity + ''' [Entity], T0.U_AB_DEBIT_CREDIT , 
sum(T0.U_AB_AMOUNT) [Amount] , ''IM'' [Cat]
from ' + @Entity + '..[@AB_STATITISTICSDATA] T0
where T0.U_AB_PERIOD = ''' +  @Period + '''
and T0.U_AB_OPER_UNIT >= ''' + @OUCodeFrom + ''' and T0.U_AB_OPER_UNIT <= ''' + @OUCodeTo + '''
group by T0.U_AB_GLCODE , T0.U_AB_PERIOD ,T0.U_AB_OPER_UNIT, T0.U_AB_DEBIT_CREDIT ) tmp
order by tmp.AcctCode'

insert into #level1 
    execute (@SQLString)
    
 ---------------  Segregating Journal memos and Statistics data
 
select * into #memojournals from #level1 T0 where T0.Cat = 'JE'
select * into #importstatistics from #level1 T0 where T0.Cat = 'IM'

---------------  including the distribution rules in the journal memos data
delete #level1 

set @SQLString = '
select ''"'' + T0.AcctCode + ''"'', ''"'' + T0.RefDate + ''"'', ''"'' + T2.PrcCode + ''"'' [OU Code] , ''"'' + T0.Entity + ''"'' , ''"'' + T0.DC + ''"'', (T2.PrcAmount/T1.OcrTotal ) * T0.Amount [Amount], '''' [Cat]
from #memojournals T0 left outer Join ' + @Entity + '.. OOCR T1 ON T1.OcrCode = T0.[OU Code] left outer Join ' + @Entity + '.. OCR1 T2 
on T1.OcrCode = T2.OcrCode '

insert into #level1 
execute (@SQLString)

----------------  Final Output
select T0.AcctCode , T0.RefDate , T0.[OU Code] , T0.Entity , T0.DC , T0.Amount from #level1 T0
union all
select T0.AcctCode , T0.RefDate , T0.[OU Code] , T0.Entity , T0.DC , T0.Amount from #importstatistics T0
order by AcctCode

--select * from #level1
--print @SQLString

end
GO
/****** Object:  StoredProcedure [dbo].[AE_SP003_PaidInvoicesList]    Script Date: 22/08/2017 11:41:16 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[AE_SP003_PaidInvoicesList]
-- Add the parameters for the stored procedure here
@FrmDate Date,
@ToDate Date
----@FrmPayMethod varchar(100),
----@ToPayMethod varchar(100)
--@WizardCode VARCHAR(20)

AS
BEGIN

select
OVPM.DocDate AS [Date Paid], OVPM.DocNum AS [Payment Voucher No], case OVPM.DocType when 'A' then '' else OVPM.CardCode END AS [Vendor Code],
 case OVPM.DocType when 'A' then OVPM.Address else OVPM.CardName END AS [Vendor Name], 
case when OVPM.DocCurr = 'SGD' then OVPM.DOCTOTAL else OVPM.DocTotalFC end as PaidDocTotal,
case when OVPM.DocCurr = 'SGD' then OVPM.DOCTOTAL else 0.00 end as TotalpaymentLC,
case when VPM2.InvType = 18 then OPCH.DocNum  else ORPC.DocNum end  AS [Invoice No], 
case when VPM2.InvType=19 then (VPM2.SumApplied)*-1 else VPM2.SumApplied END AS [Invoice AmountLC], 
case when VPM2.InvType=19 then (VPM2.AppliedFC)*-1 else VPM2.AppliedFC END AS [Invoice AmountFC], 
OPDN.DocNum AS [GRN No], case when OPDN.DocCur <> 'SGD'  then OPDN.DocTotalFc else OPDN.DocTotal end AS [GRN Amount],
OVPM.CounterRef as [ChequeNum],OVPM.DocCurr,case when VPM2.InvType = 18 then OPCH.NumAtCard else ORPC.NumAtCard end as [Invoice Ref No],
VPM2.DocEntry,  ----ORPC.DocNum as 'CN No', ORPC.NumAtCard as 'CN Ref No',
OPOR.DocNum AS [PO No], 
case when OPOR.DocCur <> 'SGD'  then OPOR.DocTotalFc else OPOR.DocTotal end AS [PO Amount],
OPDN.DocTotal - OPCH.DocTotal as [GRN Vs INV],
OPOR.DocTotal - OPCH.DocTotal as [PO Vs INV],

( dbo.AE_FN003_GetApprover(OPOR.DocEntry,1)) as 'Approver1',
( dbo.AE_FN003_GetApprover(OPOR.DocEntry,2)) as 'Approver2',
( dbo.AE_FN003_GetApprover(OPOR.DocEntry,3)) as 'Approver3',
NNM1.SeriesName, OVPM.DocType


FROM VPM2
RIGHT OUTER JOIN OVPM ON OVPM.DocEntry = VPM2.DocNum
Left Outer Join OPCH ON VPM2.DocEntry = OPCH.DocEntry 
Left Outer Join PCH1 ON PCH1.DocEntry=OPCH.DocEntry and ISNULL(PCH1.BaseRef,'')<>''
LEFT JOIN OPDN ON PCH1.BaseEntry=OPDN.DocEntry and PCH1.BaseType='20'
LEFT JOIN PDN1 ON PDN1.DocEntry=OPDN.DocEntry
Left JOIN OPOR ON (Case When PCH1.BaseType=22 THEN PCH1.BaseEntry ELSE PDN1.BaseEntry END) = OPOR.DocEntry
----left join RPC1 ON RPC1.BaseEntry= OPCH.DocEntry
LEFT JOIN ORPC ON ORPC.DocEntry = VPM2.DocEntry
----LEFT JOIN OPEX ON OVPM.DocEntry = OPEX.PaymDocNum
Left Join NNM1 on OVPM.Series = NNM1.Series

WHERE OVPM.DocDate between @FrmDate and  @ToDate  

--and  (isnull(OVPM.[PayMth],'') like @FrmPayMethod)
and OVPM.Canceled = 'N' ----and OVPM.docType = 'S'

GROUP BY
OVPM.DocDate , OVPM.DocNum , OVPM.DocType, OVPM.Address, OVPM.CardCode, OVPM.CardName , 
OVPM.DOCTOTAL ,OVPM.DOCTOTALFC, VPM2.DocNum , VPM2.InvType, VPM2.SumApplied , VPM2.AppliedFC, OPDN.DocNum , OPDN.DocTotal, 
OPOR.DocNum , OPOR.DocTotal,OPOR.DocTotalFC,OPCH.DocTotal,OPDN.DocTotalFC, OPOR.DOcEntry,OPDN.DocCur,OPOR.DocCur,
OVPM.CounterRef,OVPM.DocCurr,OPCH.NumAtCard ,NNM1.SeriesName,OPCH.DocNum,VPM2.DocEntry,ORPC.DocNum,ORPC.NumAtCard
ORDER BY    OVPM.DocDate,  
--OVPM.DocType,
OVPM.DocNum,OPCH.DocNum

end
GO
/****** Object:  StoredProcedure [dbo].[AE_SP003_PaidInvoicesList_backup]    Script Date: 22/08/2017 11:41:16 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

--[dbo].[AE_SP003_PaidInvoicesList]'20170131','20170131'


create PROCEDURE [dbo].[AE_SP003_PaidInvoicesList_backup]
-- Add the parameters for the stored procedure here
@FrmDate Date,
@ToDate Date
----@FrmPayMethod varchar(100),
----@ToPayMethod varchar(100)
--@WizardCode VARCHAR(20)

AS
BEGIN

	--DECLARE @PaymWizCod varchar(30)
	--Select @PaymWizCod =  OPEX.PaymWizCod from OPEX
	--Print @PaymWizCod


	
	
	select
OVPM.DocDate AS [Date Paid], OVPM.DocNum AS [Payment Voucher No], case OVPM.DocType when 'A' then '' else OVPM.CardCode END AS [Vendor Code], case OVPM.DocType when 'A' then OVPM.Address else OVPM.CardName END AS [Vendor Name], 
case when OVPM.DocCurr = 'SGD' then OVPM.DOCTOTAL else OVPM.DocTotalFC end as PaidDocTotal,
case when OVPM.DocCurr = 'SGD' then OVPM.DOCTOTAL else 0.00 end as TotalpaymentLC,
case when VPM2.InvType = 18 then OPCH.DocNum  else ORPC.DocNum end  AS [Invoice No], 
case when VPM2.InvType=19 then (VPM2.SumApplied)*-1 else VPM2.SumApplied END AS [Invoice AmountLC], 
case when VPM2.InvType=19 then (VPM2.AppliedFC)*-1 else VPM2.AppliedFC END AS [Invoice AmountFC], 
OPDN.DocNum AS [GRN No], case when OPDN.DocCur <> 'SGD'  then OPDN.DocTotalFc else OPDN.DocTotal end AS [GRN Amount],
OVPM.CounterRef as [ChequeNum],OVPM.DocCurr,case when VPM2.InvType = 18 then OPCH.NumAtCard else ORPC.NumAtCard end as [Invoice Ref No], 
VPM2.DocEntry,  ----ORPC.DocNum as 'CN No', ORPC.NumAtCard as 'CN Ref No',
OPOR.DocNum AS [PO No], 
case when OPOR.DocCur <> 'SGD'  then OPOR.DocTotalFc else OPOR.DocTotal end AS [PO Amount],
OPDN.DocTotal - OPCH.DocTotal as [GRN Vs INV],
OPOR.DocTotal - OPCH.DocTotal as [PO Vs INV],
( dbo.AE_FN003_GetApprover(OPOR.DocEntry,1)) as 'Approver1',
( dbo.AE_FN003_GetApprover(OPOR.DocEntry,2)) as 'Approver2',
( dbo.AE_FN003_GetApprover(OPOR.DocEntry,3)) as 'Approver3',
NNM1.SeriesName, OVPM.DocType

FROM VPM2
RIGHT OUTER JOIN OVPM ON OVPM.DocEntry = VPM2.DocNum
Left Outer Join OPCH ON VPM2.DocEntry = OPCH.DocEntry 
Left Outer Join PCH1 ON PCH1.DocEntry=OPCH.DocEntry and ISNULL(PCH1.BaseRef,'')<>''
LEFT JOIN OPDN ON PCH1.BaseEntry=OPDN.DocEntry and PCH1.BaseType='20'
LEFT JOIN PDN1 ON PDN1.DocEntry=OPDN.DocEntry
Left JOIN OPOR ON (Case When PCH1.BaseType=22 THEN PCH1.BaseEntry ELSE PDN1.BaseEntry END) = OPOR.DocEntry
----left join RPC1 ON RPC1.BaseEntry= OPCH.DocEntry
LEFT JOIN ORPC ON ORPC.DocEntry = VPM2.DocEntry
----LEFT OUTER JOIN OPEX ON OVPM.DocEntry = OPEX.PaymDocNum
Left Join NNM1 on OVPM.Series = NNM1.Series

WHERE OVPM.DocDate between @FrmDate and  @ToDate  

---and  OPEX.[PaymWizCod] = @WizardCode
and OVPM.Canceled = 'N' ----and OVPM.docType = 'S'

GROUP BY
OVPM.DocDate , OVPM.DocNum , OVPM.DocType, OVPM.Address, OVPM.CardCode, OVPM.CardName , 
OVPM.DOCTOTAL ,OVPM.DOCTOTALFC, VPM2.DocNum , VPM2.InvType, VPM2.SumApplied , VPM2.AppliedFC, OPDN.DocNum , OPDN.DocTotal, 
OPOR.DocNum , OPOR.DocTotal,OPOR.DocTotalFC,OPCH.DocTotal,OPDN.DocTotalFC, OPOR.DOcEntry,OPDN.DocCur,OPOR.DocCur,
OVPM.CounterRef,OVPM.DocCurr,OPCH.NumAtCard ,NNM1.SeriesName,OPCH.DocNum,VPM2.DocEntry,ORPC.DocNum,ORPC.NumAtCard

ORDER BY OVPM.DocType,OVPM.DocNum,OPCH.DocNum  --,CAST(OPOR.DOCNUM AS INTEGER)

END
	
		

GO
/****** Object:  StoredProcedure [dbo].[AE_SP003_PaidInvoicesList_Header]    Script Date: 22/08/2017 11:41:16 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


-- =============================================
-- Author:		<Umesh Patle>
-- Create date: <11-May-2015,,>
-- Description:	<PaidInvoicesList,,>
-- =============================================
--EXEC [AE_SP003_PaidInvoicesList_Header] 

CREATE PROCEDURE [dbo].[AE_SP003_PaidInvoicesList_Header]

AS
BEGIN

	select top(1)  isnull(T0.PrintHeadr,T0.CompnyName) CompnyName,T0.Phone1,T0.E_Mail ,T1.GlblLocNum, T0.Fax, T0.FreeZoneNo, T0.TaxIdNum, 
	T0.CompnyAddr,T1.Street, T1.StreetNo , T1.Block, T1.Building, T1.ZipCode , T1.City, T1.Country , T3.LogoImage,T1.IntrntAdrs,T0.RevOffice, T2.Name , 1 as LinkID 
	from OADM T0 with(nolock)   
	left outer join ADM1 T1 with(nolock) on 1=1  
	left outer join OADP T3 with(nolock) on 1=1  
	left outer join OCST T2 with(nolock) on T2.Country  =T1.Country  
		

END

GO
/****** Object:  StoredProcedure [dbo].[AE_SP003_PaidInvoicesList_with_Opex]    Script Date: 22/08/2017 11:41:16 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

--exec [dbo].[AE_SP003_PaidInvoicesList_with_Opex] '20160901','20160901',''

----exec [dbo].[AE_SP003_PaidInvoicesList_with_Opex] '20160824','20160824',''
CREATE PROCEDURE [dbo].[AE_SP003_PaidInvoicesList_with_Opex]
-- Add the parameters for the stored procedure here
@FrmDate Date,
@ToDate Date,
@FrmPayMethod varchar(100)
----@ToPayMethod varchar(100)
--@WizardCode VARCHAR(20)

AS

if @FrmPayMethod='Cheque'
BEGIN

	--DECLARE @PaymWizCod varchar(30)
	--Select @PaymWizCod =  OPEX.PaymWizCod from OPEX
	--Print @PaymWizCod

	if ISNULL(@FrmPayMethod,'')=''
		BEGIN
			set @FrmPayMethod='%'
		END

	select
OVPM.DocDate AS [Date Paid], OVPM.DocNum AS [Payment Voucher No], case OVPM.DocType when 'A' then '' else OVPM.CardCode END AS [Vendor Code], case OVPM.DocType when 'A' then OVPM.Address else OVPM.CardName END AS [Vendor Name], 
case when OVPM.DocCurr = 'SGD' then OVPM.DOCTOTAL else OVPM.DocTotalFC end as PaidDocTotal,
case when VPM2.InvType = 18 then OPCH.DocNum  else ORPC.DocNum end  AS [Invoice No], 
case when VPM2.InvType=19 then (VPM2.SumApplied)*-1 else VPM2.SumApplied END AS [Invoice AmountLC], 
case when VPM2.InvType=19 then (VPM2.AppliedFC)*-1 else VPM2.AppliedFC END AS [Invoice AmountFC], 
OPDN.DocNum AS [GRN No], case when OPDN.DocCur <> 'SGD'  then OPDN.DocTotalFc else OPDN.DocTotal end AS [GRN Amount],
OVPM.CounterRef as [ChequeNum],OVPM.DocCurr,case when VPM2.InvType = 18 then OPCH.NumAtCard else ORPC.NumAtCard end as [Invoice Ref No],
VPM2.DocEntry,  ----ORPC.DocNum as 'CN No', ORPC.NumAtCard as 'CN Ref No',
OPOR.DocNum AS [PO No], 
case when OPOR.DocCur <> 'SGD'  then OPOR.DocTotalFc else OPOR.DocTotal end AS [PO Amount],
OPDN.DocTotal - OPCH.DocTotal as [GRN Vs INV],
OPOR.DocTotal - OPCH.DocTotal as [PO Vs INV],
case when isnull(OVPM.PayMth,'')  = ''  then OVPM.CounterRef  
 when (OVPM.PayMth)  = 'Cheque' then OVPM.CounterRef  else OVPM.PayMth end as 'PayDocRef',
OVPM.PayMth,
( dbo.AE_FN003_GetApprover(OPOR.DocEntry,1)) as 'Approver1',
( dbo.AE_FN003_GetApprover(OPOR.DocEntry,2)) as 'Approver2',
( dbo.AE_FN003_GetApprover(OPOR.DocEntry,3)) as 'Approver3',
NNM1.SeriesName, OVPM.DocType


FROM VPM2
RIGHT OUTER JOIN OVPM ON OVPM.DocEntry = VPM2.DocNum
Left Outer Join OPCH ON VPM2.DocEntry = OPCH.DocEntry 
Left Outer Join PCH1 ON PCH1.DocEntry=OPCH.DocEntry and ISNULL(PCH1.BaseRef,'')<>''
LEFT JOIN OPDN ON PCH1.BaseEntry=OPDN.DocEntry and PCH1.BaseType='20'
LEFT JOIN PDN1 ON PDN1.DocEntry=OPDN.DocEntry
Left JOIN OPOR ON (Case When PCH1.BaseType=22 THEN PCH1.BaseEntry ELSE PDN1.BaseEntry END) = OPOR.DocEntry
----left join RPC1 ON RPC1.BaseEntry= OPCH.DocEntry
LEFT JOIN ORPC ON ORPC.DocEntry = VPM2.DocEntry
----LEFT JOIN OPEX ON OVPM.DocEntry = OPEX.PaymDocNum
Left Join NNM1 on OVPM.Series = NNM1.Series

WHERE OVPM.DocDate between @FrmDate and  @ToDate  

and  (isnull(OVPM.[PayMth],'') like @FrmPayMethod or CheckSum <> 0)
and OVPM.Canceled = 'N' ----and OVPM.docType = 'S'

GROUP BY
OVPM.DocDate , OVPM.DocNum , OVPM.DocType, OVPM.Address, OVPM.CardCode, OVPM.CardName , 
OVPM.DOCTOTAL ,OVPM.DOCTOTALFC, VPM2.DocNum , VPM2.InvType, VPM2.SumApplied , VPM2.AppliedFC, OPDN.DocNum , OPDN.DocTotal, 
OPOR.DocNum , OPOR.DocTotal,OPOR.DocTotalFC,OPCH.DocTotal,OPDN.DocTotalFC, OPOR.DOcEntry,OPDN.DocCur,OPOR.DocCur,
OVPM.CounterRef,OVPM.DocCurr,OPCH.NumAtCard ,NNM1.SeriesName,OPCH.DocNum,VPM2.DocEntry,ORPC.DocNum,ORPC.NumAtCard,
OVPM.PayMth

ORDER BY OVPM.DocDate,
----OVPM.DocType,
-OVPM.DocNum,OPCH.DocNum  --,CAST(OPOR.DOCNUM AS INTEGER)

END

else
BEGIN

	if ISNULL(@FrmPayMethod,'')=''
		BEGIN
			set @FrmPayMethod='%'
		END

	select
OVPM.DocDate AS [Date Paid], OVPM.DocNum AS [Payment Voucher No], case OVPM.DocType when 'A' then '' else OVPM.CardCode END AS [Vendor Code], case OVPM.DocType when 'A' then OVPM.Address else OVPM.CardName END AS [Vendor Name], 
case when OVPM.DocCurr = 'SGD' then OVPM.DOCTOTAL else OVPM.DocTotalFC end as PaidDocTotal,
case when VPM2.InvType = 18 then OPCH.DocNum  else ORPC.DocNum end  AS [Invoice No], 
case when VPM2.InvType=19 then (VPM2.SumApplied)*-1 else VPM2.SumApplied END AS [Invoice AmountLC], 
case when VPM2.InvType=19 then (VPM2.AppliedFC)*-1 else VPM2.AppliedFC END AS [Invoice AmountFC], 
OPDN.DocNum AS [GRN No], case when OPDN.DocCur <> 'SGD'  then OPDN.DocTotalFc else OPDN.DocTotal end AS [GRN Amount],
OVPM.CounterRef as [ChequeNum],OVPM.DocCurr,case when VPM2.InvType = 18 then OPCH.NumAtCard else ORPC.NumAtCard end as [Invoice Ref No],
VPM2.DocEntry,  ----ORPC.DocNum as 'CN No', ORPC.NumAtCard as 'CN Ref No',
OPOR.DocNum AS [PO No], 
case when OPOR.DocCur <> 'SGD'  then OPOR.DocTotalFc else OPOR.DocTotal end AS [PO Amount],
OPDN.DocTotal - OPCH.DocTotal as [GRN Vs INV],
OPOR.DocTotal - OPCH.DocTotal as [PO Vs INV],
case when isnull(OVPM.PayMth,'')  = '' then OVPM.CounterRef 
when (OVPM.PayMth)  = 'Cheque' then OVPM.CounterRef  else OVPM.PayMth end as 'PayDocRef',
OVPM.PayMth,
( dbo.AE_FN003_GetApprover(OPOR.DocEntry,1)) as 'Approver1',
( dbo.AE_FN003_GetApprover(OPOR.DocEntry,2)) as 'Approver2',
( dbo.AE_FN003_GetApprover(OPOR.DocEntry,3)) as 'Approver3',
NNM1.SeriesName, OVPM.DocType


FROM VPM2
RIGHT OUTER JOIN OVPM ON OVPM.DocEntry = VPM2.DocNum
Left Outer Join OPCH ON VPM2.DocEntry = OPCH.DocEntry 
Left Outer Join PCH1 ON PCH1.DocEntry=OPCH.DocEntry and ISNULL(PCH1.BaseRef,'')<>''
LEFT JOIN OPDN ON PCH1.BaseEntry=OPDN.DocEntry and PCH1.BaseType='20'
LEFT JOIN PDN1 ON PDN1.DocEntry=OPDN.DocEntry
Left JOIN OPOR ON (Case When PCH1.BaseType=22 THEN PCH1.BaseEntry ELSE PDN1.BaseEntry END) = OPOR.DocEntry
----left join RPC1 ON RPC1.BaseEntry= OPCH.DocEntry
LEFT JOIN ORPC ON ORPC.DocEntry = VPM2.DocEntry
----LEFT JOIN OPEX ON OVPM.DocEntry = OPEX.PaymDocNum
Left Join NNM1 on OVPM.Series = NNM1.Series

WHERE OVPM.DocDate between @FrmDate and  @ToDate  

and  (isnull(OVPM.[PayMth],'') like @FrmPayMethod)
and OVPM.Canceled = 'N' ----and OVPM.docType = 'S'

GROUP BY
OVPM.DocDate , OVPM.DocNum , OVPM.DocType, OVPM.Address, OVPM.CardCode, OVPM.CardName , 
OVPM.DOCTOTAL ,OVPM.DOCTOTALFC, VPM2.DocNum , VPM2.InvType, VPM2.SumApplied , VPM2.AppliedFC, OPDN.DocNum , OPDN.DocTotal, 
OPOR.DocNum , OPOR.DocTotal,OPOR.DocTotalFC,OPCH.DocTotal,OPDN.DocTotalFC, OPOR.DOcEntry,OPDN.DocCur,OPOR.DocCur,
OVPM.CounterRef,OVPM.DocCurr,OPCH.NumAtCard ,NNM1.SeriesName,OPCH.DocNum,VPM2.DocEntry,ORPC.DocNum,ORPC.NumAtCard,
OVPM.PayMth

ORDER BY OVPM.DocDate, 
----OVPM.DocType,
OVPM.DocNum,OPCH.DocNum  --,CAST(OPOR.DOCNUM AS INTEGER)
END
	
		

GO
/****** Object:  StoredProcedure [dbo].[AE_SP005_BUDGET_COMMITTEDAMOUNT]    Script Date: 22/08/2017 11:41:16 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

----AE_SP005_BUDGET_COMMITTEDAMOUNT'PWCL','2015','True','Test123',''

--AE_SP005_BUDGET_COMMITTEDAMOUNT'PWCL',2016,'True','Test123',''
--AE_SP005_BUDGET_COMMITTEDAMOUNT'PWCL','2016','True','Test789','71111100'
--AE_SP006_BUDGET_ACTUALSPEND'PWCL','2014','True','Prj002','71111100'
--[AE_SP005_BUDGET_COMMITTEDAMOUNT_Test]'PWCL','2014','True','Prj002','71111100'
--[AE_SP005_BUDGET_COMMITTEDAMOUNT]'PWCL','2016','True','UAT1x800','71181300'
--[AE_SP005_BUDGET_COMMITTEDAMOUNT_Test]'PWCL','2016','False','52723','71181300'
--select sum( T0.U_BudAmount ) from [@AB_PROJECTBUDGET] T0 where T0.U_PrjCode = 'Test789'





CREATE PROCEDURE [dbo].[AE_SP005_BUDGET_COMMITTEDAMOUNT]

@Entity as  Varchar(100),
@Year as varchar(10),
@Flag as Varchar(100),
@Dimension as Varchar(100),
@GLCode as varchar(100)

as

DECLARE @SQL as varchar(max)
DECLARE @Cond as varchar(1000)

Declare @DateF varchar(30)
Declare @DateT varchar(30)


set @DateF = cast(cast(@year as numeric ) -1 AS VARCHAR) +'0701'
 set @DateT =  @year + '0630'


begin

select top(1) T1.LineTotal [Column1], T1.LineTotal [Column2], T1.LineTotal [Column3], T1.LineTotal [Column4], T1.LineTotal [Column5],
 T1.LineTotal [Column6], T1.LineTotal [Column7], T1.LineTotal [Column8]
 into #Tmp from DRF1 T1 where T1.DocEntry = 1 

delete from #tmp

   set @Cond = '(''22'',''1470000113'')'
  

--- Add PO Draft with status Open Raised manually - save as draft
set @sql =
' isnull((SELECT sum(T1.LineTotal ) LineTotal
 FROM '+@Entity+' ..ODRF T0  INNER JOIN '+@Entity+' ..DRF1 T1 ON T0.[DocEntry] = T1.[DocEntry] 
left outer join  '+@Entity+' ..OWDD T2 on T0.DocEntry = T2.DocEntry 
left outer join  '+@Entity+' .. OWTM T3 on T2.[WtmCode] = T3.[WtmCode] 
where T1.LineStatus  = ''O'' and isnull(T2.[Status],'''') = '''' and T0.DocStatus = ''O'' and T0.ObjType in '+ @Cond +'
and T0.DocDate between '''+@datef+''' and '''+@DateT+''' and isnull(T3.[Active],''Y'') = ''Y'' and
(case when '''+@Flag+''' = ''True'' then isnull(T1.Project,'''') else  isnull(T1.U_AB_NONPROJECT ,'''')  end) = '''+ @Dimension +'''
and (case when '''+@Flag+''' = ''False'' then isnull(T1.AcctCode ,'''') else '''+ @GLCode +'''  end) = '''+ @GLCode +'''
and (case when '''+@Flag+''' = ''False'' then isnull(T1.Project ,'''') else ''''  end) = ''''
),0) Column1 , '


--- Add PO Draft with status Open  - Pending for Approval / Approved but not converted to PO
set @sql +=
'isnull((SELECT sum(T1.LineTotal) LineTotal  
 FROM '+@Entity+' ..ODRF T0  INNER JOIN '+@Entity+' ..DRF1 T1 ON T0.[DocEntry] = T1.[DocEntry] 
left outer join  '+@Entity+' ..OWDD T2 on T0.DocEntry = T2.DocEntry 
left outer join  '+@Entity+' .. OWTM T3 on T2.[WtmCode] = T3.[WtmCode]
where T1.LineStatus = ''O'' and  isnull(T2.[Status],'''') <> ''N'' and T0.DocStatus = ''O'' and T0.ObjType in '+ @Cond +'
and T0.DocDate between '''+@datef+''' and '''+@DateT+''' and isnull(T3.[Active],'''') = ''Y''
and (case when '''+@Flag+''' = ''True'' then isnull(T1.Project,'''') else  isnull(T1.U_AB_NONPROJECT ,'''')  end) = '''+ @Dimension +'''
and (case when '''+@Flag+''' = ''False'' then isnull(T1.AcctCode ,'''') else '''+ @GLCode +'''  end) = '''+ @GLCode +'''
and (case when '''+@Flag+''' = ''False'' then isnull(T1.Project ,'''') else ''''  end) = ''''
 ),0) Column2 ,'

 --- Add PR Document with status Open
set @sql +=
'isnull((select sum( case when isnull(T2.LineTotal,0) = 0 then  T1.LineTotal  else 0 end  )  from '+@Entity+' ..OPRQ T0 join '+@Entity+' ..PRQ1 T1 on T0.DocEntry = T1.DocEntry 
left outer join DRF1 T2 on T2.BaseEntry = T1.DocEntry and T2.BaseLine = T1.LineNum and T2.BaseType = ''1470000113''
left outer join ODRF T3 on T2.DocEntry = T3.Docentry and T3.DocStatus = ''O''
where T1.LineStatus = ''O'' and T0.DocStatus = ''O''
and T0.DocDate between '''+@datef+''' and '''+@DateT+''' 
and (case when '''+@Flag+''' = ''True'' then isnull(T1.Project,'''') else  isnull(T1.U_AB_NONPROJECT ,'''')  end) = '''+ @Dimension +'''
and (case when '''+@Flag+''' = ''False'' then isnull(T1.AcctCode ,'''') else '''+ @GLCode +'''  end) = '''+ @GLCode +'''
and (case when '''+@Flag+''' = ''False'' then isnull(T1.Project ,'''') else ''''  end) = ''''
),0) Column3,'

--- Add PO Document with status Open
set @sql +=
'isnull((select sum(T1.LineTotal) LineTotal  from '+@Entity+' ..OPOR T0 join '+@Entity+' ..POR1 T1 on T0.DocEntry = T1.DocEntry where T1.LineStatus = ''O'' and T0.DocStatus = ''O''
and T0.DocDate between '''+@datef+''' and '''+@DateT+''' 
and (case when '''+@Flag+''' = ''True'' then isnull(T1.Project,'''') else  isnull(T1.U_AB_NONPROJECT ,'''')  end) = '''+ @Dimension +'''
and (case when '''+@Flag+''' = ''False'' then isnull(T1.AcctCode ,'''') else '''+ @GLCode +'''  end) = '''+ @GLCode +'''
and (case when '''+@Flag+''' = ''False'' then isnull(T1.Project ,'''') else ''''  end) = ''''
),0) Column4,'

--- Add PO Document with status Closed But GRN in Open Status
set @sql +=
'isnull((SELECT sum(T1.[LineTotal]) LineTotal FROM '+@Entity+' ..OPOR T0  INNER JOIN '+@Entity+' ..POR1 T1 ON T0.[DocEntry] = T1.[DocEntry] join 
 '+@Entity+' ..PDN1 T2 on T2.BaseEntry = T1.DocEntry and T2.BaseLine = T1.LineNum  INNER JOIN '+@Entity+' ..OPDN T3 ON T2.[DocEntry] = T3.[DocEntry] WHERE T2.[LineStatus] = ''O''
 and T0.DocDate between '''+@datef+''' and '''+@DateT+''' 
 and (case when '''+@Flag+''' = ''True'' then isnull(T1.Project,'''') else  isnull(T1.U_AB_NONPROJECT ,'''')  end) = '''+ @Dimension +'''
 and (case when '''+@Flag+''' = ''False'' then isnull(T1.AcctCode ,'''') else '''+ @GLCode +'''  end) = '''+ @GLCode +'''
 and (case when '''+@Flag+''' = ''False'' then isnull(T1.Project ,'''') else ''''  end) = ''''
 ),0) Column5,'

--- Less PO Draft with status Open  - Approval status Rejected
set @sql +=
'isnull((SELECT sum(T1.LineTotal) LineTotal  
 FROM '+@Entity+' ..ODRF T0  INNER JOIN '+@Entity+' .. DRF1 T1 ON T0.[DocEntry] = T1.[DocEntry] 
left outer join  '+@Entity+' ..OWDD T2 on T0.DocEntry = T2.DocEntry 
left outer join  '+@Entity+' .. OWTM T3 on T2.[WtmCode] = T3.[WtmCode]
where T1.LineStatus = ''O'' and isnull(T2.[Status],'''') = ''N'' and T0.DocStatus = ''O'' and T0.ObjType in '+ @Cond +'
and T0.DocDate between '''+@datef+''' and '''+@DateT+''' and isnull( T3.[Active],'''') = ''Y''
and (case when '''+@Flag+''' = ''True'' then isnull(T1.Project,'''') else  isnull(T1.U_AB_NONPROJECT ,'''')  end) = '''+ @Dimension +'''
and (case when '''+@Flag+''' = ''False'' then isnull(T1.AcctCode ,'''') else '''+ @GLCode +'''  end) = '''+ @GLCode +'''
and (case when '''+@Flag+''' = ''False'' then isnull(T1.Project ,'''') else ''''  end) = ''''
),0) Column6,'

--- Less PO Document with status Cancel  - User manually Cancel
set @sql +=
'isnull((select sum(T1.LineTotal) LineTotal   from '+@Entity+' ..OPOR T0 join '+@Entity+' ..POR1 T1 on T0.DocEntry = T1.DocEntry where T1.LineStatus = ''C'' and T0.CANCELED = ''Y''
and T0.DocDate between '''+@datef+''' and '''+@DateT+''' 
and (case when '''+@Flag+''' = ''True'' then isnull(T1.Project,'''') else  isnull(T1.U_AB_NONPROJECT ,'''')  end) = '''+ @Dimension +'''
and (case when '''+@Flag+''' = ''False'' then isnull(T1.AcctCode ,'''') else '''+ @GLCode +'''  end) = '''+ @GLCode +'''
and (case when '''+@Flag+''' = ''False'' then isnull(T1.Project ,'''') else ''''  end) = ''''
),0) Column7,'

--- Less PO Document with status Close  - User manually Close
set @sql +=
'isnull((select sum(T1.LineTotal) LineTotal   from '+@Entity+' ..OPOR T0 join '+@Entity+' ..POR1 T1 on T0.DocEntry = T1.DocEntry where T1.LineStatus = ''C'' and T0.CANCELED = ''N'' 
and T1.TargetType = -1
and T0.DocDate between '''+@datef+''' and '''+@DateT+''' 
and (case when '''+@Flag+''' = ''True'' then isnull(T1.Project,'''') else  isnull(T1.U_AB_NONPROJECT ,'''')  end) = '''+ @Dimension +'''
and (case when '''+@Flag+''' = ''False'' then isnull(T1.AcctCode ,'''') else '''+ @GLCode +'''  end) = '''+ @GLCode +'''
and (case when '''+@Flag+''' = ''False'' then isnull(T1.Project ,'''') else ''''  end) = ''''
),0) Column8'

insert into #tmp
exec ('Select' + @SQL)


--select (T1.Column1 + T1.Column2 + T1.Column3 + T1.Column4 + T1.Column5) - T1.Column6 - T1.Column7 - T1.Column8 [CommittedAmount] from #tmp T1
select (T1.Column2 + T1.Column3 + T1.Column4 + T1.Column5) [CommittedAmount] from #tmp T1

end

GO
/****** Object:  StoredProcedure [dbo].[AE_SP006_BUDGET_ACTUALSPEND]    Script Date: 22/08/2017 11:41:16 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO



--AE_SP006_BUDGET_ACTUALSPEND'PWCL',2014,'True','Prj002',''



create PROCEDURE [dbo].[AE_SP006_BUDGET_ACTUALSPEND]

@Entity as  Varchar(100),
@Year as varchar(10),
@Flag as Varchar(100),
@Dimension as Varchar(100),
@GLCode as varchar(100)

as

DECLARE @SQL as varchar(max)

Declare @DateF varchar(30)
Declare @DateT varchar(30)

begin

set @DateF = cast(cast(@year as numeric ) -1 AS VARCHAR) +'0701'
 set @DateT =  @year + '0630'

select top(1) T1.LineTotal [Column1], T1.LineTotal [Column2], T1.LineTotal [Column3], T1.LineTotal [Column4], T1.LineTotal [Column5]
 into #Tmp from DRF1 T1 where T1.DocEntry = 1 

delete from #tmp


--- Add AP INVOICE with status Open - Base Document PO,GRPO & Direct
set @SQL =
' isnull((SELECT sum(T1.[LineTotal]) FROM '+@Entity+' ..OPCH T0  INNER JOIN '+@Entity+' ..PCH1 T1 ON T0.[DocEntry] = T1.[DocEntry]
WHERE t1.LineStatus = ''O'' and T0.DocStatus = ''O'' AND T1.BaseType IN (''20'',''22'',''-1'')
and T0.DocDate between '''+@datef+''' and '''+@DateT+''' 
and (case when '''+@Flag+''' = ''True'' then isnull(T1.Project,'''') else  isnull(T1.U_AB_NONPROJECT ,'''')  end) = '''+ @Dimension +'''
and (case when '''+@Flag+''' = ''False'' then isnull(T1.AcctCode ,'''') else '''+ @GLCode +'''  end) = '''+ @GLCode +'''
and (case when '''+@Flag+''' = ''False'' then isnull(T1.Project ,'''') else ''''  end) = '''')
,0) Column1 , '

--- Add JE & Direct
set @SQL +=
'isnull((SELECT sum(T2.[Debit] - T2.[Credit])  FROM '+@Entity+' ..[OJDT]  T1 INNER JOIN '+@Entity+' ..JDT1 T2 
ON T1.[TransId] = T2.[TransId] WHERE T1.[TransType] = 30
and T1.[TransId] NOT IN (SELECT StornoToTr FROM ojdt WHERE ISNULL(StornoToTr,'''') <> '''') AND T1.[StornoToTr] IS NULL 
and T1.TaxDate between '''+@datef+''' and '''+@DateT+'''
and (case when '''+@Flag+''' = ''True'' then isnull(T2.Project,'''') else  isnull(T2.U_AB_NONPROJECT ,'''')  end) = '''+ @Dimension +'''
and (case when '''+@Flag+''' = ''False'' then isnull(T2.Account ,'''') else '''+ @GLCode +'''  end) = '''+ @GLCode +'''
and (case when '''+@Flag+''' = ''False'' then isnull(T2.Project ,'''') else ''''  end) = '''')
,0) Column2, '

--- Add  AP Invoice with link AP credit memo 

set @sql +=
'isnull((SELECT sum(T1.[LineTotal]) FROM '+@Entity+' ..OPCH T0  INNER JOIN '+@Entity+' ..PCH1 T1 ON T0.[DocEntry] = T1.[DocEntry]
join '+@Entity+' ..RPC1 T2 on T2.BaseEntry = T1.DocEntry and T2.BaseLine = T1.LineNum INNER JOIN '+@Entity+' ..ORPC T3 ON T2.[DocEntry] = T3.[DocEntry] 
WHERE T1.[LineStatus]  = ''C'' 
and T0.DocDate between '''+@datef+''' and '''+@DateT+''' 
and (case when '''+@Flag+''' = ''True'' then isnull(T1.Project,'''') else  isnull(T1.U_AB_NONPROJECT ,'''')  end) = '''+ @Dimension +'''
and (case when '''+@Flag+''' = ''False'' then isnull(T1.AcctCode ,'''') else '''+ @GLCode +'''  end) = '''+ @GLCode +'''
and (case when '''+@Flag+''' = ''False'' then isnull(T1.Project ,'''') else ''''  end) = ''''),0)  Column3,'


--- Less AP Credit memo standalone
set @SQL +=
'isnull((SELECT sum(T1.[LineTotal]) FROM '+@Entity+' ..ORPC T0  INNER JOIN '+@Entity+' ..RPC1 T1 ON T0.[DocEntry] = T1.[DocEntry] WHERE T1.[LineStatus]  = ''O'' and [BaseType] = -1
and T0.DocDate between '''+@datef+''' and '''+@DateT+''' 
and (case when '''+@Flag+''' = ''True'' then isnull(T1.Project,'''') else  isnull(T1.U_AB_NONPROJECT ,'''')  end) = '''+ @Dimension +'''
and (case when '''+@Flag+''' = ''False'' then isnull(T1.AcctCode ,'''') else '''+ @GLCode +'''  end) = '''+ @GLCode +'''
and (case when '''+@Flag+''' = ''False'' then isnull(T1.Project ,'''') else ''''  end) = '''')
,0) Column4 , '

--- Less AP INVOICE Cancell
set @SQL +=
'isnull((SELECT sum(T1.[LineTotal]) FROM '+@Entity+' ..OPCH T0  INNER JOIN '+@Entity+' ..PCH1 T1 ON T0.[DocEntry] = T1.[DocEntry]
WHERE t1.LineStatus = ''C'' and T0.DocStatus = ''C'' and T0.CANCELED = ''y''
and T0.DocDate between '''+@datef+''' and '''+@DateT+''' 
and (case when '''+@Flag+''' = ''True'' then isnull(T1.Project,'''') else  isnull(T1.U_AB_NONPROJECT ,'''')  end) = '''+ @Dimension +'''
and (case when '''+@Flag+''' = ''False'' then isnull(T1.AcctCode ,'''') else '''+ @GLCode +'''  end) = '''+ @GLCode +'''
and (case when '''+@Flag+''' = ''False'' then isnull(T1.Project ,'''') else ''''  end) = '''')
,0) Column5'
 
 

insert into #tmp
EXEC ('Select' + @SQL)



--select (T1.Column1 + T1.Column2  + T1.Column3 ) - T1.Column4 - T1.Column5   [ActualSpend] from #tmp T1

select (T1.Column1 + T1.Column2 ) - T1.Column4    [ActualSpend] from #tmp T1

end

GO
/****** Object:  StoredProcedure [dbo].[AE_SP007_BUDGETREPORT]    Script Date: 22/08/2017 11:41:16 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


--select * from [@AB_PROJECTBUDGET] where U_PrjCode = 'Test5000'
--select * from [@AB_PROJECTBUDGET] where U_BUCode = '52723' and U_Account between '71181300' and '71181410'

--[AE_SP007_BUDGETREPORT_Test1]'2016','52723','52723','%','%','71181300','71181410'
--go
--[AE_SP007_BUDGETREPORT_Test1]'2017','52761','52761','%','%','11111100','99999999'

CREATE PROCEDURE [dbo].[AE_SP007_BUDGETREPORT]

@Year as varchar(10),
@NonProjectF as varchar(100),
@NonProjectT as varchar(100),
@ProjectF as varchar(100),
@ProjectT as varchar(100),
@GlcodeF as varchar(30),
@GlcodeT as varchar(30)

AS

Declare @MainBudget as varchar(200)
Declare @ReforecastBudget as varchar(200)
Declare @Condition as varchar(300)
Declare @Condition1 as varchar(300)
Declare @Condition2 as varchar(300)
Declare @SQL as varchar(max)

Declare @DateF varchar(30)
Declare @DateT varchar(30)

BEGIN

if @ProjectF = 'NA'
begin
 set @ProjectF = '%'
end

if @ProjectT = 'NA'
begin
 set @ProjectT = '%'
end

if @NonProjectF  = 'NA'
begin
 set @NonProjectF = '%'
end

if @NonProjectT  = 'NA'
begin
 set @NonProjectT = '%'
end
set @DateF = cast(cast(@year as numeric ) -1 AS VARCHAR) +'0701'
set @DateT =  @year + '0630'

SELECT @MainBudget = T0.[Name] FROM PWCL ..OBGS T0 WHERE year( T0.[FinancYear] ) = cast(cast(@Year as numeric) -1 as varchar) and UPPER( T0.[U_AB_MAINBUDGET]) = 'YES'
SELECT @ReforecastBudget = T0.[Name] FROM PWCL ..OBGS T0 WHERE year( T0.[FinancYear] ) = cast(cast(@Year as numeric) -1 as varchar) and   UPPER(T0.[U_AB_ACTIVE])  = 'YES'
--SELECT @MainBudget, @ReforecastBudget
--------  Condition Filter

if @ProjectF = '%' and @ProjectT = '%' and  @NonProjectF = '%' and @NonProjectT = '%'
 begin
   set @condition = 'isnull(T0.U_PrjCode,'''') LIKE ''%'''
   set @condition1 = 'isnull(T1.Project,'''') LIKE ''%'''
    set @condition2 = 'isnull(T2.Project,'''') LIKE ''%'''
   set @condition += ' and isnull(T0.U_BUCode,'''') LIKE ''%'''
   set @condition1 += ' and isnull(T1.U_AB_NONPROJECT,'''') LIKE ''%'''
    set @condition2 += ' and isnull(T2.U_AB_NONPROJECT,'''') LIKE ''%'''
 end
else if @ProjectF <> '%' and @ProjectT <> '%' and @NonProjectF <> '%' and @NonProjectT <> '%'
 begin
  set @condition = 'isnull(T0.U_PrjCode,'''') between '''+ @ProjectF +''' and '''+ @ProjectT +''''
  set @condition1 = 'isnull(T1.Project,'''') between '''+ @ProjectF +''' and '''+ @ProjectT +''''
  set @condition2 = 'isnull(T2.Project,'''') between '''+ @ProjectF +''' and '''+ @ProjectT +''''
  set @condition += ' and isnull(T0.U_BUCode,'''') between '''+ @NonProjectF  +''' and '''+ @NonProjectT  +''''
  set @condition1 += ' and isnull(T1.U_AB_NONPROJECT,'''') between '''+ @NonProjectF  +''' and '''+ @NonProjectT  +''' '
  set @condition2 += ' and isnull(T2.U_AB_NONPROJECT,'''') between '''+ @NonProjectF  +''' and '''+ @NonProjectT  +''' '
 end
 else if @ProjectF <> '%' and @ProjectT <> '%' and @NonProjectF = '%' and @NonProjectT = '%'
 begin
  set @condition = 'isnull(T0.U_PrjCode,'''') between '''+ @ProjectF +''' and '''+ @ProjectT +'''' --and  isnull(T0.U_BUCode,'''') = '''''
  set @condition1 = 'isnull(T1.Project,'''') between '''+ @ProjectF +''' and '''+ @ProjectT +'''' --and  isnull(T1.U_AB_NONPROJECT,'''') = '''''
  set @condition2 = 'isnull(T2.Project,'''') between '''+ @ProjectF +''' and '''+ @ProjectT +'''' --and  isnull(T1.U_AB_NONPROJECT,'''') = '''''
 end
 else
   begin
   set @condition = '  isnull(T0.U_BUCode,'''') between '''+ @NonProjectF  +''' and '''+ @NonProjectT  +''''
  set @condition1 = '  isnull(T1.U_AB_NONPROJECT,'''') between '''+ @NonProjectF  +''' and '''+ @NonProjectT  +''' and isnull(T1.Project,'''') = '''''
  set @condition2 = '  isnull(T2.U_AB_NONPROJECT,'''') between '''+ @NonProjectF  +''' and '''+ @NonProjectT  +''' and isnull(T2.Project,'''') = '''''
 end

 
-----  Project budget Table
-----  Fetching the Information based on the condition 

select top(1) [U_BUCode],  [U_PrjCode],  [U_Account],  case when isnull(U_PrjCode ,'') = '' then 'BU' else 'PRJ' end [Cat] INTO #TMPBUDGET1 From PWCL ..[@AB_PROJECTBUDGET]
Delete from #TMPBUDGET1

select top(1) [U_BUCode],  [U_PrjCode],  [U_Account],  case when isnull(U_PrjCode ,'') = '' then 'BU' else 'PRJ' end [Cat] INTO #TMPBUDGET2 From PWCL ..[@AB_PROJECTBUDGET]
Delete from #TMPBUDGET2

select top(1) [U_BUCode],  [U_PrjCode],  [U_Account],  case when isnull(U_PrjCode ,'') = '' then 'BU' else 'PRJ' end [Cat] INTO #TMPBUDGET3 From PWCL ..[@AB_PROJECTBUDGET]
Delete from #TMPBUDGET3

select top(1) [U_BUCode],  [U_PrjCode],  [U_Account],  case when isnull(U_PrjCode ,'') = '' then 'BU' else 'PRJ' end [Cat] INTO #TMPBUDGET4 From PWCL ..[@AB_PROJECTBUDGET]
Delete from #TMPBUDGET4


set @sql = '
SELECT isnull(T0.[U_BUCode],'''') [U_BUCode], isnull(T0.[U_PrjCode],'''') [U_PrjCode], T0.[U_Account] [U_Account] ,
case when isnull(T0.U_PrjCode ,'''') = '''' then ''BU'' else ''PRJ'' end [Cat]
FROM PWCL ..[@AB_PROJECTBUDGET]  T0
WHERE '+ @condition +' and isnull(T0.U_Account ,'''') between '''+@GlcodeF+''' and '''+@GlcodeT+'''
group by T0.[U_BUCode], T0.[U_PrjCode] , T0.[U_Account] '


insert into #TMPBUDGET1
exec (@sql)

------ PR / PO Draf Table 
set @SQL = 'SELECT isnull(T1.U_AB_NONPROJECT,'''') [U_BUCode], isnull(T1.Project,'''') [U_PrjCode], T1.AcctCode [U_Account],
case when isnull(T1.Project ,'''') = '''' then ''BU'' else ''PRJ'' end [Cat]  
 FROM ODRF T0  INNER JOIN DRF1 T1 ON T0.[DocEntry] = T1.[DocEntry] 
left outer join  OWDD T2 on T0.DocEntry = T2.DocEntry 
left outer join   OWTM T3 on T2.[WtmCode] = T3.[WtmCode]
where T1.LineStatus = ''O'' and isnull(T2.[Status],''N'') <> ''N'' and T0.DocStatus = ''O'' and T0.ObjType in (''22'',''1470000113'')
and T0.DocDate between '''+@datef+''' and '''+@DateT+''' and isnull(T3.[Active],''Y'') = ''Y''
and '+ @Condition1 +'
and isnull(T1.AcctCode ,'''') between '''+@GlcodeF+''' and '''+@GlcodeT+''' '

insert into #TMPBUDGET2 
EXEC(@sql)

------- PO 
set @SQL = 'select isnull(T1.U_AB_NONPROJECT,'''') [U_BUCode], isnull(T1.Project,'''') [U_PrjCode], T1.AcctCode [U_Account],
case when isnull(T1.Project ,'''') = '''' then ''BU'' else ''PRJ'' end [Cat] 
  from OPOR T0 join POR1 T1 on T0.DocEntry = T1.DocEntry where T1.LineStatus = ''O'' and T0.DocStatus = ''O''
and T0.DocDate between '''+@datef+''' and '''+@DateT+'''
and  '+ @Condition1 +'
and isnull(T1.AcctCode ,'''') between '''+@GlcodeF+''' and '''+@GlcodeT+''' '

----- PO status is closed but GRN is in open status
set @SQL +=
'union all SELECT isnull(T1.U_AB_NONPROJECT,'''') [U_BUCode], isnull(T1.Project,'''') [U_PrjCode], T1.AcctCode [U_Account],
case when isnull(T1.Project ,'''') = '''' then ''BU'' else ''PRJ'' end [Cat] FROM OPOR T0  INNER JOIN POR1 T1 ON T0.[DocEntry] = T1.[DocEntry] join 
 PDN1 T2 on T2.BaseEntry = T1.DocEntry and T2.BaseLine = T1.LineNum  INNER JOIN OPDN T3 ON T2.[DocEntry] = T3.[DocEntry] WHERE T2.[LineStatus] = ''O''
 and  T0.DocDate between '''+@datef+''' and '''+@DateT+'''
and  '+ @Condition1 +'
and isnull(T1.AcctCode ,'''') between '''+@GlcodeF+''' and '''+@GlcodeT+''' '


--- Add AP INVOICE with status Open - Base Document PO,GRPO & Direct
set @SQL +=
'union all SELECT isnull(T1.U_AB_NONPROJECT,'''') [U_BUCode], isnull(T1.Project,'''') [U_PrjCode], T1.AcctCode [U_Account],
case when isnull(T1.Project ,'''') = '''' then ''BU'' else ''PRJ'' end [Cat] FROM OPCH T0  INNER JOIN PCH1 T1 ON T0.[DocEntry] = T1.[DocEntry]
WHERE t1.LineStatus = ''O'' and T0.DocStatus = ''O'' AND T1.BaseType IN (''20'',''22'',''-1'')
and T0.DocDate between '''+@datef+''' and '''+@DateT+'''
and  '+ @Condition1 +'
and isnull(T1.AcctCode ,'''') between '''+@GlcodeF+''' and '''+@GlcodeT+''' '

--- Add JE & Direct
set @SQL +=
'union all SELECT isnull(T2.U_AB_NONPROJECT,'''') [U_BUCode], isnull(T2.Project,'''') [U_PrjCode], T2.Account [U_Account],
case when isnull(T2.Project ,'''') = '''' then ''BU'' else ''PRJ'' end [Cat]  FROM [OJDT]  T1 INNER JOIN JDT1 T2 
ON T1.[TransId] = T2.[TransId] WHERE T1.[TransType] = 30
and T1.[TransId] NOT IN (SELECT StornoToTr FROM ojdt WHERE ISNULL(StornoToTr,'''') <> '''') AND T1.[StornoToTr] IS NULL 
and T1.TaxDate between '''+@datef+''' and '''+@DateT+'''
and  '+ @Condition2 +'
and isnull(T2.Account ,'''') between '''+@GlcodeF+''' and '''+@GlcodeT+''' '

--- Less AP Credit memo standalone
set @SQL +=
'union all SELECT isnull(T1.U_AB_NONPROJECT,'''') [U_BUCode], isnull(T1.Project,'''') [U_PrjCode], T1.AcctCode [U_Account],
case when isnull(T1.Project ,'''') = '''' then ''BU'' else ''PRJ'' end [Cat] FROM ORPC T0  INNER JOIN RPC1 T1 ON T0.[DocEntry] = T1.[DocEntry] WHERE T1.[LineStatus]  = ''O'' and [BaseType] = -1
and T0.DocDate between '''+@datef+''' and '''+@DateT+'''
and  '+ @Condition1 +'
and isnull(T1.AcctCode ,'''') between '''+@GlcodeF+''' and '''+@GlcodeT+''' '

insert into #TMPBUDGET3 
EXEC(@sql) 
------- PR
set @SQL ='
select  isnull(T1.U_AB_NONPROJECT,'''') [U_BUCode], isnull(T1.Project,'''') [U_PrjCode], T1.AcctCode [U_Account],
case when isnull(T1.Project ,'''') = '''' then ''BU'' else ''PRJ'' end [Cat] 
from OPRQ T0 join PRQ1 T1 on T0.DocEntry = T1.DocEntry 
left outer join DRF1 T2 on T2.BaseEntry = T1.DocEntry and T2.BaseLine = T1.LineNum and T2.BaseType = ''1470000113''
left outer join ODRF T3 on T2.DocEntry = T3.Docentry and T3.DocStatus = ''O''
where T1.LineStatus = ''O'' and T0.DocStatus = ''O''
and T0.DocDate between '''+@datef+''' and '''+@DateT+'''
and  '+ @Condition1 +'
and isnull(T1.AcctCode ,'''') between '''+@GlcodeF+''' and '''+@GlcodeT+''' '


insert into #TMPBUDGET4 
EXEC(@sql)


-------- Creating the TMP Table
select Top (1) * into #TMPBUDGET from #TMPBUDGET1 
Delete from [#TMPBUDGET] 

select Top (1) * into #TMP from #TMPBUDGET1 
Delete from [#TMP] 

-------- Dumping all the information in to TMP Table
insert into [#TMP] 
select * from [#TMPBUDGET1] 
union all
select * from [#TMPBUDGET2] 
union all
select * from [#TMPBUDGET3]
union all
select * from [#TMPBUDGET4]


if @ProjectF <> '%' and @ProjectT <> '%' and @NonProjectF = '%' and @NonProjectT = '%'
begin
  insert into [#TMPBUDGET]
  select '' U_BUCode , T0.U_PrjCode, T0.U_Account, T0.Cat  from [#TMP] T0 group by T0.U_PrjCode,T0.Cat, T0.U_Account
end
else
begin
  insert into [#TMPBUDGET]
  select T0.U_BUCode , T0.U_PrjCode, T0.U_Account, T0.Cat   from [#TMP] T0 
end

------ Fetching information for #TMPBUDGET
SELECT T0.[U_BUCode] [BU], T0.[U_PrjCode] [Project], T0.[U_Account] [GL Account] ,
T0.Cat  [Cat],
case when isnull(T0.U_PrjCode ,'') = '' then 
(SELECT sum( isnull(TT0.[U_BudAmount],0)) FROM PWCL ..[@AB_PROJECTBUDGET]  TT0 WHERE TT0.[U_BUCode] = T0.U_BUCode  and  TT0.[U_Account] = T0.U_Account and TT0.[U_BudName] = @ReforecastBudget )
else 0 end [Non ProjectLatest],
case when isnull(T0.U_PrjCode ,'') <> '' then
(SELECT sum(isnull( TT0.[U_BudAmount],0)) FROM PWCL ..[@AB_PROJECTBUDGET]  TT0 WHERE TT0.[U_PrjCode] = T0.U_PrjCode  and  TT0.[U_Account] = T0.U_Account and TT0.[U_BudName] = @ReforecastBudget)
else 0 end [ProjectLatest],
case when isnull(T0.U_PrjCode ,'') = '' then 
(SELECT sum(isnull(TT0.[U_BudAmount],0)) FROM PWCL ..[@AB_PROJECTBUDGET]  TT0 WHERE TT0.[U_BUCode] = T0.U_BUCode  and  TT0.[U_Account] = T0.U_Account and TT0.[U_BudName] = @MainBudget )
else (SELECT sum( isnull(TT0.[U_BudAmount],0)) FROM PWCL ..[@AB_PROJECTBUDGET]  TT0 WHERE TT0.[U_PrjCode] = T0.U_PrjCode  and  TT0.[U_Account] = T0.U_Account and TT0.[U_BudName] = @MainBudget) 
end [OriginalBudget],
case when isnull(T0.U_PrjCode ,'') = '' then 
(SELECT sum(isnull(TT0.[U_BudAmount],0)) FROM PWCL ..[@AB_PROJECTBUDGET]  TT0 WHERE TT0.[U_BUCode] = T0.U_BUCode  and  TT0.[U_Account] = T0.U_Account and TT0.[U_BudName] = @ReforecastBudget )
else (SELECT sum( isnull(TT0.[U_BudAmount],0)) FROM PWCL ..[@AB_PROJECTBUDGET]  TT0 WHERE TT0.[U_PrjCode] = T0.U_PrjCode  and  TT0.[U_Account] = T0.U_Account and TT0.[U_BudName] = @ReforecastBudget) 
end [RevisedBudget]
INTO #TMPBUDGETResult
FROM #TMPBUDGET  T0
--WHERE T0.U_PrjCode LIKE @Project AND T0.U_BUCode LIKE @NonProject 
--AND T0.U_Account LIKE @Glcode 

--------------------  Calling the function to calculate the Committed amount & Actual amount
 SELECT T0.BU , T0.Project , T1.PrjName , T0.[GL Account] , T2.AcctName  ,T0.[Non ProjectLatest] , T0.ProjectLatest , 
  [dbo].[AE_FN003_BUDGET_COMMITTEDAMOUNT] (@Year,T0.Cat ,T0.BU , T0.Project ,T0.[GL Account] )  [Committed],
  [dbo].[AE_FN004_BUDGET_ACTUALSPEND] (@Year,T0.Cat ,T0.BU , T0.Project ,T0.[GL Account] ) [Actual],
 --[dbo].[AE_FN005_BUDGET_PRAMOUNT] (@Year,T0.Cat ,T0.BU , T0.Project ,T0.[GL Account] ) 
 0 [PR], T0.Cat ,T0.OriginalBudget , T0.RevisedBudget 
  INTO #TMPFINAL
  FROM #TMPBUDGETResult T0
  left outer join OPRJ T1 on T0.Project = T1.PrjCode 
  left outer join OACT T2 on T0.[GL Account] = T2.AcctCode 
  ----------------------------------------------------------------------------------------------------------------------------------------------------------------------
   
  SELECT T0.BU , T0.Project , T0.PrjName , T0.[GL Account] , T0.AcctName, T0.[Non ProjectLatest] , T0.ProjectLatest , T0.Committed ,T0.Actual ,
  (isnull(T0.Committed,0) + isnull(T0.Actual,0) ) [CONSUMED],
  CASE WHEN T0.Cat = 'PRJ' THEN isnull(T0.ProjectLatest,0) - (isnull(T0.Committed,0) + isnull(T0.Actual,0)) ELSE
   isnull(T0.[Non ProjectLatest],0)  - (isnull(T0.Committed,0) + isnull(T0.Actual,0)) END  [ABALANCE] , T0.PR , T0.OriginalBudget , T0.RevisedBudget 
   INTO #BUDGET
   FROM #TMPFINAL T0

    SELECT max(T0.BU) [BU] , max(T0.Project) [Project], max(T0.PrjName) [PrjName] , max(T0.[GL Account]) [GL Account] , max(T0.AcctName) [AcctName], max(isnull(T0.[Non ProjectLatest],0)) [Non ProjectLatest] , 
	max(isnull(T0.ProjectLatest,0)) [ProjectLatest]
	,max( isnull(T0.Committed,0)) Committed , max(isnull(T0.Actual,0)) Actual ,
  max(isnull( T0.[CONSUMED],0)) [CONSUMED], max(isnull(T0.ABALANCE,0)) ABALANCE ,max( isnull(T0.PR,0) )[PR] , 
  ( isnull(T0.ABALANCE,0) - isnull(T0.PR,0) ) [APR] , max(isnull(T0.OriginalBudget,0)) [OriginalBudget]
   , max(isnull(T0.RevisedBudget ,0)) [RevisedBudget]
   FROM #BUDGET T0
   group by  T0.BU , T0.Project , T0.PrjName , T0.[GL Account] , T0.AcctName,T0.[Non ProjectLatest], T0.ProjectLatest
	, T0.Committed ,T0.Actual ,T0.[CONSUMED],T0.ABALANCE,T0.PR,T0.OriginalBudget, T0.RevisedBudget
	order by T0.BU, T0.Project, cast(T0.[GL Account] as numeric )
	

END
GO
/****** Object:  StoredProcedure [dbo].[AE_SP008_BUDGETREPORT_Details]    Script Date: 22/08/2017 11:41:16 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

--[AE_SP007_BUDGETREPORT_Test]'2016','%','%','Prj001','Prj001','71111100','91111100'
--[AE_SP005_BUDGET_COMMITTEDAMOUNT]'PWCL','2014','True','Prj002','71111100'
--[AE_SP008_BUDGETREPORT_Details_Test]'2016','%','%','Test5000','Test5000','71181300','71181300','Committed'


create PROCEDURE [dbo].[AE_SP008_BUDGETREPORT_Details]

@Year as varchar(10),
@NonProjectF as varchar(100),
@NonProjectT as varchar(100),
@ProjectF as varchar(100),
@ProjectT as varchar(100),
@GlcodeF as varchar(30),
@GlcodeT as varchar(30),
@cat as varchar(50)

AS

Declare @Condition as varchar(300)
Declare @Condition1 as varchar(300)
Declare @SQLCommitt as varchar(max)
Declare @SQLActual as varchar(max)

Declare @DateF varchar(30)
Declare @DateT varchar(30)


BEGIN


if @ProjectF = 'NA'
begin
 set @ProjectF = '%'
end

if @ProjectT = 'NA'
begin
 set @ProjectT = '%'
end

if @NonProjectF  = 'NA'
begin
 set @NonProjectF = '%'
end

if @NonProjectT  = 'NA'
begin
 set @NonProjectT = '%'
end

set @DateF = cast(cast(@year as numeric ) -1 AS VARCHAR) +'0701'
 set @DateT =  @year + '0630'

select top(1) t1.ItemCode  [doctype], t0.DocNum  [docnum], T0.DocEntry  [draftkey], T0.DocDate  [docdate], t1.ItemCode  [itemcode], t1.Dscription  [desc], T1.LineTotal  [ltotal], 
t1.U_AB_NONPROJECT  [nonproject], t1.Project  [project], t1.AcctCode  [gl]
 into #Tmp from DRF1 T1 join ODRF T0 ON T0.DocEntry = T1.DocEntry  where T1.DocEntry = 1 

 delete from #Tmp 

 --select * from #Tmp 
--------  Condition Filter
 
if @ProjectF = '%' and @ProjectT = '%' and  @NonProjectF = '%' and @NonProjectT = '%'
 begin
   set @condition = 'isnull(T1.Project,'''') LIKE ''%'''
   set @condition1 = 'isnull(T1.Project,'''') LIKE ''%'''
   set @condition += ' and isnull(T1.U_AB_NONPROJECT,'''') LIKE ''%'''
   set @condition1 += ' and isnull(T1.U_AB_NONPROJECT,'''') LIKE ''%'''
 end
else if @ProjectF <> '%' and @ProjectT <> '%' and @NonProjectF <> '%' and @NonProjectT <> '%'
 begin
  set @condition = 'isnull(T1.Project,'''') between '''+ @ProjectF +''' and '''+ @ProjectT +''''
  set @condition1 = 'isnull(T1.Project,'''') between '''+ @ProjectF +''' and '''+ @ProjectT +''''
  set @condition += ' and isnull(T1.U_AB_NONPROJECT,'''') between '''+ @NonProjectF  +''' and '''+ @NonProjectT  +''''
  set @condition1 += ' and isnull(T1.U_AB_NONPROJECT,'''') between '''+ @NonProjectF  +''' and '''+ @NonProjectT  +''' '
 end
 else if @ProjectF <> '%' and @ProjectT <> '%' and @NonProjectF = '%' and @NonProjectT = '%'
 begin
  set @condition = 'isnull(T1.Project,'''') between '''+ @ProjectF +''' and '''+ @ProjectT +''''
  set @condition1 = 'isnull(T1.Project,'''') between '''+ @ProjectF +''' and '''+ @ProjectT +''''
 end
 else
   begin
   set @condition = '  isnull(T1.U_AB_NONPROJECT,'''') between '''+ @NonProjectF  +''' and '''+ @NonProjectT  +''' and isnull(T1.Project,'''') = '''''
  set @condition1 = '  isnull(T1.U_AB_NONPROJECT,'''') between '''+ @NonProjectF  +''' and '''+ @NonProjectT  +''' and isnull(T1.Project,'''') = '''''
 end

 if @GlcodeF = '%' and @GlcodeT = '%'
 begin
   set @condition += ' and T1.[AcctCode] LIKE ''%'''
    set @condition1 += ' and T1.[Account] LIKE ''%'''
 end
else
 begin
  set @condition += ' and T1.[AcctCode] between '''+ @GlcodeF +''' and '''+ @GlcodeT +''''
   set @condition1 += ' and T1.[Account] between '''+ @GlcodeF +''' and '''+ @GlcodeT +''''
 end

-----   Fetching information for Committed Quantity

--- Add PO Draft with status Open  - Pending for Approval / Approved but not converted to PO
set @SQLCommitt =
'SELECT case when T0.ObjType = ''22'' then ''PO'' else ''PR'' end [doctype] , '''' [docnum], T0.DocEntry  [draftkey],T0.DocDate ,T1.ItemCode [itemcode], T1.Dscription [desc],
T1.LineTotal [ltotal], T1.U_AB_NONPROJECT [nonproject], T1.Project [project], T1.[AcctCode] [gl]
FROM ODRF T0  INNER JOIN DRF1 T1 ON T0.[DocEntry] = T1.[DocEntry] 
left outer join  OWDD T2 on T0.DocEntry = T2.DocEntry 
left outer join   OWTM T3 on T2.[WtmCode] = T3.[WtmCode]
where T1.LineStatus = ''O'' and  isnull(T2.[Status],'''') <> ''N'' and T0.DocStatus = ''O'' and T0.ObjType in (''22'',''1470000113'')
and T0.DocDate between '''+@datef+''' and '''+@DateT+''' and isnull(T3.[Active],'''') = ''Y''
and '+ @condition +' union all '

--- Add PR Document with status Open
set @SQLCommitt +=
'select  ''PR'' [doctype] , T0.DocNum [docnum], '''' [draftkey], T0.DocDate,T1.ItemCode [itemcode], T1.Dscription [desc],
T1.LineTotal [ltotal], T1.U_AB_NONPROJECT [nonproject], T1.Project [project], T1.[AcctCode] [gl] from OPRQ T0 join PRQ1 T1 on T0.DocEntry = T1.DocEntry 
left outer join DRF1 T2 on T2.BaseEntry = T1.DocEntry and T2.BaseLine = T1.LineNum and T2.BaseType = ''1470000113''
left outer join ODRF T3 on T2.DocEntry = T3.Docentry and T3.DocStatus = ''O''
where T1.LineStatus = ''O'' and T0.DocStatus = ''O''
and T0.DocDate between '''+@datef+''' and '''+@DateT+''' 
and '+ @condition +' union all '

--- Add PO Document with status Open
set @SQLCommitt +=
'select ''PO'' [doctype] , T0.DocNum [docnum], '''' [draftkey], T0.DocDate,T1.ItemCode [itemcode], T1.Dscription [desc],
T1.LineTotal [ltotal], T1.U_AB_NONPROJECT [nonproject], T1.Project [project], T1.[AcctCode] [gl] from OPOR T0 join POR1 T1 on T0.DocEntry = T1.DocEntry  
where T1.LineStatus = ''O'' and T0.DocStatus = ''O''
and T0.DocDate between '''+@datef+''' and '''+@DateT+''' and
'+ @condition +' union all '

--- Add PO Document with status Closed But GRN in Open Status
set @SQLCommitt +=
'SELECT ''PO'' [doctype] , T0.DocNum [docnum], '''' [draftkey],T0.DocDate, T1.ItemCode [itemcode], T1.Dscription [desc],
T1.LineTotal [ltotal], T1.U_AB_NONPROJECT [nonproject], T1.Project [project], T1.[AcctCode] [gl] FROM OPOR T0  INNER JOIN POR1 T1 ON T0.[DocEntry] = T1.[DocEntry] join 
 PDN1 T2 on T2.BaseEntry = T1.DocEntry and T2.BaseLine = T1.LineNum  INNER JOIN OPDN T3 ON T2.[DocEntry] = T3.[DocEntry] WHERE T2.[LineStatus] = ''O''
 and T0.DocDate between '''+@datef+''' and '''+@DateT+''' and
 '+ @condition +' '
 
 -----   Fetching information for Actual Quantity
  --- Add AP INVOICE with status Open - Base Document PO,GRPO & Direct
set @SQLActual =
'SELECT ''AP'' [doctype] , T0.DocNum [docnum], '''' [draftkey],T0.DocDate, T1.ItemCode [itemcode], T1.Dscription [desc],
T1.LineTotal [ltotal], T1.U_AB_NONPROJECT [nonproject], T1.Project [project], T1.[AcctCode] [gl] FROM OPCH T0  INNER JOIN PCH1 T1 ON T0.[DocEntry] = T1.[DocEntry]
WHERE t1.LineStatus = ''O'' and T0.DocStatus = ''O'' AND T1.BaseType IN (''20'',''22'',''-1'')
and T0.DocDate between '''+@datef+''' and '''+@DateT+''' and
'+ @condition +' union all '

--- Add JE & Direct
set @SQLActual +=
'SELECT ''JE'' [doctype] , T0.Number [docnum], '''' [draftkey], T0.RefDate,'''' [itemcode], '''' [desc],
(T1.[Debit] - T1.[Credit]) [ltotal], T1.U_AB_NONPROJECT [nonproject], T1.Project [project], T1.[Account] [gl]  FROM [OJDT]  T0 INNER JOIN JDT1 T1 
ON T0.[TransId] = T1.[TransId] WHERE T1.[TransType] = 30
and T0.[TransId] NOT IN (SELECT StornoToTr FROM ojdt WHERE ISNULL(StornoToTr,'''') <> '''') AND T0.[StornoToTr] IS NULL 
and T0.TaxDate between '''+@datef+''' and '''+@DateT+''' and
'+ @condition1 +' union all '


--- Less AP Credit memo standalone
set @SQLActual +=
'SELECT ''APCN'' [doctype] , T0.DocNum [docnum], '''' [draftkey], T0.DocDate,T1.ItemCode [itemcode], T1.Dscription [desc],
T1.LineTotal [ltotal], T1.U_AB_NONPROJECT [nonproject], T1.Project [project], T1.[AcctCode] [gl] FROM ORPC T0  INNER JOIN RPC1 T1 ON T0.[DocEntry] = T1.[DocEntry] WHERE T1.[LineStatus]  = ''O'' and [BaseType] = -1
and T0.DocDate between '''+@datef+''' and '''+@DateT+''' and
 '+ @condition +''

 if @cat = 'Actual Spend'
  begin
   insert into #tmp
    exec(@SQLActual)
  end
else
 begin
    insert into #tmp
    exec(@SQLCommitt)
 end

   select * from #Tmp t0 order by t0.nonproject  , T0.Project , T0.gl  

END
GO
/****** Object:  StoredProcedure [dbo].[BP_Audit_Report]    Script Date: 22/08/2017 11:41:16 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

--[BP_Audit_Report] 'S1000016','S1000016','20170726','20170726'

CREATE PROCEDURE [dbo].[BP_Audit_Report]
		@FromCardcode Nvarchar(40) ,
		@ToCardcode Nvarchar(40),
		@FromDate Datetime,
		@ToDate Datetime


AS
BEGIN
--DECLARE @FromCardcode Nvarchar(40) ,
--		@ToCardcode Nvarchar(40),
--		@FromDate Datetime,
--		@ToDate Datetime

--		SET @FromCardcode = 'T4000020'
--		SET @ToCardcode = 'T4000020'
--		SET @FromDate = '20170628'
--		SET @ToDate = '20170628'

	-- SET NOCOUNT ON added to prevent extra result sets from
		-- interfering with SELECT statements.
	SET NOCOUNT ON;


if exists(select * from dbo.sysobjects t1 where t1.xtype = 'U' and t1.id = OBJECT_ID(N'#Final'))
begin
drop table #Final
end

if exists(select * from dbo.sysobjects t1 where t1.xtype = 'U' and t1.id = OBJECT_ID(N'TempACRD'))
begin
drop table TempACRD
end

if exists(select * from dbo.sysobjects t1 where t1.xtype = 'U' and t1.id = OBJECT_ID(N'TempOCRDPerson'))
begin
drop table TempOCRDPerson
end
if exists(select * from dbo.sysobjects t1 where t1.xtype = 'U' and t1.id = OBJECT_ID(N'#OCRDTemp'))
begin
drop table #OCRDTemp
end
if exists(select * from dbo.sysobjects t1 where t1.xtype = 'U' and t1.id = OBJECT_ID(N'#CRD1Temp'))
begin
drop table #CRD1Temp
end
if exists(select * from dbo.sysobjects t1 where t1.xtype = 'U' and t1.id = OBJECT_ID(N'TempOCRD'))
begin
drop table TempOCRD
end
if exists(select * from dbo.sysobjects t1 where t1.xtype = 'U' and t1.id = OBJECT_ID(N'TempCRD1'))
begin
drop table TempCRD1
end
if exists(select * from dbo.sysobjects t1 where t1.xtype = 'U' and t1.id = OBJECT_ID(N'TempACR'))
begin
drop table TempACR
end
if exists(select * from dbo.sysobjects t1 where t1.xtype = 'U' and t1.id = OBJECT_ID(N'TempOCRDF'))
begin
drop table TempOCRDF
end
if exists(select * from dbo.sysobjects t1 where t1.xtype = 'U' and t1.id = OBJECT_ID(N'#ACR2Temp'))
begin
drop table #ACR
end
if exists(select * from dbo.sysobjects t1 where t1.xtype = 'U' and t1.id = OBJECT_ID(N'#ACPRTemp'))
begin
drop table #ACPRTemp
end
if exists(select * from dbo.sysobjects t1 where t1.xtype = 'U' and t1.id = OBJECT_ID(N'TempARPC'))
begin
drop table TempARPC
end


  Select Row_Number() over (order by Col .name asc) as id,   isnull(des.Descr,Col .name) ColDesc, Col .name into #OCRDTemp
from sys .all_columns Col Inner join sys .tables tbl on Col .object_id = tbl .object_id 
left outer join CUFD des on des.TableID = tbl.name and   Col .name =  'U_'+ des.AliasID  
where tbl .name = 'OCRD' and Col .name Not in ('UpdateDate','CreateDate','LogInstanc','UserSign2','UserSign','Address','ZipCode','AddrType','Block','Building','City','Country',
'County','BillToDef','ZipCode','State1','Address','StreetNo','MailAddrTy','MailBlock','MailCity','MailCountr','MailCounty','ShipToDef','MailZipCod','State2','MailAddres',
'MailStrNo','MailBuildi','Gender'
,'AccCritria','AddID','Affiliate','AtcEntry','AutoCalBCG','AutoPost','AvrageLate','BackOrder','Balance','BalanceFC','BalanceSys','BalTrnsfrd'
,'BCACode','BlockDunn','BNKCounter','BoEDiscnt','BoEOnClct','BoEPrsnt','Box1099','Business','CardValid','CDPNum','CertBKeep','CertWHT','chainStore',
'ChecksBal','CollecAuth','CommGrCode','Commission','ConCerti','ConnBP','CpnNo','CrtfcateNO','DataSource','DatevFirst','DdctFileNo','DdctOffice',
'DdctPrcnt','DdgKey','DdtKey','DebPayAcct','Deleted','DocEntry','DpmClear','DpmIntAct','DscntObjct','DscntRel','DunnDate','DunnLevel','ITWTCode',
'KBKCode','LangCode','LetterNum','ListNum','LocMth','MainUsage','MivzExpSts','MltMthNum','MTHCounter','NINum','NTSWebSite','Number','ObjType','OKATO'
,'OKTMO','OpCode347','OprCount','OrderBalFC','OrderBalSy','OrdersBal','OtrCtlAcct','Pager','PartDelivr','PlngGroup','TaxIdIdent','TaxRndRule','ThreshOver'
,'TolrncDays','TpCusPres','TypeOfOp','TypWTReprt','SefazReply','SefazDate','SefazCheck','SCAdjust','RoleTypCod','RelCode','RcpntID','','''')

  Select Row_Number() over (order by Col .name asc) as id,   isnull(des.Descr,Col .name) ColDesc, Col .name into #CRD1Temp
from sys .all_columns Col Inner join sys .tables tbl on Col .object_id = tbl .object_id 
left outer join CUFD des on des.TableID = tbl.name and   Col .name =  'U_'+ des.AliasID  
where tbl .name = 'CRD1' and Col .name Not in ('CreateDate','UpdateDate','LogInstanc','UserSign2','UserSign','ObjType','LineNum','AdresType','CardCode','LicTradNum')

Select Row_Number() over (order by Col .name asc) as id,   isnull(des.Descr,Col .name) ColDesc, Col .name into #ACR2Temp
from sys .all_columns Col Inner join sys .tables tbl on Col .object_id = tbl .object_id 
left outer join CUFD des on des.TableID = tbl.name and   Col .name =  'U_'+ des.AliasID  
where tbl .name = 'ACR2' and Col .name Not in ('CreateDate','UpdateDate','LogInstanc','UserSign2','UserSign')

Select Row_Number() over (order by Col .name asc) as id,   isnull(des.Descr,Col .name) ColDesc, Col .name into #ACPRTemp
from sys .all_columns Col Inner join sys .tables tbl on Col .object_id = tbl .object_id 
left outer join CUFD des on des.TableID = tbl.name and   Col .name =  'U_'+ des.AliasID  
where tbl .name = 'ACPR' and Col .name Not in ('CreateDate','UpdateDate','LogInstanc','UserSign2','UserSign','updateTime','Active','DataSource',
'NFeRcpn','ObjType','Profession','CardCode','CntctCode','e_Payment','AcctName','Gender')

select ACRD .CardCode ,MAX (ACRD.LogInstanc) as LogInstanc into TempACRD
from ACRD where ACRD .CardCode >= @FromCardcode and ACRD .CardCode <= @ToCardcode  group by ACRD .CardCode

select ACRD .CardCode ,ACRD.LogInstanc as LogInstanc into TempACR
from ACRD where ACRD .CardCode >= @FromCardcode and ACRD .CardCode <= @ToCardcode  and ACRD.UpdateDate between @FromDate and @ToDate  
--and loginstanc = (select MAX(T0.LogInstanc) from ACRD  T0 where ACRD.CardCode = T0.CardCode )

select ACR1 .CardCode ,ACR1 .LineNum ,ACR1.LogInstanc as LogInstanc into TempCRD1
from ACR1 where ACR1 .CardCode >= @FromCardcode and ACR1 .CardCode <= @ToCardcode --group by ACR1 .CardCode ,ACR1 .LineNum 
and loginstanc in (select T0.LogInstanc from TempACR  T0 where ACR1.CardCode = T0.CardCode )

Select OCRD .LogInstanc+1 as LogInstanc,OCRD .CreateDate  as [Date],OCRD .CardCode as [BP Code],OCRD .CardName as [BP Name],
Case when CardType = 'C' then 'Customer' when CardType = 'S' then 'Vendor' else 'Lead' end [BP Type],
OUSR .U_NAME as [Created By],Convert(Nvarchar(MAX),null) as [Field Name],Convert (Nvarchar(MAX),'***Created***') as [Old Value],Convert (Nvarchar(MAX),null) as [New Value] into TempOCRD
from OCRD inner join OUSR on OCRD .UserSign = OUSR .INTERNAL_K 
where OCRD .CardCode >= @FromCardcode and  OCRD .CardCode <= @ToCardcode and 
Convert(Nvarchar(8), OCRD .CreateDate ,112) >= Convert(Nvarchar(8),  @FromDate ,112) and Convert(Nvarchar(8), OCRD .UpdateDate  ,112) <= Convert(Nvarchar(8),  @ToDate ,112);


Select OCRD .LogInstanc+1 as LogInstanc,OCRD .CreateDate  as [Date],OCRD .CardCode as [BP Code],OCRD .CardName as [BP Name],
Case when CardType = 'C' then 'Customer' when CardType = 'S' then 'Vendor' else 'Lead' end [BP Type],
OUSR .U_NAME as [Created By],Convert(Nvarchar(MAX),null) as [Field Name],Convert (Nvarchar(MAX),'***Created***') as [Old Value],Convert (Nvarchar(MAX),null) as [New Value] into TempOCRDPerson
from OCRD inner join OUSR on OCRD .UserSign = OUSR .INTERNAL_K 
where OCRD .CardCode >= @FromCardcode and  OCRD .CardCode <= @ToCardcode and 
Convert(Nvarchar(8), OCRD .CreateDate ,112) >= Convert(Nvarchar(8),  @FromDate ,112) and Convert(Nvarchar(8), OCRD .UpdateDate  ,112) <= Convert(Nvarchar(8),  @ToDate ,112);


Select OCRD .LogInstanc+1 as Seq ,OCRD .LogInstanc+1 as LogInstanc,OCRD .CreateDate  as [Date],OCRD .CardCode as [BP Code],OCRD .CardName as [BP Name],
Case when CardType = 'C' then 'Customer' when CardType = 'S' then 'Vendor' else 'Lead' end [BP Type],
OUSR .U_NAME as [Created By],Convert(Nvarchar(MAX),null) as [Field Name],Convert (Nvarchar(MAX),'***Created***') as [Old Value],Convert (Nvarchar(MAX),null) as [New Value] into TempARPC
from OCRD inner join OUSR on OCRD .UserSign = OUSR .INTERNAL_K 
where OCRD .CardCode >= @FromCardcode and  OCRD .CardCode <= @ToCardcode and 
Convert(Nvarchar(8), OCRD .CreateDate ,112) >= Convert(Nvarchar(8),  @FromDate ,112) and Convert(Nvarchar(8), OCRD .UpdateDate ,112) <= Convert(Nvarchar(8),  @ToDate ,112);

Declare @Count int,@Counter int,@ColName Nvarchar (100), @Query Nvarchar(max), @ColDesc Nvarchar (100), @Tmpcal NVarchar(10), @ColumnOLD NVarchar(200), @ColumnNEW NVarchar(200) ;

--select * from TempCrd1


Set @Count = (Select MAX (#OCRDTemp .id) from #OCRDTemp);
Set @Counter = 1;

While (@Counter <= @Count)

Begin

	Set @Query = '';

	Set @ColName = (Select top 1 Name from #OCRDTemp  where id = @Counter);
	Set @ColDesc = (Select top 1 ColDesc from #OCRDTemp  where id = @Counter)

	Set @Query = 'Insert into TempOCRD Select T0 .LogInstanc+1,T1 .UpdateDate  as [Date],T0 .CardCode as [BP Code],T0 .CardName as [BP Name],
	Case when T0 .CardType = ''C'' then ''Customer'' when T0 .CardType = ''S'' then ''Vendor'' else ''Lead'' end [BP Type],
	OUSR .U_NAME as [Created By],'''+ @ColDesc +''' as [Field Name],convert(Nvarchar(MAX),T0 .'+@ColName+',106) as [Old Value],Convert(Nvarchar(MAX),T1 .'+@ColName+',106) as [New Value]
	from ACRD T0 inner join ACRD T1 on T0 .CardCode = T1 .CardCode and T0 .LogInstanc = T1 .LogInstanc -1
	left outer join OUSR on T1 .UserSign2 = OUSR .INTERNAL_K 
	where  T0 .CardCode >= '''+@FromCardcode+''' and  T0 .CardCode <= '''+@ToCardcode+''' and   
	isnull(convert(Nvarchar(MAX),T0 .'+@ColName+',106) ,'''') <> isnull(convert(Nvarchar(MAX),T1 .'+@ColName+',106),'''')';
	
	--print @query
   Exec(@Query)

Set @Query = 'Insert into TempOCRD Select T1 .LogInstanc+1,T0 .UpdateDate  as [Date],T0 .CardCode as [BP Code],T0 .CardName as [BP Name],
Case when T0 .CardType = ''C'' then ''Customer'' when T0 .CardType = ''S'' then ''Vendor'' else ''Lead'' end [BP Type],
OUSR .U_NAME as [Created By],'''+ @ColDesc +''' as [Field Name],convert(Nvarchar(MAX),T1 .'+@ColName+',106) as [Old Value],
Convert(Nvarchar(MAX),T0 .'+@ColName+',106) as [New Value]
from OCRD T0 inner join ACRD T1 on T0 .CardCode = T1 .CardCode 
inner join TempACRD T2 on T0.Cardcode = T2 .Cardcode and T1 .LogInstanc = T2.LogInstanc
left outer join OUSR on T0 .UserSign2 = OUSR .INTERNAL_K 
where  T0 .CardCode >= '''+@FromCardcode+''' and  T0 .CardCode <= '''+@ToCardcode+''' and  
isnull(convert(Nvarchar(MAX),T0 .'+@ColName+',106) ,'''') <> isnull(convert(Nvarchar(MAX),T1 .'+@ColName+',106),'''')';
--print @query
  Exec(@Query)

Set @Counter = @Counter + 1;
End 

--// Bank Account Name

Insert into TempOCRD
	Select T1 .LogInstanc+1,TT .UpdateDate  as [Date],T0 .CardCode as [BP Code],T0 .CardName as [BP Name],
Case when T0 .CardType = 'C' then 'Customer' when T0 .CardType = 'S' then 'Vendor' else 'Lead' end [BP Type],
OUSR .U_NAME as [Created By],'Bank Account Name' as [Field Name],convert(Nvarchar(MAX),T1.AcctName    ,106) as [Old Value],
Convert(Nvarchar(MAX),T2.AcctName    ,106) as [New Value]
from ACRD T0 inner join ACRD TT on T0 .CardCode = TT .CardCode and T0 .LogInstanc = TT .LogInstanc -1
inner join ACRB T1 on T0 .CardCode = T1 .CardCode and T0 .LogInstanc = T1 .LogInstanc and T0.BankCode = T1.BankCode 
join ACRB T2 ON T0 .CardCode = T2 .CardCode and T0 .LogInstanc = T2 .LogInstanc -1 and T0.BankCode = T2.BankCode 
	left outer join OUSR on TT .UserSign2 = OUSR .INTERNAL_K 
where  T0 .CardCode >= @FromCardcode and  T0 .CardCode <= @ToCardcode and  
isnull(convert(Nvarchar(MAX),T2.AcctName  ,106) ,'''') <> isnull(convert(Nvarchar(MAX),T1.AcctName,106),'''')
	
	

--// first time changes
Insert into TempOCRD
Select T1 .LogInstanc+1,T0 .UpdateDate  as [Date],T0 .CardCode as [BP Code],T0 .CardName as [BP Name],
Case when T0 .CardType = 'C' then 'Customer' when T0 .CardType = 'S' then 'Vendor' else 'Lead' end [BP Type],
OUSR .U_NAME as [Created By],'Bank Account Name' as [Field Name],convert(Nvarchar(MAX),T2.AcctName  ,106) as [Old Value],
Convert(Nvarchar(MAX),Tt.AcctName  ,106) as [New Value]
from  OCRB TT JOIN  OCRD T0 ON T0.CardCode = TT.CardCode AND TT.BankCode = T0.BankCode  inner join TempACRD T1 on T0 .CardCode = T1 .CardCode 
inner join ACRB T2 on T0.Cardcode = T2 .Cardcode and T1 .LogInstanc = T2.LogInstanc
left outer join OUSR on T0 .UserSign2 = OUSR .INTERNAL_K 
where  T0 .CardCode >= @FromCardcode and  T0 .CardCode <= @ToCardcode and  
isnull(convert(Nvarchar(MAX),Tt.AcctName ,106) ,'''') <> isnull(convert(Nvarchar(MAX),T2.AcctName,106),'''')



--//Change log for Address
Set @Count = (Select MAX (#CRD1Temp .id) from #CRD1Temp);
Set @Counter = 1;

While (@Counter <= @Count)

Begin
	Set @Query = '';
	Set @ColName = (Select top 1 Name from #CRD1Temp  where id = @Counter);
	Set @ColDesc = (Select top 1 ColDesc from #CRD1Temp  where id = @Counter)

	Set @Query = 'Insert into TempOCRD Select distinct ACRD .LogInstanc,ACRD .UpdateDate  as [Date],ACRD .CardCode as [BP Code],ACRD .CardName as [BP Name],
Case when ACRD .CardType = ''C'' then ''Customer'' when ACRD .CardType = ''S'' then ''Vendor'' else ''Lead'' end [BP Type],
OUSR .U_NAME as [Created By],T1.Address+''-''+'''+ @ColDesc +''' as [Field Name],convert(Nvarchar(MAX),T0 .'+@ColName+',106) as [Old Value]
,Convert(Nvarchar(MAX),T3 .'+@ColName+',106) as [New Value]
from ACRD left outer join OUSR on ACRD .UserSign2 = OUSR .INTERNAL_K 
inner join ACR1 T0 on ACRD .Cardcode = T0 . Cardcode 
inner join ACR1 T1 on T0 .CardCode = T1 .CardCode and T0.Linenum = T1.Linenum and T0 .LogInstanc = T1 .LogInstanc 
 and ACRD .LogInstanc = (select MAX(loginstanc) from ACR1 where Address = T1.Address)
inner join CRD1 T3 on T3.CardCode = T1.CardCode and T1.LineNum = T3.LineNum 
where  T0 .CardCode >= '''+@FromCardcode+''' and  T0 .CardCode <= '''+@ToCardcode+'''  
--and isnull(convert(Nvarchar(MAX),T0 .'+@ColName+',106) ,'''') <> isnull(convert(Nvarchar(MAX),T1 .'+@ColName+',106),'''')
union
select distinct vw.LogInstanc,vw.Date,vw.[BP Code],vw.[BP Name],vw.[BP Type],vw.[Created By],vw.[Field Name],vw.[Old Value],vw.[New value] from (
Select T2 .LogInstanc,OCRD .UpdateDate  as [Date],OCRD .CardCode as [BP Code],OCRD .CardName as [BP Name],
Case when OCRD .CardType = ''C'' then ''Customer'' when OCRD .CardType = ''S'' then ''Vendor'' else ''Lead'' end [BP Type],
OUSR .U_NAME as [Created By],T2.Address +''-''+'''+ @ColDesc +''' as [Field Name],
'''' [Old Value]
,case  when isnull(convert(Nvarchar(MAX),T2 .'+@ColName+',106) ,'''') <> isnull(convert(Nvarchar(MAX),T1 .'+@ColName+',106),'''')
then Convert(Nvarchar(MAX),T2 .'+@ColName+',106) else '''' end  [New value]
--,isnull(Convert(Nvarchar(MAX),T1 .'+@ColName+',106),'''')[Filter]
,(select  count(Convert(Nvarchar(MAX),'+@ColName+',106)) from ACR1 
where Convert(Nvarchar(MAX),'+@ColName+',106) = Convert(Nvarchar(MAX),T2 .'+@ColName+',106) and cardcode = T2.cardcode  and Address = T2.Address) [Filter]
from OCRD left outer join OUSR on OCRD .UserSign2 = OUSR .INTERNAL_K 
left join  CRD1 T2 on OCRD .Cardcode = T2 . Cardcode 
left join ACR1 T0 on OCRD .Cardcode = T0 . Cardcode 
left join ACR1 T1 on T0 .CardCode = T1 .CardCode and T0.Linenum = T1.Linenum and T0 .LogInstanc = T1 .LogInstanc -1
 --and ACRD .LogInstanc = T1.LogInstanc
where  T2 .CardCode >= '''+@FromCardcode+''' and  T2 .CardCode <= '''+@ToCardcode+''' 
--AND convert(Nvarchar(MAX),T1 .'+@ColName+',106) not in (select convert(Nvarchar(MAX),'+@ColName+',106) 
--from CRD1 where CardCode >= '''+@FromCardcode+''' and CardCode <= '''+@ToCardcode+''')


) as vw where vw.[New value] <> '''' and vw.Filter = 0';
--print @query
 Exec(@Query)

	Set @Query = 'Insert into TempOCRD Select t1 .LogInstanc+1 ,OCRD .UpdateDate  as [Date],OCRD .CardCode as [BP Code],OCRD .CardName as [BP Name],
Case when OCRD .CardType = ''C'' then ''Customer'' when OCRD .CardType = ''S'' then ''Vendor'' else ''Lead'' end [BP Type],
OUSR .U_NAME as [Created By],T1.Address+''-''+'''+ @ColDesc +''' as [Field Name],convert(Nvarchar(MAX),T3 .'+@ColName+',106) as [Old Value],Convert(Nvarchar(MAX),T1 .'+@ColName+',106) as [New Value]
from OCRD Left outer join OUSR on OCRD .UserSign2 = OUSR .INTERNAL_K 
inner join TempCrd1 T2 on OCRD .Cardcode = T2 .Cardcode
inner join ACR1 T1 on T2 .CardCode = T1 .CardCode and T2.Linenum = T1.Linenum 
 and T1 .LogInstanc = T2 .LogInstanc 
 left outer join ACR1 T3 on T2 .CardCode = T3 .CardCode and T3.Linenum = T1.Linenum 
 and T3 .LogInstanc = T2 .LogInstanc -1
where T2 .CardCode >= '''+@FromCardcode+''' and  T2 .CardCode <= '''+@ToCardcode+''' 
and  
isnull(convert(Nvarchar(MAX),T3 .'+@ColName+',106) ,'''') <> isnull(convert(Nvarchar(MAX),T1 .'+@ColName+',106),'''')
';
 --print @query
 Exec(@Query)

Set @Counter = @Counter + 1;

End 

Set @Count = (Select max (TempACR.LogInstanc ) from TempACR);
Set @Counter = (Select min (TempACR.LogInstanc ) from TempACR);

While (@Counter <= @Count)

Begin
	Set @Query = ''
	set @Tmpcal = @Counter
	
set @Query = 'Insert into TempOCRD Select ACRD .LogInstanc,ACRD .UpdateDate  as [Date],ACRD .CardCode as [BP Code],ACRD .CardName as [BP Name],
Case when ACRD .CardType = ''C'' then ''Customer'' when ACRD .CardType = ''S'' then ''Vendor'' else ''Lead'' end [BP Type],
OUSR .U_NAME as [Created By],''PymCode'' as [Field Name],
--(select PymCode from ACR2 TT where TT.CardCode = ACRD.CardCode and  TT.LogInstanc = ACRD .LogInstanc -1 and TT.LineNum = 2) as [Old Value], 
--(select PymCode from ACR2 TT where TT.CardCode = ACRD.CardCode and  TT.LogInstanc = ACRD.LogInstanc and TT.LineNum = 2) as [New Value]

case when (select top 1 isnull(PymCode,'''') from ACR2 TT where TT.CardCode = ACRD.CardCode and  TT.LogInstanc = ACRD.LogInstanc and TT.LineNum = 2) <> '''' then 
(select top 1 PymCode from ACR2 TT where TT.CardCode = ACRD.CardCode and  TT.LogInstanc = ACRD.LogInstanc and TT.LineNum = 2) 
else (select top 1 PymCode from ACR2 TT 
where TT.CardCode = ACRD.CardCode and  TT.LogInstanc = ACRD.LogInstanc-1 and TT.LineNum = 2) end
as [Old Value], 
(select top 1 PymCode from CRD2 TT where TT.CardCode = ACRD.CardCode and TT.LineNum = 2) as [New Value]

from ACRD left outer join OUSR on ACRD .UserSign2 = OUSR .INTERNAL_K 
where ACRD .CardCode >= '''+@FromCardcode+''' and  ACRD .CardCode <= '''+@ToCardcode+''' and ACRD .LogInstanc = '''+@Tmpcal+''''
 exec(@Query)
--print  @Query

set @Query = 'Insert into TempOCRD Select ACRD .LogInstanc,ACRD .UpdateDate  as [Date],ACRD .CardCode as [BP Code],ACRD .CardName as [BP Name],
Case when ACRD .CardType = ''C'' then ''Customer'' when ACRD .CardType = ''S'' then ''Vendor'' else ''Lead'' end [BP Type],
OUSR .U_NAME as [Created By],''PymCode'' as [Field Name],
--(select PymCode from ACR2 TT where TT.CardCode = ACRD.CardCode and  TT.LogInstanc = ACRD .LogInstanc -1 and TT.LineNum = 3) as [Old Value], 
--(select PymCode from ACR2 TT where TT.CardCode = ACRD.CardCode and  TT.LogInstanc = ACRD.LogInstanc and TT.LineNum = 3) as [New Value]

case when (select top 1 isnull(PymCode,'''') from ACR2 TT where TT.CardCode = ACRD.CardCode and  TT.LogInstanc = ACRD.LogInstanc and TT.LineNum = 3) <> '''' then 
(select top 1 PymCode from ACR2 TT where TT.CardCode = ACRD.CardCode and  TT.LogInstanc = ACRD.LogInstanc and TT.LineNum = 3) 
else (select top 1 PymCode from ACR2 TT 
where TT.CardCode = ACRD.CardCode and  TT.LogInstanc = ACRD.LogInstanc-1 and TT.LineNum = 3) end
as [Old Value], 
(select top 1 PymCode from CRD2 TT where TT.CardCode = ACRD.CardCode and TT.LineNum = 3) as [New Value]

from ACRD left outer join OUSR on ACRD .UserSign2 = OUSR .INTERNAL_K 
where ACRD .CardCode >= '''+@FromCardcode+''' and  ACRD .CardCode <= '''+@ToCardcode+'''  and ACRD .LogInstanc = '''+@Tmpcal+''''

 exec(@Query)
--print  @Query
Set @Counter = @Counter + 1;
End 


Set @Count = (Select MAX (#ACPRTemp .id) from #ACPRTemp);
Set @Counter = 1;

While (@Counter <= @Count)

Begin
	Set @Query = '';
	Set @ColName = (Select Name from #ACPRTemp  where id = @Counter);
	Set @ColDesc = (Select ColDesc from #ACPRTemp  where id = @Counter)

SET @Query = '; WITH T1 AS(SELECT DISTINCT T1 .LogInstanc,T0 .UpdateDate  as [Date],T0 .CardCode as [BP Code],T3 .CardName as [BP Name],
Case when T3 .CardType = ''C'' then ''Customer'' when T3 .CardType = ''S'' then ''Vendor'' else ''Lead'' end [BP Type],
T4 .U_NAME as [Created By],'''+ @ColDesc +''' as [Field Name],
-- case when 
--(select max(LogInstanc) from ACPR TT where TT.CardCode = T0.CardCode  and TT.CntctCode = T1.CntctCode 
--and TT.updateDate = T1.updateDate) = T1.LogInstanc then Convert(Nvarchar(MAX),T1 .'+@ColName+',106)
-- else ''''
-- end 
--  AS [Old Value]
(select TT. '+@ColName+' from ACPR TT where  TT.CntctCode = T1.CntctCode 
and TT.updateDate = T1.updateDate and LogInstanc = (select MAX(logInstanc)-1 from ACPR where ACPR.CntctCode = T1.CntctCode 
and ACPR.updateDate = T1.updateDate)) [Old Value],
  convert(Nvarchar(MAX),T0 .'+@ColName+',106) AS [New Value]
  FROM OCPR T0
INNER JOIN ACPR T5 ON T5.CardCode = T0.CardCode
LEFT JOIN ACPR T1 ON T1.CntctCode = T0.CntctCode
LEFT JOIN ACPR T2 ON T2.CntctCode = T1.CntctCode and T1.LogInstanc = T2.LogInstanc-1
LEFT JOIN ACRD T3 ON T3.CardCode = T1.CardCode AND T3.UpdateDate = T1.updateDate
LEFT JOIN OUSR T4 on T3 .UserSign2 = T4 .INTERNAL_K
WHERE  
T1.CntctCode = T0.CntctCode AND
--T1. '+@ColName+' <> T0. '+@ColName+' AND 
T0 .CardCode >= '''+@FromCardcode+''' and  T0 .CardCode <= '''+@ToCardcode+''' 
-- and
--Convert(Nvarchar(8), T0 .UpdateDate ,112)  >= '''+Convert(Nvarchar(8),  @FromDate ,112)+''' AND 
--Convert(Nvarchar(8), T0 .UpdateDate ,112)  <= '''+Convert(Nvarchar(8),  @ToDate ,112)+'''  
--AND T0.updateDate between ''2016-09-20'' and ''2016-09-20'' 
--AND isnull(convert(Nvarchar(MAX),T0 .'+@ColName+',106) ,'''') <> isnull(convert(Nvarchar(MAX),T1 .'+@ColName+',106),'''')
--AND MAX(T3.LogInstanc) = MAX(T3.LogInstanc)
AND T1.LogInstanc = (select max(LogInstanc) from ACPR TT where TT.CardCode = T0.CardCode  and TT.CntctCode = T1.CntctCode 
and TT.updateDate = T0.updateDate
) 
group by T1.CntctCode, T0.CardCode,T3.LogInstanc,T1.LogInstanc,T0.updateDate,
T1.updateDate,T3 .CardName,T3 .CardType,T4 .U_NAME , T1. '+@ColName+', T0. '+@ColName+', T2.'+@ColName+'

union 
Select vw.LogInstanc,vw.updateDate,vw.CardCode,vw.CardName,vw.[BP Type],vw.[Created By],vw.[Field Name],vw.[OldValue],vw.[New Value] from (
Select Distinct T0.[LogInstanc],T0.updateDate,T3.CardCode,T3.CardName, 
Case when T3 .CardType = ''C'' then ''Customer'' when T3 .CardType = ''S'' then ''Vendor'' else ''Lead'' end [BP Type],
T4 .U_NAME as [Created By],'''+ @ColDesc +''' as [Field Name],
isNull(Convert(Nvarchar(MAX),T0 .'+@ColName+',106),'''')  [New Value],isNull(Convert(Nvarchar(MAX),T1 .'+@ColName+',106),'''') [FilterName],'''' [OldValue] FROM  OCPR T0
LEFT JOIN ACPR T1 ON T1.CntctCode = T0.CntctCode
LEFT JOIN ACPR T2 ON T2.CntctCode = T1.CntctCode and T1.LogInstanc = T2.LogInstanc-1
LEFT JOIN OCRD T3 ON T3.CardCode = T0.CardCode 
LEFT JOIN OUSR T4 on T3 .UserSign2 = T4 .INTERNAL_K
Where T0 .CardCode >= '''+@FromCardcode+''' and  T0 .CardCode <= '''+@ToCardcode+'''   
--and Convert(Nvarchar(8), T0 .UpdateDate ,112)  >= '''+Convert(Nvarchar(8),  @FromDate ,112)+''' AND 
--Convert(Nvarchar(8), T0 .UpdateDate ,112)  <= '''+Convert(Nvarchar(8),  @ToDate ,112)+''' 
--AND T0.updateDate  between ''2016-09-20'' and ''2016-09-20''  

) as vw where vw.FilterName = '''' and vw.CardCode <> ''''
)
Insert into TempOCRD
SELECT * FROM T1 WHERE LogInstanc IN (SELECT LogInstanc FROM T1)'

--print @query

 Exec(@Query)



Set @Counter = @Counter + 1;
End 
--Insert  into TempOCRD select * from TempOCRDPerson

Delete from TempOCRD where [Old Value] = '-1' and [New Value] = '-1'
--Delete from TempOCRD where ISNULL([Old Value],'') = ISNULL([New Value],'')

Select T0.LogInstanc , T0.Date , T0.[BP Code], T0.[BP Name] , T0.[BP Type] ,T0.[Created By] ,
(case T0.[Field Name]
when 'Notes' then 'Remarks' 
when 'Free_Text' then 'Remarks Tab - Remarks'
when 'frozenFor' then 'In Active'
when 'validFor' then 'Active'
when 'VatIdUnCmp' then 'Business Registration Number'
when 'ECVatGroup' then 'GST Code'
when 'Bill to-Block' then 'Block'
when 'Bill to-Building' then 'Building/Floor/Room'
when 'Bill to-City' then 'City'
when 'Bill to-Country' then 'Country'
when 'Bill to-County' then 'County'
when 'Bill to-Street' then 'Street / PO Box'
when 'Bill to-ZipCode' then 'Postcode'
when 'Bill to:-Block' then 'Block'
when 'Bill to:-Building' then 'Building/Floor/Room'
when 'Bill to:-City' then 'City'
when 'Bill to:-Country' then 'Country'
when 'Bill to:-County' then 'County'
when 'Bill to:-Street' then 'Street / PO Box'
when 'Bill to:-ZipCode' then 'Postcode'
when 'HousBnkAct' then 'Account'
when 'IntrntSite' then 'Web Site'
when 'LicTradNum' then 'GST Number'
when 'MandateID' then 'Sort code'
when 'Phone1' then 'Tel. 1'
when 'Phone2' then 'Tel. 2'
when 'PymCode' then 'e_Payment'
when 'BankCode' then 'Bank Code'
when 'BankCountr' then 'Bank Country' 
when 'CardName' then 'Name'
when 'Cellolar' then 'Mobile Phone'
when 'Cellular' then 'Mobile Phone'
when 'CntctPrsn' then 'Contact ID'
when 'DflAccount' then 'Account'
when 'DflBranch' then 'Branch'
when 'DflIBAN' then 'IBAN/ABA'
when 'DflSwift' then 'BIC/SWIFT Code'
when 'DflBankKey' then 'Default Bank Key'
when 'BankCtlKey' then 'Ctrl Int. ID'
when 'GroupNum' then 'Payment Terms Code'
when 'SlpCode' then 'Buyer'
when 'IndustryC' then 'Industry'
when 'FrozenComm' then 'Remarks'
when 'E_MailL' then 'E-Mail'
when 'e_Payment' then 'Payment Methods'
else T0.[Field Name]
end) [Field Name]
 , 
 (case 
   when T0.[Field Name] = 'SlpCode' then (SELECT top 1 TT0.[SlpName] FROM OSLP TT0 where TT0.[SlpCode] = T0.[Old Value]) 
   when T0.[Field Name] = 'GroupNum' then (SELECT top 1   TT0.[PymntGroup] FROM OCTG TT0 where TT0.[GroupNum] = T0.[Old Value])
   else iSNULL(T0.[Old Value],'') end) [Old Value],
(case 
     when T0.[Field Name] = 'SlpCode' then (SELECT top 1 TT0.[SlpName] FROM OSLP TT0 where TT0.[SlpCode] = T0.[New Value]) 
	 when T0.[Field Name] = 'GroupNum' then (SELECT top 1 TT0.[PymntGroup] FROM OCTG TT0 where TT0.[GroupNum] = T0.[New Value])
	 else ISNULL(T0.[New Value],'') end) [New Value]
into TempOCRDF
from TempOCRD T0 
order by T0.[BP Code], T0.LogInstanc   asc 

Delete from TempOCRDF where [New Value] = ''
Delete from TempOCRDF where [New Value] = '-1'
Delete from TempOCRDF where [Old Value] = '-1'
Delete from TempOCRDF where [Old Value] = [New Value]

select 0 LogInstanc , T0.Date , T0.[BP Code],T0.[BP Name]  ,T0.[BP Type] ,T0.[Created By],
case T0.[Field Name]
when 'e_Payment' then 'Payment Methods'
else
 T0.[Field Name] end [Field Name],
cast(T0.[Old Value] as nvarchar(max)) [Old Value], cast( T0.[New Value] as nvarchar(max)) [New Value] into #Final from TempOCRDF T0 
where T0.Date between @FromDate and @ToDate 
group by T0.LogInstanc , T0.Date , T0.[BP Code],T0.[BP Name]  ,T0.[BP Type] ,T0.[Created By], T0.[Field Name],
T0.[Old Value] , T0.[New Value]    
order by t0.[BP Code] , T0.[Date] desc

--select * from  TempOCRDPerson
--select * from  TempOCRD
--SELECT * FROM #Final

select * from(
select * from #Final
union all
select  0 LogInstanc , T0.Date , T0.[BP Code],T0.[BP Name]  ,T0.[BP Type] ,T0.[Created By], 'Bank Name' [Field Name],
cast(T1.BankName as nvarchar(max))    [Old Value] , cast(t2.BankName as nvarchar(max))  [New Value] from #Final T0 left outer join  odsc T1 on t0.[Old Value] = t1.BankCode 
left outer join odsc t2 on  t0.[New Value]  = t2.BankCode 
where T0.[Field Name] = 'Bank Code'
----union all
----select  0 LogInstanc , T0.Date , T0.[BP Code],T0.[BP Name]  ,T0.[BP Type] ,T0.[Created By], 'Account Name' [Field Name],
----cast(T1.AcctName as nvarchar(max))     [Old Value] , cast(t2.AcctName as nvarchar(max))  [New Value] from #Final T0 left outer join  OCRB T1 on t0.[Old Value] = t1.BankCode 
----left outer join OCRB t2 on  t0.[New Value]  = t2.BankCode 
----where T0.[Field Name] = 'Bank Code'
) as TmpResult 
group by TmpResult.LogInstanc , TmpResult.Date , TmpResult.[BP Code],TmpResult.[BP Name]  ,TmpResult.[BP Type] ,TmpResult.[Created By], TmpResult.[Field Name],
TmpResult.[Old Value] , TmpResult.[New Value]    
order by TmpResult.[BP Code] , TmpResult.[Date] desc

drop table #Final
Drop Table TempOCRDPerson
Drop Table #OCRDTemp
Drop Table #CRD1Temp
Drop Table TempOCRD
drop table TempACRD 
drop table TempCRD1
drop table TempACR
drop table TempOCRDF
drop table #ACR2Temp 
drop table #ACPRTemp
drop table TempARPC 
END





GO
/****** Object:  StoredProcedure [dbo].[BP_Audit_Report_26July17]    Script Date: 22/08/2017 11:41:16 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
--[BP_Audit_Report] 'T4000020','T4000073','20170601','20170628'

create PROCEDURE [dbo].[BP_Audit_Report_26July17]
		@FromCardcode Nvarchar(40) ,
		@ToCardcode Nvarchar(40),
		@FromDate Datetime,
		@ToDate Datetime


AS
BEGIN
--DECLARE @FromCardcode Nvarchar(40) ,
--		@ToCardcode Nvarchar(40),
--		@FromDate Datetime,
--		@ToDate Datetime

--		SET @FromCardcode = 'T4000020'
--		SET @ToCardcode = 'T4000020'
--		SET @FromDate = '20170628'
--		SET @ToDate = '20170628'

	-- SET NOCOUNT ON added to prevent extra result sets from
		-- interfering with SELECT statements.
	SET NOCOUNT ON;


if exists(select * from dbo.sysobjects t1 where t1.xtype = 'U' and t1.id = OBJECT_ID(N'#Final'))
begin
drop table #Final
end

if exists(select * from dbo.sysobjects t1 where t1.xtype = 'U' and t1.id = OBJECT_ID(N'TempACRD'))
begin
drop table TempACRD
end

if exists(select * from dbo.sysobjects t1 where t1.xtype = 'U' and t1.id = OBJECT_ID(N'TempOCRDPerson'))
begin
drop table TempOCRDPerson
end
if exists(select * from dbo.sysobjects t1 where t1.xtype = 'U' and t1.id = OBJECT_ID(N'#OCRDTemp'))
begin
drop table #OCRDTemp
end
if exists(select * from dbo.sysobjects t1 where t1.xtype = 'U' and t1.id = OBJECT_ID(N'#CRD1Temp'))
begin
drop table #CRD1Temp
end
if exists(select * from dbo.sysobjects t1 where t1.xtype = 'U' and t1.id = OBJECT_ID(N'TempOCRD'))
begin
drop table TempOCRD
end
if exists(select * from dbo.sysobjects t1 where t1.xtype = 'U' and t1.id = OBJECT_ID(N'TempCRD1'))
begin
drop table TempCRD1
end
if exists(select * from dbo.sysobjects t1 where t1.xtype = 'U' and t1.id = OBJECT_ID(N'TempACR'))
begin
drop table TempACR
end
if exists(select * from dbo.sysobjects t1 where t1.xtype = 'U' and t1.id = OBJECT_ID(N'TempOCRDF'))
begin
drop table TempOCRDF
end
if exists(select * from dbo.sysobjects t1 where t1.xtype = 'U' and t1.id = OBJECT_ID(N'#ACR2Temp'))
begin
drop table #ACR
end
if exists(select * from dbo.sysobjects t1 where t1.xtype = 'U' and t1.id = OBJECT_ID(N'#ACPRTemp'))
begin
drop table #ACPRTemp
end
if exists(select * from dbo.sysobjects t1 where t1.xtype = 'U' and t1.id = OBJECT_ID(N'TempARPC'))
begin
drop table TempARPC
end


  Select Row_Number() over (order by Col .name asc) as id,   isnull(des.Descr,Col .name) ColDesc, Col .name into #OCRDTemp
from sys .all_columns Col Inner join sys .tables tbl on Col .object_id = tbl .object_id 
left outer join CUFD des on des.TableID = tbl.name and   Col .name =  'U_'+ des.AliasID  
where tbl .name = 'OCRD' and Col .name Not in ('UpdateDate','CreateDate','LogInstanc','UserSign2','UserSign','Address','ZipCode','AddrType','Block','Building','City','Country',
'County','BillToDef','ZipCode','State1','Address','StreetNo','MailAddrTy','MailBlock','MailCity','MailCountr','MailCounty','ShipToDef','MailZipCod','State2','MailAddres',
'MailStrNo','MailBuildi','Gender'
,'AccCritria','AddID','Affiliate','AtcEntry','AutoCalBCG','AutoPost','AvrageLate','BackOrder','Balance','BalanceFC','BalanceSys','BalTrnsfrd'
,'BCACode','BlockDunn','BNKCounter','BoEDiscnt','BoEOnClct','BoEPrsnt','Box1099','Business','CardValid','CDPNum','CertBKeep','CertWHT','chainStore',
'ChecksBal','CollecAuth','CommGrCode','Commission','ConCerti','ConnBP','CpnNo','CrtfcateNO','DataSource','DatevFirst','DdctFileNo','DdctOffice',
'DdctPrcnt','DdgKey','DdtKey','DebPayAcct','Deleted','DocEntry','DpmClear','DpmIntAct','DscntObjct','DscntRel','DunnDate','DunnLevel','ITWTCode',
'KBKCode','LangCode','LetterNum','ListNum','LocMth','MainUsage','MivzExpSts','MltMthNum','MTHCounter','NINum','NTSWebSite','Number','ObjType','OKATO'
,'OKTMO','OpCode347','OprCount','OrderBalFC','OrderBalSy','OrdersBal','OtrCtlAcct','Pager','PartDelivr','PlngGroup','TaxIdIdent','TaxRndRule','ThreshOver'
,'TolrncDays','TpCusPres','TypeOfOp','TypWTReprt','SefazReply','SefazDate','SefazCheck','SCAdjust','RoleTypCod','RelCode','RcpntID','','''')

  Select Row_Number() over (order by Col .name asc) as id,   isnull(des.Descr,Col .name) ColDesc, Col .name into #CRD1Temp
from sys .all_columns Col Inner join sys .tables tbl on Col .object_id = tbl .object_id 
left outer join CUFD des on des.TableID = tbl.name and   Col .name =  'U_'+ des.AliasID  
where tbl .name = 'CRD1' and Col .name Not in ('CreateDate','UpdateDate','LogInstanc','UserSign2','UserSign','ObjType','LineNum','AdresType','CardCode','LicTradNum')

Select Row_Number() over (order by Col .name asc) as id,   isnull(des.Descr,Col .name) ColDesc, Col .name into #ACR2Temp
from sys .all_columns Col Inner join sys .tables tbl on Col .object_id = tbl .object_id 
left outer join CUFD des on des.TableID = tbl.name and   Col .name =  'U_'+ des.AliasID  
where tbl .name = 'ACR2' and Col .name Not in ('CreateDate','UpdateDate','LogInstanc','UserSign2','UserSign')

Select Row_Number() over (order by Col .name asc) as id,   isnull(des.Descr,Col .name) ColDesc, Col .name into #ACPRTemp
from sys .all_columns Col Inner join sys .tables tbl on Col .object_id = tbl .object_id 
left outer join CUFD des on des.TableID = tbl.name and   Col .name =  'U_'+ des.AliasID  
where tbl .name = 'ACPR' and Col .name Not in ('CreateDate','UpdateDate','LogInstanc','UserSign2','UserSign','updateTime','Active','DataSource',
'NFeRcpn','ObjType','Profession','CardCode','CntctCode','e_Payment','AcctName','Gender')

select ACRD .CardCode ,MAX (ACRD.LogInstanc) as LogInstanc into TempACRD
from ACRD where ACRD .CardCode >= @FromCardcode and ACRD .CardCode <= @ToCardcode  group by ACRD .CardCode

select ACRD .CardCode ,ACRD.LogInstanc as LogInstanc into TempACR
from ACRD where ACRD .CardCode >= @FromCardcode and ACRD .CardCode <= @ToCardcode  and ACRD.UpdateDate between @FromDate and @ToDate  
--and loginstanc = (select MAX(T0.LogInstanc) from ACRD  T0 where ACRD.CardCode = T0.CardCode )

select ACR1 .CardCode ,ACR1 .LineNum ,ACR1.LogInstanc as LogInstanc into TempCRD1
from ACR1 where ACR1 .CardCode >= @FromCardcode and ACR1 .CardCode <= @ToCardcode --group by ACR1 .CardCode ,ACR1 .LineNum 
and loginstanc in (select T0.LogInstanc from TempACR  T0 where ACR1.CardCode = T0.CardCode )

Select OCRD .LogInstanc+1 as LogInstanc,OCRD .CreateDate  as [Date],OCRD .CardCode as [BP Code],OCRD .CardName as [BP Name],
Case when CardType = 'C' then 'Customer' when CardType = 'S' then 'Vendor' else 'Lead' end [BP Type],
OUSR .U_NAME as [Created By],Convert(Nvarchar(MAX),null) as [Field Name],Convert (Nvarchar(MAX),'***Created***') as [Old Value],Convert (Nvarchar(MAX),null) as [New Value] into TempOCRD
from OCRD inner join OUSR on OCRD .UserSign = OUSR .INTERNAL_K 
where OCRD .CardCode >= @FromCardcode and  OCRD .CardCode <= @ToCardcode and 
Convert(Nvarchar(8), OCRD .CreateDate ,112) >= Convert(Nvarchar(8),  @FromDate ,112) and Convert(Nvarchar(8), OCRD .UpdateDate  ,112) <= Convert(Nvarchar(8),  @ToDate ,112);


Select OCRD .LogInstanc+1 as LogInstanc,OCRD .CreateDate  as [Date],OCRD .CardCode as [BP Code],OCRD .CardName as [BP Name],
Case when CardType = 'C' then 'Customer' when CardType = 'S' then 'Vendor' else 'Lead' end [BP Type],
OUSR .U_NAME as [Created By],Convert(Nvarchar(MAX),null) as [Field Name],Convert (Nvarchar(MAX),'***Created***') as [Old Value],Convert (Nvarchar(MAX),null) as [New Value] into TempOCRDPerson
from OCRD inner join OUSR on OCRD .UserSign = OUSR .INTERNAL_K 
where OCRD .CardCode >= @FromCardcode and  OCRD .CardCode <= @ToCardcode and 
Convert(Nvarchar(8), OCRD .CreateDate ,112) >= Convert(Nvarchar(8),  @FromDate ,112) and Convert(Nvarchar(8), OCRD .UpdateDate  ,112) <= Convert(Nvarchar(8),  @ToDate ,112);


Select OCRD .LogInstanc+1 as Seq ,OCRD .LogInstanc+1 as LogInstanc,OCRD .CreateDate  as [Date],OCRD .CardCode as [BP Code],OCRD .CardName as [BP Name],
Case when CardType = 'C' then 'Customer' when CardType = 'S' then 'Vendor' else 'Lead' end [BP Type],
OUSR .U_NAME as [Created By],Convert(Nvarchar(MAX),null) as [Field Name],Convert (Nvarchar(MAX),'***Created***') as [Old Value],Convert (Nvarchar(MAX),null) as [New Value] into TempARPC
from OCRD inner join OUSR on OCRD .UserSign = OUSR .INTERNAL_K 
where OCRD .CardCode >= @FromCardcode and  OCRD .CardCode <= @ToCardcode and 
Convert(Nvarchar(8), OCRD .CreateDate ,112) >= Convert(Nvarchar(8),  @FromDate ,112) and Convert(Nvarchar(8), OCRD .UpdateDate ,112) <= Convert(Nvarchar(8),  @ToDate ,112);

Declare @Count int,@Counter int,@ColName Nvarchar (100), @Query Nvarchar(max), @ColDesc Nvarchar (100), @Tmpcal NVarchar(10), @ColumnOLD NVarchar(200), @ColumnNEW NVarchar(200) ;

--select * from TempCrd1


Set @Count = (Select MAX (#OCRDTemp .id) from #OCRDTemp);
Set @Counter = 1;

While (@Counter <= @Count)

Begin

	Set @Query = '';

	Set @ColName = (Select top 1 Name from #OCRDTemp  where id = @Counter);
	Set @ColDesc = (Select top 1 ColDesc from #OCRDTemp  where id = @Counter)

	Set @Query = 'Insert into TempOCRD Select T0 .LogInstanc+1,T1 .UpdateDate  as [Date],T0 .CardCode as [BP Code],T0 .CardName as [BP Name],
	Case when T0 .CardType = ''C'' then ''Customer'' when T0 .CardType = ''S'' then ''Vendor'' else ''Lead'' end [BP Type],
	OUSR .U_NAME as [Created By],'''+ @ColDesc +''' as [Field Name],convert(Nvarchar(MAX),T0 .'+@ColName+',106) as [Old Value],Convert(Nvarchar(MAX),T1 .'+@ColName+',106) as [New Value]
	from ACRD T0 inner join ACRD T1 on T0 .CardCode = T1 .CardCode and T0 .LogInstanc = T1 .LogInstanc -1
	left outer join OUSR on T1 .UserSign2 = OUSR .INTERNAL_K 
	where  T0 .CardCode >= '''+@FromCardcode+''' and  T0 .CardCode <= '''+@ToCardcode+''' and   
	isnull(convert(Nvarchar(MAX),T0 .'+@ColName+',106) ,'''') <> isnull(convert(Nvarchar(MAX),T1 .'+@ColName+',106),'''')';
	
	--print @query
  Exec(@Query)

Set @Query = 'Insert into TempOCRD Select T1 .LogInstanc+1,T0 .UpdateDate  as [Date],T0 .CardCode as [BP Code],T0 .CardName as [BP Name],
Case when T0 .CardType = ''C'' then ''Customer'' when T0 .CardType = ''S'' then ''Vendor'' else ''Lead'' end [BP Type],
OUSR .U_NAME as [Created By],'''+ @ColDesc +''' as [Field Name],convert(Nvarchar(MAX),T1 .'+@ColName+',106) as [Old Value],Convert(Nvarchar(MAX),T0 .'+@ColName+',106) as [New Value]
from OCRD T0 inner join ACRD T1 on T0 .CardCode = T1 .CardCode 
inner join TempACRD T2 on T0.Cardcode = T2 .Cardcode and T1 .LogInstanc = T2.LogInstanc
left outer join OUSR on T0 .UserSign2 = OUSR .INTERNAL_K 
where  T0 .CardCode >= '''+@FromCardcode+''' and  T0 .CardCode <= '''+@ToCardcode+''' and  
isnull(convert(Nvarchar(MAX),T0 .'+@ColName+',106) ,'''') <> isnull(convert(Nvarchar(MAX),T1 .'+@ColName+',106),'''')';
--print @query
  Exec(@Query)

Set @Counter = @Counter + 1;
End 

--//Change log for Address
Set @Count = (Select MAX (#CRD1Temp .id) from #CRD1Temp);
Set @Counter = 1;

While (@Counter <= @Count)

Begin
	Set @Query = '';
	Set @ColName = (Select top 1 Name from #CRD1Temp  where id = @Counter);
	Set @ColDesc = (Select top 1 ColDesc from #CRD1Temp  where id = @Counter)

	Set @Query = 'Insert into TempOCRD Select distinct ACRD .LogInstanc,ACRD .UpdateDate  as [Date],ACRD .CardCode as [BP Code],ACRD .CardName as [BP Name],
Case when ACRD .CardType = ''C'' then ''Customer'' when ACRD .CardType = ''S'' then ''Vendor'' else ''Lead'' end [BP Type],
OUSR .U_NAME as [Created By],T1.Address+''-''+'''+ @ColDesc +''' as [Field Name],convert(Nvarchar(MAX),T0 .'+@ColName+',106) as [Old Value]
,Convert(Nvarchar(MAX),T3 .'+@ColName+',106) as [New Value]
from ACRD left outer join OUSR on ACRD .UserSign2 = OUSR .INTERNAL_K 
inner join ACR1 T0 on ACRD .Cardcode = T0 . Cardcode 
inner join ACR1 T1 on T0 .CardCode = T1 .CardCode and T0.Linenum = T1.Linenum and T0 .LogInstanc = T1 .LogInstanc 
 and ACRD .LogInstanc = (select MAX(loginstanc) from ACR1 where Address = T1.Address)
inner join CRD1 T3 on T3.CardCode = T1.CardCode and T1.LineNum = T3.LineNum 
where  T0 .CardCode >= '''+@FromCardcode+''' and  T0 .CardCode <= '''+@ToCardcode+'''  
--and isnull(convert(Nvarchar(MAX),T0 .'+@ColName+',106) ,'''') <> isnull(convert(Nvarchar(MAX),T1 .'+@ColName+',106),'''')
union
select distinct vw.LogInstanc,vw.Date,vw.[BP Code],vw.[BP Name],vw.[BP Type],vw.[Created By],vw.[Field Name],vw.[Old Value],vw.[New value] from (
Select T2 .LogInstanc,OCRD .UpdateDate  as [Date],OCRD .CardCode as [BP Code],OCRD .CardName as [BP Name],
Case when OCRD .CardType = ''C'' then ''Customer'' when OCRD .CardType = ''S'' then ''Vendor'' else ''Lead'' end [BP Type],
OUSR .U_NAME as [Created By],T2.Address +''-''+'''+ @ColDesc +''' as [Field Name],
'''' [Old Value]
,case  when isnull(convert(Nvarchar(MAX),T2 .'+@ColName+',106) ,'''') <> isnull(convert(Nvarchar(MAX),T1 .'+@ColName+',106),'''')
then Convert(Nvarchar(MAX),T2 .'+@ColName+',106) else '''' end  [New value]
--,isnull(Convert(Nvarchar(MAX),T1 .'+@ColName+',106),'''')[Filter]
,(select  count(Convert(Nvarchar(MAX),'+@ColName+',106)) from ACR1 
where Convert(Nvarchar(MAX),'+@ColName+',106) = Convert(Nvarchar(MAX),T2 .'+@ColName+',106) and cardcode = T2.cardcode  and Address = T2.Address) [Filter]
from OCRD left outer join OUSR on OCRD .UserSign2 = OUSR .INTERNAL_K 
left join  CRD1 T2 on OCRD .Cardcode = T2 . Cardcode 
left join ACR1 T0 on OCRD .Cardcode = T0 . Cardcode 
left join ACR1 T1 on T0 .CardCode = T1 .CardCode and T0.Linenum = T1.Linenum and T0 .LogInstanc = T1 .LogInstanc -1
 --and ACRD .LogInstanc = T1.LogInstanc
where  T2 .CardCode >= '''+@FromCardcode+''' and  T2 .CardCode <= '''+@ToCardcode+''' 
--AND convert(Nvarchar(MAX),T1 .'+@ColName+',106) not in (select convert(Nvarchar(MAX),'+@ColName+',106) 
--from CRD1 where CardCode >= '''+@FromCardcode+''' and CardCode <= '''+@ToCardcode+''')


) as vw where vw.[New value] <> '''' and vw.Filter = 0';
--print @query
  Exec(@Query)

----	Set @Query = 'Insert into TempOCRD Select t1 .LogInstanc+1 ,OCRD .UpdateDate  as [Date],OCRD .CardCode as [BP Code],OCRD .CardName as [BP Name],
----Case when OCRD .CardType = ''C'' then ''Customer'' when OCRD .CardType = ''S'' then ''Vendor'' else ''Lead'' end [BP Type],
----OUSR .U_NAME as [Created By],T0.Address+''-''+'''+ @ColDesc +''' as [Field Name],convert(Nvarchar(MAX),T1 .'+@ColName+',106) as [Old Value],Convert(Nvarchar(MAX),T0 .'+@ColName+',106) as [New Value]
----from OCRD Left outer join OUSR on OCRD .UserSign2 = OUSR .INTERNAL_K 
----inner join CRD1 T0 on OCRD .Cardcode = T0 . Cardcode 
----inner join ACR1 T1 on T0 .CardCode = T1 .CardCode and T0.Linenum = T1.Linenum 
------inner join TempCrd1 T2 on OCRD .Cardcode = T2 .Cardcode and T1 .LogInstanc = T2 .LogInstanc and T1 .Linenum = T2.LineNum
----where T0 .CardCode >= '''+@FromCardcode+''' and  T0 .CardCode <= '''+@ToCardcode+''' 
------and  
------isnull(convert(Nvarchar(MAX),T0 .'+@ColName+',106) ,'''') <> isnull(convert(Nvarchar(MAX),T1 .'+@ColName+',106),'''')
----';
	Set @Query = 'Insert into TempOCRD Select t1 .LogInstanc+1 ,OCRD .UpdateDate  as [Date],OCRD .CardCode as [BP Code],OCRD .CardName as [BP Name],
Case when OCRD .CardType = ''C'' then ''Customer'' when OCRD .CardType = ''S'' then ''Vendor'' else ''Lead'' end [BP Type],
OUSR .U_NAME as [Created By],T1.Address+''-''+'''+ @ColDesc +''' as [Field Name],convert(Nvarchar(MAX),T3 .'+@ColName+',106) as [Old Value],Convert(Nvarchar(MAX),T1 .'+@ColName+',106) as [New Value]
from OCRD Left outer join OUSR on OCRD .UserSign2 = OUSR .INTERNAL_K 
inner join TempCrd1 T2 on OCRD .Cardcode = T2 .Cardcode
inner join ACR1 T1 on T2 .CardCode = T1 .CardCode and T2.Linenum = T1.Linenum 
 and T1 .LogInstanc = T2 .LogInstanc 
 left outer join ACR1 T3 on T2 .CardCode = T3 .CardCode and T3.Linenum = T1.Linenum 
 and T3 .LogInstanc = T2 .LogInstanc -1
where T2 .CardCode >= '''+@FromCardcode+''' and  T2 .CardCode <= '''+@ToCardcode+''' 
and  
isnull(convert(Nvarchar(MAX),T3 .'+@ColName+',106) ,'''') <> isnull(convert(Nvarchar(MAX),T1 .'+@ColName+',106),'''')
';
 --print @query
 Exec(@Query)

Set @Counter = @Counter + 1;

End 

Set @Count = (Select max (TempACR.LogInstanc ) from TempACR);
Set @Counter = (Select min (TempACR.LogInstanc ) from TempACR);

While (@Counter <= @Count)

Begin
	Set @Query = ''
	set @Tmpcal = @Counter

	

set @Query = 'Insert into TempOCRD Select ACRD .LogInstanc,ACRD .UpdateDate  as [Date],ACRD .CardCode as [BP Code],ACRD .CardName as [BP Name],
Case when ACRD .CardType = ''C'' then ''Customer'' when ACRD .CardType = ''S'' then ''Vendor'' else ''Lead'' end [BP Type],
OUSR .U_NAME as [Created By],''PymCode'' as [Field Name],
--(select PymCode from ACR2 TT where TT.CardCode = ACRD.CardCode and  TT.LogInstanc = ACRD .LogInstanc -1 and TT.LineNum = 0) as [Old Value], 
--(select PymCode from ACR2 TT where TT.CardCode = ACRD.CardCode and  TT.LogInstanc = ACRD.LogInstanc and TT.LineNum = 0) as [New Value]

case when (select top 1 isnull(PymCode,'''') from ACR2 TT where TT.CardCode = ACRD.CardCode and  TT.LogInstanc = ACRD.LogInstanc and TT.LineNum = 0) <> '''' then 
(select top 1 PymCode from ACR2 TT where TT.CardCode = ACRD.CardCode and  TT.LogInstanc = ACRD.LogInstanc and TT.LineNum = 0) 
else (select top 1    PymCode from ACR2 TT 
where TT.CardCode = ACRD.CardCode and  TT.LogInstanc = ACRD.LogInstanc-1 and TT.LineNum = 0) end
as [Old Value], 
(select top 1 PymCode from CRD2 TT where TT.CardCode = ACRD.CardCode and TT.LineNum = 0) as [New Value]
from ACRD left outer join OUSR on ACRD .UserSign2 = OUSR .INTERNAL_K 
where ACRD .CardCode >= '''+@FromCardcode+''' and  ACRD .CardCode <= '''+@ToCardcode+''' and  ACRD .LogInstanc = '''+@Tmpcal+''''

--exec(@Query)
--PRINT @Query
set @Query = 'Insert into TempOCRD Select ACRD .LogInstanc,ACRD .UpdateDate  as [Date],ACRD .CardCode as [BP Code],ACRD .CardName as [BP Name],
Case when ACRD .CardType = ''C'' then ''Customer'' when ACRD .CardType = ''S'' then ''Vendor'' else ''Lead'' end [BP Type],
OUSR .U_NAME as [Created By],''PymCode'' as [Field Name],
--(select PymCode from ACR2 TT where TT.CardCode = ACRD.CardCode and  TT.LogInstanc = ACRD .LogInstanc -1 and TT.LineNum = 1) as [Old Value], 
--(select PymCode from ACR2 TT where TT.CardCode = ACRD.CardCode and  TT.LogInstanc = ACRD.LogInstanc and TT.LineNum = 1) as [New Value]

case when (select top 1 isnull(PymCode,'''') from ACR2 TT where TT.CardCode = ACRD.CardCode and  TT.LogInstanc = ACRD.LogInstanc and TT.LineNum = 1) <> '''' then 
(select top 1 PymCode from ACR2 TT where TT.CardCode = ACRD.CardCode and  TT.LogInstanc = ACRD.LogInstanc and TT.LineNum = 1) 
else (select top 1 PymCode from ACR2 TT 
where TT.CardCode = ACRD.CardCode and  TT.LogInstanc = ACRD.LogInstanc-1 and TT.LineNum = 1) end
as [Old Value], 
(select top 1 PymCode from CRD2 TT where TT.CardCode = ACRD.CardCode and TT.LineNum = 1) as [New Value]
NOTEPADfrom ACRD left outer join OUSR on ACRD .UserSign2 = OUSR .INTERNAL_K 
where ACRD .CardCode >= '''+@FromCardcode+''' and  ACRD .CardCode <= '''+@ToCardcode+''' and ACRD .LogInstanc = '''+@Tmpcal+''''

--exec(@Query)

set @Query = 'Insert into TempOCRD Select ACRD .LogInstanc,ACRD .UpdateDate  as [Date],ACRD .CardCode as [BP Code],ACRD .CardName as [BP Name],
Case when ACRD .CardType = ''C'' then ''Customer'' when ACRD .CardType = ''S'' then ''Vendor'' else ''Lead'' end [BP Type],
OUSR .U_NAME as [Created By],''PymCode'' as [Field Name],
--(select PymCode from ACR2 TT where TT.CardCode = ACRD.CardCode and  TT.LogInstanc = ACRD .LogInstanc -1 and TT.LineNum = 2) as [Old Value], 
--(select PymCode from ACR2 TT where TT.CardCode = ACRD.CardCode and  TT.LogInstanc = ACRD.LogInstanc and TT.LineNum = 2) as [New Value]

case when (select top 1 isnull(PymCode,'''') from ACR2 TT where TT.CardCode = ACRD.CardCode and  TT.LogInstanc = ACRD.LogInstanc and TT.LineNum = 2) <> '''' then 
(select top 1 PymCode from ACR2 TT where TT.CardCode = ACRD.CardCode and  TT.LogInstanc = ACRD.LogInstanc and TT.LineNum = 2) 
else (select top 1 PymCode from ACR2 TT 
where TT.CardCode = ACRD.CardCode and  TT.LogInstanc = ACRD.LogInstanc-1 and TT.LineNum = 2) end
as [Old Value], 
(select top 1 PymCode from CRD2 TT where TT.CardCode = ACRD.CardCode and TT.LineNum = 2) as [New Value]

from ACRD left outer join OUSR on ACRD .UserSign2 = OUSR .INTERNAL_K 
where ACRD .CardCode >= '''+@FromCardcode+''' and  ACRD .CardCode <= '''+@ToCardcode+''' and ACRD .LogInstanc = '''+@Tmpcal+''''
 exec(@Query)
--print  @Query

set @Query = 'Insert into TempOCRD Select ACRD .LogInstanc,ACRD .UpdateDate  as [Date],ACRD .CardCode as [BP Code],ACRD .CardName as [BP Name],
Case when ACRD .CardType = ''C'' then ''Customer'' when ACRD .CardType = ''S'' then ''Vendor'' else ''Lead'' end [BP Type],
OUSR .U_NAME as [Created By],''PymCode'' as [Field Name],
--(select PymCode from ACR2 TT where TT.CardCode = ACRD.CardCode and  TT.LogInstanc = ACRD .LogInstanc -1 and TT.LineNum = 3) as [Old Value], 
--(select PymCode from ACR2 TT where TT.CardCode = ACRD.CardCode and  TT.LogInstanc = ACRD.LogInstanc and TT.LineNum = 3) as [New Value]

case when (select top 1 isnull(PymCode,'''') from ACR2 TT where TT.CardCode = ACRD.CardCode and  TT.LogInstanc = ACRD.LogInstanc and TT.LineNum = 3) <> '''' then 
(select top 1 PymCode from ACR2 TT where TT.CardCode = ACRD.CardCode and  TT.LogInstanc = ACRD.LogInstanc and TT.LineNum = 3) 
else (select top 1 PymCode from ACR2 TT 
where TT.CardCode = ACRD.CardCode and  TT.LogInstanc = ACRD.LogInstanc-1 and TT.LineNum = 3) end
as [Old Value], 
(select top 1 PymCode from CRD2 TT where TT.CardCode = ACRD.CardCode and TT.LineNum = 3) as [New Value]

from ACRD left outer join OUSR on ACRD .UserSign2 = OUSR .INTERNAL_K 
where ACRD .CardCode >= '''+@FromCardcode+''' and  ACRD .CardCode <= '''+@ToCardcode+'''  and ACRD .LogInstanc = '''+@Tmpcal+''''

 exec(@Query)
--print  @Query
Set @Counter = @Counter + 1;
End 


Set @Count = (Select MAX (#ACPRTemp .id) from #ACPRTemp);
Set @Counter = 1;

While (@Counter <= @Count)

Begin
	Set @Query = '';
	Set @ColName = (Select Name from #ACPRTemp  where id = @Counter);
	Set @ColDesc = (Select ColDesc from #ACPRTemp  where id = @Counter)

--	Set @Query = 'Insert into TempOCRD Select ACRD .LogInstanc,ACRD .UpdateDate  as [Date],ACRD .CardCode as [BP Code],ACRD .CardName as [BP Name],
--Case when ACRD .CardType = ''C'' then ''Customer'' when ACRD .CardType = ''S'' then ''Vendor'' else ''Lead'' end [BP Type],
--OUSR .U_NAME as [Created By],'''+ @ColDesc +''' as [Field Name],convert(Nvarchar(MAX),T0 .'+@ColName+',106) as [Old Value],Convert(Nvarchar(MAX),T1 .'+@ColName+',106) as [New Value]
--from ACRD left outer join OUSR on ACRD .UserSign2 = OUSR .INTERNAL_K 
--inner join ACPR T0 on ACRD .Cardcode = T0 . Cardcode 
--inner join ACPR T1 on T0 .CardCode = T1 .CardCode and T0 .LogInstanc = T1 .LogInstanc -1 and ACRD .LogInstanc = T1.LogInstanc --and T0.Name = T1.Name 
--where  T0 .CardCode >= '''+@FromCardcode+''' and  T0 .CardCode <= '''+@ToCardcode+''' and 
--Convert(Nvarchar(8), ACRD .UpdateDate ,112)  >= '''+Convert(Nvarchar(8),  @FromDate ,112)+''' AND 
--Convert(Nvarchar(8), ACRD .UpdateDate ,112)  <= '''+Convert(Nvarchar(8),  @ToDate ,112)+''' AND 
--isnull(convert(Nvarchar(MAX),T0 .'+@ColName+',106) ,'''') <> isnull(convert(Nvarchar(MAX),T1 .'+@ColName+',106),'''')';


SET @Query = '; WITH T1 AS(SELECT DISTINCT T1 .LogInstanc,T0 .UpdateDate  as [Date],T0 .CardCode as [BP Code],T3 .CardName as [BP Name],
Case when T3 .CardType = ''C'' then ''Customer'' when T3 .CardType = ''S'' then ''Vendor'' else ''Lead'' end [BP Type],
T4 .U_NAME as [Created By],'''+ @ColDesc +''' as [Field Name],
-- case when 
--(select max(LogInstanc) from ACPR TT where TT.CardCode = T0.CardCode  and TT.CntctCode = T1.CntctCode 
--and TT.updateDate = T1.updateDate) = T1.LogInstanc then Convert(Nvarchar(MAX),T1 .'+@ColName+',106)
-- else ''''
-- end 
--  AS [Old Value]
(select TT. '+@ColName+' from ACPR TT where  TT.CntctCode = T1.CntctCode 
and TT.updateDate = T1.updateDate and LogInstanc = (select MAX(logInstanc)-1 from ACPR where ACPR.CntctCode = T1.CntctCode 
and ACPR.updateDate = T1.updateDate)) [Old Value],
  convert(Nvarchar(MAX),T0 .'+@ColName+',106) AS [New Value]
  FROM OCPR T0
INNER JOIN ACPR T5 ON T5.CardCode = T0.CardCode
LEFT JOIN ACPR T1 ON T1.CntctCode = T0.CntctCode
LEFT JOIN ACPR T2 ON T2.CntctCode = T1.CntctCode and T1.LogInstanc = T2.LogInstanc-1
LEFT JOIN ACRD T3 ON T3.CardCode = T1.CardCode AND T3.UpdateDate = T1.updateDate
LEFT JOIN OUSR T4 on T3 .UserSign2 = T4 .INTERNAL_K
WHERE  
T1.CntctCode = T0.CntctCode AND
--T1. '+@ColName+' <> T0. '+@ColName+' AND 
T0 .CardCode >= '''+@FromCardcode+''' and  T0 .CardCode <= '''+@ToCardcode+''' 
-- and
--Convert(Nvarchar(8), T0 .UpdateDate ,112)  >= '''+Convert(Nvarchar(8),  @FromDate ,112)+''' AND 
--Convert(Nvarchar(8), T0 .UpdateDate ,112)  <= '''+Convert(Nvarchar(8),  @ToDate ,112)+'''  
--AND T0.updateDate between ''2016-09-20'' and ''2016-09-20'' 
--AND isnull(convert(Nvarchar(MAX),T0 .'+@ColName+',106) ,'''') <> isnull(convert(Nvarchar(MAX),T1 .'+@ColName+',106),'''')
--AND MAX(T3.LogInstanc) = MAX(T3.LogInstanc)
AND T1.LogInstanc = (select max(LogInstanc) from ACPR TT where TT.CardCode = T0.CardCode  and TT.CntctCode = T1.CntctCode 
and TT.updateDate = T0.updateDate
) 
group by T1.CntctCode, T0.CardCode,T3.LogInstanc,T1.LogInstanc,T0.updateDate,
T1.updateDate,T3 .CardName,T3 .CardType,T4 .U_NAME , T1. '+@ColName+', T0. '+@ColName+', T2.'+@ColName+'

union 
Select vw.LogInstanc,vw.updateDate,vw.CardCode,vw.CardName,vw.[BP Type],vw.[Created By],vw.[Field Name],vw.[OldValue],vw.[New Value] from (
Select Distinct T0.[LogInstanc],T0.updateDate,T3.CardCode,T3.CardName, 
Case when T3 .CardType = ''C'' then ''Customer'' when T3 .CardType = ''S'' then ''Vendor'' else ''Lead'' end [BP Type],
T4 .U_NAME as [Created By],'''+ @ColDesc +''' as [Field Name],
isNull(Convert(Nvarchar(MAX),T0 .'+@ColName+',106),'''')  [New Value],isNull(Convert(Nvarchar(MAX),T1 .'+@ColName+',106),'''') [FilterName],'''' [OldValue] FROM  OCPR T0
LEFT JOIN ACPR T1 ON T1.CntctCode = T0.CntctCode
LEFT JOIN ACPR T2 ON T2.CntctCode = T1.CntctCode and T1.LogInstanc = T2.LogInstanc-1
LEFT JOIN OCRD T3 ON T3.CardCode = T0.CardCode 
LEFT JOIN OUSR T4 on T3 .UserSign2 = T4 .INTERNAL_K
Where T0 .CardCode >= '''+@FromCardcode+''' and  T0 .CardCode <= '''+@ToCardcode+'''   
--and Convert(Nvarchar(8), T0 .UpdateDate ,112)  >= '''+Convert(Nvarchar(8),  @FromDate ,112)+''' AND 
--Convert(Nvarchar(8), T0 .UpdateDate ,112)  <= '''+Convert(Nvarchar(8),  @ToDate ,112)+''' 
--AND T0.updateDate  between ''2016-09-20'' and ''2016-09-20''  

) as vw where vw.FilterName = '''' and vw.CardCode <> ''''
)
Insert into TempOCRD
SELECT * FROM T1 WHERE LogInstanc IN (SELECT LogInstanc FROM T1)'

--print @query

Exec(@Query)

Set @Counter = @Counter + 1;
End 
--Insert  into TempOCRD select * from TempOCRDPerson

Delete from TempOCRD where [Old Value] = '-1' and [New Value] = '-1'
--Delete from TempOCRD where ISNULL([Old Value],'') = ISNULL([New Value],'')

Select T0.LogInstanc , T0.Date , T0.[BP Code], T0.[BP Name] , T0.[BP Type] ,T0.[Created By] ,
(case T0.[Field Name]
when 'Notes' then 'Remarks' 
when 'Free_Text' then 'Remarks Tab - Remarks'
when 'frozenFor' then 'In Active'
when 'validFor' then 'Active'
when 'VatIdUnCmp' then 'Business Registration Number'
when 'ECVatGroup' then 'GST Code'
when 'Bill to-Block' then 'Block'
when 'Bill to-Building' then 'Building/Floor/Room'
when 'Bill to-City' then 'City'
when 'Bill to-Country' then 'Country'
when 'Bill to-County' then 'County'
when 'Bill to-Street' then 'Street / PO Box'
when 'Bill to-ZipCode' then 'Postcode'
when 'Bill to:-Block' then 'Block'
when 'Bill to:-Building' then 'Building/Floor/Room'
when 'Bill to:-City' then 'City'
when 'Bill to:-Country' then 'Country'
when 'Bill to:-County' then 'County'
when 'Bill to:-Street' then 'Street / PO Box'
when 'Bill to:-ZipCode' then 'Postcode'
when 'HousBnkAct' then 'Account'
when 'IntrntSite' then 'Web Site'
when 'LicTradNum' then 'GST Number'
when 'MandateID' then 'Sort code'
when 'Phone1' then 'Tel. 1'
when 'Phone2' then 'Tel. 2'
when 'PymCode' then 'e_Payment'
when 'BankCode' then 'Bank Code'
when 'BankCountr' then 'Bank Country' 
when 'CardName' then 'Name'
when 'Cellolar' then 'Mobile Phone'
when 'Cellular' then 'Mobile Phone'
when 'CntctPrsn' then 'Contact ID'
when 'DflAccount' then 'Account'
when 'DflBranch' then 'Branch'
when 'DflIBAN' then 'IBAN/ABA'
when 'DflSwift' then 'BIC/SWIFT Code'
when 'DflBankKey' then 'Default Bank Key'
when 'BankCtlKey' then 'Ctrl Int. ID'
when 'GroupNum' then 'Payment Terms Code'
when 'SlpCode' then 'Buyer'
when 'IndustryC' then 'Industry'
when 'FrozenComm' then 'Remarks'
when 'E_MailL' then 'E-Mail'
when 'e_Payment' then 'Payment Methods'
else T0.[Field Name]
end) [Field Name]
 , 
 (case 
   when T0.[Field Name] = 'SlpCode' then (SELECT top 1 TT0.[SlpName] FROM OSLP TT0 where TT0.[SlpCode] = T0.[Old Value]) 
   when T0.[Field Name] = 'GroupNum' then (SELECT top 1   TT0.[PymntGroup] FROM OCTG TT0 where TT0.[GroupNum] = T0.[Old Value])
   else iSNULL(T0.[Old Value],'') end) [Old Value],
(case 
     when T0.[Field Name] = 'SlpCode' then (SELECT top 1 TT0.[SlpName] FROM OSLP TT0 where TT0.[SlpCode] = T0.[New Value]) 
	 when T0.[Field Name] = 'GroupNum' then (SELECT top 1 TT0.[PymntGroup] FROM OCTG TT0 where TT0.[GroupNum] = T0.[New Value])
	 else ISNULL(T0.[New Value],'') end) [New Value]
into TempOCRDF
from TempOCRD T0 
order by T0.[BP Code], T0.LogInstanc   asc 
Delete from TempOCRDF where [New Value] = ''
Delete from TempOCRDF where [New Value] = '-1'
Delete from TempOCRDF where [Old Value] = '-1'
Delete from TempOCRDF where [Old Value] = [New Value]

select 0 LogInstanc , T0.Date , T0.[BP Code],T0.[BP Name]  ,T0.[BP Type] ,T0.[Created By],
case T0.[Field Name]
when 'e_Payment' then 'Payment Methods'
else
 T0.[Field Name] end [Field Name],
cast(T0.[Old Value] as nvarchar(max)) [Old Value], cast( T0.[New Value] as nvarchar(max)) [New Value] into #Final from TempOCRDF T0 
where T0.Date between @FromDate and @ToDate 
group by T0.LogInstanc , T0.Date , T0.[BP Code],T0.[BP Name]  ,T0.[BP Type] ,T0.[Created By], T0.[Field Name],
T0.[Old Value] , T0.[New Value]    
order by t0.[BP Code] , T0.[Date] desc

--select * from  TempOCRDPerson
--select * from  TempOCRD

select * from(
select * from #Final
union all
select  0 LogInstanc , T0.Date , T0.[BP Code],T0.[BP Name]  ,T0.[BP Type] ,T0.[Created By], 'Bank Name' [Field Name],
cast(T1.BankName as nvarchar(max))    [Old Value] , cast(t2.BankName as nvarchar(max))  [New Value] from #Final T0 left outer join  odsc T1 on t0.[Old Value] = t1.BankCode 
left outer join odsc t2 on  t0.[New Value]  = t2.BankCode 
where T0.[Field Name] = 'Bank Code'
----union all
----select  0 LogInstanc , T0.Date , T0.[BP Code],T0.[BP Name]  ,T0.[BP Type] ,T0.[Created By], 'Account Name' [Field Name],
----cast(T1.AcctName as nvarchar(max))     [Old Value] , cast(t2.AcctName as nvarchar(max))  [New Value] from #Final T0 left outer join  OCRB T1 on t0.[Old Value] = t1.BankCode 
----left outer join OCRB t2 on  t0.[New Value]  = t2.BankCode 
----where T0.[Field Name] = 'Bank Code'
) as TmpResult 
group by TmpResult.LogInstanc , TmpResult.Date , TmpResult.[BP Code],TmpResult.[BP Name]  ,TmpResult.[BP Type] ,TmpResult.[Created By], TmpResult.[Field Name],
TmpResult.[Old Value] , TmpResult.[New Value]    
order by TmpResult.[BP Code] , TmpResult.[Date] desc




drop table #Final

Drop Table TempOCRDPerson
Drop Table #OCRDTemp
Drop Table #CRD1Temp
Drop Table TempOCRD
drop table TempACRD 
drop table TempCRD1
drop table TempACR
drop table TempOCRDF
drop table #ACR2Temp 
drop table #ACPRTemp
drop table TempARPC 
END
GO
/****** Object:  StoredProcedure [dbo].[SBO_SP_TransactionNotification]    Script Date: 22/08/2017 11:41:16 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE proc [dbo].[SBO_SP_TransactionNotification] 

@object_type nvarchar(20), 				-- SBO Object Type
@transaction_type nchar(1),			-- [A]dd, [U]pdate, [D]elete, [C]ancel, C[L]ose
@num_of_cols_in_key int,
@list_of_key_cols_tab_del nvarchar(255),
@list_of_cols_val_tab_del nvarchar(255)

AS

begin

-- Return values
declare @error  int				-- Result (0 for no error)
declare @error_message nvarchar (200) 		-- Error string to be displayed
select @error = 0
select @error_message = N'Ok'

--------------------------------------------------------------------------------------------------------------------------------

--	ADD	YOUR	CODE	HERE
IF @transaction_type IN (N'A', N'U') AND
(@Object_type = N'20')
begin
if exists (SELECT T0.BaseEntry, SUM(T0.Quantity)
FROM dbo.PDN1 T0 INNER
JOIN dbo.POR1 T1 ON T1.DOCENTRY =
T0.BASEENTRY
WHERE T0.BaseType = 22 AND T0.ItemCode =
T1.ItemCode AND T0.BaseLine = T1.LineNum
and T0.DOCENTRY = @list_of_cols_val_tab_del
GROUP BY T0.BaseEntry
HAVING (SUM(T0.Quantity) > SUM(T1.Quantity)) or sum(t0.quantity) >
sum(t0.BaseOpnQty))
begin
select @Error = 10, @error_message = 'GRPO quantity is greater than PO
quantity'
end
end


----IF @transaction_type IN (N'A', N'U') AND
----(@Object_type = N'18')
----begin
----if exists (SELECT T0.BaseEntry, SUM(T0.LineTotal)
----FROM dbo.PCH1 T0 INNER
----JOIN dbo.PDN1 T1 ON T1.DOCENTRY =
----T0.BASEENTRY
----WHERE T0.BaseType = 20 AND T0.ItemCode =
----T1.ItemCode AND T0.BaseLine = T1.LineNum
----and T0.DOCENTRY = @list_of_cols_val_tab_del
----GROUP BY T0.BaseEntry
----HAVING (SUM(T0.LineTotal) > SUM(T1.LineTotal * 1.05)) or (SUM(T0.LineTotal) > SUM(T1.LineTotal + 1000)))
----begin
----select @Error = 10, @error_message = 'AP Invoice Amount is greater than GRPO
----Amount'
----end
----end
--------------------------------------------------------------------------------------------------------------------------------

---- ----AP Inv without GRPO -block
----IF @transaction_type =  'A' AND
----(@Object_type = N'18')
----begin
----if exists (SELECT T0.ItemCode 
----FROM dbo.PCH1 T0 
----Inner Join OITM T1 ON T0.Itemcode = T1.ItemCode
----WHERE T0.BaseType = -1 and T1.QryGroup1 <> 'Y'
----and T0.DOCENTRY = @list_of_cols_val_tab_del)

----begin
----select @Error = 18, @error_message = 'AP Invoice is without GRN'
----end
----end

----GRPO  without PO -block
IF @transaction_type =  'A' AND
(@Object_type = N'20')
begin
if exists (SELECT T0.ItemCode 
FROM dbo.PDN1 T0 
Inner Join OITM T1 ON T0.Itemcode = T1.ItemCode
WHERE T0.BaseType = -1  and T1.QryGroup3 <> 'Y'
and T0.DOCENTRY = @list_of_cols_val_tab_del)

begin
select @Error = 20, @error_message = 'GRN is without PO'
end
end

--show error message if the approval amount is zero---
----If  @transaction_type IN('A','U')  AND
----(@Object_type = '22')

----Begin

----If exists  (Select t.docentry

----From ODRF T

----where t.docentry=@list_of_cols_val_tab_del and
----(T.U_AB_APPROVALAMT=0))
----Begin
----Set @error=14
----Set @error_message =N'Add-on AE_PWC_IN01 disconnected. Please contact System Administrator.'
----end
----end

------ Blocking the PO / PR if the Approval amount is zero

IF @transaction_type IN (N'A', N'U') AND  @Object_type = N'112'
begin
if exists (SELECT T0.[DocNum] FROM ODRF T0 WHERE isnull(T0.[U_AB_APPROVALAMT],0)  = 0 
and T0.[Docentry] = @list_of_cols_val_tab_del and T0.[ObjType] in ('22'))
begin
select @Error = 1050, @error_message = 'Approval amount is zero, please log out SAP and login again '
end





if exists (SELECT T0.DocEntry FROM DRF1 T0 join DRF1 T1 on T0.DocEntry = T1.DocEntry WHERE isnull(T1.[U_AB_NONPROJECT],'')  = '' 
and T0.[Docentry] = @list_of_cols_val_tab_del and T0.[ObjType] in ('22','1470000113'))
begin
select @Error = 1050, @error_message = 'OU_BU Code is Empty... '
end

end


------ AP Invoice-OU_BU field mandatory

IF @transaction_type IN (N'A', N'U') AND  @Object_type = N'18'

begin

if exists (SELECT T1.DocEntry FROM PCH1 T1 WHERE isnull(T1.[U_AB_NONPROJECT],'')  = '' 

and T1.[Docentry] = @list_of_cols_val_tab_del)

begin

select @Error = 1050, @error_message = 'OU_BU Code is Empty...  '

end

end

 
---block document without approval----
IF @transaction_type IN (N'A', N'U') AND
(@Object_type = N'112')
begin
update DRF1 set [U_AB_GRPOSTAT]= 1 where DOCENTRY = @list_of_cols_val_tab_del and  ObjType in (22,1470000113)
end

IF @transaction_type IN (N'A', N'U') AND
(@Object_type = N'22')
begin
if exists (select DocEntry from POR1 where [U_AB_GRPOSTAT]= 0 and ObjType=22 and DOCENTRY = @list_of_cols_val_tab_del)
begin
select @Error = 10, @error_message = 'Approval required. Select Approving Department'
end
end

IF @transaction_type IN (N'A', N'U') AND
(@Object_type = N'1470000113')
begin
if exists (select DocEntry from PRQ1 where [U_AB_GRPOSTAT]= 0 and ObjType=1470000113 and DOCENTRY = @list_of_cols_val_tab_del)
begin
select @Error = 10, @error_message = 'Approval required. Select Approving Department.'
end
end




-- Select the return values
select @error, @error_message

end
GO
/****** Object:  StoredProcedure [dbo].[SP_AB_FORM_PurchaseRequest]    Script Date: 22/08/2017 11:41:16 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO




create PROCEDURE  [dbo].[SP_AB_FORM_PurchaseRequest]


@DocKey int

AS

BEGIN

Declare @SysCurr 	Char(3)
Declare @LocCurr	Char(3)
Declare @ComName	varchar(100)
Declare @ComAddr	varchar(200)
Declare @ComPhone	varchar(20)
Declare @ComFax  	varchar(20)


Set @SysCurr = (Select Top 1 SysCurrncy From OADM)
Set @LocCurr = (Select Top 1 MainCurncy From OADM)
set @ComName = (select Top 1 IsNull(OADM.PrintHeadr,OADM.CompnyName) from OADM)
set @ComAddr = (select Top 1 IsNull(OADM.CompnyAddr,OADM.CompnyAddr) from OADM)
set @ComPhone = (select Top 1 IsNull(OADM.Phone1,OADM.Phone1) from OADM)
set @ComFax =  (select Top 1 IsNull(OADM.Fax,OADM.Fax) from OADM)

SELECT 
	@ComName AS CompanyName,
	@ComAddr AS CompanyAddr,
	@ComPhone AS CompanyPhone,
	@ComFax AS CompanyFax,
	@SysCurr as CurrencySC,
	@LocCurr as CurrencyLC,
	(SELECT TOP 1 OADP.LogoImage from OADP) AS LogoImage,
	X.*,
	(case when X.Quantity <> 0 then X.RvLineTotalFC/X.Quantity else 0 end) as RvUnitPriceFC,
	(X.RvLineTotalFC + X.RvLineVatSumFC) as RvLineTotalIncVatSumFC,
	(X.RvLineTotal + X.RvLineVatSum) as RvLineTotalIncVatSum
FROM
(
	select	 --PurchaseRequest
		X.DocNum,
		X.DocDate,
		X.DocDueDate,
		'' as GSTRegNo,
		X.NumAtCard as CustRef,
		X.CardCode,
		X.CardName,
		X.Comments as Remark,
		X.Header as Header,
		X.Footer as Footer,
		X.Address as BillToAddr,
		X.Address2 as ShipToAddr,
				X.DocCur,
		(CASE WHEN X.DOCTOTALFC = 0 THEN (X.DOCTOTAL)
			ELSE (X.DocTotalFC) END) AS DocTotalFC,
		X.DocTotal, 
		(CASE WHEN X.DOCTOTALFC = 0 THEN (X.DocTotal - X.VatSum)
			ELSE (X.DocTotalFC-X.VatSumFC) END) AS QuotationAmountFC,
		(X.DocTotal - X.VatSum)	as InvoiceAmount,
		(CASE WHEN X.VatSumFC = 0 THEN 	(X.VatSum)
			ELSE (X.VatSumFC) END) AS VatSumFC,
		X.VatSum, 
		X.DiscSum as DiscountSum,
		case when X.DiscSumFC = 0 then X.DiscSum else X.DiscSumFC end as 'DiscountSumDocument',
		X.[Requester], 
		X.[ReqDate], 
		X.[TaxDate], 
		X1.LineVendor,
		X1.DocEntry,
		X1.VisOrder,
		X1.LineType,
		X1.ItemCode,
		OITM.ItemName,
		OITM.FrgnName as ItemFrgnName,
		X1.Dscription as Description,
		X1.Quantity,
		(X1.Quantity * X1.Price) as Total,
		X1.LineNum as LineNum,
		X1.WhsCode,
		(case when (OITM.ManSerNum='Y') then 1 else X1.Quantity end) as Quantity1,
		OITM.SalUnitMsr as QuantityUnitMsr,
		X1.Price as RvUnitPrice,
		(case when X1.Quantity <>0 then X1.LineTotal/X1.Quantity else 0 end) as RvUnitPriceLC,
		X1.Currency as RvLineCurrency,
		(case when X1.TotalFrgn=0 then X1.LineTotal else X1.TotalFrgn end) as RvLineTotalFC,
		X1.LineTotal as RvLineTotal,
		(case when X1.VatSumFrgn=0 then X1.VatSum else X1.VatSumFrgn end) as RvLineVatSumFC,
		x1.PriceAfVAT,
		X1.VatSum as RvLineVatSum,
		X1.VatPrcnt as VatPercent,
		X1.GTotal as RvLineTotalIncTax,
        X.PaytoCode,
       	X1.OcrCode,X1.OcrCode2,X1.OcrCode3,X1.OcrCode4,
        OUDP.Name as Department,
        Y.[LineSeq], 
        Y.[AftLineNum], 
        Y.[LineText]

	from OPRQ X
	inner join PRQ1 X1 on X1.DocEntry = X.DocEntry
	left join PRQ10 Y on X1.DocEntry = Y.Docentry and Y.AftLineNum = X1.VisOrder
	left join OITM  on X1.ItemCode = OITM.ItemCode
	----inner join OCRD on  OCRD.CardCode = X1.LineVendor  and OCRD.CardType = 'S'
	----inner join OCRD Z on  Z.CardName = X1.U_AI_Customer and Z.CardType = 'C'
	----inner join OCRY on OCRY.CODE = OCRD.Country
	----left join OCPR on OCPR.Name = OCRD.CntctPrsn and OCRD.CardCode = OCPR.CardCode
	----left join OPRJ on OPRJ.PrjCode =  X1.Project
	left join OUSR on ousr.U_Name =  X.ReqName
	left join OUDP on oudp.Code = OUSR.Department
	left join OHEM on ohem.userId = OUSR.USERID
	
	
	where
		X.DocEntry=@DocKey
		) X
order by  X.DocDate, X.DocNum, X.LineNum




End



GO
/****** Object:  StoredProcedure [dbo].[SP_AB_FRM_CHECKPRINTING]    Script Date: 22/08/2017 11:41:16 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create procedure [dbo].[SP_AB_FRM_CHECKPRINTING] as 
begin SELECT
	 *,
	 Substring((cONVERT(VARCHAR(10),T0.PmntDate,10)),
	 4,
	 1) AS 'D1',
	 Substring((cONVERT(VARCHAR(10),T0.PmntDate,10)),
	 5,
	 1) AS 'D2',
	 Substring((cONVERT(VARCHAR(10),T0.PmntDate,10)),
	 1,
	 1) AS 'M1',
	 Substring((cONVERT(VARCHAR(10),T0.PmntDate,10)),
	 2,
	 1) AS 'M2',
	 Substring((cONVERT(VARCHAR(10),T0.PmntDate,10)),
	 7,
	 1) AS 'Y1',
	Substring((cONVERT(VARCHAR(10),T0.PmntDate,10)),
	 8,
	 11) AS 'Y2',
	 T0.PmntDate
FROM OCHO T0 
INNER JOIN CHO1 T1 ON T0.CheckKey = T1.CheckKey 

End 


GO
/****** Object:  StoredProcedure [dbo].[SP_AB_FRM_PurchaseOrder]    Script Date: 22/08/2017 11:41:16 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE proc [dbo].[SP_AB_FRM_PurchaseOrder]
--exec sp_AB_FRM_PurchaseOrder 4
@DocKey numeric(18,0)

as

select 
T2.CompnyName,T2.Phone1,T2.Phone2,T2.GlblLocNum,T2.Fax, T2.FreeZoneNo, T2.TaxIdNum,T2.LogoImage,
T2.CompnyAddr, T2.BlockF,T2.StreetF,T2.Country,T0.DocDate,T0.DocDueDate,T2.ZipCode,T9.SeriesName,T9.[BeginStr],

------------approver--------------
T12.Stepcode as 'Stagecode', T12.UserId as 'ApproverCode', T13.U_Name as 'ApproverName',
T11.WtmCode as 'ApprovalTemplateCode',

-----------creator-----------


------------item detail--------------
T0.DocEntry,T1.ItemCode,T7.SuppCatNum,T7.FrgnName,
T1.Dscription,T1.PriceBefDi,T1.Quantity,T1.LineTotal,T1.VatSum,
T1.OcrCode, T1.OcrCode2, T1.OcrCode3, T1.OcrCode4,
T0.[Address],T0.Address2,T0.DocNum,T0.NumAtCard,T1.unitMsr,T1.ShipDate as 'DelDate',

-----------Outlet Details-------------
T10.WhsName,T10.StreetNo,T10.Block, T10.Street,T10.City,T10.ZipCode,T10.FedTaxID,T10.GlblLocNum,
-----------SalesPerson----------------
T0.TaxDate,T6.mobile,T6.Fax as 'Fax2',T6.email,T6.[firstName], T6.[middleName],

-----------BP info--------------
T4.CardName, T5.Name ContactName,T4.CntctPrsn,T0.ShipToCode,T4.E_Mail, T0.Footer, T0.Header,T4.Phone1 as 'PhoneCust', 
T0.CardCode, T0.Comments,T5.Cellolar,T0.DocCur,T4.Fax as BPFax,T8.PymntGroup,T0.U_AB_WTAX

from OPOR T0 with(nolock)
join POR1 T1 with(nolock) on T1.DocEntry=t0.DocEntry
join 
(
 select top(1) isnull(T0.PrintHeadr,T0.CompnyName) CompnyName,T0.Phone1,T0.Phone2,T1.GlblLocNum, T0.Fax, T0.FreeZoneNo, T0.TaxIdNum,T1.ZipCode,
  T0.CompnyAddr,T1.BlockF,T1.StreetF,T3.LogoImage,
  T2.Name Country
 from OADM T0 with(nolock) 
 join ADM1 T1 with(nolock) on 1=1
 join OCRY T2 with(nolock) on T2.Code=T1.Country
  join OADP T3 with(nolock) on 1=1
) T2 on 1=1
left join OHEM T6 with (nolock) on T6.empID=T0.OwnerCode
left join OCRD T4 with (nolock) on T4.CardCode=T0.CardCode
left join OCPR T5 with (nolock) on T5.CntctCode=T0.CntctCode
left join OITM T7 with (nolock) on T7.ItemCode=T1.ItemCode
left join OCTG T8 with (nolock) on T8.GroupNum=T0.GroupNum
left join NNM1 T9 with (nolock) on T9.Series=T0.Series
left join OWHS T10 with (nolock) on T10.WhsCode=T1.WhsCode
left join OWDD T11 with (nolock) on T0.DocEntry = T11.DocEntry
left  Join WDD1 T12 with (nolock) on T11.WddCode = T12.WddCode and  T11.ObjType = 22
left  join OUSR T13 with (nolock) on T12.userID = T13.UserId

left join  ODRF t14 with (nolock) on t14.docnum= t0.docnum



where T0.DocEntry=@DocKey
GO
/****** Object:  StoredProcedure [dbo].[SP_AB_FRM_PurchaseOrder_002]    Script Date: 22/08/2017 11:41:16 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE proc [dbo].[SP_AB_FRM_PurchaseOrder_002]
--exec sp_AB_FRM_PurchaseOrder_002 6
@DocKey numeric(18,0)

as

select 
T2.CompnyName,T2.Phone1,T2.Phone2,T2.GlblLocNum,T2.Fax, T2.FreeZoneNo, T2.TaxIdNum,T2.LogoImage,
T2.CompnyAddr, T2.BlockF,T2.StreetF,T2.Country,T0.DocDate,T0.DocDueDate,T2.ZipCode,T9.SeriesName,T9.[BeginStr],


------------approver--------------
T12.Stepcode as 'Stagecode', T12.UserId as 'ApproverCode',t12.UserID,
t13.U_name as 'approvername',
T11.WtmCode as 'ApprovalTemplateCode',

-------------creator-------
t14.CreatorName as CreatorName, T14.creatoremail as creatoremail,
------------item detail--------------
T0.DocEntry,T1.ItemCode,T7.SuppCatNum,T7.FrgnName,
T1.Dscription,T1.PriceBefDi,T1.Quantity,T1.LineTotal,T1.VatSum,
T1.OcrCode, T1.OcrCode2, T1.OcrCode3, T1.OcrCode4,
T0.[Address],T0.Address2,T0.DocNum,T0.NumAtCard,T1.unitMsr,T1.ShipDate as 'DelDate',

-----------Outlet Details-------------
T10.WhsName,T10.StreetNo,T10.Block, T10.Street,T10.City,T10.ZipCode,T10.FedTaxID,T10.GlblLocNum,
-----------SalesPerson----------------
T0.TaxDate,T6.mobile,T6.Fax as 'Fax2',T6.email,T6.[firstName], T6.[middleName],

-----------BP info--------------
T4.CardName, T5.Name ContactName,T4.CntctPrsn,T0.ShipToCode,T4.E_Mail, T0.Footer, T0.Header,T4.Phone1 as 'PhoneCust', 
T0.CardCode, T0.Comments,T5.Cellolar,T0.DocCur,T4.Fax as BPFax,T8.PymntGroup,T0.U_AB_WTAX

from OPOR T0 with(nolock)
join POR1 T1 with(nolock) on T1.DocEntry=t0.DocEntry
join 
(
 select top(1) isnull(T0.PrintHeadr,T0.CompnyName) CompnyName,T0.Phone1,T0.Phone2,T1.GlblLocNum, T0.Fax, T0.FreeZoneNo, T0.TaxIdNum,T1.ZipCode,
  T0.CompnyAddr,T1.BlockF,T1.StreetF,T3.LogoImage,
  T2.Name Country
 from OADM T0 with(nolock) 
 join ADM1 T1 with(nolock) on 1=1
 join OCRY T2 with(nolock) on T2.Code=T1.Country
  join OADP T3 with(nolock) on 1=1
) T2 on 1=1
left join OHEM T6 with (nolock) on T6.empID=T0.OwnerCode
left join OCRD T4 with (nolock) on T4.CardCode=T0.CardCode
left join OCPR T5 with (nolock) on T5.CntctCode=T0.CntctCode
left join OITM T7 with (nolock) on T7.ItemCode=T1.ItemCode
left join OCTG T8 with (nolock) on T8.GroupNum=T0.GroupNum
left join NNM1 T9 with (nolock) on T9.Series=T0.Series
left join OWHS T10 with (nolock) on T10.WhsCode=T1.WhsCode
left join OWDD T11 with (nolock) on T0.DocEntry = T11.DocEntry and t11.ObjType=22
left  Join WDD1 T12 with (nolock) on T11.WddCode = T12.WddCode and t11.ObjType=22
left  join OUSR T13 with (nolock) on t12.userid=t13.userid 
left join (select t2.U_NAME as CreatorName, t2.E_Mail as creatoremail,t1.DocNum from  OWDD t0
left join OPOR t1 on t0.Docentry=t1.docentry
left join ousr t2 on t2.userid=t0.OwnerID
where t0.ObjType =22) t14 on t14.docnum=t0.docnum	

 -- t12.userid=t13.userid 
--left join OUSR t14 with (nolock) on   t12.userid=t14.userid 



where T0.DocEntry=@DocKey


GO
/****** Object:  StoredProcedure [dbo].[SP_AB_FRM_PurchaseOrder_003]    Script Date: 22/08/2017 11:41:16 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create proc [dbo].[SP_AB_FRM_PurchaseOrder_003]
--exec sp_AB_FRM_PurchaseOrder_002 6
@DocKey numeric(18,0)

as

select 
T2.CompnyName,T2.Phone1,T2.Phone2,T2.GlblLocNum,T2.Fax, T2.FreeZoneNo, T2.TaxIdNum,T2.LogoImage,
T2.CompnyAddr, T2.BlockF,T2.StreetF,T2.Country,T0.DocDate,T0.DocDueDate,T2.ZipCode,T9.SeriesName,T9.[BeginStr],


------------approver--------------
T12.Stepcode as 'Stagecode', T12.UserId as 'ApproverCode',t12.UserID,
t13.U_name as 'approvername',
T11.WtmCode as 'ApprovalTemplateCode',

-------------creator-------
t14.CreatorName as CreatorName, T14.creatoremail as creatoremail,
------------item detail--------------
T0.DocEntry,T1.ItemCode,T7.SuppCatNum,T7.FrgnName,
T1.Dscription,T1.PriceBefDi,T1.Quantity,T1.LineTotal,T1.VatSum,
T1.OcrCode, T1.OcrCode2, T1.OcrCode3, T1.OcrCode4,
T0.[Address],T0.Address2,T0.DocNum,T0.NumAtCard,T1.unitMsr,T1.ShipDate as 'DelDate',

-----------Outlet Details-------------
T10.WhsName,T10.StreetNo,T10.Block, T10.Street,T10.City,T10.ZipCode,T10.FedTaxID,T10.GlblLocNum,
-----------SalesPerson----------------
T0.TaxDate,T6.mobile,T6.Fax as 'Fax2',T6.email,T6.[firstName], T6.[middleName],

-----------BP info--------------
T4.CardName, T5.Name ContactName,T4.CntctPrsn,T0.ShipToCode,T4.E_Mail, T0.Footer, T0.Header,T4.Phone1 as 'PhoneCust', 
T0.CardCode, T0.Comments,T5.Cellolar,T0.DocCur,T4.Fax as BPFax,T8.PymntGroup,T0.U_AB_WTAX

from OPOR T0 with(nolock)
join POR1 T1 with(nolock) on T1.DocEntry=t0.DocEntry
join 
(
 select top(1) isnull(T0.PrintHeadr,T0.CompnyName) CompnyName,T0.Phone1,T0.Phone2,T1.GlblLocNum, T0.Fax, T0.FreeZoneNo, T0.TaxIdNum,T1.ZipCode,
  T0.CompnyAddr,T1.BlockF,T1.StreetF,T3.LogoImage,
  T2.Name Country
 from OADM T0 with(nolock) 
 join ADM1 T1 with(nolock) on 1=1
 join OCRY T2 with(nolock) on T2.Code=T1.Country
  join OADP T3 with(nolock) on 1=1
) T2 on 1=1
left join OHEM T6 with (nolock) on T6.empID=T0.OwnerCode
left join OCRD T4 with (nolock) on T4.CardCode=T0.CardCode
left join OCPR T5 with (nolock) on T5.CntctCode=T0.CntctCode
left join OITM T7 with (nolock) on T7.ItemCode=T1.ItemCode
left join OCTG T8 with (nolock) on T8.GroupNum=T0.GroupNum
left join NNM1 T9 with (nolock) on T9.Series=T0.Series
left join OWHS T10 with (nolock) on T10.WhsCode=T1.WhsCode
left join OWDD T11 with (nolock) on T0.DocEntry = T11.DocEntry and t11.ObjType=22
left  Join WDD1 T12 with (nolock) on T11.WddCode = T12.WddCode and t11.ObjType=22
left  join OUSR T13 with (nolock) on t12.userid=t13.userid 
left join (select t2.U_NAME as CreatorName, t2.E_Mail as creatoremail,t1.DocNum from  OWDD t0
left join OPOR t1 on t0.Docentry=t1.docentry
left join ousr t2 on t2.userid=t0.OwnerID
where t0.ObjType =22) t14 on t14.docnum=t0.docnum	

 -- t12.userid=t13.userid 
--left join OUSR t14 with (nolock) on   t12.userid=t14.userid 



where T0.DocEntry=@DocKey


GO
/****** Object:  StoredProcedure [dbo].[SP_AB_FRM_PurchaseOrder_004]    Script Date: 22/08/2017 11:41:16 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
--exec [dbo].[SP_AB_FRM_PurchaseOrder_004] '7'
CREATE proc [dbo].[SP_AB_FRM_PurchaseOrder_004]
--sp_AB_FRM_PurchaseOrder 1738
@DocKey numeric(18,0)

as

begin 

Declare @StepCode varchar(100)
Declare @UserID varchar(10)
Declare @User_code varchar(50)
Declare @U_Name varchar(150)
Declare @WtmCode varchar(20)


------------ Getting Approver Information--------------
select TOP(1) @StepCode = TT1.StepCode , @UserID = TT1.UserID , @User_code = TT2.USER_CODE , @U_Name = TT2.U_NAME , @WtmCode = TT0.WtmCode 
from 
wdd1 TT1 join 
OWDD TT0 on TT1.WddCode = TT0.WddCode join 
OUSR TT2 on TT1.UserID = TT2.USERID   
where TT0.DocEntry = @DocKey and TT0.ObjType = 22 order by TT1.UpdateDate desc, TT1.updatetime desc
-------------------------------------------------------
select 
T2.CompnyName,T2.Phone1,T2.Phone2,T2.GlblLocNum,T2.Fax, T2.FreeZoneNo, T2.TaxIdNum,T2.LogoImage,
T2.CompnyAddr, T2.BlockF,T2.StreetF,T2.Country,T0.DocDate,T0.DocDueDate,T2.ZipCode,T9.SeriesName,T9.[BeginStr],
------------approver--------------
/*T12.Stepcode as 'Stagecode', T12.UserId as 'ApproverCode', T13.U_Name as 'ApproverName',
T11.WtmCode as 'ApprovalTemplateCode', */

@StepCode as 'Stagecode', @UserID as 'ApproverCode', @U_Name as 'ApproverName',
@WtmCode as 'ApprovalTemplateCode',
------------item detail--------------
T0.DocEntry,T1.ItemCode,T7.SuppCatNum,T7.FrgnName,
T1.Dscription,t16.LineText,T1.PriceBefDi,T1.Quantity,
case when t0.DocCur<>'SGD' then t1.vatsumsy
else t1.vatsum end as 'vatsum',
case when t0.DocCur<>'SGD'then t1.TotalFrgn
else T1.LineTotal end as 'linetotal',--T1.VatSum,
T1.OcrCode, T1.OcrCode2, T1.OcrCode3, T1.OcrCode4,
T0.[Address],T0.Address2,T0.DocNum,T0.NumAtCard,T1.unitMsr,T1.ShipDate as 'DelDate', TT.[U_NAME] [CreatorName], TT.[USER_CODE] [CreatorCode]
,case when t0.DocCur<>'SGD' then t0.DocTotalsy
else t0.doctotal end as 'doctotal',t1.U_AB_CONSOLIDATION,
t0.doctotal as 'LCtotal', t1.Linetotal as 'LClinetotal', t1.vatsum as'lcvatsum',t1.visorder,

-----------Outlet Details-------------
T10.WhsName,T10.StreetNo,T10.Block, T10.Street,T10.City,T10.ZipCode,T10.FedTaxID,T10.GlblLocNum,
-----------SalesPerson----------------
T0.TaxDate,T6.mobile,T6.Fax as 'Fax2',T6.email,T6.[firstName], T6.[middleName],

-----------BP info--------------
T4.CardName, T5.Name  as ContactName,T4.CntctPrsn,T0.ShipToCode,
case when t5.E_MailL!='' then t5.E_MailL
else t4.E_Mail end as E_mail,
T0.Footer, T0.Header,T4.Phone1 as 'PhoneCust', 
T0.CardCode, T0.Comments,T5.Cellolar,T0.DocCur,T4.Fax as BPFax,T8.PymntGroup,T0.U_AB_WTAX
from OPOR T0 with(nolock)
join POR1 T1 with(nolock) on T1.DocEntry=t0.DocEntry
---added on 29/04/2015-----
left join
(
	select a.DocEntry,a.AftLineNum,min(LineSeq) as LineSeq,d.LineText
	from POR10 A
	Cross Apply
	(
		Select STUFF
		((
			Select cast(LineText as nvarchar(max)) + Char(10)
			From por10 B
			where a.docentry =b.docentry and a.AftLineNum = b.AftLineNum
			order by a.docentry,b.ordernum
			for XML PAth(''),TYPE).value('.','NVARCHAR(MAX)'),1,0,'')
	) D(LineText)
	Group By a.docentry,a.aftlinenum,d.linetext
) T16 on t1.docentry= t16.docentry and t16.aftlinenum = t1.VisOrder
join OUSR TT with(nolock) on TT.USERID = T0.usersign 
join 
(
 select top(1) isnull(T0.PrintHeadr,T0.CompnyName) CompnyName,T0.Phone1,T0.Phone2,T1.GlblLocNum, T0.Fax, T0.FreeZoneNo, T0.TaxIdNum,T1.ZipCode,
  T0.CompnyAddr,T1.BlockF,T1.StreetF,T3.LogoImage,
  T2.Name Country
 from OADM T0 with(nolock) 
 join ADM1 T1 with(nolock) on 1=1
 join OCRY T2 with(nolock) on T2.Code=T1.Country
  join OADP T3 with(nolock) on 1=1
) T2 on 1=1
left join OHEM T6 with (nolock) on T6.empID=T0.OwnerCode
left join OCRD T4 with (nolock) on T4.CardCode=T0.CardCode
left join OCPR T5 with (nolock) on T5.CntctCode=T0.CntctCode
left join OITM T7 with (nolock) on T7.ItemCode=T1.ItemCode
left join OCTG T8 with (nolock) on T8.GroupNum=T0.GroupNum
left join NNM1 T9 with (nolock) on T9.Series=T0.Series
left join OWHS T10 with (nolock) on T10.WhsCode=T1.WhsCode
/*
left join OWDD T11 with (nolock) on T0.DocEntry = T11.DocEntry
inner  Join WDD1 T12 with (nolock) on T11.WddCode = T12.WddCode and  T11.ObjType = 22
left join OUSR T13 with (nolock) on T12.userID = T13.UserId 
*/
where T0.DocEntry=@DocKey

end
GO
/****** Object:  StoredProcedure [dbo].[SP_AB_FRM_PurchaseOrder_noapporal]    Script Date: 22/08/2017 11:41:16 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE proc [dbo].[SP_AB_FRM_PurchaseOrder_noapporal]
--exec sp_AB_FRM_PurchaseOrder 7
@DocKey numeric(18,0)

as

select 
T2.CompnyName,T2.Phone1,T2.Phone2,T2.GlblLocNum,T2.Fax, T2.FreeZoneNo, T2.TaxIdNum,T2.LogoImage,
T2.CompnyAddr, T2.BlockF,T2.StreetF,T2.Country,T0.DocDate,T0.DocDueDate,T2.ZipCode,T9.SeriesName,T9.[BeginStr],

------------approver--------------
--T12.Stepcode as 'Stagecode', T12.UserId as 'ApproverCode', T13.U_Name as 'ApproverName',
--T11.WtmCode as 'ApprovalTemplateCode',

-----------creator-----------
t6.u_name as creatorname,t6.e_mail as creatormail,

------------item detail--------------
T0.DocEntry,T1.ItemCode,T7.SuppCatNum,T7.FrgnName,
T1.Dscription,T1.PriceBefDi,T1.Quantity,T1.LineTotal,T1.VatSum,
T1.OcrCode, T1.OcrCode2, T1.OcrCode3, T1.OcrCode4,
T0.[Address],T0.Address2,T0.DocNum,T0.NumAtCard,T1.unitMsr,T1.ShipDate as 'DelDate',

-----------Outlet Details-------------
T10.WhsName,T10.StreetNo,T10.Block, T10.Street,T10.City,T10.ZipCode,T10.FedTaxID,T10.GlblLocNum,
-----------SalesPerson----------------
--T0.TaxDate,T6.mobile,T6.Fax as 'Fax2',T6.email,T6.[firstName], T6.[middleName],

-----------BP info--------------
T4.CardName, T5.Name ContactName,T4.CntctPrsn,T0.ShipToCode,T4.E_Mail, T0.Footer, T0.Header,T4.Phone1 as 'PhoneCust', 
T0.CardCode, T0.Comments,T5.Cellolar,T0.DocCur,T4.Fax as BPFax,T8.PymntGroup,T0.U_AB_WTAX

from OPOR T0 with(nolock)
join POR1 T1 with(nolock) on T1.DocEntry=t0.DocEntry
join 
(
 select top(1) isnull(T0.PrintHeadr,T0.CompnyName) CompnyName,T0.Phone1,T0.Phone2,T1.GlblLocNum, T0.Fax, T0.FreeZoneNo, T0.TaxIdNum,T1.ZipCode,
  T0.CompnyAddr,T1.BlockF,T1.StreetF,T3.LogoImage,
  T2.Name Country
 from OADM T0 with(nolock) 
 join ADM1 T1 with(nolock) on 1=1
 join OCRY T2 with(nolock) on T2.Code=T1.Country
  join OADP T3 with(nolock) on 1=1
) T2 on 1=1
left join OUSR T6 with (nolock) on T6.UserID=T0.UserSign2
left join OCRD T4 with (nolock) on T4.CardCode=T0.CardCode
left join OCPR T5 with (nolock) on T5.CntctCode=T0.CntctCode
left join OITM T7 with (nolock) on T7.ItemCode=T1.ItemCode
left join OCTG T8 with (nolock) on T8.GroupNum=T0.GroupNum
left join NNM1 T9 with (nolock) on T9.Series=T0.Series
left join OWHS T10 with (nolock) on T10.WhsCode=T1.WhsCode
--left join OWDD T11 with (nolock) on T0.DocEntry = T11.DocEntry
--left  Join WDD1 T12 with (nolock) on T11.WddCode = T12.WddCode and  T11.ObjType = 22
--left  join OUSR T13 with (nolock) on T12.userID = T13.UserId
--left join ousr t14 with (nolock) on t14.TPLId=t0.StationID 

where T0.DocEntry=@DocKey
GO
/****** Object:  UserDefinedFunction [dbo].[AE_FN003_BUDGET_COMMITTEDAMOUNT]    Script Date: 22/08/2017 11:41:16 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


--AE_SP005_BUDGET_COMMITTEDAMOUNT'PWCL',2014,'False','IFS 717','71111100'

--select [dbo].[AE_FN003_BUDGET_COMMITTEDAMOUNT]('2016','BU','52723','','71181300')

CREATE function [dbo].[AE_FN003_BUDGET_COMMITTEDAMOUNT]
(
@Year varchar(10),
@Cat varchar(10),
@BU varchar(100),
@Project varchar(100),
@GLCode varchar(100)
)
returns decimal(19,2)
as
begin

Declare @Column1 as decimal(19,2)
Declare @Column2 as decimal(19,2)
Declare @Column3 as decimal(19,2)
Declare @Column4 as decimal(19,2)
Declare @Column5 as decimal(19,2)
Declare @Column6 as decimal(19,2)
Declare @Column7 as decimal(19,2)
Declare @Column8 as decimal(19,2)
Declare @Dimension Varchar(100)
Declare @DateF varchar(30)
Declare @DateT varchar(30)
   
if @cat = 'PRJ'
 begin
  set @Dimension = @Project
 end
else
 begin
  set @Dimension = @BU
 end

set @DateF = cast(cast(@year as numeric ) -1 AS VARCHAR) +'0701'
 set @DateT =  @year + '0630'


   --- Add PO Draft with status Open - manually -save as draft
select @Column1 = isnull((SELECT sum(T1.LineTotal ) LineTotal
 FROM ODRF T0  INNER JOIN DRF1 T1 ON T0.[DocEntry] = T1.[DocEntry] 
left outer join  OWDD T2 on T0.DocEntry = T2.DocEntry 
left outer join   OWTM T3 on T2.[WtmCode] = T3.[WtmCode] 
where T1.LineStatus  = 'O' and isnull(T2.[Status],'') = '' and T0.DocStatus = 'O' and T0.ObjType in ('22','1470000113')
and T0.DocDate between @datef and @DateT  and isnull(T3.[Active],'Y') = 'Y' and
(case when @Cat = 'PRJ' then isnull(T1.Project,'') else  isnull(T1.U_AB_NONPROJECT ,'')  end) = @Dimension
and isnull(T1.AcctCode ,'') = @GLCode
 and (case when @Cat <> 'PRJ' then isnull(T1.Project,'') else  ''  end) = ''
),0) , 

--- Add PO Draft with status Open  - Pending for Approval / Approved but not converted to PO
@Column2 = isnull((SELECT sum(T1.LineTotal) LineTotal  
 FROM ODRF T0  INNER JOIN DRF1 T1 ON T0.[DocEntry] = T1.[DocEntry] 
left outer join  OWDD T2 on T0.DocEntry = T2.DocEntry 
left outer join   OWTM T3 on T2.[WtmCode] = T3.[WtmCode]
where T1.LineStatus = 'O' and isnull(T2.[Status],'') <> 'N' and T0.DocStatus = 'O' and T0.ObjType in ('22','1470000113')
and T0.DocDate  between @datef and @DateT and isnull(T3.[Active],'') = 'Y'
and (case when @Cat = 'PRJ' then isnull(T1.Project,'') else  isnull(T1.U_AB_NONPROJECT ,'')  end) = @Dimension
and isnull(T1.AcctCode ,'') = @GLCode
and (case when @Cat <> 'PRJ' then isnull(T1.Project,'') else  ''  end) = ''
 ),0)  ,

  --- Add PR Document with status Open
@Column3 = isnull((select sum( case when isnull(T2.LineTotal,0) = 0 then  T1.LineTotal  else 0 end  )  from OPRQ T0 join PRQ1 T1 on T0.DocEntry = T1.DocEntry 
left outer join DRF1 T2 on T2.BaseEntry = T1.DocEntry and T2.BaseLine = T1.LineNum and T2.BaseType = '1470000113'
left outer join ODRF T3 on T2.DocEntry = T3.Docentry and T3.DocStatus = 'O'
where T1.LineStatus = 'O' and T0.DocStatus = 'O'
and T0.DocDate between @datef and @DateT and
(case when @Cat = 'PRJ' then isnull(T1.Project,'') else  isnull(T1.U_AB_NONPROJECT ,'')  end) = @Dimension
and isnull(T1.AcctCode ,'') = @GLCode
and (case when @Cat <> 'PRJ' then isnull(T1.Project,'') else  ''  end) = ''
),0) ,

--- Add PO Document with status Open
@Column4 = isnull((select sum(T1.LineTotal) LineTotal  from OPOR T0 join POR1 T1 on T0.DocEntry = T1.DocEntry where T1.LineStatus = 'O' and T0.DocStatus = 'O'
and T0.DocDate between @datef and @DateT
and (case when @Cat = 'PRJ' then isnull(T1.Project,'') else  isnull(T1.U_AB_NONPROJECT ,'')  end) = @Dimension
and isnull(T1.AcctCode ,'') = @GLCode
and (case when @Cat <> 'PRJ' then isnull(T1.Project,'') else  ''  end) = ''
),0) ,

--- Add PO Document with status Closed But GRN in Open Status
@Column5 = isnull((SELECT sum(T1.[LineTotal]) LineTotal FROM OPOR T0  INNER JOIN POR1 T1 ON T0.[DocEntry] = T1.[DocEntry] join 
 PDN1 T2 on T2.BaseEntry = T1.DocEntry and T2.BaseLine = T1.LineNum  INNER JOIN OPDN T3 ON T2.[DocEntry] = T3.[DocEntry] WHERE T2.[LineStatus] = 'O'
 and T0.DocDate between @datef and @DateT
 and (case when @Cat = 'PRJ' then isnull(T1.Project,'') else  isnull(T1.U_AB_NONPROJECT ,'')  end) = @Dimension
 and isnull(T1.AcctCode ,'') = @GLCode
 and (case when @Cat <> 'PRJ' then isnull(T1.Project,'') else  ''  end) = ''
 ),0) ,

--- Less PO Draft with status Open  - Approval status Rejected
@Column6 = isnull((SELECT sum(T1.LineTotal) LineTotal  
 FROM ODRF T0  INNER JOIN  DRF1 T1 ON T0.[DocEntry] = T1.[DocEntry] 
left outer join  OWDD T2 on T0.DocEntry = T2.DocEntry 
left outer join   OWTM T3 on T2.[WtmCode] = T3.[WtmCode]
where T1.LineStatus = 'O' and isnull(T2.[Status],'') = 'N' and T0.DocStatus = 'O' and T0.ObjType in ('22','1470000113')
and T0.DocDate between @datef and @DateT and isnull(T3.[Active],'Y') = 'Y'
and (case when @Cat = 'PRJ' then isnull(T1.Project,'') else  isnull(T1.U_AB_NONPROJECT ,'')  end) = @Dimension
and isnull(T1.AcctCode ,'') = @GLCode
and (case when @Cat <> 'PRJ' then isnull(T1.Project,'') else  ''  end) = ''
),0) ,

--- Less PO Document with status Cancel  - User manually Cancel
@Column7 =  isnull((select sum(T1.LineTotal) LineTotal   from OPOR T0 join POR1 T1 on T0.DocEntry = T1.DocEntry where T1.LineStatus = 'C' and T0.CANCELED = 'Y'
and T0.DocDate between @datef and @DateT
and (case when @Cat = 'PRJ' then isnull(T1.Project,'') else  isnull(T1.U_AB_NONPROJECT ,'')  end) = @Dimension
and isnull(T1.AcctCode ,'') = @GLCode
and (case when @Cat <> 'PRJ' then isnull(T1.Project,'') else  ''  end) = ''
),0) ,

--- Less PO Document with status Close  - User manually Close
@Column8  = isnull((select sum(T1.LineTotal) LineTotal   from OPOR T0 join POR1 T1 on T0.DocEntry = T1.DocEntry where T1.LineStatus = 'C' and T0.CANCELED = 'N' 
and T1.TargetType = -1
and T0.DocDate between @datef and @DateT 
and (case when @Cat = 'PRJ' then isnull(T1.Project,'') else  isnull(T1.U_AB_NONPROJECT ,'')  end) = @Dimension
and isnull(T1.AcctCode ,'') = @GLCode
and (case when @Cat <> 'PRJ' then isnull(T1.Project,'') else  ''  end) = ''
),0) 

--return ((@Column1 + @Column2 + @Column3 + @Column4 ) - @Column5 - @Column6 - @Column7 )
return ( @Column2 + @Column3 + @Column4 + @Column5 )
end

GO
/****** Object:  UserDefinedFunction [dbo].[AE_FN003_GetApprover]    Script Date: 22/08/2017 11:41:16 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE FUNCTION [dbo].[AE_FN003_GetApprover](@DocEntry varchar(30),@Level varchar(30))
RETURNS varchar(1000)
AS 
-- Returns the stock level for the product.
BEGIN
Declare @Approver varchar(300)
Declare @count integer
declare @TempApproverMatrix table (SortId INT, [USER_CODE] Varchar(300), Name varchar(max), Docentry varchar(20), Status varchar(20))
declare @TempApprover table (ID INT, Name varchar(max))

Insert into @TempApproverMatrix 
SELECT 
 T4.SortId, T3.[USER_CODE], T3.[U_NAME],
  T0.Docentry   , T2.Status    FROM 
OPOR T0 join OWDD T1 on T0.DocEntry = T1.Docentry inner join 
WDD1 T2 on T1.WddCode = T2.WddCode INNER JOIN OUSR T3 ON T2.[USERID] = T3.[USERID]  join WTM2 T4 ON T4.WstCode = T2.StepCode
AND T1.[WtmCode] = T4.[WtmCode]
WHERE T0.Docentry = @DocEntry and T2.Status = 'Y' and T1.[ObjType] = '22'

Insert into @TempApprover 
Select distinct ST2.SortId  ,

    substring(
        (
            Select ','+ST1.Name  AS [text()]
            From @TempApproverMatrix ST1
            Where ST1.SortId  = ST2.SortId 
            ORDER BY ST1.SortId
            For XML PATH ('')
        ), 2, 1000) [Name]
From @TempApproverMatrix ST2

select @Approver = isnull(Name,'') from @TempApprover where ID = @level

--if isnull(@Approver,'') = '' and @level > 1
--begin
-- set @Approver = 'NA'
--end

RETURN @Approver 
END;

GO
/****** Object:  UserDefinedFunction [dbo].[AE_FN004_BUDGET_ACTUALSPEND]    Script Date: 22/08/2017 11:41:16 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO




--AE_SP006_BUDGET_ACTUALSPEND'PWCL',2016,'True','Prj001','71111100'


--select [dbo].[AE_FN004_BUDGET_ACTUALSPEND]('2016','PRJ','','Prj001','71111100')
create function [dbo].[AE_FN004_BUDGET_ACTUALSPEND]
(
@Year varchar(10),
@Cat Varchar(100),
@BU varchar(100),
@Project varchar(100),
@GLCode  varchar(100)
)
Returns Decimal(19,2)
as begin
Declare @Column1 as decimal(19,2)
Declare @Column2 as decimal(19,2)
Declare @Column3 as decimal(19,2)
Declare @Column4 as decimal(19,2)
Declare @Column5 as decimal(19,2)
Declare @Dimension Varchar(100)
Declare @DateF varchar(30)
Declare @DateT varchar(30)

if @cat = 'PRJ'
 begin
  set @Dimension = @Project
 end
else
 begin
  set @Dimension = @BU
 end
 
set @DateF = cast(cast(@year as numeric ) -1 AS VARCHAR) +'0701'
 set @DateT =  @year + '0630'

--- Add AP INVOICE with status Open - Base Document PO,GRPO & Direct
select @Column1= isnull((SELECT sum(T1.[LineTotal]) FROM OPCH T0  INNER JOIN PCH1 T1 ON T0.[DocEntry] = T1.[DocEntry]
WHERE t1.LineStatus = 'O' AND T1.BaseType IN ('20','22','-1')
and T0.DocDate between @datef and @DateT
and (case when @Cat = 'PRJ' then isnull(T1.Project,'') else  isnull(T1.U_AB_NONPROJECT ,'')  end) = @Dimension
and isnull(T1.AcctCode ,'') = @GLCode
and (case when @Cat <> 'PRJ' then isnull(T1.Project,'') else  ''  end) = '')
,0) 
, 
--- Add JE & Direct
@Column2 = isnull((SELECT sum(T2.[Debit] - T2.[Credit]) FROM [OJDT]  T1 INNER JOIN JDT1 T2 
ON T1.[TransId] = T2.[TransId] WHERE T1.[TransType] = 30
and T1.[TransId] NOT IN (SELECT StornoToTr FROM ojdt WHERE ISNULL(StornoToTr,'') <> '') AND T1.[StornoToTr] IS NULL 
and T1.TaxDate between @datef and @DateT
and (case when @Cat = 'PRJ' then isnull(T2.Project,'') else  isnull(T2.U_AB_NONPROJECT ,'')  end) = @Dimension
and isnull(T2.Account ,'') = @GLCode
and (case when @Cat <> 'PRJ' then isnull(T2.Project,'') else  ''  end) = ''),0) ,

@Column3 = isnull((SELECT sum(T1.[LineTotal]) FROM OPCH T0  INNER JOIN PCH1 T1 ON T0.[DocEntry] = T1.[DocEntry]
join RPC1 T2 on T2.BaseEntry = T1.DocEntry and T2.BaseLine = T1.LineNum INNER JOIN ORPC T3 ON T2.[DocEntry] = T3.[DocEntry] 
WHERE T1.[LineStatus]  = 'C'
and T0.DocDate between @datef and @DateT
and (case when @Cat = 'PRJ' then isnull(T1.Project,'') else  isnull(T1.U_AB_NONPROJECT ,'')  end) = @Dimension
and isnull(T1.AcctCode ,'') = @GLCode
and (case when @Cat <> 'PRJ' then isnull(T1.Project,'') else  ''  end) = ''),0) ,
  
--- Less AP Credit memo standalone
@Column4 = isnull((SELECT sum(T1.[LineTotal]) FROM ORPC T0  INNER JOIN RPC1 T1 ON T0.[DocEntry] = T1.[DocEntry] WHERE T1.[LineStatus]  = 'O' and [BaseType] = -1
and T0.DocDate between @datef and @DateT
and (case when @Cat = 'PRJ' then isnull(T1.Project,'') else  isnull(T1.U_AB_NONPROJECT ,'')  end) = @Dimension
and isnull(T1.AcctCode ,'') = @GLCode
and (case when @Cat <> 'PRJ' then isnull(T1.Project,'') else  ''  end) = '')
,0)  , 

--- Less AP INVOICE Cancell
@Column5 = isnull((SELECT sum(T1.[LineTotal]) FROM OPCH T0  INNER JOIN PCH1 T1 ON T0.[DocEntry] = T1.[DocEntry]
WHERE t1.LineStatus = 'O' AND T1.BaseType IN ('18')
and  T0.DocDate between @datef and @DateT
and (case when @Cat = 'PRJ' then isnull(T1.Project,'') else  isnull(T1.U_AB_NONPROJECT ,'')  end) = @Dimension
and isnull(T1.AcctCode ,'') = @GLCode
and (case when @Cat <> 'PRJ' then isnull(T1.Project,'') else  ''  end) = '')
,0)  

return ((@Column1 + @Column2  ) - @Column4 )

end

GO
/****** Object:  Table [dbo].[AB_EmailStatus]    Script Date: 22/08/2017 11:41:16 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AB_EmailStatus](
	[Sno] [int] IDENTITY(1,1) NOT NULL,
	[DocType] [nvarchar](5) NOT NULL,
	[ObjectType] [nvarchar](20) NULL,
	[Entity] [nvarchar](200) NULL,
	[EmailID] [nvarchar](max) NULL,
	[EmailBody] [nvarchar](max) NULL,
	[EmailSub] [nvarchar](100) NULL,
	[Status] [nvarchar](10) NULL,
	[ErrMsg] [nvarchar](max) NULL,
	[EmailDate] [datetime] NULL,
	[EmailTime] [nvarchar](30) NULL,
	[Fcount] [nchar](10) NULL,
	[sUser] [nvarchar](30) NULL,
	[Seq] [int] NULL,
	[DraftKey] [nvarchar](50) NULL,
 CONSTRAINT [KAB_EmailStatus] PRIMARY KEY CLUSTERED 
(
	[Sno] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO
