USE [PWCL]
GO
/****** Object:  StoredProcedure [dbo].[BP_Audit_Report]    Script Date: 03/08/2016 11:37:40 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

 --[BP_Audit_Report]'S1000471','S1000471','20160711','20160711'

ALTER PROCEDURE [dbo].[BP_Audit_Report]
		@FromCardcode Nvarchar(40) ,
		@ToCardcode Nvarchar(40),
		@FromDate Datetime,
		@ToDate Datetime


AS
BEGIN
----DECLARE @FromCardcode Nvarchar(40) ,
----		@ToCardcode Nvarchar(40),
----		@FromDate Datetime,
----		@ToDate Datetime

----		SET @FromCardcode = 'S1000471'
----		SET @ToCardcode = 'S1000471'
----		SET @FromDate = '20160711'
----		SET @ToDate = '20160711'

	-- SET NOCOUNT ON added to prevent extra result sets from
		-- interfering with SELECT statements.
	SET NOCOUNT ON;
  Select Row_Number() over (order by Col .name asc) as id,   isnull(des.Descr,Col .name) ColDesc, Col .name into #OCRDTemp
from sys .all_columns Col Inner join sys .tables tbl on Col .object_id = tbl .object_id 
left outer join CUFD des on des.TableID = tbl.name and   Col .name =  'U_'+ des.AliasID  
where tbl .name = 'OCRD' and Col .name Not in ('UpdateDate','CreateDate','LogInstanc','UserSign2','UserSign','Address','ZipCode','AddrType','Block','Building','City','Country',
'County','BillToDef','ZipCode','State1','Address','StreetNo','MailAddrTy','MailBlock','MailCity','MailCountr','MailCounty','ShipToDef','MailZipCod','State2','MailAddres',
'MailStrNo','MailBuildi'
,'AccCritria','AddID','Affiliate','AtcEntry','AutoCalBCG','AutoPost','AvrageLate','BackOrder','Balance','BalanceFC','BalanceSys','BalTrnsfrd'
,'BCACode','BlockDunn','BNKCounter','BoEDiscnt','BoEOnClct','BoEPrsnt','Box1099','Business','CardValid','CDPNum','CertBKeep','CertWHT','chainStore',
'ChecksBal','CollecAuth','CommGrCode','Commission','ConCerti','ConnBP','CpnNo','CrtfcateNO','DataSource','DatevFirst','DdctFileNo','DdctOffice',
'DdctPrcnt','DdgKey','DdtKey','DebPayAcct','Deleted','DocEntry','DpmClear','DpmIntAct','DscntObjct','DscntRel','DunnDate','DunnLevel','ITWTCode',
'KBKCode','LangCode','LetterNum','ListNum','LocMth','MainUsage','MivzExpSts','MltMthNum','MTHCounter','NINum','NTSWebSite','Number','ObjType','OKATO'
,'OKTMO','OpCode347','OprCount','OrderBalFC','OrderBalSy','OrdersBal','OtrCtlAcct','Pager','PartDelivr','PlngGroup','TaxIdIdent','TaxRndRule','ThreshOver'
,'TolrncDays','TpCusPres','TypeOfOp','TypWTReprt','SefazReply','SefazDate','SefazCheck','SCAdjust','RoleTypCod','RelCode','RcpntID')

  Select Row_Number() over (order by Col .name asc) as id,   isnull(des.Descr,Col .name) ColDesc, Col .name into #CRD1Temp
from sys .all_columns Col Inner join sys .tables tbl on Col .object_id = tbl .object_id 
left outer join CUFD des on des.TableID = tbl.name and   Col .name =  'U_'+ des.AliasID  
where tbl .name = 'CRD1' and Col .name Not in ('UpdateDate','CreateDate','LogInstanc','UserSign2','UserSign')

Select Row_Number() over (order by Col .name asc) as id,   isnull(des.Descr,Col .name) ColDesc, Col .name into #ACR2Temp
from sys .all_columns Col Inner join sys .tables tbl on Col .object_id = tbl .object_id 
left outer join CUFD des on des.TableID = tbl.name and   Col .name =  'U_'+ des.AliasID  
where tbl .name = 'ACR2' and Col .name Not in ('UpdateDate','CreateDate','LogInstanc','UserSign2','UserSign')

Select Row_Number() over (order by Col .name asc) as id,   isnull(des.Descr,Col .name) ColDesc, Col .name into #ACPRTemp
from sys .all_columns Col Inner join sys .tables tbl on Col .object_id = tbl .object_id 
left outer join CUFD des on des.TableID = tbl.name and   Col .name =  'U_'+ des.AliasID  
where tbl .name = 'ACPR' and Col .name Not in ('UpdateDate','CreateDate','LogInstanc','UserSign2','UserSign','updateTime','Active','DataSource',
'NFeRcpn','Notes1','Notes2','ObjType','Pager','Password','Position','Profession')

select ACR1 .CardCode ,ACR1 .LineNum ,MAX (ACR1.LogInstanc) as LogInstanc into TempCRD1
from ACR1 where ACR1 .CardCode >= @FromCardcode and ACR1 .CardCode <= @ToCardcode group by ACR1 .CardCode ,ACR1 .LineNum 

select ACRD .CardCode ,MAX (ACRD.LogInstanc) as LogInstanc into TempACRD
from ACRD where ACRD .CardCode >= @FromCardcode and ACRD .CardCode <= @ToCardcode  group by ACRD .CardCode

select ACRD .CardCode ,ACRD.LogInstanc as LogInstanc into TempACR
from ACRD where ACRD .CardCode >= @FromCardcode and ACRD .CardCode <= @ToCardcode  and ACRD.UpdateDate between @FromDate and @ToDate  


Select OCRD .LogInstanc+1 as LogInstanc,OCRD .CreateDate as [Date],OCRD .CardCode as [BP Code],OCRD .CardName as [BP Name],
Case when CardType = 'C' then 'Customer' when CardType = 'S' then 'Vendor' else 'Lead' end [BP Type],
OUSR .U_NAME as [Created By],Convert(Nvarchar(MAX),null) as [Field Name],Convert (Nvarchar(MAX),'***Created***') as [Old Value],Convert (Nvarchar(MAX),null) as [New Value] into TempOCRD
from OCRD inner join OUSR on OCRD .UserSign = OUSR .INTERNAL_K 
where OCRD .CardCode >= @FromCardcode and  OCRD .CardCode <= @ToCardcode and 
Convert(Nvarchar(8), OCRD .CreateDate ,112) >= Convert(Nvarchar(8),  @FromDate ,112) and Convert(Nvarchar(8), OCRD .CreateDate ,112) <= Convert(Nvarchar(8),  @ToDate ,112);


Select OCRD .LogInstanc+1 as Seq ,OCRD .LogInstanc+1 as LogInstanc,OCRD .CreateDate as [Date],OCRD .CardCode as [BP Code],OCRD .CardName as [BP Name],
Case when CardType = 'C' then 'Customer' when CardType = 'S' then 'Vendor' else 'Lead' end [BP Type],
OUSR .U_NAME as [Created By],Convert(Nvarchar(MAX),null) as [Field Name],Convert (Nvarchar(MAX),'***Created***') as [Old Value],Convert (Nvarchar(MAX),null) as [New Value] into TempARPC
from OCRD inner join OUSR on OCRD .UserSign = OUSR .INTERNAL_K 
where OCRD .CardCode >= @FromCardcode and  OCRD .CardCode <= @ToCardcode and 
Convert(Nvarchar(8), OCRD .CreateDate ,112) >= Convert(Nvarchar(8),  @FromDate ,112) and Convert(Nvarchar(8), OCRD .CreateDate ,112) <= Convert(Nvarchar(8),  @ToDate ,112);

Declare @Count int,@Counter int,@ColName Nvarchar (100), @Query Nvarchar(max), @ColDesc Nvarchar (100), @Tmpcal NVarchar(10), @ColumnOLD NVarchar(200), @ColumnNEW NVarchar(200) ;

Set @Count = (Select MAX (#OCRDTemp .id) from #OCRDTemp);
Set @Counter = 1;

While (@Counter <= @Count)

Begin

	Set @Query = '';

	Set @ColName = (Select Name from #OCRDTemp  where id = @Counter);
	Set @ColDesc = (Select ColDesc from #OCRDTemp  where id = @Counter)

	Set @Query = 'Insert into TempOCRD Select T0 .LogInstanc+1,T1 .UpdateDate  as [Date],T0 .CardCode as [BP Code],T0 .CardName as [BP Name],
	Case when T0 .CardType = ''C'' then ''Customer'' when T0 .CardType = ''S'' then ''Vendor'' else ''Lead'' end [BP Type],
	OUSR .U_NAME as [Created By],'''+ @ColDesc +''' as [Field Name],convert(Nvarchar(MAX),T0 .'+@ColName+',106) as [Old Value],Convert(Nvarchar(MAX),T1 .'+@ColName+',106) as [New Value]
	from ACRD T0 inner join ACRD T1 on T0 .CardCode = T1 .CardCode and T0 .LogInstanc = T1 .LogInstanc -1
	left outer join OUSR on T1 .UserSign2 = OUSR .INTERNAL_K 
	where  T0 .CardCode >= '''+@FromCardcode+''' and  T0 .CardCode <= '''+@ToCardcode+''' and 
	Convert(Nvarchar(8), T1 .UpdateDate ,112)  >= '''+Convert(Nvarchar(8),  @FromDate ,112)+''' AND 
	Convert(Nvarchar(8), T1 .UpdateDate ,112)  <= '''+Convert(Nvarchar(8),  @ToDate ,112)+''' AND 
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
Convert(Nvarchar(8), T1 .UpdateDate ,112)  >= '''+Convert(Nvarchar(8),  @FromDate ,112)+''' AND 
Convert(Nvarchar(8), T1 .UpdateDate ,112)  <= '''+Convert(Nvarchar(8),  @ToDate ,112)+''' AND 
isnull(convert(Nvarchar(MAX),T0 .'+@ColName+',106) ,'''') <> isnull(convert(Nvarchar(MAX),T1 .'+@ColName+',106),'''')';
--print @query
Exec(@Query)

Set @Counter = @Counter + 1;
End 


Set @Count = (Select MAX (#CRD1Temp .id) from #CRD1Temp);
Set @Counter = 1;

While (@Counter <= @Count)

Begin
	Set @Query = '';
	Set @ColName = (Select Name from #CRD1Temp  where id = @Counter);
	Set @ColDesc = (Select ColDesc from #CRD1Temp  where id = @Counter)

	Set @Query = 'Insert into TempOCRD Select ACRD .LogInstanc,ACRD .UpdateDate  as [Date],ACRD .CardCode as [BP Code],ACRD .CardName as [BP Name],
Case when ACRD .CardType = ''C'' then ''Customer'' when ACRD .CardType = ''S'' then ''Vendor'' else ''Lead'' end [BP Type],
OUSR .U_NAME as [Created By],T1.Address+''-''+'''+ @ColDesc +''' as [Field Name],convert(Nvarchar(MAX),T0 .'+@ColName+',106) as [Old Value],Convert(Nvarchar(MAX),T1 .'+@ColName+',106) as [New Value]
from ACRD left outer join OUSR on ACRD .UserSign2 = OUSR .INTERNAL_K 
inner join ACR1 T0 on ACRD .Cardcode = T0 . Cardcode 
inner join ACR1 T1 on T0 .CardCode = T1 .CardCode and T0.Linenum = T1.Linenum and T0 .LogInstanc = T1 .LogInstanc -1 and ACRD .LogInstanc = T1.LogInstanc
where  T0 .CardCode >= '''+@FromCardcode+''' and  T0 .CardCode <= '''+@ToCardcode+''' and 
Convert(Nvarchar(8), ACRD .UpdateDate ,112)  >= '''+Convert(Nvarchar(8),  @FromDate ,112)+''' AND 
Convert(Nvarchar(8), ACRD .UpdateDate ,112)  <= '''+Convert(Nvarchar(8),  @ToDate ,112)+''' AND 
isnull(convert(Nvarchar(MAX),T0 .'+@ColName+',106) ,'''') <> isnull(convert(Nvarchar(MAX),T1 .'+@ColName+',106),'''')';

Exec(@Query)

	Set @Query = 'Insert into TempOCRD Select t1 .LogInstanc+1 ,OCRD .UpdateDate  as [Date],OCRD .CardCode as [BP Code],OCRD .CardName as [BP Name],
Case when OCRD .CardType = ''C'' then ''Customer'' when OCRD .CardType = ''S'' then ''Vendor'' else ''Lead'' end [BP Type],
OUSR .U_NAME as [Created By],T0.Address+''-''+'''+ @ColDesc +''' as [Field Name],convert(Nvarchar(MAX),T1 .'+@ColName+',106) as [Old Value],Convert(Nvarchar(MAX),T0 .'+@ColName+',106) as [New Value]
from OCRD Left outer join OUSR on OCRD .UserSign2 = OUSR .INTERNAL_K 
inner join CRD1 T0 on OCRD .Cardcode = T0 . Cardcode 
inner join ACR1 T1 on T0 .CardCode = T1 .CardCode and T0.Linenum = T1.Linenum 
inner join TempCrd1 T2 on OCRD .Cardcode = T2 .Cardcode and T1 .LogInstanc = T2 .LogInstanc and T0 .Linenum = T2.LineNum
where T0 .CardCode >= '''+@FromCardcode+''' and  T0 .CardCode <= '''+@ToCardcode+''' and 
Convert(Nvarchar(8), OCRD .UpdateDate ,112)  >= '''+Convert(Nvarchar(8),  @FromDate ,112)+''' AND 
Convert(Nvarchar(8), OCRD .UpdateDate ,112)  <= '''+Convert(Nvarchar(8),  @ToDate ,112)+''' AND 
isnull(convert(Nvarchar(MAX),T0 .'+@ColName+',106) ,'''') <> isnull(convert(Nvarchar(MAX),T1 .'+@ColName+',106),'''')';

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
(select PymCode from ACR2 TT where TT.CardCode = ACRD.CardCode and  TT.LogInstanc = ACRD .LogInstanc -1 and TT.LineNum = 0) as [Old Value], 
(select PymCode from ACR2 TT where TT.CardCode = ACRD.CardCode and  TT.LogInstanc = ACRD.LogInstanc and TT.LineNum = 0) as [New Value]
from ACRD left outer join OUSR on ACRD .UserSign2 = OUSR .INTERNAL_K 
where ACRD .CardCode >= '''+@FromCardcode+''' and  ACRD .CardCode <= '''+@ToCardcode+''' and
Convert(Nvarchar(8), ACRD .UpdateDate ,112)  >= '''+Convert(Nvarchar(8),  @FromDate ,112)+''' AND 
Convert(Nvarchar(8), ACRD .UpdateDate ,112)  <= '''+Convert(Nvarchar(8),  @ToDate ,112)+''' and ACRD .LogInstanc = '''+@Tmpcal+''''

exec(@Query)
set @Query = 'Insert into TempOCRD Select ACRD .LogInstanc,ACRD .UpdateDate  as [Date],ACRD .CardCode as [BP Code],ACRD .CardName as [BP Name],
Case when ACRD .CardType = ''C'' then ''Customer'' when ACRD .CardType = ''S'' then ''Vendor'' else ''Lead'' end [BP Type],
OUSR .U_NAME as [Created By],''PymCode'' as [Field Name],
(select PymCode from ACR2 TT where TT.CardCode = ACRD.CardCode and  TT.LogInstanc = ACRD .LogInstanc -1 and TT.LineNum = 1) as [Old Value], 
(select PymCode from ACR2 TT where TT.CardCode = ACRD.CardCode and  TT.LogInstanc = ACRD.LogInstanc and TT.LineNum = 1) as [New Value]
from ACRD left outer join OUSR on ACRD .UserSign2 = OUSR .INTERNAL_K 
where ACRD .CardCode >= '''+@FromCardcode+''' and  ACRD .CardCode <= '''+@ToCardcode+''' and
Convert(Nvarchar(8), ACRD .UpdateDate ,112)  >= '''+Convert(Nvarchar(8),  @FromDate ,112)+''' AND 
Convert(Nvarchar(8), ACRD .UpdateDate ,112)  <= '''+Convert(Nvarchar(8),  @ToDate ,112)+''' and ACRD .LogInstanc = '''+@Tmpcal+''''

exec(@Query)

set @Query = 'Insert into TempOCRD Select ACRD .LogInstanc,ACRD .UpdateDate  as [Date],ACRD .CardCode as [BP Code],ACRD .CardName as [BP Name],
Case when ACRD .CardType = ''C'' then ''Customer'' when ACRD .CardType = ''S'' then ''Vendor'' else ''Lead'' end [BP Type],
OUSR .U_NAME as [Created By],''PymCode'' as [Field Name],
(select PymCode from ACR2 TT where TT.CardCode = ACRD.CardCode and  TT.LogInstanc = ACRD .LogInstanc -1 and TT.LineNum = 2) as [Old Value], 
(select PymCode from ACR2 TT where TT.CardCode = ACRD.CardCode and  TT.LogInstanc = ACRD.LogInstanc and TT.LineNum = 2) as [New Value]
from ACRD left outer join OUSR on ACRD .UserSign2 = OUSR .INTERNAL_K 
where ACRD .CardCode >= '''+@FromCardcode+''' and  ACRD .CardCode <= '''+@ToCardcode+''' and
Convert(Nvarchar(8), ACRD .UpdateDate ,112)  >= '''+Convert(Nvarchar(8),  @FromDate ,112)+''' AND 
Convert(Nvarchar(8), ACRD .UpdateDate ,112)  <= '''+Convert(Nvarchar(8),  @ToDate ,112)+''' and ACRD .LogInstanc = '''+@Tmpcal+''''

exec(@Query)

set @Query = 'Insert into TempOCRD Select ACRD .LogInstanc,ACRD .UpdateDate  as [Date],ACRD .CardCode as [BP Code],ACRD .CardName as [BP Name],
Case when ACRD .CardType = ''C'' then ''Customer'' when ACRD .CardType = ''S'' then ''Vendor'' else ''Lead'' end [BP Type],
OUSR .U_NAME as [Created By],''PymCode'' as [Field Name],
(select PymCode from ACR2 TT where TT.CardCode = ACRD.CardCode and  TT.LogInstanc = ACRD .LogInstanc -1 and TT.LineNum = 3) as [Old Value], 
(select PymCode from ACR2 TT where TT.CardCode = ACRD.CardCode and  TT.LogInstanc = ACRD.LogInstanc and TT.LineNum = 3) as [New Value]
from ACRD left outer join OUSR on ACRD .UserSign2 = OUSR .INTERNAL_K 
where ACRD .CardCode >= '''+@FromCardcode+''' and  ACRD .CardCode <= '''+@ToCardcode+''' and
Convert(Nvarchar(8), ACRD .UpdateDate ,112)  >= '''+Convert(Nvarchar(8),  @FromDate ,112)+''' AND 
Convert(Nvarchar(8), ACRD .UpdateDate ,112)  <= '''+Convert(Nvarchar(8),  @ToDate ,112)+''' and ACRD .LogInstanc = '''+@Tmpcal+''''

exec(@Query)
Set @Counter = @Counter + 1;
End 

Set @Count = (Select max (TempACR.LogInstanc ) from TempACR);
Set @Counter = (Select min (TempACR.LogInstanc ) from TempACR);

While (@Counter <= @Count)

Begin
	Set @Query = ''
	set @Tmpcal = @Counter
	
	if @Counter < @Count 
	begin
	 set @ColumnOLD = 'select AcctName from acrb TT where TT.CardCode = ACRD.CardCode and  TT.LogInstanc = ACRD .LogInstanc -1'
	 set @ColumnNEW = 'select AcctName from acrb TT where TT.CardCode = ACRD.CardCode and  TT.LogInstanc = ACRD.LogInstanc'
	end
	if @Counter = @Count 
	begin
	 set @ColumnOLD = 'select AcctName from acrb TT where TT.CardCode = ACRD.CardCode and  TT.LogInstanc = ACRD .LogInstanc'
	 set @ColumnNEW = 'select AcctName from ocrb TT where TT.CardCode = ACRD.CardCode'
	end

set @Query = 'Insert into TempOCRD Select ACRD .LogInstanc,ACRD .UpdateDate  as [Date],ACRD .CardCode as [BP Code],ACRD .CardName as [BP Name],
	Case when ACRD .CardType = ''C'' then ''Customer'' when ACRD .CardType = ''S'' then ''Vendor'' else ''Lead'' end [BP Type],
	OUSR .U_NAME as [Created By],''AcctName'' as [Field Name],
	('+ @ColumnOLD +' ) as [Old Value], 
	('+ @ColumnNEW +' ) as [New Value]
	from ACRD left outer join OUSR on ACRD .UserSign2 = OUSR .INTERNAL_K 
	where ACRD .CardCode >= '''+@FromCardcode+''' and  ACRD .CardCode <= '''+@ToCardcode+''' and
	Convert(Nvarchar(8), ACRD .UpdateDate ,112)  >= '''+Convert(Nvarchar(8),  @FromDate ,112)+''' AND 
	Convert(Nvarchar(8), ACRD .UpdateDate ,112)  <= '''+Convert(Nvarchar(8),  @ToDate ,112)+''' and ACRD .LogInstanc = '''+@Tmpcal+'''
	'
exec(@Query)
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


SET @Query = '; WITH T1 AS(SELECT DISTINCT T3 .LogInstanc,T1 .UpdateDate  as [Date],T0 .CardCode as [BP Code],T3 .CardName as [BP Name],
Case when T3 .CardType = ''C'' then ''Customer'' when T3 .CardType = ''S'' then ''Vendor'' else ''Lead'' end [BP Type],
T4 .U_NAME as [Created By],'''+ @ColDesc +''' as [Field Name],convert(Nvarchar(MAX),T1 .'+@ColName+',106) AS [Old Value],
case when 
(select max(LogInstanc) from ACPR TT where TT.CardCode = T0.CardCode  and TT.CntctCode = T1.CntctCode and TT.updateDate = T1.updateDate) 
 = T1.LogInstanc then Convert(Nvarchar(MAX),T0 .'+@ColName+',106) else Convert(Nvarchar(MAX),T2.'+@ColName+',106)
 end 
  AS [New Value]
FROM  OCPR T0
LEFT JOIN ACPR T1 ON T1.CntctCode = T0.CntctCode
LEFT JOIN ACPR T2 ON T2.CntctCode = T1.CntctCode and T1.LogInstanc = T2.LogInstanc-1
LEFT JOIN ACRD T3 ON T3.CardCode = T1.CardCode 
LEFT JOIN OUSR T4 on T3 .UserSign2 = T4 .INTERNAL_K
WHERE T1.CntctCode = T0.CntctCode AND T1. '+@ColName+' <> T0. '+@ColName+'
AND T3 .CardCode >= '''+@FromCardcode+''' and  T3 .CardCode <= '''+@ToCardcode+''' and 
Convert(Nvarchar(8), T3 .UpdateDate ,112)  >= '''+Convert(Nvarchar(8),  @FromDate ,112)+''' AND 
Convert(Nvarchar(8), T3 .UpdateDate ,112)  <= '''+Convert(Nvarchar(8),  @ToDate ,112)+''' AND 
isnull(convert(Nvarchar(MAX),T0 .'+@ColName+',106) ,'''') <> isnull(convert(Nvarchar(MAX),T1 .'+@ColName+',106),'''')
--AND MAX(T3.LogInstanc) = MAX(T3.LogInstanc)
group by T1.CntctCode, T0.CardCode,T3.LogInstanc,T1.LogInstanc,T0.updateDate,
T1.updateDate,T3 .CardName,T3 .CardType,T4 .U_NAME , T1. '+@ColName+', T0. '+@ColName+', T2.'+@ColName+'
)
Insert into TempOCRD
SELECT * FROM T1 WHERE LogInstanc = (SELECT MAX(LogInstanc) FROM T1)'

--print @query

Exec(@Query)

Set @Counter = @Counter + 1;
End 

Delete from TempOCRD where [Old Value] = '-1' and [New Value] = '-1'
Delete from TempOCRD where ISNULL([Old Value],'') = ISNULL([New Value],'')

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
when 'Cellular' then 'Mobile Phone'
when 'CntctPrsn' then 'Contact ID'
when 'DflAccount' then 'Account'
when 'DflBranch' then 'Branch'
when 'DflIBAN' then 'IBAN/ABA'
when 'DflSwift' then 'BIC/SWIFT Code'
when 'DflBankKey' then 'Default Bank Key'
when 'BankCtlKey' then 'Ctrl Int. ID'
when 'GroupNum' then 'Payment Terms Code'
else T0.[Field Name]
end) [Field Name]
 , 
 (case 
   when T0.[Field Name] = 'SlpCode' then (SELECT TT0.[SlpName] FROM OSLP TT0 where TT0.[SlpCode] = T0.[Old Value]) 
   when T0.[Field Name] = 'GroupNum' then (SELECT TT0.[PymntGroup] FROM OCTG TT0 where TT0.[GroupNum] = T0.[Old Value])
   else T0.[Old Value] end) [Old Value],
(case 
     when T0.[Field Name] = 'SlpCode' then (SELECT TT0.[SlpName] FROM OSLP TT0 where TT0.[SlpCode] = T0.[New Value]) 
	 when T0.[Field Name] = 'GroupNum' then (SELECT TT0.[PymntGroup] FROM OCTG TT0 where TT0.[GroupNum] = T0.[New Value])
	 else T0.[New Value] end) [New Value]
into TempOCRDF
from TempOCRD T0 
order by T0.[BP Code], T0.LogInstanc   asc 

select * from TempOCRDF T0 
where T0.Date between @FromDate and @ToDate 
group by T0.LogInstanc , T0.Date , T0.[BP Code],T0.[BP Name]  ,T0.[BP Type] ,T0.[Created By], T0.[Field Name],
T0.[Old Value] , T0.[New Value]    

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