

ALTER Procedure [dbo].[AE_SP_Emaillog_Statusupdate]

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