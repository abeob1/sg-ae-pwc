



CREATE Procedure AE_SP_Emaillog_Statusupdate

as

begin

DECLARE @getEmaillog CURSOR

DECLARE @Sno Varchar(100)

DECLARE @Objecttype Varchar(100)

DECLARE @Draftkey Varchar(100)

DECLARE @sUser Varchar(1000)

DECLARE @Seq integer

DECLARE @rtotalcount integer

Declare @rcount integer



---------------------- Fetching the information from Email log table



select (select max(cast(TT.Seq as integer)) from [AB_EmailStatus] TT where TT.DraftKey = T0.DraftKey ) as Rcount , 

T0.Sno , T0.ObjectType , T0.DraftKey , T0.sUser ,T0.Seq , T0.Status  into #Tmp_Emailstatus from [AB_EmailStatus] T0

delete [#Tmp_Emailstatus] where Rcount = Seq 

--------------------------------------------------------------------

--Select Sno , ObjectType , DraftKey , '%' + sUser + '%' , Seq, Rcount   from [#Tmp_Emailstatus] where status = 'Pending'



--------------------   Cursor Declartion to identify the Draft no is approved for this user

SET @getEmaillog = CURSOR FOR

Select Sno , ObjectType , DraftKey , '%' + sUser + '%' , Seq, Rcount   from [#Tmp_Emailstatus] where status = 'Pending'



OPEN @getEmaillog

FETCH NEXT

FROM @getEmaillog INTO @Sno, @Objecttype, @Draftkey, @sUser, @Seq, @rtotalcount

WHILE @@FETCH_STATUS = 0

BEGIN



-------------------  Getting information from sap table with respective draft key

SELECT TT1.[StepCode], TT0.[DocEntry], TT3.ObjType, Usercode = SUBSTRING((

    SELECT '/' + cast(T2.[USER_CODE]    as varchar) 

    FROM OWDD T0  join WDD1 T1 on T0.WddCode = T1.WddCode

	join ousr T2 on T2.USERID = T1.UserID 

    join odrf T3 on T3.DocEntry = T0.DocEntry

	WHERE T0.[DocEntry] = TT0.[DocEntry] and T1.[StepCode] = TT1.[StepCode]

    for XML PATH ('')), 1,10000) + '/',

	Status = SUBSTRING((

    SELECT '/' + cast(T1.[Status]     as varchar) 

    FROM OWDD T0  join WDD1 T1 on T0.WddCode = T1.WddCode

	join ousr T2 on T2.USERID = T1.UserID 

    join odrf T3 on T3.DocEntry = T0.DocEntry

	WHERE T0.[DocEntry] = TT0.[DocEntry] and T1.[StepCode] = TT1.[StepCode]

    for XML PATH ('')), 1,10000) + '/' INTO #Tmp_OWDD

	FROM OWDD TT0  join WDD1 TT1 on TT0.WddCode = TT1.WddCode 

    join ousr TT2 on TT2.USERID = TT1.UserID 

    join odrf TT3 on TT3.DocEntry = TT0.DocEntry

    WHERE TT0.[DocEntry] = @Draftkey and  TT3.ObjType = @Objecttype

GROUP BY TT1.[StepCode],TT0.[DocEntry],TT3.ObjType

------------------------------------------------------------------------------------------



--------------------------- Identifing the approval status

select @rcount = count(*) from #Tmp_OWDD where DocEntry = @Draftkey and ObjType = @Objecttype and Usercode like  @sUser and Status like '%Y%'





if @rcount > 0

   begin

   

     if @rtotalcount = @seq + 1

	 begin

	 	update [AB_EmailStatus] set Status = 'Open' where [DraftKey] = @Draftkey and [ObjectType] = @Objecttype and [Seq] = @rtotalcount

	 --select 'Approved' , @Draftkey , @sUser , @Seq , @rcount , 'Orignator Status Open'

	 end

     --select 'Approved' , @Draftkey , @sUser 

	 update [AB_EmailStatus] set Status = 'Open' where [Sno] = @Sno

   end

----else

----   begin

----    select 'Not Approved' , @Draftkey , @sUser 

----   end

Drop table [#Tmp_OWDD]

--------------------------------------------------------------------------------------------

FETCH NEXT

FROM @getEmaillog INTO @Sno, @Objecttype, @Draftkey, @sUser,  @Seq,  @rtotalcount

END

CLOSE @getEmaillog

DEALLOCATE @getEmaillog

--------------------------------------------------------------------------------------------

Drop table [#Tmp_Emailstatus] 

end














