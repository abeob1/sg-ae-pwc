ALTER procedure [dbo].[@AE_SP002_InsertintoBudgetTable]

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