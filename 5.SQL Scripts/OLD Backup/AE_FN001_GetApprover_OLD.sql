USE [PWCL_1]
GO
/****** Object:  UserDefinedFunction [dbo].[AE_FN003_GetApprover]    Script Date: 29/03/2017 03:47:54 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

ALTER FUNCTION [dbo].[AE_FN003_GetApprover](@DocEntry varchar(30),@Level varchar(30))
RETURNS varchar(1000)
AS 
-- Returns the stock level for the product.
BEGIN
Declare @Approver varchar(300)
Declare @count integer
declare @TempApprover table (ID INT, Code Varchar(300), Name varchar(500), Date datetime, Time datetime)

Insert into @TempApprover 
SELECT ROW_NUMBER() 
        OVER (ORDER BY T2.UpdateDate,T2.UpdateTime) AS coun, T3.[USER_CODE], T3.[U_NAME],T2.UpdateDate,T2.UpdateTime FROM 
OPOR T0 join OWDD T1 on T0.DocEntry = T1.Docentry inner join 
WDD1 T2 on T1.WddCode = T2.WddCode INNER JOIN OUSR T3 ON T2.[USERID] = T3.[USERID] 
WHERE T0.Docentry = @DocEntry and T2.Status = 'Y' and T1.[ObjType] = '22'
select @count = Count(ID) from @TempApprover
if @level = '1'
begin
 
  if @count = 1
   begin
   select @Approver = Name from @TempApprover where ID = 1
   end
  else if @count = 2
   begin
   select @Approver = Name from @TempApprover where ID = 1
   end
  else if  @count = 3
   begin
   select @Approver = Name from @TempApprover where ID = 1
   end
end
else if @level = '2'
begin
 if @count = 1
   begin
   set @Approver = 'NA'
   end
  else if @count = 2
   begin
   select @Approver = Name from @TempApprover where ID = 2
   end
  else if  @count = 3
   begin
   select @Approver = Name from @TempApprover where ID = 2
   end
end
else if @level = '3'
begin
if @count = 1
   begin
   set @Approver = 'NA'
   end
  else if @count = 2
   begin
    set @Approver = 'NA'
   end
  else if  @count = 3
   begin
   select @Approver = Name from @TempApprover where ID = 3
   end
end

RETURN @Approver 
END;