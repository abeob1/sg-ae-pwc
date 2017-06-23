USE [PWCL]
GO

/****** Object:  UserDefinedFunction [dbo].[AE_FN003_GetApprover]    Script Date: 11/5/2017 5:19:53 PM ******/
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


