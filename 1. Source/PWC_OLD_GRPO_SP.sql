USE [SBODemoSG]
GO
/****** Object:  StoredProcedure [dbo].[GRPO_NON_INV]    Script Date: 19/3/2015 12:59:52 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		Gopinath
-- Create date: 19-01-2015
-- Description:	PWC_Non Inventory GRPO
-- EXEC GRPO_NON_INV

-- =============================================
ALTER PROCEDURE [dbo].[GRPO_NON_INV]
	-- Add the parameters for the stored procedure here
	
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    declare @Dt_LastMonth as DateTime
	declare @Dt_FitstMonth as DateTime;
	declare @ReceiveDate_LastMonth as DateTime;
	declare @ReceiveDate_FitstMonth as DateTime;
	declare @ErrMsg as varchar(Max);
	Set @Dt_LastMonth=(select top 1 DATEADD(MONTH, DATEDIFF(MONTH, -1, GETDATE())-1, -1));
	set @Dt_FitstMonth =(select top 1 DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()), 0));
	Select T0.DocEntry,T0.DocEntry,SUm(T1.LineTotal-(T1.LineTotal * T0.DiscPrcnt/100)) LineTotal,Sum(T1.TotalFrgn-(T1.TotalFrgn * T0.DiscPrcnt/100)) TotalFrgn,T1.AcctCode DebitAcctCode, '208040' CreditAcctCode,T1.Currency,
	isnull(T1.OcrCode,'') OcrCode,isnull(T1.OcrCode2,'') OcrCode2,isnull(T1.OcrCode3,'') OcrCode3,isnull(T1.OcrCode4,'') OcrCode4, 
	
	@Dt_LastMonth Dt_LastMonth,@Dt_FitstMonth Dt_FitstMonth,GetDate() 'SendDate','0' SysncSt_LastMonth,'0' SysncSt_FirstMonth,@ReceiveDate_LastMonth ReceiveDate_LastMonth,@ReceiveDate_FitstMonth ReceiveDate_FitstMonth,@ErrMsg 'ErrorMsg',@ErrMsg 'ErrorMsg1'
	
	from OPDN T0 with(nolock)
	Inner Join PDN1 T1 with(nolock) on T0.DocEntry=T1.DocEntry
	Inner Join OITM T2 with(nolock) on T1.ItemCode=T2.ItemCode

	where  isnull(T2.InvntITEM,'N') ='N' and T0.DocDate <= @Dt_LastMonth  and T1.LineStatus='O'
	Group By T0.DocEntry,T1.AcctCode,T1.Currency,T1.OcrCode,T1.OcrCode2,T1.OcrCode3,T1.OcrCode4,TotalFrgn;

END
