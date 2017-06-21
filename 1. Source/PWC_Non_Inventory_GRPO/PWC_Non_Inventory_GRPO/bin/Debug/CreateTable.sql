USE [SBODemoSG]
GO

/****** Object:  Table [dbo].[AB_GRPO_NON_INV]    Script Date: 23/3/2015 11:56:25 AM ******/
DROP TABLE [dbo].[AB_GRPO_NON_INV]
GO

/****** Object:  Table [dbo].[AB_GRPO_NON_INV]    Script Date: 23/3/2015 11:56:25 AM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[AB_GRPO_NON_INV](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[DocEntry] [int] NOT NULL,
	[LineTotal] [numeric](38, 11) NULL,
	[TotalFrgn] [numeric](38, 11) NULL,
	[DebitAcctCode] [nvarchar](15) NULL,
	[CreditAcctCode] [varchar](15) NOT NULL,
	[Currency] [nvarchar](3) NULL,
	[OcrCode] [nvarchar](8) NOT NULL,
	[OcrCode2] [nvarchar](8) NOT NULL,
	[OcrCode3] [nvarchar](8) NOT NULL,
	[OcrCode4] [nvarchar](8) NOT NULL,
	[Dt_LastMonth] [datetime] NULL,
	[Dt_FitstMonth] [datetime] NULL,
	[SendDate] [datetime] NOT NULL,
	[SysncSt_LastMonth] [varchar](1) NOT NULL,
	[SysncSt_FirstMonth] [varchar](1) NOT NULL,
	[ReceiveDate_LastMonth] [datetime] NULL,
	[ReceiveDate_FitstMonth] [datetime] NULL,
	[ErrorMsg] [varchar](max) NULL,
	[ErrorMsg1] [varchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO


/****** Object:  StoredProcedure [dbo].[GRPO_NON_INV]    Script Date: 23/3/2015 11:57:25 AM ******/
DROP PROCEDURE [dbo].[GRPO_NON_INV]
GO

/****** Object:  StoredProcedure [dbo].[GRPO_NON_INV]    Script Date: 23/3/2015 11:57:25 AM ******/
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
CREATE PROCEDURE [dbo].[GRPO_NON_INV]
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
	Select T0.DocEntry,T0.DocEntry,SUm(T1.LineTotal-(T1.LineTotal * T0.DiscPrcnt/100)) LineTotal,0  TotalFrgn,T1.AcctCode DebitAcctCode, '208040' CreditAcctCode
	 ,'SGD' Currency,
	isnull(T1.OcrCode,'') OcrCode,isnull(T1.OcrCode2,'') OcrCode2,isnull(T1.OcrCode3,'') OcrCode3,isnull(T1.OcrCode4,'') OcrCode4, 
	
	@Dt_LastMonth Dt_LastMonth,@Dt_FitstMonth Dt_FitstMonth,GetDate() 'SendDate','0' SysncSt_LastMonth,'0' SysncSt_FirstMonth,@ReceiveDate_LastMonth ReceiveDate_LastMonth,@ReceiveDate_FitstMonth ReceiveDate_FitstMonth,@ErrMsg 'ErrorMsg',@ErrMsg 'ErrorMsg1'
	
	from OPDN T0 with(nolock)
	Inner Join PDN1 T1 with(nolock) on T0.DocEntry=T1.DocEntry
	Inner Join OITM T2 with(nolock) on T1.ItemCode=T2.ItemCode

	where  isnull(T2.InvntITEM,'N') ='N' and T0.DocDate <= @Dt_LastMonth  and T1.LineStatus='O' and isnull(T2.AssetItem,'N')='N'
	Group By T0.DocEntry,T1.AcctCode,T1.OcrCode,T1.OcrCode2,T1.OcrCode3,T1.OcrCode4;

END

GO


