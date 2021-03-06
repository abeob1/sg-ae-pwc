USE [PWCL_1]
GO
/****** Object:  StoredProcedure [dbo].[AE_SP003_PaidInvoicesList]    Script Date: 29/03/2017 03:47:06 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

--[dbo].[AE_SP003_PaidInvoicesList]'20160802','20160802'


ALTER PROCEDURE [dbo].[AE_SP003_PaidInvoicesList]
-- Add the parameters for the stored procedure here
@FrmDate DateTime,
@ToDate DateTime
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
	
		
