--[dbo].[SP_AB_FRM_PurchaseOrder] '7'
alter proc [dbo].[SP_AB_FRM_PurchaseOrder]
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
T1.Dscription,T1.PriceBefDi,T1.Quantity,T1.LineTotal,T1.VatSum,
T1.OcrCode, T1.OcrCode2, T1.OcrCode3, T1.OcrCode4,
T0.[Address],T0.Address2,T0.DocNum,T0.NumAtCard,T1.unitMsr,T1.ShipDate as 'DelDate', TT.[U_NAME] [CreatorName], TT.[USER_CODE] [CreatorCode],

-----------Outlet Details-------------
T10.WhsName,T10.StreetNo,T10.Block, T10.Street,T10.City,T10.ZipCode,T10.FedTaxID,T10.GlblLocNum,
-----------SalesPerson----------------
T0.TaxDate,T6.mobile,T6.Fax as 'Fax2',T6.email,T6.[firstName], T6.[middleName],

-----------BP info--------------
T4.CardName, T5.Name ContactName,T4.CntctPrsn,T0.ShipToCode,T4.E_Mail, T0.Footer, T0.Header,T4.Phone1 as 'PhoneCust', 
T0.CardCode, T0.Comments,T5.Cellolar,T0.DocCur,T4.Fax as BPFax,T8.PymntGroup,T0.U_AB_WTAX
from OPOR T0 with(nolock)
join POR1 T1 with(nolock) on T1.DocEntry=t0.DocEntry
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