USE [PWCL_1]
GO
/****** Object:  StoredProcedure [dbo].[AE_SP003_PaidInvoicesList_Header]    Script Date: 29/03/2017 05:44:14 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


-- =============================================
-- Author:		<Umesh Patle>
-- Create date: <11-May-2015,,>
-- Description:	<PaidInvoicesList,,>
-- =============================================
--EXEC [AE_SP003_PaidInvoicesList_Header] 

ALTER PROCEDURE [dbo].[AE_SP003_PaidInvoicesList_Header]

AS
BEGIN

	select top(1)  isnull(T0.PrintHeadr,T0.CompnyName) CompnyName,T0.Phone1,T0.E_Mail ,T1.GlblLocNum, T0.Fax, T0.FreeZoneNo, T0.TaxIdNum, 
	T0.CompnyAddr,T1.Street, T1.StreetNo , T1.Block, T1.Building, T1.ZipCode , T1.City, T1.Country , T3.LogoImage,T1.IntrntAdrs,T0.RevOffice, T2.Name , 1 as LinkID 
	from OADM T0 with(nolock)   
	left outer join ADM1 T1 with(nolock) on 1=1  
	left outer join OADP T3 with(nolock) on 1=1  
	left outer join OCST T2 with(nolock) on T2.Country  =T1.Country  
		

END
