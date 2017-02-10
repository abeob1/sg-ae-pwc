--EXEC AE_APPROVALGRID 'ALL','ADM3','Goh Li Ling','','ALL','','','2016-11-09','2016-11-09',''

----EXEC AE_APPROVALGRID 'ALL','GTS2','GTS2','','ALL','','','2015-01-01','2016-11-04','GHRS'
--alter PROCEDURE [dbo].[AE_APPROVALGRID](@APPTYPE NVARCHAR(100),@APPCODE NVARCHAR(100),@APPROVER NVARCHAR(100),@CREATOR NVARCHAR(100),@DOCUMENT NVARCHAR(10),@DRAFTNO NVARCHAR(10),@VENDOR NVARCHAR(100),@FROMDATE DATE,@TODATE DATE,@ENTITY NVARCHAR(100))
--AS
BEGIN

DECLARE @APPTYPE NVARCHAR(100),@APPCODE NVARCHAR(100),@APPROVER NVARCHAR(100),@CREATOR NVARCHAR(100),@DOCUMENT NVARCHAR(10),@DRAFTNO NVARCHAR(10),@VENDOR NVARCHAR(100),@FROMDATE DATE,@TODATE DATE,@ENTITY NVARCHAR(100)
SET @APPTYPE = 'ALL'
SET @APPCODE = 'ADM5'
SET @APPROVER = 'Tan Wan Chin'  
 SET @CREATOR = ''
SET @DOCUMENT = 'ALL'
SET @DRAFTNO = ''
SET @VENDOR = ''
SET @FROMDATE = '2016-11-09'
SET @TODATE = '2016-11-09'
SET @ENTITY = ''

DECLARE @CONDITION NVARCHAR(2000)

IF @APPTYPE = 'MAIN'
BEGIN
SET @CONDITION = 'and (isnull(U_Appr1,'''') = '''+ @APPCODE +''' or isnull(U_Appr2,'''') = '''+ @APPCODE +''' or isnull(U_Appr3,'''') = '''+ @APPCODE +''')'
END
ELSE IF @APPTYPE = 'BACKUP'
BEGIN
SET @CONDITION = 'and (isnull(U_Appr1B,'''') = '''+ @APPCODE +''' or isnull(U_Appr2B,'''') = '''+ @APPCODE +''' or isnull(U_Appr3B,'''') = '''+ @APPCODE +''')'
END
ELSE  
BEGIN
SET @CONDITION = 'and (isnull(U_Appr1,'''') = '''+ @APPCODE +''' or isnull(U_Appr2,'''') = '''+ @APPCODE +''' or isnull(U_Appr3,'''') = '''+ @APPCODE +''' or isnull(U_Appr1B,'''') = '''+ @APPCODE +''' or isnull(U_Appr2B,'''') = '''+ @APPCODE +''' or isnull(U_Appr3B,'''') = '''+ @APPCODE +''')'
END

DECLARE @SQL NVARCHAR(MAX),@SQLVAR NVARCHAR(MAX)

CREATE TABLE #TABLE(ENTITY NVARCHAR(50),[DOCUMENT TYPE] NVARCHAR(100),[SELECT] NVARCHAR(2),[DRAFT NO] INTEGER,[DOCUMENT NO] NVARCHAR(50),
				    [CREATOR NAME] NVARCHAR(100),[POSTING DATE] DATE,[VENDOR NAME] NVARCHAR(100),[AMOUNT(BEFORE GST)] NUMERIC(18,3),[APPROVAL GRID] NVARCHAR(100),
					[APPROVALCODE] NVARCHAR(100),[USERCODE] NVARCHAR(100),[STATUS] NVARCHAR(100),REMARKS NVARCHAR(240),APPROVER NVARCHAR(100), 
					[REMARKS BY APPROVER] NVARCHAR(240),APPSEQ INTEGER,SEQ INTEGER,STEPCODE INTEGER)


CREATE TABLE #TABLEF(ENTITY NVARCHAR(50),[DOCUMENT TYPE] NVARCHAR(100),[SELECT] NVARCHAR(2),[DRAFT NO] INTEGER,[DOCUMENT NO] NVARCHAR(50),
				    [CREATOR NAME] NVARCHAR(100),[POSTING DATE] DATE,[VENDOR NAME] NVARCHAR(100),[AMOUNT(BEFORE GST)] NUMERIC(18,3),[APPROVAL GRID] NVARCHAR(100),
					[APPROVALCODE] NVARCHAR(100),[USERCODE] NVARCHAR(100),[STATUS] NVARCHAR(100),REMARKS NVARCHAR(240),APPROVER NVARCHAR(100), 
					[REMARKS BY APPROVER] NVARCHAR(240),APPSEQ INTEGER,SEQ INTEGER,STEPCODE INTEGER)

IF ISNULL(@ENTITY, '') <> ''
BEGIN
SET @SQL = 'SELECT A.DocEntry,CASE WHEN A.ObjType=''22'' THEN ''PURCHASE ORDER'' WHEN A.ObjType=''1470000113'' THEN ''PURCHASE REQUEST'' END[Document Type],
A.ObjType,A.DocEntry[Draft No],F.SeriesName+'' - ''+cast(A.DocNum as nvarchar(30))[Document No],B.U_NAME[Creator Name],A.DocDate[Posting Date],CardName[Vendor Name],A.Doctotal-A.VatSum[DocTotal],A.U_AB_APPROVALCODE,c.WddCode,E.U_NAME,E.USERID,
E.USER_CODE,''Pending''[Status],A.Comments[Remarks],D.Status[AStatus],ROW_NUMBER() OVER(partition by c.WddCode order by c.WddCode) AS AppSeq,
cast( cast(c.WddCode as nvarchar(20)) + cast(D.[StepCode] as nvarchar(20)) as integer)[StepCode]
INTO #TmpInitialise FROM '+ @Entity +'..ODRF A
INNER JOIN '+ @Entity +'..OUSR B ON B.USERID = A.UserSign LEFT JOIN ' + @Entity + '..OWDD C ON A.DocEntry = C.DocEntry 
JOIN '+ @Entity +'..WDD1 D ON C.WddCode = D.WddCode JOIN ' + @Entity + '..OUSR E ON E.USERID = D.UserID 
JOIN '+ @Entity +'..NNM1 F ON F.Series = A.Series LEFT OUTER JOIN PWCL ..[@AB_APPROVALMATRIX] G ON G.U_ApprGridCode  =  A.U_AB_APPROVALCODE 
WHERE DocStatus = ''O'' AND A.DocDate >= '''+ CAST(@FROMDATE AS VARCHAR(50))+''' AND A.DocDate <= '''+ CAST(@TODATE AS VARCHAR(50))+'''
AND A.ObjType IN  (22,1470000113) ' + ISNULL(@CONDITION,'') + '

SELECT ''' + @ENTITY + '''[Entity],[Document Type],''''[Select],[Draft No],[Document No],[Creator Name],[Posting Date],[Vendor Name],DocTotal,U_AB_APPROVALCODE,WddCode,U_NAME,
USERID,USER_CODE,Status,Remarks,[AStatus],AppSeq,StepCode into #TmpApproval FROM #TmpInitialise
WHERE ObjType = (CASE WHEN ''' + @DOCUMENT + ''' = ''ALL'' THEN ObjType WHEN ''' + @DOCUMENT + ''' = ''PR'' THEN ''1470000113'' WHEN ''' + @DOCUMENT + ''' = ''PO'' THEN ''22'' END) 
AND [Creator Name] = (CASE WHEN ISNULL(''' + @CREATOR + ''','''') = '''' THEN  [Creator Name] ELSE ''' + @CREATOR + ''' END)
AND DocEntry = (CASE WHEN ISNULL(''' + @DRAFTNO + ''','''') = '''' THEN DocEntry ELSE ''' + @DRAFTNO + ''' END)
AND ISNULL([Vendor Name],'''') = (CASE WHEN ISNULL(''' + @VENDOR + ''','''') = '''' THEN ISNULL([Vendor Name],'''') ELSE ''' + @VENDOR + ''' END)

select *, dense_rank() OVER (PARTITION BY [draft no] ORDER BY stepcode ) [seq] into #TmpApprovalF from #TmpApproval group by [Entity],[Select],[Document Type],[Draft No],[Document No],[Creator Name],[Posting Date],[Vendor Name],DocTotal,
U_AB_APPROVALCODE,WddCode,U_NAME,USERID,USER_CODE,Status ,Remarks ,[AStatus],AppSeq,StepCode order by [Draft No] , AppSeq 

select * into #TmpApprovalF1 from #TmpApprovalF

--delete from #TmpApprovalF where StepCode in (select T0.StepCode from #TmpApprovalF T0 where T0.AStatus in (''Y'',''N''))

INSERT INTO #TABLE
select Entity,[Document Type],[Select],[Draft No],[Document No],[Creator Name],[Posting Date],[Vendor Name],DocTotal,U_AB_APPROVALCODE,
WddCode,U_NAME,[Status],[Remarks],
(select top 1 U_Name from #TmpApprovalF1 T0 where T0.[Seq]= T1.[Seq] -1 and T0.[Draft No] = T1.[Draft No]  and T0.AStatus = ''Y'' group by U_Name) as [Approvedby] ,
'''' [ApproverRemarks],AppSeq,Seq,StepCode
from #TmpApprovalF T1 where AStatus=''W'' AND U_NAME=''' + @APPROVER + ''' 
order by [Draft No],AppSeq '

set @SQLVAR ='
INSERT INTO #TABLEF
select *  from(
select * from #TABLE WHERE SEQ=1 AND [DRAFT NO] NOT IN (select [DRAFT NO] from #TABLE WHERE SEQ=1 AND [Status]= ''Y'') UNION ALL
select * from #TABLE WHERE SEQ=2 AND [DRAFT NO] IN (select [DRAFT NO] from #TABLE WHERE SEQ=1 AND [Status]=''Y'') UNION ALL
select * from #TABLE WHERE SEQ=3 AND [DRAFT NO] IN (select [DRAFT NO] from #TABLE WHERE SEQ=2 AND [Status]=''Y'') UNION ALL
select * from #TABLE WHERE SEQ=4 AND [DRAFT NO] IN (select [DRAFT NO] from #TABLE WHERE SEQ=3 AND [Status]=''Y'') UNION ALL
select * from #TABLE WHERE SEQ=5 AND [DRAFT NO] IN (select [DRAFT NO] from #TABLE WHERE SEQ=4 AND [Status]=''Y'') UNION ALL
select * from #TABLE WHERE SEQ=6 AND [DRAFT NO] IN (select [DRAFT NO] from #TABLE WHERE SEQ=5 AND [Status]=''Y'') UNION ALL
select * from #TABLE WHERE SEQ=7 AND [DRAFT NO] IN (select [DRAFT NO] from #TABLE WHERE SEQ=6 AND [Status]=''Y'') UNION ALL
select * from #TABLE WHERE SEQ=8 AND [DRAFT NO] IN (select [DRAFT NO] from #TABLE WHERE SEQ=7 AND [Status]=''Y'') UNION ALL
select * from #TABLE WHERE SEQ=9 AND [DRAFT NO] IN (select [DRAFT NO] from #TABLE WHERE SEQ=8 AND [Status]=''Y'')) x order by x.[Draft No],x.[AppSeq]

delete from #TABLEF where StepCode in (select T0.StepCode from #TABLEF T0 where T0.[Status] in (''Y'',''N''))

Drop table #TmpApprovalF1
Drop table #TmpApprovalF
Drop table #TmpApproval
Drop table #TmpInitialise'
  
EXEC(@SQL + @SQLVAR)
END
ELSE IF ISNULL(@ENTITY,'') = ''
BEGIN

DECLARE @ENTITYNAME NVARCHAR(MAX)

DECLARE C1 CURSOR FOR
SELECT A.Name FROM PWCL ..[@AB_COMPANYDATA] A
INNER JOIN [SBO-COMMON].dbo.SRGC B ON B.dbName = A.Name
WHERE A.NAME IN ('PWCL','SEAC')
OPEN C1;
FETCH NEXT FROM C1 INTO @ENTITYNAME
WHILE @@FETCH_STATUS = 0
BEGIN

SET @SQL = 'SELECT A.DocEntry,CASE WHEN A.ObjType=''22'' THEN ''PURCHASE ORDER'' WHEN A.ObjType=''1470000113'' THEN ''PURCHASE REQUEST'' END[Document Type],
A.ObjType,A.DocEntry[Draft No],F.SeriesName+'' - ''+cast(A.DocNum as nvarchar(30))[Document No],B.U_NAME[Creator Name],A.DocDate[Posting Date],CardName[Vendor Name],A.Doctotal-A.VatSum[DocTotal],A.U_AB_APPROVALCODE,c.WddCode,E.U_NAME,E.USERID,
E.USER_CODE,''Pending''[Status],A.Comments[Remarks],D.Status[AStatus],ROW_NUMBER() OVER(partition by c.WddCode order by c.WddCode) AS AppSeq,
cast( cast(c.WddCode as nvarchar(20)) + cast(D.[StepCode] as nvarchar(20)) as integer)[StepCode]
INTO #TmpInitialise FROM '+ @ENTITYNAME +'..ODRF A
INNER JOIN '+ @ENTITYNAME +'..OUSR B ON B.USERID = A.UserSign LEFT JOIN ' + @ENTITYNAME + '..OWDD C ON A.DocEntry = C.DocEntry 
JOIN '+ @ENTITYNAME +'..WDD1 D ON C.WddCode = D.WddCode JOIN ' + @ENTITYNAME + '..OUSR E ON E.USERID = D.UserID 
JOIN '+ @ENTITYNAME +'..NNM1 F ON F.Series = A.Series LEFT OUTER JOIN PWCL ..[@AB_APPROVALMATRIX] G ON G.U_ApprGridCode  =  A.U_AB_APPROVALCODE 
WHERE DocStatus = ''O'' AND A.DocDate >= '''+ CAST(@FROMDATE AS VARCHAR(50))+''' AND A.DocDate <= '''+ CAST(@TODATE AS VARCHAR(50))+'''
AND A.ObjType IN  (22,1470000113) ' + ISNULL(@CONDITION,'') + '

SELECT ''' + @ENTITYNAME + '''[Entity],[Document Type],''''[Select],[Draft No],[Document No],[Creator Name],[Posting Date],[Vendor Name],DocTotal,U_AB_APPROVALCODE,WddCode,U_NAME,
USERID,USER_CODE,Status,Remarks,[AStatus],AppSeq,StepCode into #TmpApproval FROM #TmpInitialise
WHERE ObjType = (CASE WHEN ''' + @DOCUMENT + ''' = ''ALL'' THEN ObjType WHEN ''' + @DOCUMENT + ''' = ''PR'' THEN ''1470000113'' WHEN ''' + @DOCUMENT + ''' = ''PO'' THEN ''22'' END) 
AND [Creator Name] = (CASE WHEN ISNULL(''' + @CREATOR + ''','''') = '''' THEN  [Creator Name] ELSE ''' + @CREATOR + ''' END)
AND DocEntry = (CASE WHEN ISNULL(''' + @DRAFTNO + ''','''') = '''' THEN DocEntry ELSE ''' + @DRAFTNO + ''' END)
AND ISNULL([Vendor Name],'''') = (CASE WHEN ISNULL(''' + @VENDOR + ''','''') = '''' THEN ISNULL([Vendor Name],'''') ELSE ''' + @VENDOR + ''' END)

select *, dense_rank() OVER (PARTITION BY [draft no] ORDER BY stepcode ) [seq] into #TmpApprovalF from #TmpApproval group by [Entity],[Select],[Document Type],[Draft No],[Document No],[Creator Name],[Posting Date],[Vendor Name],DocTotal,
U_AB_APPROVALCODE,WddCode,U_NAME,USERID,USER_CODE,Status ,Remarks ,[AStatus],AppSeq,StepCode order by [Draft No] , AppSeq 

select * into #TmpApprovalF1 from #TmpApprovalF

 
INSERT INTO #TABLE
select Entity,[Document Type],[Select],[Draft No],[Document No],[Creator Name],[Posting Date],[Vendor Name],DocTotal,U_AB_APPROVALCODE,
WddCode,U_NAME,[Status],[Remarks],
(select top 1 U_Name from #TmpApprovalF1 T0 where T0.[Seq]= T1.[Seq] -1 and T0.[Draft No] = T1.[Draft No]  and T0.AStatus = ''Y'' group by U_Name) as [Approvedby] ,'''' [ApproverRemarks],AppSeq,Seq,StepCode
from #TmpApprovalF T1 where AStatus=''W'' AND U_NAME=''' + @APPROVER + ''' 
order by [Draft No],AppSeq '

set @SQLVAR ='
INSERT INTO #TABLEF
select *  from(
select * from #TABLE WHERE SEQ=1 AND [DRAFT NO] NOT IN (select [DRAFT NO] from #TABLE WHERE SEQ=1 AND [Status]= ''Y'') UNION ALL
select * from #TABLE WHERE SEQ=2 AND [DRAFT NO] IN (select [DRAFT NO] from #TABLE WHERE SEQ=1 AND [Status]=''Y'') UNION ALL
select * from #TABLE WHERE SEQ=3 AND [DRAFT NO] IN (select [DRAFT NO] from #TABLE WHERE SEQ=2 AND [Status]=''Y'') UNION ALL
select * from #TABLE WHERE SEQ=4 AND [DRAFT NO] IN (select [DRAFT NO] from #TABLE WHERE SEQ=3 AND [Status]=''Y'') UNION ALL
select * from #TABLE WHERE SEQ=5 AND [DRAFT NO] IN (select [DRAFT NO] from #TABLE WHERE SEQ=4 AND [Status]=''Y'') UNION ALL
select * from #TABLE WHERE SEQ=6 AND [DRAFT NO] IN (select [DRAFT NO] from #TABLE WHERE SEQ=5 AND [Status]=''Y'') UNION ALL
select * from #TABLE WHERE SEQ=7 AND [DRAFT NO] IN (select [DRAFT NO] from #TABLE WHERE SEQ=6 AND [Status]=''Y'') UNION ALL
select * from #TABLE WHERE SEQ=8 AND [DRAFT NO] IN (select [DRAFT NO] from #TABLE WHERE SEQ=7 AND [Status]=''Y'') UNION ALL
select * from #TABLE WHERE SEQ=9 AND [DRAFT NO] IN (select [DRAFT NO] from #TABLE WHERE SEQ=8 AND [Status]=''Y'')) x order by x.[Draft No],x.[AppSeq]
delete from #TABLEF where StepCode in (select T0.StepCode from #TABLEF T0 where T0.[Status] in (''Y'',''N''))
Drop table #TmpApprovalF1
Drop table #TmpApprovalF
Drop table #TmpApproval
Drop table #TmpInitialise  
delete from #table'
  --select @ENTITYNAME
EXEC(@SQL + @SQLVAR)

FETCH NEXT FROM C1 INTO @ENTITYNAME
END
CLOSE C1;
DEALLOCATE C1;

END

--SELECT @SQL

--EXEC(@SQL)

SELECT * INTO #TMPCOMPANY FROM PWCL ..[@AB_COMPANYDATA]

SELECT ENTITY, U_AB_COMPANYNAME [ENTITY NAME],[DOCUMENT TYPE],[SELECT],[DRAFT NO],[DOCUMENT NO],[CREATOR NAME],[POSTING DATE],[VENDOR NAME],[AMOUNT(BEFORE GST)],
case when isnull(APPROVER,'') = '' then 'You Are The First Approver' else APPROVER end [APPROVED BY],
[APPROVAL GRID],APPROVALCODE,USERCODE,[STATUS], REMARKS,[REMARKS BY APPROVER] [REASON FOR NOT APPROVING] 
FROM #TABLEF
LEFT OUTER JOIN #TMPCOMPANY ON #TABLEF.ENTITY COLLATE DATABASE_DEFAULT = #TMPCOMPANY.NAME COLLATE DATABASE_DEFAULT
ORDER BY ENTITY,[DRAFT NO],APPSEQ

 DROP TABLE #TABLE
DROP TABLE #TABLEF
DROP TABLE #TMPCOMPANY

END