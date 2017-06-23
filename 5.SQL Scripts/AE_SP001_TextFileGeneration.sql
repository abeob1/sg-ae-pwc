alter procedure [dbo].[AE_SP001_TextFileGeneration]

@DateFrom varchar(20),

@DateTo varchar(20),

@OUCodeFrom varchar(20),

@OUCodeTo varchar (20),

@Entity varchar (30)

as

begin



Declare @SQLString varchar(max)

declare @year integer

Declare @monthName varchar(20)

Declare @month integer

Declare @Day varchar(5)

Declare @Period varchar(20)





select top(1) T0.Name into #period from OFPR T0

delete #period 

---update on 07052015

--------------------- Get Period from the SAP Period Table based on the Date

set @SQLString = '

SELECT isnull(RIGHT(T0.[Code],4) + ''0'' + LEFT (T0.[Code],2),''0'') [Name]

           FROM ' + @Entity + '.. OFPR T0 WHERE month(T0.[F_RefDate]) = month('''+ @DateFrom + ''')  and  year(T0.[F_RefDate]) = year(''' + @DateFrom + ''' ) '



insert into #period 

execute (@SQLString )



select @Period = Name from #period 



-------------------- Table structure creation



 select top(1) jdt1.Account [AcctCode] , jdt1.Account  [RefDate] , jdt1.OcrCode3 [OU Code], 

 jdt1.Account [Entity], jdt1.Account [DC], jdt1.Credit [Amount] , jdt1.Account [Cat] into #level1 from JDT1

 

 delete  #level1



/*

set @year = LEFT(@DateTo ,4)

set @Day = RIGHT(@DateTo ,2)

set @month = left(RIGHT(@DateTo,4),2)



select @monthName = DateName( month , DateAdd( month , @month , 0 ) - 1 )

select @Day = day(dateadd(day, -1, dateadd(month, 1, dateadd(day, 1 - day(@DateTo), @DateTo))))



select @Period = code from [@AB_IPOWERPERIOD] where Name = @monthName + ' ' + @Day



set @Period =  cast(@year as varchar) + replicate('0', 3 - LEN(@period)) + cast(@period as varchar)





set @SQLString = 'select * from (

SELECT T0.AcctCode , ''' +  @Period + ''' [RefDate]   ,isnull(T1.OcrCode3,'''') [OU Code] , ''' + @Entity + ''' [Entity],

case when T0.GroupMask = 1 then ''D''

when T0.GroupMask = 2 then ''C''

when T0.GroupMask = 3 then ''C'' end [DC], abs(sum(T1.Debit - T1.Credit)) [Amount]   

FROM ' + @Entity + '.. OACT T0 INNER JOIN ' + @Entity + '.. JDT1 T1 on T0.AcctCode  = T1.Account  INNER JOIN ' + @Entity + '.. OJDT T2 

ON T1.TransId = T2.TransId

where T2.RefDate <= ''' + @DateTo + ''' and T0.GroupMask  in (1,2,3)

--and T1.OcrCode3 >= ''' + @OUCodeFrom + ''' and T1.OcrCode3 <= ''' + @OUCodeTo + '''  

group by T0.AcctCode , T1.OcrCode3  ,T0.GroupMask

*/



------------------------  Query for getting the Balance sheet 

set @SQLString = 'select *   from (

--------------------- Balance Sheet 

SELECT T0.AcctCode , ''' +  @Period + ''' [RefDate]   ,isnull(T1.OcrCode3,'''') [OU Code] , ''' + @Entity + ''' [Entity],

case when sum(T1.Debit - T1.Credit) >= 0 then ''D'' else ''C'' end [DC], abs(sum(T1.Debit - T1.Credit)) [Amount] , ''JE'' [Cat]

FROM ' + @Entity + '.. OACT T0 INNER JOIN ' + @Entity + '.. JDT1 T1 on T0.AcctCode  = T1.Account  INNER JOIN ' + @Entity + '.. OJDT T2 

ON T1.TransId = T2.TransId

where T2.RefDate <= ''' + @DateTo + ''' and T0.GroupMask  in (1,2,3)

and T1.OcrCode3 >= ''' + @OUCodeFrom + ''' and T1.OcrCode3 <= ''' + @OUCodeTo + '''  

group by T0.AcctCode , T1.OcrCode3  ,T0.GroupMask

union all

-------------------- Profit and loss

SELECT T0.AcctCode , ''' +  @Period + ''' [Period]   ,isnull(T1.OcrCode3,'''') [OU Code] , ''' + @Entity + ''' [Entity],

case when sum(T1.Debit - T1.Credit) >= 0 then ''D''

when sum(T1.Debit - T1.Credit) < 0 then ''C''

 end [DC], abs(sum(T1.Debit - T1.Credit)) [Amount] , ''JE'' [Cat]

FROM ' + @Entity + '.. OACT T0 INNER JOIN ' + @Entity + '.. JDT1 T1 on T0.AcctCode  = T1.Account  INNER JOIN ' + @Entity + '.. OJDT T2 

ON T1.TransId = T2.TransId

where T2.RefDate >= ''' + @DateFrom + ''' and T2.RefDate <= ''' + @DateTo + ''' and T0.GroupMask not in (1,2,3)

and T1.OcrCode3 >= ''' + @OUCodeFrom + ''' and T1.OcrCode3 <= ''' + @OUCodeTo + ''' 

group by T0.AcctCode , T1.OcrCode3  ,T0.GroupMask 

union all

------------------- Import Statistics

select T0.U_AB_GLCODE , T0.U_AB_PERIOD ,isnull(T0.U_AB_OPER_UNIT,'''') [OU Code] , ''' + @Entity + ''' [Entity], T0.U_AB_DEBIT_CREDIT , 

sum(T0.U_AB_AMOUNT) [Amount] , ''IM'' [Cat]

from ' + @Entity + '..[@AB_STATITISTICSDATA] T0

where T0.U_AB_PERIOD = ''' +  @Period + '''

and T0.U_AB_OPER_UNIT >= ''' + @OUCodeFrom + ''' and T0.U_AB_OPER_UNIT <= ''' + @OUCodeTo + '''

group by T0.U_AB_GLCODE , T0.U_AB_PERIOD ,T0.U_AB_OPER_UNIT, T0.U_AB_DEBIT_CREDIT ) tmp

order by tmp.AcctCode'



insert into #level1 

    execute (@SQLString)

    

 ---------------  Segregating Journal memos and Statistics data
 
 

select * into #memojournals from #level1 T0 where T0.Cat = 'JE'

select * into #importstatistics from #level1 T0 where T0.Cat = 'IM'



---------------  including the distribution rules in the journal memos data

delete #level1 



--if @Entity = 'IAS7'

-- begin

--  set @SQLString = '

--   select ''"'' + T0.AcctCode + ''"'', ''"'' + T0.RefDate + ''"'', ''"'' + T2.PrcCode + ''"'' [OU Code] , ''"'' + T0.Entity + ''"'' , ''"'' + T0.DC + ''"'', (T2.PrcAmount/T1.OcrTotal ) * T0.Amount [Amount], '''' [Cat]

--   from #memojournals T0 left outer Join ' + @Entity + '.. OOCR T1 ON T1.OcrCode = T0.[OU Code] left outer Join ' + @Entity + '.. OCR1 T2 

--    on T1.OcrCode = T2.OcrCode where T2.ValidFrom <= ''' + @DateFrom + ''' and (T2.ValidTo >= ''' + @DateTo + ''' or T2.ValidTo is null )'

-- end

--else

-- begin

--  set @SQLString = '

--     select ''"'' + T0.AcctCode + ''"'', ''"'' + T0.RefDate + ''"'', ''"'' + T2.PrcCode + ''"'' [OU Code] , ''"'' + T0.Entity + ''"'' , ''"'' + T0.DC + ''"'', (T2.PrcAmount/T1.OcrTotal ) * T0.Amount [Amount], '''' [Cat]

--      from #memojournals T0 left outer Join ' + @Entity + '.. OOCR T1 ON T1.OcrCode = T0.[OU Code] left outer Join ' + @Entity + '.. OCR1 T2 

--       on T1.OcrCode = T2.OcrCode '

-- end



 set @SQLString = '

   select ''"'' + isnull(T0.AcctCode,'''') + ''"'' [AcctCode], ''"'' + isnull(T0.RefDate,'''') + ''"'' [RefDate], ''"'' + isnull(T2.PrcCode,'''') + ''"'' [OU Code] , ''"'' + isnull(T0.Entity,'''') + ''"'' [Entity] , ''"'' + isnull(T0.DC,'''') + ''"'' [DC], isnull( (T2.PrcAmount/T1.OcrTotal ) * T0.Amount,0.00) [Amount], '''' [Cat]
   into #level1
   from #memojournals T0 left outer Join ' + @Entity + '.. OOCR T1 ON T1.OcrCode = T0.[OU Code] left outer Join ' + @Entity + '.. OCR1 T2 

    on T1.OcrCode = T2.OcrCode where T2.ValidFrom <= ''' + @DateFrom + ''' and (T2.ValidTo >= ''' + @DateTo + ''' or T2.ValidTo is null )
	
	select T0.AcctCode , T0.RefDate , T0.[OU Code] , T0.Entity , T0.DC , T0.Amount from #level1 T0 where T0.Amount > 0 and T0.[OU Code] not like ''Cen%'' 

union all

select ''"'' + isnull(T0.AcctCode,'''') + ''"'' [AcctCode], ''"'' + isnull(T0.RefDate,'''') + ''"'' [RefDate], ''"'' + isnull(T0.[OU Code],'''') + ''"'' [OU Code] , ''"'' + isnull(T0.Entity,'''') + ''"'' [Entity] , ''"'' + isnull(T0.DC,'''') + ''"'' [DC] , isnull(T0.Amount,0.00)  [Amount] from #importstatistics T0

where T0.Amount > 0 and T0.[OU Code] not like ''Cen%'' 

order by AcctCode
	
	
	'

	
-- print @SQLString

 

--insert into #level1 

execute (@SQLString)

 

----------------  Final Output

--select T0.AcctCode , T0.RefDate , T0.[OU Code] , T0.Entity , T0.DC , T0.Amount from #level1 T0 where T0.Amount > 0

--union all

--select '"' + T0.AcctCode + '"' , '"' + T0.RefDate + '"' , '"' + T0.[OU Code] + '"' , '"' + T0.Entity + '"' , '"' + T0.DC + '"' , T0.Amount from #importstatistics T0

--where T0.Amount > 0

--order by AcctCode



--select * from #level1

--print @SQLString



end