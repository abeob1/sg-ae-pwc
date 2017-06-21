create proc sp_Invoice_Validation
@HeaderID numeric(18,0)
as

select T2.Amount PaymentAmt, sum(T1.GrossPrice*T1.Quantity) InvoiceAmt
from InvoiceHeader T0
join InvoiceLine T1 on T0.ID=t1.HeaderID
join PaymentMean T2 on T2.HeaderID=T0.ID
where T0.PaymentType in ('IN','OUT') and T0.ID=1
group by T2.Amount