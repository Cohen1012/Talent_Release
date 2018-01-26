DBCC CHECKIDENT(Interview_Info, RESEED, 0)

select Contact_Id,Name,CONVERT(varchar(100), UpdateTime, 111) UpdateTime from Contact_Info where Contact_Id = '2'

select Code_Id from Code where Contact_Id = '2'

select top 2 CONVERT(varchar(100), Contact_Date, 111) Contact_Date,Contact_Status,Remarks from Contact_Situation where Contact_Id = '1' order by Contact_Date desc

select top 2 cONVERT(varchar(100), Interview_Date, 111) Interview_Date from Interview_Info where Contact_Id = '1' order by Interview_Date desc

SELECT m.Contact_Id ,left(m.Code_Id,len(m.Code_Id)-1) as Code_Id from 
(SELECT Contact_Id,(SELECT cast(Code_Id AS NVARCHAR ) + ',' from Code 
where Contact_Id = ord.Contact_Id
FOR XML PATH('')) as Code_Id
from Code ord where Contact_Id = '1'
) M 

select Contact_Id from Contact_Info where ISNULL(Status,'NA') = ISNULL(ISNULL(null,Status),'NA') and
                                          UpdateTime >= ISNULL(null, UpdateTime) and
                                          UpdateTime <= ISNULL(null, UpdateTime) and
                                          ISNULL(Cooperation_Mode,'NA') = ISNULL(ISNULL(null,Cooperation_Mode),'NA') and
										  ISNULL(Place,'NA') Like ISNULL(ISNULL(null, Place),'NA') and
										  ISNULL(Skill,'NA') Like ISNULL(ISNULL(null, Skill),'NA')
