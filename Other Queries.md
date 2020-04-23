*****FIRST DTG TRANSITION:*****
```sql
SELECT tblVisits.PatientID, First(tblVisits.VisitDate) AS [First Switch], First(tblVisits.ARVCode) AS Regimen, First(tblVisits.NumDaysDispensed) AS Dispensed
FROM tblVisits
WHERE (((tblVisits.ARVStatusCode)=8) AND ((tblVisits.VisitDate) Between #8/1/2019# And #3/31/2020#) AND ((tblVisits.ARVCode)>99))
GROUP BY tblVisits.PatientID
ORDER BY First(tblVisits.VisitDate);
```
