##### 1: SORT APPOINTMENTS IN ASCENDING ORDER
```sql
SELECT DISTINCT 
PatientID, DateAppointmentGiven AS Given, DateOfAppointment AS Appointment 
INTO srt_apt 
FROM tblAppointments 
ORDER BY PatientID, DateAppointmentGiven;
```

##### 2: SORT VISITS IN ASCENDING ORDER 
```sql
SELECT DISTINCT tblVisits.PatientID, tblVisits.VisitDate, tblVisits.ARVStatusCode, tblVisits.ARVCode,
tblVisits.NumDaysDispensed, tblVisits.VisitTypeCode, tblVisits.TBRXIPTID,
tblVisits.NoDaysIPTDrugsDispensed, tblVisits.NowPregnant, tblVisits.NowBreastfeeding,
tblVisits.CancerScreeningID
INTO srt_vst
FROM tblVisits
ORDER BY tblVisits.PatientID, tblVisits.VisitDate;

```

##### 3: SORT STATUS IN ASCENDING ORDER 
```sql
SELECT DISTINCT 
PatientID, StatusDate, Status, StatusReason, Notes, TransferredOutWhere 
INTO srt_sts 
FROM tblStatus 
ORDER BY PatientID, StatusDate; 
```

##### 4: SORT VRL TESTS IN ASCENDING ORDER 
```sql
SELECT DISTINCT 
PatientID, TestDate, ResultDate, ResultNumeric, ResultNotes, ResultReturnDate, 
DateInitiatedEAC, DateHVLTestAfterEAC
INTO srt_tst 
FROM tblTests 
WHERE TestTypeID="VRL" 
ORDER BY PatientID, TestDate; 
```

##### 5: PICK THE LATEST APPOINTMENT GIVEN FOR THE REPORTING PERIOD 
```sql
SELECT 
srt_apt.PatientID, Last(srt_apt.Given) AS Given, Last(srt_apt.Appointment) AS [Appt Date] 
INTO LA 
FROM srt_apt 
WHERE (((srt_apt.Given)<#9/1/2020#) AND ((srt_apt.Appointment)<#4/1/2021#)) 
GROUP BY srt_apt.PatientID 
ORDER BY Last(srt_apt.Given) DESC; 
```

##### 6: PICK THE LATEST EAC SESSION FOR THE REPORTING PERIOD 
```sql
SELECT 
srt_tst.PatientID, Last(srt_tst.DateInitiatedEAC) AS [Initiated EAC], 
Last(srt_tst.DateHVLTestAfterEAC) AS [Test After EAC] 
INTO LE 
FROM srt_tst 
WHERE (((srt_tst.DateInitiatedEAC)<#9/1/2020#)) 
GROUP BY srt_tst.PatientID 
ORDER BY Last(srt_tst.DateInitiatedEAC) DESC; 
```

##### 7: PICK THE LATEST VRL TEST FOR THE REPORTING PERIOD 
```sql
SELECT 
srt_tst.PatientID, Last(srt_tst.TestDate) AS [Test Date], Last(srt_tst.ResultDate) AS [Result Date], 
Last(srt_tst.ResultNumeric) AS Results, Last(srt_tst.ResultNotes) AS Notes 
INTO LT 
FROM srt_tst 
WHERE (((srt_tst.TestDate)<#9/1/2020#)) 
GROUP BY srt_tst.PatientID 
ORDER BY Last(srt_tst.TestDate) DESC; 
```

##### 8: PICK THE LATEST STATUS FOR THE REPORTING PERIOD 
```sql
SELECT 
srt_sts.PatientID, Last(srt_sts.Status) AS [Last Status], Last(srt_sts.StatusDate) AS [Status Date], 
Last(srt_sts.StatusReason) AS Reason, Last(srt_sts.Notes) AS Notes 
INTO LS 
FROM srt_sts 
WHERE (((srt_sts.Status) Is Not Null) AND ((srt_sts.StatusDate)<#9/1/2020#)) 
GROUP BY srt_sts.PatientID 
ORDER BY Last(srt_sts.StatusDate) DESC; 
```

##### 9: PICK THE LATEST ART VISIT FOR THE REPORTING PERIOD 
```sql
SELECT 
srt_vst.PatientID, Last(srt_vst.VisitDate) AS [Last ART Visit], Last(srt_vst.NumDaysDispensed) AS Days, 
Last(srt_vst.ARVStatusCode) AS Status, Last(srt_vst.ARVCode) AS Regimen, Last(srt_vst.NowPregnant) AS Pregnant 
INTO LV 
FROM srt_vst 
WHERE (((srt_vst.NumDaysDispensed)>0) AND ((srt_vst.VisitDate)<#9/1/2020#)) 
GROUP BY srt_vst.PatientID 
ORDER BY Last(srt_vst.VisitDate) DESC; 
```

##### 10: PICK TX_NEW FROM THE PREVIOUS MONTH OF REPORTING PERIOD 
```sql
SELECT DISTINCT 
srt_vst.PatientID, srt_vst.VisitDate AS [Start ART], srt_vst.NumDaysDispensed AS Days 
INTO FER 
FROM srt_vst 
WHERE (((srt_vst.VisitDate) Between #7/1/2020# And #7/31/2020#) AND ((srt_vst.ARVStatusCode)=2)) 
ORDER BY srt_vst.VisitDate; 
```

##### 11: PICK LATEST IPT STATUS FOR THE ALL TIME 
```sql
SELECT 
srt_vst.PatientID, Last(srt_vst.VisitDate) AS [Last IPT], Last(srt_vst.TBRXIPTID) AS [Last Status] 
INTO LIPT 
FROM srt_vst 
WHERE (((srt_vst.TBRXIPTID) Like "% IPT")) 
GROUP BY srt_vst.PatientID 
ORDER BY Last(srt_vst.VisitDate) DESC; 
```

##### 12: PICK THOSE STARTED IPT 6-MO +1 MO FOR THE REPORTING PERIOD 
```sql
SELECT DISTINCT 
srt_vst.PatientID, srt_vst.VisitDate AS [Start IPT], srt_vst.TBRXIPTID 
INTO FIPT 
FROM srt_vst 
WHERE (((srt_vst.VisitDate) Between #1/1/2020# And #1/31/2020#) AND ((srt_vst.TBRXIPTID)="START IPT")) 
ORDER BY srt_vst.VisitDate; 
```

##### 13: PUT ALL CECAP DATA INTO CECAP TABLE
```sql
SELECT srt_vst.PatientID, srt_vst.VisitDate, srt_vst.NumDaysDispensed AS Drugs,
srt_vst.NowPregnant AS PG, srt_vst.NowBreastfeeding AS BF, srt_vst.CancerScreeningID AS CXCA_Code,
IIf([CXCA_Code]="1","Not Screened",IIf([CXCA_Code]="2","VIA Negative",IIf([CXCA_Code]="3","VIA Positive",
IIf([CXCA_Code]="4","Start Cryo-therapy",IIf([CXCA_Code]="5","Stop Cryo-therapy",IIf([CXCA_Code]="6","Leep/Biopsy",IIf([CXCA_Code]="7","Referral"))))))) AS CXCA_Data 
INTO cecap
FROM srt_vst
WHERE (((srt_vst.CancerScreeningID) Between "1" And "7"))
ORDER BY srt_vst.PatientID, srt_vst.VisitDate;
```

##### 14: PUT ALL CECAP SCREENING DATA INTO CECAP SCREEN TABLE
```sql
SELECT cecap.PatientID, Last(cecap.PG) AS PG, Last(cecap.BF) AS BF, Last(cecap.VisitDate) AS [Date Screened], Last(cecap.CXCA_Data) AS [Screening Results] INTO cecap_screen
FROM cecap
WHERE (((cecap.CXCA_Code) Between "2" And "3"))
GROUP BY cecap.PatientID;
```

##### 15: PUT ALL CECAP TREATMENT DATA INTO CECAP TX TABLE
```sql
SELECT cecap.PatientID, Last(cecap.VisitDate) AS [TX Date], Last(cecap.CXCA_Data) AS TX_Type INTO cecap_treatment
FROM cecap
WHERE (((cecap.CXCA_Code)="4" Or (cecap.CXCA_Code)="6"))
GROUP BY cecap.PatientID;
```

##### 16: PICK LATEST ART AND VRL DETAILS FOR THE REPORTING PERIOD 
```sql
SELECT LV.PatientID, tblExportPatients.DateOfBirth AS DoB, tblExportPatients.Sex AS Gender, 
IIf(IsNull([ReferredFromID]),"Unknown",[ReferredFromID]) AS [Entry Point], tblExportPatients.FileRef AS [File Namba], 
DateDiff("yyyy",[DoB],44074) AS Age, IIf([Gender]="Male","M","F") AS Sex, LV.[Last ART Visit] AS [ART Visit],
LV.Days, LA.[Appt Date], Round(IIf(IsNull([Appt Date]),[Days],[Appt Date]-[Last ART Visit]),0) AS Diff, 
Round(IIf([Days]>[Diff],[Days],[Diff]),0) AS Dispensed, 
IIf([Dispensed]>=180,"6 mo",IIf([Dispensed]>62,"3 mo",IIf([Dispensed]>31,"2 mo","1 mo"))) AS MMS, 
[ART Visit]+[Dispensed]+30 AS TX_CURR,
[ART Visit]+[Days]+28 AS LTFU_v0_Doc, [TX_CURR]+1 AS LTFU_v1_Adj, 
IIf(([Dispensed]<200 And [TX_CURR]>=44074),1,0) AS CURR_AUG, 
LS.[Last Status], LS.[Status Date], LT.[Test Date], LT.Results, LT.Notes, LE.[Initiated EAC], LE.[Test After EAC], 
LV.Regimen, IIf([Regimen]>99,IIf([Regimen]<106,"DTG","Other"),"Other") AS [DTG or OTH], LV.Pregnant 
FROM ((((LV LEFT JOIN LE ON LV.PatientID = LE.PatientID) LEFT JOIN LT ON LV.PatientID = LT.PatientID) 
LEFT JOIN LS ON LV.PatientID = LS.PatientID) LEFT JOIN LA ON LV.PatientID = LA.PatientID) 
LEFT JOIN tblExportPatients ON LV.PatientID = tblExportPatients.PatientID; 
```

##### 17: RAS WEEKLY MONITORING on LTFU, VRL, and CECAP
```sql
SELECT LV.PatientID, tblExportPatients.DateOfBirth AS DoB, tblExportPatients.Sex AS Gender, 
IIf(IsNull([ReferredFromID]),"Unknown",[ReferredFromID]) AS [Entry Point], tblExportPatients.FileRef 
AS [File Namba], DateDiff("yyyy",[DoB],44074) AS Age, IIf([Gender]="Male","M","F") AS Sex, LV.[Last ART Visit]
AS [ART Visit], LV.Days, LA.[Appt Date], Round(IIf(IsNull([Appt Date]),[Days],[Appt Date]-[Last ART Visit]),0) 
AS Diff, Round(IIf([Days]>[Diff],[Days],[Diff]),0) AS Dispensed, IIf([Dispensed]>=180,"6 mo",IIf([Dispensed]>62,"3 mo", 
IIf([Dispensed]>31,"2 mo","1 mo"))) AS MMS, [ART Visit]+[Dispensed]+30 AS TX_CURR, [ART Visit]+[Days]+28 
AS LTFU_v0_Doc, [TX_CURR]+1 AS LTFU_v1_Adj, IIf(([Dispensed]<200 And [TX_CURR]>=44074),1,0) AS CURR_AUG, 
LS.[Last Status], LS.[Status Date], LS.Reason, LS.Notes AS [Status Notes], LT.[Test Date], LT.[Result Date], 
LT.Results, LT.Notes AS [VRL Notes], IIf([Test Date]>43738 And [Test Date]<44105,1,0) AS Sample, 
IIf([Sample]=1 And [Results]>0,1,0) AS Result, IIf([Gender]="Female" And [CURR_AUG]=1 And [Age]>=15 And [Age] <=49,1,0) 
AS CECAP, cecap_screen.PG, cecap_screen.BF, cecap_screen.[Date Screened], cecap_screen.[Screening Results], 
cecap_treatment.[TX Date], cecap_treatment.TX_Type
FROM (((((LV LEFT JOIN LS ON LV.PatientID = LS.PatientID) LEFT JOIN LA ON LV.PatientID = LA.PatientID) 
LEFT JOIN tblExportPatients ON LV.PatientID = tblExportPatients.PatientID) LEFT JOIN LT ON LV.PatientID = LT.PatientID) 
LEFT JOIN cecap_screen ON LV.PatientID = cecap_screen.PatientID) 
LEFT JOIN cecap_treatment ON LV.PatientID = cecap_treatment.PatientID
ORDER BY LV.PatientID;
```

##### 18: IPT COMPLETION FOR THE REPORTING PERIOD 
```sql
SELECT FIPT.PatientID, FIPT.[Start IPT], LIPT.[Last IPT], LIPT.[Last Status] AS [Last IPT Status], 
LV.[Last ART Visit], LV.Days, LA.[Appt Date], LS.[Last Status], LS.[Status Date], 
IIf([Last IPT Status]="CPLT IPT","Yes","No") AS Completed 
FROM (((FIPT LEFT JOIN LIPT ON FIPT.PatientID=LIPT.PatientID) LEFT JOIN LV ON FIPT.PatientID=LV.PatientID) 
LEFT JOIN LA ON FIPT.PatientID=LA.PatientID) LEFT JOIN LS ON FIPT.PatientID=LS.PatientID 
ORDER BY FIPT.[Start IPT], LIPT.[Last IPT], LIPT.[Last Status]; 
```

##### 19: EARLY RETENTION FOR THE REPORTING PERIOD 
```sql
SELECT FER.PatientID, FER.[Start ART], FER.Days, LV.[Last ART Visit], LV.Days AS [Last ART Days], 
LS.[Last Status], LS.[Status Date], LA.[Appt Date], IIf([Last ART Visit]>[Start ART],"Yes","No") AS [First Refill] 
FROM ((FER LEFT JOIN LV ON FER.PatientID = LV.PatientID) LEFT JOIN LS ON FER.PatientID = LS.PatientID) 
LEFT JOIN LA ON FER.PatientID = LA.PatientID; 
```

##### 20: SDI FOR THE REPORTING PERIOD 
```sql
SELECT srt_vst.PatientID, tblExportPatients.ReferredFromID AS [Entry Point], 
tblExportPatients.DateConfirmedHIVPositive AS Confirmed,
srt_vst.VisitDate AS [Start ART], srt_vst.NumDaysDispensed AS Days, IIf([Start ART]-[Confirmed]<8,"Yes","No") AS SDI 
FROM tblExportPatients RIGHT JOIN srt_vst ON tblExportPatients.PatientID = srt_vst.PatientID 
WHERE (((srt_vst.VisitDate) Between #8/1/2020# And #8/31/2020#) AND ((srt_vst.ARVStatusCode)=2)) 
ORDER BY tblExportPatients.DateConfirmedHIVPositive; 
```


