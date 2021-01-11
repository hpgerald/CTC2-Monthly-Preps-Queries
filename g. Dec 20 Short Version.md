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
SELECT DISTINCT PatientID, VisitDate, ARVStatusCode, ARVCode,
NumDaysDispensed, VisitTypeCode, TBRXIPTID,
NoDaysIPTDrugsDispensed
INTO srt_vst
FROM tblVisits
ORDER BY PatientID, VisitDate;

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
WHERE (((srt_apt.Given)<#1/1/2021#) AND ((srt_apt.Appointment)<#9/1/2021#)) 
GROUP BY srt_apt.PatientID 
ORDER BY Last(srt_apt.Given) DESC; 
```

##### 7: PICK THE LATEST VRL TEST FOR THE REPORTING PERIOD 
```sql
SELECT 
srt_tst.PatientID, Last(srt_tst.TestDate) AS [Test Date], Last(srt_tst.ResultDate) AS [Result Date], 
Last(srt_tst.ResultNumeric) AS Results, Last(srt_tst.ResultNotes) AS Notes 
INTO LT 
FROM srt_tst 
WHERE (((srt_tst.TestDate)<#1/1/2021#)) 
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
WHERE (((srt_sts.Status) Is Not Null) AND ((srt_sts.StatusDate)<#1/1/2021#)) 
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
WHERE (((srt_vst.NumDaysDispensed)>0) AND ((srt_vst.VisitDate)<#1/1/2021#)) 
GROUP BY srt_vst.PatientID 
ORDER BY Last(srt_vst.VisitDate) DESC; 
```

##### 17: AGGREGATE ON PTS LEVEL
```sql
SELECT LV.PatientID, tblExportPatients.DateOfBirth AS DoB, tblExportPatients.Sex AS Gender, 
IIf(IsNull([ReferredFromID]),"Unknown",[ReferredFromID]) AS [Entry Point], tblExportPatients.FileRef 
AS [File Namba], DateDiff("yyyy",[DoB],44196) AS Age, IIf([Gender]="Male","M","F") AS Sex, LV.[Last ART Visit]
AS [ART Visit], LV.Days, LA.[Appt Date], Round(IIf(IsNull([Appt Date]),[Days],[Appt Date]-[Last ART Visit]),0) 
AS Diff, Round(IIf([Days]>[Diff],[Days],[Diff]),0) AS Dispensed, IIf([Dispensed]>=180,"6 mo",IIf([Dispensed]>62,"3 mo", 
IIf([Dispensed]>31,"2 mo","1 mo"))) AS MMS, [ART Visit]+[Dispensed]+30 AS TX_CURR, [ART Visit]+[Days]+28 
AS LTFU_v0_Doc, [TX_CURR]+1 AS LTFU_v1_Adj, IIf(([Dispensed]<200 And [TX_CURR]>=44196),1,0) AS CURR_FY21Q1, 
LS.[Last Status], LS.[Status Date], LS.Reason, LS.Notes AS [Status Notes], LT.[Test Date], LT.[Result Date], 
LT.Results, LT.Notes AS [VRL Notes], IIf([Test Date]>43830 And [Test Date]<44197,1,0) AS Sample, 
IIf([Sample]=1 And [Results]>0,1,0) AS Result
FROM ((((LV LEFT JOIN LS ON LV.PatientID = LS.PatientID) LEFT JOIN LA ON LV.PatientID = LA.PatientID) 
LEFT JOIN tblExportPatients ON LV.PatientID = tblExportPatients.PatientID) LEFT JOIN LT ON LV.PatientID = LT.PatientID) 
ORDER BY LV.PatientID;
```

##### 20: SDI FOR THE REPORTING PERIOD 
```sql
SELECT srt_vst.PatientID, tblExportPatients.ReferredFromID AS [Entry Point], 
tblExportPatients.DateConfirmedHIVPositive AS Confirmed,
srt_vst.VisitDate AS [Start ART], srt_vst.NumDaysDispensed AS Days, IIf([Start ART]-[Confirmed]<8,"Yes","No") AS SDI 
FROM tblExportPatients RIGHT JOIN srt_vst ON tblExportPatients.PatientID = srt_vst.PatientID 
WHERE (((srt_vst.VisitDate) Between #10/1/2020# And #12/31/2020#) AND ((srt_vst.ARVStatusCode)=2)) 
ORDER BY tblExportPatients.DateConfirmedHIVPositive; 
```

