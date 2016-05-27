# Microsoft VBA Code for Jalali(Shamsi) Calendar
This code is VBA functions for convert jalali date to gregorian date and vice versa. Jalali calendar is also known as 
Jalali, Persian, Khayyami, Khorshidi, Shamsi canlendar


## Origin code
this code is just vba port for js jalali js code (https://github.com/jalaali/jalaali-js) by Behrang (https://github.com/behrang)


## Description
you can find usefull comments in code 


```vba

Private Sub CommandButton1_Click()
Dim nowDate As Date
nowDate = Date
 
Dim jalaliDateArray1
jalaliDateArray1 = toJalaaliFromDateObject(nowDate)
 
'MsgBox "Now= " & jalaliDateArray1(0) & "/" & jalaliDateArray1(1) & "/" & jalaliDateArray1(2)

Dim gregorianDate As Date

gregorianDate = toGregorianDateObject(1394, 12, 19)

MsgBox "Date Object From Jalali= " & gregorianDate

Dim jalaliDateArray
jalaliDateArray = toJalaali(2016, 3, 9)

'MsgBox jalaliDateArray(0) & "/" & jalaliDateArray(1) & "/" & jalaliDateArray(2)

Dim gregorialDateArray
gregorialDateArray = toGregorian(1394, 12, 19)

'MsgBox gregorialDateArray(0) & "/" & gregorialDateArray(1) & "/" & gregorialDateArray(2)

End Sub

```

