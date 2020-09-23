<div align="center">

## BucketSort \(really really fast\)


</div>

### Description

Sorts Integer Values really fast.

on my 800mhz compu it sorts 100 000 values in 0.109 seconds...
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Rang3r](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rang3r.md)
**Level**          |Intermediate
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/rang3r-bucketsort-really-really-fast__1-29913/archive/master.zip)





### Source Code

```
<pre>
Private Sub Form_Load()
 Dim MyArr(99999) As Long
 Dim i As Long
 Dim t As Variant
 For i = LBound(MyArr) To UBound(MyArr)
  MyArr(i) = Rnd * 99999
 Next
 MsgBox "Click OK to start"
 t = Timer
 BucketSort MyArr
 MsgBox "READY!!!" & vbCrLf & "sorted 100000 values in " & Timer - t & " seconds"
 For i = LBound(MyArr) To UBound(MyArr)
  List1.AddItem MyArr(i)
 Next
End Sub
Private Sub BucketSort(ByRef Arr() As Long)
 Dim Buckets(99999) As Long
 Dim i As Long
 Dim j As Long
 Dim pos As Long
 For i = LBound(Arr) To UBound(Arr)
  Buckets(Arr(i)) = Buckets(Arr(i)) + 1
 Next
 pos = 0
 For i = LBound(Buckets) To UBound(Buckets)
  Do While Buckets(i) > 0
   Arr(pos) = i
   Buckets(i) = Buckets(i) - 1
   pos = pos + 1
  Loop
 Next
End Sub
</pre>
```

