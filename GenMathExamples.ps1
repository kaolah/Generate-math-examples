#dokolko if you want count up to 20 you need to enter 21
$dokolko = 21
#$pocet  how many examples you want to generate
$pocet=462

$priklady=@()
$pomocna = 0
$n=@()

$m=1
#examples are generated
for ($i=0; $i -lt $dokolko; $i++) {
 for ($j=0; $j -lt $dokolko; $j++) {
        $v= $i+$j
        if ($v -lt $dokolko) {
        $priklady+="$i + $j ="
        }
        if ($i -ge $j) {$priklady+="$i - $j ="}
      
        }}
$pocetkombinacii = $priklady.count

if ($pocet -le $pocetkombinacii) {

    do {
        $nah = Get-Random -Maximum $pocetkombinacii
         foreach ($nn in $n) {
             if ($nn -eq $nah) { $pomocna = 1}
            }
        if ($pomocna -eq 1) {
                    $pomocna = 0
                }           
     else {
            $n+= $nah
            $m++
        } }
     
 until ($m -gt $pocet)


  

#save examples to WORD doc
$dok = $dokolko-1
Write-Host " "
Write-Host "************************************************" -ForegroundColor Yellow
Write-Host "Generating $pocet examples up to $dok  from posible combination $pocetkombinacii" -ForegroundColor Yellow
Write-Host "************************************************" -ForegroundColor Yellow
Write-Host " "
$word=new-object -ComObject "Word.Application"
$Word.Visible = $True
 
$doc = $word.documents.Add() 
$myDoc = $word.Selection
$myDoc.Font.Size = 24
$myDoc.PageSetup.TextColumns.SetCount(3)

foreach ($mm in $n) {
    $text1 = $priklady[$mm]
    $myDoc.TypeText("$text1")
    $myDoc.TypeParagraph()
    
    }

Write-Host " "
Write-host "*********** Done ****************" -ForegroundColor Red
Write-Host " "

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
Remove-Variable doc,Word
[gc]::collect()
[gc]::WaitForPendingFinalizers()
}

else {
Write-host "********************************************************************************" -ForegroundColor red
Write-host "Please lower the number of examples, and run again. It cannot be higher than number of all possible combinations $pocetkombinacii." -ForegroundColor red
Write-host "********************************************************************************" -ForegroundColor red
    }

pause
