$aPrinterList = @()
 $StartTime = "22/04/2020 00:00:01 AM"
 $EndTime = "23/04/2020 6:00:01 PM"
 $Results = Get-WinEvent -FilterHashTable @{LogName="Print Server03/Operational"; ID=307; StartTime=$StartTime; EndTime=$EndTime;} -ComputerName "print-03"
 ForEach($Result in $Results){
 $ProperyData = [xml]$Result.ToXml()
 $PrinterName = $ProperyData.Event.UserData.DocumentPrinted.Param5
 If($PrinterName.Contains("HP-6850-03")){

 $hItemDetails = New-Object -TypeName psobject -Property @{
 DocName = $ProperyData.Event.UserData.DocumentPrinted.Param2
 UserName = $ProperyData.Event.UserData.DocumentPrinted.Param3
 MachineName = $ProperyData.Event.UserData.DocumentPrinted.Param4  
 PrinterName = $PrinterName
 PageCount = $ProperyData.Event.UserData.DocumentPrinted.Param8
 TimeCreated = $Result.TimeCreated
    }
 $aPrinterList += $hItemDetails
  }
}
 $aPrinterList | Export-Csv -LiteralPath C:\PrintServer\PrintAuditReport.csv 
