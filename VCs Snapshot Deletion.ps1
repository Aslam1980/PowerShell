Function Deletionemail{
    $CR = Read-Host "Please do provide the change Numbers using hypen seperation"
    $Outlook = New-Object -ComObject Outlook.Application
    $Mail = $Outlook.CreateItem(0)
    $Mail.To = 'dg-alight-global-mphasis-wintel@alight.com'
    $Mail.Subject = "snapshot Deletion - $CR"
    $Mail.Body = @'
    Hi Team,
    
    Snapshots Deleted.

    Regards,

   Wintel Team
 

'@
    $Mail.Send()
    #$Outlook.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null
}

Function RemoveSnapshots{
[CmdletBinding()]
param()
BEGIN{
    $progresspreference = "SilentlyContinue"
    $machines = Import-Csv 'C:\Work\Office\NGA_Patching\dcasnapdel.csv'
    $conn=Connect-VIServer -Server srv-dcb-erp-vc6.ngahr.hosting,hosdct01vc01.ngahr.hosting,srv-dcj-gen-vc6.ngahr.hosting,hosdcp01vc01.ngahr.hosting,hosdcp01vc02.ngahr.hosting -ErrorAction SilentlyContinue
   }
PROCESS{
    if($conn[0].IsConnected -eq $true){
    foreach ($machine in $machines){
    Get-Snapshot -VM $machine.Hostname.Split(".")[0] | Where Name -EQ $machine.ChangeRequest | Remove-Snapshot -ErrorAction SilentlyContinue -Confirm:$false 
    Write-Host "Snap shot Deleted for $($machine.Hostname) With Description $($machine.ChangeRequest)"
   }&Deletionemail
  }Else{
    Write-Host "Connection Unsuccessful please do check IF RSA Connected or NOT"
  }
}
END{
    Disconnect-VIServer -Server srv-dcb-erp-vc6.ngahr.hosting,hosdct01vc01.ngahr.hosting,srv-dcj-gen-vc6.ngahr.hosting,hosdcp01vc01.ngahr.hosting,hosdcp01vc02.ngahr.hosting -Confirm:$false | Out-Null
  }
}RemoveSnapshots
