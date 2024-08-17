Function send-email{
    $CR = Read-Host "Please do provide the change Numbers using hypen seperation"
    $Outlook = New-Object -ComObject Outlook.Application
    $Mail = $Outlook.CreateItem(0)
    $Mail.To = 'dg-alight-global-mphasis-wintel@alight.com'
    $Mail.Subject = "snapshot Creation - $CR"
    $Mail.Body = @'
    Hi Team,
                
    Snapshots taken.

    Regards,

    Wintel
 

'@
    $Mail.Send()
    #$Outlook.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null
}

Function NewSnapshots{
[CmdletBinding()]
param()
BEGIN{
    $progresspreference = "SilentlyContinue"
    $VMHost = Import-Csv ‘C:\Work\Office\NGA_Patching\snapshotcreation.csv'
    $conn=Connect-VIServer -Server srv-dcb-erp-vc6.ngahr.hosting,srv-dcj-gen-vc6.ngahr.hosting,hosdct01vc01.ngahr.hosting,hosdcp01vc01.ngahr.hosting,hosdcp01vc02.ngahr.hosting -ErrorAction SilentlyContinue
    }
PROCESS{
    if($conn[0].IsConnected -eq $true){
    foreach ($machine in $VMHost){
    New-Snapshot -VM $machine.Hostname.split(".")[0] -Name $machine.ChangeRequest -Description $machine.ChangeRequest -Memory:$false -Quiesce:$false  -Confirm:$false
    Write-Host "Snap shot created for $($machine.Hostname) With Description $($machine.ChangeRequest)"
  }&send-email
 }else{
    Write-Host "Connection Unsuccessful please do check IF RSA Connected or NOT"
 }
}
END{
    Disconnect-VIServer -Server srv-dcb-erp-vc6.ngahr.hosting,srv-dcj-gen-vc6.ngahr.hosting,hosdct01vc01.ngahr.hosting,hosdcp01vc01.ngahr.hosting,hosdcp01vc02.ngahr.hosting -Confirm:$false | Out-Null
  }
}NewSnapshots 