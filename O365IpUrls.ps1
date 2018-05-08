<#


    Retrieves Ips and Url for Office 365 from REST webservice
    Creates csv files for each O4365 workload

    by Pierluigi Pesce aka pierfish


#>

Function Retrieve-Ips($Workload)
{
    # webservice root URL
    $ws = "https://endpoints.office.com"

    # path where client ID and latest version number will be stored
    $datapath = $Env:TEMP + "\endpoints_clientid_latestversion.txt"

    # fetch client ID and version if data file exists; otherwise create new file
    if (Test-Path $datapath) {
        $content = Get-Content $datapath
        $clientRequestId = $content[0]
        $lastVersion = $content[1]
    }
    else {
        $clientRequestId = [GUID]::NewGuid().Guid
        $lastVersion = "0000000000"
        @($clientRequestId, $lastVersion) | Out-File $datapath
    }

    # call version method to check the latest version, and pull new data if version number is different
    $version = Invoke-RestMethod -Uri ($ws + "/version/O365Worldwide?clientRequestId=" + $clientRequestId)
    
    Write-Host "New version of Office 365 worldwide commercial service instance endpoints detected"
    
    # write the new version number to the data file
    @($clientRequestId, $version.latest) | Out-File $datapath
    
    # invoke endpoints method to get the new data
    $endpointSets = Invoke-RestMethod -Uri ($ws + "/endpoints/O365Worldwide?clientRequestId=" + $clientRequestId)

    # filter results for Allow and Optimize endpoints, and transform these into custom objects with port and category
    $flatUrls = $endpointSets | where {($_.ServiceArea -match $Workload)} | ForEach-Object {
        $endpointSet = $_
        $allowUrls = $(if ($endpointSet.allowUrls.Count -gt 0) { $endpointSet.allowUrls } else { @() })
        $optimizeUrls = $(if ($endpointSet.optimizeUrls.Count -gt 0) { $endpointSet.optimizeUrls } else { @() })
        
        $allowUrlCustomObjects = $allowUrls | ForEach-Object {
            [PSCustomObject]@{
                category = "Allow";
                url      = $_;
                # Allow URLs should permit traffic across both Allow and Optimize ports
                tcpPorts = (($endpointSet.allowTcpPorts, $endpointSet.optimizeTcpPorts) | Where-Object { $_ -ne $null }) -join ",";
                udpPorts = (($endpointSet.allowUdpPorts, $endpointSet.optimizeUdpPorts) | Where-Object { $_ -ne $null }) -join ",";
                serviceArea = $endpointSet.serviceArea;
            }
        }
        $optimizeUrlCustomObjects = $optimizeUrls | ForEach-Object {
            [PSCustomObject]@{
                category = "Optimize";
                url      = $_;
                tcpPorts = $endpointSet.optimizeTcpPorts;
                udpPorts = $endpointSet.optimizeUdpPorts;
                serviceArea = $endpointSet.serviceArea;
            }
        }
        $allowUrlCustomObjects, $optimizeUrlCustomObjects
    }

    $flatIps = $endpointSets | where {($_.ServiceArea -match $Workload)}  | ForEach-Object {
        $endpointSet = $_
        $ips = $(if ($endpointSet.ips.Count -gt 0) { $endpointSet.ips } else { @() })
        # IPv4 strings have dots while IPv6 strings have colons
        $ip4s = $ips | Where-Object { $_ -like '*.*' }
        
        $allowIpCustomObjects = @()
        if ($endpointSet.allowTcpPorts -or $endpointSet.allowUdpPorts) {
            $allowIpCustomObjects = $ip4s | ForEach-Object {
                [PSCustomObject]@{
                    category = "Allow";
                    ip = $_;
                    tcpPorts = $endpointSet.allowTcpPorts;
                    udpPorts = $endpointSet.allowUdpPorts;
                    serviceArea = $endpointSet.serviceArea;
                }
            }
        }
        $optimizeIpCustomObjects = @()
        if ($endpointSet.optimizeTcpPorts -or $endpointSet.optimizeUdpPorts) {
            $optimizeIpCustomObjects = $ip4s | ForEach-Object {
                [PSCustomObject]@{
                    category = "Optimize";
                    ip       = $_;
                    tcpPorts = $endpointSet.optimizeTcpPorts;
                    udpPorts = $endpointSet.optimizeUdpPorts;
                    serviceArea = $endpointSet.serviceArea;
                }                
            }
        }
        $allowIpCustomObjects, $optimizeIpCustomObjects
    }
    
    Write-Output "IPV4 Firewall IP Address Ranges"
    $OutCsvIP=$Workload + "_Ips.csv"
    $OutCsvIP
        
    $OutIp=@()

    foreach($ipp in $flatIps)
    {
        $OutIp+=$ipp
    }
        
    $OutIp|Export-Csv -Path $Env:TEMP\$OutCsvIP -NoTypeInformation
        
    Write-Output "URLs for Proxy Server"
    $OutCsvUrl=$Workload + "_Urls.csv"
    $OutCsvUrl
    $flatUrls|Export-Csv -Path $Env:TEMP\$OutCsvUrl -NoTypeInformation -Force

}
    


Function Show-Menu
{
     param (
           [string]$Title = 'Select Office 365 Workload'
     )
     cls
     Write-Host "================ $Title ================"
     
     Write-Host "1: Exchange"
     Write-Host "2: Sharepoint"
     Write-Host "3: Skype"
     Write-Host "4: Common"
     Write-Host "Q: Press 'Q' to quit."
}

Do
{
     Show-Menu
     $input = Read-Host "Please make a selection"
     switch ($input)
     {
           '1' {
                cls
                Retrieve-Ips("Exchange")
           } '2' {
                cls
                Retrieve-Ips("Sharepoint")
           } '3' {
                cls
                Retrieve-Ips("Skype")
           } '4' {
                cls
                Retrieve-Ips("Common")
           }'q' {
                return
           }
     }
     pause
}
until ($input -eq 'q')



