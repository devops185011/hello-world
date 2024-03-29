cls

$hostname=hostname

$pfxName="$hostname.pfx"

$fqdn=[System.Net.Dns]::GetHostByName($env:computerName).HostName

$ErrorActionPreference = "SilentlyContinue"

#############################################################################################

$server=$pfxname.split(".")[0]
#$ErrorActionPreference = "SilentlyContinue"
$pfxSrcPath="C:\temp\$hostname\$pfxName"
        
$sha2DestPath="C:\temp\$hostname\SCOM 2012R2\rootcert\SHA-2 Interm and root"

$pfxDestPath="C:\temp\$hostname\SCOM 2012R2\SupportTools\AMD64"

$log="C:\temp\$hostname\log.txt"
                
write-host "`n--------------------------------------------------------------------------------------------------`n"
               
cd $sha2DestPath
   
$dateandtime=get-date
write-output "`n********************************************************************************************************************"|out-file -FilePath $log -Append
write-output "`n#####################################################################################################################"|out-file -FilePath $log -Append
                
write-output "`n$dateandtime`n"|out-file -FilePath $log -Append  

if(test-path "C:\temp\$hostname\RemoteServerFiles")
{   remove-item -path "C:\temp\$hostname\RemoteServerFiles" -Force -Recurse
    Start-Sleep -Seconds 5                   
} 
$rootca=CERTUTIL -addstore -enterprise -f -v root "corpca2.cer"
if($? -eq $true)
{
    Write-Host "`nsuccessfully imported corpca2.cer`n" -ForegroundColor White -BackgroundColor blue
    write-output "`nsuccessfully imported corpca2.cer`n"|out-file -FilePath $log -Append 
}
else 
{
    Write-Host "`ncouldn't import corpca2.cer`n" -ForegroundColor White -BackgroundColor red
    write-output "`ncouldn't import corpca2.cer`n"|out-file -FilePath $log -Append 
}          

$ca1=CERTUTIL -addstore -enterprise -f -v CA "ncr-ent-ca1-265.cer"
if($? -eq $true)
{
    Write-Host "`nsuccessfully imported ncr-ent-ca1-265.cer`n" -ForegroundColor White -BackgroundColor blue
    write-output "`nsuccessfully imported ncr-ent-ca1-265.cer`n"|out-file -FilePath $log -Append 
}
else 
{
    Write-Host "`ncouldn't import ncr-ent-ca1-265.cer`n" -ForegroundColor White -BackgroundColor red
    write-output "`ncouldn't import ncr-ent-ca1-265.cer`n"|out-file -FilePath $log -Append 
}

$ca2=CERTUTIL -addstore -enterprise -f -v CA "ncr-ent-ca2-256.cer"
if($? -eq $true)
{
    Write-Host "`nsuccessfully imported ncr-ent-ca2-256.cer`n" -ForegroundColor White -BackgroundColor blue
    write-output "`nsuccessfully imported ncr-ent-ca2-256.cer`n"|out-file -FilePath $log -Append
}

else
{
    Write-Host "`ncouldn't import ncr-ent-ca2-256.cer`n" -ForegroundColor White -BackgroundColor red
    write-output "`ncouldn't import ncr-ent-ca2-256.cer`n"|out-file -FilePath $log -Append 
}
                             
write-host "`n--------------------------------------------------------------------------------------------------`n"      
# end of else condition, checking sha-2 existance in source path
## checking pfx source path and destination path exist or not

if (!(Test-Path $pfxSrcPath) ) 
{ 
    write-host "Pfx file doesn't exist in source path..Please check" -ForegroundColor White -BackgroundColor Red
    write-output "`nPfx file doesn't exist in source path..Please check`n"|out-file -FilePath $log -Append
    break
}

else
{
    If(!(test-path $pfxDestPath))
    {
        write-host "$pfxDestPath does not exist" -ForegroundColor White -BackgroundColor Red
       write-output "`n$pfxDestPath does not exist`n"|out-file -FilePath $log -Append
        break
    }
    else
    {
        copy-item "$pfxSrcPath" "$pfxDestPath" -Force -ErrorAction SilentlyContinue
        $mypwd = ConvertTo-SecureString -String "!golf2day" -Force -AsPlainText
                  
        cd $pfxDestPath
        $pfxDetails= cmd.exe /c CERTUTIL -f -p "!golf2day" -importpfx $pfxName
        if($? -eq $true)
        {
            write-host "`nSuccessfully installed pfx file for $server, below are the details" -ForegroundColor White -BackgroundColor blue
            write-output "`nSuccessfully installed pfx file for $server, below are the details`n"|out-file -FilePath $log -Append
            write-host "`n $pfxDetails `n" 
            write-output "`n$pfxDetails`n"|out-file -FilePath $log -Append
            write-host "`n--------------------------------------------------------------------------------------------------`n"
            cd $pfxDestPath
            $momcertImport= cmd.exe /c MOMcertimport.exe $pfxName /Password '!golf2day'

            if($? -eq $true)
            {
                write-host "`n$momcertImport" -ForegroundColor White -BackgroundColor blue
                write-output "`n$momcertImport`n"|out-file -FilePath $log -Append
            }
            else
            {
              write-host "`n$momcertImport" -ForegroundColor White -BackgroundColor red
             write-output "`n$momcertImport`n"|out-file -FilePath $log -Append
             break
            }
            Start-Sleep -Seconds 5
            $startType=get-service -DisplayName 'Microsoft Monitoring Agent'|select starttype|%{$_.StartType}
        
            if($startType -eq "Disabled")
            {
                Set-Service  -name "HealthService" -StartupType Automatic
                Start-Sleep 5
                Get-Service -Name HealthService |Start-Service 
                Start-Sleep 5

                if($? -eq $true)
                {
                write-host "`nAgent service was disabled, enabling it" -ForegroundColor Cyan
                write-host "`n`n Successfully started  microsoft monitoring agent service for $server" -ForegroundColor White -BackgroundColor blue
                write-output "`n Agent service was disabled, enabling it"|out-file -FilePath $log -Append
                write-output "`n`n Successfully started  microsoft monitoring agent service for $server"|out-file -FilePath $log -Append
                }
                else
                {
                write-host "`n`n Couldn't start microsoft monitoring agent service for $server" -ForegroundColor White -BackgroundColor red
                write-output "`n`n Couldn't start microsoft monitoring agent service for $server"|out-file -FilePath $log -Append
                }
            }

            else
            {
              Restart-Service -DisplayName "microsoft monitoring agent" -Force
                if($? -eq $true)
                {
                write-host "`n`n Successfully restarted  microsoft monitoring agent service for $server" -ForegroundColor White -BackgroundColor blue
                write-output "`n`n Successfully restarted  microsoft monitoring agent service for $server"|out-file -FilePath $log -Append
                }
                else
                {
                write-host "`n`n Couldn't restart microsoft monitoring agent service for $server" -ForegroundColor White -BackgroundColor red
                write-output "`n`n Couldn't restart microsoft monitoring agent service for $server"|out-file -FilePath $log -Append
                }
            }
                        
                        
        $ncr="$hostname.ncr.com"
        $gsm="$hostname.gsm.ncr.com"

        if($fqdn -eq $ncr)
        {
            Copy-Item -Path "\\$hostname\c$\temp\$hostname\$pfxname" -Destination "\\susday321\e$\certs\$pfxname" -ErrorAction SilentlyContinue
            if($? -eq $true)
            {
                Write-host "`nBackup of $pfxname stored to \\susday321\e$\certs\$pfxname" -ForegroundColor white -BackgroundColor blue
                Write-Output "`nBackup of $pfxname stored to \\susday321\e$\certs\$pfxname"|out-file -FilePath $log -Append
                
            }                
        }
        elseif($fqdn -eq $gsm)
        {
            Copy-Item -Path "\\$hostname\c$\temp\$hostname\$pfxname" -Destination "\\susday4933\e$\certs\$pfxname" 
            if($? -eq $true)
            {
                 Write-host "`nBackup of $pfxname stored to \\susday4933\e$\certs\$pfxname" -ForegroundColor white -BackgroundColor blue
                 Write-Output "`nBackup of $pfxname stored to \\susday4933\e$\certs\$pfxname"|out-file -FilePath $log -Append  
            }
        }

        else
        {
            Copy-Item -Path "\\$hostname\c$\temp\$hostname\$pfxname" -Destination "\\susday321\e$\certs\$pfxname" 
            if($? -eq $true)
            {
                Write-host "`nBackup of $pfxname stored to \\susday321\e$\certs\$pfxname" -ForegroundColor white -BackgroundColor blue
                Write-Output "`nBackup of $pfxname stored to \\susday321\e$\certs\$pfxname"|out-file -FilePath $log -Append
            }
        }
                        
                       
                        
    } #end of if cheking pfx file installed or not

    else 
    {
        write-host "pfx file couldn't be installed for $server" -ForegroundColor White -BackgroundColor red
    }


 }##end of else block checking pfx Destination file exist or not

 } #end of else block checking pfx Source file exist or not
        



 # If running in the console, wait for input before closing.
if ($Host.Name -eq "ConsoleHost")
{
    Write-Host "Press any key to exit..."
    $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyUp") > $null
}

