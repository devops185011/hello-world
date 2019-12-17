cls

if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    [System.Windows.Forms.Messagebox]::Show("Not running as administrator!, To proceed Run as admin");

    exit
}

$qlook=read-host "Enter your QuickLook id e.g. - ku185011"

$qid=Get-childItem -path "c:\"|?{$_.name -match $qlook}|select name|%{$_.Name}
if($qlook -eq $qid)
{


Write-Host "`n$qlook folder exist`n" -ForegroundColor White -BackgroundColor Blue
}

else
{
write-host "`nNo such folder exist - C:\$qlook" -ForegroundColor White -BackgroundColor red

if ($Host.Name -eq "ConsoleHost")
{
   Write-Host -NoNewLine "`n`nPress any key to exit..." -ForegroundColor Cyan; 
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
exit


}

}


##############################################################################################################################

$path="C:\$qlook"

$serverFile="$path\serverlist.txt"
##$serverFileBackup="$path\serverlistBackup.txt"

$backup = "$path\Result"


$dateandtime=get-date

$logname=(Get-Date).ToString("dd-MM-yyyy")

$certslog="C:\Certificate_generation_and_installtion\certslog.txt"


If(!(test-path "$path\logs\$logname.txt"))
{
   $logpath=New-Item -ItemType Directory -Force -Path "$path\logs"
}

$log="$path\logs\$logname.txt"


If(!(test-path $backup))
{
   $backupPath=New-Item -ItemType Directory -Force -Path $backup
}
###############################################################################################################################



$infContent='[NewRequest]
Subject = "cn=test,ou=ITS,o=NCR Corporation,l=Dayton,st=Ohio,c=US"
KeySpec = 1
KeyLength = 2048
Exportable = TRUE
MachineKeySet = TRUE
KeyUsage = 0xF0
HashAlgorithm = SHA256

[EnhancedKeyUsageExtension]
OID=1.3.6.1.5.5.7.3.1
OID=1.3.6.1.5.5.7.3.2'


$src="$backup\test.inf"

Set-Content -Path $src -Value $infContent -ErrorAction SilentlyContinue

$ErrorActionPreference="SilentlyContinue"



$cases=read-host "Select your choice
                   1. Select 1, if you want to create a req file
                   2. select 2, if you want to accept certificate and generate a pfx file
                   "
switch ($cases)
{
1{
    $confirmation = Read-Host "`nAre you Sure You Want To Proceed [y/n]:"
    if ($confirmation -eq 'y' -or $confirmation -eq 'Y' -or $confirmation -eq 'yes' -or $confirmation -eq 'Yes' -or $confirmation -eq 'YES') 
    { 
                
                 write-output "`n********************************************************************************************************************"|out-file -FilePath $log -Append
                 write-output "`n#####################################################################################################################"|out-file -FilePath $log -Append
                
                 write-output "`n$dateandtime`n"|out-file -FilePath $log -Append
       

        
          # Move-Item "$path\*" "C:\SCOM2012R2_Agent_Package_with_CU5" -Exclude serverlist.txt,log.txt,generate-cert.bat,Scripts,RemoteServerFiles,Result -force -ErrorAction SilentlyContinue 
       
       
       
        copy-item "$backup\*" "C:\SCOM2012R2_Agent_Package_with_CU5" -Recurse -force -ErrorAction SilentlyContinue

              


        $req="$backup\req"

        if(!(test-path $req))
        {
            New-Item -ItemType Directory -Force -Path $req
        }
        else
        {
             Move-Item -Path "$req\*" -Destination "$backup" -Force -ErrorAction SilentlyContinue
        }
        
        
        (Get-Content $serverFile) | Foreach-Object {  $_ -replace $SearchString, $Replacement } | Set-Content $serverFile

        $fileNames=Get-Content -path "$serverFile" 

        foreach($tmp in $fileNames)
        {
                
                
                write-host "`n=============================================================================================="
                
                
                write-host "`nserver:$tmp" -ForegroundColor black -BackgroundColor white
                
                write-output "----------------------------------------------------------"|out-file -FilePath $log -Append
                write-output "`nserver:$tmp`n"|out-file -FilePath $log -Append
            $temp=$tmp.split(".")[0]
            $dest="$backup\$temp.inf"
            $reqFile="$req\$temp.req"
            Copy-Item "$src" "$dest" -Force
            
            (Get-Content -path $dest -Raw) -replace 'test',"$tmp" |Set-Content $dest -Force

            if($? -eq $true)
            {
                write-host "`n$temp.inf created" -ForegroundColor White -BackgroundColor blue
                
                write-output "`n$temp.inf created"|out-file -FilePath $log -Append
                
                certreq -new "$dest" "$reqFile"|out-null
        
                if($? -eq $true)
                {
                     write-host "$temp.req created" -ForegroundColor White -BackgroundColor blue
                     Write-Output "$temp.req created"|out-file -FilePath $log -Append
                }
                 

            }
            




            copy-item "$dest" "C:\SCOM2012R2_Agent_Package_with_CU5" -Recurse -force -ErrorAction SilentlyContinue

                 if($? -eq $true)
                 {
                 
                write-host "`nbackup of $temp.inf stored to C:\SCOM2012R2_Agent_Package_with_CU5 "
                
                write-output "backup of $temp.inf stored to C:\SCOM2012R2_Agent_Package_with_CU5"|out-file -FilePath $log -Append
                 
                
                 }
            
              copy-item "$reqFile" "C:\SCOM2012R2_Agent_Package_with_CU5" -Recurse -force -ErrorAction SilentlyContinue

                 if($? -eq $true)
                 {
                 
                write-host "backup of $temp.req stored to C:\SCOM2012R2_Agent_Package_with_CU5 " 
                
                
                write-output "backup of $temp.req stored to C:\SCOM2012R2_Agent_Package_with_CU5"|out-file -FilePath $log -Append
                                 
                 }





        }  ###end of for loop
         
        $date=(Get-Date).ToString("dd-MM-yyyy_HHmmss")
        $certName="certs.zip"
        $attachments="$backup\$date$certName"
        $source = "$req"
        $archive = "$attachments"

        Add-Type -assembly "system.io.compression.filesystem"
        [io.compression.zipfile]::CreateFromDirectory($source, $archive)
    
        
        write-host "`n=============================================================================================="

        $choice=read-host "`nSelect your choice
                   a. Select a, if you want to send a test mail to yourself
          
                   b. select b, if you want to send mail to NCR PKI
                   "
        switch ($choice)
        {
        a{
            $smtp = "relay1.daytonoh.ncr.com"
            $to = "$qlook@ncr.com"
            $from = "ma230070@ncr.com"
            $subject = "Certificate request" 
            $body= @"
             Hi Team,
             Please use SCOM template to generate the required certificates for attached request.
             
             PS:- Reply All
            

             Thanks,
             SCOM team
             ($qlook)
"@
             send-MailMessage -SmtpServer $smtp -To $to -From $from -Subject $subject -Body $body  -Attachments $attachments 

             if($? -eq $true)
             {
                write-host "`n$qlook sent the TEST mail to $to for cert req" -ForegroundColor White -BackgroundColor blue

            
                
                Write-Output "`n$qlook sent the TEST mail to $to for cert req"|out-file -FilePath $log -Append
                
                Write-Output "==============================================================="|out-file -FilePath $log -Append
             }
             else
             {
                write-host "`ntest mail not sent" -ForegroundColor White -BackgroundColor red
                
                
                Write-Output "`ntest mail not sent"|out-file -FilePath $log -Append
                
                
                Write-Output "==============================================================="|out-file -FilePath $log -Append
             }
    
             break 
           } #end of choice a

        b{
            $smtp = "relay1.daytonoh.ncr.com"
            $to=$null
            $to = "NCR, PKI <PKI.NCR@ncr.com>"
            $from = "ma230070@ncr.com"
            $subject = "Certificate request" 
            $cc=@("ss185240@ncr.com","SB185328@ncr.com","Cecilia.Slonaker@ncr.com","ku185011@ncr.com")
            $body= @"
             Hi Team,

             Please use SCOM template to generate the required certificates for attached request.
            
             PS:- Reply All
             

             Thanks,
             SCOM team
             ($qlook)
"@
             send-MailMessage -SmtpServer $smtp -To $to -From $from -cc $cc -Subject $subject -Body $body  -Attachments $attachments 

             if($? -eq $true)
              {
                write-host "`n $qlook sent the mail to PKI.NCR@ncr.com for cert req" -ForegroundColor White -BackgroundColor blue
                 Write-Output " $qlook sent the mail to PKI.NCR@ncr.com for cert req"|out-file -FilePath $log -Append

            
                Write-Output "==============================================================="|out-file -FilePath $log -Append

              }
             else
              {
                 write-host "`n mail not sent to  $to " -ForegroundColor White -BackgroundColor red
                         Write-Output " mail not sent to  $to"|out-file -FilePath $log -Append
                         
                Write-Output "==============================================================="|out-file -FilePath $log -Append
         
              }
              

             break
             
           }##end of choice b

        
        default {write-host "Something else happened" -ForegroundColor white -BackgroundColor red; break}

        }#end of switch choice
          

            
        } #end of confirmation block

    break
    }  #end of 1 switch

   
   
    2{
     write-output "`n********************************************************************************************************************"|out-file -FilePath $log -Append
                
                 write-output "`n$dateandtime`n"|out-file -FilePath $log -Append

    
    $certs=Get-ChildItem "$path\*.cer"|select name|%{$_.name}|out-file "$backup\certs.txt" -Force

               
            $servers=Get-Content "$backup\certs.txt"


             if(($servers -eq $null))
                {

                
                write-host "`nTo Proceed please put certificates in path $path" -ForegroundColor Yellow
                
                write-output "`nTo Proceed please put certificates in path $path"|out-file -FilePath $log -Append
                            
                
                write-host "`n=============================================================================================="
                break;
                }

            foreach ($tmp in $servers)
            {
                $temp=$tmp.split(".")[0]

                
                $tempCer="$temp.cer"
                
                
                write-host "`n=============================================================================================="
                
                
                write-host "`nserver:$temp" -ForegroundColor black -BackgroundColor white
                
                write-output "----------------------------------------------------------------------------------"|out-file -FilePath $log -Append
                write-output "`nserver:$temp`n"|out-file -FilePath $log -Append
                
                $cer=Get-ChildItem -path "$path\*.cer"|?{$_.name -match $tempCer}|select name|%{$_.Name}

                
               
                cd $path

                $cereq=certreq -accept -q "$temp.cer" 
                


              if($? -eq $true)
              {
                 write-host "`nSuccessfully accepted certificate for server $temp by $qlook`n" -ForegroundColor Green

                 
                 write-output "`nSuccessfully accepted certificate for server $temp by $qlook"|out-file $certslog -Append

                 

                 Write-Output "`nSuccessfully accepted certificate for server $temp by $qlook"|out-file -FilePath $log -Append
         
                    write-host "`n pfx file creation in PROGRESS...please wait `n" -ForegroundColor Cyan

 
                  
                    
                    $remotePfx="$path\$temp"
                       
                   if(test-path $remotePfx)
                   {  
                       remove-item -LiteralPath $remotePfx -Force -Recurse -ErrorAction SilentlyContinue
                       Start-Sleep -Seconds 5
                       $newdir=New-Item -ItemType Directory -Force -Path $remotePfx -ErrorAction SilentlyContinue
                   }
                         $source="C:\Certificate_generation_and_installtion\RemoteServerFiles"
                        
                    
                        $source1="C:\Certificate_generation_and_installtion\RemoteServerFiles\SCOM 2012R2"
                        
                        $source2="C:\Certificate_generation_and_installtion\RemoteServerFiles\cert-install.ps1"
                        
                        $source3="C:\Certificate_generation_and_installtion\RemoteServerFiles\cert-install.txt"

                        $destination="$remotePfx"
                        cp -Recurse $source $destination
                        
                        cp -Recurse $source1 $destination
                        cp -Recurse $source2 $destination
                        cp -Recurse $source3 $destination





                        if(test-path "$remotePfx\RemoteServerFiles")
                      {  
                       remove-item -path "$remotePfx\RemoteServerFiles" -Force -Recurse
                       Start-Sleep -Seconds 5
                      }
                        

                      
                Move-Item "$path\$tempCer" "C:\SCOM2012R2_Agent_Package_with_CU5" -force -ErrorAction SilentlyContinue
                
                if($? -eq $true)
                {
                    write-host "`nBackup of $tempcer stored to C:\SCOM2012R2_Agent_Package_with_CU5 " -ForegroundColor white -BackgroundColor blue

                    Write-Output "`nBackup of $tempcer stored to C:\SCOM2012R2_Agent_Package_with_CU5"|out-file -FilePath $log -Append
         
                }
                else
                {
                    write-host "`nBackup of $tempcer couldn't be stored to C:\SCOM2012R2_Agent_Package_with_CU5 " -ForegroundColor white -BackgroundColor red
                    write-output "`nBackup of $tempcer couldn't be stored to C:\SCOM2012R2_Agent_Package_with_CU5 "|out-file -FilePath $log -Append

                }
                        
                    $mypwd = ConvertTo-SecureString -String "!golf2day" -Force -AsPlainText

                    $cn="$temp"

                 Get-ChildItem -Path cert:\localMachine\my|?{$_.Subject -like "*$cn*"}|sort-object|select subject, Thumbprint,NotAfter,NotBefore -First 1

                $sub=Get-ChildItem -Path cert:\localMachine\my|?{$_.Subject -like "*$cn*"}|sort-object|select subject -First 1|%{$_.subject}

                 $sub1=$sub.split(",")[0];

                $sub2=$sub1.Split(".")[0]

                $nameofPfx=$sub2.Split("=")[1]
                 
                 if($nameofPfx -like "*$temp*")

                 {
                 $thumbprint=Get-ChildItem -Path cert:\localMachine\my|?{$_.Subject -like "*$cn*"}|sort-object| select Thumbprint -First 1|%{$_.Thumbprint}
                    
                 $pfxFile=Export-PfxCertificate -Cert "cert:\localMachine\my\$thumbprint" -FilePath "$remotePfx\$temp.pfx" -ChainOption EndEntityCertOnly -NoProperties -Password $mypwd -Force 
                 
                                     
                    if($?)
                    { write-host "`npfx file created for server $temp and  is stroed in $remotepfx\$temp.pfx" -ForegroundColor White -BackgroundColor blue
                        

                     Write-Output "`npfx file created for server $temp and Result is stroed in $remotepfx\$temp.pfx"|out-file -FilePath $log -Append
         
         
                        
                        copy-item "$remotepfx\$temp.pfx" "C:\SCOM2012R2_Agent_Package_with_CU5" -Recurse -force -ErrorAction SilentlyContinue
                       
                      
                        if($? -eq $true)
                        {

                            write-host "`nBackup of $temp.pfx is stored to C:\SCOM2012R2_Agent_Package_with_CU5 " -ForegroundColor white -BackgroundColor blue
                            
                     Write-Output "`nBackup of $temp.pfx is stored to C:\SCOM2012R2_Agent_Package_with_CU5"|out-file -FilePath $log -Append
                
                        }

                        else
                        {
                        
                          write-host "`n Backup of $temp.pfx couln't be stored to C:\SCOM2012R2_Agent_Package_with_CU5 " -ForegroundColor white -BackgroundColor red
                          Write-Output "`nBackup of $temp.pfx couln't be stored to C:\SCOM2012R2_Agent_Package_with_CU5"|out-file -FilePath $log -Append
                
                       
                
                        }
                          
                          
                        (Get-Content -path "$remotepfx\cert-install.txt" -Raw) -replace 'test',"$temp" |Set-Content "$remotepfx\cert-install.txt" -Force
                        Rename-Item -Path "$remotepfx\cert-install.txt" -NewName "$remotepfx\cert-install.bat"
                         $ncr="$temp.ncr.com"
                          $gsm="$temp.gsm.ncr.com"

                          if($tmp -eq $ncr)
                          {

                            
                            
                            
                            Copy-Item -Path "$path\$temp\$temp.pfx" -Destination "\\susday321\e$\certs\$temp.pfx" -ErrorAction SilentlyContinue

                            if($? -eq $true)
                            {
                                Write-host "`nBackup of $temp.pfx stored to \\susday321\e:\certs\$temp.pfx" -ForegroundColor white -BackgroundColor blue

                                
                             Write-Output "`nBackup of $temp.pfx stored to \\susday321\e:\certs\$temp.pfx"|out-file -FilePath $log -Append
                
                            }
                            

                          }
             
                         elseif($tmp -eq $gsm)
                          {
                          
                            
                            
                            
                             Copy-Item -Path "$path\$temp\$temp.pfx"-Destination "\\susday321\e$\certs\$temp.pfx" -ErrorAction SilentlyContinue
                            if($? -eq $true)
                            {
                                Write-host "`nBackup of $temp.pfx stored to \\susday321\e:\certs\$temp.pfx" -ForegroundColor white -BackgroundColor blue
                           
                
                             Write-Output "`nBackup of $temp.pfx stored to \\susday321\e:\certs\$temp.pfx"|out-file -FilePath $log -Append
                
                            }
             
                          }

                         else
                          {
                          
                
                            
                            
                             Copy-Item -Path "$path\$temp\$temp.pfx" -Destination "\\susday321\e$\certs\$temp.pfx" -ErrorAction SilentlyContinue

                            if($? -eq $true)
                            {
                            Write-host "`nBackup of $temp.pfx stored to \\susday321\e:\certs\$temp.pfx" -ForegroundColor white -BackgroundColor blue
                            
                
                            
                             Write-Output "`nBackup of $temp.pfx stored to \\susday321\e:\certs\$temp.pfx"|out-file -FilePath $log -Append
                
                            }
             
                          }
                                              

                     }
                    else
                    {
                        write-host "`npfx file can't be created for server :$temp" -ForegroundColor  White  -BackgroundColor Red
                        Write-Output "`npfx file can't be created for server :$temp"|out-file -FilePath $log -Append
                
                
                write-host "`n==============================================================================================" 
                       
                    }       
                
                } ###end of if checking if pfx file name is correct or not
                else
                {
                
                write-host "`nsome problem with pfx file please check :$temp" -ForegroundColor  White  -BackgroundColor Red
                        Write-Output "`nsome problem with pfx file please check :$temp"|out-file -FilePath $log -Append
                
                
                }
                
            }###end of if cheking certificate already accepted or not
            
            else
            {
               
               $confirmcert=Get-ChildItem -Path cert:\localMachine\my|?{$_.Subject -match $temp}|select Thumbprint, subject,NotAfter,NotBefore -Last 1|ft

               if(!($confirmcert))

               {
               
               write-host "`nSome Problem with certificate $temp.cer, Can't install...Please check`n" -ForegroundColor white -BackgroundColor red

               
                write-output "`n `nSome Problem with certificate $temp.cer, Can't install...Please check`n"|out-file -FilePath $log -Append

                continue
               }

                $userNam=$null
                $usern=$null

               $userNam=select-string -path $certslog -Pattern "$temp" | select line|%{$_.line}|select -Last 1
                $usern=$userNam.split()[-1]

                if($usern)
                {
                 write-host "`nCertifcate for $temp is already accepted by $usern`n" -ForegroundColor green
 
                write-output "`n `nCertifcate for $temp is already accepted by $usern`n"|out-file -FilePath $log -Append
                }
                else
                {
                
                 write-host "`nCertifcate for $temp is already accepted " -ForegroundColor green
 
                write-output "`n `nCertifcate for $temp is already accepted"|out-file -FilePath $log -Append
                
                }
                

               $confirm1 = Read-Host "`nDo you still want to create pfx [y/n]:"
                if ($confirm1 -eq 'y' -or $confirm1 -eq 'Y' -or $confirm1 -eq 'yes' -or $confirm1 -eq 'Yes' -or $confirm1 -eq 'YES') 
                {
                     
                    write-host "`n pfx file creation in PROGRESS...please wait `n"  -ForegroundColor Cyan
                    
                    $remotePfx="$path\$temp"
                       
                   if(test-path $remotePfx)
                   {  
                       remove-item -LiteralPath $remotePfx -Force -Recurse -ErrorAction SilentlyContinue
                       Start-Sleep -Seconds 5
                       $newdir=New-Item -ItemType Directory -Force -Path $remotePfx -ErrorAction SilentlyContinue
                   }
                    
                    $source="C:\Certificate_generation_and_installtion\RemoteServerFiles"
                        
                    
                        $source1="C:\Certificate_generation_and_installtion\RemoteServerFiles\SCOM 2012R2"
                        
                        $source2="C:\Certificate_generation_and_installtion\RemoteServerFiles\cert-install.ps1"
                        
                        $source3="C:\Certificate_generation_and_installtion\RemoteServerFiles\cert-install.txt"

                        $destination="$remotePfx"
                        cp -Recurse $source $destination
                        
                        cp -Recurse $source1 $destination
                        cp -Recurse $source2 $destination
                        cp -Recurse $source3 $destination
                     if(test-path "$remotePfx\RemoteServerFiles")
                      {  
                       remove-item -path "$remotePfx\RemoteServerFiles" -Force -Recurse
                       Start-Sleep -Seconds 5
                       
                    
                      }
                        
                              
                Move-Item "$path\$tempCer" "C:\SCOM2012R2_Agent_Package_with_CU5" -force -ErrorAction SilentlyContinue
                
                if($? -eq $true)
                {
                    write-host "`nBackup of $tempcer stored to C:\SCOM2012R2_Agent_Package_with_CU5 " -ForegroundColor white -BackgroundColor blue

                    Write-Output "`nBackup of $tempcer stored to C:\SCOM2012R2_Agent_Package_with_CU5"|out-file -FilePath $log -Append
         
                }
                else
                {
                    write-host "`nBackup of $tempcer couldn't be stored to C:\SCOM2012R2_Agent_Package_with_CU5 " -ForegroundColor white -BackgroundColor red
                    write-output "`nBackup of $tempcer couldn't be stored to C:\SCOM2012R2_Agent_Package_with_CU5 "|out-file -FilePath $log -Append

                }

                    $mypwd = ConvertTo-SecureString -String "!golf2day" -Force -AsPlainText
                    

                  $cn="$temp"

                 Get-ChildItem -Path cert:\localMachine\my|?{$_.Subject -like "*$cn*"}|sort-object|select subject, Thumbprint,NotAfter,NotBefore -First 1

                 $sub=Get-ChildItem -Path cert:\localMachine\my|?{$_.Subject -like "*$cn*"}|sort-object|select subject -First 1|%{$_.subject}

                 $sub1=$sub.split(",")[0];

                $sub2=$sub1.Split(".")[0]

                $nameofPfx=$sub2.Split("=")[1]
                 
                 
                 if($nameofPfx -like "*$temp*")


                 {
                 $thumbprint=Get-ChildItem -Path cert:\localMachine\my|?{$_.Subject -like "*$cn*"}|sort-object| select Thumbprint -First 1|%{$_.Thumbprint}
                    
                 $pfxFile=Export-PfxCertificate -Cert "cert:\localMachine\my\$thumbprint" -FilePath "$remotePfx\$temp.pfx" -ChainOption EndEntityCertOnly -NoProperties -Password $mypwd -Force 
                 
                                     
                    if($?)
                    { write-host "`npfx file created for server $temp and  is stroed in $remotepfx\$temp.pfx" -ForegroundColor White -BackgroundColor blue
                        

                     Write-Output "`npfx file created for server $temp and Result is stroed in $remotepfx\$temp.pfx"|out-file -FilePath $log -Append
         
         
                        
                        copy-item "$remotepfx\$temp.pfx" "C:\SCOM2012R2_Agent_Package_with_CU5" -Recurse -force -ErrorAction SilentlyContinue
                       
                      
                        if($? -eq $true)
                        {

                            write-host "`nBackup of $temp.pfx is stored to C:\SCOM2012R2_Agent_Package_with_CU5 " -ForegroundColor white -BackgroundColor blue
                            
                     Write-Output "`nBackup of $temp.pfx is stored to C:\SCOM2012R2_Agent_Package_with_CU5"|out-file -FilePath $log -Append
                
                        }

                        else
                        {
                        
                          write-host "`n Backup of $temp.pfx couln't be stored to C:\SCOM2012R2_Agent_Package_with_CU5 " -ForegroundColor white -BackgroundColor red
                          Write-Output "`nBackup of $temp.pfx couln't be stored to C:\SCOM2012R2_Agent_Package_with_CU5"|out-file -FilePath $log -Append
                
                       
                
                        }
                          
                          
                        (Get-Content -path "$remotepfx\cert-install.txt" -Raw) -replace 'test',"$temp" |Set-Content "$remotepfx\cert-install.txt" -Force
                        Rename-Item -Path "$remotepfx\cert-install.txt" -NewName "$remotepfx\cert-install.bat"
                         $ncr="$temp.ncr.com"
                          $gsm="$temp.gsm.ncr.com"

                          if($tmp -eq $ncr)
                          {

                            
                            
                            
                            Copy-Item -Path "$path\$temp\$temp.pfx" -Destination "\\susday321\e$\certs\$temp.pfx" -ErrorAction SilentlyContinue

                            if($? -eq $true)
                            {
                                Write-host "`nBackup of $temp.pfx stored to \\susday321\e:\certs\$temp.pfx" -ForegroundColor white -BackgroundColor blue

                                
                             Write-Output "`nBackup of $temp.pfx stored to \\susday321\e:\certs\$temp.pfx"|out-file -FilePath $log -Append
                
                            }
                            

                          }
             
                         elseif($tmp -eq $gsm)
                          {
                          
                            
                            
                            
                             Copy-Item -Path "$path\$temp\$temp.pfx"-Destination "\\susday321\e$\certs\$temp.pfx" -ErrorAction SilentlyContinue
                            if($? -eq $true)
                            {
                                Write-host "`nBackup of $temp.pfx stored to \\susday321\e:\certs\$temp.pfx" -ForegroundColor white -BackgroundColor blue
                           
                
                             Write-Output "`nBackup of $temp.pfx stored to \\susday321\e:\certs\$temp.pfx"|out-file -FilePath $log -Append
                
                            }
             
                          }

                         else
                          {
                          
                
                            
                            
                             Copy-Item -Path "$path\$temp\$temp.pfx" -Destination "\\susday321\e$\certs\$temp.pfx" -ErrorAction SilentlyContinue

                            if($? -eq $true)
                            {
                            Write-host "`nBackup of $temp.pfx stored to \\susday321\e:\certs\$temp.pfx" -ForegroundColor white -BackgroundColor blue
                            
                
                            
                             Write-Output "`nBackup of $temp.pfx stored to \\susday321\e:\certs\$temp.pfx"|out-file -FilePath $log -Append
                
                            }
             
                          }
                                              

                     }
                    else
                    {
                        write-host "`npfx file can't be created for server :$temp" -ForegroundColor  White  -BackgroundColor Red
                        Write-Output "`npfx file can't be created for server :$temp"|out-file -FilePath $log -Append
                
                
                write-host "`n==============================================================================================" 
                       
                    }       
                
               } ###end of if checking pfxfile name is correct or not
               
               else
                {
                
                write-host "`nsome problem with pfx file please check :$temp" -ForegroundColor  White  -BackgroundColor Red
                        Write-Output "`nsome problem with pfx file please check :$temp"|out-file -FilePath $log -Append
                
                
                }
                





                 }   ###end of if creating pfx when certificate is already accepted
                 else
                 {
                  write-host "`nskipping $tempcer" -ForegroundColor Cyan

                  Move-Item "$path\$tempCer" "C:\SCOM2012R2_Agent_Package_with_CU5" -force -ErrorAction SilentlyContinue
                
                     if($? -eq $true)
                    {
                      write-host "`nBackup of $tempcer stored to C:\SCOM2012R2_Agent_Package_with_CU5 " -ForegroundColor white -BackgroundColor blue

                         Write-Output "`nBackup of $tempcer stored to C:\SCOM2012R2_Agent_Package_with_CU5"|out-file -FilePath $log -Append
         
                    }
                    else
                    {
                         write-host "`nBackup of $tempcer couldn't be stored to C:\SCOM2012R2_Agent_Package_with_CU5 " -ForegroundColor white -BackgroundColor red
                         write-output "`nBackup of $tempcer couldn't be stored to C:\SCOM2012R2_Agent_Package_with_CU5 "|out-file -FilePath $log -Append

                    }
                
                    continue
                 
                 } #end of els eblock skipping if certificate already accepted
            
            }###end of else block cheking certificate already accepted or not
           

              } #end of for loop

           ###end of confirmation block
         break;
    }     #end of 2 switch


default {write-host "Something else happened" -ForegroundColor white -BackgroundColor red; break}

}






 # If running in the console, wait for input before closing.
if ($Host.Name -eq "ConsoleHost")
{
   Write-Host -NoNewLine "`n`nPress any key to exit..." -ForegroundColor Cyan; 
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');


}
