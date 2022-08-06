#special thanks to my dad who without which none of this would have been possible

# To execute the script without agreeing with the execution policy
Set-ExecutionPolicy Bypass -Scope Process
#Install-Module importexcel
# Defines the directory where the file is located
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12;
if (!(Get-Module -Name "importexcel ")) {
    write-host "Install module"
    Install-Module importexcel -Scope CurrentUser -AllowClobber -Force
   }
$dir = $PSScriptRoot
$run = 0
$new_dict = $dir +"\results"
if(Test-Path -Path $new_dict){
    continue
}else{
    New-Item -Path ($dir + "\results") -ItemType "directory"
    $user = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    Start-sleep -Seconds .25
}
<# $ACL = Get-ACL -Path $new_dict
$AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($user,"Read","Allow")
$ACL.SetAccessRule($AccessRule)
$ACL | Set-Acl -Path $new_dict
(Get-ACL -Path $new_dict).Access | Format-Table IdentityReference,FileSystemRights,AccessControlType,IsInherited,InheritanceFlags -AutoSize #>
#$dir = $PSScriptRoot
#Write-Output $dir
while ($run -eq 0) {

    $dir = $PSScriptRoot
    $store = Read-Host -Prompt "What store do you want info about?"
    $print_data = @(@())
    $color_data = @(@())
    $color_data += 0
    $color_data += 1
    $first = 0
    foreach($file in (Get-ChildItem $dir -Exclude *.ps1, *.txt, *.pdf, "results")) { #this part turns each file into a csv and then gets the data from it
        $newname = $file.FullName -replace '\.xls$', '.csv'
        $ExcelWB = new-object -comobject excel.application
        $Workbook = $ExcelWB.Workbooks.Open($file.FullName) 
        $Workbook.SaveAs($newname,6)
        $Workbook.Close($false)
        $ExcelWB.quit()
        Write-Output $newname
        Start-sleep -Seconds .25
        $a = @()
        #Start-sleep -Seconds 5
        $a = Get-Content $newname | Select-Object -First 5 | Select-Object -Last 1
        #Write-Output $a
        if($a -match $store){ #if the store matches the invoice
            $data = @()
            $data = Get-Content $newname | Select-Object -Skip 12
            #Write-Output $data[0]
            $temp = @()
            $counter = 0
            $data = Get-Content $newname | Select-Object -Skip 12
            For($k=0;$k -lt $data.Length; $k += 2){#get the length of the order
                #Get this sorted, something needs to be done so that the duplicate stores don't just add to the bottom of list
                $temp = $data[$k].split(',')
                if($temp[1] -eq "" -or $temp[1] -eq " "){
                    break
                }
                else{
                    $counter += 2
                }
            }
            $data = Get-Content $newname | Select-Object -Skip 12 #skips to the actual sale stuff

            For($i=0; $i -lt $counter; $i += 2){#populates a 2d array of everything that was in the csv
                $temp = $data[$i].split(",")
                if($first -eq 0){
                    if($i -eq 0){$print_data += 0}
                    if($i -eq 0){
                        $print_data += , $temp
                    }else{
                        $print_data += , $temp
                        $c_temp = $data[$i - 1].split(",")
                        $color_data += , $c_temp
                    }
                    if($i -eq ($counter - 2) -or $i -eq $counter){
                        $first = 1
                    }
                }
                else{
                    $found = 0
                    if($temp[1]-match "Style"){
                        continue
                    }
                    else{
                        if($i -eq 2){
                            $c_temp = $data[$i - 1].split(",")
                            Start-sleep -Seconds .1
                            #Write-Output $temp[1]
                            #Write-Output $print_data[2][1]
                            if($print_data[2][1] -match $temp[1]){
                                $found = 1
                                For($e=5;$e -le 12; $e++){
                                    if($temp[$e] -eq "" -or $temp[$e] -eq " "){
                                        continue
                                    }else{
                                        if($print_data[2][$e] -eq "" -or $print_data[2][$e] -eq " "){
                                            $print_data[$d][$e] = $temp[$e]
                                            
                                        }else{
                                            $inside = [int]$print_data[2][$e]
                                            $outside = [int]$temp[$e]
                                            $total = $inside + $outside
                                            $print_data[2][$e] = [string]$total

                                        }

                                    }

                                }
                            }

                        }else{
                            
                            For($d = 3;$d -lt $print_data.Length; $d++){
                                $c_temp = $data[$i - 1].split(",")
                                
                                if($print_data[$d][1] -match $temp[1]){
                                    $found = 1
                                    For($e=5;$e -le 12; $e++){
                                        if($temp[$e] -eq "" -or $temp[$e] -eq " "){
                                            continue
                                        }else{
                                            if($print_data[$d][$e] -eq "" -or $print_data[$d][$e] -eq " "){
                                                $print_data[$d][$e] = $temp[$e]
                                            }else{
                                                $inside = [int]$print_data[$d][$e]
                                                $outside = [int]$temp[$e]
                                                $total = $inside + $outside
                                                $print_data[$d][$e] = $total

                                            }
                                        }
                                    }
                                }
                            }
                        }
                        
                        if($found -eq 0){
                            $print_data += , $temp
                            $c_temp = $data[$i - 1].split(",")
                            $color_data += , $c_temp
                        }
                        else{
                            $found = 0
                        }
                    }
                    
                }
            }
                
        }

        }
        $dir = $PSScriptRoot
        foreach($file in (Get-ChildItem $dir  *.csv)) {
        $yay = "yay"
        Write-Output $file
        Remove-Item $file.FullName
        }
        $actual_print = @(@())

        $actual_color = @(@())
        For($i=2; $i -le ($color_data.Length - 1);$i++){
        $index_3 = 0
        $temp = @("10") * 10
        For($j=2;$j -lt 13; $j++){
            if($j -eq 4){
                continue
            }else{
                $temp[$index_3] = $color_data[$i][$j]
                $index_3 += 1
            }

        }
        $actual_color += , $temp
        $temp.Clear
        }
        #Write-Output $print_data.Length
        For($i=0; $i -le ($print_data.Length -1);$i++){
        $temp = @("10") * 10
        $index_2 = 0
        if($i -ge 1){
            For($j=1;$j -lt 13; $j++){
                #Write-Output $j
                if($j -eq 1 -or $j -eq 2){
                    $temp[$index_2] = $print_data[$i][$j]
                    #Write-Output $print_data[$i][$j]
                    $index_2 += 1
                }
                if($j -ge 5 -and $j -le 12){
                    #Write-Output $print_data[$i][$j]
                    $temp[$index_2] = $print_data[$i][$j]
                    $index_2 += 1
                }
            }
            if($i -ge 2){
                $actual_print += , $actual_color[$i - 2]
            }
            
            $actual_print += , $temp
            $temp.Clear
            }

        }
        #Write-Output $actual_print[1]
        $newname = $dir + "\results.csv"
        $outputfile = $dir + "\" + "results" + "\" + $store + "_" + "results.xls"
        #Write-Output $print_data[1]
        $actual_print | % { $_ -join ','} | Out-File $newname
        Import-CSV $newname | Export-Excel $outputfile
        #Import-CSV $inputfile | Export-Excel $outputfile
        Remove-Item $newname
        Write-Output "All Done!"
        $end_message = "Your data is saved in a new excel file called " + $outputfile
        Write-Output $end_message
        $again = Read-Host -Prompt "Do you want to do another store [Yes] or [No]?"
        if($again -match "yes"){
            $run = 0
        }else{
            $run = 1
        }
}