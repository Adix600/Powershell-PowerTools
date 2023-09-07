#Author Adrian Hempel

Write-Host -f Magenta ("Wczytywanie...");
import-module ActiveDirectory
import-module ExchangeOnlineManagement
Add-Type -AssemblyName System.Windows.Forms

function Get-FileName()
{
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    $OpenFileDialog.filter = "SpreadSheet (*csv)|*.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.FileName
    $OpenFileDialog.Title = "Wybierz plik z listą użytkowników"
}

function get-sanitizedUTF8Input{
    Param(
        [String]$inputString
    )
    $replaceTable = @{"ą"="a";"ć"="c";"ę"="e";"ł"="l";"ń"="n";"ó"="o";"ś"="s";"ż"="z";"ź"="z"}

    foreach($key in $replaceTable.Keys){
        $inputString = $inputString -Replace($key,$replaceTable.$key)
    }
    $inputString = $inputString -replace '[^a-zA-Z0-9]', ''
    return $inputString
}


$users = @();
Write-Host -f Magenta ("Wczytywanie zakończone!");
Write-Host -f Magenta ("Łączę z Exchange Online...");
Connect-ExchangeOnline
Write-Host -f Magenta ("Połączono!");
Write-Host -f Magenta ("Pobieram dane użytkowników...");
$exusers = Get-Mailbox -ResultSize unlimited | Select-Object DisplayName, @{Name="EmailAddresses";Expression={($_.EmailAddresses | Where-Object {$_ -clike "SMTP*"} | ForEach-Object {$_ -replace "smtp:",""}) -join ","}} | Sort-Object DisplayName
Write-Host -f Magenta ("Dane Pobrane!");
Write-Host -f Magenta ("Oczytuję plik csv...");
$cusers = import-csv -Path (Get-FileName) -Delimiter ';' | Sort-Object -Property Surname
Write-Host -f Magenta ("Plik odczytany!");

foreach ($cuser in $cusers){
$found = $false
   foreach ($exuser in $exusers){
        $username=$cuser.Name+$cuser.Surname
        $username=$username.replace(' ' , '').replace('-' , '').ToLower()
        $exusername= ($exuser.DisplayName.Split("|")[0]).replace(' ' , '').replace('-' , '').ToLower()
        $username = get-sanitizedUTF8Input($username)
        $exusername = get-sanitizedUTF8Input($exusername)
        #write-host($username, $exusername)
        if(($username -eq $exusername)-or($username.contains($exusername))-or($exusername.contains($username))){
            $found = $true
            $users += $exuser.EmailAddresses
        }
   }
   if($found){
        write-host -f Green ("Znaleziono: $found ",$cuser.Name," ",$cuser.surname)
   }
   else{
        write-host -f Red ("Znaleziono: $found ",$cuser.Name," ",$cuser.surname)
   }
   
}

$DL = Read-Host("Podaj ID listy dystrybucyjnej ")
Write-Host -f Magenta ("Przetwarzam dane...!");

$dlusers =  Get-DistributionGroupMember -Identity $DL -ResultSize Unlimited | Select -Expand PrimarySmtpAddress | Sort-Object
$deletelist = @()
$addlist = @()

foreach($user in $users){
$isnotmember = $true
    foreach($dluser in $dlusers){
        $username=$user.Split("@")[0]
        $dlusername=$dluser.Split("@")[0]
        if($username -eq $dlusername){
            Write-host -f Yellow ("Użytkownik należy już do listy dystrybucyjnej:",$user)
            $isnotmember = $false
        }      
    }
    #Add-DistributionGroupMember –Identity $GroupEmailID -Member $user.userprincipalname
    if($isnotmember){
        $addlist += $user
    }  
}

foreach($dluser in $dlusers){
    $delete = $true
    foreach($user in $users){
        $username=$user.Split("@")[0]
        $dlusername=$dluser.Split("@")[0]
        if($username -eq $dlusername){
            $delete = $false
        }      
    }
    if($delete){
        $deletelist += $dluser
    }

}

Write-Host -f Magenta ("Przetwarzanie zakończone!");

$deletelist | Out-GridView -OutputMode Multiple -Title "Użytkowicy do usunięcia z listy dystrybucyjnej" | foreach{Remove-DistributionGroupMember -Identity $DL -Member $_;Write-Host -f Green ("Usunięto:",$_)}
$addlist | Out-GridView -OutputMode Multiple -Title "Użytkowicy do dodania do listy dystrybucyjnej" | foreach{Add-DistributionGroupMember -Identity $DL -Member $_; Write-Host -f Green ("Dodano:",$_)}
