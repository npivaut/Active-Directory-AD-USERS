 <#
 Auteur: Pivaut Nicolas
 Date: 11/09/2022
 Description : script de création d’utilisateurs dans l'AD
 Version : 0.1 : Version initiale 
 Version : 0.2 : Ajout bloc try/catch
 #>




Import-Module ActiveDirectory
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
$dialog = New-Object System.Windows.Forms.OpenFileDialog
$dialog.InitialDirectory = $StartDir
$dialog.Filter = "CSV (*.csv)| *.csv" 
$dialog.ShowDialog() | Out-Null
$CSVFile = $dialog.FileName
if ([System.IO.File]::Exists($CSVFile)) {
    Write-Host "Importing CSV..."
    $CSV = Import-Csv -LiteralPath "$CSVFile"
} else {
    Write-Host "File path specified was not valid"
    Exit
}

  
foreach($user in $CSV) {

    try
{
   $EtatService = Get-Service -Name $NomService -ErrorAction Stop
   Write-Host -ForegroundColor Green "Etat du service correctement récupéré !"

    $SecurePassword = ConvertTo-SecureString "$($user.'First Name'[0])$($user.'Last Name')$($user.'Employee ID')!@#" -AsPlainText -Force

    $Username = "$($user.'First Name').$($user.'Last Name')"
    $Username = $Username.Replace(" ", "")


    New-ADUser -Name "$($user.'First Name') $($user.'Last Name')" `
                -GivenName $user.'First Name' `
                -Surname $user.'Last Name' `
                -UserPrincipalName $Username `
                -SamAccountName $Username `
                -EmailAddress $user.'Email Address' `
                -Description $user.Description `
                -OfficePhone $user.'Office Phone' `
                -Path "$($user.'Organizational Unit')" `
                -ChangePasswordAtLogon $true `
                -AccountPassword $SecurePassword `
                -Enabled $true
                # Write to host that we created a new user
    Write-Host "Created $Username / $($user.'Email Address')"

    }
catch
{
   Write-Host $_.Exception.Message -ForegroundColor Red
}
finally
{
   $Error.Clear()
}

Write-Host "Contenu des erreurs : $Error"

     
    # If groups is not null... groupe de securite
    if ($User.'Add Groups (csv)' -ne "") {
        $User.'Add Groups (csv)'.Split(",") | ForEach {
            Add-ADGroupMember -Identity $_ -Members "$($user.'First Name').$($user.'Last Name')"
            Write-Host "Added $Username to $_ group" # Log to console
        }
    }

    Write-Host "Created user $Username with groups $($User.'Add Groups (csv)')"
}
Read-Host -Prompt "Script complete... Press enter to exit."
