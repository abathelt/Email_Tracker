#Credentials are email address and AD password
$UserCredential = Get-Credential
$date = Get-Date
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session -DisableNameChecking

Echo " "
Echo "[START] ***************************************"
#informations needed
$Email = Read-Host -Prompt 'Input the user email address'

Echo "[Info] Search is limited to 7 days"
$StartDate = Read-Host -Prompt 'Input the starting date MM/DD/YYYY'
$EndDate = Read-Host -Prompt 'Input the ending date MM/DD/YYYY'

#06/06/2021

Get-MessageTrace -SenderAddress $Email -StartDate $StartDate –EndDate $EndDate | Export-Csv -Path "C:\temp\MessageTrack.csv”
Get-MessageTrace -SenderAddress $Email -Status Delivered | Export-Csv -Path "C:\temp\MessageTrack_Delivered.csv”

Echo "[Info] Results were exported to C:\Temp folder"

Remove-PSSession *

Echo "[END] ***************************************"