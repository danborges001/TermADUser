param(
[string]$Email
)
#Output for error possible debug message in perl script
Get-ADUser -filter {EmailAddress -like $Email}
#Set variables, perform terminate procedures on user account
$Username = Get-ADUser -filter {EmailAddress -like $Email}
$TermDate = Get-Date -UFormat "%Y-%m-%d"
Move-ADObject  $Username -TargetPath "OU=TermedUsers,OU=Actian Corporation,DC=Actian,DC=com"
Set-ADUser $Username.SAMAccountName -Clear Department,Title,telephoneNumber,Company,Manager,Mobile,HomePhone
Set-ADUser $Username.SAMAccountName -Replace @{Description="Termed "+ $TermDate}
Set-ADUser $Username.SAMAccountName -Enabled $False
$GroupArray = Get-ADGroup -property * -filter * | Where-Object {$_.member -like $Username}
ForEach ($Group in $GroupArray)
	{
	Remove-ADGroupMember -Identity $NewGroup -Member $Username.SAMaccountName -Confirm:$false
	}
#param(
#[string]$Email
#)
#Get-ADUser -filter {EmailAddress -like $Email} | Move-ADObject -TargetPath "OU=TermedUsers,OU=Actian Corporation,DC=Actian,DC=com"
#Get-ADUser -filter {EmailAddress -like $Email}
#$Username = Get-ADUser -filter {EmailAddress -like $Email} | Select-Object SAMAccountName
#Set-ADUser $Username.SAMAccountName -Clear Department,Title,telephoneNumber,Company,Manager,Mobile,HomePhone
#$TermDate = Get-date -UFormat "%Y-%m-%d"
#Set-ADUser $Username.SamAccountName -Replace @{Description="Termed "+ $TermDate}
#Set-ADUser $Username.SamAccountName -Enabled $False
#$GroupArray = Get-ADPrincipalGroupMembership $UserName.SAMAccountName | Select name | Where-Object {$_.name -ne "Domain Users"}
#ForEach ($Group in $GroupArray)
#	{
#	Remove-ADGroupMember -Identity $Group.Name -Member $Username.SAMAccountName -Confirm:$false
#	}