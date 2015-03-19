param(
[string]$Email
)
#Output for error possible debug message in perl script
Get-ADUser -filter {EmailAddress -like $Email}
#Set variables, perform terminate procedures on user account
$Username = Get-ADUser -filter {EmailAddress -like $Email}
$TermDate = Get-Date -UFormat "%Y-%m-%d"
Move-ADObject -TargetPath "OU=TermedUsers,OU=Actian Corporation,DC=Actian,DC=com"
Set-ADUser $Username.SAMAccountName -Clear Department,Title,telephoneNumber,Company,Manager,Mobile,HomePhone
Set-ADUser $Username.SamAccountName -Replace @{Description="Termed "+ $TermDate}
Set-ADUser $Username.SamAccountName -Enabled $False
$GroupArray = Get-ADGroup -property * -filter * | Where-Object {$_.member -like $Username}
ForEach ($Group in $GroupArray)
	{
	Remove-ADGroupMember -Identity $NewGroup -Member $Username.SAMaccountName -Confirm:$false
	}          