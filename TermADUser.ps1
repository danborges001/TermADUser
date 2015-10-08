param(
[string]$Email
)
#Set variables, perform terminate procedures on user account
$OUPath = "Path to OU for Termed users"
$Username = Get-ADUser -filter {EmailAddress -like $Email}
$TermDate = Get-Date -UFormat "%Y-%m-%d"
Move-ADObject  $Username -TargetPath $OUPath
Set-ADUser $Username.SAMAccountName -Clear Department,Title,telephoneNumber,Company,Manager,Mobile,HomePhone
Set-ADUser $Username.SAMAccountName -Replace @{Description="Termed "+ $TermDate}
Set-ADUser $Username.SAMAccountName -Enabled $False
Set-ADObject $Username -Replace @{msExchHideFromAddressLists=$True}
$GroupArray = Get-ADGroup -property * -filter * | Where-Object {$_.member -like $Username}
ForEach ($Group in $GroupArray)
	{
	Remove-ADGroupMember -Identity $Group -Member $Username.SAMaccountName -Confirm:$false
	}
#Output for error possible debug message in perl script
Get-ADUser -filter {EmailAddress -like $Email}