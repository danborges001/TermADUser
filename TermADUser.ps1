Import-Module ActiveDirectory
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
Add-Type -AssemblyName Microsoft.VisualBasic

#Get credentials from user

Function Get-UserCredentials(){
	$Script:UserCredential = Get-Credential
}

#This function gets the current date, formats the date,
#moves the user to the termed user OU,
#clears Department, telephoneNumber, Company, Manager, Mobile, HomePhone fields
#replaces Description with date of termination + formatted date
#Disables the account

Function Set-ADAccount(){
	Write-Host "Disabling AD account and removing group memberships..."
	
	$TermDate = Get-Date -UFormat "%Y-%m-%d"
	
	Set-ADUser $UserName.SAMAccountName `
	-Clear Department,Title,telephoneNumber,Company,Manager,Mobile,HomePhone `
	-Replace @{Description="Termed "+ $TermDate} `
	-Enabled $False
	
	Set-ADObject $UserName `
	-Replace @{msExchHideFromAddressLists=$True}
	
	Move-ADObject $UserName `
	-TargetPath ""
	
	Write-Host "Done!"
}

#Archives group membership and e-mails to a person if option is selected, then removes group memberships

Function Remove-GroupMemberships(){
	$GroupArray = Get-ADPrincipalGroupMembership $UserName.SAMAccountName `
	| Where-Object {$_.Name -ne "Domain Users"}	
	
	
	If ($GroupArray.count -gt 0){
		If ($RadioButton4.Checked -eq $True){
			ForEach ($Group in $GroupArray){
				Add-Content "$home\$Email.txt" $Group.name
			}
			
			Send-MailMessage `
			-To "$ArchiveRequester" `
			-From "Termination Script <noreply@actian.com>" `
			-Subject "User Group Membership for $Email" `
			-Body "File attached" `
			-Attachments "$HOME\$Email.txt" `
			-SMTPserver ""
			
			Remove-Item $HOME\$Email.txt
		}
		
		ForEach ($Group in $GroupArray){
			Remove-ADGroupMember `
			-Identity $Group.DistinguishedName `
			-Member $Username.SAMAccountName `
			-Confirm:$False
		}
	}
}

#This function connects to the Microsoft Online Service (MSOL) works with all functions

Function Connect-ExchangeOnline(){
	Connect-MSOLService -Credential $UserCredential
	
	$MSOLSession = New-PSSession `
	-ConfigurationName Microsoft.Exchange `
	-ConnectionUri https://outlook.office365.com/powershell-liveid/ `
	-Credential $UserCredential `
	-Authentication Basic `
	-AllowRedirection
	
	Import-PSSession $MSOLSession `
	-AllowClobber `
	-DisableNameChecking
}

#This function removes any mobile devices linked to the exchange mailbox

Function Remove-MobileDevices(){
	$SavedErrorAction=$Global:ErrorActionPreference
    $Global:ErrorActionPreference='stop'
	Try{
		Write-Host "Removing Mobile Devices"	
	
		Get-MobileDevice -MailBox $Email | Remove-MobileDevice -Confirm:$False
	
		Write-Host "Done!"
	}Catch [System.Management.Automation.RemoteException]{
		$ErrorCode = $Error[0].Exception
		Write-Host $ErrorCode
	}Finally{
        $Global:ErrorActionPreference=$SavedErrorAction
	}
}

#This function removes the Exchange Unified Messaging voice mail box

Function Disable-UnifiedMessaging(){
    $SavedErrorAction=$Global:ErrorActionPreference
    $Global:ErrorActionPreference='stop'    
	Try{
		Write-Host "Disabling Unified Messaging Mailbox"
		
		Disable-UMMailbox $Email -Confirm:$False
		
		Write-Host "Done!"
    }Catch [System.Management.Automation.RemoteException]{
		$ErrorCode = $Error[0].Exception
		Write-Host $ErrorCode
	}Finally{
        $Global:ErrorActionPreference=$SavedErrorAction
	}
}

#This function converts the user mailbox to shared and removes the E4 O365 license

Function Convert-MailBox(){
	$SavedErrorAction=$Global:ErrorActionPreference
    $Global:ErrorActionPreference='stop'    
	Try{
		Get-Mailbox -Identity $Email
		Write-Host "Converting user mailbox into a shared mailbox"
		
		$SavedErrorAction=$Global:ErrorActionPreference
		$Global:ErrorActionPreference='stop'
		
		Set-Mailbox $Email -Type Shared
		$IsShared=(Get-MailBox $Email).IsShared
		
		Write-Host "Done!"
		Write-Host "Waiting for O365 to report that the mailbox is shared before unlicensing..."
		
		Start-Sleep -s 10
		If ($IsShared -eq $False){
			While($IsShared -eq $False){
				$IsShared=(Get-MailBox $Email).IsShared
				Write-Host "."
				Start-Sleep -s 10
			}		
		}
		If ($IsShared -eq $True){
			Try{
				Write-Host "Removing MSOL license"
				
				Set-MSOLUserLicense `
				-UserPrincipalName $Email `
				-RemoveLicenses ""
				
				Write-Host "Done!"
			}Catch{
				$ErrorCode = $Error[0].Exception
				Write-Host $ErrorCode
			}Finally{
				$Global:ErrorActionPreference=$SavedErrorAction
			}
		}
	}Catch [System.Management.Automation.RemoteException]{
		$ErrorCode = $Error[0].Exception
		Write-Host $ErrorCode
	}Finally{
        $Global:ErrorActionPreference=$SavedErrorAction
	}
}


#This function removes the Lync account

Function Disable-Lync(){
	If (!((Get-PSSession).ComputerName -Like "")){
		$LyncSession = New-PSSession `
		-ConnectionURI "" `
		-Credential $UserCredential
		
		Import-PsSession $LyncSession
	}
	$SavedErrorAction = $Global:ErrorActionPreference
	$Global:ErrorActionPreference='stop'
	Try{
		Write-Host "Disabling Lync"
			
		Disable-CSUser $Email
			
		Write-Host "Done!"
	}Catch [System.Management.Automation.RemoteException]{
		$ErrorCode = $Error[0].Exception
		Write-Host $ErrorCode
	}Finally{
		$Global:ErrorActionPreference=$SavedErrorAction
	}
}

#This performs the open file action

Function Get-FileName(){    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
	$OpenFileDialog.InitialDirectory = $HOME
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.Filename
}

#Make sure you are connected to Microsoft Exchange Online

Function Check-PSSession(){
	If (!((Get-PSSession).ComputerName -Like "outlook.office365.com") -or !((Get-PSSession).State -Like "Opened")){
		Get-UserCredentials
		Connect-ExchangeOnline
	}
}

#Removes sessions

Function End-Sessions(){
	Get-PSSession | Remove-PSSession
}

#Validates that the e-mail address for the user to be terminated is valid

Function Validate-Input(){	
	$Script:UserName = (Get-ADuser -Filter {EmailAddress -Like $Email})
	$Script:SAMAccount = $UserName.SAMAccountName
		
	Try{
		Get-ADUser $SAMAccount
		Show-Announcement "User $SAMAccount validated, continuing..."
	}Catch{
		Show-Announcement "The e-mail address you entered for $Email did not work, please enter a new e-mail address:"
		Show-Form2
	}	
}

#Pop-up that alerts the user that the script is finished

Function show-announcement($Message){
	 [Microsoft.VisualBasic.Interaction]::MsgBox("$Message",'OkOnly,MsgBoxSetForeground,Information','Complete!')
	}

#Logic of the script runs the 2 guaranteed functions and then depending on the checkboxes checked, runs those triggers by OK button

Function Button-Click(){
	If ($RadioButton1.Checked -eq $True){
		$Script:InputFile = Get-FileName
		$Script:Emails = Get-Content $InputFile
	}Else{
		$Script:Emails = $Textbox1.Text
	}
	If ($RadioButton4.Checked -eq $True){
		$Script:ArchiveRequester = $Textbox2.Text
	}
	ForEach ($Email in $Emails){
		Validate-Input
		$Script:UserName = (Get-ADuser -Filter {EmailAddress -Like $Email})
		If ($Checkbox5.Checked -eq $True){
			Set-ADAccount
			Remove-GroupMemberships
		}
		If ($Checkbox4.Checked -eq $True){
			Check-PSSession
			Remove-MobileDevices
		}
		If ($Checkbox3.Checked -eq $True){
			Check-PSSession
			Disable-UnifiedMessaging
		}
		If ($Checkbox2.Checked -eq $True){
			Disable-Lync
		}
		If ($Checkbox1.Checked -eq $True){
			Check-PSSession
			Convert-MailBox
		}
	}
}

#Main windows form for the script

Function Show-Form(){

	$Form = New-Object System.Windows.Forms.Form
	$Form.Text = "User Termination"
	$Form.Size = New-Object System.Drawing.Size(500,300)
	$Form.StartPosition = "CenterScreen"

	$OKButton = New-Object System.Windows.Forms.Button
	$OKButton.Location = New-Object System.Drawing.Size(75,200)
	$OKButton.Size = New-Object System.Drawing.Size(75,23)
	$OKButton.Text = "OK"
	$OKButton.TabIndex = 6
	$OKButton.Add_Click({Button-Click;$Form.Close()})
	$Form.Controls.Add($OKButton)

	$CancelButton = New-Object System.Windows.Forms.Button
	$CancelButton.Location = New-Object System.Drawing.Size(150,200)
	$CancelButton.Size = New-Object System.Drawing.Size(75,23)
	$CancelButton.Text = "Cancel"
	$CancelButton.TabIndex = 7
	$CancelButton.Add_Click({End-Sessions;$Form.Close()})
	$Form.Controls.Add($CancelButton)

	$GroupBox1 = New-Object System.Windows.Forms.GroupBox
	$GroupBox1.Location = New-Object System.Drawing.Size(320,20)
	$GroupBox1.Size = New-Object System.Drawing.Size(140,60)
	$GroupBox1.Text = "Select Input Method"
	$Form.Controls.Add($GroupBox1)

	$GroupBox2 = New-Object System.Windows.Forms.GroupBox
	$GroupBox2.Location = New-Object System.Drawing.Size(280,100)
	$GroupBox2.Size = New-Object System.Drawing.Size(200,120)
	$GroupBox2.Text = "Group Memberships"
	$GroupBox2.Enabled = $False
	$Form.Controls.Add($GroupBox2)

	$Label1 = New-Object System.Windows.Forms.Label
	$Label1.Location = New-Object System.Drawing.Size(10,20)
	$Label1.Size = New-Object System.Drawing.Size(300,20)
	$Label1.Text = "Enter the e-mail addresses of the employees to terminate:"
	$Form.Controls.Add($Label1)
	
	$Label2 = New-Object System.Windows.Forms.Label
	$Label2.Location = New-Object System.Drawing.Size(15,55)
	$Label2.Size = New-Object System.Drawing.Size(160,30)
	$Label2.Text = "Email address of the user requesting archive:"
	$GroupBox2.Controls.Add($Label2)

	$TextBox1 = New-Object System.Windows.Forms.TextBox
	$TextBox1.Location = New-Object System.Drawing.Size(10,40)
	$TextBox1.Size = New-Object System.Drawing.Size(260,40)
	$TextBox1.TabIndex = 0
	$Form.Controls.Add($TextBox1)

	$TextBox2 = New-Object System.Windows.Forms.TextBox
	$TextBox2.Location = New-Object System.Drawing.Size(15,85)
	$TextBox2.Size = New-Object System.Drawing.Size(120,40)
	$TextBox2.Enabled = $False
	$GroupBox2.Controls.Add($TextBox2)

	$Checkbox1 = New-Object System.Windows.Forms.Checkbox
	$Checkbox1.Location = New-Object System.Drawing.Size(10,180)
	$Checkbox1.Size = New-Object System.Drawing.Size(180,20) 
	$Checkbox1.Text = "Convert to Shared Mailbox"
	$Checkbox1.TabIndex = 5
	$Form.Controls.Add($Checkbox1)

	$Checkbox2 = New-Object System.Windows.Forms.Checkbox
	$Checkbox2.Location = New-Object System.Drawing.Size(10,160)
	$Checkbox2.Size = New-Object System.Drawing.Size(180,20)
	$Checkbox2.Text = "Disable Lync"
	$Checkbox2.TabIndex = 4
	$Form.Controls.Add($Checkbox2)

	$Checkbox3 = New-Object System.Windows.Forms.Checkbox
	$Checkbox3.Location = New-Object System.Drawing.Size(10,140)
	$Checkbox3.Size = New-Object System.Drawing.Size(180,20)
	$Checkbox3.Text = "Disable Unified Messaging"
	$Checkbox3.TabIndex = 3
	$Form.Controls.Add($Checkbox3)

	$Checkbox4 = New-Object System.Windows.Forms.Checkbox
	$Checkbox4.Location = New-Object System.Drawing.Size(10,120)
	$Checkbox4.Size = New-Object System.Drawing.Size(180,20)
	$Checkbox4.Text = "Remove Mobile Devices"
	$Checkbox4.TabIndex = 2
	$Form.Controls.Add($Checkbox4)

	$Checkbox5 = New-Object System.Windows.Forms.Checkbox 
	$Checkbox5.Location = New-Object System.Drawing.Size(10,100)
	$Checkbox5.Size = New-Object System.Drawing.Size(280,20)
	$Checkbox5.Text = "Disable AD Account/Purge Group Memberships"
	$Checkbox5.Add_CheckStateChanged({ `
	$GroupBox2.Enabled = $CheckBox5.Checked})
	$Checkbox5.TabIndex = 1
	$Form.Controls.Add($Checkbox5)

	$Checkbox6 = New-Object System.Windows.Forms.Checkbox 
	$Checkbox6.Location = New-Object System.Drawing.Size(10,80)
	$Checkbox6.Size = New-Object System.Drawing.Size(180,20)
	$Checkbox6.Text = "Select/Deselect All"
	$Checkbox6.TabIndex = 9
	$Checkbox6.Add_CheckStateChanged({ `
	$Checkbox1.Checked = $CheckBox6.Checked; `
	$Checkbox2.Checked = $CheckBox6.Checked; `
	$Checkbox3.Checked = $CheckBox6.Checked; `
	$Checkbox4.Checked = $CheckBox6.Checked; `
	$Checkbox5.Checked = $CheckBox6.Checked})
	$Form.Controls.Add($Checkbox6)

	$RadioButton1 = New-Object System.Windows.Forms.RadioButton
	$RadioButton1.Location = new-object System.Drawing.Point(15,15)
	$RadioButton1.Size = New-Object System.Drawing.Size(80,20)
	$RadioButton1.Text = "CSV"
	$RadioButton1.Add_Click({ `
	$Textbox1.Enabled = $False; `
	$Textbox1.Text = ""})
	$GroupBox1.Controls.Add($RadioButton1)

	$RadioButton2 = New-Object System.Windows.Forms.RadioButton
	$RadioButton2.Location = new-object System.Drawing.Point(15,35)
	$RadioButton2.Size = New-Object System.Drawing.Size(80,20)
	$RadioButton2.Text = "Text Box"
	$RadioButton2.Checked = $true
	$RadioButton2.Add_Click({$Textbox1.Enabled = $True})
	$GroupBox1.Controls.Add($RadioButton2)

	$RadioButton3 = New-Object System.Windows.Forms.RadioButton
	$RadioButton3.Location = New-Object System.Drawing.Size(15,15)
	$RadioButton3.Size = New-Object System.Drawing.Size(100,20)
	$RadioButton3.Text = "Purge"
	$RadioButton3.Checked = $true
	$RadioButton3.Add_Click({ `
	$Textbox2.Enabled = $False; `
	$Textbox2.Text = ""})
	$GroupBox2.Controls.Add($RadioButton3)

	$RadioButton4 = New-Object System.Windows.Forms.RadioButton
	$RadioButton4.Location = New-Object System.Drawing.Size(15,35)
	$RadioButton4.Size = New-Object System.Drawing.Size(120,20)
	$RadioButton4.Text = "Purge and Archive"
	$RadioButton4.Add_Click({$Textbox2.Enabled = $True})
	$GroupBox2.Controls.Add($RadioButton4)

	$Form.Add_Shown({$Form.Activate()})
	[void] $Form.ShowDialog()

}

#Windows form for when the e-mail address for the user to be terminated is not valid

Function Show-Form2(){
	$Form2 = New-Object System.Windows.Forms.Form
	$Form2.Text = "Emergency Account Manipulation Tool"
	$Form2.Size = New-Object System.Drawing.Size(350,150)
	$Form2.StartPosition = "CenterScreen"

	$OKButton2 = New-Object System.Windows.Forms.Button
	$OKButton2.Location = New-Object System.Drawing.Size(10,80)
	$OKButton2.Size = New-Object System.Drawing.Size(75,23)
	$OKButton2.Text = "OK"
	$OKButton2.Add_Click({Get-Variables; $Form2.Close()})
	$Form2.Controls.Add($OKButton2)

	$CancelButton2 = New-Object System.Windows.Forms.Button
	$CancelButton2.Location = New-Object System.Drawing.Size(85,80)
	$CancelButton2.Size = New-Object System.Drawing.Size(75,23)
	$CancelButton2.Text = "Cancel"
	$CancelButton2.Add_Click({exit})
	$Form2.Controls.Add($CancelButton2)

	$Label2 = New-Object System.Windows.Forms.Label
	$Label2.Location = New-Object System.Drawing.Size(10,20)
	$Label2.Size = New-Object System.Drawing.Size(340,20)
	$Label2.Text = "Please enter a valid e-mail address"
	$Form2.Controls.Add($Label2)

	$TextBox1 = New-Object System.Windows.Forms.TextBox
	$TextBox1.Location = New-Object System.Drawing.Size(10,40)
	$TextBox1.Size = New-Object System.Drawing.Size(260,40)
	$Form2.Controls.Add($TextBox1)
	
	$Form2.Add_Shown({$Form2.Activate()})
	[void] $Form2.ShowDialog()

}

#Calls the form and kicks off the script

Show-Form