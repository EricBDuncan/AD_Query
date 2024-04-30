<#PSScriptInfo
.VERSION 1
.GUID 468bfeb4-bb80-4c97-8ec6-0e66b0cbe7e0
.AUTHOR Eric Duncan
.COMPANYNAME kalyeri
.COPYRIGHT
MIT License

Copyright (c) 2024 Eric Duncan

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
.TAGS
.LICENSEURI https://mit-license.org/
.PROJECTURI
.ICONURI
.EXTERNALMODULEDEPENDENCIES
.REQUIREDSCRIPTS
.EXTERNALSCRIPTDEPENDENCIES
.RELEASENOTES
.TODO
#>
<#
.SYNOPSIS PowerShell-based GUI Active Directioy User Lookup Tool.
.DESCRIPTION
PowerShell-based GUI Active Directioy User Lookup Tool that does not use the AD PowerShell cmdlets, so it will work on any domain joined computer.
The code isn't the prettiest but this was initally a GUI POC.
I lost the link to the website that gave most of the GUI code examples; kudos to whomever this author is.
#>
param()

#Window CLI-Console show/hide
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
    # Hide = 0,
    # ShowNormal = 1,
    # ShowMinimized = 2,
    # ShowMaximized = 3,
    # Maximize = 3,
    # ShowNormalNoActivate = 4,
    # Show = 5,
    # Minimize = 6,
    # ShowMinNoActivate = 7,
    # ShowNoActivate = 8,
    # Restore = 9,
    # ShowDefault = 10,
    # ForceMinimized = 11

[Console.Window]::ShowWindow([Console.Window]::GetConsoleWindow(), 2) | Out-Null

function get-user ($user,$query) {
	begin {
	$Search = New-Object DirectoryServices.DirectorySearcher("(&(objectCategory=user)($query=*$user*))")
	$props="Name","Mail","UserPrincipalName","Samaccountname","Company","Department","Description","Manager","WhenCreated","msDS-UserPasswordExpiryTimeComputed"
	foreach ($property in $props) {$search.PropertiesToLoad.add($property) | Out-Null}
	$Results = $Search.FindAll()
	[int]$rcount=($Results).count
	
	
	}
	process {
	
	if ($rcount -le 1) {
		foreach ($property in $props) {"$property = $($Results.properties[$property])`n"}
		$expire=[datetime]::FromFileTime("$($Results.properties["msDS-UserPasswordExpiryTimeComputed"])")
		"Password expires on $expire"
	} ELSEIF ($rcount -gt 1) {
			[System.Windows.Forms.MessageBox]::Show("Found $rcount users...")	| out-null
			For ($i = 0; $i -lt $rcount) {
				$newResults=$Results[$i]
				foreach ($property in $props) {"$property = $($newResults.properties[$property])`n"}
				$expire=[datetime]::FromFileTime("$($newResults.properties["msDS-UserPasswordExpiryTimeComputed"])")
				"Password expires on $expire"
				$div
				$i++
			}
		} ELSE {"User not found."}
	}
}

function ShowResult () {
	$Label3.Text=get-user -user $TextBox.Text -query $query | out-string
}

Add-Type -assembly System.Windows.Forms
Add-Type -assembly System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
$main_form = New-Object System.Windows.Forms.Form
$main_form.Text ='Local Active Directory User Lookup'
$main_form.Width = 600
$main_form.Height = 500
$main_form.AutoSize = $true
$main_form.StartPosition = "CenterScreen"
$main_form.KeyPreview = $True
$main_form.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {ShowResult}
    }
)
$main_form.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
    {$main_form.Close()}})
[string]$div='-' * 50
$font = New-Object System.Drawing.Font("Arial", 14)

$Combo=New-Object System.Windows.Forms.ComboBox
#$Combo.DropDownStyle= "DropDownList"
$combo.Location = New-Object System.Drawing.Point(3, 20)
$Combo.AutoSize = $True
#Define a custom class for ComboBox items
class ComboBoxItem {
    [string]$DisplayName
    [string]$Value
	# Override ToString method to display DisplayName in ComboBox
    [string] ToString() {
        return $this.DisplayName
    }
}
# Add custom items to the ComboBox
$item1 = [ComboBoxItem]::new()
$item1.DisplayName = "AD Account"
$item1.Value = "samaccountname"

$item2 = [ComboBoxItem]::new()
$item2.DisplayName = "Email Address"
$item2.Value = "Mail"

$item3 = [ComboBoxItem]::new()
$item3.DisplayName = "Name"
$item3.Value = "Name"

$item4 = [ComboBoxItem]::new()
$item4.DisplayName = "Department"
$item4.Value = "Department"

$item5 = [ComboBoxItem]::new()
$item5.DisplayName = "Description"
$item5.Value = "Description"

$combo.Items.Add($item1) | out-null
$combo.Items.Add($item2) | out-null
$combo.Items.Add($item3) | out-null
$combo.Items.Add($item4) | out-null
$combo.Items.Add($item5) | out-null

# Define an event handler for SelectedIndexChanged event
$comboBox_SelectedIndexChanged = {
    $selectedItem = $combo.SelectedItem
    if ($selectedItem -ne $null) {
        $selectedValue = $selectedItem.Value
        #[System.Windows.Forms.MessageBox]::Show("Selected Value: $selectedValue")
		$global:query=$selectedValue
    }
}
 $selectedItem = $combo.SelectedItem
  $selectedValue = $selectedItem.Value
$selectedValue
# Attach event handler to SelectedIndexChanged event
$combo.add_SelectedIndexChanged($comboBox_SelectedIndexChanged)

# Add ComboBox to the form
$main_form.Controls.Add($combo)

$Label = New-Object System.Windows.Forms.Label
$Label.Text = "Select type of query to search by:"
$Label.Location = New-Object System.Drawing.Point(3,5)
$Label.AutoSize = $true
$main_form.Controls.Add($Label)

$textbox = New-Object System.Windows.Forms.textbox
$textbox.Width = 150
$textbox.Location = New-Object System.Drawing.Size(3,45)
$x=$TextBox.Text
$main_form.Controls.Add($textbox)
#global:x=$objTextBox.Text;

#$Label3 = New-Object System.Windows.Forms.Label
$Label3 = New-Object System.Windows.Forms.textbox
$label3.font = $font
$Label3.Multiline = $true
$Label3.ReadOnly = $true
$Label3.AcceptsTab = $true
$Label3.AcceptsReturn = $true
$Label3.Width = 550
$Label3.Height = 450
$Label3.ScrollBars = "both"
$Label3.WordWrap = $false
$Label3.Text = ""
$Label3.Location  = New-Object System.Drawing.Point(3,90)
$Label3.AutoSize = $true
$main_form.Controls.Add($Label3)

$Button = New-Object System.Windows.Forms.Button
$Button.Location = New-Object System.Drawing.Size(160,45)
$Button.Size = New-Object System.Drawing.Size(75,17)
$Button.Text = "Check"
$Button.Add_Click({
	ShowResult
	})
$main_form.Controls.Add($Button)



#$Label3.Text =  [datetime]::FromFileTime((Get-ADUser -identity $ComboBox.selectedItem -Properties pwdLastSet).pwdLastSet).ToString('MM dd yy : hh ss')

$main_form.Topmost = $True
$main_form.Add_Shown({$main_form.Activate()})
$main_form.ShowDialog()
