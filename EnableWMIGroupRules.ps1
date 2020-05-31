# ------------------------------------------------------------------------------------------------------------------------------------------------------------
#  Name        		:	EnableWMIGroupRules.ps1
#  Author      		:	Harsh Arora
#  Description     	:	This powershell script modifies existing Firewall Group Rules Profile type as well as Enabled state for a profile which will be passed  
#                       as one of the command-line argument and Add WMI related firewall rules for public profile using powershell cmdlets 
#  Arguments       	:	$GroupName, $Direction, $Profile
# ------------------------------------------------------------------------------------------------------------------------------------------------------------

$GroupName = $args[0]
$Direction = $args[1]
$Profile = $args[2]
$Logfile = "$env:SystemDrive\EnableWMIGroupRules.log"
$ListWMIGroupRules = "$env:SystemDrive\ListWMIGroupRules.log"
$FirewallRuleNames = @('Windows Management Instrumentation (ASync-In)', 'Windows Management Instrumentation (DCOM-In)', 'Windows Management Instrumentation (WMI-In)')
$FirewallRuleProfile = 'Public'

function LogWrite([string]$logstring)
{
	$DateTime = Get-Date -Format g
	Add-content $Logfile -value "$DateTime $logstring"
}

ForEach ($name in $GroupName)
{
   # Retrieves firewall rules based on Private, Public profile type and dumping into a file
   Get-NetFirewallRule -DisplayGroup $name -Direction In | Where-Object {$_.Profile -eq 'Private , Public'} | Out-File $ListWMIGroupRules -Append
   
   $Return = Get-NetFirewallRule -DisplayGroup $name -Direction In | Where-Object {$_.Profile -eq 'Private , Public'} |  
   
   # Modifying existing firewall group rule profile and enabled state
   Set-NetFirewallRule -Profile $Profile -Enabled True 
   
   If($? -eq $false)
    {
        LogWrite "Can't modify Firewall Group Rule Name: '$name' for Profile: '$Profile' in Direction: '$Direction' on Install System Type: $env:INSTALL_SYSTEM_TYPE returned: $?"     
    }
    Else
    {
        LogWrite "Modified existing Firewall Group Rule Name: '$name' for Profile: '$Profile' in Direction: '$Direction' on Install System Type: $env:INSTALL_SYSTEM_TYPE returned: $?"     
    }
}

# Add WMI related Firewall Rules For Public Profile.
ForEach ($name in $FirewallRuleNames)
{
   If($name -like '*ASync*')
   {
      $Return = New-NetFirewallRule -DisplayName $name -Group $GroupName -Profile $FirewallRuleProfile -Enabled False -Action Allow -Direction $Direction -Program '%systemroot%\system32\wbem\unsecapp.exe' -RemoteAddress 'LocalSubnet' -Protocol 'TCP' 
      
      If($? -eq $false)
      {
         LogWrite "Can't add Firewall Rule Name: '$name' with Group Name: '$GroupName' for Profile: '$FirewallRuleProfile' in Direction: '$Direction' on Install System Type: $env:INSTALL_SYSTEM_TYPE returned: $?"     
      }
      Else
      {
         LogWrite "Adding a Firewall Rule Name: '$name' with Group Name: '$GroupName' for Profile: '$FirewallRuleProfile' in Direction: '$Direction' on Install System Type: $env:INSTALL_SYSTEM_TYPE returned: $?"     
      }
   }
   
   ElseIf($name -like '*DCOM*')
   {
      $Return = New-NetFirewallRule -DisplayName $name -Group $GroupName -Profile $FirewallRuleProfile -Enabled False -Action Allow -Direction $Direction -Program '%SystemRoot%\system32\svchost.exe' -RemoteAddress 'LocalSubnet' -Protocol 'TCP' -LocalPort '135' 
      
      If($? -eq $false)
      {
         LogWrite "Can't add Firewall Rule Name: '$name' with Group Name: '$GroupName' for Profile: '$FirewallRuleProfile' in Direction: '$Direction' on Install System Type: $env:INSTALL_SYSTEM_TYPE returned: $?"     
      }
      Else
      {
         LogWrite "Adding a Firewall Rule Name: '$name' with Group Name: '$GroupName' for Profile: '$FirewallRuleProfile' in Direction: '$Direction' on Install System Type: $env:INSTALL_SYSTEM_TYPE returned: $?"     
     }
   }
   
   ElseIf($name -like '*WMI*')
   {
      $Return = New-NetFirewallRule -DisplayName $name -Group $GroupName -Profile $FirewallRuleProfile -Enabled False -Action Allow -Direction $Direction -Program '%SystemRoot%\system32\svchost.exe' -RemoteAddress 'LocalSubnet' -Protocol 'TCP'  
      
      If($? -eq $false)
      {
         LogWrite "Can't add Firewall Rule Name: '$name' with Group Name: '$GroupName' for Profile: '$FirewallRuleProfile' in Direction: '$Direction' on Install System Type: $env:INSTALL_SYSTEM_TYPE returned: $?"     
      }
      Else
      {
         LogWrite "Adding a Firewall Rule Name: '$name' with Group Name: '$GroupName' for Profile: '$FirewallRuleProfile' in Direction: '$Direction' on Install System Type: $env:INSTALL_SYSTEM_TYPE returned: $?"     
     }
   }
   
   Else
   {
      LogWrite " Not found, Firewall Rule Name: '$name' for Profile: '$FirewallRuleProfile' in Direction: '$Direction' on Install System Type: $env:INSTALL_SYSTEM_TYPE."     
   }
 }