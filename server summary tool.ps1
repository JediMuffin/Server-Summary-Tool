

function test_decom{

$Header = @"
<style>
TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse; width: 85%;}
TD {border-width: 1px; padding: 0px; border-style: solid; border-color: black;}
TH {border: 1px solid black; background: #dddddd; padding: 5px;}
</style>
"@ 

#get hostname
$hostname = Invoke-Command $server {"$env:computername"}
$winver = Invoke-Command $server {(Get-WmiObject -Class Win32_OperatingSystem).Caption}

$report += "<center>";
$report += "<h1>Server Summary Report $hostname</h1><h3>$winver</h3><h4>Local Administrators</h4>";

#get local admins
$admins = Invoke-Command $server {net localgroup administrators}
$report += "<table>"
foreach($admin in $admins){
$report += "<tr><td>$admin</td></tr>"
}
$report += "</table>"


#get IP address
$report += "<h4>IP Addresses</h4>"
$report += Invoke-Command $server {
function get_ip{
        Get-WmiObject Win32_NetworkAdapterConfiguration | Select-Object Description, 
        @{Name='IpAddress';Expression={$_.IpAddress -join '; '}}, 
        @{Name='IpSubnet';Expression={$_.IpSubnet -join '; '}}, 
        @{Name='DefaultIPgateway';Expression={$_.DefaultIPgateway -join '; '}}
        }
 get_ip 
 } | Select Description, IpAddress | ConvertTo-Html -head $header

#get roles
$report += "<h4>Roles and Features</h4>";
$report += Invoke-Command $server {  Import-Module ServerManager; Get-WindowsFeature | Where-Object { $_.Installed -eq $true -and $_.SubFeatures.Count -eq 0}} | Select DisplayName, Name | ConvertTo-Html -head $header 


#get installed apps
$report += "<br><h4>Installed Applications</h4>"
$report += Invoke-Command $server { Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* }|  Select DisplayName  | ConvertTo-Html -head $header 

#get running services
$report += "<h4>Running Services</h4>"
$report += Invoke-Command $server { Get-Service | Where-Object {$_.Status -like '*Running*'}}  | Select-Object Name, DisplayName, Status | ConvertTo-Html -head $header 

#get listening ports
$report += "<br><h4>Listening Ports</h4>"
$report += Invoke-Command $server { 
function get-listeningtcpconnections {            
[cmdletbinding()]            
param(            
)            
            
try {            
    $TCPProperties = [System.Net.NetworkInformation.IPGlobalProperties]::GetIPGlobalProperties()            
    $Connections = $TCPProperties.GetActiveTcpListeners()            
    foreach($Connection in $Connections) {            
        if($Connection.address.AddressFamily -eq "InterNetwork" ) { $IPType = "IPv4" } else { $IPType = "IPv6" }            
                    
        $OutputObj = New-Object -TypeName PSobject            
        $OutputObj | Add-Member -MemberType NoteProperty -Name "LocalAddress" -Value $connection.Address            
        $OutputObj | Add-Member -MemberType NoteProperty -Name "ListeningPort" -Value $Connection.Port            
        $OutputObj | Add-Member -MemberType NoteProperty -Name "IPV4Or6" -Value $IPType            
        $OutputObj            
    }            
            
} catch {            
    Write-Error "Failed to get listening connections. $_"            
}           
} get-listeningtcpconnections } | Select LocalAddress, ListeningPort, IPV4Or6 | ConvertTo-Html -head $header
#

#get shared files
$report += "<br><h4>Shared Folders</h4>"
$report += Invoke-Command $server { get-wmiobject -class win32_share -property *} | select name, path, description | ConvertTo-Html -head $header

$report += "</center>";

#outfile
$report | Out-File -FilePath "\\corporate.ingrammicro.com\mts-global\MTS-Operations\Server-decom\validation_archive\$hostname.html"

}

function winform{
    #region_font

    #endregion

    #region_appearance
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $Form                            = New-Object system.Windows.Forms.Form
    $Form.ClientSize                 = '505,527'
    $Form.text                       = "Server Summary Validation v.1"
    $Form.TopMost                    = $false

#groupboxes
    $Groupbox1                       = New-Object system.Windows.Forms.Groupbox
    $Groupbox1.height                = 179
    $Groupbox1.width                 = 457
    $Groupbox1.text                  = "Output"
    $Groupbox1.location              = New-Object System.Drawing.Point(22,331)

    $Groupbox2                       = New-Object system.Windows.Forms.Groupbox
    $Groupbox2.height                = 219
    $Groupbox2.width                 = 458
    $Groupbox2.text                  = "Servers for validation:"
    $Groupbox2.location              = New-Object System.Drawing.Point(22,100)


    $Groupbox3                       = New-Object system.Windows.Forms.Groupbox
    $Groupbox3.height                = 69
    $Groupbox3.width                 = 458
    $Groupbox3.text                  = "Insert Server:"
    $Groupbox3.location              = New-Object System.Drawing.Point(22,15)

#textboxes

    #textbox for console
    $listbox_console                        = New-Object system.Windows.Forms.ListBox
    $listbox_console.width                  = 430
    $listbox_console.height                 = 136
    $listbox_console.location               = New-Object System.Drawing.Point(7,25)
    $listbox_console.Font                   = 'Microsoft Sans Serif,10'

    #textbox for server
    $txtbox_server                        = New-Object system.Windows.Forms.TextBox
    $txtbox_server.multiline              = $false
    $txtbox_server.width                  = 291
    $txtbox_server.height                 = 20
    $txtbox_server.location               = New-Object System.Drawing.Point(17,28)
    $txtbox_server.Font                   = 'Microsoft Sans Serif,10'  

    $listbox_server                        = New-Object system.Windows.Forms.ListBox
    $listbox_server.text                   = "listBox"
    $listbox_server.width                  = 290
    $listbox_server.height                 = 188
    $listbox_server.location               = New-Object System.Drawing.Point(13,21)

#buttons
    $btn_ping                        = New-Object system.Windows.Forms.Button
    $btn_ping.text                   = "Ping"
    $btn_ping.width                  = 128
    $btn_ping.height                 = 30
    $btn_ping.location               = New-Object System.Drawing.Point(320,21)
    $btn_ping.Font                   = 'Microsoft Sans Serif,10'

    $btn_validate                    = New-Object system.Windows.Forms.Button
    $btn_validate.text               = "Validate"
    $btn_validate.width              = 127
    $btn_validate.height             = 30
    $btn_validate.location           = New-Object System.Drawing.Point(321,130)
    $btn_validate.Font               = 'Microsoft Sans Serif,10'


    $btn_insert                      = New-Object system.Windows.Forms.Button
    $btn_insert.text                 = "Insert Server"
    $btn_insert.width                = 129
    $btn_insert.height               = 30
    $btn_insert.location             = New-Object System.Drawing.Point(320,28)
    $btn_insert.Font                 = 'Microsoft Sans Serif,10'

    $btn_help                        = New-Object system.Windows.Forms.Button
    $btn_help.text                   = "Help"
    $btn_help.width                  = 128
    $btn_help.height                 = 30
    $btn_help.location               = New-Object System.Drawing.Point(320,179)
    $btn_help.Font                   = 'Microsoft Sans Serif,10'

    $btn_remove                      = New-Object system.Windows.Forms.Button
    $btn_remove.text                 = "Clear All Items"
    $btn_remove.width                = 127
    $btn_remove.height               = 30
    $btn_remove.location             = New-Object System.Drawing.Point(320,62)
    $btn_remove.Font                 = 'Microsoft Sans Serif,10'

    $btn_del                      = New-Object system.Windows.Forms.Button
    $btn_del.text                 = "Remove Item"
    $btn_del.width                = 127
    $btn_del.height               = 30
    $btn_del.location             = New-Object System.Drawing.Point(320,96)
    $btn_del.Font                 = 'Microsoft Sans Serif,10'

    $Form.controls.AddRange(@($Groupbox1,$Groupbox2,$Groupbox3))
    $Groupbox1.controls.AddRange(@($listbox_console))
    $Groupbox2.controls.AddRange(@($listbox_server,$btn_ping,$btn_validate,$btn_help,$btn_remove,$btn_del))
    $Groupbox3.controls.AddRange(@($txtbox_server,$btn_insert))

#endregion

    #region_behavior
    #insert server
    $btn_insert.Add_Click({ 
        $servers += $txtbox_server.Text;
        $listbox_server.Items.Add($servers)
        $txtbox_server.Text = "";
     })

    #ping servers
    $btn_ping.Add_Click({ 
        foreach($item in $listbox_server.Items){
            Try{
            $ErrorActionPreference = 'Stop'
            $output = Test-Connection -Computer $item;
            $output = $item + " is reachable"; 
            }
            Catch{
            $output = $item + " is unreachable"
            }

            $listbox_console.Items.Add($output);
        }
     })



    #validate servers
    $btn_validate.Add_Click({
     #declare flag
     $value;
    
    foreach($server in $listbox_server.Items){

    #check if winrm is enabled
    try{ 
        $errorActionPreference = "Stop" 
        $result = Invoke-Command -ComputerName $server { 1 } 
        $value = 1;
    } 
    catch{ 
        Write-Verbose $_ 
        $value = 0; 
    }
   
    #winrm is enabled
    if($value -eq 1){

            try{
            $ErrorActionPreference = "SilentlyContinue"
            test_decom;
            $listbox_console.Items.Add("Report generated for node: " + $server)
            }
            catch{
            $listbox_console.Items.Add($server + " failed to connect. Please apply script locally.")
            }
    }

    #winrm is disabled
    else{
    try{
        $ErrorActionPreference = "SilentlyContinue"
        Write-Host($server)
        Write-Host($value) 
        Write-Host("reaches") 
        schtasks /create /S $server /tn "\ServerValidation" /xml "\\corporate.ingrammicro.com\mts-global\MTS-Operations\Server-decom\validation_script\Servervalidation.xml"
        start-sleep 2
        schtasks /RUN /S $server /tn "\ServerValidation"
        start-sleep 10
        schtasks /Delete /S $server /tn "\ServerValidation" /F
    }
    catch{
   
        $listbox_console.Items.Add($server + " failed to connect. please apply script locally.")
                }
            }
        }      
     })



    #help button
    $btn_help.Add_Click({ 
    [System.Windows.Forms.MessageBox]::Show("Please see KBA below" , "Help dialog")
     })

    #remove all from list
    $btn_remove.Add_Click({ 


    $listbox_server.Items.Clear()
         
     })

     #remove selected item from list
    $btn_del.Add_Click({ 

            while($listbox_server.SelectedItems) {
                $listbox_server.Items.Remove($listbox_server.SelectedItems[0])
            }
         
     })


#endregion

#region_show
$form.ShowDialog();
#endregion
}

#GUI
winform;
