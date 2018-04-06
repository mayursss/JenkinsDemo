$Now = (Get-date)
$Computername = $env:COMPUTERNAME
$report ="$env:temp\System_Health_report_"+$Now.ToString('MM-dd-yyyy')+".html"
Set-Content $report ""
$fragments = @()
#$ImagePath = "C:\Users\ms458j\Documents\PS-Scripts\css\db.png"
#$ImageBits =  [Convert]::ToBase64String((Get-Content $ImagePath -Encoding Byte))
#$ImageFile = Get-Item $ImagePath
#$ImageType = $ImageFile.Extension.Substring(1) #strip off the leading .
#$ImageTag = "<Img src='data:image/$ImageType;base64,$($ImageBits)' Alt='$($ImageFile.Name)' style='float:left' width='120' height='120' hspace=10>"
#

$top = @"
<table>
<tr>
<td class='transparent'>$ImageTag</td><td class='transparent'><H1>System Report - $Computername</H1></td>
</tr>
</table>
"@
 
$fragments+=$top

$fragments+="<a href='javascript:toggleAll();' title='Click to toggle all sections'>+/-</a>"
########################
$Text = "Operating System"
$div = $Text.Replace(" ","_")
$fragments+= "<a href='javascript:toggleDiv(""$div"");' title='click to collapse or expand this section'><h2>$Text</h2></a><div id=""$div"">"

$fragments+= Get-Ciminstance -ClassName win32_operatingsystem |
Select @{Name="Operating System";Expression= {$_.Caption}},Version,InstallDate |
ConvertTo-Html -Fragment -As List
$fragments+="</div>"
########################
$Text = "System Information"
$div = $Text.Replace(" ","_")
$fragments+= "<a href='javascript:toggleDiv(""$div"");' title='click to collapse or expand this section'><h2>$Text</h2></a><div id=""$div"">"
 
$cs = Get-CimInstance -ClassName Win32_computersystem #-ComputerName $Computername  
$proc = Get-CimInstance -ClassName win32_processor #-ComputerName $Computername 

$data1 = [ordered]@{
TotalPhysicalMemGB = $cs.TotalPhysicalMemory/1GB -as [int]
NumProcessors = $cs.NumberOfProcessors
NumLogicalProcessors = $cs.NumberOfLogicalProcessors
HyperVisorPresent = $cs.HypervisorPresent
DeviceID = $proc.DeviceID
Name = $proc.Name
MaxClock = $proc.MaxClockSpeed
L2size = $proc.L2CacheSize
L3Size = $proc.L3CacheSize
 
}
$fragments+= New-Object -TypeName PSObject -Property $data1 | ConvertTo-Html -Fragment -As List
$fragments+="</div>"

########################
$Text = "Disk Information"
$div = $Text.Replace(" ","_")
$fragments+= "<a href='javascript:toggleDiv(""$div"");' title='click to collapse or expand this section'><h2>$Text</h2></a><div id=""$div"">"

$htmldata = @()
$diskdata = Get-CimInstance -ClassName win32_logicaldisk 
foreach ($disk in $diskdata){
    $htmlTab = New-Object -TypeName PSObject -Property @{
    Disk= $disk.DeviceID #$DiskTab.Disk
    SizeGB = "{0:N2}" -f $($disk.size/1GB) #$DiskTab.SizeGB
    FreeGB = [math]::round($disk.Freespace/1gb,2) -as [decimal]  #$DiskTab.FreeGB
    PctFree = [math]::round(($disk.freespace/$disk.size)*100,2) -as [decimal] #$DiskTab.PctFree
    }
    $htmldata += $htmlTab
}

[xml]$html = $htmldata | select disk,SizeGB,FreeGB,PctFree | ConvertTo-Html -Fragment

 for ($i=1;$i -le $html.table.tr.count-1;$i++)
    {
          if ($html.table.tr[$i].td[3].ToDecimal($_) -le 20)
          {
            $class = $html.CreateAttribute("class")
            $class.value = 'alert'
            $html.table.tr[$i].childnodes[3].attributes.append($class) | out-null
          }
}

$fragments+= $html.InnerXml
$fragments+="</div>"

########################

$Text = "EventLog Info"
$div = $Text.Replace(" ","_")
$fragments+= "<a href='javascript:toggleDiv(""$div"");' title='click to collapse or expand this section'><h2>$Text</h2></a><div id=""$div"">"


[xml]$html = Get-Eventlog -List | 
Select @{Name="Max(K)";Expression = {"{0:n0}" -f $_.MaximumKilobytes }},
@{Name="Retain";Expression = {$_.MinimumRetentionDays }},
OverFlowAction,
@{Name="Entries";Expression = {"{0:n0}" -f $_.entries.count}},
@{Name="Log";Expression = {$_.LogDisplayname}} | convertto-html -Fragment
 
for ($i=1;$i -le $html.table.tr.count-1;$i++) {
  if ($html.table.tr[$i].td[3] -eq 0) {
    $class = $html.CreateAttribute("class")
    $class.value = 'alert'
    $html.table.tr[$i].attributes.append($class) | out-null
  }
}

$fragments+= $html.InnerXml

$fragments+="</div>"


$fragments+= "<p class='footer'>$(get-date)</p>"
$head = @"
<Title>System Report - $($env:COMPUTERNAME)</Title>
<style>
body { background-color:#E5E4E2;
       font-family:Monospace;
       font-size:10pt; }
td, th { border:0px solid black; 
         border-collapse:collapse;
         white-space:pre; }
th { color:white;
     background-color:black; }
table, tr, td, th { padding: 2px; margin: 0px ;white-space:pre; }
tr:nth-child(odd) {background-color: lightgray}
table { width:95%;margin-left:5px; margin-bottom:20px;}
h2 {
 font-family:Tahoma;
 color:#6D7B8D;
}
.alert {
 color: red; 
 }
.footer 
{ color:green; 
  margin-left:10px; 
  font-family:Tahoma;
  font-size:8pt;
  font-style:italic;
}
.transparent {
background-color:#E5E4E2;
}
</style>
<script type='text/javascript' src='https://ajax.googleapis.com/ajax/libs/jquery/1.4.4/jquery.min.js'>
</script>
<script type='text/javascript'>
function toggleDiv(divId) {
   `$("#"+divId).toggle();
}
function toggleAll() {
    var divs = document.getElementsByTagName('div');
    for (var i = 0; i < divs.length; i++) {
        var div = divs[i];
        `$("#"+div.id).toggle();
    }
}
</script>
"@


$convertParams = @{ 
 head = $head 
 body = $fragments
}


convertto-html @convertParams | out-file $report


##############################################
################# Email ######################
##############################################
$smtpServer = "smtp.it.att.com"
$From = "ms458j@intl.att.com"
$To = "ms458j@intl.att.com"

write-output "          Creating email"
$messageSubject = 'System Health report for '+$now.ToString('dddd MMMM dd yyyy')
 
$Msg = New-Object Net.Mail.MailMessage # Create new e-mail message object 
$Smtp = New-Object Net.Mail.SmtpClient($SmtpServer) # Create new smtp client object 
$Msg.From = $From # Add from address to the message 
$Msg.To.Add($To) # Add to address to the message 
#$Msg.CC.Add($CcRecipients) # Add CC address(es) to the message 
$Msg.Subject = $messageSubject # Add the message subject to email
$Msg.IsBodyHTML = $true # set email format to html
[string]$body=get-content $report
$Msg.Body = ($body,'<H5>Thank you</H5>','<H3>ATT-EMAS Team</H3>') # Add the bodytext to the message 
$Msg.Body
write-output "          Sending email"
$Smtp.Send($Msg) # Send the message 
$Msg.Attachments.Dispose()
#Invoke-Item $report
