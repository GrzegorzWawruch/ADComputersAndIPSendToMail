   
 function Get-ADComputerIP {  
    param( 
        $Computers
    )
    if(-not $Computers) {
        $Computers = Get-ADComputer -Filter * -Properties IPv4Address | Select-Object Name,IPv4Address | Out-GridView -OutputMode Multiple -Title "Active Directory Machines" | Select-Object -ExpandProperty Name
    }

        $header = @"
<style>

    h1 {

        font-family: Arial, Helvetica, sans-serif;
        color: #e68a00;
        font-size: 28px;

    }
    
    h2 {

        font-family: Arial, Helvetica, sans-serif;
        color: #000099;
        font-size: 16px;

    }
  
   table {
		font-size: 12px;
		border: 0px; 
		font-family: Arial, Helvetica, sans-serif;
	} 
	
    td {
		padding: 4px;
		margin: 0px;
		border: 0;
	}
	
    th {
        background: #395870;
        background: linear-gradient(#49708f, #293f50);
        color: #fff;
        font-size: 11px;
        text-transform: uppercase;
        padding: 10px 15px;
        vertical-align: middle;
	}

    tbody tr:nth-child(even) {
        background: #f0f0f2;
    }

        #data {

        font-family: Arial, Helvetica, sans-serif;
        color: #ff3300;
        font-size: 12px;
    }
</style>
"@
    $ComputerInfo = foreach($computer in $Computers)
    {
        $item = Get-ADComputer -Identity $computer -Properties IPv4Address

        [PSCustomObject]@{
            AD_Computers_Names = $item.Name
            AD_Computers_IP = $item.IPv4Address
        }
    }

    #Prepare data for add to message

    $ToSend = $ComputerInfo | ConvertTo-Html -Title "Adresy IP Komputerów w domenie AD" -Head $header -PostContent "<p id='data'>Utworzono: $(Get-Date)</p>"
    #$ToSend | Out-File -FilePath "C:\Users\$env:USERNAME\Desktop\ADCompandIP.html"
    #$ToSend


    # Login to SMTP server
    $SmtpUsername = "###"
    $SmtpPassword = ConvertTo-SecureString -String "###" -AsPlainText -Force
    $SmtpCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $SmtpUsername, $SmtpPassword

    # Email Configuration
    $From = "###"
    $To = "###"
    $Subject = "Active Directory Computers with IP"
    $Body = $ToSend
    $SmtpServer = "ssl0.ovh.net"
    $SmtpPort = 587
    $UseSSL = $true

    # Create the SMTP client
    $Client = New-Object System.Net.Mail.SmtpClient($SmtpServer, $SmtpPort)
    $Client.EnableSsl = $UseSSL
    $Client.Credentials = $SmtpCredential

    # Message object
    $Message = New-Object System.Net.Mail.MailMessage($From, $To, $Subject, $Body)
    $Message.IsBodyHtml = $true

    # Send the email
    try {
        $Client.Send($Message)
        Write-Host "Email sent successfully."
    }
    catch {
        Write-Host "Error occurred during send message process: $_"
    }
}

# Run the function
Get-ADComputerIP