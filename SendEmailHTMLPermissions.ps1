Param (
    [Parameter(Position=0, Mandatory=$true)][string] $AdressList
)

$header = '<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<style> 
        .rotate90
        {
            transform:rotate(180deg);
            -ms-transform:rotate(180deg); /* IE 9 */
            -webkit-transform:rotate(180deg); /* Opera, Chrome, and Safari */
        }
    </style>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Last Day Installations Results</title>
</head>
<body>

    <table width="1200" border="0" cellspacing="0" cellpadding="0">
      <tr>
      </tr>
      <tr bgcolor="#3d90bd">
        <td align="left" valign="top" bgcolor="#3d90bd"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="35" align="left" valign="top">&nbsp;</td>
            <td align="left" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td align="center" valign="middle" bgcolor="#3d90bd">
                  <div style="font-family: Verdana, Geneva, sans-serif; color:#898989; font-size:12px;"></div></td>
              </tr>
              <tr>
                <td align="left" valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:12px; color:#525252;">
				<table BORDERCOLOR=Black  border="1" cellspacing="3" cellpadding="5" width="500" align="center" style="background-color:#FFFFFF;text-align: center;margin-bottom: 5px">
				<tr><td>'
$footer = '</table>
				</td>
              </tr>
            </table></td>
            <td width="35" align="left" valign="top">&nbsp;</td>
          </tr>
        </table><p></p></td>
      </tr>
      <tr>
        <td align="left" valign="top" bgcolor="#3d90bd" style="background-color:#3d90bd;"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="35">&nbsp;</td>
PV Infra team</a>
            <td width="35">&nbsp;</td>
          </tr>
        </table></td>
      </tr>
  </table>

</body>
</html>
'



")

$Connection = New-Object System.Data.SQLClient.SQLConnection
$Connection.Open()
$Command = New-Object System.Data.SQLClient.SQLCommand
$Command.Connection = $Connection
$Command.CommandText = $SQLQuery
$Reader = $Command.ExecuteReader()

while ($Reader.Read()) {
        $Data = $Data + '<tr>'
        for ($i=0; $i -lt $reader.FieldCount; $i++)
        {
            if ($Reader.GetValue($i) -match 'fail')
            {
                $Data = $Data +  '<td bgcolor="#ff4d4d">' +$Reader.GetValue($i) +'</td>'
            }
            elseif ($Reader.GetValue($i) -match 'SUCCESS' -Or $Reader.GetValue($i) -match 'pass')
            {
                $Data = $Data +  '<td bgcolor="#00cc00">' +$Reader.GetValue($i) +'</td>'
            }
            else
            {
                $Data = $Data +  '<td>' +$Reader.GetValue($i) +'</td>'
            }
        }
        $Data = $Data + '</tr>'
}
$Connection.Close()

$Data = $Data + '</table>
				</td>
              </tr>
              <tr>
                <td align="center" valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:12px; color:#525252;">
				<table BORDERCOLOR=Black  border="1" cellspacing="3" cellpadding="5" width="1000" align="center" style="background-color:#FFFFFF;text-align: center">
				<tr><td>' 

####Secound Table

$SQLQuery= $("SELECT * 
")

$Connection = New-Object System.Data.SQLClient.SQLConnection
$Connection.Open()
$Command = New-Object System.Data.SQLClient.SQLCommand
$Command.Connection = $Connection
$Command.CommandText = $SQLQuery
$Reader = $Command.ExecuteReader()

while ($Reader.Read()) {
        $Data = $Data + '<tr>'
        for ($i=0; $i -lt $reader.FieldCount; $i++)
        {
            if ($Reader.GetValue($i) -match 'False')
            {
                $Data = $Data +  '<td bgcolor="#ff4d4d">' + 'Failed' +'</td>'
            }
            elseif ($Reader.GetValue($i) -match 'SUCCESS' -Or $Reader.GetValue($i) -match 'True')
            {
                $Data = $Data +  '<td bgcolor="#00cc00">' + 'Passed' +'</td>'
            }
            else
            {
                $Data = $Data +  '<td>' +$Reader.GetValue($i) +'</td>'
            }
        }
        $Data = $Data + '</tr>'
}
$Connection.Close()


$html = $header + $Data + $footer
Write-Host $html

$SendTo = $AdressList
$emailSmtpServer = "smtp.gmail.com"
$emailSmtpServerPort = "587"
$emailMessage = New-Object System.Net.Mail.MailMessage
foreach ($adress in $SendTo){$emailMessage.To.Add( $adress )}
$date = Get-Date
$emailMessage.IsBodyHtml = $true
$emailMessage.Body = $html ##Get-Content $HtmlFilePath
$SMTPClient = New-Object System.Net.Mail.SmtpClient( $emailSmtpServer , $emailSmtpServerPort )
$SMTPClient.Credentials = New-Object System.Net.NetworkCredential( $emailSmtpUser , $emailSmtpPass );
$SMTPClient.EnableSsl = $true
$SMTPClient.Credentials = New-Object System.Net.NetworkCredential( $emailSmtpUser , $emailSmtpPass );
$SMTPClient.Send( $emailMessage )
