#requires -Version 2

<#
        .SYNOPSIS
        Execute few queries from MSSQL DB and send the results as a pretty HTML 
#>
[CmdletBinding()]
param
(
    [Parameter(Mandatory = $true, Position = 0)]
    [System.String]
    $AddressList
)
function Get-QueryResultAsHTMLTable
{
    <#
            .SYNOPSIS
            Execute the query and parse the result as XML
            .DESCRIPTION
            Execute an MSSQL query and reteave the data, parsing the data into an html table with the headers provided
            .EXAMPLE
            Get-QueryResultAsHTMLTable
            -Data 'Column 1</td><td>Column 2</td><td>Column 3</td></tr>',
            -SQLQuery $('SELECT * FROM Table')
            -ConnectionString 'Server=Server;Database=DB;User Id=user;Password=pass'
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, Position = 0)]
        [System.String]
        $Data,
        
        [Parameter(Mandatory = $true, Position = 1)]
        [Object]
        $SQLQuery,
        
        [Parameter(Mandatory = $true, Position = 2)]
        [System.String]
        $ConnectionString
    )
    
    $Connection = New-Object -TypeName System.Data.SQLClient.SQLConnection
    $Connection.ConnectionString = $ConnectionString
    $Connection.Open()
    $Command = New-Object -TypeName System.Data.SQLClient.SQLCommand
    $Command.Connection = $Connection
    $Command.CommandText = $SQLQuery
    $Reader = $Command.ExecuteReader()
    
    while ($Reader.Read()) 
    {
        $Data = $Data + '<tr>'
        for ($i = 0; $i -lt $Reader.FieldCount; $i++)
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
    return $Data
}

function Send-HTMLmail
{
    <#
            .SYNOPSIS
            Send an email with HTML body
            .DESCRIPTION
            use this command to send email with html body to several users
            .EXAMPLE
            Send-HTMLmail
            -SendTo specify array of contacts
            -emailSmtpServer choose the smtp server
            -emailSmtpServerPort smtp server port
            -emailSmtpUser username for smtp server
            -emailSmtpPass password for smtp server
            -emailFrom user name to send from
            -Subject the subject of the email (auto add date to the title)
            - the html body
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, Position = 0)]
        [System.Object]
        $SendTo,
        
        [Parameter(Mandatory = $false, Position = 1)]
        [System.String]
        $emailSmtpServer = 'smtp.gmail.com',
        
        [Parameter(Mandatory = $false, Position = 2)]
        [System.String]
        $emailSmtpServerPort = '587',
        
        [Parameter(Mandatory = $false, Position = 3)]
        [System.String]
        $emailSmtpUser = 'user@gmail.com',
        
        [Parameter(Mandatory = $false, Position = 4)]
        [System.String]
        $emailSmtpPass = 'Password',

        [Parameter(Mandatory = $false, Position = 5)]
        [System.String]
        $emailFrom = 'user@gmail.com',

        [Parameter(Mandatory = $false, Position = 6)]
        [System.String]
        $Subject = 'Email title',
        
        [Parameter(Mandatory = $true, Position = 7)]
        [System.String]
        $html
    )
    $emailMessage = New-Object -TypeName System.Net.Mail.MailMessage

    $emailMessage.From = $emailFrom
    foreach ($adress in $SendTo)
    {
        $emailMessage.To.Add( $adress )
    }
    $date = Get-Date
    $emailMessage.Subject = "$Subject - $date"
    $emailMessage.IsBodyHtml = $true
    $emailMessage.Body = $html ##Get-Content $HtmlFilePath
    $SMTPClient = New-Object -TypeName System.Net.Mail.SmtpClient -ArgumentList ( $emailSmtpServer , $emailSmtpServerPort )
    $SMTPClient.Credentials = New-Object -TypeName System.Net.NetworkCredential -ArgumentList ( $emailSmtpUser , $emailSmtpPass )
    $SMTPClient.EnableSsl = $true
    $SMTPClient.Credentials = New-Object -TypeName System.Net.NetworkCredential -ArgumentList ( $emailSmtpUser , $emailSmtpPass )
    $SMTPClient.Send( $emailMessage )
}


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
    <title>This is the email title</title>
    </head>
    <body>

    <table width="583" border="0" cellspacing="0" cellpadding="0">
    <tr>
    </tr>
    <tr bgcolor="#3d90bd">
    <td align="left" valign="top" bgcolor="#3d90bd"><table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
    <td width="35" align="left" valign="top">&nbsp;</td>
    <td align="left" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
    <td align="center" valign="middle" bgcolor="#3d90bd">
    <div style="color:#FFFFFF; font-family:Roboto-Regular, Times, serif; font-size:34px;">Multi Platform Results from Last Night - 6.2.3</div>
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
    <td height="50" align="center" valign="middle" style="color:#FFFFFF; font-size:11px; font-family:Arial, Helvetica, sans-serif;"><b>Company NAME</b><br>For more info contact us:</br><a href="mailto:main@email.com?Subject=Errors%20In%20Test" target="_top">
    PV Infra team</a>
    <td width="35">&nbsp;</td>
    </tr>
    </table></td>
    </tr>
    </table>

    </body>
    </html>
'

##First Query

$Data = 'Operation System</td><td>Count</td><td>Results</td></tr>'

$SQLQuery = $('select [Os],count([TestName]),[Result]	
        FROM [TeamCity].[dbo].[MultiPlatformCommands]
        where CONVERT(DATE, [ResultTime]) = CONVERT(DATE, CURRENT_TIMESTAMP)
        group by [Result], [OS]
        order by [Result]
')

$Data = Get-QueryResultAsHTMLTable -Data $Data -SQLQuery $SQLQuery -ConnectionString 'Server=Server;Database=DB;User Id=user;Password=pass'

$Data = $Data + '</table>
    </td>
    </tr>
    <tr>
    <td align="center" valign="top" style="font-family:Arial, Helvetica, sans-serif; font-size:12px; color:#525252;">
    <table BORDERCOLOR=Black  border="1" cellspacing="3" cellpadding="5" width="500" align="center" style="background-color:#FFFFFF;text-align: center">
<tr><td>' 

##Secound Query

$Data = $Data +  'Test Name</td><td>Operation System</td><td>Results</td></tr>'

$SQLQuery = $("select [TestName] , [Os],[Result]
        FROM [TeamCity].[dbo].[MultiPlatformCommands]
        where CONVERT(DATE, [ResultTime]) = CONVERT(DATE, CURRENT_TIMESTAMP)
        and Result = 'Fail'
")

$Data = Get-QueryResultAsHTMLTable -Data $Data -SQLQuery $SQLQuery -ConnectionString 'Server=Server;Database=DB;User Id=user;Password=pass'

## build the complete HTML body
$html = $header + $Data + $footer

Send-HTMLmail -SendTo $AddressList -html $html
