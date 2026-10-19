# --- CONFIG — edit these values -----------------------------------------------
$SqlServer     = "YOUR_SERVER\SQLEXPRESS"  # e.g. "localhost" or "MYPC\SQLEXPRESS"
$SqlDatabase   = "YOUR_DATABASE"           # e.g. "AdventureWorks"
$ThresholdMB   = 900                       # Alert when free space drops below this
$SmtpServer    = "smtp.office365.com"      # Your SMTP server
$SmtpPort      = 587
$SmtpUser      = "sender@yourdomain.com"   # Sender email (with SMTP auth)
$SmtpPassword  = "YOUR_APP_PASSWORD"       # App password (not your login password)
$ToEmail       = "recipient@yourdomain.com" # Who receives the alert
$FromEmail     = "sender@yourdomain.com"   # Must match $SmtpUser for most providers
# --------------------------------------------------------------------------------

$Query = @"
SELECT
    name AS FileName,
    size/128.0 AS TotalSizeMB,
    size/128.0 - CAST(FILEPROPERTY(name, 'SpaceUsed') AS INT)/128.0 AS FreeSpaceMB,
    physical_name AS PhysicalPath
FROM sys.database_files
WHERE type = 0
"@

try {
    $result = Invoke-Sqlcmd `
        -ServerInstance $SqlServer `
        -Database $SqlDatabase `
        -Query $Query `
        -ErrorAction Stop

    $free  = [math]::Round($result.FreeSpaceMB, 2)
    $total = [math]::Round($result.TotalSizeMB, 2)

    if ($free -lt $ThresholdMB) {

        $subject   = "SQL Server MDF Space Alert - Only $free MB free"
        $checkedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

        # ─── Inline-CSS HTML email body (Outlook-safe) ──────────────────────────
        $body = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>SQL Server MDF Space Alert</title>
</head>
<body style="margin:0; padding:0; background-color:#fdf0e3; font-family:Segoe UI, Arial, sans-serif;">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background-color:#fdf0e3; padding:24px 0;">
    <tr>
      <td align="center">
        <table role="presentation" width="600" cellpadding="0" cellspacing="0" style="background-color:#ffffff; border:2px solid #d2691e; border-radius:6px; overflow:hidden;">

          <tr>
            <td style="background-color:#d2691e; padding:18px 24px; border-bottom:3px solid #a8530f;">
              <table role="presentation" width="100%" cellpadding="0" cellspacing="0">
                <tr>
                  <td style="font-size:20px; font-weight:bold; color:#ffffff;">
                    &#9888; SQL Server MDF space alert
                  </td>
                </tr>
              </table>
            </td>
          </tr>

          <tr>
            <td style="padding:24px;">
              <p style="margin:0 0 16px; font-size:15px; color:#5a3a1e; line-height:1.6;">
                Free space on the database data file has dropped below the configured threshold. Review the details below and take action if needed.
              </p>

              <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background-color:#fdf0e3; border:1px solid #f0c39a; border-radius:4px; margin-bottom:20px;">
                <tr>
                  <td style="padding:14px 18px; border-bottom:1px solid #f0c39a; font-size:13px; color:#a8530f; width:140px;">Server</td>
                  <td style="padding:14px 18px; border-bottom:1px solid #f0c39a; font-size:14px; color:#5a3a1e; font-weight:bold;">$SqlServer</td>
                </tr>
                <tr>
                  <td style="padding:14px 18px; border-bottom:1px solid #f0c39a; font-size:13px; color:#a8530f;">Database</td>
                  <td style="padding:14px 18px; border-bottom:1px solid #f0c39a; font-size:14px; color:#5a3a1e; font-weight:bold;">$SqlDatabase</td>
                </tr>
                <tr>
                  <td style="padding:14px 18px; border-bottom:1px solid #f0c39a; font-size:13px; color:#a8530f;">Data file</td>
                  <td style="padding:14px 18px; border-bottom:1px solid #f0c39a; font-size:14px; color:#5a3a1e;">$($result.FileName)</td>
                </tr>
                <tr>
                  <td style="padding:14px 18px; border-bottom:1px solid #f0c39a; font-size:13px; color:#a8530f;">Path</td>
                  <td style="padding:14px 18px; border-bottom:1px solid #f0c39a; font-size:13px; color:#7a4a22; word-break:break-all;">$($result.PhysicalPath)</td>
                </tr>
                <tr>
                  <td style="padding:14px 18px; border-bottom:1px solid #f0c39a; font-size:13px; color:#a8530f;">Total size</td>
                  <td style="padding:14px 18px; border-bottom:1px solid #f0c39a; font-size:14px; color:#5a3a1e;">$total MB</td>
                </tr>
                <tr>
                  <td style="padding:14px 18px; font-size:13px; color:#a8530f;">Free space</td>
                  <td style="padding:14px 18px;">
                    <span style="display:inline-block; background-color:#d2691e; color:#ffffff; font-size:14px; font-weight:bold; padding:4px 12px; border-radius:4px;">
                      $free MB &mdash; below $ThresholdMB MB threshold
                    </span>
                  </td>
                </tr>
              </table>

              <p style="margin:0; font-size:12px; color:#a8753f;">
                Checked at $checkedAt
              </p>
            </td>
          </tr>

          <tr>
            <td style="background-color:#fdf0e3; padding:14px 24px; font-size:12px; color:#b58655; border-top:1px solid #f0c39a;">
              Automated alert from MDF space monitor &middot; do not reply to this message
            </td>
          </tr>

        </table>
      </td>
    </tr>
  </table>
</body>
</html>
"@
        # ─────────────────────────────────────────────────────────────────────────

        try {
            $smtpClient = New-Object System.Net.Mail.SmtpClient($SmtpServer, $SmtpPort)
            $smtpClient.EnableSsl = $true
            $smtpClient.Credentials = New-Object System.Net.NetworkCredential($SmtpUser, $SmtpPassword)

            $mailMessage = New-Object System.Net.Mail.MailMessage
            $mailMessage.From    = $FromEmail
            $mailMessage.To.Add($ToEmail)
            $mailMessage.Subject = $subject
            $mailMessage.Body    = $body
            $mailMessage.IsBodyHtml = $true

            $smtpClient.Send($mailMessage)
            Write-Output "ALERT sent — Free space: $free MB"
        }
        catch {
            Write-Error "Email send failed: $_"
        }
        finally {
            if ($mailMessage) { $mailMessage.Dispose() }
            if ($smtpClient)  { $smtpClient.Dispose() }
        }

    } else {
        Write-Output "OK — Free space: $free MB (threshold: $ThresholdMB MB)"
    }
}
catch {
    Write-Error "Script failed: $_"
}

Read-Host "Press Enter to exit"
