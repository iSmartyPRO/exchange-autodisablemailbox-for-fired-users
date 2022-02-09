[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
$config = [Object](Get-Content '.\config.json' | Out-String | ConvertFrom-Json)
$encoding = [System.Text.Encoding]::UTF8

Function Get-FiredUsers($month = 9) {
  <# Получить список уволенных пользователей с имеющимися электронными почтами и более 6 месяцев (по умолчанию) #>
  return  Get-ADUser -SearchBase $config.SearchBase -Filter * -Properties mail, whenChanged | ? {$_.mail -ne $null -AND $_.whenChanged -lt $(Get-Date).AddMonths((-$month))}
}

Function ComposeHtmlMail($users) {
  <# Сформировать HTML шаблон для отправки письма #>
  $userCount = ($users | measure).Count
  $description = "Добрый день, <br><br> Данное письмо было автоматически сгенерировано PowerShell скриптом установленный на сервере $($config.serverName).<br><br>Список пользователей, которым были отключены корпоративные почтовые ящики"
  $tUsers = ""
  ForEach($user in $users) {
    $tUsers += "<tr><td style='border: 1px solid #1C6EA4; margin: 5px;'>$($user.name)</td><td style='border: 1px solid #1C6EA4; margin: 5px;'>$($user.whenChanged)</td></tr>"
  }
  $table = `
        "<table style='width: 600px; border: 1px solid #1C6EA4;' cellspacing='0' cellpadding='0'>
          <tr style='border: 1px solid #1C6EA4; background: #1C6EA4; color: #ffffff'>
            <td  style='border: 1px solid #1C6EA4; margin: 5px;text-align: center;font-weight: 700;'>Имя пользователя</td>
            <td style='border: 1px solid #1C6EA4; margin: 5px;text-align: center;font-weight: 700;'>Когда изменен</td>
          </tr>
          $($tUsers)
        </table>"
  $footer = "Всего: $($userCount) пользователей"
  return $description + $table + $footer
}

Function Send-Email($html) {
  Send-MailMessage -From $config.from -To $config.to -Subject $config.subject -SmtpServer $config.mailServer -BodyAsHtml $html -encoding $encoding
}

Function DisableMails($users) {
  $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $config.exchangePowerShell -Authentication Kerberos
  Import-PSSession $Session
  ForEach($user in $users) {
    Disable-Mailbox $user.SAMAccountName -Confirm:$false
    <# Write-Host $user.SAMAccountName #>
  }
  Remove-PSSession $Session
}

$users = Get-FiredUsers($config.filterMonth)
if($users) {
  $userCount = ($users | measure).Count
  $html = ComposeHtmlMail($users)

  DisableMails($users)

  Send-Email($html)
}
