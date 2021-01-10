Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#Your variables
$ExchangeUser = "Exchange user for create mailbox" 
$ExchangeServer = "IP or DNS name Exchange server"
$Domain = "@your_domain"
$DefaultPassword = "your_Password"
$ExchangeDatabase = "Database_Name"
$PathLog = "Path_for_whrite_logs"
$EmailArchive = "Path for unload mailbox"

#Start UI form
$Form1 = New-Object System.Windows.Forms.Form
$Form1.Text = "Create/ disable/ reset password Active Directory user"
$Form1.Size = New-Object System.Drawing.Size(830,320)
$Form1.StartPosition = "CenterScreen"
# Icon
$Form1.Icon = [Drawing.Icon]::ExtractAssociatedIcon((Get-Command powershell).Path)

 # Translite from RUS to EN
 function TranslitRU2LAT{
    param([string]$inString)
    $Translit = @{ 
    [char]'а' = "a";[char]'А' = "A";
    [char]'б' = "b";[char]'Б' = "B";
    [char]'в' = "v";[char]'В' = "V";
    [char]'г' = "g";[char]'Г' = "G";
    [char]'д' = "d";[char]'Д' = "D";
    [char]'е' = "e";[char]'Е' = "E";
    [char]'ё' = "yo";[char]'Ё' = "Yo";
    [char]'ж' = "zh";[char]'Ж' = "Zh";
    [char]'з' = "z";[char]'З' = "Z";
    [char]'и' = "i";[char]'И' = "I";
    [char]'й' = "y";[char]'Й' = "Y";
    [char]'к' = "k";[char]'К' = "K";
    [char]'л' = "l";[char]'Л' = "L";
    [char]'м' = "m";[char]'М' = "M";
    [char]'н' = "n";[char]'Н' = "N";
    [char]'о' = "o";[char]'О' = "O";
    [char]'п' = "p";[char]'П' = "P";
    [char]'р' = "r";[char]'Р' = "R";
    [char]'с' = "s";[char]'С' = "S";
    [char]'т' = "t";[char]'Т' = "T";
    [char]'у' = "u";[char]'У' = "U";
    [char]'ф' = "f";[char]'Ф' = "F";
    [char]'х' = "kh";[char]'Х' = "Kh";
    [char]'ц' = "ts";[char]'Ц' = "Ts";
    [char]'ч' = "ch";[char]'Ч' = "Ch";
    [char]'ш' = "sh";[char]'Ш' = "Sh";
    [char]'щ' = "sch";[char]'Щ' = "Sch";
    [char]'ъ' = "";[char]'Ъ' = "";
    [char]'ы' = "y";[char]'Ы' = "Y";
    [char]'ь' = "";[char]'Ь' = "";
    [char]'э' = "e";[char]'Э' = "E";
    [char]'ю' = "yu";[char]'Ю' = "Yu";
    [char]'я' = "ya";[char]'Я' = "Ya"
    }
    $outChars = ""
   # $translit.Text = ""
   # $inString = $FamTB.Text +"."+ $nameTB.Text[0] +"."+ $Init.Text[0]
    foreach ($c in $inChars = $inString.ToCharArray())
        {
        if ($Translit[$c] -cne $Null ) 
            {$outChars += $Translit[$c]}
        else
            {$outChars += $c}
        }
    $outChars
    }

#create new AD user
function NewUser { 

#Add variables for create user

    #in translite
$nameTB = TranslitRU2LAT $nameTBRU.Text
$FamTB = TranslitRU2LAT $FamTBRU.Text
$Init = TranslitRU2LAT $InitRU.Text


$Description = Get-ADUser -Identity $userDest.text -Properties Description #| Select-Object Description
$Title = Get-ADUser -Identity $userDest.text -Properties Title # | Select-Object Title
$Company = Get-ADUser -Identity $userDest.text -Properties Company
$Department = Get-ADUser -Identity $userDest.text -Properties Department
$userOU = Get-ADUser -Identity $userDest.text -Properties CanonicalName
$userOU1 = ($userOU.DistinguishedName -split ",",2)[1]
$ADname = $nameTBRU.Text + " " + $InitRU.Text[0] +". "+ $FamTBRU.Text
$ADdispname = $FamTBRU.Text + " " + $nameTBRU.Text 
$ADusPrNm = $FamTB +"."+ $nameTB[0] +"."+ $Init[0] + $Domain
$SamAccountNameLogin = $FamTB +"."+ $nameTB[0] +"."+ $Init[0]
$telephoneNumbers = $tel.Text
$mobile = $mob.Text
$middle = $Init.Text

#logs to window
$w1="Перевел в транслит нового пользователя = $($ADname) 
У пользователя $($userDest.text) следующие атрибуты:
Description = $($Description.Description)
Title = $($Title.Title)
OU = $($userOU1)

"

#Create user
New-ADUser -Name $ADname `
-DisplayName $ADdispname `
-GivenName $nameTBRU.Text `
-Surname $FamTBRU.Text `
-Initials $InitRU.Text[0] `
-Description $Description.Description `
-Title $Title.Title `
-Company $Company.Company `
-Department $Department.Department `
-UserPrincipalName $ADusPrNm `
-SamAccountName $SamAccountNameLogin `
-OfficePhone $telephoneNumbers `
-OtherName $middle `
-mobile $mobile `
-Path $userOU1 `
-ChangePasswordAtLogon $true `
-AccountPassword (ConvertTo-SecureString $DefaultPassword -AsPlainText -force) -Enabled $true

#Copy user group from destination user to new user
Get-ADUser -Identity $userDest.text -Properties memberof | Select-Object -ExpandProperty memberof | Add-ADGroupMember -Members $SamAccountNameLogin

#logs
$polz= Get-ADUser -Identity $SamAccountNameLogin -Properties * | Select-Object  DisplayName, GivenName, Surname, Initials, Title, Description, UserPrincipalName, SamAccountName, distinguishedName | fl | out-string
$w2="Пользователь создан
$($polz)
"

#create E-mail 
if ($checkbox.Checked -eq $true) {

$UserCredential = Get-Credential -Credential $ExchangeUser
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://+ $ExchangeServer +/PowerShell/  -Authentication Kerberos -Credential $UserCredential
Import-PSSession $Session -DisableNameChecking
Enable-Mailbox -Identity $SamAccountNameLogin -Database $ExchangeDatabase

while (([bool] (Get-ADUser -Identity $SamAccountNameLogin -Properties *).mail -eq $false)) {
    sleep 2
    $Result.Text = "Создание почтового ящика, ожидайте. $(Get-Date -Format T)"
} 
$Result.Text ="Почтовый ящик создан"

$post = Get-ADUser -Identity $SamAccountNameLogin -Properties * | Select-Object  mail | fl | out-string
$w3="Почта создана$($post)"

Remove-PSSession $Session
}

#log output
$Result.Text = $w1, $w2, $w3

#Whrite logs to file
$filelog = "
****************
Создание пользователя $(Get-Date -Format U)
    Скрипт запущен от пользователя $((Get-WMIObject -class Win32_ComputerSystem | select username).username)
права как у пользователя $($userDest.text)
    Список групп пользователя:
$((Get-ADPrincipalGroupMembership $userDest.text).name | fl | Out-String)
    Новый пользователь:
$($ADname)
Title = $($Title.Title)
OU = $($userOU1)
$($w3)
"
$filelog | out-file \\+ $PathLog +\Log.txt -append

#message for copy to service desk system
$o1="Логин $($SamAccountNameLogin)
Пароль стандартный " 
$out.Text = $o1
}


# Disable AD user
function DisableAccount {
$Login = $Loginform.text 

Disable-ADAccount $Login

     #check availability E-mail
if ([bool] (Get-ADUser -Identity $Login -Properties *).mail -eq $true) {
$UserCredential = Get-Credential -Credential $ExchangeUser
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://+ $ExchangeServer +/PowerShell/  -Authentication Kerberos -Credential $UserCredential
Import-PSSession $Session -DisableNameChecking
sleep 5
$BatchName = "DisableAccount"
$PrimaryPath = $EmailArchive + $Login + ".pst"
#Unload primary mailbox
New-MailboxExportRequest -Mailbox $Login -BatchName $BatchName -FilePath $PrimaryPath

#Unload archive mailbox
if ([bool] (Get-Mailbox -Identity $Login).ArchiveDatabase -eq $true) {
$ArhivePath = $EmailArchive + $Login + "_Archive.pst"
New-MailboxExportRequest -Mailbox $Login -BatchName $BatchName -FilePath $ArhivePath -IsArchive


}
#waiting for the completion of unloading mailboxes
while ((Get-MailboxExportRequest -BatchName $BatchName | Where {($_.Status -eq “Queued”) -or ($_.Status -eq “InProgress”)})) {
    sleep 15
    $Result.Text = "Скрипт работает. Ожидаем завершения... $(Get-Date -Format U)"
} 
$Result.Text ="Выгрузка законченна"

#Deleting primary mailbox after upload to archive folder
Get-MailboxExportRequest -BatchName $BatchName -Status Completed | Remove-MailboxExportRequest -Confirm:$false 
$DistribList = Get-DistributionGroup
foreach($List in $DistribList.name){     
        Remove-DistributionGroupMember -Identity $List -Member $Login -Confirm:$false -ErrorAction Ignore  
} 

#Deleting archive mailbox after upload to archive folder
if ([bool] (Get-Mailbox -Identity $Login).ArchiveDatabase -eq $true) {

if ([bool] (Get-Item $ArhivePath -ErrorAction SilentlyContinue) -eq $true){
Disable-Mailbox -Identity $Login -Archive -Confirm:$false
$OutArchive =$Login + "_Archive.pst = $([math]::round(((Get-Item $ArhivePath).length/1MB),0)) MB "   
    }
        else {
        $erroutarch = "Ошибка выгрузки архива. Файл не создан"
    }

}
#Deleting primary mailbox after upload to archive folder
if ([bool] (Get-Item $PrimaryPath -ErrorAction SilentlyContinue) -eq $true){
Disable-Mailbox -Identity $Login -Confirm:$false
$outpst = $Login + ".pst = $([math]::round(((Get-Item $PrimaryPath).length/1MB),0)) MB "
    }
        else{
        $erroutpst = "Ошибка выгрузки основного ящика. Файл не создан"
}
Remove-PSSession $Session

    }

$filelog = "
****************
Отключение пользователя $(Get-Date -Format U)
    Скрипт запущен от пользователя $((Get-WMIObject -class Win32_ComputerSystem | select username).username)
Отключен пользователь $((Get-ADUser -Identity $Login -Properties *).DisplayName)
Файл выгрузки:
$OutArchive
$outpst
$erroutarch
$erroutpst
"
$filelog | out-file \\+ $PathLog +\Log.txt -append

 $Result.Text ="Скрипт выполнен
Пользователь $((Get-ADUser -Identity $Login -Properties *).DisplayName) отключен
Файл выгрузки:
$OutArchive
$outpst
$erroutarch
$erroutpst
" 

$out.Text = "Учетная запись пользователя отключенна"
}

#confirmation of action "reset password"
function mesbox2 {
$msbox = [System.Windows.Forms.MessageBox]::Show('Вы точно хотите сбросить пароль пользователю?', 'Сброс пароля', 'YesNo', 'Warning')

# check the result:
if ($msbox -eq 'Yes')
{
  (Resetpassw)
}
else
{
  $Result.Text = "Действие отменено"
}
}

#confirmation of action "disable user"
function mesbox {
$msbox = [System.Windows.Forms.MessageBox]::Show('Вы точно хотите отключить пользователя?', 'Отключение пользователя', 'YesNo', 'Warning')

# check the result:
if ($msbox -eq 'Yes')
{
  (DisableAccount)
}
else
{
  $Result.Text = "Действие отменено"
}
}

#Find AD user
Function FindUser{
$find = '*' + $FamTB1.Text + '*'

$userline = Get-ADUser -f {name -like $find } -Properties *

foreach($Listusr in $userline){     
  $result.AppendText("
  Пользователь- $($Listusr.name)
  Логин- $($Listusr.samaccountname)
  Учетная запись активна- $($Listusr.enabled)
  Должность- $($Listusr.title)
  Группы пользователя - 
$((Get-ADPrincipalGroupMembership $Listusr.samaccountname).name | Out-String)
  *****************************
  ")
} 

}
#Reset password for AD user
Function Resetpassw {
$Login = $Loginform.text

Set-ADAccountPassword $Login -Reset -NewPassword (ConvertTo-SecureString -AsPlainText -String $DefaultPassword -force)
Set-ADUser -Identity $Login -ChangePasswordAtLogon $true
$Result.Text = "Пароль пользователя $($Login) сброшен.
Лог:
$(Get-ADUser $Login -Properties * |select name, passworde* | fl | Out-String)"

$filelog = "
****************
Сброс пароля пользователя $(Get-Date -Format U)
    Скрипт запущен от пользователя $((Get-WMIObject -class Win32_ComputerSystem | select username).username)
Сброшен пароль пользователю $(Get-ADUser $Login -Properties * | select name | fl | Out-String)
"
$filelog | out-file \\+ $PathLog +\Log.txt -append
}

#Create UI form
$GroupBox1 = New-Object System.Windows.Forms.GroupBox
$GroupBox1.Text = "Создание пользователя"
$GroupBox1.AutoSize = $true
$GroupBox1.Location  = New-Object System.Drawing.Point(10,10)
$Form1.Controls.Add($GroupBox1)


$Label2 = New-Object System.Windows.Forms.Label
$Label2.Location = New-Object System.Drawing.Point(10,20)
$Label2.Size = New-Object System.Drawing.Size(60,20)
$Label2.Text = "Фамилия:"
$GroupBox1.Controls.Add($Label2)

$FamTBRU = New-Object System.Windows.Forms.TextBox
$FamTBRU.Location = New-Object System.Drawing.Point(80,17)
$FamTBRU.Size = New-Object System.Drawing.Size(140,20)
$GroupBox1.Controls.Add($FamTBRU)

$Label1 = New-Object System.Windows.Forms.Label
$Label1.Location = New-Object System.Drawing.Point(10,50) 
$Label1.Size = New-Object System.Drawing.Size(30,20)
$Label1.Text = "Имя:"
$GroupBox1.Controls.Add($Label1)

$nameTBRU = New-Object System.Windows.Forms.TextBox
$nameTBRU.Location = New-Object System.Drawing.Point(80,47) 
$nameTBRU.Size = New-Object System.Drawing.Size(140,20)
$GroupBox1.Controls.Add($nameTBRU)

$Label3 = New-Object System.Windows.Forms.Label
$Label3.Location = New-Object System.Drawing.Point(10,80)
$Label3.Size = New-Object System.Drawing.Size(60,20)
$Label3.Text = "Отчество:"
$GroupBox1.Controls.Add($Label3)

$InitRU = New-Object System.Windows.Forms.TextBox
$InitRU.Location = New-Object System.Drawing.Point(80,77)
$InitRU.Size = New-Object System.Drawing.Size(140,20)
$GroupBox1.Controls.Add($InitRU)

$Label4 = New-Object System.Windows.Forms.Label
$Label4.Location = New-Object System.Drawing.Point(10,110)
$Label4.Size = New-Object System.Drawing.Size(60,20)
$Label4.Text = "Телефон:"
$GroupBox1.Controls.Add($Label4)

$tel = New-Object System.Windows.Forms.TextBox
$tel.Location = New-Object System.Drawing.Point(80,107)
$tel.Size = New-Object System.Drawing.Size(140,20)
$GroupBox1.Controls.Add($tel)

$Label5 = New-Object System.Windows.Forms.Label
$Label5.Location = New-Object System.Drawing.Point(10,140)
$Label5.Size = New-Object System.Drawing.Size(70,20)
$Label5.Text = "Мобильный:"
$GroupBox1.Controls.Add($Label5)

$mob = New-Object System.Windows.Forms.TextBox
$mob.Location = New-Object System.Drawing.Point(80,137)
$mob.Size = New-Object System.Drawing.Size(140,20)
$GroupBox1.Controls.Add($mob)


$Label4 = New-Object System.Windows.Forms.Label
$Label4.Location = New-Object System.Drawing.Point(10,170)
$Label4.Size = New-Object System.Drawing.Size(70,20)
$Label4.Text = "Права как у:"
$GroupBox1.Controls.Add($Label4)

$userDest = New-Object System.Windows.Forms.TextBox
$userDest.Location = New-Object System.Drawing.Point(80,167)
$userDest.Size = New-Object System.Drawing.Size(140,40)
$GroupBox1.Controls.Add($userDest)
 
$button1 = New-Object System.Windows.Forms.Button
$button1.Text="Создать"
$button1.Location = New-Object System.Drawing.Point(10,195)
$button1.Size = New-Object System.Drawing.Size(60,20)
$button1.add_click($Function:NewUser) 
$Form1.AcceptButton = $button1
$GroupBox1.Controls.Add($button1)
 
$Result = New-Object System.Windows.Forms.TextBox 
$Result.Location = New-Object System.Drawing.Point(240,10)
$Result.Size = New-Object System.Drawing.Size(330,258)
$Result.ReadOnly = $true
$Result.MultiLine = $True 
$Result.ScrollBars = "Vertical" 
$Result.Text = ""
$Form1.Controls.Add($Result)

$out = New-Object System.Windows.Forms.TextBox 
$out.Location = New-Object System.Drawing.Point(575,196)
$out.Size = New-Object System.Drawing.Size(210,70)
$out.MultiLine = $True 
$out.Text = ""
$Form1.Controls.Add($out)

$checkbox = New-Object System.Windows.Forms.checkbox 
$checkbox.Location = new-object System.Drawing.Point(80,195) 
$checkbox.size = New-Object System.Drawing.Size(60,20) 
$checkbox.Text = "Почта" 
$GroupBox1.Controls.Add($checkbox) 

$GroupBox2 = New-Object System.Windows.Forms.GroupBox
$GroupBox2.Text = "Поиск пользователя"
$GroupBox2.Size = New-Object System.Drawing.Size(227,80)
$GroupBox2.Location  = New-Object System.Drawing.Point(575,10)
$Form1.Controls.Add($GroupBox2)

$Label21 = New-Object System.Windows.Forms.Label
$Label21.Location = New-Object System.Drawing.Point(10,20)
$Label21.Size = New-Object System.Drawing.Size(70,20)
$Label21.Text = "Кого ищем:"
$GroupBox2.Controls.Add($Label21)

$FamTB1 = New-Object System.Windows.Forms.TextBox
$FamTB1.Location = New-Object System.Drawing.Point(80,17)
$FamTB1.Size = New-Object System.Drawing.Size(140,20)
$GroupBox2.Controls.Add($FamTB1)

$button2 = New-Object System.Windows.Forms.Button
$button2.Text="Найти"
$button2.Location = New-Object System.Drawing.Point(10,47)
$button2.Size = New-Object System.Drawing.Size(60,20)
$button2.add_click({FindUser})
$Form1.AcceptButton = $button2
$GroupBox2.Controls.Add($button2)

$GroupBox3 = New-Object System.Windows.Forms.GroupBox
$GroupBox3.Text = "Отключить пользователя\ Сбросить пароль"
$GroupBox3.AutoSize = $true
$GroupBox3.Location  = New-Object System.Drawing.Point(575,90)
$Form1.Controls.Add($GroupBox3)

$button3 = New-Object System.Windows.Forms.Button
$button3.Text="Отключить"
$button3.Location = New-Object System.Drawing.Point(10,57)
$button3.Size = New-Object System.Drawing.Size(70,20)
$button3.add_click({mesbox})
$Form1.AcceptButton = $button3
$GroupBox3.Controls.Add($button3)

$button4 = New-Object System.Windows.Forms.Button
$button4.Text="Сбросить пароль"
$button4.Location = New-Object System.Drawing.Point(110,57)
$button4.Size = New-Object System.Drawing.Size(110,20)
$button4.add_click({mesbox2})
$Form1.AcceptButton = $button4
$GroupBox3.Controls.Add($button4)

$LabelLogin = New-Object System.Windows.Forms.Label
$LabelLogin.Location = New-Object System.Drawing.Point(10,30)
$LabelLogin.Size = New-Object System.Drawing.Size(60,20)
$LabelLogin.Text = "Логин:"
$GroupBox3.Controls.Add($LabelLogin)

$Loginform = New-Object System.Windows.Forms.TextBox
$Loginform.Location = New-Object System.Drawing.Point(80,27)
$Loginform.Size = New-Object System.Drawing.Size(140,20)
$GroupBox3.Controls.Add($Loginform)



$Form1.Topmost = $True
$Form1.ShowDialog()
