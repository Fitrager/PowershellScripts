#Так как скрипт в открытом доступе, то все чувствительные данные удалены

Remove-Variable -Name * -Force -ErrorAction SilentlyContinue #Очистка все сохраненных переменных

$Admin="";
$DomainController="";

Write-Host "Тип заявки 1: [Для выбора создание пользователя в Active Directory напишите 1]"
Write-Host "Тип заявки 2: [Для создания пользователя в ServiceDesk напишите 2]"
Write-Host "Тип заявки 3: [Для создания пользователя в Exchange напишите 3]"
Write-Host "Тип заявки 4: [Для создания пользователя в Lync напишите 4]"
Write-Host "Тип заявки 5: [Для смены UPN и создания Google Workspace напишите 5]"
Write-Host "Тип заявки 6: [Для выхода из скрипта нажмите 6]"

$UserCredential = Get-Credential -credential $Admin #Получения данных для подключения
$Login = Read-Host "Введите логин пользователя для проверки существования в AD"

#Проверка существования пользователя
$PShellsession = New-PSSession -computername $DomainController -Credential $UserCredential
Invoke-Command -Session $PShellsession -ScriptBlock {
$CreateUser = $(try {Get-ADUser $Using:Login} catch {$Null})
If ($CreateUser -ne $Null) 
    {
     Write-Host "$Using:Login Уже существует" -foregroundcolor "green"
     Get-ADUser $Using:Login -property createTimeStamp |select name,createTimeStamp
     break
    }
    else 
    {
     Write-Host "$Using:Login НЕ существует" -foregroundcolor "green"
     Read-Host "Для продолжения нажмите Enter"
    }
}
Disconnect-PSSession $PShellsession

#Получение данных с ServiceDesk
$RequestIDServiceDesk = Read-Host "Введите номер заявки"
$SQLServer = "SERVICEDESK\SQLSERVICE";
$SQLCatalog = "ServiceDeskBase";
$SQLLogin = "";
$SQLPassword = "";
$SQLConnection = New-Object System.Data.SqlClient.SqlConnection
$SQLConnection.ConnectionString = "Server=$SQLServer; Database=$SQLCatalog; User ID=$SQLLogin; Password=$SQLPassword;"
$SQLConnection.Open()
$SQLCmdLogin = $SQLConnection.CreateCommand()
$SQLCmdLogin.CommandText = "Select RequestService.Theme From dbo.RequestService where RequestID = $RequestIDServiceDesk"
$objReaderLogin = $SQLCmdLogin.ExecuteReader()
while ($objReaderLogin.Read()) 
      {
      $Global:ServiceDeskLogin = $objReaderLogin.GetValue(0) #Запись логина в переменную
      }
      $objReaderLogin.Close()
$SQLCmdPassword = $SQLConnection.CreateCommand()
$SQLCmdPassword.CommandText = "Select RequestService.PasswordRequest From dbo.RequestService where RequestID = $RequestIDServiceDesk"
$objReaderPassword = $SQLCmdPassword.ExecuteReader()
while ($objReaderPassword.Read()) 
      {
      $Global:ServiceDeskPassword = $objReaderPassword.GetValue(0) #Запись пароля в переменную
      }
$objReaderPassword.Close()
$SQLCmdFullName = $SQLConnection.CreateCommand()
$SQLCmdFullName.CommandText = "SELECT UPPER(LEFT(U.FirstName,1))+LOWER(SUBSTRING(U.FirstName,2,LEN(U.FirstName))) as FirstName,
             UPPER(LEFT(U.LastName,1))+LOWER(SUBSTRING(U.LastName,2,LEN(U.LastName))) as LastName FROM (
            SELECT LTRIM(RTRIM(T.FirstName)) as FirstName,
              LTRIM(RTRIM(T.LastName)) as LastName  FROM (
            SELECT SUBSTRING(R.UserNameCyrillic, 1,charindex(' ',R.UserNameCyrillic)) as FirstName,
                SUBSTRING(R.UserNameCyrillic, charindex(' ', R.UserNameCyrillic),len(R.UserNameCyrillic)-charindex(' ',R.UserNameCyrillic)+1) as LastName
            FROM RequestService as R Where RequestID = $RequestIDServiceDesk
            ) as T
            ) as U"
$objReaderFullName = $SQLCmdFullName.ExecuteReader()
while ($objReaderFullName.Read()) 
       {
       $Global:ServiceDeskGivenName = $objReaderFullName.GetValue(0) #Запись имени в переменную
       $Global:ServiceDeskSurname = $objReaderFullName.GetValue(1) #Запись фамилии в переменную
       }
$objReaderFullName.Close()
#Проверка верности данных
Write-Host "Имя пользователя на кириллице: $Global:ServiceDeskGivenName" -foregroundcolor "green"
Write-Host "Фамилия пользователя на кириллице: $Global:ServiceDeskSurname" -foregroundcolor "green"
Write-Host "Логин пользователя: $Global:ServiceDeskLogin" -foregroundcolor "green"
$userNameAccount = $Global:ServiceDeskGivenName + " " + $Global:ServiceDeskSurname
$upn = $Global:ServiceDeskLogin + "@ektu.kz"
Write-Host "UPN пользователя: $upn" -foregroundcolor "green"
Read-Host "Если все данные верны, то нажмите Enter"
$SQLConnection.Close() 

$ADLogin = $Global:ServiceDeskLogin
$ADUserNameAccount = $userNameAccount
$ADGivenName = $Global:ServiceDeskGivenName
$ADSurName = $Global:ServiceDeskSurname
$ADPassword = $Global:ServiceDeskPassword
while ($True)
{
$Case = Read-Host "Выберите тип заявки цифрой и нажмите Enter"
    Switch ($Case)
    {
        1{
        #Создание пользователя в Active Directory
        $PShellsession = New-PSSession -computername $DomainController -Credential $UserCredential
        Invoke-Command -Session $PShellsession -ScriptBlock {
        New-ADUser -Name $using:ADUserNameAccount `
        -GivenName $using:ADGivenName `
        -Surname $using:ADSurName `
        -SamAccountName $using:ADLogin `
        -DisplayName $using:ADUserNameAccount `
        -UserPrincipalName $using:upn `
        -Path "OU=Все,DC=ektu,DC=kz" -AccountPassword (ConvertTo-SecureString $using:ADPassword -AsPlainText -force) -Enabled $true -PasswordNeverExpires $true
        Write-Host "$using:ADLogin создан. Учетной записи в Active Directory у него не было " -foregroundcolor "green"
        }
        Disconnect-PSSession $PShellsession
        }
        2{
        #Создание пользователя в ServiceDesk
        Read-Host "Для создания пользователя в ServiceDesk нажмите Enter"
        $SQLServer = "SERVICEDESK\SQLSERVICE";
        $SQLCatalog = "ServiceDeskBase";
        $SQLLogin = "";
        $SQLPassword = "";
        $SQLConnection = New-Object System.Data.SqlClient.SqlConnection
        $SQLConnection.ConnectionString = "Server=$SQLServer; Database=$SQLCatalog; User ID=$SQLLogin; Password=$SQLPassword;"
        $SQLConnection.Open()
        $DepartmentID = '80'
        $IDRole= '1'
        $SQLCmdInsert = $SQLConnection.CreateCommand()
        $SQLCmdInsert.CommandText = "INSERT INTO [dbo].[Employee]([Login],[AutorName],[AutorSurname],[DepartmentID],[IDRole]) 
                                    VALUES ('$Global:ServiceDeskLogin',N'$Global:ServiceDeskGivenName',N'$Global:ServiceDeskSurname',$DepartmentID,$IDRole)"
        $objReaderInsert = $SQLCmdInsert.ExecuteNonQuery()
        $SQLConnection.Close()
        } 
        3{
        #Создание пользователя в MS Exchange
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://ExchangeSrv.ektu.kz/PowerShell/ -Authentication Kerberos -Credential $UserCredential
        Import-PSSession $Session –DisableNameChecking
        Enable-Mailbox -Identity $Global:ServiceDeskLogin -Database "MailDatabase" -RetentionPolicy "Default 6 Mounth Delete" -PrimarySmtpAddress $upn
        Remove-PSSession $session
        }
        4{
        #Создание пользователя в MS Lync
        $so = New-PSSessionOption -SkipRevocationCheck:$true -SkipCACheck:$true -SkipCNCheck:$true 
        $Sessionlync = New-PSSession -ConnectionUri https://lync2010.ektu.kz/ocspowershell -Credential $UserCredential -SessionOption $so
        Import-PSSession $Sessionlync
        Enable-CsUser –Identity $upn –RegistrarPool lync2010.ektu.kz –SipAddress "sip:$upn"
        Remove-PSSession $Sessionlync
        }
        5{
        #Создание пользователя в Google Workspace
        $PShellsession = New-PSSession -computername $DomainController -Credential $UserCredential
        Invoke-Command -Session $PShellsession -ScriptBlock {
        $NewUPN = $Using:ADLogin +"@edu.ektu.kz"
        Set-ADUser $Using:ADLogin –userPrincipalName $NewUPN }
        Disconnect-PSSession $PShellsession
        
        Import-Module PSGSuite
        $upnEdu = $Global:ServiceDeskLogin + "@edu.ektu.kz"
        
        Write-Host "Имя пользователя на кириллице: $Global:ServiceDeskGivenName" -foregroundcolor "green"
        Write-Host "Фамилия пользователя на кириллице: $Global:ServiceDeskSurname" -foregroundcolor "green"
        Write-Host "Логин пользователя: $Global:ServiceDeskLogin" -foregroundcolor "green"
        Write-Host "Пароль пользователя: $Global:ServiceDeskPassword" -foregroundcolor "green"
        Write-Host "UPN пользователя: $upnEdu" -foregroundcolor "green"
        Read-Host "Для продолжения нажмите enter"
        
        Import-Module PSGSuite
        $ConfigName = "GSuite"
        $P12KeyPath = "key.p12"
        $AppEmail = ""
        $AdminEmail = ""
        $CustomerID = ""
        $Domain = "edu.ektu.kz"
        $ServiceAccountClientID = ""
        $GooglePassword = ""
        Set-PSGSuiteConfig -ConfigName $ConfigName `
        -P12KeyPath $P12KeyPath -AppEmail $AppEmail `
        -AdminEmail $AdminEmail -CustomerID $CustomerID `
        -Domain $Domain  -ServiceAccountClientID $ServiceAccountClientID
        $Password = ConvertTo-SecureString -String $GooglePassword -AsPlainText -Force
        New-GSUser -PrimaryEmail $upnEdu -GivenName $Global:ServiceDeskGivenName -FamilyName $Global:ServiceDeskSurname -Password $Password -OrgUnitPath "/Новые сотрудники"
        Write-Host "Пользователь $NewUserUPN Google Workspace создан"
        }
        6{ 
        #Выход из цикла
        Write-Host "Спасибо, запускайте еще!"
        break exit 
        }
    }
}
:exit 