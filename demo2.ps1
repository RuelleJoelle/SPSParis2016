
#Création de l’objet d’identification utilisateur
$credential = Get-Credential

##################################################################################################################################################################################################################
###### Demo Skype ######

Import-Module LyncOnlineConnector

#Connexion à Office 365 et ouverture d'une session à l’aide des informations dֹ’identification fournies.
$session = New-CsOnlineSession -Credential $credential

#Import de la session utilisateur dans la session courante
Import-PSSession $session -AllowClobber

#Récupération de propriétés de tous les utilisateurs ayant un compte Skype
Get-CsOnlineUser| select DisplayName, UserPrincipalName, Enabled, LastSyncTimeStamp, UsageLocation

#Export de toute les proprités utilisateurs dans un fichier CSV + tri par DisplayName
Get-CsOnlineUser| select DisplayName, UserPrincipalName, Enabled, LastSyncTimeStamp | Sort-Object DisplayName| Export-Csv c:\SPSParis\SkypeUsers.csv -NoTypeInformation

#Récupération du nom du tenant Skype 
Get-CsTenant | fl DisplayName

#Récupération des paramètres de configuration des réunions
Get-CsMeetingConfiguration


##################################################################################################################################################################################################################
###### Demo Exchange + Skype ###### 

#Initialisation d'une connexion persistante à Exchange
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $credential -Authentication Basic -AllowRed

#Import de la session dans la session courante
Import-PSSession $Session -AllowClobber

#Etablissement de la connexion avec Azure Active Directory 
Connect-MsolService -Credential $credential

#Requête de récupération des utilisateurs ayant une licence
$x = Get-MsolUser  | Where-Object {$_.isLicensed -eq "TRUE"}


foreach ($i in $x)
    {
      #Récupération d'informations sur la boîte aux lettres de l'utilisateur stocké dans l'objet $i
      $y = Get-Mailbox -Identity $i.UserPrincipalName
      
      If($y)
      {
        #Ajout de l'information IsMailboxEnabled dans l'objet $i
        $i | Add-Member -MemberType NoteProperty -Name IsMailboxEnabled -Value $y.IsMailboxEnabled
      }

      #Récupération d'informations sur le compte Lync de l'utilisateur stocké dans l'objet $i
      $y = Get-CsOnlineUser -Identity $i.UserPrincipalName

      #Ajout de l'information EnabledForSkype dans l'objet $i
      $i | Add-Member -MemberType NoteProperty -Name EnabledForSkype -Value $y.Enabled
    }

#Affichage des informations récupérées ci-dessus
$x | Select-Object DisplayName, IsLicensed, IsMailboxEnabled, EnabledForSkype


##################################################################################################################################################################################################################
###### Demo creation signature personnalisée ######
 
#Paramétrage du répertoire de sauvegarde des signatures au format HTLM
$save_location = 'c:\SPSParis\Signatures\'

#Récupération de tous les utilisateurs 
$users = Get-MsolUser 
 
foreach ($user in $users) {
  $DisplayName= “$($user.DisplayName)”
  $title = "$($User.Title)"
  $MobilePhone = "$($User.MobilePhone)"
  $UserPrincipalName = "$($User.UserPrincipalName)"

  #Construction et sauvegarde de la signature au format HTML
  $output_file = $save_location + $DisplayName + ".html"
  Write-Host "Création de la signature au format html pour " $DisplayName
  "<span style=`"font-family: calibri,sans-serif;`"><strong> SPS Paris 2016 </strong><br /><strong>" + $DisplayName + "</strong><br />", $title + " - " + $MobilePhone + "<br />", $UserPrincipalName + "<br />", "</span><br />"| Out-File $output_file
  
  #$signHTML = (Get-Content $output_file)
  #Set-MailboxMessageConfiguration –Identity $user.UserPrincipalName -AutoAddSignature $True  -SignatureHtml   $signHTML
}

#Requête de récupération d'un utilisateur par son DisplayName
$Myuser = Get-MsolUser  | Where-Object {$_.DisplayName -eq "Joelle Ruelle"}
$MyuserDisplayName= “$($Myuser.DisplayName)”
#Affectation de la signature à un utilisateur à partir de son fichier HTML
$output_file_user = $save_location + $MyuserDisplayName + ".html"
$signHTML = (Get-Content $output_file_user)
Write-Host "Affectation de la signature de " $MyuserDisplayName
Set-MailboxMessageConfiguration –Identity $Myuser.UserPrincipalName -AutoAddSignature $True  -SignatureHtml   $signHTML

##################################################################################################################################################################################################################
###### Demo Outlook Web : Ajout de signature (exemple du yOS Lyon - Etienne Bailly ISTEP) ######

#Définition du fichier contenant la signature.
$fichHTML = "C:\SPSParis\signature.html"

#Récupération d'un utilisateur via son nom
$x = Get-MsolUser

#Affectation de la signature à l'utilisateur
 $x |ForEach { $signHTML = (Get-Content $fichHTML) -f $_.DisplayName, $_.Title,  $_.MobilePhone, $_.UserPrincipalName

Set-MailboxMessageConfiguration –Identity $_.UserPrincipalName -AutoAddSignature $True  -SignatureHtml $signHTML

}


##################################################################################################################################################################################################################
###### Demo Exchange ######

#Afficher des informations détaillées sur le trafic des messages dans l'organisation.
Get-MailTrafficReport -StartDate 05/01/2016 -EndDate 05/27/2016  | Format-Table Date,EventType,MessageCount, Direction


##################################################################################################################################################################################################################
###### Demo Exchange + PowerBI ######

Import-Module "C:\SPSParis\powerbi-powershell-modules-master\Modules\PowerBIPS" –Force

#Paramétrage du jeton d'authentification OAuth
$authToken = Get-PBIAuthToken -clientId "yourClientId" #clientId venant de Azure


#Paramétrage du schéma du dataset
 $dataSetSchema = @{
        name = "SPSParis"    
        ; tables = @(
            @{  name = "MailTrafficReport"
                ; columns = @( 
                    @{ name = "Date"; dataType = "DateTime"   }
                    , @{ name = "MessageCount"; dataType = "Int64"  }                  
                    ) 
            })
    }  
 
#Création du DataSet dans PowerBI   
$createdDataSet = New-PBIDataSet -authToken $authToken -dataSet $dataSetSchema -Verbose 

#Récupération des données sur le trafic des messages dans l'organisation.
$myDAtas = Get-MailTrafficReport -StartDate 01/01/2016 -EndDate 05/27/2016  | Select Date,MessageCount

# Insertion des données en masse
$myDAtas | Add-PBITableRows -authToken $authToken -dataSetName "SPSParis" -tableName "MailTrafficReport" -Verbose

# Insertion des données par lot de 5
$myDAtas | Add-PBITableRows -authToken $authToken -dataSetName "SPSParis" -tableName "MailTrafficReport" -batchSize 5 -Verbose

 
##################################################################################################################################################################################################################
#Deconnexion de Office 365
get-PSSession | remove-PSSession