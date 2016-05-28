#Import du module Azure Active Directory module dans notre session PowerShell
Import-Module MSOnline

#Création de l’objet d’identification utilisateur
$credential = get-credential

#Etablissement de la connexion avec Azure Active Directory 
Connect-MsolService -Credential $credential

#Listage des propriété et méthode pour la commande 
Get-MsolUser | Get-Member

#Récupération des propriétés des utilisateurs
Get-MsolUser | Select UserPrincipalName, DisplayName, WhenCreated, Licenses

#Récupération de tous les abonnements achetés
Get-MsolSubscription 


