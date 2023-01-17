# voir ou existantes
#Get-ADOrganizationalUnit

# ===========================================
# Ajout Organizational Unit Utilisateurs
# New-ADorganizationalUnit -Name "Utilisateurs" -Path "DC=morgan,DC=lan"
# OK DONE !

# ===========================================
# Ajout Organization Unit Children Lille / Paris
# New-ADOrganizationalUnit -Name "Lille" -Path "OU=utilisateurs,DC=morgan,DC=lan"
# New-ADOrganizationalUnit -Name "Paris" -Path "OU=utilisateurs,DC=morgan,DC=lan"
# OK DONE !


# ===========================================
# Création de groupes de sécurité Direction / Informatique / Comptabilité
# New-ADGroup -Name "Direction" -GroupCategory Security -GroupScope DomainLocal -Path "OU=utilisateurs,DC=morgan,DC=lan" -Description "Groupe de sécurité pour les dirigeants du bureau de direction."
# New-ADGroup -Name "Informatique" -GroupCategory Security -GroupScope DomainLocal -Path "OU=utilisateurs,DC=morgan,DC=lan" -Description "Groupe de sécurité pour les informaticiens."
# New-ADGroup -Name "Comptabilite" -GroupCategory Security -GroupScope DomainLocal -Path "OU=utilisateurs,DC=morgan,DC=lan" -Description "Groupe de sécurité pour les membres de la comptabilité"
# OK DONE !


# ===========================================
# Création des users selon le fichiers users.csv et placement dans les bonnes OU et groupes de sécurité
# on importe le CSV en ignorant la 1ère ligne car elle contient uniquement le header
$users = Import-Csv -Path ".\users.csv" -Delimiter ","

function createId {
    param (
        [Parameter(Mandatory)][String]$lastName,
        [Parameter(Mandatory)][String]$firstName 
    )

    $id = $lastName.toLower()[0] + $firstName.toLower()
    return $id
}

function get_city_path {
    param(
        [Parameter(Mandatory)][String]$city
    )

    switch( $city ){
         "Paris"{
            return "OU=Paris,OU=Utilisateurs,DC=formation,DC=local"
        }
         "Lille"{
            return "OU=Lille,OU=Utilisateurs,DC=morgan,DC=lan"
        }
        default {
            throw "Unrecognized city : $city"
        }
    }
}

function Remove-StringLatinCharacters
{
    PARAM ([string]$String)
    [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes($String))
}

function get_user_group{
    param(
        [Parameter(Mandatory)][String]$group
    )
    $groupWithoutAccent = Remove-StringLatinCharacters -String $group
    switch ($groupWithoutAccent){
        "Direction"{
            return "Direction"
        }
        "Comptabilitee"{
            return "Comptabilite"
        }
        "Informatique"{
            return "Informatique"
        }
        default {
            Write-Output "ERROR"
            Write-Output "$groupWithoutAccent"
            throw "Unrecognized Group : $group !"
        }

    }
}

function createAdUser {
    param(
        [Parameter(Mandatory)][String]$userId,
        [Parameter(Mandatory)][String]$group,
        [Parameter(Mandatory)][String]$city
    )
    $BUSINESS_DOMAIN = "formation.local"

    # we construct necessary informations
    $useremail = $userId + "@" + $BUSINESS_DOMAIN

    # we get the correct AD path for user creation in correct city
    $userPath = get_city_path -City $city

    # ========================================
    # Real addition of user
    # ========================================
    New-ADUser -Name $UserId -Path $userPath -OtherAttributes @{'mail'=$useremail}
    # ========================================

    $userGroup = get_user_group -Group $group
    # ========================================
    # Make the new user member group
    # ========================================
    Add-AdGroupMember -Identity $userGroup -Members $userId

}

foreach ($newUser in $users){
    $newUserFirstName = $newUser.Firstname
    $newUserLastName = $newUser.Lastname
    # $newUserCity = $newUser.Lastg
    # $newUserService = $newUser.Service

    # we create an userId with first letter of firstName and the rest with lastName, all in lowercase
    $newUserId = createId -LastName $newUserLastName -FirstName $newUserFirstName

    Write-Output "Creation of user for $newUserFirstName $newUserLastName"
    createAdUser -UserId $newUserId -Group $newUserService -City $newUserCity

}