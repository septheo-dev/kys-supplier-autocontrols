$SAPWindowTitle = "SAP DP5(1)/100 SAP Easy Acess - User Menu for "
Add-Type -AssemblyName System.Windows.Forms

# Chemin du fichier texte à lire
$cheminFichier = "PATH_TO_TEXT\texte.txt"

# Tableau pour stocker les noms de variables
$variables = @("raison_sociale", "rue", "numero_rue", "code_postal", "ville", "departement", "siren", "siret", "tva", "bic", "iban", "categorie_achat", "entite_juridique_1", "entite_juridique_2")

# Tableau pour stocker les valeurs des variables
$valeurs = @()

# Lecture du fichier ligne par ligne
foreach ($ligne in Get-Content $cheminFichier) {

    # Ajout de la valeur de la ligne courante au tableau des valeurs
    $valeurs += $ligne.Trim()
}

# Affectation des valeurs aux variables correspondantes
for ($i = 0; $i -lt $variables.Length; $i++) {

    # Affectation de la valeur de la variable correspondante
    Set-Variable -Name $variables[$i] -Value $valeurs[$i]
}

# Affichage des variables pour vérification
$variables | ForEach-Object { Write-Host "$_ = $($ExecutionContext.InvokeCommand.ExpandString("`$$_"))" }

$categorie = Read-Host "Entrez la categorie d'achat"
$entite_juridique_1 = Read-Host "Entrez l'entité juridique 1 :"
$entite_juridique_2 = Read-Host "Entrez l'entité juridique 2 :"

#VARIABLES :

# Ouvrir l'application SAP à partir de l'URL
#$URL = "C:\Users\Public\Desktop\SAPHIR.url"
#Start-Process $URL -Wait
Start-Sleep -Seconds 7

# Écrit "BP" dans la zone de recherche
[System.Windows.Forms.SendKeys]::SendWait("BP")

# Attend que le texte soit saisi
Start-Sleep -Milliseconds 1000

# Clic sur la touche "Entrée"
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")

Start-Sleep -Seconds 2

# Envoyer la combinaison de touches Ctrl + F5
[System.Windows.Forms.SendKeys]::SendWait("^{F5}")

Start-Sleep -Seconds 2

# Envoyer la combinaison de touches TAB
[System.Windows.Forms.SendKeys]::SendWait("{TAB}")

Start-Sleep -Milliseconds 125

[System.Windows.Forms.SendKeys]::SendWait("ZK14")

Start-Sleep -Milliseconds 1000

##################

[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")

Start-Sleep -Milliseconds 500

[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")

Start-Sleep -Milliseconds 500

1..9 | ForEach-Object { [System.Windows.Forms.SendKeys]::SendWait("{UP}") }

Start-Sleep -Milliseconds 500
##################

[System.Windows.Forms.SendKeys]::SendWait("ZK")

Start-Sleep -Milliseconds 2500

[System.Windows.Forms.SendKeys]::SendWait("{TAB}")

Start-Sleep -Milliseconds 1500

[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")

Start-Sleep -Milliseconds 1500

[System.Windows.Forms.SendKeys]::SendWait("00")

Start-Sleep -Milliseconds 500

[System.Windows.Forms.SendKeys]::SendWait("{TAB}")

Start-Sleep -Milliseconds 500

[System.Windows.Forms.SendKeys]::SendWait($raison_sociale)

Start-Sleep -Milliseconds 500

1..4 | ForEach-Object { [System.Windows.Forms.SendKeys]::SendWait("{TAB}") }

[System.Windows.Forms.SendKeys]::SendWait("SUEZ HQ")

[System.Windows.Forms.SendKeys]::SendWait("{TAB}")

[System.Windows.Forms.SendKeys]::SendWait("SUEZ HQ")

[System.Windows.Forms.SendKeys]::SendWait("{TAB}")

Start-Sleep -Milliseconds 500

[System.Windows.Forms.SendKeys]::SendWait($rue)

Start-Sleep -Milliseconds 250		

[System.Windows.Forms.SendKeys]::SendWait("{TAB}")

Start-Sleep -Milliseconds 100

[System.Windows.Forms.SendKeys]::SendWait($numero_rue)

Start-Sleep -Milliseconds 250

[System.Windows.Forms.SendKeys]::SendWait("{TAB}")

Start-Sleep -Milliseconds 250

[System.Windows.Forms.SendKeys]::SendWait($code_postal)

Start-Sleep -Milliseconds 250

[System.Windows.Forms.SendKeys]::SendWait("{TAB}")

Start-Sleep -Milliseconds 250

[System.Windows.Forms.SendKeys]::SendWait($ville)

Start-Sleep -Milliseconds 250

[System.Windows.Forms.SendKeys]::SendWait("{TAB}")

Start-Sleep -Milliseconds 250

[System.Windows.Forms.SendKeys]::SendWait("FR")

Start-Sleep -Milliseconds 250

[System.Windows.Forms.SendKeys]::SendWait("{TAB}")

Start-Sleep -Milliseconds 250

[System.Windows.Forms.SendKeys]::SendWait($departement)

Start-Sleep -Milliseconds 250

[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")

Start-Sleep -Milliseconds 4000

[System.Windows.Forms.SendKeys]::SendWait("{F7}")

Start-Sleep -Milliseconds 1500

[System.Windows.Forms.SendKeys]::SendWait("{F7}")

Start-Sleep -Milliseconds 2000

1..6 | ForEach-Object { [System.Windows.Forms.SendKeys]::SendWait("{TAB}") }

Start-Sleep -Milliseconds 250

[System.Windows.Forms.SendKeys]::SendWait($categorie)

Start-Sleep -Milliseconds 250

[System.Windows.Forms.SendKeys]::SendWait("{TAB}")

Start-Sleep -Milliseconds 250

[System.Windows.Forms.SendKeys]::SendWait(" ")

Start-Sleep -Milliseconds 250

1..10 | ForEach-Object { [System.Windows.Forms.SendKeys]::SendWait("{TAB}") }

Start-Sleep -Milliseconds 250

Start-Sleep -Milliseconds 100

Start-Sleep -Milliseconds 100

[System.Windows.Forms.SendKeys]::SendWait("FR1")

Start-Sleep -Milliseconds 100

[System.Windows.Forms.SendKeys]::SendWait("{TAB}")

Start-Sleep -Milliseconds 100

[System.Windows.Forms.SendKeys]::SendWait($siren)

Start-Sleep -Milliseconds 100

[System.Windows.Forms.SendKeys]::SendWait("{TAB}")

Start-Sleep -Milliseconds 100

[System.Windows.Forms.SendKeys]::SendWait("FR3")

Start-Sleep -Milliseconds 500

[System.Windows.Forms.SendKeys]::SendWait("{TAB}")

Start-Sleep -Milliseconds 100

[System.Windows.Forms.SendKeys]::SendWait($siret)

Start-Sleep -Milliseconds 500

[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")

Start-Sleep -Milliseconds 3000

################

[System.Windows.Forms.SendKeys]::SendWait("{F7}")

Start-Sleep -Milliseconds 4000

[System.Windows.Forms.SendKeys]::SendWait("{F7}")

Start-Sleep -Milliseconds 2000

1..5 | ForEach-Object { [System.Windows.Forms.SendKeys]::SendWait("{TAB}") }

Start-Sleep -Milliseconds 500

[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")

Start-Sleep -Milliseconds 1500

[System.Windows.Forms.SendKeys]::SendWait($iban)

Start-Sleep -Milliseconds 250

[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")

Start-Sleep -Milliseconds 3000

[System.Windows.Forms.SendKeys]::SendWait("{TAB}")

Start-Sleep -Milliseconds 1500

[System.Windows.Forms.SendKeys]::SendWait("{F2}")

Start-Sleep -Milliseconds 1500

1..5 | ForEach-Object { [System.Windows.Forms.SendKeys]::SendWait("{TAB}") }

Start-Sleep -Milliseconds 250

[System.Windows.Forms.SendKeys]::SendWait($bic)

Start-Sleep -Milliseconds 250

[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")

Start-Sleep -Milliseconds 1500

###############

###############

[System.Windows.Forms.SendKeys]::SendWait("^{F3}")

Start-Sleep -Milliseconds 1500

1..4 | ForEach-Object { [System.Windows.Forms.SendKeys]::SendWait("{TAB}") }

Start-Sleep -Milliseconds 100

# Ajout de la logique pour tester l'entit� 1
if ($entite_juridique_1) {
    if ($entite_juridique_1 -eq "G100" -or $entite_juridique_1 -eq "G200") {
        # Si l'entit� 1 est G100 ou G200
        [System.Windows.Forms.SendKeys]::SendWait($entite_juridique_1)
        Start-Sleep -Milliseconds 100
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
        Start-Sleep -Milliseconds 1500

        
        1..2 | ForEach-Object { [System.Windows.Forms.SendKeys]::SendWait("{TAB}") }
        
        Start-Sleep -Milliseconds 750

        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")

        Start-Sleep -Milliseconds 300


    } else {
        # Si l'entit� 1 est E100
        [System.Windows.Forms.SendKeys]::SendWait($entite_juridique_1)
        Start-Sleep -Milliseconds 300
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
        Start-Sleep -Milliseconds 1500
        [System.Windows.Forms.SendKeys]::SendWait("EUR")
        Start-Sleep -Milliseconds 500
        [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
        Start-Sleep -Milliseconds 500
        [System.Windows.Forms.SendKeys]::SendWait("Z060")
        Start-Sleep -Milliseconds 500
        [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
        Start-Sleep -Milliseconds 500
        [System.Windows.Forms.SendKeys]::SendWait("DDP")
        Start-Sleep -Milliseconds 500
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")

        Start-Sleep -Milliseconds 750

        1..3 | ForEach-Object { [System.Windows.Forms.SendKeys]::SendWait("{UP}") }

        Start-Sleep -Milliseconds 100

        1..3 | ForEach-Object { [System.Windows.Forms.SendKeys]::SendWait("{TAB}") }

        Start-Sleep -Milliseconds 100

        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")

        Start-Sleep -Milliseconds 500

        1..2 | ForEach-Object { [System.Windows.Forms.SendKeys]::SendWait("{TAB}") }

        
    }
}

Start-Sleep -Milliseconds 1000

# Ajout de la logique pour tester l'entit� 2
if ($entite_juridique_2) {
    if ($entite_juridique_2 -eq "G100" -or $entite_juridique_2 -eq "G200") {
        # Si l'entit� 2 est G100 ou G200
        [System.Windows.Forms.SendKeys]::SendWait($entite_juridique_2)
        Start-Sleep -Milliseconds 100
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
        Start-Sleep -Milliseconds 1500
        # ... Ajoutez le reste de vos �tapes pour G200 ...
    } else {
        # Si l'entit� 2 est E100
        [System.Windows.Forms.SendKeys]::SendWait($entite_juridique_2)
        Start-Sleep -Milliseconds 100
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
        Start-Sleep -Milliseconds 1500
        [System.Windows.Forms.SendKeys]::SendWait("EUR")
        Start-Sleep -Milliseconds 500
        [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
        Start-Sleep -Milliseconds 500
        [System.Windows.Forms.SendKeys]::SendWait("Z060")
        Start-Sleep -Milliseconds 500
        [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
        Start-Sleep -Milliseconds 500
        [System.Windows.Forms.SendKeys]::SendWait("DDP")
        Start-Sleep -Milliseconds 500
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
    }
}
