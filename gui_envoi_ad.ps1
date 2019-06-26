###############################################################################
#                   SCRIPT D'ENVOI DE MESSAGE DE DOMAINE                      #
#                           0.2 26 JUIN 2019                                  #
#                           CLEMENT FOURSANS                                  #
###############################################################################


#Importation du module Active Directory
try {
    Import-Module ActiveDirectory
}

catch {
    [System.Windows.MessageBox]::Show("Impossible d'importer le module Active Directory. Merci de l'exécuter sur un serveur AD.","Erreur")
}

#Paramétrage de l'interface graphique####################################################################################
#########################################################################################################################

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '400,400'
$Form.text                       = "Formulaire d`'envoi de message"
$Form.TopMost                    = $false

$Label4                          = New-Object system.Windows.Forms.Label
$Label4.text                     = "Unite d`'Organisation"
$Label4.AutoSize                 = $true
$Label4.visible                  = $false
$Label4.width                    = 25
$Label4.height                   = 10
$Label4.location                 = New-Object System.Drawing.Point(11,84)
$Label4.Font                     = 'Microsoft Sans Serif,10'

$BoxOU                           = New-Object system.Windows.Forms.TextBox
$BoxOU.multiline                 = $false
$BoxOU.width                     = 294
$BoxOU.height                    = 20
$BoxOU.visible                   = $false
$BoxOU.location                  = New-Object System.Drawing.Point(7,114)
$BoxOU.Font                      = 'Microsoft Sans Serif,10'

$BoxComputer                     = New-Object system.Windows.Forms.TextBox
$BoxComputer.multiline           = $false
$BoxComputer.width               = 295
$BoxComputer.height              = 20
$BoxComputer.visible             = $false
$BoxComputer.location            = New-Object System.Drawing.Point(6,114)
$BoxComputer.Font                = 'Microsoft Sans Serif,10'

$Label3                          = New-Object system.Windows.Forms.Label
$Label3.text                     = "Poste"
$Label3.AutoSize                 = $true
$Label3.visible                  = $false
$Label3.width                    = 25
$Label3.height                   = 10
$Label3.location                 = New-Object System.Drawing.Point(10,84)
$Label3.Font                     = 'Microsoft Sans Serif,10'

$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "Option d`'envoi du message"
$Label1.AutoSize                 = $true
$Label1.width                    = 25
$Label1.height                   = 10
$Label1.location                 = New-Object System.Drawing.Point(6,15)
$Label1.Font                     = 'Microsoft Sans Serif,10'

$BoutonPoste                     = New-Object system.Windows.Forms.RadioButton
$BoutonPoste.text                = "Poste"
$BoutonPoste.AutoSize            = $true
$BoutonPoste.width               = 104
$BoutonPoste.height              = 20
$BoutonPoste.location            = New-Object System.Drawing.Point(10,52)
$BoutonPoste.Font                = 'Microsoft Sans Serif,10'

$BoutonOU                        = New-Object system.Windows.Forms.RadioButton
$BoutonOU.text                   = "Unite d`'Organisation"
$BoutonOU.AutoSize               = $true
$BoutonOU.width                  = 104
$BoutonOU.height                 = 20
$BoutonOU.location               = New-Object System.Drawing.Point(98,49)
$BoutonOU.Font                   = 'Microsoft Sans Serif,10'

$BoutonDomaine                   = New-Object system.Windows.Forms.RadioButton
$BoutonDomaine.text              = "Domaine"
$BoutonDomaine.AutoSize          = $true
$BoutonDomaine.width             = 104
$BoutonDomaine.height            = 20
$BoutonDomaine.location          = New-Object System.Drawing.Point(298,49)
$BoutonDomaine.Font              = 'Microsoft Sans Serif,10'

$Label2                          = New-Object system.Windows.Forms.Label
$Label2.text                     = "Contenu du message"
$Label2.AutoSize                 = $true
$Label2.width                    = 25
$Label2.height                   = 10
$Label2.location                 = New-Object System.Drawing.Point(10,162)
$Label2.Font                     = 'Microsoft Sans Serif,10'

$Button1                         = New-Object system.Windows.Forms.Button
$Button1.text                    = "Envoyer"
$Button1.width                   = 102
$Button1.height                  = 30
$Button1.location                = New-Object System.Drawing.Point(10,216)
$Button1.Font                    = 'Microsoft Sans Serif,10'

$BoxMsg                          = New-Object system.Windows.Forms.TextBox
$BoxMsg.multiline                = $false
$BoxMsg.width                    = 295
$BoxMsg.height                   = 20
$BoxMsg.location                 = New-Object System.Drawing.Point(7,189)
$BoxMsg.Font                     = 'Microsoft Sans Serif,10'

$Form.controls.AddRange(@($Label4,$BoxOU,$BoxComputer,$Label3,$Label1,$BoutonPoste,$BoutonOU,$BoutonDomaine,$Label2,$Button1,$BoxMsg))

#########################################################################################################################

#Paramétrage de l'affichage des boites de saisie de texte################################################################
#########################################################################################################################
$BoutonPoste.Add_Click(
{
$Label4.Visible = $false
$BoxOU.Visible = $false
$Label3.Visible = $true
$BoxComputer.Visible = $true
})

$BoutonOU.Add_Click(
{
$Label3.Visible = $false
$BoxComputer.Visible = $false
$Label4.Visible = $true
$BoxOU.Visible = $true
})

$BoutonDomaine.Add_Click(
{
$Label3.Visible = $false
$BoxComputer.Visible = $false
$Label4.Visible = $false
$BoxOU.Visible = $false
})

#########################################################################################################################

#Paramétrage de l'exécution au click sur le bouton "envoyer"#############################################################
#########################################################################################################################
$Button1.Add_click(   
{
#Si aucun RadioButton n'est coché, affichage d'une boite de dialogue
if (($BoutonPoste.Checked -match "False") -and ($BoutonOU.Checked -match "False") -and ($BoutonDomaine.Checked -match "False"))
{
[System.Windows.MessageBox]::Show("Veuillez selectionner une option d'envoi","Erreur")
}


else {
#Initialisation des variables du log
$StartTime = (Get-Date).ToShortDateString()+", "+(Get-Date).ToLongTimeString()
#Stockage du texte des boites de saisie dans une variable
$txt = $BoxComputer.Text
$msg = $BoxMsg.Text

#Si le bouton "Poste" est coché, le $Mode devient "Poste"
    if ($BoutonPoste.Checked -eq $true)
        {
        $Mode="Poste"
        }
#Si le bouton "OU" est coché, le $Mode devient "OU"
    elseif ($BoutonOU.Checked -eq $true)
        {
            try {
                #Récupération de la liste des ordinateurs de l'UO
                $Mode="OU"
                $ou = $BoxOU.Text
                $req = Get-ADComputer -SearchBase $ou -Filter {OperatingSystem -like '*Windows*'} | Select-Object -ExpandProperty name
            }

            catch {
                #Si l'UO n'existe pas, retourne un message d'erreur
                [System.Windows.MessageBox]::Show("Unité d'Organisation introuvable.","Erreur")
            }
    }

#Si le bouton "Domaine" est coché, le $Mode devient "Domaine"
    elseif ($BoutonDomaine.Checked -eq $true)
        {
        #Récupération de la liste des ordinateurs du Domaine
        $Mode="Domaine"
        $req = Get-ADComputer -Filter {OperatingSystem -like '*Windows*'} | Select-Object -ExpandProperty name
        }
    
    #La liste $comp est égale au contenu de la requête, ou au nom du Poste
    switch($Mode)
            {
            "Poste" { $comp = $txt }
            "OU" { $comp = $req}
            "Domaine" { $comp = $req}
            }
    #Pour chaque $Computer dans la liste $comp
    foreach($computer in $comp)
        {
        
        #Test de l'état de la machine sur le réseau
        $test = Test-Connection -CN $computer -Count 1 -BufferSize 16 -Quiet
        
        #Si la machine est en ligne
        if ($test -match $true)
        {
        #Envoi du message
        Write-Host "envoi $msg à $computer" -ForegroundColor Green
        Invoke-WmiMethod  -Path Win32_Process -Name Create -ArgumentList "msg * $msg" -ComputerName $computer
        $compon++
        }
        #Sinon, affichage en rouge
        else { Write-Host "$computer injoignable" -ForegroundColor Red }
        }
#Initialisation des variables de fin de log
$EndTime = (Get-Date).ToShortDateString()+", "+(Get-Date).ToLongTimeString()
$TimeTaken = New-TimeSpan -Start $StartTime -End $EndTime

#Création d'un objet dont les valeurs sont les paramètres
$log =  @{
        HeureDeDebut = $StartTime
        HeureDeFin = $EndTime
        TempsEcoule = $TimeTaken
}

$duree = New-Object psobject -Property $log
$result = @($duree.tempsecoule,$compon)
$result = $result -join " de durée totale`nNombre de postes atteints : "

#Affichage log avec le temps écoulé
[System.Windows.MessageBox]::Show($result , "Résultat")
#Remise à 0 des variables
$StartTime,$EndTime,$TimeTaken,$numbcomp=$null
}})
[void]$Form.ShowDialog()