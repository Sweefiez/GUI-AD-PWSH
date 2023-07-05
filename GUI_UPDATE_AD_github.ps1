Set-ExecutionPolicy -Force Unrestricted

# This loop checks whether a session is open. As long as the session is not open, it continuously asks for the admin login and password to connect to a domain controller.
while($dc.state -ne 'Opened'){

# The lines below display a window to fill in a text field to enter your admin login 
$form = New-Object System.Windows.Forms.Form
$form.Size = New-Object System.Drawing.Size(200,150)
$form.Text = "Login admin"

# Create a text field
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Size = New-Object System.Drawing.Size(100,20)
$textBox.Location = New-Object System.Drawing.Point(15,45)

# Create a button
$button_ok = New-Object System.Windows.Forms.Button
$button_ok.AutoSize = $true
$button_ok.Location = New-Object System.Drawing.Point(95,80)
$button_ok.Text = "Valider"
$button_ok.Add_Click({$form.Close()})

# Create a label
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Size(10, 15)
$label.AutoSize = $true
$label.Text = "Entrez votre login d'admin :"

# Adds the widgets defined above 
$form.Controls.Add($textBox)
$form.Controls.Add($label)
$form.Controls.Add($button_ok)

# Display window
$form.ShowDialog() | Out-Null

# The variable $login recovers the login written in the text field of the window created above, the variable $password displays a window to write the password securely.
$login = $textBox.Text
$password = Read-Host -Prompt "Entrez votre mot de passe" -AsSecureString
# This variable $credential creates a login and password pair in order to connect to the domain controller.
$credential = New-Object System.Management.Automation.PSCredential -ArgumentList $login, $password

# Define a list of machine names
$machines = "DC01", "DC02", "DC03"

# Initialize the variable that will store the name of the switched-on machine
$poweredOnMachine = $null

# For each machine in the list
foreach ($machine in $machines) {
    # Try to connect to the machine using WMI
  try {
    # If the connection is successful, the machine is switched on
    if((Test-Connection $machine -Count 1 -Quiet) -eq "True"){
        # Store the name of the machine switched on in the variable
        $poweredOnMachine = $machine
        # Exit the loop
        break
    }
  } catch {
    # If the connection fails, the machine is not switched on
    # Go to the next machine in the list
  }
}

# Allows you to create a session on the lit domain controller defined above, in order to import the AD module and make any necessary changes. 
$dc = New-PSSession –ComputerName $poweredOnMachine -Credential $credential
}
# Imports the module named ActiveDirectory from the domain controller so that it can be updated in AD 
Import-Module –PSSession $dc –Name ActiveDirectory -Force

# Imports the modules contained in the Functions.ps1 file
Import-Module -Name C:\0-projet_phpXpepsi\Liste_tel\Fonctions.ps1 -Force
# Window creation
$window = New-Object System.Windows.Forms.Form
$window.Text = "Mise à jour de l'Active Directory"
$window.Size = New-Object System.Drawing.Size(550,400)
$window.StartPosition = "CenterScreen"
$window.BackColor = [System.Drawing.Color]::FromArgb(0xEB, 0xEE, 0xEF)

# Create a button with "Update fixed numbers" written inside
$ButtonFixe = New-Object System.Windows.Forms.Button
$ButtonFixe.Location = New-Object System.Drawing.Size(80,200)
$ButtonFixe.Text = "Mise à jour des numéros fixes"
$ButtonFixe.BackColor = [System.Drawing.Color]::FromArgb(0x26, 0x46, 0x91)
$ButtonFixe.ForeColor = [System.Drawing.Color]::White
$ButtonFixe.AutoSize = $true

# Create a button with "Mobile number update" written inside
$ButtonMobile = New-Object System.Windows.Forms.Button
$ButtonMobile.Location = New-Object System.Drawing.Size(80,240)
$ButtonMobile.Text = "Mise à jour des numéros mobiles"
$ButtonMobile.BackColor = [System.Drawing.Color]::FromArgb(0x26, 0x46, 0x91)
$ButtonMobile.ForeColor = [System.Drawing.Color]::White
$ButtonMobile.AutoSize = $true

# Create a button with "Function update" written inside
$ButtonFonction = New-Object System.Windows.Forms.Button
$ButtonFonction.Location = New-Object System.Drawing.Size(310,200)
$ButtonFonction.Text = "Mise à jour de la fonction"
$ButtonFonction.BackColor = [System.Drawing.Color]::FromArgb(0x26, 0x46, 0x91)
$ButtonFonction.ForeColor = [System.Drawing.Color]::White
$ButtonFonction.AutoSize = $true

# Create a button with "Total update" written inside
$ButtonComplet.Location = New-Object System.Drawing.Size(310,240)
$ButtonComplet.Text = "Mise à jour complète"
$ButtonComplet.BackColor = [System.Drawing.Color]::FromArgb(0x26, 0x46, 0x91)
$ButtonComplet.ForeColor = [System.Drawing.Color]::White
$ButtonComplet.AutoSize = $true


# Creation of a button with "Validate" written inside
$confirmButton = New-Object System.Windows.Forms.Button
$confirmButton.Location = New-Object System.Drawing.Size(150,200)
$confirmButton.Text = "Valider"
$confirmButton.BackColor = [System.Drawing.Color]::FromArgb(0x26, 0x46, 0x91)
$confirmButton.ForeColor = [System.Drawing.Color]::White
$confirmButton.Visible = $false
$confirmButton.AutoSize = $true


# Creation of a button with "Back" written inside
$backButton = New-Object System.Windows.Forms.Button
$backButton.Location = New-Object System.Drawing.Size(250,200)
$backButton.Text = "Retour"
$backButton.BackColor = [System.Drawing.Color]::FromArgb(0x26, 0x46, 0x91)
$backButton.Visible = $false
$backButton.AutoSize = $true

# Create a button with "Quit" written inside
$quitButton = New-Object System.Windows.Forms.Button
$quitButton.Location = New-Object System.Drawing.Size(225,300)
$quitButton.Text = "Quitter"
$quitButton.BackColor = [System.Drawing.Color]::FromArgb(0xBB, 0x00, 0x4D)
$quitButton.ForeColor = [System.Drawing.Color]::White
$quitButton.AutoSize = $true

# Creating ToolTips for each button. ToolTips are small informative texts that appear when the mouse is held over the button for a few seconds.
$MAJFixeToolTip = New-Object System.Windows.Forms.ToolTip
$MAJFixeToolTip.SetToolTip($ButtonFixe, "Ce bouton permet de mettre à jour les téléphones fixe de chaque utilisateur de l'Active Directory")
$MAJMobileToolTip = New-Object System.Windows.Forms.ToolTip
$MAJMobileToolTip.SetToolTip($ButtonMobile, "Ce bouton permet de mettre à jour les téléphones mobiles de chaque utilisateur de l'Active Directory")
$MAJFonctionToolTip = New-Object System.Windows.Forms.ToolTip
$MAJFonctionToolTip.SetToolTip($ButtonFonction, "Ce bouton permet de mettre à jour les focntions de chaque utilisateur de l'Active Directory")
$MAJCompleteToolTip = New-Object System.Windows.Forms.ToolTip
$MAJCompleteToolTip.SetToolTip($ButtonComplet, "Ce bouton permet de mettre à jour tous les paramètres de chaque utilisateur de l'Active Directory")
$MAJQuitToolTip = New-Object System.Windows.Forms.ToolTip
$MAJQuitToolTip.SetToolTip($quitButton, "Ce bouton permet de quitter la fenêtre")

# Add an image at the top of the window to make a banner
$image1 = New-Object System.Windows.Forms.PictureBox
$image1.Location = New-Object System.Drawing.Size(5,0)
$image1.Size = New-Object System.Drawing.Size(520,210)
$image1.Image = [System.Drawing.Image]::FromFile('C:\Images\Pictures.png')

#Ajout d'une barre de progression pour visualiser l'avancement de la fonction
$ProgressBar = New-Object System.Windows.Forms.ProgressBar
$ProgressBar.Style = 'Continuous'
$ProgressBar.Minimum = 0
$ProgressBar.Maximum = 100

$ProgressBar.Width = 470
$ProgressBar.Height = 30
$ProgressBar.Top = 210
$ProgressBar.Left = 30

# Ajout d'une boite de texte afin de visualiser en temps réel les informations qui sont mise à jour
$TextBox = New-Object System.Windows.Forms.TextBox
$TextBox.Multiline = $true
$TextBox.ReadOnly = $true
$TextBox.ScrollBars = "Vertical"
$TextBox.Location = New-Object System.Drawing.Point(50,260)
$TextBox.Size = New-Object System.Drawing.Size(430,130)

# Add the above defibi controls to the window
$window.Controls.Add($ButtonFixe)
$window.Controls.Add($ButtonMobile)
$window.Controls.Add($ButtonFonction)
$window.Controls.Add($ButtonComplet)
$window.Controls.Add($confirmButton)
$window.Controls.Add($backButton)
$window.Controls.Add($quitButton)
$window.Controls.Add($image1)


# Event management
$ButtonFixe.Add_Click({
#Displays a confirmation message and shows confirmation and back buttons

$messageBox = [System.Windows.Forms.MessageBox]::Show("Êtes-vous sûr de vouloir mettre à jour les numéros de téléphones fixes ?", "Confirmation", [System.Windows.Forms.MessageBoxButtons]::YesNo)
if ($messageBox -eq "Yes") {
# Executes the MAJFixes function created in the "Functions" file
MAJFixes -ProgressBar $ProgressBar
[System.Windows.MessageBox]::Show("Opération réalisée avec succès !")
}
$confirmButton.Visible = $false
$backButton.Visible = $false
})

$ButtonMobile.Add_Click({
#Displays a confirmation message and shows confirmation and back buttons

$messageBox = [System.Windows.Forms.MessageBox]::Show("Êtes-vous sûr de vouloir mettre à jour les numéros de téléphones mobiles ?", "Confirmation", [System.Windows.Forms.MessageBoxButtons]::YesNo)
if ($messageBox -eq "Yes") {
# Executes the MAJMobiles function created in the "Functions" file
MAJMobiles -ProgressBar $ProgressBar
[System.Windows.MessageBox]::Show("Opération réalisée avec succès !")
}
$confirmButton.Visible = $false
$backButton.Visible = $false
})

$ButtonFonction.Add_Click({
#Displays a confirmation message and shows confirmation and back buttons

$messageBox = [System.Windows.Forms.MessageBox]::Show("Êtes-vous sûr de vouloir mettre à jour les fonctions des employés ?", "Confirmation", [System.Windows.Forms.MessageBoxButtons]::YesNo)
if ($messageBox -eq "Yes") {
# Executes the MAJEmployes function created in the "Functions" file
MAJEmployes -ProgressBar $ProgressBar
[System.Windows.MessageBox]::Show("Opération réalisée avec succès !")
}
$confirmButton.Visible = $false
$backButton.Visible = $false
})

$ButtonComplet.Add_Click({
#Displays a confirmation message and shows confirmation and back buttons

$messageBox = [System.Windows.Forms.MessageBox]::Show("Êtes-vous sûr de vouloir faire une mise à jour complète ?", "Confirmation",[System.Windows.Forms.MessageBoxButtons]::YesNo)
if ($messageBox -eq "Yes") {
# Execute the MAJFixes, MAJMobiles and MAJEmployes functions created in the "Functions" file, asking for the right file in a pop-up window at the start of each function execution.
[System.Windows.MessageBox]::Show("Choisissez le ficheir Mitel svp")
MAJFixes -ProgressBar $ProgressBar
[System.Windows.MessageBox]::Show("Choisissez le ficheir Orange svp")
MAJMobiles -ProgressBar $ProgressBar
[System.Windows.MessageBox]::Show("Choisissez le ficheir RH svp")
MAJEmployes -ProgressBar $ProgressBar
[System.Windows.MessageBox]::Show("Opération réalisée avec succès !")
}
$confirmButton.Visible = $false
$backButton.Visible = $false
})

$confirmButton.Add_Click({
# Executes the action associated with the selected button

Write-Host "Action associée au bouton sélectionné"
$confirmButton.Visible = $false
$backButton.Visible = $false
})

$backButton.Add_Click({
#Return to selection window

$confirmButton.Visible = $false
$backButton.Visible = $false
})

$quitButton.Add_Click({
# Close the window
$messageBox = [System.Windows.Forms.MessageBox]::Show("Êtes-vous sûr de vouloir quitter la fenêtre ?", "Confirmation", [System.Windows.Forms.MessageBoxButtons]::YesNo)
if ($messageBox -eq "Yes") {
# Executes the Close action to close the window
$window.Close()
}
$confirmButton.Visible = $false
$backButton.Visible = $false
})
# Displays the window

$window.StartPosition = "CenterScreen"
$window.ShowDialog()

# Disconnect the session created on the domain controller
Remove-PSSession -Session $dc
