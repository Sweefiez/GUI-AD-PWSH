# Program containing the functions necessary for the proper operation of the HMI for updating the AD

# This function is used to update landline telephone numbers in the AD.

function MAJFixes{
    # these lines are used to search your computer files for the document containing the fixed numbers

    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
    $FileBrowser.Filter = "csv file (*.csv) |*.csv"
    $FileBrowser.InitialDirectory = "C:\"
    $FileBrowser.ShowDialog()

    # These lines are used to retrieve the file name and path so that it can be processed in any 
    # where the file is stored
    $FixePath = $FileBrowser.FileName
    $FixeFolder = ($FixePath.Split('\')[-1])
    $FolderFile = $FixePath.Substring(0,($FixePath.Length-$FixeFolder.Length))
    
    # Here we check if the path is not null, if the path is null then we don't execute the program
    if ($FixePath -notlike $null){

        # Ici je sélectionne que les colonnes qui sont utiles aux traitement pour après et je les renomme pour faciliter 
        #le traitement
        Import-Csv -Path $FixePath -Delimiter ";" | 
        Select-Object @{ expression={$_.'="Nom"'}; label="Nom"},
        @{ expression={$_.'="Prénom"'}; label="Prenom"},
        @{ expression={$_.'="numéro public"'}; label="numero_public" },
        @{ expression={$_.'="numéro père"'}; label="numero_pere" },
        @{ expression={$_.'="numéro interne"'}; label="numero_interne" } | 
        Export-Csv -Path "$FolderFile\Temp_Result.CSV" -Delimiter ";"-Encoding UTF8 -NoTypeInformation

        # Here I select only the columns that are necessary for the processing and rename them to facilitate the processing
        Import-Csv -Path $FixePath -Delimiter ";" | 
        Select-Object @{ expression={$_.'="Nom"'}; label="Nom"},
        @{ expression={$_.'="Prénom"'}; label="Prenom"},
        @{ expression={$_.'="numéro public"'}; label="numero_public" },
        @{ expression={$_.'="numéro père"'}; label="numero_pere" },
        @{ expression={$_.'="numéro interne"'}; label="numero_interne" } | 
        Export-Csv -Path "$FolderFile\Temp_Result.CSV" -Delimiter ";" -Encoding UTF8 -NoTypeInformation

        # Here I import the temporary CSV file created with only the necessary information
        $newcsv = Import-Csv -Path "$FolderFile\Temp_Result.CSV" -Delimiter ";"
        foreach($row in $newcsv){
            # This loop removes unnecessary characters to make it more readable and easier to manipulate the data
            $row.Nom = $row.Nom -replace '="','' -replace '"',''
            $row.numero_interne = $row.numero_interne -replace '="','' -replace '"',''
            $row.numero_pere = $row.numero_pere -replace '="','' -replace '"',''
            $row.numero_public = $row.numero_public -replace '="','' -replace '"',''
            $row.Prenom = $row.Prenom -replace '="','' -replace '"',''
            # The line below adds a column named "nom_AD" to easily handle the information in AD processing
            $row | Add-Member -MemberType NoteProperty -Name "nom_AD" -Value ($row.Prenom+" "+$row.Nom)
            # This line takes the first letter of the first name and the full name as the SamAccountName in AD for easier processing later
            $row.nom_AD = $row.nom_AD -replace '" "',' '
        }

        # Export the data with the new assignments made in the above loop
        $newcsv | Export-Csv -Path "$FolderFile\Result.CSV" -Delimiter ";" -Encoding UTF8 -NoTypeInformation

        # Remove the temporary file created above
        Remove-Item -Path "$FolderFile\Temp_Result.CSV"
        # Here we retrieve the processed file to perform the update in AD
        $file_maj_fixes = Import-Csv -Path "$FolderFile\Result.CSV" -Delimiter ";"
        # This loop takes each line from the file and checks if the username exists in AD, if it does, it checks if a primary number is specified in the corresponding column. If a primary number is specified, it adds it; otherwise, it adds the secondary number.
        Foreach($User in $file_maj_fixes){
            $Nom_AD = $User.nom_AD
            if (Get-ADUser -Filter "Name -like ""$Nom_AD"""){
                $num_pere = $User.numero_pere
                $phone_public = $User.numero_public
                if ($phone_public -like $null -and $num_pere -like $null){
                    $ipPhone = $User.numero_interne
                    $identity = (Get-ADUser -Filter "Name -like ""$Nom_AD""").SamAccountName
                    Set-ADUser -Identity $identity -Replace @{ipPhone = $ipPhone}
                    #Write-Host "Person : $Nom_AD, Tel ip : $ipPhone"
                }
                if ($phone_public -like $null -and $num_pere -notlike $null){
                    $ipPhone = $User.numero_pere
                    $identity = (Get-ADUser -Filter "Name -like ""$Nom_AD""").SamAccountName
                    Set-ADUser -Identity $identity -Replace @{ipPhone = $ipPhone}
                    #Write-Host "Person : $Nom_AD, Tel ip : $ipPhone"
                }                                                                                                                                                                                          
                if ($phone_public -notlike $null -and $num_pere -like $null){                                                                                                                                                           
                    $OffPhone = $User.numero_public
                    $ipPhone = $User.numero_interne
                    $identity = (Get-ADUser -Filter "Name -like ""$Nom_AD""").SamAccountName
                    Set-ADUser -Identity $identity -Replace @{ipPhone = $ipPhone} -OfficePhone $OffPhone
                    #Write-Host "Person : $Nom_AD, Tel ip : $ipPhone, Office Phone : $OffPhone"
                }
                if($phone_public -notlike $null -and $num_pere -notlike $null){
                    $OffPhone = $User.numero_public
                    $ipPhone = $User.numero_pere
                    $identity = (Get-ADUser -Filter "Name -like ""$Nom_AD""").SamAccountName
                    Set-ADUser -Identity $identity -Replace @{ipPhone = $ipPhone} -OfficePhone $OffPhone
                    #Write-Host "Person : $Nom_AD, Tel ip : $ipPhone, Office Phone : $OffPhone"
                }
            }
        }
    }
        # Deletes the modified file so that there is no trace of it on the machine of the person executing the program.
        Remove-Item -Path "$FolderFile\Result.CSV"
}

# This loop updates the mobile phone numbers in AD

function MAJMobiles{
    # This program selects a CSV file, extracts the necessary data, and saves it to a new CSV file.
    # Then, the program uses the information from this new file to update the mobile phone information of users in the active directory.

    # Creating a dialog window to select the CSV file
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
    $FileBrowser.Filter = "csv file (*.csv) |*.csv"
    $FileBrowser.InitialDirectory = "C:\"
    $FileBrowser.ShowDialog()

    # Storing the selected file path and name
    $MobilesPath = $FileBrowser.FileName
    $MobilesFolder = ($MobilesPath.Split('\')[-1])
    $FolderFile = $MobilesPath.Substring(0,($MobilesPath.Length-$MobilesFolder.Length))

    # Checking if the file was correctly selected
    if ($MobilesPath -notlike $null){
        # Extraction des données du fichier CSV sélectionné et stockage dans un nouveau fichier CSV
        Import-Csv -Path $MobilesPath -Delimiter ";" -Encoding Default| `
        Select-Object @{ expression={$_."nom utilisateur"}; label='nom_utilisateur' },@{ expression={$_."numéro"}; label='numero_mobile' } | Where-Object {$_ -notmatch ' - '} | `
        Where-Object {$_ -notmatch 'BE'} | `
        Export-Csv -Path "$FolderFile\Temp_Result_Tel_Mobiles.CSV" -Delimiter ";" -Encoding UTF8 -NoTypeInformation

        # Loading the data from the new CSV file into memory
        $ResultMobilesCSV = Import-csv -Path "$FolderFile\Temp_Result_Tel_Mobiles.CSV" -Delimiter ";" -Encoding Default

        # Loop that reads each line of the imported CSV file
        $i = 0
        foreach ($row in $ResultMobilesCSV){
            # The variable $name retrieves the username from each line of the file
            $name = $ResultMobilesCSV.nom_utilisateur
            # Splitting the name and checking its length, if it is greater than 2, then the condition is true and the code is executed
            if ((($name[$i].split(" ")).Length -gt 2) -eq $true){
                # Checking the number of uppercase words because all surnames are in uppercase.
                # Here it checks if there are 3, 2, or 1 uppercase words in order to place the first name first followed by the last name
                if (((($name[$i] -split " ") -cmatch '[A-Z]{2}').Length -eq 3) -eq $true){
                    $row | Add-Member -MemberType NoteProperty -Name "nom_AD" -Value (($name[$i].split(" "))[-1]+" "+($name[$i].split(" "))[0]+" "+($name[$i].split(" "))[1]+" "+($name[$i].split(" "))[2]) -Force
                }
                if (((($name[$i] -split " ") -cmatch '[A-Z]{2}').Length -eq 2) -eq $true){
                    $row | Add-Member -MemberType NoteProperty -Name "nom_AD" -Value (($name[$i].split(" "))[-1]+" "+($name[$i].split(" "))[0]+" "+($name[$i].split(" "))[1]) -Force
                }
                if (((($name[$i] -split " ") -cmatch '[A-Z]{2}').Length -eq 1) -eq $true){
                    $row | Add-Member -MemberType NoteProperty -Name "nom_AD" -Value (($name[$i].split(" "))[1]+" "+($name[$i].split(" "))[-1]+" "+($name[$i].split(" "))[0]) -Force
                }
            }
            # Reverse the first name and last name to have the first name first
            if ((($name[$i].split(" ")).Length -eq 2) -eq $true){
                $row | Add-Member -MemberType NoteProperty -Name "nom_AD" -Value (($name[$i].split(" "))[1]+" "+($name[$i].split(" "))[0]) -Force
            }
            if((($name[$i]).split(" ")).length -eq 1){
                $row | Add-Member -MemberType NoteProperty -Name "nom_AD" -Value ($name[$i]) -Force
            }
            $i += 1
        }

        # This line saves the modified information into a CSV file named "result_PARC20122022.csv"
        $ResultMobilesCSV | Export-csv "$FolderFile\Result_Tel_Mobiles.CSV" -Encoding UTF8 -Force -NoTypeInformation

        # This line removes the temporary file created at the beginning of the script to avoid unnecessary files
        Remove-item -Path "$FolderFile\Temp_Result_Tel_Mobiles.CSV"

        # Import the previously processed file
        $file_maj_mobiles = Import-Csv -Path "$FolderFile\Result_Tel_Mobiles.CSV" -Delimiter "," -Encoding Default
        # The $Nom_AD variable retrieves the name from the "nom_ad" column, and the $Phone variable retrieves the phone number from the "numero_mobile" column
        for($j=0;$j -lt ($file_maj_mobiles.Length);$j++){
            $Nom_AD = ($file_maj_mobiles.nom_AD)[$j]
            $Phone = ($file_maj_mobiles.numero_mobile)[$j]
            # Check if the person exists in the AD, if so, execute the code
            if (Get-ADUser -Filter "Name -like ""$Nom_AD"""){
                # The $identity variable dynamically retrieves the user's identity, which consists of the first letter of the first name followed by the first 7 letters of the last name
                $identity = (Get-ADUser -Filter "Name -like ""$Nom_AD""").SamAccountName
                # Update the mobile phone field in the AD
                Set-ADUser -Identity $identity -MobilePhone $Phone
                #Write-Host "$Nom_AD -- $Phone" 
            }     
        }
    }
    # Delete the modified file to leave no trace on the machine of the person executing the program
    Remove-item -Path "$FolderFile\Result_Tel_Mobiles.CSV"
}

function MAJEmployes{

    # Create a window to choose a file
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
    # Filter the displayed files to show only CSV files
    $FileBrowser.Filter = "csv file (*.csv) |*.csv"
    # Default directory when opening the selection window
    $FileBrowser.InitialDirectory = "C:\"
    # Show the window to choose a file
    $FileBrowser.ShowDialog()

    # Store the full path of the selected file
    $EmploiPath = $FileBrowser.FileName

    # If a file has been selected
    if ($EmploiPath -notlike $null){
        # Read the CSV file using a delimiter of ";"
        $EmploiCSV = Import-csv -Path $EmploiPath -Delimiter ";" -Encoding Default
        # For each row in the CSV file
        foreach($row_emploi in $EmploiCSV){
            # Combine the first name and last name of the person into a full name
            $nom_AD = $row_emploi.Prénom +" "+ $row_emploi.Nom
            # Store the person's job
            $emploi = $row_emploi.Emploi
            # Store the person's company
            $societe = $row_emploi.Société
            $Nsociete = $societe -replace ' S.A.S','' -replace 'SAS ',''
            $Date_depart = $row_emploi.'Date départ'
            $Date_today = Get-Date -Format 'dd/MM/yyyy'

            # Search for the person in Active Directory using their full name
            if (Get-ADUser -Filter "Name -like ""$nom_AD""")
            {
                if($Date_depart -like $null){
                    # Store the username of the found person in Active Directory
                    $identity = (Get-ADUser -Filter "Name -like ""$nom_AD""").SamAccountName
                    # Update the person's job and description in Active Directory
                    Set-ADUser -Identity $identity -Title $emploi -Description $emploi -Company $Nsociete
                    #Write-Host "$nom_AD --- $emploi --- $Nsociete"
                }
                if($Date_depart -notlike $null){
                    $Date_depart_obj = [datetime]::ParseExact($Date_depart, "dd/MM/yyyy", $null)
                    $Date_today_obj = [datetime]::ParseExact($Date_today, "dd/MM/yyyy", $null)
                    if($Date_depart_obj -ge $Date_today_obj){
                        # Store the username of the found person in Active Directory
                        $identity = (Get-ADUser -Filter "Name -like ""$nom_AD""").SamAccountName
                        # Update position and person description in Active Directory
                        Set-ADUser -Identity $identity -Title $emploi -Description $emploi -Company $Nsociete
                        #Write-Host "$nom_AD --- $emploi --- $Nsociete"
                    }
                }
            }
        }
    }
}
