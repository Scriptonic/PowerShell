#Isaac Estell
#6/15/2022
#
#Creates WOL Links based off a given company directory. Has the functionality to Update, Create, Delete links if the user is new, gets a differnet pc, or is leaving
#the company. Names the files based off of the hostname which follows a certain naming protocol. Also pulls the users MAC as it is required to make the link.

function createShortcut{
    Param(
        $hostname,
        $MAC,
        $fullName
        )

    if(-not (($hostname -eq "") -or ($MAC -eq ""))){
        #File Paths
        $shortcutExePath = '\\ww-file16-01\IT\BootUpLinks\program\wakemeonlan-x64\WakeMeOnLan.exe'  
        $shortcutFilePath = '\\ww-file16-01\IT\BootUpLinks\'

        #Calculate the hostname without WW- or -PC
        $hostname = $hostname.ToUpper()
        $username = $hostname.Replace("WW-", "").Replace("-LT", "").Replace("-PC", "")

        #Format the MAC address
        $MAC = $MAC.Replace('-', '').ToUpper()

        #Check if file already exists. If false, creates the shortcut file
        if((Test-Path -Path ($shortcutFilePath + $username + ".lnk")) -eq $false)
        {
            $SourceFilePath = $shortcutExePath
            $WScriptObj = New-Object -ComObject ("WScript.Shell")
            $shortcut = $WscriptObj.CreateShortcut($shortcutFilePath + $username + ".lnk")
            $shortcut.Description = $fullName
            $shortcut.TargetPath = ($SourceFilePath)
            $shortcut.Arguments = ("/wakeup " + $MAC)
            $shortcut.Save()
        }
    }
}

function getShortcutInfo{

    #Pulling all users from SeatTable who have both a MAC and ComputerName(HostName)
    $connString = "Data Source=WW-SEAT-SVR;Initial Catalog=WW-SEAT-DB;User ID=Seat;Password=Chartblock38"
    $userInfo = Invoke-SQLcmd -Query "SELECT Name, ComputerName, MAC From SeatTable Where MAC != '' and ComputerName != ''" -ConnectionString $connString
    $existingLinks = Get-ChildItem -Path "\\ww-file16-01\IT\BootUpLinks\" -Filter '*.lnk'
    $errorUsers = [System.Collections.ArrayList]::new()
    $count = 1

    #Loops through and writes to host if the file already exists
    foreach($user in $userInfo){
        
        Write-Host("User: $count")
        $count = $count + 1

        #Gets the current hostname from seattable
        $hostname = $user.ComputerName.Replace("WW-", "").Replace("-PC", "").Replace("-LT", "").ToUpper()
        $needsFileCreated = $true

        #Loops through all current links and checks if there are any already created for the user.
        foreach($existingLink in $existingLinks){

            #Get the mac, name, and description of the existing link
            $Shell = New-Object -ComObject WScript.Shell
            $existingLinkDescription = $Shell.CreateShortcut($existingLink.FullName).Description
            $existingLinkMAC = $Shell.CreateShortcut($existingLink.FullName).Arguments.Replace("/wakeup ", "")
            $existingLinkName = $existingLink.Name.Replace(".lnk", "").ToUpper()
           
            #If the $existingLinkName equals $hostname then a file has already been created for the user. Then the program runs a couple more 
            #checks to see if the file needs to be deleted and then recreated due to updated information
            if($existingLinkName -eq $hostname){

                $hostname
                $existingLinkName
                $existingLinkMAC
                $user.MAC.Replace("-", "").ToUpper()


                #If the current MAC equals the existing files MAC then the file does not need to be updated and it exists
                #Else delete the old link, and move on to the next file.
                if($existingLinkMAC -eq $user.MAC.Replace("-", "").ToUpper()){
                    Write-Host("Does not need a new file.`n")
                    $needsFileCreated = $false
                }
                else{
                    Remove-Item -Path ("\\ww-file16-01\IT\BootUpLinks\" + $existingLink.Name)
                    "Deleted an incorect file for" + $user.Name | Out-File -FilePath "\\ww-file16-01\IT\BootUpLinks\logs\logs.txt" -Append   
                }
                                                                                                  
            }
            elseif($existingLinkDescription -eq $user.Name){
                Remove-Item -Path ("\\ww-file16-01\IT\BootUpLinks\" + $existingLink.Name)
                "Deleted an old file for " + $user.Name | Out-File -FilePath "\\ww-file16-01\IT\BootUpLinks\logs\logs.txt" -Append
            }   
        }

        #If $needsFileCreated is true, then create a new link
        if($needsFileCreated){

            $user.ComputerName
            $user.Name
            $user.MAC
            Write-Host ("File created.`n")

            createShortcut -hostname $user.ComputerName -MAC $user.MAC -fullName $user.Name
            "Shortcut created for " + $user.Name | Out-File -FilePath "\\ww-file16-01\IT\BootUpLinks\logs\logs.txt" -Append
        }
    }

    #Print all users that had errors to a text file
    $errorUsers | Out-File -FilePath "\\ww-file16-01\IT\BootUpLinks\errors\errors.txt"
}
getShortcutInfo
