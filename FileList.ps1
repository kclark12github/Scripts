param([string]$Root="$($env:COMPUTERNAME)",[string]$BackupFolder,[string]$LogPath)
function Format-Elapsed {
    Param($Start, $End)
    $Elapsed = ""
    $ts = New-TimeSpan -start $Start -end $End
    if ($ts.Days -gt 0) {$Elapsed += "$($ts.Days) Days, "}
    if ($ts.Hours -gt 0) {$Elapsed += "$($ts.Hours) Hours, "}
    if ($ts.Minutes -gt 0) {$Elapsed += "$($ts.Minutes) Minutes, "}
    $Elapsed += "{0}.{1:000} Seconds" -f $ts.Seconds, $ts.Milliseconds
    return $Elapsed
}
function Get-TargetFolder {
  param ($Folder)
    if ($Folder.Name.StartsWith("Act of War")) {return "Atari\Act Of War\" + $Folder.Name}
    if ($Folder.Name.StartsWith("LEGO")) {return "Traveller's Tales\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Akalabeth")) {return "Origin Systems\Ultima\Akalabeth"}
    if ($Folder.Name.StartsWith("Alpha Centauri")) {return "2k Games\Civilization\Sid Meier's Alpha Centuri"}
    if ($Folder.Name.StartsWith("Arena")) {return "is Software\Quake\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Army Men - Toys in Space")) {return "2K Games\Army Men\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Army Men II")) {return "2K Games\Army Men\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Army Men RTS")) {return "2K Games\Army Men\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Army Men")) {return "2K Games\Army Men\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Assassins Creed")) {return "EA Games\Assassin's Creed\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Atlantic Fleet")) {return "Killerfish Games\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Aven Colony")) {return "Team17 Digital\" + $Folder.Name}
    if ($Folder.Name.StartsWith("AvP Classic")) {return "Fox Interactive\Alien vs Predator\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Baldur's Gate")) {return "Beamdog\Baldur's Gate\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Battle Chess")) {return "Interplay\Battle Chess\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Battlespire")) {return "Bethesda Softworks\Elder Scrolls\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Battlestar Galactica Deadlock")) {return "Slitherine\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Battlezone")) {return "Activision\Battlezone\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Beyond Zork")) {return "Infocom\Zork\" + $Folder.Name}
    if ($Folder.Name.StartsWith("BioShock")) {return "2K Games\BioShock\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Blood 2")) {return "id Software\Blood\Blood II - The Chosen"}
    if ($Folder.Name.StartsWith("Blood")) {return "id Software\Blood\Blood"}
    if ($Folder.Name.StartsWith("Caesar IV")) {return "Sierra Studios\Caesar\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Caesar 3")) {return "Sierra Studios\Caesar\Caesar III"}
    if ($Folder.Name.StartsWith("Caesar II")) {return "Sierra Studios\Caesar\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Caesar")) {return "Sierra Studios\Caesar\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Call To Power 2")) {return "2K Games\Civilization\Civilization - Call To Power"}
    if ($Folder.Name.StartsWith("Carmageddon Max Pack")) {return "Interplay\Carmageddon\Carmageddon"}
    if ($Folder.Name.StartsWith("CarmageddonMaxDamage")) {return "Interplay\Carmageddon\Carmageddon Max Damage"}
    if ($Folder.Name.StartsWith("Carmageddon")) {return "Interplay\Carmageddon\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Castles")) {return "Interplay\Castles\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Champions of Krynn")) {return "Strategic Simulations Inc\D&D Krynn Series\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Chinese Chess")) {return "Interplay\Battle Chess\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Civilization III Complete")) {return "2K Games\Civilization\Civilization III\Civilization III Complete"}
    if ($Folder.Name.StartsWith("Civilization IV Complete")) {return "2K Games\Civilization\Civilization IV\Civilization IV Complete"}
    if ($Folder.Name.StartsWith("Close Combat")) {return "Slitherine\Close Combat\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Commandos")) {return "Kalypso Media Digital\Commandos\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Curse of the Azure Bonds")) {return "Strategic Simulations Inc\AD&D Forgotten Realms\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Daggerfall")) {return "Bethesda Softworks\Elder Scrolls\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Dark Sun")) {return "Strategic Simulations Inc\D&D Dark Sun Series\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Dead Space")) {return "EA Games\Dead Space\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Death Knights of Krynn")) {return "Strategic Simulations Inc\D&D Krynn Series\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Delta Force")) {return "THQ\Delta Force\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Descent")) {return "Interplay\Descent\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Deus Ex")) {return "Square Enix\Deus Ex\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Diablo")) {return "Blizzard\Diablo\" + $Folder.Name}
    if ($Folder.Name.StartsWith("DOOM 3 BFG")) {return "id Software\Doom\Doom 3 BFG (GOG)"}
    if ($Folder.Name.StartsWith("DOOM 2")) {return "id Software\Doom\Doom 2 (GOG)"}
    if ($Folder.Name.StartsWith("DOOM")) {return "id Software\Doom\Doom (GOG)"}
    if ($Folder.Name.StartsWith("Dragonshpere")) {return "MicroProse\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Dungeon Hack")) {return "Strategic Simulations Inc\AD&D Forgotten Realms\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Dungeon Keeper")) {return "EA Games\Dungeon Keeper\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Dungeons")) {return "Kalypso Media Digital\Dungeons\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Earth 2150")) {return "TopWare\Earth 2150\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Emperor - Rise of the Middle Kingdom")) {return "Sierra Studios\Empire Earth\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Empire Earth")) {return "Sierra Studios\Empire Earth\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Eye of the Beholder")) {return "Strategic Simulations Inc\AD&D Forgotten Realms\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Falcon")) {return "Spectrum Holobyte\Falcon\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Fallout")) {return "Bethesda Softworks\Fallout\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Far Cry")) {return "UbiSoft\Far Cry\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Final DOOM")) {return "id Software\Doom\Final Doom"}
    if ($Folder.Name.StartsWith("Freespace")) {return "Interplay\Descent\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Gateway to the Savage Frontier")) {return "Strategic Simulations Inc\AD&D Forgotten Realms\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Ground Control")) {return "Sierra Studios\Ground Control\" + $Folder.Name}
    if ($Folder.Name.StartsWith("HC")) {return "UbiSoft\Heroes Chronicles\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Heroes of Might and Magic V")) {return "UbiSoft\Heroes of Might and Magic\HOMM 5"}
    if ($Folder.Name.StartsWith("HOMM")) {return "UbiSoft\Heroes of Might and Magic\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Hillsfar")) {return "Strategic Simulations Inc\AD&D Forgotten Realms\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Homeworld")) {return "Sierra Studios\Homeworld\" + $Folder.Name}
    if ($Folder.Name.StartsWith("HOMM")) {return "UbiSoft\Heroes of Might and Magic\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Icewind Dale Enhanced Edition")) {return "Beamdog\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Independence War")) {return "Infogrames\Independence War\" + $Folder.Name}
    if ($Folder.Name.StartsWith("JA2 - Unfinished Business")) {return "Sir Tech\Jagged Alliance\Jagged Alliance 2\Jagged Alliance 2 Unfinished Business"}
    if ($Folder.Name.StartsWith("Jagged Alliance 2 Wildfire")) {return "Sir Tech\Jagged Alliance\Jagged Alliance 2\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Jagged Alliance 2")) {return "Sir Tech\Jagged Alliance\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Jagged Alliance - DG")) {return "Sir Tech\Jagged Alliance\Jagged Alliance\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Jagged Alliance")) {return "Sir Tech\Jagged Alliance\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Leisure Suit Larry")) {return "Sierra Studios\Leisure Suit Larry\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Lords of the Fallen")) {return "CI Games\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Martian Dreams")) {return "Origin Systems\Ultima\Ultima Worlds of Adventure 2"}
    if ($Folder.Name.StartsWith("Master of Orion")) {return "MicroProse\Master of Orion\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Medal of Honor")) {return "EA Games\Medal of Honor\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Menzoberranzan")) {return "Strategic Simulations Inc\AD&D Forgotten Realms\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Metal Fatigue")) {return "Night Dive Studios\Metal Fatigue\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Might and Magic")) {return "UbiSoft\Might and Magic\" + $Folder.Name}
    if ($Folder.Name.Contains("Monkey Island")) {return "LucasArts\Monkey Island\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Morrowind")) {return "Bethesda Softworks\Elder Scrolls\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Myst")) {return "UbiSoft\Myst\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Neverwinter Nights")) {return "Beamdog\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Oblivion")) {return "Bethesda Softworks\Elder Scrolls\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Pandora First Contact")) {return "Slitherine\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Panzer General")) {return "Strategic Simulations Inc\Panzer General\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Pharaoh")) {return "Sierra Studios\Pharaoh\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Pirates")) {return "MicroProse\Pirates\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Planetfall")) {return "Infocom\Zork\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Planetscape")) {return "Beamdog\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Pool of Radiance")) {return "Strategic Simulations Inc\AD&D Forgotten Realms\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Pools of Darkness")) {return "Strategic Simulations Inc\AD&D Forgotten Realms\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Prince of Persia")) {return "UbiSoft\Prince of Persia\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Privateer 2")) {return "Origin Systems\Wing Commander\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Quake III")) {return "id Software\Quake\Quake 3"}
    if ($Folder.Name.StartsWith("Quake II")) {return "id Software\Quake\Quake 2"}
    if ($Folder.Name.StartsWith("Quake")) {return "id Software\Quake\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Railroad Tycoon")) {return "Gathering of Developers\Railroad Tycoon\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Ravenloft")) {return "Strategic Simulations Inc\D&D Ravenloft Series\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Realms of Arkania")) {return "Sir Tech\Realms of Arkania\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Red Faction")) {return "THQ\Red Faction\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Redguard")) {return "Bethesda Softworks\Elder Scrolls\" + $Folder.Name}
    if ($Folder.Name.StartsWith("RiME")) {return "Grey Box\" + $Folder.Name}
    if ($Folder.Name.Contains("Wolfenstein")) {return "id Software\Castle Wolfenstein\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Return To Zork")) {return "Infocom\Zork\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Riven")) {return "UbiSoft\Myst\" + $Folder.Name}
    if ($Folder.Name.StartsWith("RollerCoaster Tycoon")) {return "MicroProse\RollerCoaster Tycoon\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Secret of the Silver Blades")) {return "Strategic Simulations Inc\AD&D Forgotten Realms\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Sid Meier's Covert Action")) {return "MicroProse\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Sid Meier's Pirates")) {return "MicroProse\Pirates\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Sid Meier's Railroads")) {return "2K Games\Railroads!"}
    if ($Folder.Name.StartsWith("Silent Hunter")) {return "Strategic Simulations Inc\Silent Hunter\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Silent Service")) {return "MicroProse\Silent Service\" + $Folder.Name}
    if ($Folder.Name.StartsWith("SimCity")) {return "EA Games\SimCity\" + $Folder.Name}
    if ($Folder.Name.StartsWith("SOMA")) {return "Frictional Games\" + $Folder.Name}
    if ($Folder.Name.Contains("Spear of Destiny")) {return "id Software\Castle Wolfenstein\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Star Trek - Judgment Rites")) {return "Interplay\Star Trek - Judgement Rites\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Star Trek - Starfleet Academy")) {return "Interplay\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Star Trek - Starfleet Command")) {return "Interplay\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Star Trek 25th Anniversary")) {return "Interplay\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Star Wars - Battlefront")) {return "LucasArts\Star Wars - Battlefront\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Star Wars - Empire At War")) {return "LucasArts\Star Wars - Empire At War\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Star Wars - KotOR")) {return "LucasArts\Star Wars - Knights of the Old Republic"}
    if ($Folder.Name.StartsWith("Star Wars - KotOR2")) {return "LucasArts\Star Wars - Knights of the Old Republic 2 - The Sith Lords"}
    if ($Folder.Name.StartsWith("Star Wars - Rebel Assault")) {return "LucasArts\Star Wars - Rebel Assault\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Star Wars - Rogue Squadron 3D")) {return "LucasArts\Star Wars - Rogue Squadron"}
    if ($Folder.Name.StartsWith("Star Wars - TIE Fighter")) {return "LucasArts\Star Wars - TIE Fighter\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Star Wars - X-Wing")) {return "LucasArts\Star Wars - X-Wing\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Star Wars - XvT")) {return "LucasArts\Star Wars - X-Wing vs TIE Fighter\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Star Wars Jedi Knight - Mysteries of the Sith")) {return "LucasArts\Star Wars - Jedi Knight\Star Wars - Jedi Knight - Dark Forces 2\Mysteries of the Sith"}
    if ($Folder.Name.StartsWith("Star Wars Jedi Knight")) {return "LucasArts\Star Wars - Jedi Knight\" + $Folder.Name}
    if ($Folder.Name.StartsWith("STAR WARS The Force Unleashed")) {return "LucasArts\Star Wars - The Force Unleashed\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Starship Titanic")) {return "Completely Unexpected Productions\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Stronghold - Kingdom Simulator")) {return "Stormfront Studios\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Sudden Strike")) {return "Kalypso Media Digital\Sudden Strike\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Sunless Sea")) {return "Failbetter Games\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Surviving Mars")) {return "Paradox Interactive\" + $Folder.Name}
    if ($Folder.Name.StartsWith("The Bard's Tale")) {return "Interplay\Bard's Tale, The\" + $Folder.Name}
    if ($Folder.Name.StartsWith("The Bards Tale")) {return "Interplay\Bard's Tale, The\" + $Folder.Name}
    if ($Folder.Name.StartsWith("The Dark Queen of Krynn")) {return "Strategic Simulations Inc\D&D Krynn Series\" + $Folder.Name}
    if ($Folder.Name.StartsWith("The Lion King")) {return "Disney\" + $Folder.Name}
    if ($Folder.Name.StartsWith("The Savage Empire")) {return "Origin Systems\Ultima\Ultima Worlds of Adventure 1\" + $Folder.Name}
    if ($Folder.Name.Contains("Settlers")) {return "UbiSoft\" + $Folder.Name}
    if ($Folder.Name.Contains("Witcher")) {return "CD Projekt Red\Witcher, The\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Thief")) {return "Square Enix\Thief\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Tomb Raider")) {return "Eidos Interactive\Tomb Raider\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Total Annihilation")) {return "Wargaming.net\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Treasures of the Savage Frontier")) {return "Strategic Simulations Inc\AD&D Forgotten Realms\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Turok2 EX")) {return "Night Dive Studios\Turok\Turok 2 - Seeds of Evil"}
    if ($Folder.Name.StartsWith("Turok")) {return "Night Dive Studios\Turok\Turok - Dinosaur Hunter"}
    if ($Folder.Name.StartsWith("Ultima 4 - Quest")) {return "Origin Systems\Ultima\Ultima 4"}
    if ($Folder.Name.StartsWith("Ultima 7 - Serpent")) {return "Origin Systems\Ultima\Ultima 7"}
    if ($Folder.Name.StartsWith("Ultima Underworld 2")) {return "Origin Systems\Ultima\Ultima Underworld II"}
    if ($Folder.Name.StartsWith("Ultima Underworld")) {return "Origin Systems\Ultima\Ultima Underworld I"}
    if ($Folder.Name.StartsWith("Ultima")) {return "Origin Systems\Ultima\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Unlimited Adventures")) {return "Strategic Simulations Inc\AD&D Forgotten Realms\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Unreal")) {return "Infogrames\Unreal\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Uru")) {return "UbiSoft\Myst\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Victory At Sea")) {return "Evil Twin Artworks\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Warcraft")) {return "Blizzard\Warcraft\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Wing Commander")) {return "Origin Systems\Wing Commander\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Wizardry 6")) {return "Sir Tech\Wizardry\Wizardry 05 - The Heart of the Maelstrom"}
    if ($Folder.Name.StartsWith("Wizardry 7 Gold")) {return "Sir Tech\Wizardry\Wizardry 07a - Wizardry Gold"}
    if ($Folder.Name.StartsWith("Wizardry 8")) {return "Sir Tech\Wizardry\Wizardry 08"}
    if ($Folder.Name.StartsWith("World in Conflict - Complete Edition")) {return "UbiSoft\World in Conflict\World in Conflict - Collector's Edition"}
    if ($Folder.Name.StartsWith("X-COM")) {return "MicroProse\X-COM\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Xenonauts")) {return "Goldhawk Interactive\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Zork Grand Inquisitor")) {return "Infocom\Zork\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Zork Nemesis")) {return "Infocom\Zork\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Zork Zero")) {return "Infocom\Zork\Zork 0 - The Revenge of Megaboz"}
    if ($Folder.Name.StartsWith("Zork 2")) {return "Infocom\Zork\Zork 2 - The Wizard of Frobozz"}
    if ($Folder.Name.StartsWith("Zork 3")) {return "Infocom\Zork\Zork 3 - The Dungeon Master"}
    if ($Folder.Name.StartsWith("Zork")) {return "Infocom\Zork\Zork 1 - The Great Underground Empire"}
    return $Folder.Name
}
function Send-Mail {
    Param($AppName, $EmailBody, $LogPath)
    $PW = ConvertTo-SecureString $env:SMTP_PW -AsPlainText -Force
    $Creds = New-Object System.Management.Automation.PSCredential ($env:SMTP_USER, $PW)
    $Server = $env:SMTP_ADDRESS+":"+$env:SMTP_PORT
    Send-MailMessage -From "$env:USERNAME <$env:My_EMAIL>" -To "$env:USERNAME <$env:My_EMAIL>" -Subject "$AppName Succeeded!" -Body $EmailBody -BodyAsHtml -Attachments $LogPath -SmtpServer $env:SMTP_ADDRESS -Port $env:SMTP_PORT -Credential $Creds
}
function Write-Log {
    Param($Message, $Path = ".")
    function TS {return "[{0:MM/dd/yy} {0:HH:mm:ss tt}]" -f (Get-Date)}
    Write-Message -Message "$(TS) $Message" -Path $Path
}
function Write-Message {
    Param($Message, $Path = ".")
    "$Message" | Tee-Object -FilePath $Path -Append | Write-Output
}
function Write-Separator {
    Param($Path = ".")
    #                      1         2         3         4         5         6         7         8
    #             12345678901234567890123456789012345678901234567890123456789012345678901234567890
    $Separator = "--------------------------------------------------------------------------------"
    Write-Message -Message $Separator -Path $Path
}
function Write-Files {
    Param($Path = ".", $OutFile, $LogPath)
    $errMessage = ""
    $Message = "Listing Files from $Path to $OutFile"
    Write-Message -Message $Message -Path $LogPath

    Get-ChildItem -Path "$Path" -Force -Recurse -ErrorVariable errMessage | Select Fullname,Length,CreationTime,LastWriteTime | Export-Csv -Path "$OutFile" -NoTypeInformation
    if (-not $errMessage.Equals("")) {Write-Message -Message $errMessage -Path $LogPath}
    "<br />" #Return break tags appended to message
}
.{
    $AppName = "FileList"
    $StartTime = Get-Date
    $Message = "[$AppName © $("{0:yyyy}" -f $StartTime), Ken Clark                       $("{0:MM/dd/yy} {0:hh:mm:ss tt}" -f $StartTime)]"

    if ($BackupFolder.Equals("")) {$BackupFolder = "$($Env:OneDrive)\Backups\$Root\"}  #$($Env:BackupRoot)\
    if ($LogPath.Equals("")) {$LogPath = "$($BackupFolder)$Root.FileList.log"}
    Write-Message -Message $Message -Path $LogPath
    $EmailBody = $Message.Replace("©", "&copy;") + "<br />"
    $Message = "Root: $Root; BackupFolder: $BackupFolder; LogPath: $LogPath;"
    Write-Message -Message $Message -Path $LogPath
    $EmailBody += "$Message<br /><br />"

    if ($Root.Contains("TEST")) {
        $EmailBody += Write-Files -Path "\\Alpha\Public\Pluralsight\Entity Framework Core 2 - Getting Started\03 - demos\M3 Demos\SamuraiApp\packages\Microsoft.Extensions.DependencyInjection.Abstractions.2.0.0\lib" -OutFile "$($BackupFolder)TEST\Test.csv" -LogPath $LogPath
    }

    if ($Root.Contains("NAS")) {
        $EmailBody += Write-Files -Path "\\Alpha\Public" -OutFile "$($BackupFolder)ALPHA\Alpha Public Files.csv" -LogPath $LogPath
        $EmailBody += Write-Files -Path "\\Alpha\Backups" -OutFile "$($BackupFolder)ALPHA\Alpha Backups Files.csv" -LogPath $LogPath
        #$EmailBody += Write-Files -Path "\\Bravo\Public" -OutFile "$($BackupFolder)BRAVO\Bravo Public Files.csv" -LogPath $LogPath
        $EmailBody += Write-Files -Path "\\Charlie\Public" -OutFile "$($BackupFolder)CHARLIE\Charlie Public Files.csv" -LogPath $LogPath
        $EmailBody += Write-Files -Path "\\Delta\Public" -OutFile "$($BackupFolder)DELTA\Delta Public Files.csv" -LogPath $LogPath
        $EmailBody += Write-Files -Path "\\Echo\Public" -OutFile "$($BackupFolder)ECHO\Echo Public Files.csv" -LogPath $LogPath
        $EmailBody += Write-Files -Path "\\Foxtrot\Public" -OutFile "$($BackupFolder)FOXTROT\Foxtrot Public Files.csv" -LogPath $LogPath
        $EmailBody += Write-Files -Path "\\Golf\Public" -OutFile "$($BackupFolder)GOLF\Golf Public Files.csv" -LogPath $LogPath
        $EmailBody += Write-Files -Path "\\Golf\Backups" -OutFile "$($BackupFolder)GOLF\Golf Backups Files.csv" -LogPath $LogPath
        $EmailBody += Write-Files -Path "\\Hotel\Public" -OutFile "$($BackupFolder)HOTEL\Hotel Public Files.csv" -LogPath $LogPath
        $EmailBody += Write-Files -Path "\\Hotel\Backups" -OutFile "$($BackupFolder)HOTEL\Hotel Backups Files.csv" -LogPath $LogPath
    }

    if ($Root.Contains("7X4L243")) {
        $EmailBody += Write-Files -Path "\\7X4L243\C" -OutFile "$($BackupFolder)7X4L243\7X4L243 C Files.csv" -EmailBody $EmailBody -LogPath $LogPath
        $EmailBody += Write-Files -Path "\\7X4L243\D" -OutFile "$($BackupFolder)7X4L243\7X4L243 D Files.csv" -EmailBody $EmailBody -LogPath $LogPath
        #$EmailBody += Write-Files -Path "\\7X4L243\E" -OutFile "$($BackupFolder)B54J71N E Files.txt" -EmailBody $EmailBody -LogPath $LogPath
    }

    if ($Root.Contains("AAIUCI4")) {
        $EmailBody += Write-Files -Path "\\AAIUCI4\C" -OutFile "$($BackupFolder)AAIUCI4\AAIUCI4 C Files.csv" -EmailBody $EmailBody -LogPath $LogPath
        $EmailBody += Write-Files -Path "\\AAIUCI4\D" -OutFile "$($BackupFolder)AAIUCI4\AAIUCI4 D Files.csv" -EmailBody $EmailBody -LogPath $LogPath
    }
    if ($Root.Contains("GGGSCP1")) {
        $EmailBody += Write-Files -Path "\\GGGSCP1\C" -OutFile "$($BackupFolder)GGGSCP1\GGGSCP1 C Files.csv" -EmailBody $EmailBody -LogPath $LogPath
        $EmailBody += Write-Files -Path "\\GGGSCP1\D" -OutFile "$($BackupFolder)GGGSCP1\GGGSCP1 D Files.csv" -EmailBody $EmailBody -LogPath $LogPath
    }
    
    $EndTime = Get-Date
    $Message = "`n$AppName Complete @ $("{0:hh:mm:ss tt}" -f $EndTime) (Elapsed: $(Format-Elapsed -Start $StartTime -End $EndTime))"
    Write-Message -Message $Message -Path $LogPath
    $EmailBody += "$Message<br /><br />"
    Write-Separator -Path $LogPath


    #write-output "$EmailBody"


    &"$PSScriptRoot\eMailResults.ps1" -Subject "$Root.$AppName Complete" -Body "$EmailBody" -LogFile $LogPath -AsHTML
    #Send-Mail -AppName $AppName -EmailBody $EmailBody -LogPath $LogPath
}