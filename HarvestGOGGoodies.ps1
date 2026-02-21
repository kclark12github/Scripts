[cmdletbinding()]
Param([string]$Root,
    [string]$BackupFolder,
    [string]$LogPath)

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
function Old-Get-TargetFolder {
  param ($Folder)

    if ($Folder.Name.Contains("688(I) HunterKiller")) {return "Jane's Combat Simulations\Jane's Combat Simulations - 688(i) Hunter-Killer"}
    if ($Folder.Name.Contains("Close Combat")) {return "Close Combat\" + $Folder.Name}





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
function Get-TargetFolder {
  param ($Folder)

    if ($Folder.Name.Contains("688(I) HunterKiller")) {return "Jane's Combat Simulations\Jane's Combat Simulations - 688(i) Hunter-Killer"}
    if ($Folder.Name.Contains("Act of War")) {return "Act of War\" + $Folder.Name}

    if ($Folder.Name.Contains("Curse of the Azure Bonds")) {return "AD&D Forgotten Realms\" + $Folder.Name}
    if ($Folder.Name.Contains("Demon Stone")) {return "AD&D Forgotten Realms\" + $Folder.Name}
    if ($Folder.Name.Contains("Dungeon Hack")) {return "AD&D Forgotten Realms\" + $Folder.Name}
    if ($Folder.Name.Contains("Eye of the Beholder")) {return "AD&D Forgotten Realms\" + $Folder.Name}
    if ($Folder.Name.Contains("Gateway to the Savage Frontier")) {return "AD&D Forgotten Realms\" + $Folder.Name}
    if ($Folder.Name.Contains("Hillsfar")) {return "AD&D Forgotten Realms\" + $Folder.Name}
    if ($Folder.Name.Contains("Icewind Dale")) {return "AD&D Forgotten Realms\" + $Folder.Name}
    if ($Folder.Name.Contains("Menzoberranzan")) {return "AD&D Forgotten Realms\" + $Folder.Name}
    if ($Folder.Name.Contains("Pool of Radiance")) {return "AD&D Forgotten Realms\" + $Folder.Name}
    if ($Folder.Name.Contains("Pools of Darkness")) {return "AD&D Forgotten Realms\" + $Folder.Name}
    if ($Folder.Name.Contains("Secret of the Silver Blades")) {return "AD&D Forgotten Realms\" + $Folder.Name}
    if ($Folder.Name.Contains("Treasures of the Savage Frontier")) {return "AD&D Forgotten Realms\" + $Folder.Name}
    if ($Folder.Name.Contains("Unlimited Adventures")) {return "AD&D Forgotten Realms\" + $Folder.Name}
   
    if ($Folder.Name.Contains("Agatha Christie")) {return "Agatha Christie\" + $Folder.Name}
    if ($Folder.Name.Contains("Age of Wonders")) {return "Age of Wonders\" + $Folder.Name}
    if ($Folder.Name.Contains("Army Men")) {return "Army Men\" + $Folder.Name}
    if ($Folder.Name -eq "Assassins Creed") {return "Assassin's Creed\Assassin's Creed"}
    if ($Folder.Name.Contains("Assassin's Creed")) {return "Assassin's Creed\" + $Folder.Name.Replace("Assassins","Assassin's")}
    if ($Folder.Name.Contains("Alien vs Predator") -or $Folder.Name.Contains("AvP")) {return "Alien vs Predator\" + $Folder.Name}

    if ($Folder.Name.StartsWith("Dreams of Desire")) {return "AVF\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Helping the Hotties")) {return "AVF\" + $Folder.Name}
    if ($Folder.Name.StartsWith("House Party")) {return "AVF\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Inside Jennifer")) {return "AVF\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Leap of Faith")) {return "AVF\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Lust Academy")) {return "AVF\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Lust Epidemic")) {return "AVF\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Lust Theory")) {return "AVF\" + $Folder.Name}
    if ($Folder.Name.StartsWith("LustyVerse Passion on Display")) {return "AVF\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Sisterly Lust")) {return "AVF\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Summer with Mia Season 1")) {return "AVF\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Treasure of Nadia")) {return "AVF\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Unreal Lust Theory")) {return "AVF\" + $Folder.Name}

    if ($Folder.Name.Contains("B-17 Flying Fortress")) {return "B-17 Flying Fortress\" + $Folder.Name}
    if ($Folder.Name.Contains("Baldur's Gate")) {return "Baldur's Gate\" + $Folder.Name}

    if ($Folder.Name -eq "The Bards Tale" -or $Folder.Name -eq "The Bard's Tale") {return "Bard's Tale, The\Bard's Tale, The"}
    if ($Folder.Name -eq "The Bard's Tale Trilogy") {return "Bard's Tale, The\Bard's Tale Trilogy, The"}
    if ($Folder.Name -eq "The Bard's Tale IV Director's Cut") {return "Bard's Tale, The\Bard's Tale IV Director's Cut, The"}

    if ($Folder.Name.Contains("Batman Arkham")) {return "Batman\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Battle Chess") -or $Folder.Name.StartsWith("Chinese Chess")) {return "Battle Chess\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Battlefield 1942")) {return "Battlefield 1942\" + $Folder.Name}
    if ($Folder.Name.Contains("Battlestar Galactica")) {return "Battlestar Galactica\" + $Folder.Name}
    if ($Folder.Name.Contains("Battlezone")) {return "Battlezone\" + $Folder.Name}
    if ($Folder.Name.Contains("BioShock")) {return "BioShock\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Blood")) {return "Blood\" + $Folder.Name}
    if ($Folder.Name.Contains("Borderlands")) {return "Borderlands\" + $Folder.Name}
    if ($Folder.Name.Contains("Bridge Constructor")) {return "Bridge Constructor\" + $Folder.Name}
    if ($Folder.Name.Contains("Caesar")) {return "Caesar\" + $Folder.Name}
    if ($Folder.Name.Contains("Call To Power 2")) {return "Call To Power\" + $Folder.Name}
    if ($Folder.Name.Contains("Carmageddon")) {return "Carmageddon\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Castles")) {return "Castles\" + $Folder.Name}

    if ($Folder.Name -eq "Alpha Centauri") {return "Civilization\Sid Meier's Alpha Centauri"}
    if ($Folder.Name.Contains("Civilization") -or $Folder.Name -eq "Civ.Net") {return "Civilization\" + $Folder.Name}

    if ($Folder.Name.StartsWith("Close Combat")) {return "Close Combat\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Commandos")) {return "Commandos\" + $Folder.Name}
    if ($Folder.Name.Contains("Crysis")) {return "Crysis\" + $Folder.Name}

    if ($Folder.Name.Contains("D&D") -or $Folder.Name.Contains("Dragonshard") -or $Folder.Name.Contains("Stronghold - Kingdom Simulator")) {return "D&D\" + $Folder.Name}
    if ($Folder.Name.Contains("Dark Queen of Krynn")) {return "D&D Krynn Series\Dark Queen of Krynn, The"}
    if ($Folder.Name.Contains("Krynn")) {return "D&D Krynn Series\" + $Folder.Name}
    if ($Folder.Name.Contains("Dark Sun")) {return "D&D Dark Sun Series\" + $Folder.Name}
    if ($Folder.Name.Contains("Ravenloft")) {return "D&D Ravenloft Series\" + $Folder.Name}

    if ($Folder.Name.Contains("Dead Space")) {return "Dead Space\" + $Folder.Name}
    if ($Folder.Name.Contains("Delta Force")) {return "Delta Force\" + $Folder.Name}
    if ($Folder.Name.Contains("Descent") -or $Folder.Name.Contains("Freespace")) {return "Descent\" + $Folder.Name}
    if ($Folder.Name.Contains("Desperados")) {return "Desperados\" + $Folder.Name}
    if ($Folder.Name.Contains("Deus Ex")) {return "Deus Ex\" + $Folder.Name}
    if ($Folder.Name.Contains("Dishonored")) {return "Dishonored\" + $Folder.Name}

    if ($Folder.Name -eq "DOOM") {return "Doom\Doom (GOG)"}
    if ($Folder.Name.StartsWith("DOOM") -or $Folder.Name.StartsWith("Doom") -or $Folder.Name -eq "Final Doom") {return "Doom\" + $Folder.Name}

    if ($Folder.Name.StartsWith("Dragon Age")) {return "Dragon Age\" + $Folder.Name}
    if ($Folder.Name.Contains("Dungeon Keeper")) {return "Dungeon Keeper\" + $Folder.Name}
    if ($Folder.Name.Contains("Dungeon Siege")) {return "Dungeon Siege\" + $Folder.Name}
    if ($Folder.Name.Contains("Dungeons")) {return "Dungeons\" + $Folder.Name}
    if ($Folder.Name.Contains("Earth 2150")) {return "Earth 2150\" + $Folder.Name}
    if ($Folder.Name.Contains("Elder Scrolls") -or $Folder.Name.Contains("Battlespire") -or $Folder.Name.Contains("Daggerfall") -or $Folder.Name.Contains("Morrowind") -or $Folder.Name -eq "Oblivion" -or $Folder.Name.Contains("Redguard") -or $Folder.Name.Contains("Skyrim")) {return "Elder Scrolls\" + $Folder.Name}
    if ($Folder.Name.Contains("Empire Earth") -or $Folder.Name.Contains("Emperor - Rise of the Middle Kingdom")) {return "Empire Earth\" + $Folder.Name}
    if ($Folder.Name.Contains("ENDLESS Space") -or $Folder.Name.Contains("Endless Space")) {return "Endless Space\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Falcon")) {return "Falcon\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Fallout")) {return "Fallout\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Far Cry")) {return "Far Cry\" + $Folder.Name}
    if ($Folder.Name.Contains("Fleet Defender")) {return "Fleet Defender\" + $Folder.Name}
    if ($Folder.Name.Contains("Great Naval Battles")) {return "Great Naval Battles\" + $Folder.Name}
    if ($Folder.Name.Contains("Ground Control")) {return "Ground Control\" + $Folder.Name}
    if ($Folder.Name.StartsWith("HC")) {return "Heroes Chronicles\" + $Folder.Name}
    if ($Folder.Name.Contains("Heroes of the Lance") -or $Folder.Name.Contains("Dragons of Flame")) {return "Heroes of the Lance\" + $Folder.Name}
    if ($Folder.Name.Contains("Heroes of Might and Magic") -or $Folder.Name.Contains("HoMM")) {return "Heroes of Might and Magic\" + $Folder.Name}
    if ($Folder.Name.Contains("Heretic") -or $Folder.Name.Contains("HeXen")) {return "Heretic\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Hitman")) {return "Hitman\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Homeworld")) {return "Homeworld\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Independence War")) {return "Independence War\" + $Folder.Name}

    if ($Folder.Name -eq "Infernal Machine") {return "Indiana Jones\Indiana Jones and the Infernal Machine"}
    if ($Folder.Name.StartsWith("Indiana Jones")) {return "Indiana Jones\" + $Folder.Name}

    if ($Folder.Name -eq "Jagged Alliance - DG") {return "Jagged Alliance\Jagged Alliance\Jagged Alliance - DG"}
    if ($Folder.Name -eq "JA2 - Unfinished Business") {return "Jagged Alliance\Jagged Alliance 2\Jagged Alliance 2 - Unfinished Business"}
    if ($Folder.Name -eq "Jagged Alliance 2 Wildfire") {return "Jagged Alliance\Jagged Alliance 2\Jagged Alliance 2 - Wildfire"}
    if ($Folder.Name.StartsWith("Jagged Alliance")) {return "Jagged Alliance\" + $Folder.Name}

    if ($Folder.Name.StartsWith("Jazz Jackrabbit")) {return "Jazz Jackrabbit\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Just Cause")) {return "Just Cause\" + $Folder.Name}
    if ($Folder.Name.StartsWith("KKND")) {return "KKND\" + $Folder.Name}
    if ($Folder.Name.Contains("LEGO - The Hobbit")) {return "LEGO\LEGO The Hobbit"}
    if ($Folder.Name.StartsWith("LEGO")) {return "LEGO\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Leisure Suit Larry") -or $Folder.Name.Contains("lslwddd")) {return "Leisure Suit Larry\" + $Folder.Name}
    if ($Folder.Name.Contains("M1 Tank Platoon")) {return "M1 Tank Platoon\" + $Folder.Name}
    if ($Folder.Name.Contains("Master of Orion")) {return "Master of Orion\" + $Folder.Name}
    if ($Folder.Name.Contains("MDK")) {return "MDK\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Medal of Honor")) {return "Medal of Honor\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Metal Gear")) {return "Metal Gear\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Metro")) {return "Metro\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Might and Magic")) {return "Might and Magic\" + $Folder.Name}
    if ($Folder.Name.Contains("Monkey Island")) {return "Monkey Island\" + $Folder.Name}
    if ($Folder.Name.Contains("MW5Mercs")) {return "MechWarrior\MechWarrior 5 - Mercenaries"}
    if ($Folder.Name.Contains("Myst ") -or $Folder.Name.Contains("Riven") -or $Folder.Name.StartsWith("Uru")) {return "Myst\" + $Folder.Name}
    if ($Folder.Name.Contains("Neverwinter Nights") -or $Folder.Name.StartsWith("NWN")) {return "Neverwinter Nights\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Panzer General")) {return "Panzer General\" + $Folder.Name}

    if ($Folder.Name -eq "Pirates") {return "Pirates\" + $Folder.Name}
    if ($Folder.Name -eq "Pirates GOLD") {return "Pirates\" + $Folder.Name}
    if ($Folder.Name -eq "Privateer 2 - The Darkening") {return "Pirates\" + $Folder.Name}
    
    if ($Folder.Name.StartsWith("Planescape Torment")) {return "Planescape Torment\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Populous")) {return "Populous\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Prince of Persia")) {return "Prince of Persia\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Pharaoh")) {return "Pharaoh\" + $Folder.Name}
    if ($Folder.Name.StartsWith("QT")) {return "QuickTime"}

    if ($Folder.Name -eq "Arena") {return "Quake\Quake 3\Quake 3 Arena"}
    if ($Folder.Name -eq "Quake III") {return "Quake\" + $Folder.Name.Replace("III","3")}
    if ($Folder.Name.StartsWith("Quake III")) {return "Quake\Quake 3\" + $Folder.Name.Replace("III","3")}
    if ($Folder.Name -eq "Quake II") {return "Quake\" + $Folder.Name.Replace("II","2")}
    if ($Folder.Name.StartsWith("Quake II")) {return "Quake\Quake 2\" + $Folder.Name.Replace("II","2")}
    if ($Folder.Name.StartsWith("Quake")) {return "Quake\" + $Folder.Name}

    if ($Folder.Name.StartsWith("Railroad Tycoon")) {return "Railroad Tycoon\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Realms of Arkania")) {return "Realms of Arkania\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Red Baron")) {return "Red Baron\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Red Faction")) {return "Red Faction\" + $Folder.Name}
    if ($Folder.Name.StartsWith("RollerCoaster Tycoon") -or $Folder.Name.Contains("Corkscrew Follies")) {return "RollerCoaster Tycoon\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Secret Files")) {return "Secret Files\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Serious Sam")) {return "Serious Sam\" + $Folder.Name}
    if ($Folder.Name.Contains("Settlers")) {return "Settlers\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Sherlock Holmes")) {return "Sherlock Holmes\" + $Folder.Name}

    if ($Folder.Name -eq "Sid Meier's Alpha Centauri Planetary Pack") {return "Civilization\" + $Folder.Name}
    if ($Folder.Name -eq "Sid Meier's Pirates") {return "Pirates\Sid Meier's Pirates!"}
    if ($Folder.Name -eq "Sid Meier's Railroads") {return "Sid Meier's Railroads!"}

    if ($Folder.Name.StartsWith("Silent Hunter")) {return "Silent Hunter\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Silent Service")) {return "Silent Service\" + $Folder.Name}
    if ($Folder.Name.StartsWith("SimCity") -or $Folder.Name.StartsWith("SIM City")) {return "SIM City\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Sniper Ghost Warrior")) {return "Sniper Ghost Warrior\" + $Folder.Name}
    if ($Folder.Name.Contains("Soldier of Fortune")) {return "Soldier of Fortune\" + $Folder.Name}

    if ($Folder.Name -eq "Star Trek - Judgment Rites") {return "Star Trek - Judgement Rites"}
    if ($Folder.Name -eq "Star Trek Armada") {return "Star Trek - Armada\Star Trek - Armada"}
    if ($Folder.Name -eq "Star Trek Armada II") {return "Star Trek - Armada\Star Trek - Armada II"}
    if ($Folder.Name -eq "Star Trek 25th Anniversary") {return "Star Trek - 25th Anniversary"}
    if ($Folder.Name -eq "Star Trek Away Team") {return "Star Trek - Away Team"}
    if ($Folder.Name -eq "Star Trek Bridge Commander") {return "Star Trek - Bridge Commander"}
    if ($Folder.Name -eq "Star Trek Elite Force") {return "Star Trek - Voyager Elite Force"}
    if ($Folder.Name -eq "Star Trek Elite Force II") {return "Star Trek - Elite Force II"}
    if ($Folder.Name -eq "Star Trek Encyclopedia") {return "Star Trek - Encyclopedia"}
    if ($Folder.Name -eq "Star Trek Hidden Evil") {return "Star Trek - Hidden Evil"}
    if ($Folder.Name -eq "Star Trek SFC3") {return "Star Trek - Starfleet Command\Star Trek - Starfleet Command III"}
    if ($Folder.Name -eq "Star Trek TNG Technical Manual") {return "Star Trek - TNG Technical Manual"}

    if ($Folder.Name -eq "Star Wars - Dark Forces" -or $Folder.Name -eq "STAR WARS™ Dark Forces Remaster") {return "Star Wars - Dark Forces\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Star Wars - Battlefront")) {return "Star Wars - Battlefront\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Star Wars - Empire At War")) {return "Star Wars - Empire At War\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Star Wars - KotOR")) {return "Star Wars - Knights of the Old Republic\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Star Wars - Rebel Assault")) {return "Star Wars - Rebel Assault\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Star Wars - TIE Fighter")) {return "Star Wars - TIE Fighter\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Star Wars - X-Wing")) {return "Star Wars - X-Wing\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Star Wars - X-Wing vs TIE Fighter") -or $Folder.Name -eq "Star Wars - XvT") {return "Star Wars - X-Wing vs TIE Fighter\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Star Wars Jedi Knight")) {return "Star Wars - Jedi Knight\" + $Folder.Name}
    if ($Folder.Name.Contains("STAR WARS The Force Unleashed")) {return "Star Wars - Force Unleashed, The\" + $Folder.Name}
    if ($Folder.Name -eq "STAR WARS Bounty Hunter") {return "Star Wars - Bounty Hunter"}
    if ($Folder.Name -eq "STAR WARS Episode I Jedi Power Battles") {return "Star Wars - Episode I - Jedi Power Battles"}
    if ($Folder.Name -eq "STAR WARS Racer") {return "Star Wars - Racer"}
    if ($Folder.Name -eq "Star Wars Trilogy CD-ROM") {return "Star Wars - Trilogy CD-ROM"}

    if ($Folder.Name.StartsWith("Starcraft")) {return "Starcraft\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Starship Troopers")) {return "Starship Troopers\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Steel Panthers")) {return "Steel Panthers\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Styx")) {return "Styx\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Sudden Strike")) {return "Sudden Strike\" + $Folder.Name}
    if ($Folder.Name.Contains("Syberia")) {return "Syberia\" + $Folder.Name}
    if ($Folder.Name.Contains("System Shock")) {return "System Shock\" + $Folder.Name}
    if ($Folder.Name.Contains("The Dig")) {return "Dig, The"}
    if ($Folder.Name.Contains("The Escapists")) {return "Escapists\" + $Folder.Name}
    if ($Folder.Name.Contains("The Even More Incredible Machine")) {return "Even More Incredible Machine, The"}
    if ($Folder.Name.Contains("The Lion King")) {return "Lion King, The"}
    if ($Folder.Name.Contains("The Saboteur")) {return "Saboteur, The"}

    if ($Folder.Name -eq "The Witcher 2") {return "Witcher, The\Witcher 2, The"}
    if ($Folder.Name -eq "The Witcher 3 Wild Hunt GOTY") {return "Witcher, The\Witcher 3 Wild Hunt GOTY, The"}
    if ($Folder.Name -eq "The Witcher Adventure Game") {return "Witcher, The\Witcher Adventure Game, The"}
    if ($Folder.Name -eq "The Witcher Enhanced Edition") {return "Witcher, The\Witcher Enhanced Edition, The"}

    if ($Folder.Name.StartsWith("Thief")) {return "Thief\" + $Folder.Name}
    if ($Folder.Name.Contains("Tomb Raider")) {return "Tomb Raider\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Tropico")) {return "Tropico\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Turok")) {return "Turok\" + $Folder.Name}

    if ($Folder.Name.StartsWith("Ultima ") -or $Folder.Name.Contains("Akalabeth") -or $Folder.Name.Contains("Martian Dreams") -or $Folder.Name.Contains("The Savage Empire")) {return "Ultima\" + $Folder.Name}

    if ($Folder.Name.StartsWith("Unreal")) {return "Unreal\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Victory At Sea")) {return "Victory At Sea\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Victory at Sea")) {return "Victory At Sea\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Warcraft")) {return "Warcraft\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Wargame Construction Set")) {return "Wargame Construction Set\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Wasteland 2 Director's Cut")) {return "Wasteland\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Wing Commander")) {return "Wing Commander\" + $Folder.Name}

    if ($Folder.Name.StartsWith("Wizardry 1")) {return "Wizardry\Wizardry 1 - Proving Grounds of the Mad Overlord"}
    if ($Folder.Name.StartsWith("Wizardry 2")) {return "Wizardry\Wizardry 2 - The Knight of Diamonds"}
    if ($Folder.Name.StartsWith("Wizardry 3")) {return "Wizardry\Wizardry 3 - The Legacy of Llylgamyn"}
    if ($Folder.Name.StartsWith("Wizardry 4")) {return "Wizardry\Wizardry 4 - The Return of Werdna"}
    if ($Folder.Name.StartsWith("Wizardry 5")) {return "Wizardry\Wizardry 5 - The Heart of the Maelstrom"}
    if ($Folder.Name.StartsWith("Wizardry 6")) {return "Wizardry\Wizardry 6 - Bane of the Cosmic Forge"}
    if ($Folder.Name.StartsWith("Wizardry 7")) {return "Wizardry\Wizardry 7 - Crusaders of the Dark Savant"}
    if ($Folder.Name.StartsWith("Wizardry 7 Gold")) {return "Wizardry\Wizardry 7a - Wizardry Gold"}
    if ($Folder.Name.StartsWith("Wizardry 8")) {return "Wizardry\Wizardry 8"}
    if ($Folder.Name.StartsWith("Wizardry Labyrinth of Lost Souls")) {return "Wizardry\" + $Folder.Name}
    if ($Folder.Name.StartsWith("Wizardry")) {return "Wizardry\" + $Folder.Name}

    if ($Folder.Name.Contains("Wolfenstein") -or $Folder.Name.Contains("Spear of Destiny")) {return "Wolfenstein\" + $Folder.Name}
    if ($Folder.Name.Contains("World in Conflict")) {return "World in Conflict\" + $Folder.Name}
    if ($Folder.Name.Contains("Worms")) {return "Worms\" + $Folder.Name}
    if ($Folder.Name.Contains("X-COM")) {return "X-COM\" + $Folder.Name}
    if ($Folder.Name -eq "Zork Zero") {return "Zork\Zork 0 - The Revenge of Megaboz"}
    if ($Folder.Name -eq "Zork") {return "Zork\Zork 1 - The Great Underground Empire"}
    if ($Folder.Name -eq "Zork 2") {return "Zork\Zork 2 - The Wizard of Frobozz"}
    if ($Folder.Name -eq "Zork 3") {return "Zork\Zork 3 - The Dungeon Master"}
    if ($Folder.Name.Contains("Zork") -or $Folder.Name -eq "Planetfall") {return "Zork\" + $Folder.Name}

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
.{
    #Param([string]$Root,[string]$BackupFolder,[string]$LogPath)
 
    if ($Root.Equals("")) {$Root = "D:\Games\GOG Galaxy\Games"}   #if (!($PSBoundParameters.ContainsKey('Root'))) {$Root = "D:\Games\GOG Galaxy\Games"}
    if (![System.IO.Directory]::Exists($Root)) { #if ((Test-Path -Path $Root)) {
        $Message = "Error: Specified -Root folder (""$Root"") not found!"
        Write-Host $Message -ForegroundColor Red
        exit
    }

    if ($BackupFolder.Equals("")) {$BackupFolder = "\\Alpha\Public\Software\Games\PC"}   #if (!($PSBoundParameters.ContainsKey('BackupFolder'))) {$BackupFolder = "\\Alpha\Public\Software\Games\PC\"}
    #if ($BackupFolder.Equals("")) {$BackupFolder = "\\Bravo\Public\Software\Games\PC"}   #if (!($PSBoundParameters.ContainsKey('BackupFolder'))) {$BackupFolder = "\\Bravo\Public\Software\Games\PC\"}
    #if ((Test-Path -Path $BackupFolder)) {
    if (![System.IO.Directory]::Exists($BackupFolder)) {
        $Message ="Error: Specified -BackupFolder folder (""$BackupFolder"") not found!"
        Write-Host $Message -ForegroundColor Red
        exit
    }
    if (!$BackupFolder.EndsWith("\")) {$BackupFolder = $BackupFolder + "\"}

    #if (!($PSBoundParameters.ContainsKey('LogPath'))) {$LogPath = "$($BackupFolder)HarvestGOGGoodies.log"}
    if ($LogPath.Equals("")) {$LogPath = "$($BackupFolder)HarvestGOGGoodies.log"}
    
    $AppName = "HarvestGOGGoodies"
    $StartTime = Get-Date
    $Moved = 0
    $Skipped = 0
    $Message = "[$AppName © $("{0:yyyy}" -f $StartTime), Ken Clark                       $("{0:MM/dd/yy} {0:hh:mm:ss tt}" -f $StartTime)]"
    Write-Message $Message -Path $LogPath
    $EmailBody = $Message.Replace("©", "&copy;") + "<br />"
    $Message = "Harvesting GOG Goodies from $($(Get-Item $Root).FullName) to $BackupFolder"
    Write-Message $Message -Path $LogPath
    $EmailBody += $Message + "<br /><br />"
    $test = $false

    $folders = Get-ChildItem $Root | Where {$_.PSIsContainer -eq $true} | Sort-Object
    foreach ($iFolder in $folders)
    {
        #$subFolderItems = Get-ChildItem -LiteralPath $iFolder.FullName -recurse -force | Where-Object {$_.PSIsContainer -eq $false} | Measure-Object -property Length -sum | Select-Object Sum
        #$iFolder.Name + "|" + "{0:N2}" -f ($subFolderItems.sum / 1MB) + " MB"
        #$iFolder.Name
        #$i.Name + "\" + $Downloads.Name
    
        $Downloads = Get-ChildItem -LiteralPath $iFolder.FullName -Recurse -Depth 2 -force | Where {$_.Name -Like "!Downloads"}
        $sourceFolder = $Downloads.FullName + "\"
        $targetFolder = $BackupFolder + "$(Get-TargetFolder -Folder $iFolder)"

        Write-Message "`t$($iFolder.FullName)" -Path $LogPath
        if (($Downloads -eq $null) -or ($iFolder.FullName -eq "\")) {
            $Message = "Warning: No Goodies found for $($iFolder.Name)!"
            Write-Message "`t`t$Message" -Path $LogPath
            $EmailBody += "$Message<br />"
        } else {
            if (!(Test-Path -path $targetFolder)) {
                #If the target doesn't exist, there's no point in checking details, just copy the whole thing...
                $Message = "Creating $($targetFolder)"
                Write-Message "`t`t$Message" -Path $LogPath
                $EmailBody += "$Message<br />"
                if(-not $test) {$NewFolder = New-Item $targetFolder -Type Directory}
                
                $Message = "Moving $($sourceFolder) to $($targetFolder)"
                Write-Message "`t`t$Message" -Path $LogPath
                $EmailBody += "$Message<br />"

                #TODO: Removed -LiteralPath as it interferred with the desired wild-carding, may have introduced an issue by doing so...
                if(-not $test) {Move-Item "$($sourceFolder)*" -Destination $targetFolder -Force    }
                $Moved += 1
            } else {
                #Now sweat the details and only copy anything not already there...
                $Items = Get-ChildItem -LiteralPath $Downloads.FullName -Recurse -force
                foreach ($iItem in $Items) {
                    $targetFile = $iItem.FullName.Replace($iFolder.FullName, $targetFolder).Replace("\!Downloads",$null)
                    #Write-Message "`t`tConsidering $($iItem.FullName) against $($targetFile)..." -Path $LogPath

                    #$doCopy = $false
                    $doCopy = $true
                    if (!(Test-Path -path $targetFile)) {
                        Write-Message "`t`t$targetFile does not exist!... (doCopy)" -Path $LogPath
                        #$doCopy = $true
                    } else {
                        $targetItem = Get-Item -LiteralPath $targetFile
                        if ($iItem.Length -ne $targetItem.Length) {
                            $Message = "$($iItem.Name) Length $($iItem.Length) does not match $($targetItem.Name) Length $($targetItem.Length)!... (doCopy)"
                            Write-Message "`t`t$Message" -Path $LogPath
                            $EmailBody += "$Message<br />"
                            $doCopy = $true
                        }
                    }

                    if ($doCopy) {
                        if ($iItem.Name -eq "\") {
                            Write-Message "`t`tERROR: iItem.Name is $($iItem.Name)" -Path $LogPath
                            $Skipped += 1
                        } else {
                            #Write-Message "`t`tMoving $($iItem.FullName) to $($targetFile)..." -Path $LogPath
                            $Message = "Moving $($iItem.Name) to $($targetFolder)..."
                            Write-Message "`t`t$Message" -Path $LogPath
                            $EmailBody += "$Message<br />"
                            if(-not $test) {
                                #Move-Item -LiteralPath $iItem.FullName -Destination $targetFile -Force
                                Copy-Item -LiteralPath $iItem.FullName -Destination $targetFile -Force
                                Remove-Item -LiteralPath $iItem.FullName -Force
                            }
                            $Moved += 1
                        }
                    } else {
                        $Skipped += 1
                    }
                }
            }
        }
    }
    $EndTime = Get-Date
    $Message = "`n$AppName Complete @ $("{0:hh:mm:ss tt}" -f $EndTime) (Elapsed: $(Format-Elapsed -Start $StartTime -End $EndTime))"
    Write-Message $Message -Path $LogPath
    $EmailBody += "<br /><dl>$Message"
    $Message = "`t$("{0:#,##0}" -f $Moved) Files Harvested ($("{0:#,##0}" -f $Skipped) Skipped)"
    Write-Message $Message -Path $LogPath
    $EmailBody += "<dd>$Message</dd></dl>"
    Write-Separator -Path $LogPath

    &"$PSScriptRoot\eMailResults.ps1" -Subject "$Root.$AppName Complete" -Body "$EmailBody" -LogFile $LogPath -AsHTML
    #Send-Mail -AppName $AppName -EmailBody $EmailBody -LogPath $LogPath
}