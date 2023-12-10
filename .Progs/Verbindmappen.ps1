<#
.SYNOPSIS
 Dit programmaatje verzorgt een makkelijke manier om gedeelde netwerkmappen
 te verbinden als netwerk schijven.
.DESCRIPTION
 Het programma maakt netwerk schijven aan van gedeelde netwerkmappen aan de 
 hand van specificaties vastgelegd in het bestand VerbindMappen.[mappennaam].csv.  
 Dit bestand moet in dezelfde map staan als #Verbind.cmd waar je het 
 programma mee start.
 Het formaat van het csv bestand is als volgt:

SchijfLetter,NetwerkMap
A:,nassie.onstein52.local\homes\rob
B:,nassie.onstein52.local\netboot
H:,nassie.onstein52.local\hardware
P:,nassie.onstein52.local\printer
S:,nassie.onstein52.local\software
T:,nassie.onstein52.local\torrent-inbox
U:,nassie.onstein52.local\download

 De eerste regel bevat de tekst exact zoals hierboven weergegeven, gevolgd 
 door een of meer regels die een vrije schijfletter bevatten, gevolgd door
 een komma en meteen daarna de netwerknaam van de gedeelde netwerk map.
 De mapping wordt gemaakt met de opgegeven driveletter, naar de gespecificeerde
 gedeelde map.
 
 Het programma accepteert een aantal parameter waarden die tijdens de oproep
 kunnen worden meegeven: 
  -VerversInlogInfo [Waar of NietWaar, dus als parameter aanwezig is, wordt opnieuw
                     om inlog details gevraagd]
  -MappenBestand [Mappennaam gebruikt in het Mappenbestand] 
	
 Bij aanroep kunnen er twee dingen gebeuren:
 1. Als je nog nooit wachtwoord bestanden hebt gemaakt, zal het programma 
    vragen om het accounts en het wachtwoorden welke gebruikt moet worden om de 
	netwerkmappen te kunnen verbinden
 2. Als er reeds  wachtwoord bestanden zijn aangemaakt, wordt meteen 
    geprobeerd de netwerkmappen te verbinden

 PROGRAMMADETAILS
 Behalve dat het programma het VerbindMappen.[mappennaam].csv bestand leest, maakt het 
 ook bestanden aan met daarin de accountnaam en een beveiligde/versleutelde versie
 van het wachtwoord voor elke host die in het VerbindMappen.[mappennaam].csv staat.
 De naam van dit bestand wordt als volgt samengesteld:
   .\Cache\AccountInfo.[hostname].[mappennaam]
 Als de optie -MappenBestand met een MappenBestandNaam wordt 
 meegegeven aan het programma, is de naam van het beveiligde/verstleutelde 
 wachtwoord bestand .\Cache\AccountInfo.[hostnaam].[MappenNaam].

 NOOT: dit versleutelde bestand kan alleen gebruikt worden op de computer
       waar het bestand op is aangemaakt en door de (aangelogde) gebruiker 
       die het bestand heeft aangemaakt. Andere gebruikers kunnen deze
	   gegevens niet gebruiken om het wachtwoord te achterhalen of om 
	   de mappen te verbinden.
 EasyConnect gebruikt de mappennaam 'default' als geen 
 -MappenBestand optie is gespecificeerd en leest dus het Verbindingen.Default.csv
 om de verbindingen te maken. Het leest de Verbindingen.[MappenBestandNaam].csv
 als deze optie wel is meegegeven tijdens de aanroep van het programma en zal
 alle verbindingsopdrachten die hier in staan verwerken.
 Een verbindings opdracht is simpelweg een regel die de volgende informatie
 bevat:
 <schijfletter>,<Gedeelde mapnaam op de NAS>
 Waarbij schijfletter een (vrije) schijfletter op de computer moet zijn,
 bijvoorbeeld E: of H: of M: en
 <gedeelde mapnaam op de NAS> de volledige naam moet zijn van de gedeelde
 map, bijvoorbeeld nassie\homes\kees of nassie\audio of, als je liever 
 werkt met volledige gekwalificeerde namen nassie.onstein52.local\homes\kees
 #>

param ( 
    [Switch]$VerversInlogInfo = $False,
    [string]$MappenBestand = "Default",
    [switch]$Help = $False
)

# Functie die wacht op het indrukken van de "any" toest :-)
Function pause ($message) {
    # Check if running Powershell ISE
    if (($psISE) -or ($psEditor)) {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show("$message")
    }
    else {
        Write-Host "$message" -ForegroundColor Yellow
        $x = $host.ui.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
}

# Algemene Error en Stop codeblock
$ErrorStop = {
    Pause "Uitvoering van het programma gestopt `n `nDruk op een toets om verder te gaan..."
}

# Functie die n wachtwoor bestand leest en/of creeert.
# Init tabel welke bijhoudt of een ID/WW al is ingetoetst in deze run
$Script:BestandReedsVerverst = @()
Function Get-AccountPasswordForHost {
    Param (
        $AccountBestandnaam,
        $Hostname,
        [Switch]$Vervang = $False
    )
    if ((-not (Test-Path $AccountBestandnaam)) -or 
        ($Vervang)) {
        if ($Script:BestandReedsVerverst -notcontains $AccountBestandnaam) {

            # Er is geen AccountBestand gevonden, OF er is een expliciete ververs actie voor de bestanden aangevraagd
            $Accountnaam = ""
            try {
                $Accountnaam = Read-Host "Wat is de inlognaam die je wilt gebruiken voor het maken van de verbindingen voor host $($Hostname)? " -ErrorAction SilentlyContinue
            }
            Catch { }
            If ($Accountnaam -eq "") {
                "Geen accountnaam ingevoerd" | Out-Host
                & $ErrorStop
                Exit
            }
		
            $Wachtwoord = ""
            Try {
                $Wachtwoord = Read-Host "Wat is het wachtwoord wat hoort bij deze inlognaam?" -AsSecureString  -ErrorAction SilentlyContinue
            }
            Catch { }
            If ($Wachtwoord -eq "") {
                "Geen wachtwoord ingevoerd" | Out-Host
                & $ErrorStop
                Exit
            }
	
            $VeiligWachtwoord = $Wachtwoord | ConvertFrom-SecureString
	
            "$Accountnaam,$VeiligWachtwoord" | Out-File $AccountBestandnaam
            "Accountnaam en wachtwoord zijn veilig opgeslagen."
            $Script:BestandReedsVerverst += $AccountBestandnaam
        }
    }
    $VeiligWachtwoord = $null
    $Accountnaam, $temp = (Get-Content $AccountBestandnaam).Split(",")
    Try {
        [System.Security.SecureString]$VeiligWachtwoord = $temp | ConvertTo-SecureString -ErrorAction SilentlyContinue
    }
    Catch { }
    if ($Null -eq $VeiligWachtwoord) {
        "De opgeslagen account gegevens voor host $Hostname zijn ongeldig. Probeer ze opnieuw aan te maken`ndoor optie -VerversInlogInfo te gebruiken" | Out-Host
        & $ErrorStop
        Exit
    }
    $LoginInfo = "" | Select-Object "Account", "VeiligWachtwoord"
    $LoginInfo.Account = $Accountnaam
    $LoginInfo.VeiligWachtwoord = $VeiligWachtwoord
    $LoginInfo
}


######################
# Script begint hier #
######################
clear-host
$ErrorActionPreference = 'Stop'

If ($Help) {
    Get-Help ".\.Progs\Verbindmappen.ps1"
    & $ErrorStop
    Return
}

# Script omgevings waarden
$ScriptPath = (split-path -parent $MyInvocation.MyCommand.Definition)
$ScriptPathroot = $ScriptPath.replace("\.Progs", "")
if (!(test-path "$ScriptPathroot\.Cache")) {
    mkdir -Path "$ScriptPathroot\.Cache" | Out-Null
}
$AccountBestandnaamTemplate = "{0}\.Cache\AccountInfo.##HOSTNAME##.$MappenBestand" -f $ScriptPathroot
$VerbindMappenBestandNaam = "{0}\VerbindMappen.$MappenBestand.csv" -f $ScriptPathroot

# Vraag af wat er ingevoerd is
if ("$MappenBestand" -eq "") {
    "Parameter -Mappenbestand heeft geen waarde toegewezen gekregen." | Out-Host
    & $ErrorStop
    Return
}
# Kijk of het Verbndmappenbestand al bestaat
if (!(Test-Path $VerbindMappenBestandNaam)) {
    "Kan het bestand $VerbindMappenBestandNaam niet vinden.`nMaak dit aan en vul deze met verbind opdrachten" | Out-Host
    # Aanmaken voorbeeld bestand:
    "SchijfLetter,NetwerkMap" | Out-File -FilePath $VerbindMappenBestandNaam -Encoding ASCII
    "H:,nas.mijnnetwerk.local\homes\kees" | Out-File -FilePath $VerbindMappenBestandNaam -Append -Encoding ASCII
    "A:,nas.mijnnetwerk.local\audio" | Out-File -FilePath $VerbindMappenBestandNaam -Append -Encoding ASCII
    "V:,nas.mijnnetwerk.local\video" | Out-File -FilePath $VerbindMappenBestandNaam -Append -Encoding ASCII
    "Een voorbeeld van dit bestand is aangemaakt, pas dit aan naar jouw wensen`n" | Out-Host
    # Maak ook een cmd bestand aan
    if ($MappenBestand -ne "Default") {
        $CmdFile = "{0}\#Verbind.$MappenBestand.cmd" -f $env:_eDir
        if (!(Test-Path $CmdFile)) {
            "@Echo off" | Out-File -FilePath $CmdFile -Encoding ASCII
            "Call #Verbind.cmd -MappenBestand $MappenBestand %1 %2 %3 %4 %5 %6 %7 %8 %9" | Out-File -FilePath $CmdFile -Encoding ASCII -Append
		
            "Om het aanmaken van deze verbindingen te vereenvoudigen heb ik bestand`n$CMDFile`nvoor je gemaakt" | Out-Host
            "Als je dit bestand 'dubbeklikt', wordt het verbinden van deze mappen gestart" | Out-Host
        }
    }
    & $ErrorStop
    Return
}

# Als we hier komen is er een VerbindMappen verzoek gedaan, en is het csv bestand gevonden
$VerbindOpdrachten = @(Import-Csv $VerbindMappenBestandNaam -ErrorAction SilentlyContinue)
If ($VerbindOpdrachten.Count -le 0) {
    "Het bestand $VerbindMappenBestandNaam heeft`neen ongeldige inhoud (geen elementen). Controleer de inhoud van dit bestand" | Out-Host
    & $ErrorStop
    Return
}
If (($null -eq $VerbindOpdrachten[0].SchijfLetter) -or
    ($null -eq $VerbindOpdrachten[0].NetwerkMap)) {
    "Het bestand $VerbindMappenBestandNaam heeft`neen ongeldige inhoud. Controleer de inhoud van dit bestand" | Out-Host
    & $ErrorStop
    Return
}

"Aanmaken van netwerkschijven mappenbestand $MappenBestand wordt gestart" | Out-Host

# Start met maken van de nieuwe verbindingen
$TryCount = 0

Do {
    $TryCount++
    $ErrorsFound = 0
    "Het maken van verbindingen, poging $TryCount" | Out-Host

    # Bereid een Com-object voor
    $map = new-object -ComObject WScript.Network
    # Maak alle bestaande verbindingen ongedaan
    #Try { $ComThings = $map.EnumNetworkDrives() }Catch { } 
    $ComThings = Get-PSDrive | Where {"$($_.DisplayRoot)".StartsWith("\\")}
    foreach ($ComThing in $ComThings) {
#        if ($ComThing.Length -ne 0) {
#            if ($ComThing.SubString(0, 2) -ne "\\") {
#                "Bezig om bestaande verbinding met $ComThing te verbreken..." | Out-Host
                try {
#                    $map.RemoveNetworkDrive($ComThing, "True")
                    net use "$($ComThing):" /d 
                    Start-Sleep -Milliseconds 1000
#                    "Klaar met verwijderen $($ComThing.Root)" | Out-Host
                }
                Catch {
                    "Het verbreken van de verbinding met $($ComThing.Root) is mislukt.`nHerstel deze fout en probeer opnieuw" | Out-Host
                    "{0}" -f $_
                    "Suggestie: probeer het commando " | Out-Host
                    "    net use * /d /yes   (of 'net use * /d /ja' bij nederlands talige windows)" | Out-Host
                    "in een opdrachtprompt" | Out-Host
                    & $ErrorStop
                    Return
                }
 #           }
 #       }
    }
    ForEach ($Opdracht in $VerbindOpdrachten) {
        $HostName = $Opdracht.NetwerkMap.Split("\")[0]
        $AccountBestandnaam = $AccountBestandnaamTemplate.Replace("##HOSTNAME##", $HostName)
        $LoginObject = Get-AccountPasswordForHost -AccountBestandnaam $AccountBestandnaam -Hostname $Hostname -Vervang:$VerversInlogInfo
        # Vertaal veilig wachtwoord naar leesbaar wachtwoord
        $credential = New-Object System.Management.Automation.PsCredential($LoginObject.Account, $LoginObject.VeiligWachtwoord)
        $helderWachtwoord = ($credential.GetNetworkCredential()).Password
        $Netwerkmap = "\\" + $Opdracht.NetwerkMap
        Try {
            $map.MapNetworkDrive($Opdracht.SchijfLetter, $Netwerkmap, $false, $LoginObject.Account, $helderWachtwoord) | Out-Null
            "Het maken van schijfletter {0} is voor map {1} is gelukt." -f $Opdracht.SchijfLetter, $Netwerkmap | Out-Host
        }
        Catch {
            $ErrorsFound++
            "Het maken van schijfletter {0} is voor map {1} is mislukt, foutmelding:" -f $Opdracht.SchijfLetter, $Netwerkmap | Out-Host
            "{0}" -f $Error[0] | Out-Host
        }
    }
    Start-sleep -Seconds 2
} While (($TryCount -lt 3) -and ($ErrorsFound -ne 0))

If ($ErrorsFound -ne 0) {
    "`nHet ging niet helemaal goed!`n-----------------------------`n" | Out-Host
    "Bij meldingen over een ongeldig wachtwoord of account, moet het account/`nwachtwoord bestand opnieuw aangemaakt worden." | Out-Host
    "Dit doe je door een opdrachtprompt te openen en naar deze folder te gaan,`ndoor commando" | Out-Host
    "`n   cd /d {0}`n" -f $env:_Edir | Out-Host
    "uit te voeren, en daarna het commando" | out-host
    "(of dubbelklik op het #Opdrachtprompt.cmd bestand in deze folder`n`n" | Out-Host
    if ($MappenBestand -ne "Default") {
        $ExtraText = " -MappenBestand $MappenBestand"
    }
    "`n   #Verbind$ExtraText -VerversInlogInfo`n`n" | Out-Host
    "Bij meldingen over een ongeldige of onvindbare netwerk schijven moet je bestand" | Out-Host
    "$VerbindMappenBestandNaam" | Out-Host
    "nog eens goed nakijken of alle schijfletters vrij zijn en de netwerkmappen goed`nzijn opgegeven.`n" | Out-Host

    & $ErrorStop
    Return
}
Else {
    "`nDit venster sluit vanzelf over 3 seconden..."
    Start-sleep -Seconds 3
    Return
}