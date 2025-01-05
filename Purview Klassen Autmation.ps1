#Installiere nötige Packete
# Check if ExchangeOnlineManagement is installed
if (-not (Get-Module -Name ExchangeOnlineManagement -ListAvailable)) {
    Write-Output "ExchangeOnlineManagement module is not installed. Installing now..."
    
    # Install the module
    try {
        Install-Module -Name ExchangeOnlineManagement -Force -Scope CurrentUser
        Write-Output "ExchangeOnlineManagement module installed successfully."
    } catch {
        Write-Error "Failed to install ExchangeOnlineManagement module: $_"
    }
} else {
    Write-Output "ExchangeOnlineManagement module is already installed."
}

# Check if AIPService is installed
if (-not (Get-Module -Name AIPService -ListAvailable)) {
    Write-Output "AIPService module is not installed. Installing now..."
    
    # Install the module
    try {
        Install-Module -Name AIPService -Force -Scope CurrentUser
        Write-Output "AIPService module installed successfully."
    } catch {
        Write-Error "Failed to install AIPService module: $_"
    }
} else {
    Write-Output "AIPService module is already installed."
}



# Check the AIP Service status
Import-Module AIPService
Connect-AipService

$serviceStatus = Get-AIPService

if ($serviceStatus.Status -eq "Disabled") {
    Write-Output "AIP Service is currently disabled. Enabling now..."
    
    # Enable the AIP Service
    Enable-AIPService
    
    # Confirm the service is enabled
    $newStatus = Get-AIPService
    if ($newStatus.Status -eq "Enabled") {
        Write-Output "AIP Service has been successfully enabled."
    } else {
        Write-Error "Failed to enable the AIP Service. Please check the service configuration."
	sleep 5
	exit
    }
} else {
    Write-Output "AIP Service is already enabled."
}

# Disconnect from AIP Service
Disconnect-AipService





# Verbindung zu Microsoft Purview herstellen
Connect-IPPSSession

$FirmenName = "Firma Contoso"


    # Labels erstellen
$labels = @(
	@{Name="Öffentlich"; Tooltip="Für öffentlich zugängliche Informationen"; Description="Dieses Label ist für Daten vorgesehen, die öffentlich geteilt werden können."; ContentMarkingText=""; EncryptionEnabled=$false},
        @{Name="Intern"; Tooltip="Nur für interne Nutzung"; Description="Dieses Label ist für Daten vorgesehen, die innerhalb des Unternehmens bleiben sollen."; ContentMarkingText=""; EncryptionEnabled=$false},
        @{Name="Vertraulich"; Tooltip="Vertrauliche Informationen"; Description="Dieses Label schützt vertrauliche Daten."; ContentMarkingText="Vertraulich - $FirmenName"; EncryptionEnabled=$true;EncryptionOfflineAccessDays=7 },
        @{Name="Streng Vertraulich"; Tooltip="Hochsensible Informationen"; Description="Dieses Label schützt hochsensible Daten und wendet erweiterte Verschlüsselung an."; ContentMarkingText="Streng Vertraulich - $FirmenName"; EncryptionEnabled=$true;EncryptionOfflineAccessDays=7}
)

    # Labels durchlaufen und erstellen
foreach ($label in $labels) {
        $labelParams = @{
        Name = $label.Name
        DisplayName = $label.Name
        Tooltip = $label.Tooltip
        Comment = $label.Description
        EncryptionEnabled = $label.EncryptionEnabled
        EncryptionPromptUser = $true
        ContentMarkingText = $label.ContentMarkingText
        ApplyContentMarkingFooterAlignment = "Center"
        ApplyContentMarkingFooterEnabled = $true
        ApplyContentMarkingFooterFontSize = 10
        ApplyContentMarkingFooterText = $label.ContentMarkingText
        ContentType = "File, Email"
}

if ($label.EncryptionEnabled) {
        $labelParams.EncryptionOfflineAccessDays = $label.EncryptionOfflineAccessDays
}

New-Label @labelParams
}
# Verbindung beenden
Disconnect-ExchangeOnline -Confirm:$false

# Verbindung zu Azure Information Protection herstellen
Connect-AipService

# Veröffentlichungsrichtlinie für die Gruppe GF erstellen
$policyNameGF = "Richtlinie Geschäftsführung"
$policyDescriptionGF = "Diese Richtlinie stellt sicher, dass die Geschäftsführung alle Labels nutzen kann."

New-AipServicePolicy `
           -Name $policyNameGF `
           -Description $policyDescriptionGF `
           -Labels "Öffentlich, Intern, Vertraulich, Streng Vertraulich" `
           -DefaultLabelId "Öffentlich" `
           -ScopeId "GF" `
           -Mandatory $true

# Veröffentlichungsrichtlinie für die Gruppe Mitarbeiter erstellen
$policyNameMitarbeiter = "Richtlinie Mitarbeiter"
$policyDescriptionMitarbeiter = "Diese Richtlinie stellt sicher, dass Mitarbeiter nur eingeschränkte Labels nutzen können."

New-AipServicePolicy `
       -Name $policyNameMitarbeiter `
       -Description $policyDescriptionMitarbeiter `
       -Labels "Öffentlich, Intern" `
       -DefaultLabelId "Öffentlich" `
       -ScopeId "Mitarbeiter" `
       -Mandatory $true
