Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# --- Configuration des ports ---
$TcpPorts = @(3074) + (27014..27050)
$UdpPorts = @(3074,3478) + (4379..4380) + (27000..27030)

# --- Création de la fenêtre ---
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Black Ops II - UPnP Port Manager" Height="500" Width="600"
        ResizeMode="CanResize" WindowStartupLocation="CenterScreen">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        
        <!-- Titre -->
        <TextBlock Grid.Row="0" Text="Gestionnaire de ports UPnP pour Black Ops II" 
                   FontSize="16" FontWeight="Bold" HorizontalAlignment="Center" 
                   Margin="0,0,0,20"/>
        
        <!-- Boutons -->
        <StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,0,0,20">
            <Button Name="btnAdd" Content="Ouvrir les ports" Width="150" Height="40" 
                    Margin="0,0,10,0" Background="#4CAF50" Foreground="White" 
                    FontSize="14" FontWeight="Bold"/>
            <Button Name="btnRemove" Content="Fermer les ports" Width="150" Height="40" 
                    Background="#f44336" Foreground="White" 
                    FontSize="14" FontWeight="Bold"/>
        </StackPanel>
        
        <!-- Zone de logs -->
        <Border Grid.Row="2" BorderBrush="Gray" BorderThickness="1" CornerRadius="5">
            <ScrollViewer Name="scrollViewer" VerticalScrollBarVisibility="Auto">
                <TextBlock Name="txtLogs" TextWrapping="Wrap" Margin="10" 
                           FontFamily="Consolas" FontSize="12" 
                           Background="Black" Foreground="Lime" Padding="10"/>
            </ScrollViewer>
        </Border>
    </Grid>
</Window>
"@

# --- Chargement de l'interface ---
$window = [Windows.Markup.XamlReader]::Parse($xaml)
$btnAdd = $window.FindName("btnAdd")
$btnRemove = $window.FindName("btnRemove")
$txtLogs = $window.FindName("txtLogs")
$scrollViewer = $window.FindName("scrollViewer")

# --- Fonction pour ajouter des logs ---
function Add-Log {
    param([string]$Message, [string]$Color = "Lime")
    
    $timestamp = Get-Date -Format "HH:mm:ss"
    $logEntry = "[$timestamp] $Message`n"
    
    $window.Dispatcher.Invoke([action]{
        $txtLogs.Text += $logEntry
        $scrollViewer.ScrollToBottom()
    })
}

# --- Fonction pour obtenir l'IP locale ---
function Get-LocalIP {
    try {
        $LocalIP = (Get-NetIPAddress -AddressFamily IPv4 `
                    | Where-Object { $_.PrefixOrigin -ne "WellKnown" -and $_.IPAddress -notlike "169.*" } `
                    | Select-Object -First 1 -ExpandProperty IPAddress)
        
        if ($LocalIP) {
            Add-Log "IP locale détectée : $LocalIP"
            return $LocalIP
        } else {
            Add-Log "❌ Impossible de déterminer l'IP locale" "Red"
            return $null
        }
    } catch {
        Add-Log "❌ Erreur lors de la détection de l'IP : $($_.Exception.Message)" "Red"
        return $null
    }
}

# --- Fonction pour gérer les ports UPnP ---
function Set-UPnPPorts {
    param([string]$Action)
    
    Add-Log "=== Début de l'action : $Action ==="
    
    # Désactiver les boutons pendant le traitement
    $btnAdd.IsEnabled = $false
    $btnRemove.IsEnabled = $false
    
    try {
        # Obtenir l'IP locale
        $LocalIP = Get-LocalIP
        if (-not $LocalIP) {
            return
        }
        
        # Initialiser UPnP
        Add-Log "Initialisation de UPnP..."
        $UPnP = New-Object -ComObject HNetCfg.NATUPnP
        $Maps = $UPnP.StaticPortMappingCollection
        
        if (-not $Maps) {
            Add-Log "❌ Routeur ne répond pas via UPnP (IGD)" "Red"
            return
        }
        
        Add-Log "✓ Connexion UPnP établie avec le routeur"
        
        # Fonction interne pour traiter un port (identique au CLI original)
        function Process-Port {
            param([int]$Port, [string]$Proto)
            
            try {
                $exists = $null
                try {
                    $exists = $Maps.Item($Port, $Proto)
                } catch {
                    # Item() lève une exception si le port n'existe pas
                    $exists = $null
                }
                
                switch ($Action) {
                    "Add" {
                        if ($exists) {
                            Add-Log "◦ $Port/$Proto existe déjà – ignoré" "Yellow"
                        } else {
                            try {
                                $Maps.Add($Port, $Proto, $Port, $LocalIP, $true, "BO2 $Proto $Port")
                                Add-Log "✓ Ajouté : $Port/$Proto → $LocalIP"
                            } catch {
                                Add-Log "❌ Erreur ajout $Port/$Proto : $($_.Exception.Message)" "Red"
                            }
                        }
                    }
                    "Remove" {
                        if (-not $exists) {
                            Add-Log "◦ $Port/$Proto n'existe pas – ignoré" "Yellow"
                        } else {
                            try {
                                $Maps.Remove($Port, $Proto)
                                Add-Log "✗ Supprimé : $Port/$Proto"
                            } catch {
                                Add-Log "❌ Erreur suppression $Port/$Proto : $($_.Exception.Message)" "Red"
                            }
                        }
                    }
                }
            } catch {
                Add-Log "❌ Erreur critique sur $Port/$Proto : $($_.Exception.Message)" "Red"
            }
        }
        
        # Traitement des ports TCP
        Add-Log "Traitement des ports TCP..."
        $TcpPorts | ForEach-Object { Process-Port $_ "TCP" }
        
        # Traitement des ports UDP
        Add-Log "Traitement des ports UDP..."
        $UdpPorts | ForEach-Object { Process-Port $_ "UDP" }
        
        Add-Log "=== Action '$Action' terminée avec succès ==="
        
    } catch {
        Add-Log "❌ Erreur critique : $($_.Exception.Message)" "Red"
    } finally {
        # Réactiver les boutons
        $btnAdd.IsEnabled = $true
        $btnRemove.IsEnabled = $true
    }
}

# --- Événements des boutons ---
$btnAdd.Add_Click({
    Set-UPnPPorts -Action "Add"
})

$btnRemove.Add_Click({
    Set-UPnPPorts -Action "Remove"
})

# --- Initialisation ---
Add-Log "Interface UPnP Black Ops II démarrée"
Add-Log "Ports TCP : $($TcpPorts -join ', ')"
Add-Log "Ports UDP : $($UdpPorts -join ', ')"
Add-Log "Prêt à configurer les ports UPnP..."

# --- Affichage de la fenêtre ---
$window.ShowDialog() | Out-Null
