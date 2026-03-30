# Lists all certs that could be smart card related
# Shows the provider (how to tell physical vs virtual vs YubiKey)

Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.HasPrivateKey } | ForEach-Object {
    $provider = ""
    try {
        $key = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($_)
        if ($key) { $provider = $key.Key.Provider.Provider }
    } catch {}
    if (-not $provider) {
        try { $provider = $_.PrivateKey.CspKeyContainerInfo.ProviderName } catch {}
    }

    $isSmartCardLogon = $_.EnhancedKeyUsageList.ObjectId -contains "1.3.6.1.4.1.311.20.2.2"
    $hasClientAuth = $_.EnhancedKeyUsageList.ObjectId -contains "1.3.6.1.5.5.7.3.2"

    # Determine type based on provider
    $type = "Unknown"
    if ($provider -match "Smart Card" -and $provider -match "Virtual") { $type = "VIRTUAL SMART CARD" }
    elseif ($provider -match "Smart Card|Base CSP") { $type = "PHYSICAL SMART CARD" }
    elseif ($provider -match "YubiKey|Yubi") { $type = "YUBIKEY" }
    elseif ($isSmartCardLogon) { $type = "SMART CARD (by OID)" }
    elseif ($hasClientAuth) { $type = "CLIENT AUTH" }

    if ($isSmartCardLogon -or $hasClientAuth -or $provider -match "Smart|Yubi|Virtual|PIV") {
        Write-Host "`n--- $type ---" -ForegroundColor Cyan
        Write-Host "  Expires:  $($_.NotAfter.ToString('yyyy-MM-dd'))"
        Write-Host "  Provider: $provider"
        Write-Host "  EKU:      $($_.EnhancedKeyUsageList.FriendlyName -join ', ')"
        Write-Host "  Thumb:    $($_.Thumbprint.Substring(0,8))..."
    }
}
Write-Host ""
