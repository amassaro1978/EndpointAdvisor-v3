Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.HasPrivateKey } | ForEach-Object {
    [PSCustomObject]@{
        Subject  = $_.Subject.Substring(0, [Math]::Min(60, $_.Subject.Length))
        Issuer   = $_.Issuer.Substring(0, [Math]::Min(60, $_.Issuer.Length))
        EKU      = ($_.EnhancedKeyUsageList.FriendlyName -join ', ')
        KeyUsage = ($_.Extensions | Where-Object { $_ -is [System.Security.Cryptography.X509Certificates.X509KeyUsageExtension] } | ForEach-Object { $_.KeyUsages })
        Expires  = $_.NotAfter.ToString('yyyy-MM-dd')
        Thumb    = $_.Thumbprint.Substring(0,8)
    }
} | Format-Table -AutoSize
