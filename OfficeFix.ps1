if ([bool](([System.Security.Principal.WindowsIdentity]::GetCurrent()).groups -match "S-1-5-32-544")) {
    if ($(New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Common\Identity").GetValue("EnableADAL") -eq $null) {
        New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Common\Identity" -name EnableADAL -value 1  
    }
    if ($(New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Common\Identity").GetValue("Version") -eq $null) {
        New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\15.0\Common\Identity" -name Version -Value 1
    }
}