
Add-Type -AssemblyName PresentationFramework

[xml]$xaml = Get-Content -Path ".\InstallMSIGUI.xaml"
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Bind UI elements
$csvPath = $window.FindName("CsvPath")
$msiPath = $window.FindName("MsiPath")
$username = $window.FindName("Username")
$password = $window.FindName("Password")
$outputBox = $window.FindName("OutputBox")
$browseCsv = $window.FindName("BrowseCsv")
$browseMsi = $window.FindName("BrowseMsi")
$startInstall = $window.FindName("StartInstall")

# Browse file dialogs
$browseCsv.Add_Click({
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Filter = "CSV files (*.csv)|*.csv"
    if ($dlg.ShowDialog()) { $csvPath.Text = $dlg.FileName }
})

$browseMsi.Add_Click({
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Filter = "MSI files (*.msi)|*.msi"
    if ($dlg.ShowDialog()) { $msiPath.Text = $dlg.FileName }
})

# Job results queue
$syncHash = [hashtable]::Synchronized(@{})
$syncHash.Results = @()

# Install logic
$startInstall.Add_Click({
    $outputBox.Clear()
    $cred = New-Object System.Management.Automation.PSCredential (
        $username.Text,
        (ConvertTo-SecureString $password.Password -AsPlainText -Force)
    )

    if (-not (Test-Path $csvPath.Text) -or -not (Test-Path $msiPath.Text)) {
        $outputBox.AppendText("❌ Please provide valid CSV and MSI paths.`n")
        return
    }

    $targets = Import-Csv $csvPath.Text
    $jobs = @()

    foreach ($target in $targets) {
        $job = Start-Job -ScriptBlock {
            param($computer, $msiPath, $cred)
            try {
                $session = New-PSSession -ComputerName $computer -Credential $cred -ErrorAction Stop
                Copy-Item -ToSession $session -Path $msiPath -Destination "C:\Windows\Temp\installer.msi"
                Invoke-Command -Session $session -ScriptBlock {
                    Start-Process "msiexec.exe" -ArgumentList "/i C:\Windows\Temp\installer.msi /qn /norestart" -Wait
                }
                Remove-PSSession $session
                "$computer : ✅ Installed successfully"
            } catch {
                "$computer : ❌ $_"
            }
        } -ArgumentList $target.ComputerName, $msiPath.Text, $cred

        $jobs += $job
    }

    # Background timer to monitor job completion
    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromSeconds(1)
    $timer.Add_Tick({
        foreach ($job in $jobs) {
            if ($job.State -eq 'Completed') {
                $result = Receive-Job -Job $job
                $outputBox.AppendText("$result`n")
                Remove-Job $job
                $jobs = $jobs | Where-Object { $_.Id -ne $job.Id }
            }
        }

        if ($jobs.Count -eq 0) {
            $timer.Stop()
            $outputBox.AppendText("✅ All installations completed.`n")
        }
    })
    $timer.Start()
})

# Show window
$window.ShowDialog() | Out-Null
