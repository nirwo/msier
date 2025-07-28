
Add-Type -AssemblyName PresentationFramework

[xml]$xaml = Get-Content -Path ".\InstallMSIGUI.xaml"
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

$csvPath      = $window.FindName("CsvPath")
$msiPath      = $window.FindName("MsiPath")
$username     = $window.FindName("Username")
$password     = $window.FindName("Password")
$outputBox    = $window.FindName("OutputBox")
$browseCsv    = $window.FindName("BrowseCsv")
$browseMsi    = $window.FindName("BrowseMsi")
$startInstall = $window.FindName("StartInstall")
$exportLog    = $window.FindName("ExportLog")
$progressBar  = $window.FindName("InstallProgress")

$logBuffer = New-Object System.Collections.ObjectModel.ObservableCollection[string]

function Log($msg) {
    $ts = (Get-Date).ToString("HH:mm:ss")
    $fullMsg = "[$ts] $msg"
    $window.Dispatcher.Invoke([action]{
        $outputBox.AppendText("$fullMsg`n")
        $outputBox.ScrollToEnd()
    })
    $logBuffer.Add($fullMsg)
}

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

$exportLog.Add_Click({
    $path = "$env:USERPROFILE\\Desktop\\msi_install_log_$(Get-Date -Format yyyyMMdd_HHmmss).txt"
    $logBuffer | Out-File -Encoding UTF8 -FilePath $path
    Log "Log saved to: $path"
})

$startInstall.Add_Click({
    $outputBox.Clear()
    $logBuffer.Clear()
    $progressBar.Value = 0

    if (-not (Test-Path $csvPath.Text) -or -not (Test-Path $msiPath.Text)) {
        Log "❌ Please provide valid CSV and MSI paths."
        return
    }

    try {
        $cred = New-Object System.Management.Automation.PSCredential (
            $username.Text,
            (ConvertTo-SecureString $password.Password -AsPlainText -Force)
        )
    } catch {
        Log "❌ Invalid credentials: $_"
        return
    }

    $targets = Import-Csv $csvPath.Text
    $total = $targets.Count
    $completed = 0
    $jobs = @()

    foreach ($target in $targets) {
        $job = Start-Job -ScriptBlock {
            param($computer, $msiPath, $cred)
            function SafeLog($m) {
                "$computer : $m"
            }

            if (-not (Test-Connection -ComputerName $computer -Count 1 -Quiet)) {
                return SafeLog "❌ Host unreachable"
            }

            try {
                $session = New-PSSession -ComputerName $computer -Credential $cred -ErrorAction Stop
                Copy-Item -ToSession $session -Path $msiPath -Destination "C:\Windows\Temp\installer.msi"
                Invoke-Command -Session $session -ScriptBlock {
                    Start-Process "msiexec.exe" -ArgumentList "/i C:\Windows\Temp\installer.msi /qn /norestart" -Wait
                }
                Remove-PSSession $session
                return SafeLog "✅ Installation completed"
            } catch {
                return SafeLog "❌ $_"
            }
        } -ArgumentList $target.ComputerName, $msiPath.Text, $cred

        $jobs += $job
    }

    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromSeconds(1)
    $timer.Add_Tick({
        foreach ($job in $jobs.ToArray()) {
            if ($job.State -eq 'Completed') {
                $result = Receive-Job -Job $job
                Log $result
                Remove-Job $job
                $jobs = $jobs | Where-Object { $_.Id -ne $job.Id }
                $completed++
                $progressBar.Value = ($completed / $total) * 100
            }
        }

        if ($jobs.Count -eq 0) {
            $timer.Stop()
            Log "✅ All installations completed."
        }
    })
    $timer.Start()
})

$window.ShowDialog() | Out-Null
