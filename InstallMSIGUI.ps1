
Add-Type -AssemblyName PresentationFramework

[xml]$xaml = Get-Content -Path ".\InstallMSIGUI.xaml"
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Bind controls
$csvPath = $window.FindName("CsvPath")
$msiPath = $window.FindName("MsiPath")
$username = $window.FindName("Username")
$password = $window.FindName("Password")
$outputBox = $window.FindName("OutputBox")
$browseCsv = $window.FindName("BrowseCsv")
$browseMsi = $window.FindName("BrowseMsi")
$startInstall = $window.FindName("StartInstall")

# Browse CSV
$browseCsv.Add_Click({
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Filter = "CSV files (*.csv)|*.csv"
    if ($dlg.ShowDialog()) { $csvPath.Text = $dlg.FileName }
})

# Browse MSI
$browseMsi.Add_Click({
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Filter = "MSI files (*.msi)|*.msi"
    if ($dlg.ShowDialog()) { $msiPath.Text = $dlg.FileName }
})

# Start Installation
$startInstall.Add_Click({
    $outputBox.Text = ""
    $cred = New-Object System.Management.Automation.PSCredential (
        $username.Text,
        (ConvertTo-SecureString $password.Password -AsPlainText -Force)
    )

    $targets = Import-Csv $csvPath.Text
    foreach ($target in $targets) {
        Start-Job -ScriptBlock {
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
        } -ArgumentList $target.ComputerName, $msiPath.Text, $cred | Out-Null
    }

    Register-ObjectEvent -InputObject ([System.Management.Automation.Job]::GetJobs()) -EventName StateChanged -Action {
        $job = $Event.SourceArgs[0]
        if ($job.State -eq 'Completed') {
            $result = Receive-Job -Job $job
            $window.Dispatcher.Invoke([action] { $outputBox.AppendText("$result`n") })
            Remove-Job $job
        }
    }
})

$window.ShowDialog() | Out-Null
