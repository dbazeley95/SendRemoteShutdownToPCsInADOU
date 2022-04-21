###### Script to remotely shutdown Windows PCs

# Requirements: 
# -Script has to be run by an administrator of the system being shutdown
# -Must be able to ping the machine before a shutdown is attempted

# Get date for Output File
$ShutdownDate = Get-Date -Format "dddd dd/mm/yyyy HH:mm"

# Import all AD computers from the OU and nested OUs
$IncludedDevices = Get-ADComputer -Filter * -SearchBase "OU=Windows Workstations,DC=Example,DC=net" | select -expandproperty Name

# Specify a list of AD computers that you wish to not power off but are in the OU's specified
$ExcludedDevices = @(
                    'TEST1',
                    'TEST2'
    )

# Compares and prepares the array for a list of computers in the OU and and not in the exclusion list
$PCNamesToShutdown = $IncludedDevices | Where {$ExcludedDevices -NotContains $_}

# Loop to check whether each PC in the array is online and shutdown if true, else write out PC is offline
ForEach ($PC in $PCNamesToShutdown){
           
       If (Test-Connection -BufferSize 32 -Count 1 -ComputerName $PC -Quiet) {
        
        shutdown.exe /m \\$PC /s /t 300
        Write-Host "$ShutdownDate - A shutdown has been attempted for $PC!"
        
        } Else {
        
        Write-Host "$ShutdownDate - $PC was already switched off!"
        
        }
  
}