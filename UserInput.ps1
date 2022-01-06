
# try block to execute
try
{
    # variable to hold user selection
    $userInput = 0

    # while loop | 5 exits
    while ($userInput -ne 5)
    {
        # show user menu to select from
        Write-Host -ForegroundColor DarkGreen -BackgroundColor White
            "
            1. List Log Files Appended To DailyLog.txt
            2. List Files In Directory
            3. List Current CPU and Memory Usage
            4. List Running Processes
            5. Exit The Script Execution
            "
        # get user input
        Write-Host "Please Choose An Option (1-5):"
        $userInput = Read-Host

        # switch statement based on user input
        switch($userInput)
        {
            1 # user selected 1
            {
                "Date: " + (Get-Date) | Out-File -FilePath $PSScriptRoot\DailyLog.txt -Append 
                Get-ChildItem -Path $PSScriptRoot -Filter *.log | Out-File -FilePath $PSScriptRoot\DailyLog.txt -Append
            }

            2 # user selected 2
            {
                Get-ChildItem "$PSScriptRoot" | Sort-Object Name | Format-Table -AutoSize -Wrap | Out-File -FilePath "$PSScriptRoot\directoryContents.txt"
            }

            3 # user selected 3
            {
                $counterParameters = “\Processor(_Total)\% Processor Time”, “\Memory\Committed Bytes”
                Get-Counter -Counter $counterParameters -MaxSamples 4 -SampleInterval 5 
            }

            4 # user selected 4
            {
                Get-Process | Select-Object ID, Name, VM | Sort-Object VM | Out-GridView
            }
            
            5 # user selected 5
            {
                # exits while loop
            }

        }
    }
}

# catch system out of memory error
Catch [System.OutOfMemoryException]
{
    Write-Host -ForegroundColor Red "OutOfMemoryException Occured"
}

# catch all other errors
Catch
{
    Write-Host -ForegroundColor Red "Error: $($_.Exception.Message)"
}

