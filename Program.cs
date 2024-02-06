using System;
using System.Management.Automation;
using System.Threading.Tasks;

namespace ConsoleApp
{
    class Program
    {
        static async Task Main(string[] args)
        {
            try
            {
                await ExecutePowerShellCommandsAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        public static async Task ExecutePowerShellCommandsAsync()
        {
            try
            {
                Console.WriteLine("PowerShell User Lookup Tool");
                Console.Write("Enter the 6-digit extension for the username: ");
                string usernameExtension = Console.ReadLine();

                string username = $"TL{usernameExtension}";

                using (PowerShell ps = PowerShell.Create())
                {
                    ps.AddScript($@"# Ensure the ActiveDirectory module is imported
Import-Module ActiveDirectory

# Define the suffixes and their adjustments
$suffixes = @{{ 
    ""null"" = '';  # No suffix for 'null'
    ""E"" = 'e';    # 'E' suffix turns to 'e' at the end of the numeric string
    ""G"" = 'TGL';  # Directly use 'TGL' without 'TL' prefix
    ""W"" = 'TVL'   # Directly use 'TVL' without 'TL' prefix
}}

function Get-UserRange {{
    param(
        [int]$Min = 0,
        [int]$Max = 999999
    )
    $range = $null
    do {{
        $start = Read-Host ""Enter the start of the range (between $Min and $Max)""
        $end = Read-Host ""Enter the end of the range (between $Min and $Max)""
        if ($start -match '^\d+$' -and $end -match '^\d+$') {{
            $start = [int]$start
            $end = [int]$end
            if ($start -le $end -and $start -ge $Min -and $end -le $Max) {{
                $range = @($start, $end)
            }} else {{
                Write-Host ""Start must be less than or equal to End, and both must be between $Min and $Max.""
            }}
        }} else {{
            Write-Host ""Please enter valid numbers for both start and end of the range.""
        }}
    }} while (-not $range)
    return $range
}}

$range = Get-UserRange
$rangeStart = $range[0]
$rangeEnd = $range[1]

# Prompt the user for the output CSV file name
$fileName = Read-Host ""Enter the name for the output CSV file (without extension)""
$fileName = if (-not $fileName.EndsWith('.csv')) {{ ""$fileName.csv"" }} else {{ $fileName }}
$desktopPath = [System.Environment]::GetFolderPath(""Desktop"")
$fullPath = Join-Path -Path $desktopPath -ChildPath $fileName

# Confirm save location
$confirmation = Read-Host ""File will be saved to your desktop as '$fileName'. Type 'yes' to confirm.""
if ($confirmation -eq 'yes') {{
    $outputData = @()

    foreach ($suffixKey in $suffixes.Keys) {{
        $suffixValue = $suffixes[$suffixKey]
        for ($i = $rangeStart; $i -le $rangeEnd; $i++) {{
            $numericPart = ""{{0:D6}}"" -f $i
            $username = switch ($suffixKey) {{
                ""null"" {{ ""TL$numericPart"" }}
                ""E""    {{ ""TL$numericPart$suffixValue"" }}
                default {{ ""$suffixValue$numericPart"" }}
            }}

            try {{
                $user = Get-ADUser -Filter ""SamAccountName -eq '$username'"" -Properties Name, LastLogon
                if ($user -ne $null) {{
                    $lastLogonTime = [DateTime]::FromFileTime($user.LastLogon)
                    $outputData += New-Object PSObject -Property @{{
                        Username = $username
                        User = $user.Name
                        LastLogon = $lastLogonTime.ToString('g')
                    }}
                }} else {{
                    $outputData += New-Object PSObject -Property @{{
                        Username = $username
                        User = 'Not Found'
                        LastLogon = 'N/A'
                    }}
                }}
            }} catch {{
                Write-Host ""An error occurred while attempting to retrieve user $username. Error: $_""
            }}
        }}
    }}

    $outputData | Export-Csv -Path $fullPath -NoTypeInformation
    Write-Host ""Data exported to $fullPath""
}} else {{
    Write-Host ""Operation cancelled by user.""
}}
");
                    var results = await ps.InvokeAsync();
                    foreach (var result in results)
                    {
                        Console.WriteLine(result);
                    }
                }
            }
            catch
            {
                Console.WriteLine("An error occurred while executing PowerShell commands.");
            }
        }
    }
}
