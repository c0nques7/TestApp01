using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics.Metrics;
using System.Management.Automation;
using System.Threading;
using System.Threading.Tasks;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                ExecutePowerShellCommands();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        public static void ExecutePowerShellCommands()
        {
            try
            {
                Console.WriteLine("PowerShell User Lookup Tool");
                Console.Write("Enter the range of usernames beginning with 'TL': ");
                string usernameRange = Console.ReadLine();

                int maxDegreeOfParallelism = 3;

                using (PowerShell ps = PowerShell.Create())
                {
                    ps.AddScript(@"
                            $usernameExtension = Read-Host "Enter the 6 - digit extension for the username"
    $username = "TL$usernameExtension"

    $filter = "SamAccountName -eq `'$username`'"
    $user = Get - ADUser - Filter $filter - Properties LastLogon

    if ($user - and $user.LastLogon) {
        $lastLogonTime = [DateTime]::FromFileTime($user.LastLogon)
        Write - Host "User: $($user.Name) | Last Logon: $($lastLogonTime.ToString('g'))"
    }
                    elseif($user) {
                        Write - Host "User: $($user.Name) | Last Logon: Not available"
    } else
                    {
                        Write - Host "User with username '$username' not found."
    }
                }
catch
            {
                Write - Host "An error occurred: $($_.Exception.Message)"
}
            ");
                    ps.Invoke();
                }

                Console.WriteLine("PowerShell script execution completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}