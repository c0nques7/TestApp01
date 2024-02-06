using System;
using System.Collections.Generic;
using System.Management.Automation;
using System.Threading;
using System.Threading.Tasks;

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
                            $domainName = 'yourdomain.com'
                            $usernameRange = '" + usernameRange + @"'
                            $users = Get-ADUser -Filter {SamAccountName -like 'TL' + $usernameRange + '*'} -SearchBase ""DC=$domainName"" -Properties LastLogon

                            foreach ($user in $users) {
                                $lastLogonTime = [DateTime]::FromFileTime($user.LastLogon)
                                Write-Host ""User: $($user.Name) | Last Logon: $lastLogonTime""
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