using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using System.Collections.ObjectModel;


namespace NAVinfo
{
    class functions
    {

        public bool ExecuteCommand(string exeDir, string args)
        {
            try
            {
                ProcessStartInfo procStartInfo = new ProcessStartInfo();

                procStartInfo.FileName = exeDir;
                procStartInfo.Arguments = args;
                procStartInfo.RedirectStandardOutput = true;
                procStartInfo.UseShellExecute = false;
                procStartInfo.CreateNoWindow = true;

                using (Process process = new Process())
                {
                    process.StartInfo = procStartInfo;
                    process.Start();

                    process.WaitForExit();

                    string result = process.StandardOutput.ReadToEnd();
                    Console.WriteLine(result);
                }
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine("*** Error occured executing the following commands.");
                Console.WriteLine(exeDir);
                Console.WriteLine(args);
                Console.WriteLine(ex.Message);
                return false;
            }


        }


        public bool KeyExists(RegistryKey baseKey, string subKeyName)
        {
            RegistryKey ret = baseKey.OpenSubKey(subKeyName);

            return ret != null;
        }


        public string resetWiFi()
        {
            using (PowerShell PS = PowerShell.Create() )
            {
                PS.AddScript("Disable-NetAdapter -Name 'Wi-Fi'-Confirm:$false; ping localhost -n 1; Enable-NetAdapter -Name 'Wi-Fi'-Confirm:$false");

                Collection<PSObject> PSoutput = PS.Invoke();
                string output = "";
                foreach (PSObject outputitem in PSoutput)
                {
                    if (outputitem != null)
                    {
                        output += outputitem.ToString();

                    }
                }
                return output;

            }
        }

        public string resetF5VPN()
        {
            using (PowerShell PS = PowerShell.Create())
            {
                PS.AddScript(@"$F5ID = (& 'c:\Program Files(x86)\F5 VPN\f5fpc.exe' -info | findstr /r ^[1-9] | ForEach-Object {$_.Substring(0,7)})
                                & 'c:\Program Files (x86)\F5 VPN\f5fpc.exe' - stop / s $F5ID");

                Collection<PSObject> PSoutput = PS.Invoke();
                string output = "";
                foreach (PSObject outputitem in PSoutput)
                {
                    if (outputitem != null)
                    {
                        output += outputitem.ToString();

                    }
                }
                return output;

            }

        }


        public string resetOutlook()
        {
            using (PowerShell PS = PowerShell.Create())
            {
                PS.AddScript(@"
                                function ResetOutlook{


                                . c:\windows\staging\Functions.ps1;

				                                FinnEpostadresse;
                                                $ostpath = '$env:USERPROFILE\appdata\local\microsoft\outlook\$env:USERNAME.ost';
                                                $result = Invoke - Command - ScriptBlock { c:\windows\staging\profiler.exe Outlook $awepost $ostpath};
                                 ResetOutlook;
                                ");

                Collection<PSObject> PSoutput = PS.Invoke();
                string output = "";
                foreach (PSObject outputitem in PSoutput)
                {
                    if (outputitem != null)
                    {
                        output += outputitem.ToString();

                    }
                }
                return output;

            }
        }

    }
}
