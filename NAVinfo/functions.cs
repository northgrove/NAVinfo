using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using System.Collections.ObjectModel;
using System.Net;


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

        public string getEmail(string exeDir, string args)
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

                    return result;
                }
                
            }
            catch (Exception ex)
            {
                Console.WriteLine("*** Error occured executing the following commands.");
                Console.WriteLine(exeDir);
                Console.WriteLine(args);
                Console.WriteLine(ex.Message);
                return ex.Message;
            }


        }

        public bool KeyExists(RegistryKey baseKey, string subKeyName)
        {
            RegistryKey ret = baseKey.OpenSubKey(subKeyName);

            return ret != null;
        }


        public string resetWiFi()
        {

            return "wifi";
            
            /**
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
            **/

        }

        public string resetF5VPN()
        {
            using (PowerShell PS = PowerShell.Create())
            {
                PS.AddScript(@"$F5ID = (& 'c:\Program Files(x86)\F5 VPN\f5fpc.exe' -info | findstr /r ^[1-9] | ForEach-Object {$_.Substring(0,7)})
                                & 'c:\Program Files (x86)\F5 VPN\f5fpc.exe' -stop / s $F5ID");

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


        public string resetOutlook(bool FullClean)
        {
            var output = "";
            var username = Environment.GetEnvironmentVariable("USERNAME");
            RegistryKey NAVKey = Registry.CurrentUser.OpenSubKey(@"Software\NAV", true);
            //var email = NAVKey.GetValue("EmailLoggedonUser").ToString();
            //var email2 = email.Remove(email.Length - 2);
            var email3 = getEmail("whoami", "/upn");
            var email4 = System.Text.RegularExpressions.Regex.Replace(email3, "[^A-Za-z0-9.@]", "");
            var userprofile = Environment.GetEnvironmentVariable("USERPROFILE");
            var ostpath = userprofile + "\\appdata\\local\\microsoft\\outlook\\" + username + ".ost";
            var ostdir = userprofile + "\\appdata\\local\\microsoft\\outlook\\*";

            try
            {
               ExecuteCommand("pskill", "/accepteula outlook");
                output += "Stengte Outlook # ";
            }
            catch
            {
                output += "klarte ikke å strenge outlook #";
            }
            
            try
            {
                ExecuteCommand("Powershell", @"remove-item HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook\Profiles\Outlook -Recurse -Force");
                output += "fjernet outlook profil fra registry #";
            }
            catch (Exception ex)
            {
                output += "klarte ikke å slette outlook profilen i registry # " + ex;
            }

            /* Fjernet for å benytte default profil wizard i outlook
            try
            {
                ExecuteCommand("c:\\windows\\Nsystem\\profiler.exe", "Outlook " + email4 + " " + ostpath);
                output += "Opprettet ny outlook profil #";
            }
            catch
            {
                output += "Klarte ikke å opprette ny profil #";
            }
            */
            try
            {
                ExecuteCommand("Powershell", @"new-item HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook\Profiles\Outlook");
                output += "La til ny outlook profil i registry #";
            }
            catch (Exception ex)
            {
                output += "klarte ikke å legge til ny outlook profil i registry # " + ex;
            }
            if (FullClean)
            {
                try
                {
                    ExecuteCommand("Powershell", @"remove-item -path " + ostdir + " -recurse -force");
                    output += "Slettet Outlook cache i brukerprofilen # ";
                }
                catch
                {
                    output += "Klarte ikke å slette Outlook cache i brukerprofilen # ";
                }
            }
            if (!FullClean)
            {
                output += "FullClean er ikke valgt # ";
            }

            try
            {
                //ExecuteCommand("c:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE", "");
                //ExecuteCommand("outlook", "");
                System.Threading.Thread.Sleep(1000);
                System.Diagnostics.Process.Start("outlook");
                output += "Startet outlook på nytt #";
            }
            catch
            {
                output += "klarte ikke å starte outlook på nytt #";
            }



            return output.ToString();

            
        }


        public static bool CheckForInternetConnection()
        {
            try
            {
                using (var client = new WebClient())
                {
                    using (client.OpenRead("http://clients3.google.com/generate_204"))
                    {
                        return true;
                    }
                }
            }
            catch
            {
                return false;
            }
        }

    }
}
