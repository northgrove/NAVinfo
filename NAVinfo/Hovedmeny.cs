using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Management;
using System.Net;
using System.Net.NetworkInformation;
using Microsoft.Win32;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Deployment.Application;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace NAVinfo
{
    public partial class Hovedmeny : Form
    {
        RegistryKey HKCU = Registry.CurrentUser;

       

        public Hovedmeny()
        {
            InitializeComponent();
            Console.SelectedIndexChanged += new EventHandler(tabControl1_SelectedIndexChanged);

            
            // sjekker om NAVinfo registry key finnes, lager om ikke
            var navinfofs = new functions();
            if (navinfofs.KeyExists(HKCU, @"Software\NAVinfo"))
            {
                
            }
            else
            {
                HKCU.CreateSubKey(@"Software\NAVinfo");
           
            };


            // Setter versjonsinfo

            if (System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed)
            {
                label21.Text = System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString();
            }

            label24.Text = Environment.OSVersion.ToString();


        }




       


        private void Hovedmeny_Load(object sender, EventArgs e)
        {
            
            
            // Finner serienummer på maskin
            ManagementObjectSearcher ComSerial = new ManagementObjectSearcher("SELECT * FROM Win32_BaseBoard");

            foreach (ManagementObject wmi in ComSerial.Get())
            {
                try
                {
                    label12.Text = wmi.GetPropertyValue("SerialNumber").ToString();
                }
                catch { }
            }

            // Setter maskinnavn
            label11.Text = Environment.MachineName;



            // Finner IP-adressene på maskinen
            IPAddress[] localIPs = Dns.GetHostAddresses(Dns.GetHostName());
            foreach (IPAddress addr in localIPs)
            {
                //label14.Text = label14.Text + addr + Environment.NewLine;
                textBox2.Text += addr + Environment.NewLine;
            }


            // Sjekker om det finnes en aktiv nettverkstilkobling
            bool connection = NetworkInterface.GetIsNetworkAvailable();
            
            if (connection == true)

            {
                labelwifi.BackColor = Color.Green;
                labelwifi.Text = "Tilkoblet";

            }
            else
            {
                labelwifi.BackColor = Color.Red;
                labelwifi.Text = "Ikke tilkoblet";
            }

            // sjekker internett forbindelse
            if (functions.CheckForInternetConnection())
            {
                labelinet.BackColor = Color.Green;
                labelinet.Text = "Tilkoblet";
            }
            else
            {
                labelinet.BackColor = Color.Red;
                labelinet.Text = "Ikke tilkoblet";
            }

            // Sjekker om betteriet har status "lader"
            var strom = SystemInformation.PowerStatus.PowerLineStatus.ToString();
            if (strom == "Online")
            {
                labelstrom.BackColor = Color.Green;
                labelstrom.Text = "Tilkoblet";
            }
            else
            {
                labelstrom.BackColor = Color.Red;
                labelstrom.Text = "Ikke tilkoblet";
            }

            // pålogget bruker:
            label13.Text = Environment.UserName;



        }

        // link til passordbytte
        private void button4_Click(object sender, EventArgs e)
        {
            // laster passord.nav.no i eget browser vindu
            byttpassord passordbytte = new byttpassord();
            passordbytte.Show();
        }


        // Link til Office 365 hjelp
        private void button6_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("C:\\Windows\\Mob\\kom i gang.exe");
        }

        // link til office 365 hjelp
        private void button7_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://navno.sharepoint.com/sites/Office365Hjelp");
        }




        // Trigger Functions ut i fra hvilken fane som velges
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch ((sender as TabControl).SelectedIndex)
            {
                case 6:
                    //lesEventlog();
                    break;
                case 4:
                    mapPrint();
                    break;
                case 5:
                    openStrom();
                    break;
            }


        }



/**
        // Lister installerte applikasjoner og viser i fane
        public void GetInstalledApps()
        {
            listView1.Columns.Add("Navn", 400);
            string uninstallKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
            using (RegistryKey rk = Registry.LocalMachine.OpenSubKey(uninstallKey))
            {
                foreach (string skName in rk.GetSubKeyNames())
                {
                    using (RegistryKey sk = rk.OpenSubKey(skName))
                    {
                        try
                        {
                            listView1.Items.Add(sk.GetValue("DisplayName").ToString());
                        }
                        catch (Exception ex)
                        { }
                    }
                }
            }
        }
**/

        // Leser Eventloggen og viser i fane
        /**
        public void lesEventlog()
        {
            listView2.Sorting = SortOrder.Descending;
            listView2.Columns.Add("ID", 25);
            listView2.Columns.Add("Dato", 120);
            listView2.Columns.Add("Source", 80);
            listView2.Columns.Add("Feil", 200);
            foreach (System.Diagnostics.EventLogEntry log in eventLog1.Entries)
            {
                string[] arr = new string[4];
                ListViewItem item;
                arr[0] = log.InstanceId.ToString();
                arr[1] = log.TimeGenerated.ToString();
                arr[2] = log.Source.ToString();
                arr[3] = log.Message.ToString();
                item = new ListViewItem(arr);
                listView2.Items.Add(item);
            }
            
        }
    **/

        public void lesNAVlogger()
        {

        }

        // Sjekker om NAV Skrivere er installert
        private void mapPrint()
        {
            listViewPrint.Items.Clear();
            listViewPrint.Columns.Add("Printer", 200);
            listViewPrint.Columns.Add("Status", 200);

            var print2 = Printer.IsPrinterInstalled("FargeDuplexCloud");
            if (print2)
            {
                string[] printer = { "FargeDuplex", "Tilkoblet" };
                var listViewItem = new ListViewItem(printer);
                listViewPrint.Items.Add(listViewItem);
            }
            else
            {
                string[] printer = { "FargeDuplex", "Frakoblet" };
                var listViewItem = new ListViewItem(printer);
                listViewPrint.Items.Add(listViewItem);
            }
            var print3 = Printer.IsPrinterInstalled("SortDuplexCloud");
            if (print3)
            {
                string[] printer = { "SortDuplex", "Tilkoblet" };
                var listViewItem = new ListViewItem(printer);
                listViewPrint.Items.Add(listViewItem);
            }
            else
            {
                string[] printer = { "SortDuplex", "Frakoblet" };
                var listViewItem = new ListViewItem(printer);
                listViewPrint.Items.Add(listViewItem);
            }
            var print4 = Printer.IsPrinterInstalled("FargeSimplexCloud");
            if (print4)
            {
                string[] printer = { "FargeSimplex", "Tilkoblet" };
                var listViewItem = new ListViewItem(printer);
                listViewPrint.Items.Add(listViewItem);
            }
            else
            {
                string[] printer = { "FargeSimplex", "Frakoblet" };
                var listViewItem = new ListViewItem(printer);
                listViewPrint.Items.Add(listViewItem);
            }
            var print5 = Printer.IsPrinterInstalled("SortSimplexCloud");
            if (print5)
            {
                string[] printer = { "SortSimplex", "Tilkoblet" };
                var listViewItem = new ListViewItem(printer);
                listViewPrint.Items.Add(listViewItem);
            }
            else
            {
                string[] printer = { "SortSimplex", "Frakoblet" };
                var listViewItem = new ListViewItem(printer);
                listViewPrint.Items.Add(listViewItem);
            }

        }

        //Leser strøminstillinger ved valg av strøm fanen
        public void openStrom()
        {
            RegistryKey NAVinfoKey = Registry.CurrentUser.OpenSubKey(@"Software\NAVinfo", true);

            var stromvalgAC = NAVinfoKey.GetValue("stromvalgAC");
            //var stromvalgAC = "0";
            var stromvalgDC = NAVinfoKey.GetValue("stromvalgDC");
            //var stromvalgDC = "0";
            switch (stromvalgAC)
            {
                case "0":
                    checkedListBox1.SetItemChecked(0, true);
                    break;
                case "1":
                    checkedListBox1.SetItemChecked(1, true);
                    break;
                case "2":
                    checkedListBox1.SetItemChecked(2, true);
                    break;
                case "3":
                    checkedListBox1.SetItemChecked(3, true);
                    break;
            }

            switch (stromvalgDC)
            {
                case "0":
                    checkedListBox2.SetItemChecked(0, true);
                    break;
                case "1":
                    checkedListBox2.SetItemChecked(1, true);
                    break;
                case "2":
                    checkedListBox2.SetItemChecked(2, true);
                    break;
                case "3":
                    checkedListBox2.SetItemChecked(3, true);
                    break;
            }


        }




        // Installerer alle skrivere 
        private void button8_Click(object sender, EventArgs e)
        {
            ProcessStartInfo mssettings = new ProcessStartInfo("ms-settings:printers");
            mssettings.WindowStyle = ProcessWindowStyle.Maximized;
            
            Process.Start(mssettings);

            //System.Diagnostics.Process.Start("ms-settings:printers");
            Printerveiledning installprinter = new Printerveiledning();

            int screenHeight = Screen.PrimaryScreen.WorkingArea.Height - 50;
            int screenWidth = Screen.PrimaryScreen.WorkingArea.Width - 750;

            

            installprinter.StartPosition = FormStartPosition.Manual;
           
            installprinter.Location = new Point(screenWidth);

            installprinter.Show();
            

        }



        // Strøminstillinger
        private void checkedListBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (checkedListBox2.CheckedItems.Count > 1)
            {
                Int32 checkedItemIndex = checkedListBox2.CheckedIndices[0];
                checkedListBox2.ItemCheck -= checkedListBox2_ItemCheck;
                checkedListBox2.SetItemChecked(checkedItemIndex, false);
                checkedListBox2.ItemCheck += checkedListBox2_ItemCheck;
            }
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (checkedListBox1.CheckedItems.Count > 1)
            {
                Int32 checkedItemIndex = checkedListBox1.CheckedIndices[0];
                checkedListBox1.ItemCheck -= checkedListBox1_ItemCheck;
                checkedListBox1.SetItemChecked(checkedItemIndex, false);
                checkedListBox1.ItemCheck += checkedListBox1_ItemCheck;
            }

        }
        private void checkedListBox1_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (checkedListBox1.CheckedItems.Count == 1)
            {
                Boolean isCheckedItemBeingUnchecked = (e.CurrentValue == CheckState.Checked);
                if (isCheckedItemBeingUnchecked)
                {
                    e.NewValue = CheckState.Checked;
                }
                else
                {
                    Int32 checkedItemIndex = checkedListBox1.CheckedIndices[0];
                    checkedListBox1.ItemCheck -= checkedListBox1_ItemCheck;
                    checkedListBox1.SetItemChecked(checkedItemIndex, false);
                    checkedListBox1.ItemCheck += checkedListBox1_ItemCheck;
                }

                return;
            }
        }
        private void checkedListBox2_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (checkedListBox2.CheckedItems.Count == 1)
            {
                Boolean isCheckedItemBeingUnchecked = (e.CurrentValue == CheckState.Checked);
                if (isCheckedItemBeingUnchecked)
                {
                    e.NewValue = CheckState.Checked;
                }
                else
                {
                    Int32 checkedItemIndex = checkedListBox2.CheckedIndices[0];
                    checkedListBox2.ItemCheck -= checkedListBox2_ItemCheck;
                    checkedListBox2.SetItemChecked(checkedItemIndex, false);
                    checkedListBox2.ItemCheck += checkedListBox2_ItemCheck;
                }

                return;
            }
        }


        //Setter og lagrer strøminstillinger
        private void button9_Click(object sender, EventArgs e)
        {
            var powercmd = new functions();
            RegistryKey NAVinfoKey = Registry.CurrentUser.OpenSubKey(@"Software\NAVinfo", true);
            if (checkedListBox1.SelectedIndex == 0)
            {
                
                var runpowercmd = powercmd.ExecuteCommand("c:\\windows\\system32\\powercfg.exe", "-setacvalueindex 381b4222-f694-41f0-9685-ff5bb260df2e 4f971e89-eebd-4455-a8de-9e59040e7347 5ca83367-6e45-459f-a27b-476b1d01c936 0");
                NAVinfoKey.SetValue("stromvalgAC", "0");
            }
            else if (checkedListBox1.SelectedIndex == 1)
            {
                
                var runpowercmd = powercmd.ExecuteCommand("c:\\windows\\system32\\powercfg.exe", "-setacvalueindex 381b4222-f694-41f0-9685-ff5bb260df2e 4f971e89-eebd-4455-a8de-9e59040e7347 5ca83367-6e45-459f-a27b-476b1d01c936 1");
                NAVinfoKey.SetValue("stromvalgAC", "1");
            }
            else if (checkedListBox1.SelectedIndex == 2)
            {
                
                var runpowercmd = powercmd.ExecuteCommand("c:\\windows\\system32\\powercfg.exe", "-setacvalueindex 381b4222-f694-41f0-9685-ff5bb260df2e 4f971e89-eebd-4455-a8de-9e59040e7347 5ca83367-6e45-459f-a27b-476b1d01c936 2");
                NAVinfoKey.SetValue("stromvalgAC", "2");
            }
            else if (checkedListBox1.SelectedIndex == 3)
            {
                
                var runpowercmd = powercmd.ExecuteCommand("c:\\windows\\system32\\powercfg.exe", "-setacvalueindex 381b4222-f694-41f0-9685-ff5bb260df2e 4f971e89-eebd-4455-a8de-9e59040e7347 5ca83367-6e45-459f-a27b-476b1d01c936 3");
                NAVinfoKey.SetValue("stromvalgAC", "3");
            }


            if (checkedListBox2.SelectedIndex == 0)
            {
                
                var runpowercmd = powercmd.ExecuteCommand("c:\\windows\\system32\\powercfg.exe", "-setdcvalueindex 381b4222-f694-41f0-9685-ff5bb260df2e 4f971e89-eebd-4455-a8de-9e59040e7347 5ca83367-6e45-459f-a27b-476b1d01c936 0");
                NAVinfoKey.SetValue("stromvalgDC", "0");
            }
            else if (checkedListBox2.SelectedIndex == 1)
            {
                
                var runpowercmd = powercmd.ExecuteCommand("c:\\windows\\system32\\powercfg.exe", "-setdcvalueindex 381b4222-f694-41f0-9685-ff5bb260df2e 4f971e89-eebd-4455-a8de-9e59040e7347 5ca83367-6e45-459f-a27b-476b1d01c936 1");
                NAVinfoKey.SetValue("stromvalgDC", "1");
            }
            if (checkedListBox2.SelectedIndex == 2)
            {
               
                var runpowercmd = powercmd.ExecuteCommand("c:\\windows\\system32\\powercfg.exe", "-setdcvalueindex 381b4222-f694-41f0-9685-ff5bb260df2e 4f971e89-eebd-4455-a8de-9e59040e7347 5ca83367-6e45-459f-a27b-476b1d01c936 2");
                NAVinfoKey.SetValue("stromvalgDC", "2");
            }
            if (checkedListBox2.SelectedIndex == 3)
            {
                
                var runpowercmd = powercmd.ExecuteCommand("c:\\windows\\system32\\powercfg.exe", "-setdcvalueindex 381b4222-f694-41f0-9685-ff5bb260df2e 4f971e89-eebd-4455-a8de-9e59040e7347 5ca83367-6e45-459f-a27b-476b1d01c936 3");
                NAVinfoKey.SetValue("stromvalgDC", "3");
            }

            button9.Text = "Lagret OK";
        }


        //resetter outlook
        private void button5_Click(object sender, EventArgs e)
        {
            button5.Text = "Resetter..";
            button5.BackColor = Color.Orange;
            var command = new functions();
            bool FullClean = false;

            try
            {
                if (checkBox1.Checked == true)
                {
                    FullClean = true;
                }

                textBox1.AppendText(command.resetOutlook(FullClean));
                button5.BackColor = Color.Green;
                button5.Text = "Resatt OK";
            }
            catch (Exception err)
            {
                button5.Text = "Feilet";
                button5.BackColor = Color.Red;
                textBox1.AppendText(err.ToString());
            }

        }

        private void button10_Click(object sender, EventArgs e)
        {
            var command = new functions();
            RegistryKey NAVinfoKey = Registry.CurrentUser.OpenSubKey(@"Software\NAVinfo", true);

            try
            {
                var resett = command.ExecuteCommand("c:\\windows\\system32\\powercfg.exe", "-restoredefaultschemes");
                button10.Text = "Resatt";
                NAVinfoKey.SetValue("stromvalgAC", "1");
                NAVinfoKey.SetValue("stromvalgDC", "2");
                if (!(resett))
                {
                    button10.Text = "Feilet";
                }
            }
            catch
            {
                button10.Text = "Feilet";
            }


        }


        //resetter wifi + VPN
        private void button3_Click(object sender, EventArgs e)
        {
            button3.Text = "Resetter..";
            button3.BackColor = Color.Orange;
            functions reset = new functions();
            //textBox1.AppendText(reset.resetWiFi()) ;
            //textBox1.AppendText(reset.resetF5VPN());
            eventLog1.Source = "NAV-Status";
            eventLog1.WriteEntry("Reset WiFi", EventLogEntryType.Information, 1 );

            button3.Text = "Resatt OK";
            button3.BackColor = Color.Green;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("c:\\windows\\notepad.exe", "c:\\temp\\NAV-User.log");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("c:\\windows\\notepad.exe", "c:\\temp\\NAV-System.log");
        }

        private void button11_Click(object sender, EventArgs e)
        {
            if (File.Exists(@"C:\temp\nav-user.log"))
            {
                File.Delete(@"c:\temp\nav-user.log");
                            }

            if (File.Exists(@"C:\temp\nav-system.log"))
            {
                File.Delete(@"c:\temp\nav-system.log");
            }

            if (Directory.Exists(@"C:\temp\logger"))
            {
                Directory.Delete(@"c:\temp\logger", true);
            }

            foreach (string f in Directory.EnumerateFiles(@"c:\temp\", "logger*.zip"))
            {
                File.Delete(f);
            }


            button11.Text = "Slettet !!";
        }

        private void button12_Click(object sender, EventArgs e)
        {
            // Samler NAV logger
            var bruker = Environment.UserName;
            
            try
            {
                if(!(Directory.Exists("C:\\temp\\logger")))
                {
                    System.IO.Directory.CreateDirectory("C:\\temp\\logger");
                }
                    
                File.Copy("C:\\temp\\NAV-user.log", "C:\\temp\\logger\\NAV-user.log", true);
                File.Copy("C:\\temp\\NAV-system.log", "C:\\temp\\logger\\NAV-system.log", true);
               
            }
            catch (Exception ex)
            {
                textBox1.Text += "Klarte ikke å kopiere NAV-User.log og NAV-System.log: " +ex;
            }

            //samler Application eventloggen
            try
            {
                EventLog evtLog = new EventLog("Application");  // Event Log type
                evtLog.MachineName = ".";  // dot is local machine
                string path = "C:\\temp\\logger\\Application.log";
                if (File.Exists(path))
                {
                    File.Delete(path);
                }

                
                using (StreamWriter sw = File.CreateText(path))
                {
                    foreach (EventLogEntry evtEntry in evtLog.Entries)
                    {
                        if (evtEntry.TimeWritten > DateTime.Now.AddDays(-1))
                        {
                            sw.WriteLine(evtEntry.TimeWritten + "#" + evtEntry.EntryType.ToString() + "#" + evtEntry.Source.ToString() + "#" + evtEntry.Message.ToString() + ";");

                        }



                    }
                }

                evtLog.Close();
            }
            catch (Exception ex)
            {
                textBox1.Text += "Klarte ikke å hente inn application event loggen" + ex;
            }

            //SAmler System evnt loggen
            try
            {
                EventLog evtLog = new EventLog("System");  // Event Log type
                evtLog.MachineName = ".";  // dot is local machine
                string path = "C:\\temp\\logger\\System.log";
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
                
                using (StreamWriter sw = File.CreateText(path))
                {
                    foreach (EventLogEntry evtEntry in evtLog.Entries)
                    {

                        if (evtEntry.TimeWritten > DateTime.Now.AddDays(-1))
                        {
                            sw.WriteLine(evtEntry.TimeWritten + "#" + evtEntry.EntryType.ToString() + "#" + evtEntry.Source.ToString() + "#" + evtEntry.Message.ToString() + ";");

                        }

                    }
                }

                evtLog.Close();
            }
            catch (Exception ex)
            {
                textBox1.Text += "Klarte ikke å hente inn System event loggen " + ex;
            }

            var sysinfofile = "C:\\temp\\sysinfo.cmd";
            try
            {
                if (!(File.Exists(sysinfofile)))
                {
                    File.WriteAllText(sysinfofile, "systeminfo >> c:\\temp\\logger\\systeminfo.txt" + Environment.NewLine);
                    File.AppendAllText(sysinfofile, "ipconfig /all >> c:\\temp\\logger\\ipconfig.txt" + Environment.NewLine);
                    File.AppendAllText(sysinfofile, "tasklist >> c:\\temp\\logger\\tasklist.txt" + Environment.NewLine);
                    File.AppendAllText(sysinfofile, "netsh wlan show interfaces >> c:\\temp\\logger\\wireless.txt" + Environment.NewLine);
                    File.AppendAllText(sysinfofile, "netsh wlan show networks >> c:\\temp\\logger\\tilgjengeligeWlan.txt" + Environment.NewLine);
                    File.AppendAllText(sysinfofile, "wmic startup >> c:\\temp\\logger\\startup.txt" + Environment.NewLine);



                }
                System.Diagnostics.Process.Start(sysinfofile).WaitForExit(15000);
            }
            catch (Exception ex)
            {
                textBox1.Text += "Samling av SysInf feilet" + ex;
            }





            try
            {

                Random rnd = new Random();
                int random = rnd.Next(1, 999999);


                System.IO.Compression.ZipFile.CreateFromDirectory("C:\\temp\\logger", "c:\\temp\\logger_" + random + ".zip");



                Outlook.Application app = new Outlook.Application();
                
                Outlook.MailItem mailItem = app.CreateItem(Outlook.OlItemType.olMailItem);
                //Outlook.MailItem mailItem = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
                mailItem.Subject = "Diagnostic informasjon fra " + Environment.MachineName + " " + bruker;
                //mailItem.To = "someone@example.com";
                //mailItem.Body = "This is the message.";
                mailItem.Attachments.Add("c:\\temp\\logger_" + random + ".zip");
                mailItem.Importance = Outlook.OlImportance.olImportanceLow;
                mailItem.Display(true);

                File.Delete("c:\\temp\\logger_" + random + ".zip");

                if (Directory.Exists(@"C:\temp\logger"))
                {
                    Directory.Delete(@"c:\temp\logger", true);
                }

            }
            catch (Exception ex)
            {
                textBox1.Text += "Klarte ikke å lage mail til brukerstøtte " + ex;
            }

        }

        private void button13_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("ms-settings:printers");
        }

        private void button14_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("ms-settings:powersleep");
        }

        private void button15_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("companyportal:");
        }



        private void button16_Click_1(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://forms.office.com/Pages/ResponsePage.aspx?id=NGU2YsMeYkmIaZtVNSedC00WkBGW6YxFtyQ4FAm2fj1UQzhONjlXT0U1SDBMWU9ZWktTVEdMMlFIVy4u");
            
        }

        private void button17_Click(object sender, EventArgs e)
        {
            button17.BackColor = Color.Orange;
            button17.Text = "Resetter..";
            var command = new functions();
            var resettSM = command.ExecuteCommand("powershell", "-executionpolicy bypass c:\\Windows\\Mob\\ResetStartmenyPS.ps1");
            if(resettSM)
            {
                button17.BackColor = Color.Green;
                button17.Text = "OK";
            }
            if (!(resettSM))
            {
                button17.BackColor = Color.Red;
                button17.Text = "Feilet";
                textBox1.Text = "Restting av Start-Menyen feilet";
            }

        }

        private void button18_Click(object sender, EventArgs e)
        {
            if (!(Directory.Exists("C:\\temp\\logger")))
            {
                System.IO.Directory.CreateDirectory("C:\\temp\\logger");
            }

            ScreenCapture sc = new ScreenCapture();
            // capture entire screen, and save it to a file
            Image img = sc.CaptureScreen();
            // display image in a Picture control named imageDisplay
            this.pictureBox5.Image = img;
            
            // capture this window, and save it
            //sc.CaptureWindowToFile(this.Handle, "C:\\temp\\logger\\screen.gif", ImageFormat.Gif);
            sc.CaptureScreenToFile("C:\\temp\\logger\\screen.gif", ImageFormat.Gif);
            
        }

        private void button19_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(@"softwarecenter:");
        }

        private void label14_Click(object sender, EventArgs e)
        {

        }

        private void label14_Click_1(object sender, EventArgs e)
        {

        }

        private void label28_Click(object sender, EventArgs e)
        {

        }

        private void button20_Click(object sender, EventArgs e)
        {

            button20.Text = "Resetter ...";
            eventLog1.Source = "NAV-Status";
            eventLog1.WriteEntry("Reset SCCM", EventLogEntryType.Information, 11);

            button20.Text = "OK";
            button20.BackColor = Color.Green;
        }

        private void label21_Click(object sender, EventArgs e)
        {

        }
        //Fjerner UAC dimming
        private void button21_Click(object sender, EventArgs e)
        {
            button21.Text = "UAC";
            button21.BackColor = Color.Orange;
            functions reset = new functions();
            //textBox1.AppendText(reset.resetWiFi()) ;
            //textBox1.AppendText(reset.resetF5VPN());
            eventLog1.Source = "NAV-Status";
            eventLog1.WriteEntry("Endre UAC", EventLogEntryType.Information, 21);

            button21.Text = "OK";
            button21.BackColor = Color.Green;
        }

        private void label30_Click(object sender, EventArgs e)
        {

        }
    }

    }

