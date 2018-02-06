using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Management;
using System.Net;
using System.Net.NetworkInformation;
using Microsoft.Win32;
using System.Diagnostics;

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
           
            }

            
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
                label14.Text = label14.Text + addr + Environment.NewLine;
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

            // Sjekker om betteriet har status "lader"
            var battery = BatteryChargeStatus.Charging;
            if (battery.ToString() == "Charging")
            {
                labelstrom.BackColor = Color.Green;
                labelstrom.Text = "Tilkoblet";
            }
            else
            {
                labelstrom.BackColor = Color.Red;
                labelstrom.Text = "Ikke tilkoblet";
            }





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
            System.Diagnostics.Process.Start("https://navno.sharepoint.com/Office365hjelp");
        }

        // link til office 365 hjelp
        private void button7_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://navno.sharepoint.com/Office365hjelp");
        }




        // Trigger Functions ut i fra hvilken fane som velges
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch ((sender as TabControl).SelectedIndex)
            {
                case 1:
                    GetInstalledApps();
                    break;
                case 5:
                    lesEventlog();
                    break;
                case 3:
                    mapPrint();
                    break;
                case 6:
                    openStrom();
                    break;
            }


        }




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

        // Leser Eventloggen og viser i fane
        public void lesEventlog()
        {
            listView2.Columns.Add("EventID", 100);
            listView2.Columns.Add("Source", 100);
            listView2.Columns.Add("Feil", 200);
            foreach (System.Diagnostics.EventLogEntry log in eventLog1.Entries)
            {
                string[] arr = new string[3];
                ListViewItem item;
                arr[0] = log.InstanceId.ToString();
                arr[1] = log.Source.ToString();
                arr[2] = log.Message.ToString();
                item = new ListViewItem(arr);
                listView2.Items.Add(item);
            }
        }


        // Sjekker om NAV Skrivere er installert
        private void mapPrint()
        {
            listViewPrint.Columns.Add("Printer", 100);
            listViewPrint.Columns.Add("Status", 200);

            var print1 = Printer.IsPrinterInstalled("Fax");
            if (print1)
            {
                string[] printer = { "Fax", "Tilkoblet" };
                var listViewItem = new ListViewItem(printer);
                listViewPrint.Items.Add(listViewItem);
            }
            else
            {
                string[] printer = { "Fax", "Frakoblet" };
                var listViewItem = new ListViewItem(printer);
                listViewPrint.Items.Add(listViewItem);
            }

            var print2 = Printer.IsPrinterInstalled("FargeDuplex");
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
            var print3 = Printer.IsPrinterInstalled("SortDuplex");
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
            var print4 = Printer.IsPrinterInstalled("FargeSimplex");
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
            var print5 = Printer.IsPrinterInstalled("SortSimplex");
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
            var stromvalgDC = NAVinfoKey.GetValue("stromvalgDC");
            switch(stromvalgAC)
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
            listViewPrint.Items.Clear();

                var status = "";
                var fargeduplex = Printer.AddPrinter("\\\\a01psvw005.adeo.no\\FargeDuplex IKSS");
                    if (fargeduplex)
                    {
                        status = "Tilkoblet";
                    }
                    else
                    {
                        status = "feilet";
                    }
                string[] printer = { "FargeDuplex", status };
                var listViewItem = new ListViewItem(printer);
                listViewPrint.Items.Add(listViewItem);
     
            
            //Printer.AddPrinter("\\\\a01psvw005.adeo.no\\SortDuplex IKSS");


        }

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
            var command = new functions();

            try
            {
                command.ExecuteCommand("pskill", "/accepteula outlook");
                command.ExecuteCommand("Powershell", @"remove-item HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook\Profiles\Outlook -Recurse -Force");

                textBox1.AppendText(command.resetOutlook());



                command.ExecuteCommand(@"c:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE", "");



                button5.Text = "Resatt OK";
            }
            catch
            {
                button5.Text = "Feilet";
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
            functions reset = new functions();
            textBox1.AppendText(reset.resetWiFi()) ;
            textBox1.AppendText(reset.resetF5VPN());
            button3.Text = "Resatt";
        }
    }

    }

