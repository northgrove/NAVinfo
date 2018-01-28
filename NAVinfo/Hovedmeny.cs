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

namespace NAVinfo
{
    public partial class Hovedmeny : Form
    {
       

        public Hovedmeny()
        {
            InitializeComponent();
            tabControl1.SelectedIndexChanged += new EventHandler(tabControl1_SelectedIndexChanged);
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

        private void button4_Click(object sender, EventArgs e)
        {
            // laster passord.nav.no i eget browser vindu
            byttpassord passordbytte = new byttpassord();
            passordbytte.Show();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://navno.sharepoint.com/Office365");
        }

        private void button7_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://navno.sharepoint.com/Office365");
        }


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


        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch((sender as TabControl).SelectedIndex)
            {
                case 1:
                    GetInstalledApps();
                    break;
                case 4:
                    lesEventlog();
                    break;
            }
            

        }

    }

    }

