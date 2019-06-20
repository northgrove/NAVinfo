using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;


namespace NAVinfo
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            if (System.Diagnostics.Process.GetProcessesByName(System.IO.Path.GetFileNameWithoutExtension(System.Reflection.Assembly.GetEntryAssembly().Location)).Count() > 1) System.Diagnostics.Process.GetCurrentProcess().Kill();
            InitializeComponent();
            this.Hide();

        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Hovedmeny hmeny = new Hovedmeny();
            hmeny.Show();
            hmeny.Activate();
        }

        // private void opplæringToolStripMenuItem_Click(object sender, EventArgs e)
        // {
        //     System.Diagnostics.Process.Start("https://navno.sharepoint.com/sites/intranett-it/SitePages/Office365.aspx");
        // }
        // private void toolStripMenuItem4_Click(object sender, EventArgs e)
        // {
        //     // oppretter eventid 1 i NAV source for Reset nettverkstilkobling
        //     eventLog1.Source = "NAV-Status";
        //     eventLog1.WriteEntry("Reset nettverkstilkobling", EventLogEntryType.Information, 1);
        // }
        // private void toolStripMenuItem5_Click(object sender, EventArgs e)
        // {
        //     // laster https://aka.ms/sspr i eget browser vindu
        //     // Gammel kode: byttpassord passordbytte = new byttpassord();
        //     // Gammel kode: passordbytte.Show();
        //     System.Diagnostics.Process.Start("https://aka.ms/sspr");
        // }
        private void toolStripMenuItem6_Click(object sender, EventArgs e)
        {
            // oppretter eventid 25 i NAV source for Reset Outlookprofil
            eventLog1.Source = "NAV-Status";
            eventLog1.WriteEntry("Reset Outlookprofil", EventLogEntryType.Information, 25);
        }
        private void toolStripMenuItem7_Click(object sender, EventArgs e)
        {
            // oppretter eventid 22 i NAV source for Reparere Bitlocker
            eventLog1.Source = "NAV-Status";
            eventLog1.WriteEntry("Reparer Bitlocker", EventLogEntryType.Information, 22);
        }

        // private void toolStripMenuItem8_Click(object sender, EventArgs e)
        // {
        //     // oppretter eventid 23 i NAV source for reset Chrome Nettleserdata
        //     eventLog1.Source = "NAV-Status";
        //     eventLog1.WriteEntry("Reset Chrome Nettleserdata", EventLogEntryType.Information, 23);
        // }

        private void toolStripMenuItem9_Click(object sender, EventArgs e)
        {
            // oppretter eventid 24 i NAV source for Aktiver Basic Auth
            eventLog1.Source = "NAV-Status";
            eventLog1.WriteEntry("Aktiver Basic Auth", EventLogEntryType.Information, 24);
        }
        private void toolStripMenuItem10_Click(object sender, EventArgs e)
        {
            // Terminerer NAVstatus.exe
            // System.Diagnostics.Process.Start("notepad.exe");
            Application.Exit();
        }
        private void toolStripMenuItem11_Click(object sender, EventArgs e)
        {
            // Terminerer NAVstatus.exe
            // System.Diagnostics.Process.Start("notepad.exe");
            eventLog1.Source = "NAV-Status";
            eventLog1.WriteEntry("Aktiver Ethernet", EventLogEntryType.Information, 20);
        }
        private void toolStripMenuItem12_Click(object sender, EventArgs e)
        {
            // Terminerer NAVstatus.exe
            // System.Diagnostics.Process.Start("notepad.exe");
            eventLog1.Source = "NAV-Status";
            eventLog1.WriteEntry("Aktiver UAC Fjernkontroll", EventLogEntryType.Information, 21);
        }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {

        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            // åpner hovedmenyen ved dobbeltklikk på tray-iconet
            Hovedmeny hmeny = new Hovedmeny();
            hmeny.Show();
            hmeny.Activate();
        }
        private void notifyIcon1_MouseClick(object sender, MouseEventArgs e)
        {
            //åpner hovedmenyen ved left-klikk på tray-iconet
            if (e.Button == MouseButtons.Left)
            {
                
                Hovedmeny hmeny = new Hovedmeny();
                hmeny.Show();
                hmeny.Activate();
            }

        }

        // private void applikasjonerToolStripMenuItem_Click(object sender, EventArgs e)
        // {
        //     System.Diagnostics.Process.Start(@"softwarecenter:");
        // }
    }
}
