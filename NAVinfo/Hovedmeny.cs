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

namespace NAVinfo
{
    public partial class Hovedmeny : Form
    {
        public Hovedmeny()
        {
            InitializeComponent();
        }

        private void Hovedmeny_Load(object sender, EventArgs e)
        {

            ManagementObjectSearcher ComSerial = new ManagementObjectSearcher("SELECT * FROM Win32_BaseBoard");

            foreach (ManagementObject wmi in ComSerial.Get())
            {
                try
                {
                    label12.Text = wmi.GetPropertyValue("SerialNumber").ToString();
                }
                catch { }
            }


            label11.Text = Environment.MachineName;




            IPAddress[] localIPs = Dns.GetHostAddresses(Dns.GetHostName());
            foreach (IPAddress addr in localIPs)
            {


                label14.Text = label14.Text + addr + Environment.NewLine;

            }



            
            
            
           
         
        }

        }
    }

