using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace officeLibSample
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            officeLib.officeLib test = new officeLib.officeLib();
            List<officeLib.traceAct> func = test.getActivityRecord();
            bool officeStatus = test.chkInstall();
            MessageBox.Show("Installed Office : " + officeStatus.ToString());
            foreach (officeLib.traceAct result in func)
            {
                MessageBox.Show(result.usedFile + " / " + result.officeType + " / " + result.timeRecord);
            }

        }
    }
}
