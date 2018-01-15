using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace InvAddIn
{
    public partial class FileBrowser : Form
    {
        public FileBrowser(string s)
        {
            InitializeComponent();
            openFileDialog1.InitialDirectory = s;
            openFileDialog1.Filter = "Autodesk Inventor Assemblies (*.iam)|*.iam";            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.DialogResult dr = openFileDialog1.ShowDialog();

            if (dr == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void FileBrowser_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
        }

        public string getFileChosen()
        {
            return openFileDialog1.FileName;
        }
       

    }
}
