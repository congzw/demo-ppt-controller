using System;
using System.IO;
using System.Windows.Forms;

namespace DemoApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            this.TopMost = true;
            Helper = new PptHelper();
        }

        public PptHelper Helper { get; set; }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            Helper.Open(Path.GetFullPath("Hello.pptx"));
        }

        private void btnPre_Click(object sender, EventArgs e)
        {
            Helper.Pre();
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            Helper.Next();
        }

        private void btnInfo_Click(object sender, EventArgs e)
        {
            var pptInfo = Helper.GetPptInfo();
            this.textBox1.Text = pptInfo.ToString();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Helper.Close();
        }
    }
}
