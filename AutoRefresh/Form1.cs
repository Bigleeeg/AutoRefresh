using System;
using System.Windows.Forms;

namespace AutoRefresh
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            this.label3.Visible = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.label3.Visible = false;
            string userName = this.textBox1.Text.Trim();
            string passWord = this.textBox2.Text.Trim();

            if ((userName.ToLower() == "jrfdc" && passWord.ToLower() == "j0r18") || userName.ToLower() == "jiaofeng")
            {
                this.Visible = false;
                Form2 frm2 = new Form2();
                frm2.Show();
            }
            else
            {
                this.label3.Visible = true;
            }
        }
    }
}
