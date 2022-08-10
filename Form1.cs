using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace CRM_Publisher_V2
{
    public partial class Form1 : Form
    {
        public string Email;
        public string Password;
        public Form1()
        {
            InitializeComponent();
        }

        private void btn_Cancel2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btn_OK_Click(object sender, EventArgs e)
        {
            //Form1 form = new Form1();

            if (!IsValid(textBox1.Text))
            {
                MessageBox.Show("Please enter a valid email address!");
            }
            else
            {
                Email = textBox1.Text;
                this.Close();
            }
            Password = textBox2.Text;
          
            

            
        }
        
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            //MailAddress zae = new MailAddress("nizar.s@ablaviation.com");
            updateButton();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            updateButton();
        }
        private void updateButton()
        {
            
            btnOK.Enabled = textBox1.Text != string.Empty && textBox2.Text != string.Empty ;
            
        }
      
        private static bool IsValid(string email)
        {

            
            if(email.Length > 17)
            {
                if(email.Substring(email.Length-16,16)=="@ablaviation.com")
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }
    }
}
