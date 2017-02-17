using OpenPop.Common;
using OpenPop.Mime;
using OpenPop.Pop3;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DesktopEmail
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }
        #region Load
        private void Form1_Load(object sender, EventArgs e)
        {
            checkBox1.Checked = Properties.Settings.Default.checkbox;
            textBox5.Text = Properties.Settings.Default.Password;
            textBox6.Text = Properties.Settings.Default.Email;
        }
        #endregion
        #region zapis
        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            Properties.Settings.Default.Save();
            if (checkBox1.Checked == true)
            {
                Properties.Settings.Default.checkbox = checkBox1.Checked;
                Properties.Settings.Default.Email = textBox6.Text;
                Properties.Settings.Default.Password = textBox5.Text;
                Properties.Settings.Default.Save();
            }
            else if (checkBox1.Checked == false)
            {
                textBox6.Clear();
                textBox5.Clear();
                Properties.Settings.Default.Email = textBox6.Text;
                Properties.Settings.Default.Password = textBox5.Text;
                Properties.Settings.Default.Save();
            }
        }
        #endregion
        #region Wysyłanie Maila
        private void SendMail()
        {

            try
            {
                MailMessage message = new MailMessage();
                message.From = new MailAddress(textBox6.Text);
                message.Subject = textBox2.Text;
                message.Body = richTextBox1.Text;
                message.IsBodyHtml = true;
                message.BodyEncoding = Encoding.UTF8;
                foreach (string s in textBox1.Text.Split(';'))
                    message.To.Add(s);
                if (textBox3.Text != "")
                {
                    message.Attachments.Add(new Attachment(textBox3.Text));
                }
                SmtpClient client = new SmtpClient();
                client.Credentials = new NetworkCredential(textBox6.Text, textBox5.Text);
                client.Timeout = 10000;
                //client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.Host = comboBox1.Text;
                client.Port = 587;
                client.EnableSsl = true;
                client.Send(message);
                MessageBox.Show("Wysłano E-mail.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region Zmiana na HTML
        private void zmien(string a, string b)
        {
            #region Old
            //string[] lines = richTextBox1.Lines;
            //for (int i = 0; i < lines.Length; i++)
            //{
            //    if (richTextBox1.Text.Length == 0)
            //    {
            //        for (int j = 0; i < lines.Length; i++)
            //            lines[i] = Environment.NewLine;
            //    }
            //    else
            //        lines[i] = a  + lines[i] + b;
            //}
            //richTextBox1.Lines = lines;
            //richTextBox1.SelectedText = "test" + lines;
            //richTextBox1.SelectionStart = 0;
            // richTextBox1.SelectionLength = richTextBox1.Text.Length;
            // if (richTextBox1.Text.Length > 0)
            //     {
            //    richTextBox1.Text = a + " " + richTextBox1.SelectedText + " " + b;
            //     }
            #endregion

            if (richTextBox1.SelectedText.Length > 0)
            {
                richTextBox1.SelectedText = a + " " + richTextBox1.SelectedText + " " + b;
            }

            
            
        }
        #endregion


        private void button2_Click(object sender, EventArgs e)
        {
            SendMail();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolTip ToolTip1 = new System.Windows.Forms.ToolTip();
            ToolTip1.SetToolTip(this.button3, "rozmiar czcionki");
            zmien("<font size=6>", "</font>");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox3.Text = openFileDialog1.FileName.ToString();
            }
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            webBrowser1.DocumentText = richTextBox1.Text;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolTip ToolTip2 = new System.Windows.Forms.ToolTip();
            ToolTip2.ShowAlways = true;
            ToolTip2.SetToolTip(this.button4, "Pogrubienie tekstu");
            zmien("<b>", "</b>");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolTip ToolTip3 = new System.Windows.Forms.ToolTip();
            ToolTip3.ShowAlways = true;
            ToolTip3.SetToolTip(this.button5, "Podkreślenie tekstu");
            zmien("<u>", "</u>");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolTip ToolTip4 = new System.Windows.Forms.ToolTip();
            ToolTip4.ShowAlways = true;
            ToolTip4.SetToolTip(this.button6, "Ustawia tło strony");
            zmien("<body style=background-color:blue>", "</body>");
        }

        private void button7_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolTip ToolTip5 = new System.Windows.Forms.ToolTip();
            ToolTip5.ShowAlways = true;
            ToolTip5.SetToolTip(this.button7, "Tekst pochylony");
            zmien("<i>", "</i>");
        }

        private void button8_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolTip ToolTip6 = new System.Windows.Forms.ToolTip();
            ToolTip6.ShowAlways = true;
            ToolTip6.SetToolTip(this.button8, "Zmiana koloru czcionki");
            zmien("<font color=red>", "</font>");
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox2.Text = "Wybierz czcionkę";
            System.Windows.Forms.ToolTip ToolTip7 = new System.Windows.Forms.ToolTip();
            ToolTip7.ShowAlways = true;
            ToolTip7.SetToolTip(this.comboBox2, "rodzaj czcionki");

            switch (comboBox2.SelectedIndex)
            {
                case 0:
                    zmien("<font face= Verdana>", "</font>");
                    break;
                case 1:
                    zmien("<font face= Arial>", "</font>");
                    break;
                case 2:
                    zmien("<font face= Times New Roman>", "</font>");
                    break;
            }
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolTip ToolTip8 = new System.Windows.Forms.ToolTip();
            ToolTip8.ShowAlways = true;
            ToolTip8.SetToolTip(this.comboBox3, "Rozmiar nagłówka");
            switch (comboBox3.SelectedIndex)
            {
                case 0:
                    zmien("<h1>", "</h1>");
                    break;
                case 1:
                    zmien("<h2>", "</h2>");
                    break;
                case 2:
                    zmien("<h3>", "</h3>");
                    break;
                case 3:
                    zmien("<h4>", "</h4>");
                    break;
            }
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolTip ToolTip9 = new System.Windows.Forms.ToolTip();
            ToolTip9.ShowAlways = true;
            ToolTip9.SetToolTip(this.comboBox4, "Wyrównanie tekstu");
            switch (comboBox4.SelectedIndex)
            {
                case 0:
                    zmien("<center>", "</center>");
                    break;
                case 1:
                    zmien("<p align=left>", " </p>");
                    break;
                case 2:
                    zmien("<p align=right>", "</p>");
                    break;
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolTip ToolTip10 = new System.Windows.Forms.ToolTip();
            ToolTip10.ShowAlways = true;
            ToolTip10.SetToolTip(this.button9, "Przekreślenie tekstu");
            zmien("<del>", "</del>");
        }

        private void button10_Click(object sender, EventArgs e)
        {
            string a = "<li>";
            string b = "</li>";
            System.Windows.Forms.ToolTip ToolTip11 = new System.Windows.Forms.ToolTip();
            ToolTip11.ShowAlways = true;
            ToolTip11.SetToolTip(this.button10, "Puktowanie tekstu");
            zmien(a, b);
        }

        private DataTable dt;
       private Pop3Client client = new Pop3Client();

        private void button11_Click(object sender, EventArgs e)
        {


            dt = new DataTable("Inbox");
            dt.Columns.Add("ID");
            dt.Columns.Add("Temat");
            dt.Columns.Add("Sender");
            dt.Columns.Add("Email");
            dt.Columns.Add("Tekst");
            dt.Columns.Add("Czas");
            dataGridView1.DataSource = dt;

            try
            {

                client.Connect(comboBox5.Text, 995, true);
                client.Authenticate(textBox6.Text, textBox5.Text, OpenPop.Pop3.AuthenticationMethod.UsernameAndPassword);
                int count = client.GetMessageCount();

                string htmlContained = "";
             

               
                
                    if (client.Connected)
                    {
                    //   for (int i = count; i > count - 100 && i >= 0; i--)
                    //     for (int i = 1;  i <=100 && i <= count; i--)
                    //  for (int i = 1; i <= count && i <= 100 ; i++)
                    // for (int i = count; i >= 100; i--)
                

                    for (int i = count; i > count - Convert.ToInt32(textBox4.Text) && i >= 1; i--)
                         {
                       
                            OpenPop.Mime.Message message = client.GetMessage(i);

                        
                            OpenPop.Mime.MessagePart html = message.FindFirstHtmlVersion();
                        OpenPop.Mime.MessagePart file = message.FindFirstMessagePartWithMediaType("") ;


                            if (html != null)
                            {

                                htmlContained = html.GetBodyAsText();
                        
                            
                            }
                            else
                            {
                                html = message.FindFirstPlainTextVersion();
                               
                                    htmlContained = html.GetBodyAsText();
                           

                            
                        }
                      
                            string name = message.Headers.Subject;
                            if (name == "")
                            {
                                name = "Brak Tematu";

                           
                            }

                            string nadawca = message.Headers.From.DisplayName;

                            if (nadawca == "")
                            {
                                nadawca = "Brak Informacji";
                            }

                            dt.Rows.Add(new object[] { i.ToString(), name.ToString(), nadawca.ToString(), message.Headers.From.Address, htmlContained, message.Headers.DateSent });

                        }
                    }
                    client.Disconnect();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
             

            }



        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            Form1 form = new Form1();
            if (e.RowIndex < dt.Rows.Count && e.RowIndex >= 0)
            {
                webBrowser2.DocumentText = dt.Rows[e.RowIndex]["Tekst"].ToString();
            }
      


            }

        private void dataGridView1_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            
                Form2 form = new Form2(dt.Rows[e.RowIndex]["Tekst"].ToString());
                form.Show();
         
        }

        private void button11_MouseClick(object sender, MouseEventArgs e)
                   {
            try
            {
                if (client.Connected)
                    client.Disconnect();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;

            if (!Char.IsDigit(ch) && ch != 8 && ch != 44)
            {
                e.Handled = true;
            }
        }
    }
}


