using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ComWrapper1C
{
    partial class StartDialog : Form
    {
        public Dictionary<string, string> connectionData = new Dictionary<string, string>(4);
        public bool close = false;
        public bool dataDone = false;

        public StartDialog(string dbName)
        {
            InitializeComponent();
            close = false;
            db.Text = dbName;
            server.Text = "srv-kasup";
            textBox1.Text = "C:\\Program Files (x86)\\COM_CADLib1C\\Template CADLib DB\\Base with archive.cde";
            connectionData.Add("db_template", textBox1.Text);
        }

        private void create_Click(object sender, EventArgs e)
        {
            connectionData.Add("server", server.Text);
            connectionData.Add("db", db.Text.Replace(',', '_'));
            if (this.radioButton2.Checked == true)
            {
                if (connectionData.ContainsKey("db_login") == false && connectionData.ContainsKey("db_password") == false)
                {
                    connectionData.Add("db_login", db_login.Text);
                    connectionData.Add("db_password", db_password.Text);
                }
                else
                {
                    connectionData["db_login"] = db_login.Text;
                    connectionData["db_password"] = db_password.Text;
                }
            }
            else
            {
                if (connectionData.ContainsKey("db_login") == false && connectionData.ContainsKey("db_password") == false)
                {
                    connectionData.Add("db_login", "null");
                    connectionData.Add("db_password", "null");
                }
                else
                {
                    connectionData["db_login"] = "null";
                    connectionData["db_password"] = "null";
                }
            }
            dataDone = true;
            this.Close();
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked == true)
            {
                radioButton1.Checked = false;
                db_login.Enabled = true;
                db_password.Enabled = true;
            }
            else
            {
                radioButton1.Checked = true;
                db_login.Enabled = false;
                db_password.Enabled = false;
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked == true)
            {
                radioButton2.Checked = false;
                db_login.Enabled = false;
                db_password.Enabled = false;
            }
            else
            {
                radioButton2.Checked = true;
                db_login.Enabled = true;
                db_password.Enabled = true;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog getTemplateDialog = new OpenFileDialog();
            getTemplateDialog.Multiselect = true;
            DialogResult result = getTemplateDialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                if (connectionData.ContainsKey("db_template") == false)
                {
                    connectionData.Add("db_template", getTemplateDialog.FileName);
                }
                else connectionData["db_template"] = getTemplateDialog.FileName;
                textBox1.Text = getTemplateDialog.FileName;
            }
        }

        private void server_TextChanged(object sender, EventArgs e)
        {
            if (server.Text.Length > 0)
            {
                db.Enabled = true;
                if (db.Text.Length > 0)
                {
                    radioButton1.Enabled = true;
                    radioButton2.Enabled = true;
                    button2.Enabled = true;
                    if (radioButton1.Checked == true)
                    {
                        if (textBox1.Text.Split('.')[textBox1.Text.Split('.').Length - 1] == "cde")
                        {
                            button1.Enabled = true;
                        }
                    }
                    else
                    {
                        if (db_login.Text.Length > 0)
                        {
                            if (textBox1.Text.Split('.')[textBox1.Text.Split('.').Length-1] == "cde")
                            {
                                button1.Enabled = true;
                            }
                        }
                    }
                }
            }
            else
            {
                db.Enabled = false;
                radioButton1.Enabled = false;
                radioButton2.Enabled = false;
                button2.Enabled = false;
                button1.Enabled = false;
            }
        }

        private void db_TextChanged(object sender, EventArgs e)
        {
            if (db.Text.Length > 0)
            {
                radioButton1.Enabled = true;
                radioButton2.Enabled = true;
                button2.Enabled = true;
                if (radioButton1.Checked == true)
                {
                    if (textBox1.Text.Split('.')[textBox1.Text.Split('.').Length - 1] == "cde")
                    {
                        button1.Enabled = true;
                    }
                }
                else
                {
                    if (db_login.Text.Length > 0)
                    {
                        if (textBox1.Text.Split('.')[textBox1.Text.Split('.').Length - 1] == "cde")
                        {
                            button1.Enabled = true;
                        }
                    }
                }
            }
            else
            {
                radioButton1.Enabled = false;
                radioButton2.Enabled = false;
                button2.Enabled = false;
                button1.Enabled = false;
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            string[] checkCDE = textBox1.Text.Split('.');
            if (checkCDE[checkCDE.Length-1] == "cde")
            {
                button1.Enabled = true;
            }
            else
            {
                button1.Enabled = false;
                MessageBox.Show("Шаблон должен быть в формате *.cde", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void StartDialog_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (dataDone != true) close = true;
        }
    }
}
