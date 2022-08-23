using System;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Media;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Serialization;
using System.IO;

namespace UrlopyDelegacje
{
	public class Ustawienia : Form
	{
		private IContainer components = null;

		private GroupBox groupBox1;

		private TextBox textBox1;

		private Button button1;
        private GroupBox groupBox2;
        private TextBox textBox2;
        private TextBox textBox4;
        private TextBox textBox3;
        private Button button3;
        private CheckBox checkBox1;
        private Button button2;

		public Ustawienia()
		{
			InitializeComponent();
            Form1 form = new Form1();
		}

		private void Button1_Click(object sender, EventArgs e)
		{
			try
			{
				Form1 form = Application.OpenForms.OfType<Form1>().FirstOrDefault();
				form.WrtieConfigXML(Convert.ToInt32(textBox1.Text));
				SystemSounds.Hand.Play();
				MessageBox.Show("Wartość zapisana", "Sukces");
				form.FillForm();
                if (checkBox1.Checked == true)
                {
                    CreateAPICOnfig("create");
                }
                form.GetKraj();
                Close();
			}
			catch
			{
				SystemSounds.Beep.Play();
				MessageBox.Show("Nieprawidłowa wartość", "Błąd");
			}
		}

		private void Ustawienia_Load(object sender, EventArgs e)
		{
			textBox1.Text = Form1.dniurlwsz.ToString();
            ChckAPIConfig();
        }

		private void Button2_Click(object sender, EventArgs e)
		{
			textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            textBox4.Clear();
        }

		protected override void Dispose(bool disposing)
		{
			if (disposing && components != null)
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		private void InitializeComponent()
		{
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ustawienia));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.button3 = new System.Windows.Forms.Button();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.textBox1);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(303, 52);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Ilość dni urlopowych";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(6, 19);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(291, 20);
            this.textBox1.TabIndex = 0;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(12, 200);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(144, 68);
            this.button1.TabIndex = 1;
            this.button1.Text = "Zapisz";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.Button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(171, 200);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(144, 68);
            this.button2.TabIndex = 2;
            this.button2.Text = "Wyczyść";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.Button2_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.button3);
            this.groupBox2.Controls.Add(this.checkBox1);
            this.groupBox2.Controls.Add(this.textBox4);
            this.groupBox2.Controls.Add(this.textBox3);
            this.groupBox2.Controls.Add(this.textBox2);
            this.groupBox2.Location = new System.Drawing.Point(12, 70);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(303, 124);
            this.groupBox2.TabIndex = 3;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Ustawienia API";
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(171, 15);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(126, 23);
            this.button3.TabIndex = 5;
            this.button3.Text = "Sprawdź połączenie";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(6, 19);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(58, 17);
            this.checkBox1.TabIndex = 4;
            this.checkBox1.Text = "Włącz";
            this.checkBox1.UseVisualStyleBackColor = true;
            this.checkBox1.Click += new System.EventHandler(this.checkBox1_Click);
            // 
            // textBox4
            // 
            this.textBox4.Location = new System.Drawing.Point(6, 94);
            this.textBox4.Name = "textBox4";
            this.textBox4.Size = new System.Drawing.Size(291, 20);
            this.textBox4.TabIndex = 2;
            this.textBox4.Text = "country";
            // 
            // textBox3
            // 
            this.textBox3.Location = new System.Drawing.Point(6, 68);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(291, 20);
            this.textBox3.TabIndex = 1;
            this.textBox3.Text = "api_key";
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(6, 42);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(291, 20);
            this.textBox2.TabIndex = 0;
            this.textBox2.Text = "URI";
            // 
            // Ustawienia
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(327, 280);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Ustawienia";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Ustawienia";
            this.Load += new System.EventHandler(this.Ustawienia_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);

		}

        private void checkBox1_Click(object sender, EventArgs e)
        {
            if (checkBox1.Checked == false)
            {
                DisableAPI(true);
                CreateAPICOnfig("remove");
            }
            else
            {
                DisableAPI(false);
                CreateAPICOnfig("create");
            }
        }

        private void CreateAPICOnfig(string _action)
        {
            Form1 form = Application.OpenForms.OfType<Form1>().FirstOrDefault();
            switch (_action)
            {
                case "create":
                    try
                    {
                        XmlDocument xmlDocument2 = new XmlDocument();
                        xmlDocument2.LoadXml("<API></API>");
                        XmlElement xmlElement3 = xmlDocument2.CreateElement("Value");
                        XmlAttribute xmlAttribute2 = xmlDocument2.CreateAttribute("URI");
                        xmlAttribute2.Value = textBox2.Text;
                        xmlElement3.Attributes.Append(xmlAttribute2);
                        xmlElement3.InnerText = textBox3.Text;
                        xmlDocument2.DocumentElement.AppendChild(xmlElement3);
                        XmlAttribute xmlAttribute3 = xmlDocument2.CreateAttribute("country");
                        xmlAttribute3.Value = textBox4.Text;
                        xmlElement3.Attributes.Append(xmlAttribute3);
                        xmlDocument2.PreserveWhitespace = true;
                        xmlDocument2.Save(form.APIconfigFileFullPath);
                    }
                    catch
                    {
                        SystemSounds.Beep.Play();
                        MessageBox.Show("Błąd podczas tworzenia pliku konfiguracyjnego API", "Błąd");
                    }
                    break;
                case "remove":
                    try
                    {
                        File.Delete(form.APIconfigFileFullPath);
                    }
                    catch
                    {
                        SystemSounds.Beep.Play();
                        MessageBox.Show("Błąd podczas usuwania pliku konfiguracyjnego API", "Błąd");
                    }
                    break;
                case "read":
                    try
                    {
                        XmlDocument xmlDocument = new XmlDocument();
                        xmlDocument.Load(form.APIconfigFileFullPath);
                        XmlNodeList elementsByTagName = xmlDocument.GetElementsByTagName("Value");
                        foreach (XmlElement item in elementsByTagName)
                        {
                            textBox2.Text = item.GetAttribute("URI");
                            textBox4.Text = item.GetAttribute("country");
                            textBox3.Text = item.InnerText;
                        }
                    }
                    catch
                    {
                        SystemSounds.Beep.Play();
                        MessageBox.Show("Błąd podczas odczytu pliku konfiguracyjnego API", "Błąd");
                    }
                    break;
            }
        }
        void ChckAPIConfig()
        {
            Form1 form = Application.OpenForms.OfType<Form1>().FirstOrDefault();
            if (File.Exists(form.APIconfigFileFullPath))
            {
                checkBox1.Checked = true;
                CreateAPICOnfig("read");
            }
            else
            {
                checkBox1.Checked = false;
                DisableAPI(true);
            }
        }
        void DisableAPI(bool _state)
        {
            if (_state)
            {
                button3.Enabled = false;
                textBox2.Enabled = false;
                textBox3.Enabled = false;
                textBox4.Enabled = false;
            }
            else
            {
                button3.Enabled = true;
                textBox2.Enabled = true;
                textBox3.Enabled = true;
                textBox4.Enabled = true;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Form1 form = Application.OpenForms.OfType<Form1>().FirstOrDefault();
            CreateAPICOnfig("create");
            form.CheckWebReq();
        }
    }
}
