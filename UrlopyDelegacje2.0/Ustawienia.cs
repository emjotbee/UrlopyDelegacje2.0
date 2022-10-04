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
        private GroupBox groupBox3;
        private Button button6;
        private Button button5;
        private Button button4;
        private TextBox textBox6;
        private Label label2;
        private TextBox textBox5;
        private Label label1;
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
            FillKonfiguracja();
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
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.button6 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.textBox6 = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox5 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.textBox1);
            this.groupBox1.Location = new System.Drawing.Point(14, 14);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.groupBox1.Size = new System.Drawing.Size(354, 60);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Ilość dni urlopowych";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(7, 22);
            this.textBox1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(339, 23);
            this.textBox1.TabIndex = 0;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(14, 358);
            this.button1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(168, 78);
            this.button1.TabIndex = 1;
            this.button1.Text = "Zapisz";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.Button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(200, 358);
            this.button2.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(168, 78);
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
            this.groupBox2.Location = new System.Drawing.Point(14, 81);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.groupBox2.Size = new System.Drawing.Size(354, 143);
            this.groupBox2.TabIndex = 3;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Ustawienia API";
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(200, 17);
            this.button3.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(147, 27);
            this.button3.TabIndex = 5;
            this.button3.Text = "Sprawdź połączenie";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(7, 22);
            this.checkBox1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(57, 19);
            this.checkBox1.TabIndex = 4;
            this.checkBox1.Text = "Włącz";
            this.checkBox1.UseVisualStyleBackColor = true;
            this.checkBox1.Click += new System.EventHandler(this.checkBox1_Click);
            // 
            // textBox4
            // 
            this.textBox4.Location = new System.Drawing.Point(7, 108);
            this.textBox4.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.textBox4.Name = "textBox4";
            this.textBox4.Size = new System.Drawing.Size(339, 23);
            this.textBox4.TabIndex = 2;
            this.textBox4.Text = "country";
            // 
            // textBox3
            // 
            this.textBox3.Location = new System.Drawing.Point(7, 78);
            this.textBox3.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(339, 23);
            this.textBox3.TabIndex = 1;
            this.textBox3.Text = "api_key";
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(7, 48);
            this.textBox2.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(339, 23);
            this.textBox2.TabIndex = 0;
            this.textBox2.Text = "URI";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.button6);
            this.groupBox3.Controls.Add(this.button5);
            this.groupBox3.Controls.Add(this.button4);
            this.groupBox3.Controls.Add(this.textBox6);
            this.groupBox3.Controls.Add(this.label2);
            this.groupBox3.Controls.Add(this.textBox5);
            this.groupBox3.Controls.Add(this.label1);
            this.groupBox3.Location = new System.Drawing.Point(14, 230);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(354, 122);
            this.groupBox3.TabIndex = 4;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Konfiguracja";
            // 
            // button6
            // 
            this.button6.Location = new System.Drawing.Point(7, 89);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(339, 23);
            this.button6.TabIndex = 7;
            this.button6.Text = "Otwórz folder roboczy";
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new System.EventHandler(this.button6_Click);
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(288, 53);
            this.button5.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(59, 23);
            this.button5.TabIndex = 6;
            this.button5.Text = "Dodaj";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(288, 20);
            this.button4.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(59, 23);
            this.button4.TabIndex = 5;
            this.button4.Text = "Dodaj";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // textBox6
            // 
            this.textBox6.Enabled = false;
            this.textBox6.Location = new System.Drawing.Point(119, 53);
            this.textBox6.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.textBox6.Name = "textBox6";
            this.textBox6.Size = new System.Drawing.Size(161, 23);
            this.textBox6.TabIndex = 3;
            this.textBox6.Text = "brak";
            this.textBox6.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(7, 53);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(31, 15);
            this.label2.TabIndex = 2;
            this.label2.Text = "PWS";
            // 
            // textBox5
            // 
            this.textBox5.Enabled = false;
            this.textBox5.Location = new System.Drawing.Point(119, 20);
            this.textBox5.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.textBox5.Name = "textBox5";
            this.textBox5.Size = new System.Drawing.Size(161, 23);
            this.textBox5.TabIndex = 1;
            this.textBox5.Text = "brak";
            this.textBox5.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(7, 28);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(105, 15);
            this.label1.TabIndex = 0;
            this.label1.Text = "Wniosek urlopowy";
            // 
            // Ustawienia
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(382, 445);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.MaximizeBox = false;
            this.Name = "Ustawienia";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Ustawienia";
            this.Load += new System.EventHandler(this.Ustawienia_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
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

        void FillKonfiguracja()
        {          
            string sourceFileNametmp = Path.Combine(Form1.dataPath, "template.docx");
            if (File.Exists(sourceFileNametmp))
            {
                textBox5.Text = sourceFileNametmp;
                button4.Text = "Otwórz";
            }
            string sourceFileNamepws = Path.Combine(Form1.dataPath, "pws_template.doc");
            if (File.Exists(sourceFileNamepws))
            {
                textBox6.Text = sourceFileNamepws;
                button5.Text = "Otwórz";
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Details details = new Details();
            if (button4.Text == "Otwórz")
            {
                details.OpenWniosek(Path.Combine(Form1.dataPath, "template.docx"));
            }
            else
            {
                details.LoadTemplate("template.docx");
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Details details = new Details();
            if (button4.Text == "Otwórz")
            {
                details.OpenWniosek(Path.Combine(Form1.dataPath, "pws_template.doc"));
            }
            else
            {
                details.LoadTemplate("pws_template.doc");
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("explorer.exe",@$"{Form1.dataPath}");
        }
    }
}
