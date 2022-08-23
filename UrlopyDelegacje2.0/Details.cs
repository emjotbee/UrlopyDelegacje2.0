using System;
using System.ComponentModel;
using System.Collections;
using System.Drawing;
using System.Linq;
using System.Media;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using System.IO;
using RestSharp;
using System.Xml;
using System.Xml.Serialization;
using System.Text;

namespace UrlopyDelegacje
{
	public class Details : Form
	{
		private Form principalForm = System.Windows.Forms.Application.OpenForms.OfType<Form1>().FirstOrDefault();

		private Urlop urlop;

		private IContainer components = null;

		private GroupBox groupBox1;

		private Label label1;

		private TextBox Do;

		private TextBox Od;

		private Label label2;

		private TextBox IloscDni;

		private Label label4;

		private TextBox Komentarz;

		private Label label3;

		private Button button3;

		private Button button2;

		private Button button1;

		private Label label5;

		private TextBox ID;

		private DateTimePicker dateTimePickerDo;
        private TextBox textBox1;
        private Label label6;
        private DataGridView dataGridView1;
        private Button button4;
        private DataGridViewTextBoxColumn WyjazdZ;
        private DataGridViewTextBoxColumn Data;
        private DataGridViewTextBoxColumn PrzyjazdDo;
        private DataGridViewTextBoxColumn DataPrz;
        private DataGridViewTextBoxColumn SrodkiLokomocji;
        private DataGridViewTextBoxColumn Koszt;
        private OpenFileDialog openFileDialog1;
        private SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.CheckBox sniadania;
        private DateTimePicker dateTimePickerOd;

		public Details()
		{
			InitializeComponent();
		}

		public void Details_Load(object sender, EventArgs e)
		{
			Form1 form = System.Windows.Forms.Application.OpenForms.OfType<Form1>().FirstOrDefault();
			urlop = form.GetUrlop();
			Text = form.GetUrlop().Od.ToString("dd/MM/yyyy") + " - " + form.GetUrlop().Do.ToString("dd/MM/yyyy");
			FillForm(form.GetUrlop());
			SetPickersAndButtons();
            EnableDelegacja();
        }

		private void Details_FormClosing(object sender, FormClosingEventArgs e)
		{
			principalForm.Enabled = true;
		}

		private void FillForm(Urlop _url)
		{
			ID.Text = _url.ID.ToString();
			Od.Text = _url.Od.ToString("dd/MM/yyyy");
			Do.Text = _url.Do.ToString("dd/MM/yyyy");
			Komentarz.Text = _url.Comments;
			IloscDni.Text = _url.DniIlosc.ToString();
            if (_url.Delegacja)
            {
                textBox1.Text = _url.Zwrot.ToString();
            }
            else
            {
                textBox1.Text = _url.Swieto;
            }

        }

		private void Button2_Click(object sender, EventArgs e)
		{
			Form1 form = System.Windows.Forms.Application.OpenForms.OfType<Form1>().FirstOrDefault();
			FillForm(form.GetUrlop());
		}

		private void DateTimePickerOd_ValueChanged(object sender, EventArgs e)
		{
			Od.Text = dateTimePickerOd.Value.ToString("dd/MM/yyyy");
		}

		private void DateTimePickerDo_ValueChanged(object sender, EventArgs e)
		{
			Do.Text = dateTimePickerDo.Value.ToString("dd/MM/yyyy");
		}

		private void SetPickersAndButtons()
		{
			dateTimePickerOd.Value = urlop.Od;
			dateTimePickerDo.Value = urlop.Do;
			if (!string.IsNullOrEmpty(urlop.WniosekPath))
			{
				button3.Enabled = true;
			}
		}

		private void Button1_Click(object sender, EventArgs e)
		{
            Cursor.Current = Cursors.WaitCursor;
            int num = Form1.listaUrlopow.Urlopy.Count();
			Form1 form = System.Windows.Forms.Application.OpenForms.OfType<Form1>().FirstOrDefault();
			form.Button2_Click(sender, e);
			form.SetPickersForNewUrlop(dateTimePickerOd.Value, dateTimePickerDo.Value);
			form.CreateNewUrlop(form.DniWyk(),this.urlop.Delegacja);
			if (Form1.listaUrlopow.Urlopy.Count() != num)
			{
				Form1.listaUrlopow.Urlopy.Add(this.urlop);
            }
            else
            {
				Urlop urlop = Form1.listaUrlopow.Urlopy.Find((Urlop ele) => ele.Od == dateTimePickerOd.Value);
				urlop.Comments = Komentarz.Text;
				urlop.WniosekPath = this.urlop.WniosekPath;
			}
			form.SaveToXml();
			form.FillList();
			form.FillForm();
            Cursor.Current = Cursors.Default;
            Close();
		}

		private void Komentarz_TextChanged(object sender, EventArgs e)
		{
			button1.Enabled = true;
		}

		private void DateTimePickerOd_Enter(object sender, EventArgs e)
		{
			button1.Enabled = true;
		}

		private void DateTimePickerDo_Enter(object sender, EventArgs e)
		{
			button1.Enabled = true;
		}

		private void OpenWniosek(string _path)
		{
			Microsoft.Office.Interop.Word.Application obj = (Microsoft.Office.Interop.Word.Application)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("000209FF-0000-0000-C000-000000000046")));
			obj.Visible = true;
			Microsoft.Office.Interop.Word.Application application = obj;
			Documents documents = application.Documents;
			object FileName = _path;
			object ConfirmConversions = Type.Missing;
			object ReadOnly = false;
			object AddToRecentFiles = Type.Missing;
			object PasswordDocument = Type.Missing;
			object PasswordTemplate = Type.Missing;
			object Revert = Type.Missing;
			object WritePasswordDocument = Type.Missing;
			object WritePasswordTemplate = Type.Missing;
			object Format = Type.Missing;
			object Encoding = Type.Missing;
			object Visible = true;
			object OpenAndRepair = Type.Missing;
			object DocumentDirection = Type.Missing;
			object NoEncodingDialog = Type.Missing;
			object XMLTransform = Type.Missing;
			Document document = documents.Open(ref FileName, ref ConfirmConversions, ref ReadOnly, ref AddToRecentFiles, ref PasswordDocument, ref PasswordTemplate, ref Revert, ref WritePasswordDocument, ref WritePasswordTemplate, ref Format, ref Encoding, ref Visible, ref OpenAndRepair, ref DocumentDirection, ref NoEncodingDialog, ref XMLTransform);
		}

		private void Button3_Click(object sender, EventArgs e)
		{
			try
			{
				OpenWniosek(urlop.WniosekPath);
			}
			catch
			{
				SystemSounds.Beep.Play();
				MessageBox.Show("Nie można otworzyć pliku", "Błąd");
			}
		}

		protected override void Dispose(bool disposing)
		{
			if (disposing && components != null)
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

        void EnableDelegacja()
        {
            Form1 form = System.Windows.Forms.Application.OpenForms.OfType<Form1>().FirstOrDefault();
            if ((form.GetUrlop()).Delegacja)
            {
                dataGridView1.Enabled = true;
                dataGridView1.DefaultCellStyle.BackColor = SystemColors.Window;
                dataGridView1.DefaultCellStyle.ForeColor = SystemColors.ControlText;
                dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = SystemColors.Window;
                dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = SystemColors.ControlText;
                dataGridView1.ReadOnly = false;
                dataGridView1.EnableHeadersVisualStyles = true;
                button4.Enabled = true;
                label3.Text = "Miejsce";
                label6.Text = "Zwrot";
            }
            else
            {
                dataGridView1.DefaultCellStyle.BackColor = SystemColors.Control;
                dataGridView1.DefaultCellStyle.ForeColor = SystemColors.GrayText;
                dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = SystemColors.Control;
                dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = SystemColors.GrayText;
                dataGridView1.ReadOnly = true;
                dataGridView1.EnableHeadersVisualStyles = false;
                button4.Enabled = false;
                checkBox1.Enabled = false;
                sniadania.Enabled = false;
            }
        }

        private void InitializeComponent()
		{
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Details));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.button4 = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.WyjazdZ = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Data = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PrzyjazdDo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DataPrz = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.SrodkiLokomocji = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Koszt = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.dateTimePickerDo = new System.Windows.Forms.DateTimePicker();
            this.dateTimePickerOd = new System.Windows.Forms.DateTimePicker();
            this.button3 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.ID = new System.Windows.Forms.TextBox();
            this.IloscDni = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.Komentarz = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.Do = new System.Windows.Forms.TextBox();
            this.Od = new System.Windows.Forms.TextBox();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.sniadania = new System.Windows.Forms.CheckBox();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.sniadania);
            this.groupBox1.Controls.Add(this.checkBox1);
            this.groupBox1.Controls.Add(this.button4);
            this.groupBox1.Controls.Add(this.dataGridView1);
            this.groupBox1.Controls.Add(this.textBox1);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.dateTimePickerDo);
            this.groupBox1.Controls.Add(this.dateTimePickerOd);
            this.groupBox1.Controls.Add(this.button3);
            this.groupBox1.Controls.Add(this.button2);
            this.groupBox1.Controls.Add(this.button1);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.ID);
            this.groupBox1.Controls.Add(this.IloscDni);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.Komentarz);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.Do);
            this.groupBox1.Controls.Add(this.Od);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(461, 506);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Szczegóły";
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(6, 399);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(120, 17);
            this.checkBox1.TabIndex = 20;
            this.checkBox1.Text = "Wyjazd zagraniczny";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(211, 393);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(243, 23);
            this.button4.TabIndex = 19;
            this.button4.Text = "Generuj polecenie wyjzadu służbowego";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.WyjazdZ,
            this.Data,
            this.PrzyjazdDo,
            this.DataPrz,
            this.SrodkiLokomocji,
            this.Koszt});
            this.dataGridView1.Enabled = false;
            this.dataGridView1.Location = new System.Drawing.Point(5, 178);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(450, 211);
            this.dataGridView1.TabIndex = 17;
            // 
            // WyjazdZ
            // 
            this.WyjazdZ.HeaderText = "Wyjazd z";
            this.WyjazdZ.Name = "WyjazdZ";
            // 
            // Data
            // 
            dataGridViewCellStyle1.Format = "g";
            dataGridViewCellStyle1.NullValue = null;
            this.Data.DefaultCellStyle = dataGridViewCellStyle1;
            this.Data.HeaderText = "Data i godzina";
            this.Data.Name = "Data";
            // 
            // PrzyjazdDo
            // 
            this.PrzyjazdDo.HeaderText = "Przyjazd do";
            this.PrzyjazdDo.Name = "PrzyjazdDo";
            // 
            // DataPrz
            // 
            dataGridViewCellStyle2.Format = "g";
            dataGridViewCellStyle2.NullValue = null;
            this.DataPrz.DefaultCellStyle = dataGridViewCellStyle2;
            this.DataPrz.HeaderText = "Data i godzina";
            this.DataPrz.Name = "DataPrz";
            // 
            // SrodkiLokomocji
            // 
            this.SrodkiLokomocji.HeaderText = "Środki lokomocji";
            this.SrodkiLokomocji.Name = "SrodkiLokomocji";
            // 
            // Koszt
            // 
            dataGridViewCellStyle3.Format = "N2";
            dataGridViewCellStyle3.NullValue = "0";
            this.Koszt.DefaultCellStyle = dataGridViewCellStyle3;
            this.Koszt.HeaderText = "Koszt";
            this.Koszt.Name = "Koszt";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(76, 152);
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.Size = new System.Drawing.Size(379, 20);
            this.textBox1.TabIndex = 16;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(11, 159);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(53, 13);
            this.label6.TabIndex = 15;
            this.label6.Text = "Dni świąt.";
            // 
            // dateTimePickerDo
            // 
            this.dateTimePickerDo.Location = new System.Drawing.Point(439, 74);
            this.dateTimePickerDo.Name = "dateTimePickerDo";
            this.dateTimePickerDo.Size = new System.Drawing.Size(16, 20);
            this.dateTimePickerDo.TabIndex = 14;
            this.dateTimePickerDo.ValueChanged += new System.EventHandler(this.DateTimePickerDo_ValueChanged);
            this.dateTimePickerDo.Enter += new System.EventHandler(this.DateTimePickerDo_Enter);
            // 
            // dateTimePickerOd
            // 
            this.dateTimePickerOd.Location = new System.Drawing.Point(439, 49);
            this.dateTimePickerOd.Name = "dateTimePickerOd";
            this.dateTimePickerOd.Size = new System.Drawing.Size(16, 20);
            this.dateTimePickerOd.TabIndex = 13;
            this.dateTimePickerOd.ValueChanged += new System.EventHandler(this.DateTimePickerOd_ValueChanged);
            this.dateTimePickerOd.Enter += new System.EventHandler(this.DateTimePickerOd_Enter);
            // 
            // button3
            // 
            this.button3.Enabled = false;
            this.button3.Location = new System.Drawing.Point(316, 422);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(139, 77);
            this.button3.TabIndex = 12;
            this.button3.Text = "Otwórz wniosek";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.Button3_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(150, 422);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(160, 77);
            this.button2.TabIndex = 11;
            this.button2.Text = "Załaduj ponownie";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.Button2_Click);
            // 
            // button1
            // 
            this.button1.Enabled = false;
            this.button1.Location = new System.Drawing.Point(5, 422);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(139, 77);
            this.button1.TabIndex = 10;
            this.button1.Text = "Zapisz";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.Button1_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(11, 29);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(18, 13);
            this.label5.TabIndex = 9;
            this.label5.Text = "ID";
            // 
            // ID
            // 
            this.ID.Location = new System.Drawing.Point(77, 22);
            this.ID.Name = "ID";
            this.ID.ReadOnly = true;
            this.ID.Size = new System.Drawing.Size(378, 20);
            this.ID.TabIndex = 8;
            // 
            // IloscDni
            // 
            this.IloscDni.Location = new System.Drawing.Point(77, 126);
            this.IloscDni.Name = "IloscDni";
            this.IloscDni.ReadOnly = true;
            this.IloscDni.Size = new System.Drawing.Size(378, 20);
            this.IloscDni.TabIndex = 7;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(11, 133);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(46, 13);
            this.label4.TabIndex = 6;
            this.label4.Text = "Ilość dni";
            // 
            // Komentarz
            // 
            this.Komentarz.Location = new System.Drawing.Point(77, 100);
            this.Komentarz.Name = "Komentarz";
            this.Komentarz.Size = new System.Drawing.Size(378, 20);
            this.Komentarz.TabIndex = 5;
            this.Komentarz.TextChanged += new System.EventHandler(this.Komentarz_TextChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(11, 107);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(57, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "Komentarz";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(11, 81);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(21, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Do";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(11, 55);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(21, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Od";
            // 
            // Do
            // 
            this.Do.Location = new System.Drawing.Point(77, 74);
            this.Do.Name = "Do";
            this.Do.ReadOnly = true;
            this.Do.Size = new System.Drawing.Size(356, 20);
            this.Do.TabIndex = 1;
            // 
            // Od
            // 
            this.Od.Location = new System.Drawing.Point(77, 48);
            this.Od.Name = "Od";
            this.Od.ReadOnly = true;
            this.Od.Size = new System.Drawing.Size(356, 20);
            this.Od.TabIndex = 0;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // sniadania
            // 
            this.sniadania.AutoSize = true;
            this.sniadania.Location = new System.Drawing.Point(132, 399);
            this.sniadania.Name = "sniadania";
            this.sniadania.Size = new System.Drawing.Size(73, 17);
            this.sniadania.TabIndex = 21;
            this.sniadania.Text = "Śniadania";
            this.sniadania.UseVisualStyleBackColor = true;
            // 
            // Details
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(485, 523);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Details";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Details";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Details_FormClosing);
            this.Load += new System.EventHandler(this.Details_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

		}

        private void button4_Click(object sender, EventArgs e)
        {
            string path = Path.Combine(Form1.dataPath, "pws_template.doc");
            if (!File.Exists(path))
            {
                LoadTemplate();
            }
            else
            {
                SaveWniosek();
            }
        }
        private void LoadTemplate()
        {         
            string destFileName = Path.Combine(Form1.dataPath, "pws_template.doc");
            SystemSounds.Beep.Play();
            MessageBox.Show("Załaduj szablon polecenia wyjazdu służbowego", "Błąd");
            openFileDialog1.FileName = "pws_szablon.doc";
            openFileDialog1.Title = "Załaduj szablon polecenia wyjazdu służbowego";
            openFileDialog1.Filter = "Text files (*.doc)|*.doc|All files (*.*)|*.*";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    File.Copy(openFileDialog1.FileName, destFileName);
                    SystemSounds.Hand.Play();
                    MessageBox.Show("Szablon zapisany", "Sukces");
                }
                catch
                {
                    SystemSounds.Beep.Play();
                    MessageBox.Show("Błąd podczas zapisu", "Błąd");
                }
            }
        }

        private void SaveWniosek()
        {
            Form1 form = System.Windows.Forms.Application.OpenForms.OfType<Form1>().FirstOrDefault();
            string sourceFileName = Path.Combine(Form1.dataPath, "pws_template.doc");
            saveFileDialog1.FileName = Od.Text.Replace("/", "-") + "_" + Do.Text.Replace("/", "-") + "_" + Komentarz.Text;
            saveFileDialog1.Title = "Zapisz polecenie wyjazdu służbowego";
            saveFileDialog1.Filter = "Text files (*.doc)|*.doc|All files (*.*)|*.*";
            if (saveFileDialog1.ShowDialog() != DialogResult.OK)
            {
                return;
            }
            string lokom = "";
            int sumadoj = 0;
            double diety = 0;
            File.Copy(sourceFileName, saveFileDialog1.FileName, overwrite: true);
            object fileName = saveFileDialog1.FileName;
            Microsoft.Office.Interop.Word.Application obj = (Microsoft.Office.Interop.Word.Application)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("000209FF-0000-0000-C000-000000000046")));
            obj.Visible = true;
            Microsoft.Office.Interop.Word.Application application = obj;
            Documents documents = application.Documents;
            object FileName = fileName;
            object ConfirmConversions = Type.Missing;
            object ReadOnly = false;
            object AddToRecentFiles = Type.Missing;
            object PasswordDocument = Type.Missing;
            object PasswordTemplate = Type.Missing;
            object Revert = Type.Missing;
            object WritePasswordDocument = Type.Missing;
            object WritePasswordTemplate = Type.Missing;
            object Format = Type.Missing;
            object Encoding = Type.Missing;
            object Visible = true;
            object OpenAndRepair = Type.Missing;
            object DocumentDirection = Type.Missing;
            object NoEncodingDialog = Type.Missing;
            object XMLTransform = Type.Missing;
            Document document = documents.Open(ref FileName, ref ConfirmConversions, ref ReadOnly, ref AddToRecentFiles, ref PasswordDocument, ref PasswordTemplate, ref Revert, ref WritePasswordDocument, ref WritePasswordTemplate, ref Format, ref Encoding, ref Visible, ref OpenAndRepair, ref DocumentDirection, ref NoEncodingDialog, ref XMLTransform);
            document.Activate();
            try
            {
                form.FindAndReplace(application, "<datanow>", DateTime.Today.ToString("dd/MM/yyyy"));
                form.FindAndReplace(application, "<dataod>", Od.Text);
                form.FindAndReplace(application, "<datado>", Do.Text);
                form.FindAndReplace(application, "<komentarz>", Komentarz.Text);
                for (int i = 0; i < dataGridView1.Rows.Count -1; i++)
                {
                    form.FindAndReplace(application, $"<wyjazdz{i}>", dataGridView1.Rows[i].Cells["WyjazdZ"].Value.ToString());
                    form.FindAndReplace(application, $"<data{i}>", dataGridView1.Rows[i].Cells["Data"].Value.ToString());
                    form.FindAndReplace(application, $"<przyjazddo{i}>", dataGridView1.Rows[i].Cells["PrzyjazdDo"].Value.ToString());
                    form.FindAndReplace(application, $"<dataprz{i}>", dataGridView1.Rows[i].Cells["DataPrz"].Value.ToString());
                    form.FindAndReplace(application, $"<lokom{i}>", dataGridView1.Rows[i].Cells["SrodkiLokomocji"].Value.ToString());
                    if (string.IsNullOrEmpty(lokom))
                    {
                        lokom += dataGridView1.Rows[i].Cells["SrodkiLokomocji"].Value.ToString();
                    }
                    else
                    {
                        if (!lokom.Contains(dataGridView1.Rows[i].Cells["SrodkiLokomocji"].Value.ToString()))
                        {
                            lokom += "," + dataGridView1.Rows[i].Cells["SrodkiLokomocji"].Value.ToString();
                        }
                    }
                    form.FindAndReplace(application, $"<koszt{i}>", dataGridView1.Rows[i].Cells["Koszt"].Value.ToString() + " zł");
                    sumadoj += Convert.ToInt32(dataGridView1.Rows[i].Cells["Koszt"].Value);
                }
                for (int i = 0; i < 4; i++)
                {
                    form.FindAndReplace(application, $"<wyjazdz{i}>", "");
                    form.FindAndReplace(application, $"<data{i}>", "");
                    form.FindAndReplace(application, $"<przyjazddo{i}>", "");
                    form.FindAndReplace(application, $"<dataprz{i}>", "");
                    form.FindAndReplace(application, $"<lokom{i}>", "");
                    form.FindAndReplace(application, $"<koszt{i}>", "");
                }
                form.FindAndReplace(application, "<lokom>", lokom);
                form.FindAndReplace(application, "<sumadoj>", sumadoj);
                if (checkBox1.Checked)
                {
                    diety = Math.Round((Convert.ToDouble(IloscDni.Text) * 49) * GetCurrency());
                }
                else
                {
                    diety = Convert.ToInt32(IloscDni.Text) * 30;
                }
                if (sniadania.Checked)
                {
                    diety = Math.Round(diety - (diety * 0.25));
                }
                form.FindAndReplace(application, "<diety>", diety);
                form.FindAndReplace(application, "<razem>", Convert.ToString(diety + sumadoj));
                SystemSounds.Hand.Play();
                MessageBox.Show("Wniosek wygenerowany pomyślnie. Zapisz plik.", "Sukces");
                try
                {
                    form.GetUrlop().WniosekPath = saveFileDialog1.FileName;
                    form.GetUrlop().Zwrot = diety + sumadoj;
                    form.SaveToXml();
                }
                catch
                {
                }
            }
            catch
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Błąd podczas zapisu", "Błąd");
                documents.Close();
                File.Delete(saveFileDialog1.FileName);                
            }
        }

        double GetCurrency()
        {
            IRestResponse response;
            var client = new RestClient("https://api.nbp.pl/api/exchangerates/rates/a/eur/");
            client.Timeout = -1;
            var request = new RestRequest(Method.GET);
            request.AddHeader("Accept", "text/xml");
            response = client.Execute(request);
            System.Threading.Thread.Sleep(1000);
            using (var ms = new MemoryStream(response.RawBytes))
            {
                XmlSerializer serializer = new XmlSerializer(typeof(ExchangeRatesSeries));
                ExchangeRatesSeries feedObject = (ExchangeRatesSeries)serializer.Deserialize(ms);
                return Convert.ToDouble(feedObject.Rates.Rate.Mid);
            }
        }

        // NOTE: Generated code may require at least .NET Framework 4.5 or .NET Core/Standard 2.0.
        /// <remarks/>
        [System.SerializableAttribute()]
        [System.ComponentModel.DesignerCategoryAttribute("code")]
        [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
        [System.Xml.Serialization.XmlRootAttribute(Namespace = "", IsNullable = false)]
        public partial class ExchangeRatesSeries
        {

            private string tableField;

            private string currencyField;

            private string codeField;

            private ExchangeRatesSeriesRates ratesField;

            /// <remarks/>
            public string Table
            {
                get
                {
                    return this.tableField;
                }
                set
                {
                    this.tableField = value;
                }
            }

            /// <remarks/>
            public string Currency
            {
                get
                {
                    return this.currencyField;
                }
                set
                {
                    this.currencyField = value;
                }
            }

            /// <remarks/>
            public string Code
            {
                get
                {
                    return this.codeField;
                }
                set
                {
                    this.codeField = value;
                }
            }

            /// <remarks/>
            public ExchangeRatesSeriesRates Rates
            {
                get
                {
                    return this.ratesField;
                }
                set
                {
                    this.ratesField = value;
                }
            }
        }

        /// <remarks/>
        [System.SerializableAttribute()]
        [System.ComponentModel.DesignerCategoryAttribute("code")]
        [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
        public partial class ExchangeRatesSeriesRates
        {

            private ExchangeRatesSeriesRatesRate rateField;

            /// <remarks/>
            public ExchangeRatesSeriesRatesRate Rate
            {
                get
                {
                    return this.rateField;
                }
                set
                {
                    this.rateField = value;
                }
            }
        }

        /// <remarks/>
        [System.SerializableAttribute()]
        [System.ComponentModel.DesignerCategoryAttribute("code")]
        [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
        public partial class ExchangeRatesSeriesRatesRate
        {

            private string noField;

            private System.DateTime effectiveDateField;

            private decimal midField;

            /// <remarks/>
            public string No
            {
                get
                {
                    return this.noField;
                }
                set
                {
                    this.noField = value;
                }
            }

            /// <remarks/>
            [System.Xml.Serialization.XmlElementAttribute(DataType = "date")]
            public System.DateTime EffectiveDate
            {
                get
                {
                    return this.effectiveDateField;
                }
                set
                {
                    this.effectiveDateField = value;
                }
            }

            /// <remarks/>
            public decimal Mid
            {
                get
                {
                    return this.midField;
                }
                set
                {
                    this.midField = value;
                }
            }
        }


    }
}
