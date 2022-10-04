using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Media;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Serialization;
using Microsoft.Office.Interop.Word;
using RestSharp;

namespace UrlopyDelegacje
{
	public class Form1 : Form
	{
		public static ListaUrlopow listaUrlopow = new ListaUrlopow();

		private static string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);

		public static string dataPath = Path.Combine(appDataPath, "UrlopyDelegacje") ;

		private string urlopyFileFullPath = Path.Combine(dataPath, DateTime.Now.Year + "Urlopy.xml");

		private string swietaFileFullPath = Path.Combine(dataPath, DateTime.Now.Year + "Swieta.xml");

		private string configFileFullPath = Path.Combine(dataPath, "Config.xml");

		public string APIconfigFileFullPath = Path.Combine(dataPath, "APIConfig.xml");

		public static int dniurlwsz;

		private string version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();

		private IContainer components = null;

		private DataGridView dataGridView1;

		private BackgroundWorker backgroundWorker1;

		private GroupBox groupBox1;

		private MonthCalendar monthCalendar1;

		private GroupBox groupBox2;

		private DateTimePicker PickerDo;

		private DateTimePicker PickerOd;

		private Label label5;

		private Label label4;

		private Label label3;

		private Label label2;

		private Label label1;

		private TextBox textBox3;

		private TextBox textBox2;

		private TextBox textBox1;

		private Button button2;

		private Button button1;

		private GroupBox groupBox3;

		private GroupBox groupBox4;

		private Button button3;

		private OpenFileDialog openFileDialog1;

		private SaveFileDialog saveFileDialog1;

		private MenuStrip menuStrip1;

		private ToolStripMenuItem dasdToolStripMenuItem;

		private ToolStripMenuItem konfiguracjaToolStripMenuItem;

		private ToolStripMenuItem ustawieniaToolStripMenuItem;

		private ToolStripMenuItem wersjaToolStripMenuItem;

		private ToolStripComboBox toolStripComboBox1;
        private Label label6;
        private GroupBox groupBox5;
        private Label label8;
        private DateTimePicker dateTimePicker1;
        private Button button5;
        private Label label7;
        private GroupBox groupBox6;
        private Label label10;
        private Label label9;
        private Button button6;
        private System.Windows.Forms.CheckBox checkBox1;
        private GroupBox groupBox7;
        private TextBox textBox6;
        private TextBox textBox5;
        private TextBox textBox4;
        private Button button4;

		public Form1()
		{
			InitializeComponent();
			Text = "UrlopyDelegacje " + version + " ©2022 Autor: Kamil Kłonica";
		}

		private void Form1_Load(object sender, EventArgs e)
		{
			CheckWebReq();
			CreateConfigXML(26);
			CreateXML(listaUrlopow);
			FillList();
			FillForm();
			FillCalendar();
			toolStripComboBox1.DropDownStyle = ComboBoxStyle.DropDownList;
			toolStripComboBox1.Items.Add(DateTime.Now.Year + "Urlopy.xml");
			toolStripComboBox1.Text = DateTime.Now.Year + "Urlopy.xml";
			GetKraj();
		}

		private void Button1_Click(object sender, EventArgs e)
		{
			Cursor.Current = Cursors.WaitCursor;
			CreateNewUrlop(DniWyk(),checkBox1.Checked);
			FillList();
			FillForm();
			Cursor.Current = Cursors.Default;
		}

		private void CreateXML(ListaUrlopow lista)
		{
			if (File.Exists(urlopyFileFullPath))
			{
				XmlSerializer xmlSerializer = new XmlSerializer(lista.GetType());
				using (Stream stream = File.OpenRead(urlopyFileFullPath))
				{
					lista = (ListaUrlopow)xmlSerializer.Deserialize(stream);
				}
				listaUrlopow = lista;
			}
			else
			{
				Directory.CreateDirectory(dataPath);
				SaveToXml();
			}
		}

		public void CreateNewUrlop(int dniwyk,bool isDelegacja)
		{
			Urlop urlop = new Urlop();
			if (dniwyk >= dniurlwsz)
			{
				SystemSounds.Beep.Play();
				MessageBox.Show("Wszystkie dni urlopowe wykorzystane!", "Błąd");
				return;
			}
			if ((PickerDo.Value.Date - PickerOd.Value.Date).Days - urlop.DniWeekend(PickerOd.Value.Date, PickerDo.Value.Date) + 1 > Convert.ToInt32(textBox3.Text))
			{
				SystemSounds.Beep.Play();
				MessageBox.Show("Brak wystarczającej ilości dni!", "Błąd");
				return;
			}
			if (!CheckIfPossible())
			{
				SystemSounds.Beep.Play();
				MessageBox.Show("Tworzony urlop pokrywa się z już istniejącym!", "Błąd");
				return;
			}
			try
			{
				Urlop item = new Urlop(PickerOd.Value.Date, PickerDo.Value.Date, isDelegacja);
				listaUrlopow.Urlopy.Add(item);
				SaveToXml();
				monthCalendar1.SetSelectionRange(PickerOd.Value.Date, PickerDo.Value.Date);
				SetPickers("clear");
			}
			catch
			{
				SystemSounds.Beep.Play();
				MessageBox.Show("Błędna data", "Błąd");
			}
		}

		public void FillList()
		{
			dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
			dataGridView1.Rows.Clear();
			dataGridView1.ColumnCount = 6;
			dataGridView1.Columns[0].Name = "ID";
			dataGridView1.Columns[0].Visible = false;
			dataGridView1.Columns[1].Name = "Od";
			dataGridView1.Columns[2].Name = "Do";
			dataGridView1.Columns[3].Name = "Ilość dni";
			dataGridView1.Columns[4].Name = "Data";
			dataGridView1.Columns[4].Visible = false;
			dataGridView1.Columns[5].Name = "Komentarz";
			foreach (Urlop item in listaUrlopow.Urlopy)
			{
				dataGridView1.Rows.Add(item.ID, item.Od.ToString("dd/MM/yyyy"), item.Do.ToString("dd/MM/yyyy"), item.DniIlosc, item.Od, item.Comments);
			}
			Sort();
			foreach (DataGridViewColumn column in dataGridView1.Columns)
			{
				column.SortMode = DataGridViewColumnSortMode.NotSortable;
			}
			foreach (DataGridViewRow item2 in (IEnumerable)dataGridView1.Rows)
			{
				if (Convert.ToDateTime(item2.Cells["Data"].Value) > DateTime.Now)
				{
					item2.DefaultCellStyle.BackColor = Color.Yellow;
				}
				if (!string.IsNullOrEmpty(GetUrlop(Convert.ToInt64(item2.Cells["ID"].Value)).Swieto))
                {
					item2.DefaultCellStyle.BackColor = Color.Green;
				}
				if ((GetUrlop(Convert.ToInt64(item2.Cells["ID"].Value)).Delegacja))
				{
					item2.DefaultCellStyle.BackColor = Color.LightBlue;
				}
			}
		}

		public void FillForm()
		{
			int num = dniurlwsz;
			textBox1.Text = num.ToString();
			textBox2.Text = DniWyk().ToString();
			textBox3.Text = (num - DniWyk()).ToString();
		}

		public int DniWyk()
		{
			int num = 0;
			foreach (Urlop item in listaUrlopow.Urlopy)
			{
				if(!item.Delegacja)
                {
					num += item.DniIlosc;
				}
			}
			return num;
		}

		public void Button2_Click(object sender, EventArgs e)
		{
			RemoveUrlop();
			FillList();
			FillForm();
			SaveToXml();
			button3.Enabled = false;
			button4.Enabled = false;
		}

		public void RemoveUrlop()
		{
			try
			{
				if (dataGridView1.CurrentCell != null)
				{
					int rowIndex = dataGridView1.CurrentCell.RowIndex;
					long num = Convert.ToInt64(dataGridView1.Rows[rowIndex].Cells[0].Value);
					foreach (Urlop item in listaUrlopow.Urlopy.ToList())
					{
						if (item.ID == num)
						{
							listaUrlopow.Urlopy.Remove(item);
						}
					}
				}
				else
				{
					throw new Exception();
				}				
			}
			catch
			{
				SystemSounds.Beep.Play();
				MessageBox.Show("Brak urlopu do usunięcia", "Błąd");
			}
		}

		public void SaveToXml()
		{
			XmlSerializer xmlSerializer = new XmlSerializer(typeof(ListaUrlopow));
			FileStream fileStream = File.Create(urlopyFileFullPath);
			xmlSerializer.Serialize(fileStream, listaUrlopow);
			fileStream.Close();
		}

		private void DataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
		{
			try
			{
				int rowIndex = dataGridView1.CurrentCell.RowIndex;
				long num = Convert.ToInt64(dataGridView1.Rows[rowIndex].Cells[0].Value);
				foreach (Urlop item in listaUrlopow.Urlopy.ToList())
				{
					if (item.ID == num)
					{
						DateTime od = item.Od;
						DateTime @do = item.Do;
						monthCalendar1.SetSelectionRange(od, @do);
                        if (item.Delegacja)
                        {
							button3.Enabled = false;
                        }
                        else
                        {
							button3.Enabled = true;
						}
					}
				}
				button4.Enabled = true;
			}
			catch
			{
				button3.Enabled = false;
				button4.Enabled = false;
			}
		}

		private void SetPickers(string value)
		{
			if (value == "clear")
			{
				PickerOd.Value = DateTime.Today;
				PickerDo.Value = DateTime.Today;
			}
			else if (value == "set")
			{
				PickerOd.MinDate = DateTime.Today;
				PickerDo.MinDate = DateTime.Today;
			}
		}

		private void PickerOd_Leave(object sender, EventArgs e)
		{
			PickerDo.Value = PickerOd.Value;
		}

		private void Button3_Click(object sender, EventArgs e)
		{
			string path = Path.Combine(dataPath, "template.docx");
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
			string destFileName = Path.Combine(dataPath, "template.docx");
			SystemSounds.Beep.Play();
			MessageBox.Show("Załaduj szablon wniosku urlopowego", "Błąd");
			openFileDialog1.FileName = "szablon.docx";
			openFileDialog1.Title = "Załaduj szablon wniosku urlopowego";
			openFileDialog1.Filter = "Text files (*.docx)|*.docx|All files (*.*)|*.*";
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
			int rowIndex = dataGridView1.CurrentCell.RowIndex;
			string sourceFileName = Path.Combine(dataPath, "template.docx");
			string text = dataGridView1.Rows[rowIndex].Cells[1].Value.ToString();
			string text2 = dataGridView1.Rows[rowIndex].Cells[2].Value.ToString();
			string replaceWithText = dataGridView1.Rows[rowIndex].Cells[3].Value.ToString();
			saveFileDialog1.FileName = text.Replace("/", "-") + "_" + text2.Replace("/", "-");
			saveFileDialog1.Title = "Zapisz wniosek urlopowy";
			saveFileDialog1.Filter = "Text files (*.docx)|*.docx|All files (*.*)|*.*";
			if (saveFileDialog1.ShowDialog() != DialogResult.OK)
			{
				return;
			}
			try
			{
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
				FindAndReplace(application, "<roknow>", DateTime.Today.Year);
				FindAndReplace(application, "<rokprv>", DateTime.Today.Year - 1);
				FindAndReplace(application, "<datenow>", DateTime.Today.ToString("dd/MM/yyyy"));
				FindAndReplace(application, "<dateod>", text);
				FindAndReplace(application, "<datedo>", text2);
				FindAndReplace(application, "<days>", replaceWithText);
				SystemSounds.Hand.Play();
				MessageBox.Show("Wniosek wygenerowany pomyślnie. Zapisz plik.", "Sukces");
				try
				{
					GetUrlop().WniosekPath = saveFileDialog1.FileName;
					SaveToXml();
				}
				catch
				{
				}
			}
			catch
			{
				SystemSounds.Beep.Play();
				MessageBox.Show("Błąd podczas zapisu", "Błąd");
			}
		}

		public List<DateTime> FillSwieta()
		{
			List<DateTime> list = new List<DateTime>();
			try
			{
				list.Add(new DateTime(DateTime.Now.Year, 1, 1));
				list.Add(new DateTime(DateTime.Now.Year, 4, 10));
				list.Add(new DateTime(DateTime.Now.Year, 4, 13));
				list.Add(new DateTime(DateTime.Now.Year, 5, 1));
				list.Add(new DateTime(DateTime.Now.Year, 10, 3));
				list.Add(new DateTime(DateTime.Now.Year, 12, 25));
				list.Add(new DateTime(DateTime.Now.Year, 12, 26));
			}
			catch
			{
				SystemSounds.Beep.Play();
				MessageBox.Show("Błąd podczas wyznaczania dni wolnych od pracy", "Błąd");
			}
			return list;
		}

		private void FillCalendar()
		{
			monthCalendar1.AnnuallyBoldedDates = FillSwieta().ToArray();
		}

		private bool CheckIfPossible()
		{
			int num = 0;
			foreach (Urlop item in listaUrlopow.Urlopy)
			{
				DateTime dateTime = PickerOd.Value.Date;
				while (dateTime.Date <= PickerDo.Value.Date)
				{
					DateTime dateTime2 = item.Do.Date;
					while (dateTime2.Date <= item.Od.Date)
					{
						if (dateTime.Date == dateTime2.Date)
						{
							num++;
						}
						dateTime2 = dateTime2.AddDays(1.0);
					}
					dateTime = dateTime.AddDays(1.0);
				}
			}
			if (num > 0)
			{
				return false;
			}
			return true;
		}

		public void FindAndReplace(Microsoft.Office.Interop.Word.Application doc, object findText, object replaceWithText)
		{	
			object MatchCase = false;
			object MatchWholeWord = true;
			object MatchWildcards = false;
			object MatchSoundsLike = false;
			object MatchAllWordForms = false;
			object Forward = true;
			object Format = false;
			object MatchKashida = false;
			object MatchDiacritics = false;
			object MatchAlefHamza = false;
			object MatchControl = false;
			object obj = false;
			object obj2 = true;
			object Replace = 2;
			object Wrap = 1;
			doc.Selection.Find.Execute(ref findText, ref MatchCase, ref MatchWholeWord, ref MatchWildcards, ref MatchSoundsLike, ref MatchAllWordForms, ref Forward, ref Wrap, ref Format, ref replaceWithText, ref Replace, ref MatchKashida, ref MatchDiacritics, ref MatchAlefHamza, ref MatchControl);
		}

		private void Sort()
		{
			dataGridView1.Sort(dataGridView1.Columns["Data"], ListSortDirection.Ascending);
		}

		public void CreateConfigXML(int dni)
		{
			try
			{
				if (File.Exists(configFileFullPath))
				{
					XmlDocument xmlDocument = new XmlDocument();
					xmlDocument.Load(configFileFullPath);
					XmlNodeList elementsByTagName = xmlDocument.GetElementsByTagName("Value");
					foreach (XmlElement item in elementsByTagName)
					{
						if (item.GetAttribute("Year") == DateTime.Now.Year.ToString())
						{
							dniurlwsz = Convert.ToInt32(item.InnerText);
						}
					}
					if (dniurlwsz == 0)
					{
						XmlElement xmlElement2 = xmlDocument.CreateElement("Value");
						XmlAttribute xmlAttribute = xmlDocument.CreateAttribute("Year");
						xmlAttribute.Value = DateTime.Now.Year.ToString();
						xmlElement2.Attributes.Append(xmlAttribute);
						xmlElement2.InnerText = "26";
						xmlDocument.DocumentElement.AppendChild(xmlElement2);
						xmlDocument.PreserveWhitespace = true;
						xmlDocument.Save(configFileFullPath);
					}
				}
				else
				{
					dniurlwsz = dni;
					Directory.CreateDirectory(dataPath);
					XmlDocument xmlDocument2 = new XmlDocument();
					xmlDocument2.LoadXml("<dniurlwsz></dniurlwsz>");
					XmlElement xmlElement3 = xmlDocument2.CreateElement("Value");
					XmlAttribute xmlAttribute2 = xmlDocument2.CreateAttribute("Year");
					xmlAttribute2.Value = DateTime.Now.Year.ToString();
					xmlElement3.Attributes.Append(xmlAttribute2);
					xmlElement3.InnerText = "26";
					xmlDocument2.DocumentElement.AppendChild(xmlElement3);
					xmlDocument2.PreserveWhitespace = true;
					xmlDocument2.Save(configFileFullPath);
				}
			}
			catch
			{
				SystemSounds.Beep.Play();
				MessageBox.Show("Błąd podczas tworzenia pliku konfiguracyjnego", "Błąd");
			}
		}

		private void UstawieniaToolStripMenuItem_Click(object sender, EventArgs e)
		{
			Ustawienia ustawienia = new Ustawienia();
			ustawienia.Show();
		}

		public void WrtieConfigXML(int dni)
		{
			try
			{
				dniurlwsz = dni;
				bool flag = false;
				Directory.CreateDirectory(dataPath);
				XmlDocument xmlDocument = new XmlDocument();
				xmlDocument.Load(configFileFullPath);
				XmlNodeList xmlNodeList = xmlDocument.SelectNodes("//dniurlwsz/Value");
				foreach (XmlNode item in xmlNodeList)
				{
					if (item.Attributes["Year"].Value == toolStripComboBox1.SelectedItem.ToString().Substring(0, 4))
					{
						item.InnerText = dni.ToString();
						flag = true;
					}
				}
				if (!flag)
				{
					XmlElement xmlElement = xmlDocument.CreateElement("Value");
					XmlAttribute xmlAttribute = xmlDocument.CreateAttribute("Year");
					xmlAttribute.Value = toolStripComboBox1.SelectedItem.ToString().Substring(0, 4);
					xmlElement.Attributes.Append(xmlAttribute);
					xmlElement.InnerText = dni.ToString();
					xmlDocument.DocumentElement.AppendChild(xmlElement);
					xmlDocument.PreserveWhitespace = true;
					xmlDocument.Save(configFileFullPath);
				}
				xmlDocument.Save(configFileFullPath);
			}
			catch
			{
				SystemSounds.Beep.Play();
				MessageBox.Show("Błąd podczas zapisywania pliku konfiguracyjnego", "Błąd");
			}
		}

		private void WersjaToolStripMenuItem_Click(object sender, EventArgs e)
		{
			MessageBox.Show("Wersja " + version + "\nAutor: Kamil Kłonica\nEmail: kamil.klonica@komanord.pl\n©2022", "UrlopyDelegacje " + version );
		}

		private void FillComboBox()
		{
			DirectoryInfo directoryInfo = new DirectoryInfo(dataPath);
			FileInfo[] files = directoryInfo.GetFiles("20*");
			FileInfo[] array = files;
			foreach (FileInfo fileInfo in array)
			{
				if (!toolStripComboBox1.Items.Contains(fileInfo.Name))
				{
					toolStripComboBox1.Items.Add(fileInfo.Name);
				}
			}
		}

		private void KonfiguracjaToolStripMenuItem_Click(object sender, EventArgs e)
		{
			FillComboBox();
		}

		private void ToolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e)
		{
			DateTime date = Convert.ToDateTime("01/02/" + toolStripComboBox1.SelectedItem.ToString().Substring(0, 4));
			urlopyFileFullPath = Path.Combine(dataPath, toolStripComboBox1.SelectedItem.ToString());
			CreateXML(listaUrlopow);
			RefreshConfig();
			FillList();
			FillForm();
			dasdToolStripMenuItem.HideDropDown();
			if (toolStripComboBox1.SelectedItem.ToString().Substring(0, 4) == DateTime.Now.Year.ToString())
			{
				monthCalendar1.SetDate(DateTime.Today);
			}
			else
			{
				monthCalendar1.SetDate(date);
			}
		}

		private void DasdToolStripMenuItem_Click(object sender, EventArgs e)
		{
			FillComboBox();
		}

		private void RefreshConfig()
		{
			XmlDocument xmlDocument = new XmlDocument();
			xmlDocument.Load(configFileFullPath);
			XmlNodeList elementsByTagName = xmlDocument.GetElementsByTagName("Value");
			foreach (XmlElement item in elementsByTagName)
			{
				if (item.GetAttribute("Year") == toolStripComboBox1.SelectedItem.ToString().Substring(0, 4))
				{
					dniurlwsz = Convert.ToInt32(item.InnerText);
				}
			}
		}

		private void DataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
		{
			Details details = new Details();
			details.Show();
			base.Enabled = false;
		}

		public Urlop GetUrlop()
		{
			int rowIndex = dataGridView1.CurrentCell.RowIndex;
			long urlopID = Convert.ToInt64(dataGridView1.Rows[rowIndex].Cells[0].Value);
			return listaUrlopow.Urlopy.Find((Urlop ele) => ele.ID == urlopID);
		}
		public Urlop GetUrlop(long _ID)
		{
			return listaUrlopow.Urlopy.Find((Urlop ele) => ele.ID == _ID);
		}

		public void SetPickersForNewUrlop(DateTime _od, DateTime _do)
		{
			PickerOd.Value = _od;
			PickerDo.Value = _do;
		}

		private void Button4_Click(object sender, EventArgs e)
		{
			openFileDialog1.Title = "Załącz wniosek urlopowy";
			openFileDialog1.FileName = "";
			openFileDialog1.Filter = "Text files (*.docx)|*.docx|All files (*.*)|*.*";
			if (openFileDialog1.ShowDialog() == DialogResult.OK)
			{
				try
				{
					GetUrlop().WniosekPath = openFileDialog1.FileName;
					SaveToXml();
					SystemSounds.Hand.Play();
					MessageBox.Show("Wniosek dodany pomyślnie!", "Sukces");
				}
				catch
				{
					SystemSounds.Beep.Play();
					MessageBox.Show("Błąd podczas dodawania wniosku", "Błąd");
				}
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
		public IRestResponse CheckSwieto(int year, int month, int day)
		{
			string uri = "";
			string country = "";
			string api_key = "";
			if (File.Exists(APIconfigFileFullPath))
			{
				XmlDocument xmlDocument = new XmlDocument();
				xmlDocument.Load(APIconfigFileFullPath);
				XmlNodeList elementsByTagName = xmlDocument.GetElementsByTagName("Value");
				foreach (XmlElement item in elementsByTagName)
				{
					uri = item.GetAttribute("URI");
					country = item.GetAttribute("country");
					api_key = item.InnerText;
				}
                try
                {
					IRestResponse response;
					var client = new RestClient(uri);
					client.Timeout = -1;
					var request = new RestRequest(Method.GET);
					request.AddParameter("api_key", api_key);
					request.AddParameter("country", country);
					request.AddParameter("year", year);
					request.AddParameter("month", month);
					request.AddParameter("day", day);
					response = client.Execute(request);
					System.Threading.Thread.Sleep(1000);
					return response;
                }
                catch
                {
					IRestResponse response;
					var client = new RestClient("https://holidays.abstractapi.com/v1/");
					client.Timeout = -1;
					var request = new RestRequest(Method.GET);
					response = client.Execute(request);
					System.Threading.Thread.Sleep(1000);
					return response;
				}
            }
            else
            {
				IRestResponse response;
				var client = new RestClient("https://holidays.abstractapi.com/v1/");
				client.Timeout = -1;
				var request = new RestRequest(Method.GET);
				response = client.Execute(request);
				System.Threading.Thread.Sleep(1000);
				return response;
			}
		}

		public void CheckWebReq()
        {
			if (File.Exists(APIconfigFileFullPath))
			{
				if (!CheckSwieto(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day).IsSuccessful)
				{
					SystemSounds.Beep.Play();
					MessageBox.Show("Web service nie odpowiada!", "Błąd");
					label6.Text = "Web service brak odp.";
					label6.BackColor = Color.Red;
                }
                else
                {
					label6.Text = "Web service połączony";
					label6.BackColor = Color.Lime;
				}
            }
            else
            {
				label6.Text = "Web service niedostępny";
				label6.BackColor = Color.Yellow;
			}
		}
		private void InitializeComponent()
		{
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.monthCalendar1 = new System.Windows.Forms.MonthCalendar();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.PickerDo = new System.Windows.Forms.DateTimePicker();
            this.PickerOd = new System.Windows.Forms.DateTimePicker();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.button4 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.dasdToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.konfiguracjaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripComboBox1 = new System.Windows.Forms.ToolStripComboBox();
            this.ustawieniaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.wersjaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.label6 = new System.Windows.Forms.Label();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.label8 = new System.Windows.Forms.Label();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.button5 = new System.Windows.Forms.Button();
            this.label7 = new System.Windows.Forms.Label();
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.button6 = new System.Windows.Forms.Button();
            this.label10 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.groupBox7 = new System.Windows.Forms.GroupBox();
            this.textBox6 = new System.Windows.Forms.TextBox();
            this.textBox5 = new System.Windows.Forms.TextBox();
            this.textBox4 = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.groupBox6.SuspendLayout();
            this.groupBox7.SuspendLayout();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(14, 29);
            this.dataGridView1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.Size = new System.Drawing.Size(596, 373);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.DataGridView1_CellClick);
            this.dataGridView1.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.DataGridView1_CellDoubleClick);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.monthCalendar1);
            this.groupBox1.Location = new System.Drawing.Point(618, 125);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.groupBox1.Size = new System.Drawing.Size(308, 387);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Kalendarz";
            // 
            // monthCalendar1
            // 
            this.monthCalendar1.CalendarDimensions = new System.Drawing.Size(1, 2);
            this.monthCalendar1.Enabled = false;
            this.monthCalendar1.FirstDayOfWeek = System.Windows.Forms.Day.Monday;
            this.monthCalendar1.Location = new System.Drawing.Point(31, 42);
            this.monthCalendar1.Margin = new System.Windows.Forms.Padding(10);
            this.monthCalendar1.MaxSelectionCount = 99;
            this.monthCalendar1.Name = "monthCalendar1";
            this.monthCalendar1.ShowWeekNumbers = true;
            this.monthCalendar1.TabIndex = 3;
            // 
            // textBox3
            // 
            this.textBox3.Enabled = false;
            this.textBox3.Location = new System.Drawing.Point(142, 72);
            this.textBox3.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(41, 23);
            this.textBox3.TabIndex = 9;
            // 
            // textBox2
            // 
            this.textBox2.Enabled = false;
            this.textBox2.Location = new System.Drawing.Point(142, 45);
            this.textBox2.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(41, 23);
            this.textBox2.TabIndex = 8;
            // 
            // textBox1
            // 
            this.textBox1.Enabled = false;
            this.textBox1.Location = new System.Drawing.Point(142, 19);
            this.textBox1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(41, 23);
            this.textBox1.TabIndex = 7;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(8, 80);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(58, 15);
            this.label5.TabIndex = 6;
            this.label5.Text = "Pozostało";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(8, 53);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(93, 15);
            this.label4.TabIndex = 5;
            this.label4.Text = "Wykorzystanych";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(8, 27);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(126, 15);
            this.label3.TabIndex = 4;
            this.label3.Text = "Liczba dni urlopowych";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.checkBox1);
            this.groupBox2.Controls.Add(this.button2);
            this.groupBox2.Controls.Add(this.button1);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.PickerDo);
            this.groupBox2.Controls.Add(this.PickerOd);
            this.groupBox2.Location = new System.Drawing.Point(14, 408);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.groupBox2.Size = new System.Drawing.Size(301, 181);
            this.groupBox2.TabIndex = 2;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Dodaj nowy urlop";
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(38, 159);
            this.checkBox1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(201, 19);
            this.checkBox1.TabIndex = 10;
            this.checkBox1.Text = "Zaznacz, jeżeli tworzysz delegację";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // button2
            // 
            this.button2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.button2.Location = new System.Drawing.Point(156, 87);
            this.button2.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(131, 72);
            this.button2.TabIndex = 9;
            this.button2.Text = "Usuń";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.Button2_Click);
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.button1.Location = new System.Drawing.Point(8, 87);
            this.button1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(131, 70);
            this.button1.TabIndex = 8;
            this.button1.Text = "Dodaj";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.Button1_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(5, 59);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(22, 15);
            this.label2.TabIndex = 7;
            this.label2.Text = "Do";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(5, 30);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(23, 15);
            this.label1.TabIndex = 6;
            this.label1.Text = "Od";
            // 
            // PickerDo
            // 
            this.PickerDo.Location = new System.Drawing.Point(38, 52);
            this.PickerDo.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.PickerDo.Name = "PickerDo";
            this.PickerDo.Size = new System.Drawing.Size(248, 23);
            this.PickerDo.TabIndex = 5;
            // 
            // PickerOd
            // 
            this.PickerOd.Location = new System.Drawing.Point(38, 22);
            this.PickerOd.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.PickerOd.Name = "PickerOd";
            this.PickerOd.Size = new System.Drawing.Size(248, 23);
            this.PickerOd.TabIndex = 4;
            this.PickerOd.Leave += new System.EventHandler(this.PickerOd_Leave);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.textBox3);
            this.groupBox3.Controls.Add(this.textBox2);
            this.groupBox3.Controls.Add(this.textBox1);
            this.groupBox3.Controls.Add(this.label5);
            this.groupBox3.Controls.Add(this.label4);
            this.groupBox3.Controls.Add(this.label3);
            this.groupBox3.Location = new System.Drawing.Point(734, 24);
            this.groupBox3.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Padding = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.groupBox3.Size = new System.Drawing.Size(192, 102);
            this.groupBox3.TabIndex = 10;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Dane ";
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.button4);
            this.groupBox4.Controls.Add(this.button3);
            this.groupBox4.Location = new System.Drawing.Point(331, 411);
            this.groupBox4.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Padding = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.groupBox4.Size = new System.Drawing.Size(279, 102);
            this.groupBox4.TabIndex = 11;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Wniosek urlopowy";
            // 
            // button4
            // 
            this.button4.Enabled = false;
            this.button4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.button4.Location = new System.Drawing.Point(136, 22);
            this.button4.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(135, 74);
            this.button4.TabIndex = 1;
            this.button4.Text = "Załącz wniosek";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.Button4_Click);
            // 
            // button3
            // 
            this.button3.Enabled = false;
            this.button3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.button3.Location = new System.Drawing.Point(7, 22);
            this.button3.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(122, 74);
            this.button3.TabIndex = 0;
            this.button3.Text = "Generuj wniosek";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.Button3_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.dasdToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Padding = new System.Windows.Forms.Padding(7, 2, 0, 2);
            this.menuStrip1.Size = new System.Drawing.Size(931, 24);
            this.menuStrip1.TabIndex = 12;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // dasdToolStripMenuItem
            // 
            this.dasdToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.konfiguracjaToolStripMenuItem,
            this.ustawieniaToolStripMenuItem,
            this.wersjaToolStripMenuItem});
            this.dasdToolStripMenuItem.Name = "dasdToolStripMenuItem";
            this.dasdToolStripMenuItem.Size = new System.Drawing.Size(38, 20);
            this.dasdToolStripMenuItem.Text = "Plik";
            this.dasdToolStripMenuItem.Click += new System.EventHandler(this.DasdToolStripMenuItem_Click);
            // 
            // konfiguracjaToolStripMenuItem
            // 
            this.konfiguracjaToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripComboBox1});
            this.konfiguracjaToolStripMenuItem.Name = "konfiguracjaToolStripMenuItem";
            this.konfiguracjaToolStripMenuItem.Size = new System.Drawing.Size(200, 22);
            this.konfiguracjaToolStripMenuItem.Text = "Załaduj poprzedni okres";
            // 
            // toolStripComboBox1
            // 
            this.toolStripComboBox1.Name = "toolStripComboBox1";
            this.toolStripComboBox1.Size = new System.Drawing.Size(121, 23);
            this.toolStripComboBox1.SelectedIndexChanged += new System.EventHandler(this.ToolStripComboBox1_SelectedIndexChanged);
            // 
            // ustawieniaToolStripMenuItem
            // 
            this.ustawieniaToolStripMenuItem.Name = "ustawieniaToolStripMenuItem";
            this.ustawieniaToolStripMenuItem.Size = new System.Drawing.Size(200, 22);
            this.ustawieniaToolStripMenuItem.Text = "Ustawienia";
            this.ustawieniaToolStripMenuItem.Click += new System.EventHandler(this.UstawieniaToolStripMenuItem_Click);
            // 
            // wersjaToolStripMenuItem
            // 
            this.wersjaToolStripMenuItem.Name = "wersjaToolStripMenuItem";
            this.wersjaToolStripMenuItem.Size = new System.Drawing.Size(200, 22);
            this.wersjaToolStripMenuItem.Text = "Wersja";
            this.wersjaToolStripMenuItem.Click += new System.EventHandler(this.WersjaToolStripMenuItem_Click);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.BackColor = System.Drawing.Color.Lime;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.label6.Location = new System.Drawing.Point(54, 23);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(120, 13);
            this.label6.TabIndex = 13;
            this.label6.Text = "Web service połączony";
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.label8);
            this.groupBox5.Controls.Add(this.dateTimePicker1);
            this.groupBox5.Controls.Add(this.button5);
            this.groupBox5.Location = new System.Drawing.Point(331, 519);
            this.groupBox5.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Padding = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.groupBox5.Size = new System.Drawing.Size(280, 70);
            this.groupBox5.TabIndex = 14;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "Sprawdź święto";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(10, 31);
            this.label8.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(36, 15);
            this.label8.TabIndex = 13;
            this.label8.Text = "Dzień";
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.CustomFormat = "";
            this.dateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dateTimePicker1.Location = new System.Drawing.Point(57, 24);
            this.dateTimePicker1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(94, 23);
            this.dateTimePicker1.TabIndex = 10;
            // 
            // button5
            // 
            this.button5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.button5.Location = new System.Drawing.Point(159, 13);
            this.button5.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(113, 51);
            this.button5.TabIndex = 2;
            this.button5.Text = "Sprawdź";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(7, 23);
            this.label7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(39, 15);
            this.label7.TabIndex = 14;
            this.label7.Text = "Status";
            // 
            // groupBox6
            // 
            this.groupBox6.Controls.Add(this.button6);
            this.groupBox6.Controls.Add(this.label10);
            this.groupBox6.Controls.Add(this.label9);
            this.groupBox6.Controls.Add(this.label6);
            this.groupBox6.Controls.Add(this.label7);
            this.groupBox6.Location = new System.Drawing.Point(618, 520);
            this.groupBox6.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Padding = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.groupBox6.Size = new System.Drawing.Size(308, 69);
            this.groupBox6.TabIndex = 15;
            this.groupBox6.TabStop = false;
            this.groupBox6.Text = "Web service";
            // 
            // button6
            // 
            this.button6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.button6.Location = new System.Drawing.Point(182, 12);
            this.button6.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(117, 51);
            this.button6.TabIndex = 14;
            this.button6.Text = "Zmień";
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new System.EventHandler(this.button6_Click);
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(54, 47);
            this.label10.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(27, 15);
            this.label10.TabIndex = 16;
            this.label10.Text = "Kraj";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(7, 47);
            this.label9.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(27, 15);
            this.label9.TabIndex = 15;
            this.label9.Text = "Kraj";
            // 
            // groupBox7
            // 
            this.groupBox7.Controls.Add(this.textBox6);
            this.groupBox7.Controls.Add(this.textBox5);
            this.groupBox7.Controls.Add(this.textBox4);
            this.groupBox7.Location = new System.Drawing.Point(618, 23);
            this.groupBox7.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.groupBox7.Name = "groupBox7";
            this.groupBox7.Padding = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.groupBox7.Size = new System.Drawing.Size(107, 103);
            this.groupBox7.TabIndex = 16;
            this.groupBox7.TabStop = false;
            this.groupBox7.Text = "Legenda";
            // 
            // textBox6
            // 
            this.textBox6.BackColor = System.Drawing.Color.LightBlue;
            this.textBox6.Location = new System.Drawing.Point(7, 73);
            this.textBox6.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.textBox6.Name = "textBox6";
            this.textBox6.ReadOnly = true;
            this.textBox6.Size = new System.Drawing.Size(93, 23);
            this.textBox6.TabIndex = 12;
            this.textBox6.Text = "Delegacja";
            // 
            // textBox5
            // 
            this.textBox5.BackColor = System.Drawing.Color.Green;
            this.textBox5.Location = new System.Drawing.Point(7, 46);
            this.textBox5.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.textBox5.Name = "textBox5";
            this.textBox5.ReadOnly = true;
            this.textBox5.Size = new System.Drawing.Size(92, 23);
            this.textBox5.TabIndex = 11;
            this.textBox5.Text = "Zawiera święto";
            // 
            // textBox4
            // 
            this.textBox4.BackColor = System.Drawing.Color.Yellow;
            this.textBox4.Location = new System.Drawing.Point(7, 20);
            this.textBox4.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.textBox4.Name = "textBox4";
            this.textBox4.ReadOnly = true;
            this.textBox4.Size = new System.Drawing.Size(93, 23);
            this.textBox4.TabIndex = 10;
            this.textBox4.Text = "Nadchodzący";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(931, 603);
            this.Controls.Add(this.groupBox7);
            this.Controls.Add(this.groupBox6);
            this.Controls.Add(this.groupBox5);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.menuStrip1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            this.groupBox6.ResumeLayout(false);
            this.groupBox6.PerformLayout();
            this.groupBox7.ResumeLayout(false);
            this.groupBox7.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

		}

		private void button5_Click(object sender, EventArgs e)
		{
			if (CheckSwieto(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day).Content.Contains("National"))
			{
				SystemSounds.Hand.Play();
				MessageBox.Show($"{dateTimePicker1.Value.ToShortDateString()} jest świętem w {GetKraj()}","Sprawdź święto");
			}
            else
            {
				SystemSounds.Hand.Play();
				MessageBox.Show($"{dateTimePicker1.Value.ToShortDateString()} NIE jest świętem w {GetKraj()}","Sprawdź święto");
			}
		}

		public string GetKraj()
        {
			CheckWebReq();
			dateTimePicker1.MinDate = new DateTime(DateTime.Now.Year, 1, 1);
			dateTimePicker1.MaxDate = new DateTime(DateTime.Now.Year, 12, 31);
			string country = "Brak danych";
			if (File.Exists(APIconfigFileFullPath))
			{
				XmlDocument xmlDocument = new XmlDocument();
				xmlDocument.Load(APIconfigFileFullPath);
				XmlNodeList elementsByTagName = xmlDocument.GetElementsByTagName("Value");
				foreach (XmlElement item in elementsByTagName)
				{
					country = item.GetAttribute("country");
					label10.Text = country;
				}
				button5.Enabled = true;
				return country;
			}
            else
            {
				label10.Text = country;
				button5.Enabled = false;				
				return country;
			}				
        }

        private void button6_Click(object sender, EventArgs e)
        {
			Ustawienia ustawienia = new Ustawienia();
			ustawienia.Show();
		}
    }
}
