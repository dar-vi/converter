using Newtonsoft.Json.Serialization;
using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;
using System.IO.Pipes;
using CsvHelper;
using Microsoft.VisualStudio.TestPlatform.ObjectModel;
using ClosedXML;
using ClosedXML.Excel;

namespace converter
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }
        public class Document
        {
            public string Name { get ; set; }
            public string StartDate { get; set; }
            public string DeadLine { get; set; }
            public Document(string Name = "", string StartDate = "", string DeadLine = "")
            {
                this.Name = Name;
                this.StartDate = StartDate;
                this.DeadLine = DeadLine;
            }
        }
        public static class ReadDataFromCsv
        {
            public static List<Document> Read(string fileName)
            {
                StreamReader documents = new StreamReader(fileName);
                List<Document> output = new List<Document>();
                if (documents != null)
                {
                    CsvHelper.CsvReader scv = new CsvHelper.CsvReader(documents, CultureInfo.InvariantCulture);
                    output = scv.GetRecords<Document>().ToList();
                }
                documents.Dispose();
                return output;
            }
        }

        public static class ReadDataFromXml
        {
            public static List<Document> Read(string fileName)
            {
                var serializer = new XmlSerializer(typeof(List<Document>));
                using (var reader = new StreamReader(fileName))
                {
                    return (List<Document>)serializer.Deserialize(reader);
                }
            }
        }
        public static class ReadDataFromJson
        {
            public static List<Document> Read(string fileName)
            {
                using (var reader = new StreamReader(fileName))
                {
                    string json = reader.ReadToEnd();
                    return JsonConvert.DeserializeObject<List<Document>>(json);
                }
            }
        }

        public static class ReadDataFromExcel
        {
            public static List<Document> Read(string fileName)
            {
                XLWorkbook documents = new XLWorkbook();
                List<Document> output = new List<Document>();
                var ws = documents.Worksheets.Worksheet(1);
                int row = 1;
                while (ws.Cell($"A{row}").Value.ToString() != "")
                {
                    output.Add(new Document(ws.Cell($"A{row}").Value.ToString(),
                        ws.Cell($"B{row}").Value.ToString(),
                        ws.Cell($"C{row}").Value.ToString()));
                    row++;
                }
                documents.Dispose();
                return output;
            }
        }
        List<Document> documents = new List<Document>();

        private void MainForm_Load(object sender, EventArgs e)
        {
            // Настройка диалоговых окон
            openFileDialog2.Filter = "JSON files (*.json)|*.json|CSV files (*.csv)|*.csv|XML files (*.xml)|*.xml|XLSX files (*.xlsx)|*.xlsx";
            saveFileDialog2.Filter = "JSON files (*.json)|*.json|CSV files (*.csv)|*.csv|XML files (*.xml)|*.xml|XLSX files (*.xlsx)|*.xlsx";
        }


        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            // Открытие диалогового окна выбора файла
            DialogResult result = openFileDialog2.ShowDialog();
            if (result == DialogResult.OK)
            {
                // Чтение данных из исходного файла "Document"
                string fileName = openFileDialog2.FileName;

                switch (comboBox5.Text)
                {
                    case ".json":
                        documents = ReadDataFromJson.Read(fileName);
                        break;
                    case ".csv":
                        documents = ReadDataFromCsv.Read(fileName);
                        break;
                    case ".xml":
                        documents = ReadDataFromXml.Read(fileName);
                        break;
                    case ".xlsx":
                        documents = ReadDataFromExcel.Read(fileName);
                        break;
                    default:
                        MessageBox.Show("Unsupported file format");
                        return;
                }
            }
        }

        private void btnConvert_Click(object sender, EventArgs e)
        {
            // Проверка наличия данных
            if (documents.Count == 0)
            {
                MessageBox.Show("No data to convert");
                return;
            }
            // Конвертация и сохранение данных в выбранный формат
            string fileName = saveFileDialog2.FileName;
            string extension = Path.GetExtension(fileName);
            var format = "";
            switch (comboBox6.Text)
            {
                case ".json":
                    format = "json";
                    break;
                case ".csv":
                    format = "csv";
                    break;
                case ".xml":
                    format = "xml";
                    break;
                case ".xlsx":
                    format = "xlsx";
                    break;

            }
            if (format == "json")
            {
                var jsonSerializerSettings = new JsonSerializerSettings
                {
                    ContractResolver = new CamelCasePropertyNamesContractResolver(),
                    Formatting = Formatting.Indented
                };

                var jsonData = JsonConvert.SerializeObject(documents, jsonSerializerSettings);

                using (var writer = new StreamWriter(saveFileDialog2.FileName))
                {
                    writer.Write(jsonData);
                }
            }
            else if (format == "csv")
            {
                using (var writer = new StreamWriter(saveFileDialog2.FileName))
                using (var csv = new CsvHelper.CsvWriter(writer, CultureInfo.InvariantCulture))
                {
                    csv.WriteRecords(documents);
                }
            }
            else if (format == "xml")
            {
                var xmlSerializer = new XmlSerializer(typeof(List<Document>));

                using (var writer = new StreamWriter(saveFileDialog2.FileName))
                {
                    xmlSerializer.Serialize(writer, documents);
                }
            }
            else if (format == "xlsx")
            {
                var file = new FileInfo("output.xlsx");
                using (var package = new ExcelPackage(file))
                {
                    // Добавление листа и заголовков столбцов
                    var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                    worksheet.Cells[1, 1].Value = "Name";
                    worksheet.Cells[1, 2].Value = "StartDate";
                    worksheet.Cells[1, 3].Value = "DeadLine";

                    // Заполнение таблицы данными из списка documents
                    int row = 2;
                    foreach (var doc in documents)
                    {
                        worksheet.Cells[row, 1].Value = doc.Name;
                        worksheet.Cells[row, 2].Value = doc.StartDate;
                        worksheet.Cells[row, 3].Value = doc.DeadLine;
                        row++;
                    }

                    // Сохранение файла
                    package.Save();
                }
            }
            MessageBox.Show("Конвертация завершена успешно!");
        }

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.comboBox5 = new System.Windows.Forms.ComboBox();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.menuStrip2 = new System.Windows.Forms.MenuStrip();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.comboBox6 = new System.Windows.Forms.ComboBox();
            this.saveFileDialog2 = new System.Windows.Forms.SaveFileDialog();
            this.openFileDialog2 = new System.Windows.Forms.OpenFileDialog();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // comboBox5
            // 
            this.comboBox5.FormattingEnabled = true;
            this.comboBox5.Items.AddRange(new object[] {
            ".json",
            ".csv",
            ".xml",
            ".xlsx"});
            this.comboBox5.Location = new System.Drawing.Point(313, 121);
            this.comboBox5.Name = "comboBox5";
            this.comboBox5.Size = new System.Drawing.Size(121, 21);
            this.comboBox5.TabIndex = 0;
            this.comboBox5.SelectedIndexChanged += new System.EventHandler(this.comboBox5_SelectedIndexChanged);
            // 
            // menuStrip1
            // 
            this.menuStrip1.Location = new System.Drawing.Point(0, 24);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1085, 24);
            this.menuStrip1.TabIndex = 1;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // menuStrip2
            // 
            this.menuStrip2.Location = new System.Drawing.Point(0, 0);
            this.menuStrip2.Name = "menuStrip2";
            this.menuStrip2.Size = new System.Drawing.Size(1085, 24);
            this.menuStrip2.TabIndex = 2;
            this.menuStrip2.Text = "menuStrip2";
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.button1.Location = new System.Drawing.Point(139, 120);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(144, 23);
            this.button1.TabIndex = 3;
            this.button1.Text = "Открыть файл";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.btnSelectFile_Click);
            // 
            // button2
            // 
            this.button2.BackColor = System.Drawing.SystemColors.Info;
            this.button2.Location = new System.Drawing.Point(750, 122);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(129, 25);
            this.button2.TabIndex = 4;
            this.button2.Text = "Скачать файл";
            this.button2.UseVisualStyleBackColor = false;
            this.button2.Click += new System.EventHandler(this.btnConvert_Click);
            // 
            // comboBox6
            // 
            this.comboBox6.FormattingEnabled = true;
            this.comboBox6.Items.AddRange(new object[] {
            ".json",
            ".csv",
            ".xml",
            ".xlsx"});
            this.comboBox6.Location = new System.Drawing.Point(611, 122);
            this.comboBox6.Name = "comboBox6";
            this.comboBox6.Size = new System.Drawing.Size(121, 21);
            this.comboBox6.TabIndex = 5;
            // 
            // openFileDialog2
            // 
            this.openFileDialog2.FileName = "openFileDialog2";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(345, 105);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(49, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "Формат";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(619, 106);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(99, 13);
            this.label2.TabIndex = 7;
            this.label2.Text = "Конвертировать в";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.BackColor = System.Drawing.Color.MistyRose;
            this.label3.Font = new System.Drawing.Font("Arial", 13F, System.Drawing.FontStyle.Bold);
            this.label3.Location = new System.Drawing.Point(435, 69);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(206, 21);
            this.label3.TabIndex = 8;
            this.label3.Text = "КОНВЕРТЕР ФАЙЛОВ";
            this.label3.Click += new System.EventHandler(this.label3_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.ErrorImage = ((System.Drawing.Image)(resources.GetObject("pictureBox1.ErrorImage")));
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.InitialImage = ((System.Drawing.Image)(resources.GetObject("pictureBox1.InitialImage")));
            this.pictureBox1.Location = new System.Drawing.Point(923, 134);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(150, 185);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 9;
            this.pictureBox1.TabStop = false;
            // 
            // Form1
            // 
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
            this.ClientSize = new System.Drawing.Size(1085, 331);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.comboBox6);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.comboBox5);
            this.Controls.Add(this.menuStrip1);
            this.Controls.Add(this.menuStrip2);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}