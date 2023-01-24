using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;
using System.Text.Json;
using System.Text.Json.Serialization;
using Aspose.Pdf;
using Aspose.Pdf.Text;


//TODO Обработать разрывы страниц в PDF
//TODO Сделать объединение всех PDF
// TODO JSON



namespace ExcelApp
{
    public partial class Form1 : Form
    {

        protected string _path;
        protected string[] dir;

        protected DirectoryInfo pdfFolder;

        protected DirectoryInfo rootFolder;
        protected FileInfo[] localFiles;

        protected DirectoryInfo childFolder;
        protected FileInfo[] objectiveFiles;

        List<SmetaFile> localData = new List<SmetaFile>();
        List<SmetaFile> objectiveData = new List<SmetaFile>();


        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {

        }
        private void startNumberTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != 8)
            {
                e.Handled = true;
            }
        }

        private void BtnSelectFolder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();

            if (fbd.ShowDialog() == DialogResult.OK)
            {
                _path = fbd.SelectedPath;



                rootFolder = new DirectoryInfo(_path);

                if (Directory.Exists($"{_path}\\TEMPdf"))
                {
                    Directory.Delete($"{_path}\\TEMPdf", true);
                }

                if (rootFolder.Exists)
                {
                    localFiles = rootFolder.GetFiles(".", SearchOption.TopDirectoryOnly);

                    dir = Directory.GetDirectories(_path);
                    if (dir.Length > 1)
                    {
                        MessageBox.Show("Количество папок превышает допустимое значение");
                        labelNameFolder.Text = "";
                        return;
                    }
                    else
                    {
                        labelNameFolder.Text = _path;

                        childFolder = new DirectoryInfo(dir[0]);
                        objectiveFiles = childFolder.GetFiles(".", SearchOption.TopDirectoryOnly);

                        infoTextBox.AppendText($"Кол-во всех файлов: {localFiles.Length + objectiveFiles.Length}\n" + Environment.NewLine +
                            $"Кол-во файлов в корневой папке: {localFiles.Length}\n" + Environment.NewLine +
                            $"Кол-во файлов в дочерней папке: {objectiveFiles.Length}\n" + Environment.NewLine +
                            $"Кол-во папок: {dir.Length}" + Environment.NewLine + Environment.NewLine);

                        string[] fileNames = new string[localFiles.Length + objectiveFiles.Length];

                        Directory.GetFiles(_path, ".", SearchOption.TopDirectoryOnly).ToList()
                            .ForEach(f => infoTextBox.AppendText($"\n- {Path.GetFileName(f)}" + Environment.NewLine));

                        infoTextBox.AppendText(Environment.NewLine + $"\n\nПапка {dir[0]}" + Environment.NewLine);

                        Directory.GetFiles(dir[0], ".", SearchOption.TopDirectoryOnly).ToList()
                            .ForEach(f => infoTextBox.AppendText($"\n{Path.GetFileName(f)}"));
                    }
                }
            }
        }



        private void BtnBuild_Click(object sender, EventArgs e)
        {
            if (localFiles != null)
            {
                int countPages = 1;
                int documentNumber = 1;

                if (startNumberTextBox.Text != "")
                {
                    countPages = Convert.ToInt32(startNumberTextBox.Text);
                }

                Excel.Application app = new Excel.Application();

                for (int i = 0; i < objectiveFiles.Length; i++) /// шаблон для объектных смет
                {
                    string filePath = $"{childFolder}\\{objectiveFiles[i]}";
                    Excel.Workbook ObjWorkBook = app.Workbooks.Open($@"{filePath}");
                    Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

                    Regex regex = new Regex(@"(\w*)-(\w*)-(\w*)");
                    MatchCollection match = regex.Matches(ObjWorkSheet.Range["B10"].Value.ToString());

                    int pages = ObjWorkBook.Sheets[1].PageSetup.Pages.Count; /// кол-во страниц на листе
                    countPages += pages;

                    objectiveData.Add(new SmetaFile(
                        match[0].ToString(), // код сметы
                        ObjWorkSheet.Range["B7"].Value.ToString(), // Наименование
                        ObjWorkSheet.Range["F14"].Value.ToString(), // Сумма денег
                        ObjWorkBook.Sheets[1].PageSetup.Pages.Count, // кол-во страниц на листе
                        objectiveFiles[i]));

                    ObjWorkBook.Close();

                    documentNumber++;
                }

                for (int i = 0; i < localFiles.Length; i++) // шаблон для локальных смет
                {
                    string filePath = $"{rootFolder}\\{localFiles[i]}";
                    Excel.Workbook ObjWorkBook = app.Workbooks.Open($@"{filePath}");
                    Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

                    Regex regex = new Regex(@"(\w*)-(\w*)-(\w*)");
                    MatchCollection match = regex.Matches(ObjWorkSheet.Range["A18"].Value.ToString());

                    var pages = ObjWorkBook.Sheets[1].PageSetup.Pages.Count;
                    countPages += pages; // кол-во страниц на листе


                    localData.Add(new SmetaFile(
                        match[0].ToString(), // код сметы
                        ObjWorkSheet.Range["A20"].Value.ToString(), // Наименование
                        ObjWorkSheet.Range["C28"].Value.ToString(), // Сумма денег
                        ObjWorkBook.Sheets[1].PageSetup.Pages.Count, // кол-во страниц на листе
                        localFiles[i]));

                    ObjWorkBook.Close();
                    documentNumber++;
                }

                localData = localData.OrderBy(x => x.Code).ThenBy(x => x.Name).ToList(); // Сортировка по коду и названию
                objectiveData = objectiveData.OrderBy(x => x.Code).ThenBy(x => x.Name).ToList(); // Сортировка по коду и названию



                ///// конвертер Excel to PDF
                //int countCompleted = 0;
                //Directory.CreateDirectory($"{_path}\\TEMPdf");
                //foreach (var file in objectiveData)
                //{
                //    string filePath = $"{_path}\\ОС\\{file.FolderInfo}";
                //    Excel.Workbook workbook = app.Workbooks.Open(filePath);
                //    string tempPDFPath = $"{_path}\\TEMPdf\\{file.FolderInfo}";
                //    app.ActiveWorkbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, tempPDFPath);
                //    workbook.Close();
                //    countCompleted++;
                //    labelCompleted.Text = $"Кол-во обработанных файлов: {countCompleted} из {localFiles.Length + objectiveFiles.Length}";
                //}
                //foreach (var file in localData)
                //{
                //    string filePath = $"{_path}\\{file.FolderInfo}";
                //    Excel.Workbook workbook = app.Workbooks.Open(filePath);
                //    string tempPDFPath = $"{_path}\\TEMPdf\\{file.FolderInfo}";
                //    app.ActiveWorkbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, tempPDFPath);
                //    workbook.Close();
                //    countCompleted++;
                //    labelCompleted.Text = $"Кол-во обработанных файлов: {countCompleted} из {localFiles.Length + objectiveFiles.Length}";
                //}


                app.Quit();







            }
            else
            {
                MessageBox.Show($"Ошибка! Вы не выбрали папку");
            }
        }



        private void Read_Button_Click(object sender, EventArgs e)
        {
            if (localFiles != null)
            {
                Word.Application app = new Word.Application();
                app.Visible = false;
                var wordDoc = app.Documents.Add();

                object oMissing = Type.Missing;

                var Paragraph = wordDoc.Paragraphs.Add();
                var tableRange = Paragraph.Range;

                var header = wordDoc.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                header.Range.Tables.Add(tableRange, 1, 6, oMissing, oMissing).set_Style("Сетка таблицы");
                header.Range.Text = "Содержание";
                



                //Word.Table tbl = wordDoc.Tables[1];

                //tbl.set_Style("Сетка таблицы");
                //tbl.Cell(1, 1).Range.Text = "N п/п";
                //tbl.Cell(1, 2).Range.Text = "N сметы";
                //tbl.Cell(1, 3).Range.Text = "Наименование";
                //tbl.Cell(1, 4).Range.Text = "Всего тыс.руб.";
                //tbl.Cell(1, 5).Range.Text = "Стр.";
                //tbl.Cell(1, 6).Range.Text = "Часть";

                app.ActiveDocument.SaveAs2($@"{_path}\TEST.docx");
                app.Quit();
            }
            else
            {
                MessageBox.Show($"Ошибка! Вы не выбрали папку");
            }
        }


    }
}









