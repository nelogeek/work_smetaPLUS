
using iTextSharp.text;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace ExcelAPP
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
        readonly Stopwatch stopWatch = new Stopwatch();
        readonly string DesktopFolder = $@"{Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)}\Книга смет";

        public Form1()
        {
            InitializeComponent();
            backgroundWorker.WorkerReportsProgress = true;
            backgroundWorker.WorkerSupportsCancellation = true;
        }

        private void BtnSelectFolder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();

            if (fbd.ShowDialog() == DialogResult.OK)
            {
                _path = fbd.SelectedPath;
                rootFolder = new DirectoryInfo(_path);

                DeleteTempFiles();

                if (rootFolder.Exists)
                {
                    localFiles = rootFolder.GetFiles(".", SearchOption.TopDirectoryOnly);
                    foreach (var file in localFiles)
                    {
                        string fileName = file.Name;

                        Regex regex = new Regex(@".*", RegexOptions.RightToLeft);
                        MatchCollection match = regex.Matches(fileName);
                        string fileNameStr = match[0].ToString();
                        string[] fileType = fileNameStr.Split('.');
                        if (fileType[fileType.Length - 1] != "xlsx" && fileType[fileType.Length - 1] != "xls")
                        {
                            MessageBox.Show($"В папке находится недопустимый файл");
                            return;
                        }
                    }
                    dir = Directory.GetDirectories(_path);
                    if (dir.Length == 1)
                    {
                        labelNameFolder.Text = _path;

                        childFolder = new DirectoryInfo(dir[0]);
                        objectiveFiles = childFolder.GetFiles(".", SearchOption.TopDirectoryOnly);

                        infoTextBox.Clear(); // очистка TextBox

                        infoTextBox.Text = $"Общее количество страниц: {FullBookPageCounter}" + Environment.NewLine;

                        infoTextBox.AppendText(
                            $"Кол-во всех файлов: {localFiles.Length + objectiveFiles.Length}\n" + Environment.NewLine +
                            $"Кол-во папок: {dir.Length}" + Environment.NewLine +
                            $"Кол-во объектных файлов: {objectiveFiles.Length}\n" + Environment.NewLine +
                            $"Кол-во локальных файлов: {localFiles.Length}\n" + Environment.NewLine);

                        infoTextBox.AppendText(Environment.NewLine + $"Объектные файлы:" + Environment.NewLine);

                        Directory.GetFiles(dir[0], ".", SearchOption.TopDirectoryOnly).ToList()
                            .ForEach(f => infoTextBox.AppendText($"\n- {Path.GetFileName(f)}" + Environment.NewLine));

                        infoTextBox.AppendText(Environment.NewLine + $"Локальные файлы:" + Environment.NewLine);

                        Directory.GetFiles(_path, ".", SearchOption.TopDirectoryOnly).ToList()
                            .ForEach(f => infoTextBox.AppendText($"\n- {Path.GetFileName(f)}" + Environment.NewLine));
                    }
                    else
                    {
                        MessageBox.Show("В сметах должна быть только одна папка, которая должна содержать объектные сметы!");
                        labelNameFolder.Text = "Добавьте папку с объектными сметами\"ОС\"";
                        return;
                    }
                    pdfFolder = new DirectoryInfo($"{_path}\\TEMPdf");
                }
            }
        }

        private void BtnBuild_Click(object sender, EventArgs e)
        {
            if (backgroundWorker.IsBusy != true)
            {
                DisableButton();
                backgroundWorker.RunWorkerAsync();
            }
        }

        private void BackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled == true)
            {
                labelCompleted.Text = "Отмена!";
                EnabledButton();
            }
            else if (e.Error != null)
            {
                labelCompleted.Text = "Ошибка: " + e.Error.Message;
                EnabledButton();
            }
            else
            {
                EnabledButton();
                infoTextBox.Clear();
            }
        }

        private void BackgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            labelCompleted.Text = e.UserState.ToString();
        }

        private void BackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            if (_path != null)
            {
                if (Directory.Exists($"{DesktopFolder}"))
                {
                    DialogResult dialogResult = MessageBox.Show("Вы точно хотите заменить папку 'Книга смет'?", "Подтверждение замены папки", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        Directory.Delete(DesktopFolder, true);

                        RunBackgroundWorker_DoWork();
                    }
                    else if (dialogResult == DialogResult.No)
                    {
                        backgroundWorker.ReportProgress(1, "Сборка остановлена...");
                        return;
                    }
                }
                else
                    RunBackgroundWorker_DoWork();
            }
            else
            {
                MessageBox.Show($"Ошибка! Вы не выбрали папку");
                backgroundWorker.ReportProgress(1, "Сборка остановлена...");
                return;
            }
        }

        protected void DisableButton()
        {
            this.StartNumberNumeric.Enabled = false;
            this.CountPagePZNumeric.Enabled = false;
            this.btnBuild.Enabled = false;
            this.btnSelectFolder.Enabled = false;
            this.TwoSidedPrintCheckBox.Enabled = false;
            this.SplitBookContentCheckBox.Enabled = false;
            this.RdPdToggle.Enabled = false;
        }
        protected void EnabledButton()
        {
            this.StartNumberNumeric.Enabled = true;
            this.CountPagePZNumeric.Enabled = true;
            this.btnBuild.Enabled = true;
            this.btnSelectFolder.Enabled = true;
            this.TwoSidedPrintCheckBox.Enabled = true;
            this.SplitBookContentCheckBox.Enabled = true;
            this.RdPdToggle.Enabled = true;
        }

        private bool ExcelParser()
        {
            Excel.Application app = new Excel.Application
            {
                DisplayAlerts = false,
                Visible = false,
                ScreenUpdating = false
            };

            Excel.Workbook eWorkbook;
            Excel.Worksheet eWorksheet;

            try
            {
                for (int i = 0; i < objectiveFiles.Length; i++) //Шаблон для объектных смет
                {
                    string filePath = $"{childFolder}\\{objectiveFiles[i]}";
                    eWorkbook = app.Workbooks.Open($@"{filePath}");
                    eWorksheet = (Excel.Worksheet)eWorkbook.Sheets[1];
                    eWorksheet.PageSetup.Orientation = XlPageOrientation.xlLandscape;
                    Regex regex = new Regex(@"(\w*)-(\w*)-(\w*)");
                    string code = regex.Matches(eWorksheet.Range["E8"].Value.ToString())[0].ToString();
                    string ShortCode = code.Replace("p", "").Replace("р", "").Replace("OC-", "").Replace("ОС-", "");
                    string money = eWorksheet.Range["G12"].Value.ToString();
                    string nameDate = eWorksheet.Range["C5"].Value.ToString();
                    string date = eWorksheet.Range["C18"].Value.ToString().Split(new string[] { " цен " }, StringSplitOptions.None)[1];
                    nameDate += $"\n(в ценах на {date})";

                    if (RdPdToggle.Checked)
                    {
                        eWorksheet.Range["E8"].Replace("ОБЪЕКТНЫЙ СМЕТНЫЙ РАСЧЕТ (СМЕТА)", "ОБЪЕКТНАЯ СМЕТА");
                    }

                    int pages = eWorkbook.Sheets[1].PageSetup.Pages.Count; /// кол-во страниц на листе

                    objectiveData.Add(new SmetaFile(
                        code, // код сметы
                        eWorksheet.Range["C5"].Value.ToString(), // наименование
                        nameDate, // Наименование
                        money, // Сумма денег
                        pages, // кол-во страниц на листе
                        objectiveFiles[i],
                        ShortCode));

                    if (AutoPageBreakeToolStripMenuItem.Checked)
                    {
                        // TODO pageBreaker
                        PageBreaker(eWorksheet);
                    }


                    money = null;
                    pages = 0;
                    nameDate = null;
                    date = null;
                    eWorkbook.Save();
                    eWorkbook.Close(false);
                }

                for (int j = 0; j < localFiles.Length; j++) //Шаблон для локальных смет
                {
                    string filePath = $"{rootFolder}\\{localFiles[j]}";
                    eWorkbook = app.Workbooks.Open($@"{filePath}");
                    eWorksheet = (Excel.Worksheet)eWorkbook.Sheets[1];
                    eWorksheet.PageSetup.Orientation = XlPageOrientation.xlLandscape;

                    Regex regex = new Regex(@"(\w*)-(\w*)-(\w*)");
                    MatchCollection match = regex.Matches(eWorksheet.Range["A18"].Value.ToString());

                    regex = new Regex(@"(\w*)-(\w*)");
                    string shortCode = regex.Matches(match[0].Value.ToString())[0].ToString();

                    string money = eWorksheet.Range["C28"].Value.ToString().Replace("(", "").Replace(")", "");
                    if (money == "0")
                        money = eWorksheet.Range["D28"].Value.ToString().Replace("(", "").Replace(")", "");

                    string nameDate = eWorksheet.Range["A20"].Value.ToString();
                    string date = eWorksheet.Range["D26"].Value.ToString();
                    nameDate += $"\n(в ценах на {date})";

                    if (RdPdToggle.Checked)
                    {
                        eWorksheet.Range["A18"].Replace("ЛОКАЛЬНЫЙ СМЕТНЫЙ РАСЧЕТ (СМЕТА)", "ЛОКАЛЬНАЯ СМЕТА");
                    }


                    int pages = eWorksheet.PageSetup.Pages.Count; /// кол-во страниц на листе

                    if (AutoPageBreakeToolStripMenuItem.Checked)
                    {
                        // TODO pageBreaker
                        PageBreaker(eWorksheet);
                    }

                    localData.Add(new SmetaFile(
                        match[0].ToString(), // код сметы
                        eWorksheet.Range["A20"].Value.ToString(), // наименование
                        nameDate, // Наименование c датой
                        money, // Сумма денег
                        pages, // кол-во страниц на листе
                        localFiles[j],
                        shortCode));

                    money = null;
                    pages = 0;
                    nameDate = null;
                    date = null;
                    eWorkbook.Save();
                    eWorkbook.Close(false);
                }

                localData = localData.OrderBy(x => x.Code).ThenBy(x => x.Name).ToList(); // Сортировка по коду и названию
                objectiveData = objectiveData.OrderBy(x => x.Code).ThenBy(x => x.Name).ToList(); // Сортировка по коду и названию

                app.Quit();
                eWorkbook = null;
                eWorksheet = null;
                GC.Collect();

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка! Неверный шаблон сметы");
                MessageBox.Show(ex.Message.ToString());
                Console.WriteLine(ex.StackTrace);
                Console.WriteLine(ex.Message.ToString());
                backgroundWorker.CancelAsync();
                DeleteTempFiles();
                DeleteTempVar();
                app.Quit();
                eWorkbook = null;
                eWorksheet = null;
                GC.Collect();

                backgroundWorker.ReportProgress(1, "Сборка остановлена...");

                return false;
            }
        }

        protected bool ExcelConverter()
        {
            Excel.Application app = new Excel.Application
            {
                DisplayAlerts = false,
                Visible = false,
                ScreenUpdating = false

            };
            Excel.Workbook eWorkbook;
            Excel.Worksheet eWorksheet;

            try
            {
                /// конвертер Excel to PDF
                int countCompleted = 0;
                Directory.CreateDirectory($"{_path}\\TEMPdf");
                foreach (var file in objectiveData)
                {
                    string filePath = $"{_path}\\ОС\\{file.FolderInfo}";
                    eWorkbook = app.Workbooks.Open(filePath);
                    eWorksheet = (Excel.Worksheet)eWorkbook.Sheets[1];
                    string tempPDFPath = $"{_path}\\TEMPdf\\{file.FolderInfo}";
                    eWorksheet.PageSetup.RightFooter = ""; ///Удаление нумерации станиц в Excel

                    app.ActiveWorkbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, tempPDFPath);
                    eWorkbook.Close(false);
                    countCompleted++;
                }
                foreach (var file in localData)
                {
                    string filePath = $"{_path}\\{file.FolderInfo}";
                    string tempPDFPath = $"{_path}\\TEMPdf\\{file.FolderInfo}";
                    eWorkbook = app.Workbooks.Open(filePath);
                    eWorksheet = (Excel.Worksheet)eWorkbook.Sheets[1];

                    eWorksheet.PageSetup.RightFooter = ""; ///Удаление нумерации стpаниц в Excel
                    app.ActiveWorkbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, tempPDFPath);
                    eWorkbook.Close(false);
                    countCompleted++;
                }
                app.Quit();
                eWorkbook = null;
                eWorksheet = null;
                GC.Collect();
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка конвертации в pdf");
                MessageBox.Show(ex.Message.ToString());
                backgroundWorker.CancelAsync();
                DeleteTempFiles();
                DeleteTempVar();

                app.Quit();
                eWorkbook = null;
                GC.Collect();

                backgroundWorker.ReportProgress(1, "Сборка остановлена...");

                return false;
            }
        }

        protected bool PdfMerge()
        {
            try
            {
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                string fileNameConcatPdf = $"{_path}\\TEMPdf\\smetaBook.pdf";
                string fileNameSmetaPdf = $"{_path}\\TEMPdf\\Сметы.pdf";
                string fileNameTitlePdf = $"{_path}\\TEMPdf\\Содержание.pdf";

                List<SmetaFile> tempFilesList = objectiveData;
                tempFilesList.AddRange(localData);

                SmetaFile lastUsedDocument = null;

                //Объединение PDF
                PdfDocument inputPdfDocument;
                if (partsBookCheckBox.Checked)
                {
                    int bookNumber = 1;
                    int i = 0;
                    while(lastUsedDocument != tempFilesList[tempFilesList.Count - 1])
                    {
                        
                        PdfDocument outputSmetaPdfDocument = new PdfDocument();
                        for (; i < tempFilesList.Count; i++)
                        {
                            var smetaFile = tempFilesList[i];
                            inputPdfDocument = PdfReader.Open($"{pdfFolder}\\{smetaFile.FolderInfo}.pdf", PdfDocumentOpenMode.Import);
                            int pageCountInputDocument = inputPdfDocument.PageCount;

                            if (outputSmetaPdfDocument.PageCount + pageCountInputDocument < pagesInPartBookNumeric.Value + 50)
                            {
                                for (int j = 0; j < pageCountInputDocument; j++)
                                {
                                    PdfPage page = inputPdfDocument.Pages[j];
                                    outputSmetaPdfDocument.AddPage(page);
                                }
                                lastUsedDocument = smetaFile;
                                inputPdfDocument.Close();
                            }
                            else
                                break;
                        }
                        outputSmetaPdfDocument.Save($@"{DesktopFolder}\Сметы{bookNumber}.pdf");
                        outputSmetaPdfDocument.Close();
                        AddPageNumberSmetaITextSharp($@"{DesktopFolder}\Сметы{bookNumber}.pdf");
                        bookNumber++;
                    }
                }
                else
                {
                    if (SplitBookContentCheckBox.Checked)
                    {
                        PdfDocument outputSmetaPdfDocument = new PdfDocument();
                        foreach (var file in tempFilesList)
                        {
                            inputPdfDocument = PdfReader.Open($"{pdfFolder}\\{file.FolderInfo}.pdf", PdfDocumentOpenMode.Import);
                            for (int i = 0; i < inputPdfDocument.PageCount; i++)
                            {
                                PdfPage page = inputPdfDocument.Pages[i];
                                outputSmetaPdfDocument.AddPage(page);
                            }
                            inputPdfDocument.Close();
                        }
                        outputSmetaPdfDocument.Save(fileNameSmetaPdf);
                        outputSmetaPdfDocument.Close();
                    }
                    else
                    {
                        PdfDocument outputPdfDocument = new PdfDocument();
                        //Добавляем содержание
                        inputPdfDocument = PdfReader.Open(fileNameTitlePdf, PdfDocumentOpenMode.Import);
                        for (int i = 0; i < inputPdfDocument.PageCount; i++)
                        {
                            PdfPage page = inputPdfDocument.Pages[i];
                            outputPdfDocument.AddPage(page);
                        }
                        //Добавляем сметы
                        foreach (var file in tempFilesList)
                        {
                            inputPdfDocument = PdfReader.Open($"{pdfFolder}\\{file.FolderInfo}.pdf", PdfDocumentOpenMode.Import);
                            for (int i = 0; i < inputPdfDocument.PageCount; i++)
                            {
                                PdfPage page = inputPdfDocument.Pages[i];
                                outputPdfDocument.AddPage(page);
                            }
                        }
                        outputPdfDocument.Save(fileNameConcatPdf);
                        inputPdfDocument.Close();
                        outputPdfDocument.Close();
                    }
                    if (SplitBookContentCheckBox.Checked) //Нумерация страниц
                    {
                        //AddPageNumberTitleITextSharp(fileNameTitlePdf);
                        AddPageNumberSmetaITextSharp(fileNameSmetaPdf);
                    }
                    else
                    {
                        AddPageNumberITextSharp(fileNameConcatPdf);
                    }
                }
                
                return true;
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка сборки книги");
                backgroundWorker.CancelAsync();
                DeleteTempFiles();
                DeleteTempVar();
                backgroundWorker.ReportProgress(1, "Сборка остановлена...");
                return false;
            }
        }

        protected void AddPageNumberTitleITextSharp(string fileTitlePath)
        {
            try
            {
                byte[] bytesTitle = File.ReadAllBytes(fileTitlePath);

                iTextSharp.text.Font blackFont = FontFactory.GetFont("Arial", 8, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
                using (MemoryStream stream = new MemoryStream())
                {
                    iTextSharp.text.pdf.PdfReader readerTitle = new iTextSharp.text.pdf.PdfReader(bytesTitle);
                    int pagesTitle = readerTitle.NumberOfPages;

                    using (iTextSharp.text.pdf.PdfStamper stamper = new iTextSharp.text.pdf.PdfStamper(readerTitle, stream))
                    {
                        int startPageNumber = Convert.ToInt32(StartNumberNumeric.Value) - 1;

                        //Нумерация страниц содержания
                        if (TwoSidedPrintCheckBox.Checked)
                            for (int i = 1; i <= pagesTitle; i++)
                            {
                                if ((i + startPageNumber) % 2 == 0)
                                    iTextSharp.text.pdf.ColumnText.ShowTextAligned(stamper.GetUnderContent(i), Element.ALIGN_RIGHT, new Phrase((i + startPageNumber).ToString(), blackFont), 30f, 810f, 0);
                                else
                                    iTextSharp.text.pdf.ColumnText.ShowTextAligned(stamper.GetUnderContent(i), Element.ALIGN_RIGHT, new Phrase((i + startPageNumber).ToString(), blackFont), 565f, 810f, 0);
                            }
                        else
                            for (int i = 1; i <= pagesTitle; i++)
                                iTextSharp.text.pdf.ColumnText.ShowTextAligned(stamper.GetUnderContent(i), Element.ALIGN_RIGHT, new Phrase((i + startPageNumber).ToString(), blackFont), 565f, 15f, 0);
                    }
                    bytesTitle = stream.ToArray();
                    readerTitle.Close();
                }
                File.WriteAllBytes(fileTitlePath, bytesTitle);
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка нумерации содержания");
                DeleteTempFiles();
                DeleteTempVar();
                backgroundWorker.ReportProgress(1, "Сборка остановлена...");
                backgroundWorker.CancelAsync();
            }
        }

        protected void AddPageNumberSmetaITextSharp(string filePath)
        {
            try
            {
                byte[] bytes = File.ReadAllBytes(filePath);
                PdfDocument titleDocument = PdfReader.Open($"{_path}\\TEMPdf\\Содержание.pdf");

                iTextSharp.text.Font blackFont = FontFactory.GetFont("Arial", 8, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
                using (MemoryStream stream = new MemoryStream())
                {
                    iTextSharp.text.pdf.PdfReader reader = new iTextSharp.text.pdf.PdfReader(bytes);
                    int titlePages = titleDocument.PageCount;
                    int pagesBook = reader.NumberOfPages;
                    //int afterTitleNumericPages = Convert.ToInt32(afterTitleNumeric.Value);

                    using (iTextSharp.text.pdf.PdfStamper stamper = new iTextSharp.text.pdf.PdfStamper(reader, stream))
                    {
                        int startPageNumber = Convert.ToInt32(StartNumberNumeric.Value) - 1;
                        int pagesPzCount = Convert.ToInt32(CountPagePZNumeric.Value);

                        if (TwoSidedPrintCheckBox.Checked)
                        {
                            if ((startPageNumber + titlePages) % 2 == 1)
                                titlePages++;
                            if (pagesPzCount % 2 == 1)
                                pagesPzCount++;

                            for (int i = 1; i <= pagesBook; i++)
                            {
                                if ((startPageNumber + titlePages + pagesPzCount + i) % 2 == 0)
                                    iTextSharp.text.pdf.ColumnText.ShowTextAligned(stamper.GetUnderContent(i), Element.ALIGN_RIGHT, new Phrase((i + startPageNumber + pagesPzCount + titlePages).ToString(), blackFont), 810f, 575f, 0);
                                else
                                    iTextSharp.text.pdf.ColumnText.ShowTextAligned(stamper.GetUnderContent(i), Element.ALIGN_RIGHT, new Phrase((i + startPageNumber + pagesPzCount + titlePages).ToString(), blackFont), 810f, 15f, 0);
                            }
                        }
                        else
                        {
                            if ((startPageNumber + titlePages) % 2 == 1)
                                titlePages++;
                            if (pagesPzCount % 2 == 1)
                                pagesPzCount++;
                            for (int i = 1; i <= pagesBook; i++)
                                iTextSharp.text.pdf.ColumnText.ShowTextAligned(stamper.GetUnderContent(i), Element.ALIGN_RIGHT, new Phrase((i + startPageNumber + pagesPzCount + titlePages).ToString(), blackFont), 810f, 15f, 0);
                        }
                    }
                    titleDocument.Close();
                    bytes = stream.ToArray();
                    reader.Close();
                }
                File.WriteAllBytes(filePath, bytes);
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка нумерации смет");
                DeleteTempFiles();
                DeleteTempVar();
                backgroundWorker.ReportProgress(1, "Сборка остановлена...");
                backgroundWorker.CancelAsync();
            }
        }

        protected void AddPageNumberITextSharp(string filePath)
        {
            try
            {
                byte[] bytes = File.ReadAllBytes(filePath);
                byte[] bytesTitle = File.ReadAllBytes($"{_path}\\TEMPdf\\Содержание.pdf");

                iTextSharp.text.Font blackFont = FontFactory.GetFont("Arial", 8, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
                using (MemoryStream stream = new MemoryStream())
                {
                    iTextSharp.text.pdf.PdfReader reader = new iTextSharp.text.pdf.PdfReader(bytes);
                    iTextSharp.text.pdf.PdfReader readerOnlyTitle = new iTextSharp.text.pdf.PdfReader(bytesTitle);
                    int titlePages = readerOnlyTitle.NumberOfPages;
                    int pagesBook = reader.NumberOfPages;
                    //int afterTitleNumericPages = Convert.ToInt32(afterTitleNumeric.Value);

                    using (iTextSharp.text.pdf.PdfStamper stamper = new iTextSharp.text.pdf.PdfStamper(reader, stream))
                    {
                        int startPageNumber = Convert.ToInt32(StartNumberNumeric.Value) - 1;
                        int pagesPzCount = Convert.ToInt32(CountPagePZNumeric.Value);

                        if (TwoSidedPrintCheckBox.Checked)
                        {
                            //Нумерация страниц содержания
                            //for (int i = 1; i <= titlePages; i++)
                            //    iTextSharp.text.pdf.ColumnText.ShowTextAligned(stamper.GetUnderContent(i), Element.ALIGN_RIGHT, new Phrase((i + startPageNumber).ToString(), blackFont), 565f, 15f, 0);
                            if ((startPageNumber + titlePages) % 2 == 1)
                                titlePages++;
                            if (pagesPzCount % 2 == 1)
                                pagesPzCount++;

                            for (int i = 1; i <= pagesBook; i++)
                            {
                                if ((startPageNumber + titlePages + pagesPzCount + i) % 2 == 0)
                                    iTextSharp.text.pdf.ColumnText.ShowTextAligned(stamper.GetUnderContent(i), Element.ALIGN_RIGHT, new Phrase((i + startPageNumber + pagesPzCount + titlePages).ToString(), blackFont), 810f, 575f, 0);
                                else
                                    iTextSharp.text.pdf.ColumnText.ShowTextAligned(stamper.GetUnderContent(i), Element.ALIGN_RIGHT, new Phrase((i + startPageNumber + pagesPzCount + titlePages).ToString(), blackFont), 810f, 15f, 0);
                            }
                        }
                        else
                        {
                            if ((startPageNumber + titlePages) % 2 == 1)
                                titlePages++;
                            if (pagesPzCount % 2 == 1)
                                pagesPzCount++;
                            for (int i = 1; i <= pagesBook; i++)
                                iTextSharp.text.pdf.ColumnText.ShowTextAligned(stamper.GetUnderContent(i), Element.ALIGN_RIGHT, new Phrase((i + startPageNumber + pagesPzCount + titlePages).ToString(), blackFont), 810f, 15f, 0);
                        }
                    }
                    bytes = stream.ToArray();
                    reader.Close();
                }
                File.WriteAllBytes(filePath, bytes);
            }
            catch (Exception)
            {
                DeleteTempFiles();
                DeleteTempVar();
                MessageBox.Show("Ошибка нумерации книги");
                backgroundWorker.ReportProgress(1, "Сборка остановлена...");
                backgroundWorker.CancelAsync();
            }
        }

        protected bool TitleGeneration()
        {
            // ---------------- Генерация содержания ----------------------------------------------------------------------------

            int NumberDocument = 1; // номер документа
            int row = 1; // номер заполняемой строки таблицы

            Word.Application wordApp = new Word.Application
            {

                Visible = false,
                ScreenUpdating = false
            };


            try
            {

                if (objectiveData.Count != 0)
                {
                    object oMissing = Type.Missing;
                    Object defaultTableBehavior = Word.WdDefaultTableBehavior.wdWord9TableBehavior;
                    Object autoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitWindow;


                    var wDocument = wordApp.Documents.Add();


                    // настройка полей документа
                    wDocument.PageSetup.TopMargin = wordApp.InchesToPoints(0.5f);
                    wDocument.PageSetup.BottomMargin = wordApp.InchesToPoints(0.5f);
                    wDocument.PageSetup.LeftMargin = wordApp.InchesToPoints(0.4f);
                    wDocument.PageSetup.RightMargin = wordApp.InchesToPoints(0.4f);
                    wDocument.PageSetup.HeaderDistance = 20f;

                    wDocument.Sections[1].PageSetup.OddAndEvenPagesHeaderFooter = -1; // -1 = true  - настройка: четные-нечетные страницы

                    Word.Range headerRange = wDocument.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;

                    wDocument.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
                    wDocument.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.RestartNumberingAtSection = true;
                    wDocument.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.StartingNumber = (int)StartNumberNumeric.Value; // номер первой страницы

                    // колонтитул нечетной страницы
                    wDocument.Tables.Add(headerRange, 1, 6, ref defaultTableBehavior, ref autoFitBehavior);
                    Word.Table headerTable = headerRange.Tables[1];

                    headerTable.Borders.Enable = 0;
                    Word.Range rangePageNum = headerTable.Range.Cells[headerTable.Range.Cells.Count].Range;
                    rangePageNum.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    Word.Field fld = rangePageNum.Document.Fields.Add(rangePageNum, oMissing, "Page", false);
                    Word.Range rangeFieldPageNum = fld.Result;
                    rangeFieldPageNum.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                    headerTable.Cell(1, 6).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                    headerTable.Cell(1, 6).Range.Font.Size = 12;

                    headerTable.Rows.Add();
                    headerTable.Cell(2, 3).Range.Text = "Содержание";
                    headerTable.Cell(2, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    headerTable.Cell(2, 3).Range.Font.Name = "Times New Roman";
                    headerTable.Cell(2, 3).Range.Font.Size = 12;
                    headerTable.Cell(2, 3).Range.Font.Italic = 1;
                    headerTable.Cell(2, 3).Range.Font.Bold = 1;
                    headerTable.Cell(2, 3).Range.Font.Color = Word.WdColor.wdColorBlack;

                    // заполнение таблицы
                    headerTable.Rows.Add();
                    headerTable.Rows[3].Borders.Enable = 1;
                    headerTable.Cell(3, 1).Range.Text = "N п/п";
                    headerTable.Cell(3, 2).Range.Text = "N сметы";
                    headerTable.Cell(3, 3).Range.Text = "Наименование";
                    headerTable.Cell(3, 4).Range.Text = "Всего тыс.руб.";
                    headerTable.Cell(3, 5).Range.Text = "Стр.";
                    headerTable.Cell(3, 6).Range.Text = "Часть";
                    // изменение параметров таблицы
                    headerTable.Rows[3].Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    headerTable.Rows[3].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    headerTable.Rows[3].Range.Font.Name = "Times New Roman";
                    headerTable.Rows[3].Range.Font.Italic = 0;
                    headerTable.Rows[3].Range.Font.Bold = 1;
                    headerTable.Rows[3].Range.Font.Size = 10;
                    headerTable.Rows[3].Range.Font.Color = Word.WdColor.wdColorBlack;
                    // ширина ячеек таблицы
                    headerTable.Columns[1].PreferredWidth = 6f;
                    headerTable.Columns[2].PreferredWidth = 9f;
                    headerTable.Columns[3].PreferredWidth = 30f;
                    headerTable.Columns[4].PreferredWidth = 9f;
                    headerTable.Columns[5].PreferredWidth = 4f;
                    headerTable.Columns[6].PreferredWidth = 4f;

                    // колонтитул четных страниц
                    headerRange = wDocument.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages].Range;

                    wDocument.Tables.Add(headerRange, 1, 6, ref defaultTableBehavior, ref autoFitBehavior);
                    headerTable = headerRange.Tables[1];

                    headerTable.Borders.Enable = 0;
                    rangePageNum = headerTable.Range.Cells[1].Range;
                    rangePageNum.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    fld = rangePageNum.Document.Fields.Add(rangePageNum, oMissing, "Page", false);
                    rangeFieldPageNum = fld.Result;
                    rangeFieldPageNum.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                    headerTable.Cell(1, 6).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                    headerTable.Cell(1, 6).Range.Font.Size = 12;

                    headerTable.Rows.Add();
                    headerTable.Cell(2, 3).Range.Text = "Содержание";
                    headerTable.Cell(2, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    headerTable.Cell(2, 3).Range.Font.Name = "Times New Roman";
                    headerTable.Cell(2, 3).Range.Font.Size = 12;
                    headerTable.Cell(2, 3).Range.Font.Italic = 1;
                    headerTable.Cell(2, 3).Range.Font.Bold = 1;
                    headerTable.Cell(2, 3).Range.Font.Color = Word.WdColor.wdColorBlack;

                    // заполнение таблицы
                    headerTable.Rows.Add();
                    headerTable.Rows[3].Borders.Enable = 1;
                    headerTable.Cell(3, 1).Range.Text = "N п/п";
                    headerTable.Cell(3, 2).Range.Text = "N сметы";
                    headerTable.Cell(3, 3).Range.Text = "Наименование";
                    headerTable.Cell(3, 4).Range.Text = "Всего тыс.руб.";
                    headerTable.Cell(3, 5).Range.Text = "Стр.";
                    headerTable.Cell(3, 6).Range.Text = "Часть";
                    // изменение параметров таблицы
                    headerTable.Rows[3].Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    headerTable.Rows[3].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    headerTable.Rows[3].Range.Font.Name = "Times New Roman";
                    headerTable.Rows[3].Range.Font.Italic = 0;
                    headerTable.Rows[3].Range.Font.Bold = 1;
                    headerTable.Rows[3].Range.Font.Size = 10;
                    headerTable.Rows[3].Range.Font.Color = Word.WdColor.wdColorBlack;
                    // ширина ячеек таблицы
                    headerTable.Columns[1].PreferredWidth = 6f;
                    headerTable.Columns[2].PreferredWidth = 9f;
                    headerTable.Columns[3].PreferredWidth = 30f;
                    headerTable.Columns[4].PreferredWidth = 9f;
                    headerTable.Columns[5].PreferredWidth = 4f;
                    headerTable.Columns[6].PreferredWidth = 4f;

                    // основная часть
                    wDocument.Paragraphs.Add(ref oMissing);
                    var Paragraph = wDocument.Paragraphs[1];
                    var Range = Paragraph.Range;
                    var Table = wDocument.Tables.Add(Range, 1, 6, ref defaultTableBehavior, ref autoFitBehavior);
                    // ширина ячеек таблицы
                    Table.Columns[1].PreferredWidth = 6f;
                    Table.Columns[2].PreferredWidth = 9f;
                    Table.Columns[3].PreferredWidth = 30f;
                    Table.Columns[4].PreferredWidth = 9f;
                    Table.Columns[5].PreferredWidth = 4f;
                    Table.Columns[6].PreferredWidth = 4f;
                    //---
                    Table.Range.Rows[row].Range.Font.Name = "Times New Roman";
                    Table.Rows[row].Range.Font.Size = 10;
                    Table.Cell(row, 1).Range.Font.Size = 9;
                    Table.Rows[row].Range.Font.Bold = 0;
                    Table.Rows[row].Range.Font.Color = Word.WdColor.wdColorBlack;
                    Table.Cell(row, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    Table.Cell(row, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    Table.Cell(row, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    Table.Cell(row, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                    Table.Cell(row, 5).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    Table.Cell(row, 6).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    Table.Rows[row].Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop;
                    Table.Borders.Enable = 0;
                    // заполнение таблицы
                    Table.Rows.Add();
                    row++;
                    // пояснительная записка
                    Table.Rows.Add();
                    Table.Cell(row, 1).Range.Text = NumberDocument.ToString();
                    Table.Cell(row, 3).Range.Text = "Пояснительная записка" + "\n";
                    row++;
                    // ОБЪЕКТНЫЕ СМЕТЫ
                    Table.Rows.Add();
                    Table.Cell(row, 3).Range.Text = "ОБЪЕКТНЫЕ СМЕТЫ" + "\n";
                    // изменение параметров строки
                    Table.Cell(row, 3).Range.Font.Bold = 1;
                    Table.Cell(row, 3).Range.Font.Size = 14;
                    Table.Cell(row, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    row++;

                    // вывод объектной сметы
                    foreach (var data in objectiveData)
                    {
                        NumberDocument++;
                        Table.Rows.Add();
                        Table.Cell(row, 1).Range.Text = NumberDocument.ToString();
                        Table.Cell(row, 2).Range.Text = data.Code;
                        Table.Cell(row, 3).Range.Text = data.NameDate + "\n";
                        Table.Cell(row, 4).Range.Text = data.Price;
                        // изменение параметров строки
                        Table.Cell(row, 3).Range.Font.Bold = 0;
                        Table.Cell(row, 3).Range.Font.Size = 10;
                        Table.Cell(row, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                        row++;
                    }

                    // удаление повторяющихся номеров объектных смет
                    List<Pair> pairs = new List<Pair>();
                    foreach (var data in objectiveData)
                    {
                        pairs.Add(new Pair() { Key = data.ShortCode, Value = data.Name });
                    }
                    var setDict = pairs.GroupBy(x => x.Key.Trim()).Select(y => y.FirstOrDefault());

                    // шапка локальных смет
                    Table.Rows.Add();
                    Table.Cell(row, 3).Range.Text = "ЛОКАЛЬНЫЕ СМЕТЫ" + "\n";
                    // изменение параметров строки
                    Table.Cell(row, 3).Range.Font.Bold = 1;
                    Table.Cell(row, 3).Range.Font.Size = 14;
                    Table.Cell(row, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    row++;

                    // вывод локальных смет относительно объектных
                    foreach (var oData in setDict)
                    {
                        Table.Rows.Add();
                        Table.Cell(row, 3).Range.Text = oData.Value + "\n";
                        // изменение параметров строки
                        Table.Cell(row, 3).Range.Font.Bold = 0;
                        Table.Cell(row, 3).Range.Font.Size = 10;
                        Table.Cell(row, 3).Range.Font.Italic = 1;
                        Table.Cell(row, 3).Range.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
                        Table.Cell(row, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        row++;

                        foreach (var lData in localData)
                        {
                            if (lData.ShortCode == oData.Key)
                            {
                                NumberDocument++;
                                Table.Rows.Add();
                                Table.Cell(row, 1).Range.Text = NumberDocument.ToString();
                                Table.Cell(row, 2).Range.Text = lData.Code;
                                Table.Cell(row, 3).Range.Text = lData.NameDate + "\n";
                                Table.Cell(row, 4).Range.Text = lData.Price;
                                // изменение параметров строки
                                Table.Cell(row, 3).Range.Font.Size = 10;
                                Table.Cell(row, 3).Range.Font.Italic = 0;
                                Table.Cell(row, 3).Range.Font.Bold = 0;
                                Table.Cell(row, 3).Range.Font.Underline = Word.WdUnderline.wdUnderlineNone;
                                Table.Cell(row, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                                row++;
                            }
                        }
                    }

                    //нумерация страниц
                    int pagesInTitle = wDocument.ComputeStatistics(WdStatistic.wdStatisticPages, false) - 1; // кол-во страниц в содержании
                    int pageNumber = (int)StartNumberNumeric.Value + pagesInTitle; // номер страницы

                    row = 2;

                    if (TwoSidedPrintCheckBox.Checked)
                    {
                        //TODO
                        // добавление страниц после содержания 

                        // нумерация ПЗ
                        if ((pageNumber % 2) == 0)
                        {
                            pageNumber += 1;
                        }
                        else
                        {
                            pageNumber += 2;
                        }
                        Table.Cell(row, 5).Range.Text = pageNumber.ToString();
                        pageNumber += (int)CountPagePZNumeric.Value - 1;
                        row += 2;

                        // нумерация сметы
                        if ((pageNumber % 2) == 0)
                        {
                            pageNumber += 1;
                        }
                        else
                        {
                            pageNumber += 2;
                        }

                        foreach (var data in objectiveData) // объектные сметы
                        {
                            Table.Cell(row, 5).Range.Text = pageNumber.ToString();
                            pageNumber += data.PageCount;
                            row++;
                        }


                        row++;

                        foreach (var oData in setDict) // локальные сметы
                        {
                            row++;

                            foreach (var lData in localData)
                            {
                                if (lData.ShortCode == oData.Key)
                                {
                                    Table.Cell(row, 5).Range.Text = pageNumber.ToString();
                                    pageNumber += lData.PageCount;
                                    row++;
                                }
                            }
                        }


                    }
                    else
                    {
                        // добавление страниц после содержания TODO

                        // нумерация ПЗ
                        pageNumber++;
                        Table.Cell(row, 5).Range.Text = pageNumber.ToString();
                        pageNumber += (int)CountPagePZNumeric.Value - 1;
                        row += 2;

                        // нумерация сметы
                        foreach (var data in objectiveData) // объектные сметы
                        {
                            Table.Cell(row, 5).Range.Text = pageNumber.ToString();
                            pageNumber += data.PageCount;
                            row++;
                        }
                        row++;

                        foreach (var oData in setDict) // локальные сметы
                        {
                            row++;

                            foreach (var lData in localData)
                            {
                                if (lData.ShortCode == oData.Key)
                                {
                                    Table.Cell(row, 5).Range.Text = pageNumber.ToString();
                                    pageNumber += lData.PageCount;
                                    row++;
                                }
                            }
                        }
                    }

                    wDocument.SaveAs2($"{pdfFolder}\\Содержание.docx");
                    wDocument.ExportAsFixedFormat($"{pdfFolder}\\Содержание.pdf", Word.WdExportFormat.wdExportFormatPDF);
                    wDocument.Close(Word.WdSaveOptions.wdDoNotSaveChanges, Word.WdOriginalFormat.wdOriginalDocumentFormat, false);

                }
                else
                {

                    object oMissing = Type.Missing;
                    Object defaultTableBehavior = Word.WdDefaultTableBehavior.wdWord9TableBehavior;
                    Object autoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitWindow;

                    var wDocument = wordApp.Documents.Add();

                    // настройка полей документа
                    wDocument.PageSetup.TopMargin = wordApp.InchesToPoints(0.5f);
                    wDocument.PageSetup.BottomMargin = wordApp.InchesToPoints(0.5f);
                    wDocument.PageSetup.LeftMargin = wordApp.InchesToPoints(0.65f);
                    wDocument.PageSetup.RightMargin = wordApp.InchesToPoints(0.4f);
                    wDocument.PageSetup.HeaderDistance = 20f;

                    Word.Range headerRange = wDocument.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;

                    wDocument.Tables.Add(headerRange, 1, 6, ref defaultTableBehavior, ref autoFitBehavior);
                    Word.Table headerTable = headerRange.Tables[1];

                    headerTable.Borders.Enable = 0;
                    Word.Range rangePageNum = headerTable.Range.Cells[headerTable.Range.Cells.Count].Range;
                    object oCollapseStart = Word.WdCollapseDirection.wdCollapseStart;
                    rangePageNum.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    Word.Field fld = rangePageNum.Document.Fields.Add(rangePageNum, oMissing, "Page", false);
                    Word.Range rangeFieldPageNum = fld.Result;
                    rangeFieldPageNum.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    headerTable.Cell(1, 6).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                    headerTable.Cell(1, 6).Range.Font.Size = 12;

                    headerTable.Rows.Add();
                    headerTable.Cell(2, 3).Range.Text = "Содержание";
                    headerTable.Cell(2, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    headerTable.Cell(2, 3).Range.Font.Name = "Times New Roman";
                    headerTable.Cell(2, 3).Range.Font.Size = 12;
                    headerTable.Cell(2, 3).Range.Font.Italic = 1;
                    headerTable.Cell(2, 3).Range.Font.Bold = 1;
                    headerTable.Cell(2, 3).Range.Font.Color = Word.WdColor.wdColorBlack;

                    // заполнение таблицы
                    headerTable.Rows.Add();
                    headerTable.Rows[3].Borders.Enable = 1;
                    headerTable.Cell(3, 1).Range.Text = "N п/п";
                    headerTable.Cell(3, 2).Range.Text = "N сметы";
                    headerTable.Cell(3, 3).Range.Text = "Наименование";
                    headerTable.Cell(3, 4).Range.Text = "Всего тыс.руб.";
                    headerTable.Cell(3, 5).Range.Text = "Стр.";
                    headerTable.Cell(3, 6).Range.Text = "Часть";
                    // изменение параметров таблицы
                    headerTable.Rows[3].Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    headerTable.Rows[3].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    headerTable.Rows[3].Range.Font.Name = "Times New Roman";
                    headerTable.Rows[3].Range.Font.Italic = 0;
                    headerTable.Rows[3].Range.Font.Bold = 1;
                    headerTable.Rows[3].Range.Font.Size = 10;
                    headerTable.Rows[3].Range.Font.Color = Word.WdColor.wdColorBlack;
                    // ширина ячеек таблицы
                    headerTable.Columns[1].PreferredWidth = 6f;
                    headerTable.Columns[2].PreferredWidth = 9f;
                    headerTable.Columns[3].PreferredWidth = 30f;
                    headerTable.Columns[4].PreferredWidth = 9f;
                    headerTable.Columns[5].PreferredWidth = 4f;
                    headerTable.Columns[6].PreferredWidth = 4f;

                    // основная часть
                    wDocument.Paragraphs.Add(ref oMissing);
                    var Paragraph = wDocument.Paragraphs[1];
                    var Range = Paragraph.Range;
                    var Table = wDocument.Tables.Add(Range, 1, 6, ref defaultTableBehavior, ref autoFitBehavior);
                    // ширина ячеек таблицы
                    Table.Columns[1].PreferredWidth = 6f;
                    Table.Columns[2].PreferredWidth = 9f;
                    Table.Columns[3].PreferredWidth = 30f;
                    Table.Columns[4].PreferredWidth = 9f;
                    Table.Columns[5].PreferredWidth = 4f;
                    Table.Columns[6].PreferredWidth = 4f;
                    //---
                    Table.Range.Rows[row].Range.Font.Name = "Times New Roman";
                    Table.Rows[row].Range.Font.Size = 10;
                    Table.Cell(row, 1).Range.Font.Size = 9;
                    Table.Rows[row].Range.Font.Bold = 0;
                    Table.Rows[row].Range.Font.Color = Word.WdColor.wdColorBlack;
                    Table.Cell(row, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    Table.Cell(row, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    Table.Cell(row, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    Table.Cell(row, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                    Table.Cell(row, 5).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    Table.Cell(row, 6).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    Table.Rows[row].Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop;
                    Table.Borders.Enable = 0;
                    // заполнение таблицы
                    Table.Rows.Add();
                    row++;
                    // пояснительная записка
                    Table.Rows.Add();
                    Table.Cell(row, 1).Range.Text = NumberDocument.ToString();
                    Table.Cell(row, 3).Range.Text = "Пояснительная записка" + "\n";
                    row++;
                    // шапка локальных смет
                    Table.Rows.Add();
                    Table.Cell(row, 3).Range.Text = "ЛОКАЛЬНЫЕ СМЕТЫ" + "\n";
                    // изменение параметров строки
                    Table.Cell(row, 3).Range.Font.Bold = 1;
                    Table.Cell(row, 3).Range.Font.Size = 14;
                    Table.Cell(row, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    row++;
                    // вывод локальных смет
                    foreach (var lData in localData)
                    {
                        NumberDocument++;
                        Table.Rows.Add();
                        Table.Cell(row, 1).Range.Text = NumberDocument.ToString();
                        Table.Cell(row, 2).Range.Text = lData.Code;
                        Table.Cell(row, 3).Range.Text = lData.NameDate + "\n";
                        Table.Cell(row, 4).Range.Text = lData.Price;
                        // изменение параметров строки
                        Table.Cell(row, 3).Range.Font.Size = 10;
                        Table.Cell(row, 3).Range.Font.Italic = 0;
                        Table.Cell(row, 3).Range.Font.Bold = 0;
                        Table.Cell(row, 3).Range.Font.Underline = Word.WdUnderline.wdUnderlineNone;
                        Table.Cell(row, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                        Table.Cell(row, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                        Table.Cell(row, 4).Range.Font.Size = 10;
                        row++;

                    }




                    //нумерация страниц
                    int pagesInTitle = wDocument.ComputeStatistics(WdStatistic.wdStatisticPages, false) - 1;
                    int pageNumber = (int)StartNumberNumeric.Value + pagesInTitle;
                    row = 2;
                    if (TwoSidedPrintCheckBox.Checked)
                    {
                        // добавление страниц после содержания TODO

                        // нумерация ПЗ
                        if ((pageNumber % 2) == 0)
                        {
                            pageNumber += 1;
                        }
                        else
                        {
                            pageNumber += 2;
                        }
                        Table.Cell(row, 5).Range.Text = pageNumber.ToString();
                        pageNumber += (int)CountPagePZNumeric.Value - 1;
                        row++;

                        // нумерация сметы
                        if ((pageNumber % 2) == 0)
                        {
                            pageNumber += 1;
                        }
                        else
                        {
                            pageNumber += 2;
                        }

                        foreach (var data in objectiveData) // объектные сметы
                        {
                            Table.Cell(row, 5).Range.Text = pageNumber.ToString();
                            pageNumber += data.PageCount;
                            row++;
                        }


                        row++;



                        foreach (var lData in localData)
                        {

                            Table.Cell(row, 5).Range.Text = pageNumber.ToString();
                            pageNumber += lData.PageCount;
                            row++;

                        }



                    }
                    else
                    {
                        // добавление страниц после содержания TODO

                        // нумерация ПЗ
                        pageNumber++;
                        Table.Cell(row, 5).Range.Text = pageNumber.ToString();
                        pageNumber += (int)CountPagePZNumeric.Value - 1;
                        row++;

                        // нумерация сметы
                        foreach (var data in objectiveData) // объектные сметы
                        {
                            Table.Cell(row, 5).Range.Text = pageNumber.ToString();
                            pageNumber += data.PageCount;
                            row++;
                        }
                        row++;



                        foreach (var lData in localData)
                        {

                            Table.Cell(row, 5).Range.Text = pageNumber.ToString();
                            pageNumber += lData.PageCount;
                            row++;

                        }

                    }

                    wDocument.SaveAs2($"{pdfFolder}\\Содержание.docx");
                    wDocument.ExportAsFixedFormat($"{pdfFolder}\\Содержание.pdf", Word.WdExportFormat.wdExportFormatPDF);
                    wDocument.Close(Word.WdSaveOptions.wdDoNotSaveChanges, Word.WdOriginalFormat.wdOriginalDocumentFormat, false);
                }
                return true;
            }
            catch (Exception ex)
            {
                DeleteTempFiles();
                DeleteTempVar();
                MessageBox.Show("Ошибка генерации содержания");
                MessageBox.Show(ex.Message.ToString());
                backgroundWorker.CancelAsync();
                backgroundWorker.ReportProgress(1, "Сборка остановлена...");

                return false;
            }
            finally
            {
                wordApp.Quit();
            }


        }
        protected void RunBackgroundWorker_DoWork()
        {
            backgroundWorker.ReportProgress(1, "Сборка начата...");
            stopWatch.Start(); //Запуск секундомера (Время сборки)

            if (!ExcelParser()) return;
            if (!ExcelConverter()) return;
            if (!TitleGeneration()) return;
            if (!CreateDesktopFolder()) return;
            if (!PdfMerge()) return;
            if (!MoveFiles()) return;

            DeleteTempFiles();
            DeleteTempVar();

            stopWatch.Stop(); //Остановка секундомера
            TimeSpan ts = stopWatch.Elapsed;
            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);
            string Time = $"Время сборки: {elapsedTime}";
            backgroundWorker.ReportProgress(1, Time);
        }

        protected void PageBreaker(Excel.Worksheet eWorksheet)
        {
            // разделение (разрыв) страниц
            try
            {
                int lastUsedRow = eWorksheet.Cells.Find("*", System.Reflection.Missing.Value,
                       System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                       Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                       false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
                var a = eWorksheet.HPageBreaks;
                eWorksheet.ResetAllPageBreaks();

                for (int p = 1; p <= a.Count; p++)
                {
                    int i = a[p].Location.Row;
                    a.Add(eWorksheet.Range[$"A{i}"]);
                }
                int lastPageBreak = a[a.Count].Location.Row;
                if (lastUsedRow - lastPageBreak < 13)
                {
                    a[a.Count].Delete();
                    a.Add(eWorksheet.Range[$"A{lastUsedRow - 13}"]);
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
                Console.WriteLine(ex.StackTrace);
                Console.WriteLine(ex.Message.ToString());
            }
        }

        protected int FullBookPageCounter //Счетчик общего количества страниц
        {
            get
            {
                int numberPagesInBooks = 0;
                Excel.Application app = new Excel.Application { DisplayAlerts = false, Visible = false, ScreenUpdating = false };
                Workbook eWorkbook;
                try
                {
                    infoTextBox.Text = "Идет подсчет страниц...";
                    for (int i = 0; i < objectiveFiles.Length; i++)
                    {
                        eWorkbook = app.Workbooks.Open($"{childFolder}\\{objectiveFiles[i]}");
                        numberPagesInBooks += eWorkbook.Sheets[1].PageSetup.Pages.Count;
                        eWorkbook.Close();
                    }
                    for (int j = 0; j < localFiles.Length; j++)
                    {
                        eWorkbook = app.Workbooks.Open($"{rootFolder}\\{localFiles[j]}");
                        numberPagesInBooks += eWorkbook.Sheets[1].PageSetup.Pages.Count;
                        eWorkbook.Close();
                    }
                    return numberPagesInBooks;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка!");
                    Console.WriteLine(ex.StackTrace);
                    Console.WriteLine(ex.Message.ToString());
                    DeleteTempFiles();
                    DeleteTempVar();
                    return 0;
                }
                finally
                {
                    app.Quit();
                    eWorkbook = null;
                    GC.Collect();
                }
            }
        }

        protected bool CreateDesktopFolder()
        {
            if (!System.IO.Directory.Exists(DesktopFolder))
            {
                System.IO.Directory.CreateDirectory(DesktopFolder);
                return true;
            }
            return false;
        }

        protected bool MoveFiles()
        {
            try
            {
                if (SplitBookContentCheckBox.Checked)
                {
                    File.Move($@"{_path}\TEMPdf\Содержание.pdf", $@"{DesktopFolder}\Содержание.pdf");
                    if(!partsBookCheckBox.Checked)
                        File.Move($@"{_path}\TEMPdf\Сметы.pdf", $@"{DesktopFolder}\Сметы.pdf");
                    File.Move($@"{_path}\TEMPdf\Содержание.docx", $@"{DesktopFolder}\Содержание.docx");
                }
                else
                {
                    File.Move($@"{_path}\TEMPdf\smetaBook.pdf", $@"{DesktopFolder}\smetaBook.pdf");
                }
                return true;
            }
            catch (Exception)
            {
                DeleteTempFiles();
                DeleteTempVar();
                MessageBox.Show("Ошибка перемещения файлов на рабочий стол");
                backgroundWorker.CancelAsync();
                backgroundWorker.ReportProgress(1, "Сборка остановлена...");
                return false;
            }
        }

        protected void DeleteTempFiles()
        {
            if (Directory.Exists($"{_path}\\TEMPdf"))
            {
                Directory.Delete($"{_path}\\TEMPdf", true);
            }
        }

        protected void DeleteTempVar()
        {
            _path = null;
            dir = null;
            pdfFolder = null;
            rootFolder = null;
            localFiles = null;
            childFolder = null;
            objectiveFiles = null;
            localData = new List<SmetaFile>();
            objectiveData = new List<SmetaFile>();
            GC.Collect();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Word.Application wordApp = new Word.Application
            {

                Visible = true,
                ScreenUpdating = true
            };


            object oMissing = Type.Missing;
            Object defaultTableBehavior = Word.WdDefaultTableBehavior.wdWord9TableBehavior;
            Object autoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitWindow;
            

            var wDocument = wordApp.Documents.Add();
            

            // настройка полей документа
            wDocument.PageSetup.TopMargin = wordApp.InchesToPoints(0.5f);
            wDocument.PageSetup.BottomMargin = wordApp.InchesToPoints(0.5f);
            wDocument.PageSetup.LeftMargin = wordApp.InchesToPoints(0.65f);
            wDocument.PageSetup.RightMargin = wordApp.InchesToPoints(0.4f);
            wDocument.PageSetup.HeaderDistance = 20f;

            wDocument.Sections[1].PageSetup.OddAndEvenPagesHeaderFooter = -1; // -1 = true  - настройка: четные-нечетные страницы

            Word.Range headerRange = wDocument.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;

            wDocument.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
            wDocument.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.RestartNumberingAtSection = true;
            wDocument.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.StartingNumber = (int)StartNumberNumeric.Value; // номер первой страницы

            // колонтитул нечетной страницы
            wDocument.Tables.Add(headerRange, 1, 6, ref defaultTableBehavior, ref autoFitBehavior);
            Word.Table headerTable = headerRange.Tables[1];

            headerTable.Borders.Enable = 0;
            Word.Range rangePageNum = headerTable.Range.Cells[headerTable.Range.Cells.Count].Range;
            rangePageNum.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            Word.Field fld = rangePageNum.Document.Fields.Add(rangePageNum, oMissing, "Page", false);
            Word.Range rangeFieldPageNum = fld.Result;
            rangeFieldPageNum.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            headerTable.Cell(1, 6).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            headerTable.Cell(1, 6).Range.Font.Size = 12;

            headerTable.Rows.Add();
            headerTable.Cell(2, 3).Range.Text = "Содержание";
            headerTable.Cell(2, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            headerTable.Cell(2, 3).Range.Font.Name = "Times New Roman";
            headerTable.Cell(2, 3).Range.Font.Size = 12;
            headerTable.Cell(2, 3).Range.Font.Italic = 1;
            headerTable.Cell(2, 3).Range.Font.Bold = 1;
            headerTable.Cell(2, 3).Range.Font.Color = Word.WdColor.wdColorBlack;

            // заполнение таблицы
            headerTable.Rows.Add();
            headerTable.Rows[3].Borders.Enable = 1;
            headerTable.Cell(3, 1).Range.Text = "N п/п";
            headerTable.Cell(3, 2).Range.Text = "N сметы";
            headerTable.Cell(3, 3).Range.Text = "Наименование";
            headerTable.Cell(3, 4).Range.Text = "Всего тыс.руб.";
            headerTable.Cell(3, 5).Range.Text = "Стр.";
            headerTable.Cell(3, 6).Range.Text = "Часть";
            // изменение параметров таблицы
            headerTable.Rows[3].Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            headerTable.Rows[3].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            headerTable.Rows[3].Range.Font.Name = "Times New Roman";
            headerTable.Rows[3].Range.Font.Italic = 0;
            headerTable.Rows[3].Range.Font.Bold = 1;
            headerTable.Rows[3].Range.Font.Size = 10;
            headerTable.Rows[3].Range.Font.Color = Word.WdColor.wdColorBlack;
            // ширина ячеек таблицы
            headerTable.Columns[1].PreferredWidth = 6f;
            headerTable.Columns[2].PreferredWidth = 9f;
            headerTable.Columns[3].PreferredWidth = 30f;
            headerTable.Columns[4].PreferredWidth = 9f;
            headerTable.Columns[5].PreferredWidth = 4f;
            headerTable.Columns[6].PreferredWidth = 4f;

            // колонтитул четных страниц
            headerRange = wDocument.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages].Range;

            wDocument.Tables.Add(headerRange, 1, 6, ref defaultTableBehavior, ref autoFitBehavior);
            headerTable = headerRange.Tables[1];

            headerTable.Borders.Enable = 0;
            rangePageNum = headerTable.Range.Cells[1].Range;
            rangePageNum.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            fld = rangePageNum.Document.Fields.Add(rangePageNum, oMissing, "Page", false);
            rangeFieldPageNum = fld.Result;
            rangeFieldPageNum.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            
            headerTable.Cell(1, 6).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            headerTable.Cell(1, 6).Range.Font.Size = 12;

            headerTable.Rows.Add();
            headerTable.Cell(2, 3).Range.Text = "Содержание";
            headerTable.Cell(2, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            headerTable.Cell(2, 3).Range.Font.Name = "Times New Roman";
            headerTable.Cell(2, 3).Range.Font.Size = 12;
            headerTable.Cell(2, 3).Range.Font.Italic = 1;
            headerTable.Cell(2, 3).Range.Font.Bold = 1;
            headerTable.Cell(2, 3).Range.Font.Color = Word.WdColor.wdColorBlack;

            // заполнение таблицы
            headerTable.Rows.Add();
            headerTable.Rows[3].Borders.Enable = 1;
            headerTable.Cell(3, 1).Range.Text = "N п/п";
            headerTable.Cell(3, 2).Range.Text = "N сметы";
            headerTable.Cell(3, 3).Range.Text = "Наименование";
            headerTable.Cell(3, 4).Range.Text = "Всего тыс.руб.";
            headerTable.Cell(3, 5).Range.Text = "Стр.";
            headerTable.Cell(3, 6).Range.Text = "Часть";
            // изменение параметров таблицы
            headerTable.Rows[3].Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            headerTable.Rows[3].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            headerTable.Rows[3].Range.Font.Name = "Times New Roman";
            headerTable.Rows[3].Range.Font.Italic = 0;
            headerTable.Rows[3].Range.Font.Bold = 1;
            headerTable.Rows[3].Range.Font.Size = 10;
            headerTable.Rows[3].Range.Font.Color = Word.WdColor.wdColorBlack;
            // ширина ячеек таблицы
            headerTable.Columns[1].PreferredWidth = 6f;
            headerTable.Columns[2].PreferredWidth = 9f;
            headerTable.Columns[3].PreferredWidth = 30f;
            headerTable.Columns[4].PreferredWidth = 9f;
            headerTable.Columns[5].PreferredWidth = 4f;
            headerTable.Columns[6].PreferredWidth = 4f;
        }

    }
}