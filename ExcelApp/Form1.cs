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
using System.Text.RegularExpressions;
using System.Windows.Forms;
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

        Stopwatch stopWatch = new Stopwatch();

        string DesktopFolder = $@"{Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)}\Книга смет";

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

                        infoTextBox.AppendText($"Общее количество страниц: {fullBookPageCounter()}" + Environment.NewLine +
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

        private void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
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
            }
        }

        private void backgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            labelCompleted.Text = e.UserState.ToString();
        }

        private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            if (_path != null)
            {
                if (Directory.Exists($"{DesktopFolder}"))
                {
                    DialogResult dialogResult = MessageBox.Show("Вы точно хотите заменить папку 'Книга смет'?", "Подтверждение замены папки", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        Directory.Delete(DesktopFolder, true);

                        runBackgroundWorker_DoWork();
                    }
                    else if (dialogResult == DialogResult.No)
                    {
                        backgroundWorker.ReportProgress(1, "Сборка остановлена...");
                        return;
                    }
                }
                else
                    runBackgroundWorker_DoWork();
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
            //this.numericUpDown1.Enabled = false;
            this.afterTitleNumeric.Enabled = false;
            this.CountPagePZNumeric.Enabled = false;
            this.btnBuild.Enabled = false;
            this.btnSelectFolder.Enabled = false;
            this.TwoSidedPrintCheckBox.Enabled = false;
            this.SplitBookContentCheckBox.Enabled = false;
        }
        protected void EnabledButton()
        {
            this.StartNumberNumeric.Enabled = true;
            //this.numericUpDown1.Enabled = true;
            this.afterTitleNumeric.Enabled = true;
            this.CountPagePZNumeric.Enabled = true;
            this.btnBuild.Enabled = true;
            this.btnSelectFolder.Enabled = true;
            this.TwoSidedPrintCheckBox.Enabled = true;
            this.SplitBookContentCheckBox.Enabled = true;
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
                    eWorksheet.PageSetup.Orientation = XlPageOrientation.xlLandscape; // TODO
                    Regex regex = new Regex(@"(\w*)-(\d*)-(\d*)");
                    string code = regex.Matches(eWorksheet.Range["E8"].Value.ToString())[0].ToString() + "p"; // TODO переделать
                    string ShortCode = code.Substring(3,5);
                    string money = eWorksheet.Range["G12"].Value.ToString();
                    string nameDate = eWorksheet.Range["C5"].Value.ToString();
                    string date = eWorksheet.Range["C18"].Value.ToString().Split(new string[] { " цен " }, StringSplitOptions.None)[1];
                    nameDate += $"\n(в ценах на {date})";

                    int pages = eWorkbook.Sheets[1].PageSetup.Pages.Count; /// кол-во страниц на листе

                    objectiveData.Add(new SmetaFile(
                        code, // код сметы
                        eWorksheet.Range["C5"].Value.ToString(),
                        nameDate, // Наименование
                        money, // Сумма денег
                        pages, // кол-во страниц на листе
                        objectiveFiles[i],
                        ShortCode));

                    pageBreaker();

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
                    string shortCode = regex.Matches(match[0].Value.ToString())[0].ToString(); // TODO сделать шорт-код

                    string money = eWorksheet.Range["C28"].Value.ToString().Replace("(", "").Replace(")", "");
                    if (money == "0")
                        money = eWorksheet.Range["D28"].Value.ToString().Replace("(", "").Replace(")", "");

                    string nameDate = eWorksheet.Range["A20"].Value.ToString();
                    string date = eWorksheet.Range["D26"].Value.ToString();
                    nameDate += $"\n(в ценах на {date})";

                    int pages = eWorksheet.PageSetup.Pages.Count; /// кол-во страниц на листе

                    pageBreaker();

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

                List<SmetaFile> tempFilesArray = objectiveData;
                tempFilesArray.AddRange(localData);

                //Объединение PDF
                PdfDocument inputPdfDocument;

                if (SplitBookContentCheckBox.Checked)
                {
                    PdfDocument outputSmetaPdfDocument = new PdfDocument();
                    foreach (var file in tempFilesArray)
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
                    foreach (var file in tempFilesArray)
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
                
                //Нумерация страниц
                if (SplitBookContentCheckBox.Checked)
                {
                    AddPageNumberTitleITextSharp(fileNameTitlePdf);
                    AddPageNumberSmetaITextSharp(fileNameSmetaPdf);
                } else
                {
                    AddPageNumberITextSharp(fileNameConcatPdf);
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
                        if(TwoSidedPrintCheckBox.Checked)
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
                    int afterTitleNumericPages = Convert.ToInt32(afterTitleNumeric.Value);

                    using (iTextSharp.text.pdf.PdfStamper stamper = new iTextSharp.text.pdf.PdfStamper(reader, stream))
                    {
                        int startPageNumber = Convert.ToInt32(StartNumberNumeric.Value) - 1;
                        int pagesPzCount = Convert.ToInt32(CountPagePZNumeric.Value);

                        if (TwoSidedPrintCheckBox.Checked)
                        {
                            if((startPageNumber + titlePages) % 2 == 1)
                                titlePages++;
                            if (pagesPzCount % 2 == 1)
                                pagesPzCount++;

                            for (int i = 1; i <= pagesBook; i++)
                            {
                                if ((startPageNumber + titlePages + pagesPzCount + i) % 2 == 0)
                                    iTextSharp.text.pdf.ColumnText.ShowTextAligned(stamper.GetUnderContent(i), Element.ALIGN_RIGHT, new Phrase((i + afterTitleNumericPages + startPageNumber + pagesPzCount + titlePages).ToString(), blackFont), 810f, 575f, 0);
                                else
                                    iTextSharp.text.pdf.ColumnText.ShowTextAligned(stamper.GetUnderContent(i), Element.ALIGN_RIGHT, new Phrase((i + afterTitleNumericPages + startPageNumber + pagesPzCount + titlePages).ToString(), blackFont), 810f, 15f, 0);
                            }
                        }
                        else
                            for (int i = 1; i <= pagesBook; i++)
                                iTextSharp.text.pdf.ColumnText.ShowTextAligned(stamper.GetUnderContent(i), Element.ALIGN_RIGHT, new Phrase((i + afterTitleNumericPages + startPageNumber + pagesPzCount + titlePages).ToString(), blackFont), 810f, 15f, 0);
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
                    int pages = reader.NumberOfPages;
                    int afterTitleNumericPages = Convert.ToInt32(afterTitleNumeric.Value);

                    using (iTextSharp.text.pdf.PdfStamper stamper = new iTextSharp.text.pdf.PdfStamper(reader, stream))
                    {
                        int startPageNumber = Convert.ToInt32(StartNumberNumeric.Value) - 1;
                        int pagesPzCount = Convert.ToInt32(CountPagePZNumeric.Value);

                        if (TwoSidedPrintCheckBox.Checked)
                        {
                            bool flag = true;
                            //Нумерация страниц содержания
                            for (int i = 1; i <= titlePages; i++)
                                iTextSharp.text.pdf.ColumnText.ShowTextAligned(stamper.GetUnderContent(i), Element.ALIGN_RIGHT, new Phrase((i + startPageNumber).ToString(), blackFont), 565f, 15f, 0);
                            for (int i = 1 + titlePages; i <= pages; i++)
                            {
                                if (flag)
                                {
                                    iTextSharp.text.pdf.ColumnText.ShowTextAligned(stamper.GetUnderContent(i), Element.ALIGN_RIGHT, new Phrase((i + afterTitleNumericPages + startPageNumber + pagesPzCount).ToString(), blackFont), 810f, 15f, 0);
                                    flag = false;
                                }
                                else
                                {
                                    iTextSharp.text.pdf.ColumnText.ShowTextAligned(stamper.GetUnderContent(i), Element.ALIGN_RIGHT, new Phrase((i + afterTitleNumericPages + startPageNumber + pagesPzCount).ToString(), blackFont), 810f, 575f, 0);
                                    flag = true;
                                }
                            }
                        }
                        else
                        {
                            //Нумерация страниц содержания
                            for (int i = 1; i <= titlePages; i++)
                                iTextSharp.text.pdf.ColumnText.ShowTextAligned(stamper.GetUnderContent(i), Element.ALIGN_RIGHT, new Phrase((i + startPageNumber).ToString(), blackFont), 565f, 15f, 0);
                            for (int i = titlePages + 1; i <= pages; i++)
                                iTextSharp.text.pdf.ColumnText.ShowTextAligned(stamper.GetUnderContent(i), Element.ALIGN_RIGHT, new Phrase((i + afterTitleNumericPages + startPageNumber + pagesPzCount).ToString(), blackFont), 810f, 15f, 0);
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
            Word.Application app = new Word.Application
            {
                Visible = false,
                ScreenUpdating = false
            };

            try
            {
                if (objectiveData.Count != 0)
                {
                    var wDocument = app.Documents.Add();

                    object oMissing = Type.Missing;
                    Object defaultTableBehavior = Word.WdDefaultTableBehavior.wdWord9TableBehavior;
                    Object autoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitWindow;

                    wDocument.Paragraphs.Add(ref oMissing);
                    wDocument.Paragraphs.Add(ref oMissing);

                    var wParagraph = wDocument.Paragraphs[1];
                    var wRange = wParagraph.Range;
                    wParagraph.SpaceAfter = 0; // межстрочный интервал
                    wRange.Text = "Содержание";
                    wRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wRange.Font.Name = "Times New Roman";
                    wRange.Font.Size = 12;
                    wRange.Font.Italic = 1;
                    wRange.Font.Bold = 1;
                    wRange.Font.Color = Word.WdColor.wdColorGray55;

                    wParagraph = wDocument.Paragraphs[2];
                    wRange = wParagraph.Range;
                    var wTable1 = wDocument.Tables.Add(wRange, 1, 6, ref defaultTableBehavior, ref autoFitBehavior);

                    int row = 1;

                    // Шапка
                    wTable1.Cell(row, 1).Range.Text = "N п/п";
                    wTable1.Cell(row, 2).Range.Text = "N сметы";
                    wTable1.Cell(row, 3).Range.Text = "Наименование";
                    wTable1.Cell(row, 4).Range.Text = "Всего тыс.руб.";
                    wTable1.Cell(row, 5).Range.Text = "Стр.";
                    wTable1.Cell(row, 6).Range.Text = "Часть";
                    // изменение параметров шапки
                    wTable1.Range.Rows[row].Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    wTable1.Range.Rows[row].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wTable1.Range.Rows[row].Range.Font.Name = "Times New Roman";
                    wTable1.Range.Rows[row].Cells.Borders.Enable = 0; // отключение границ в содержании
                    wTable1.Rows[row].Range.Font.Color = Word.WdColor.wdColorGray55;
                    wTable1.Rows[row].Range.Font.Bold = 1;
                    row++;



                    wTable1.Rows.Add();
                    row++;

                    int NumberDocument = 1; // номер документа

                    // пояснительная записка
                    wTable1.Rows.Add();
                    wTable1.Cell(row, 1).Range.Text = NumberDocument.ToString();
                    wTable1.Cell(row, 3).Range.Text = "Пояснительная записка" + "\n";
                    // изменение параметров строки
                    wTable1.Cell(row, 1).Range.Font.Size = 9;
                    wTable1.Cell(row, 3).Range.Font.Size = 10;
                    wTable1.Rows[row].Range.Font.Bold = 0;
                    wTable1.Rows[row].Range.Font.Color = Word.WdColor.wdColorBlack;
                    wTable1.Cell(row, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    wTable1.Rows[row].Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop;

                    //---
                    row++;


                    // шапка объектной сметы
                    wTable1.Rows.Add();
                    wTable1.Cell(row, 3).Range.Text = "ОБЪЕКТНЫЕ СМЕТЫ" + "\n";
                    // изменение параметров строки
                    wTable1.Cell(row, 3).Range.Font.Bold = 1;
                    wTable1.Cell(row, 3).Range.Font.Size = 14;
                    wTable1.Cell(row, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    row++;

                    // вывод объектной сметы
                    foreach (var data in objectiveData)
                    {
                        NumberDocument++;
                        wTable1.Rows.Add();
                        wTable1.Cell(row, 1).Range.Text = NumberDocument.ToString();
                        wTable1.Cell(row, 2).Range.Text = data.Code;
                        wTable1.Cell(row, 3).Range.Text = data.NameDate + "\n";
                        wTable1.Cell(row, 4).Range.Text = data.Price;

                        // изменение параметров строки
                        wTable1.Cell(row, 1).Range.Font.Size = 9;
                        wTable1.Cell(row, 2).Range.Font.Size = 10;
                        wTable1.Cell(row, 3).Range.Font.Bold = 0;
                        wTable1.Cell(row, 3).Range.Font.Size = 10;
                        wTable1.Cell(row, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                        wTable1.Cell(row, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                        wTable1.Cell(row, 4).Range.Font.Size = 10;
                        //---

                        row++;
                    }

                    // шапка локальных смет
                    wTable1.Rows.Add();
                    wTable1.Cell(row, 3).Range.Text = "ЛОКАЛЬНЫЕ СМЕТЫ" + "\n";
                    // изменение параметров строки
                    wTable1.Cell(row, 3).Range.Font.Bold = 1;
                    wTable1.Cell(row, 3).Range.Font.Size = 14;
                    wTable1.Cell(row, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    row++;

                    // удаление повторяющихся номеров объектных смет
                    List<Pair> pairs = new List<Pair>();
                    foreach (var data in objectiveData)
                    {
                        pairs.Add(new Pair() { Key = data.ShortCode, Value = data.Name });
                    }
                    var setDict = pairs.GroupBy(x => x.Key.Trim()).Select(y => y.FirstOrDefault());

                    // вывод локальных смет
                    foreach (var oData in setDict)
                    {
                        //MessageBox.Show(oData.Value.ToString());
                        wTable1.Rows.Add();
                        // вывод объектной сметы
                        wTable1.Cell(row, 3).Range.Text = oData.Value + "\n";
                        // изменение параметров строки
                        wTable1.Cell(row, 3).Range.Font.Bold = 0;
                        wTable1.Cell(row, 3).Range.Font.Size = 10;
                        wTable1.Cell(row, 3).Range.Font.Italic = 1;
                        wTable1.Cell(row, 3).Range.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
                        wTable1.Cell(row, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                        //---
                        row++;



                        // вывод соответствующих локальных смет
                        foreach (var lData in localData)
                        {
                            if (lData.ShortCode == oData.Key)
                            {
                                NumberDocument++;
                                wTable1.Rows.Add();
                                wTable1.Cell(row, 1).Range.Text = NumberDocument.ToString();
                                wTable1.Cell(row, 2).Range.Text = lData.Code;
                                wTable1.Cell(row, 3).Range.Text = lData.NameDate + "\n";
                                wTable1.Cell(row, 4).Range.Text = lData.Price;

                                // изменение параметров строки
                                wTable1.Cell(row, 3).Range.Font.Size = 10;
                                wTable1.Cell(row, 3).Range.Font.Italic = 0;
                                wTable1.Cell(row, 3).Range.Font.Bold = 0;
                                wTable1.Cell(row, 3).Range.Font.Underline = Word.WdUnderline.wdUnderlineNone;
                                wTable1.Cell(row, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                                wTable1.Cell(row, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                                wTable1.Cell(row, 4).Range.Font.Size = 10;
                                row++;
                            }
                        }
                    }



                    wTable1.Columns[1].PreferredWidth = 6f;
                    wTable1.Columns[2].PreferredWidth = 9f;
                    wTable1.Columns[3].PreferredWidth = 30f;
                    wTable1.Columns[4].PreferredWidth = 9f;
                    wTable1.Columns[5].PreferredWidth = 8f;
                    wTable1.Columns[6].PreferredWidth = 5f;

                    wTable1.Range.Rows[1].Cells.Borders.Enable = 1; // добавление границ в шапке
                    wTable1.Rows[1].Cells[1].Borders[Word.WdBorderType.wdBorderTop].Color = Word.WdColor.wdColorGray30;
                    wTable1.Rows[1].Cells[1].Borders[Word.WdBorderType.wdBorderRight].Color = Word.WdColor.wdColorGray30;
                    wTable1.Rows[1].Cells[1].Borders[Word.WdBorderType.wdBorderBottom].Color = Word.WdColor.wdColorGray30;
                    wTable1.Rows[1].Cells[1].Borders[Word.WdBorderType.wdBorderLeft].Color = Word.WdColor.wdColorGray30;

                    wTable1.Rows[1].Cells[2].Borders[Word.WdBorderType.wdBorderTop].Color = Word.WdColor.wdColorGray30;
                    wTable1.Rows[1].Cells[2].Borders[Word.WdBorderType.wdBorderRight].Color = Word.WdColor.wdColorGray30;
                    wTable1.Rows[1].Cells[2].Borders[Word.WdBorderType.wdBorderBottom].Color = Word.WdColor.wdColorGray30;
                    wTable1.Rows[1].Cells[2].Borders[Word.WdBorderType.wdBorderLeft].Color = Word.WdColor.wdColorGray30;

                    wTable1.Rows[1].Cells[3].Borders[Word.WdBorderType.wdBorderTop].Color = Word.WdColor.wdColorGray30;
                    wTable1.Rows[1].Cells[3].Borders[Word.WdBorderType.wdBorderRight].Color = Word.WdColor.wdColorGray30;
                    wTable1.Rows[1].Cells[3].Borders[Word.WdBorderType.wdBorderBottom].Color = Word.WdColor.wdColorGray30;
                    wTable1.Rows[1].Cells[3].Borders[Word.WdBorderType.wdBorderLeft].Color = Word.WdColor.wdColorGray30;

                    wTable1.Rows[1].Cells[4].Borders[Word.WdBorderType.wdBorderTop].Color = Word.WdColor.wdColorGray30;
                    wTable1.Rows[1].Cells[4].Borders[Word.WdBorderType.wdBorderRight].Color = Word.WdColor.wdColorGray30;
                    wTable1.Rows[1].Cells[4].Borders[Word.WdBorderType.wdBorderBottom].Color = Word.WdColor.wdColorGray30;
                    wTable1.Rows[1].Cells[4].Borders[Word.WdBorderType.wdBorderLeft].Color = Word.WdColor.wdColorGray30;

                    wTable1.Rows[1].Cells[5].Borders[Word.WdBorderType.wdBorderTop].Color = Word.WdColor.wdColorGray30;
                    wTable1.Rows[1].Cells[5].Borders[Word.WdBorderType.wdBorderRight].Color = Word.WdColor.wdColorGray30;
                    wTable1.Rows[1].Cells[5].Borders[Word.WdBorderType.wdBorderBottom].Color = Word.WdColor.wdColorGray30;
                    wTable1.Rows[1].Cells[5].Borders[Word.WdBorderType.wdBorderLeft].Color = Word.WdColor.wdColorGray30;

                    wTable1.Rows[1].Cells[6].Borders[Word.WdBorderType.wdBorderTop].Color = Word.WdColor.wdColorGray30;
                    wTable1.Rows[1].Cells[6].Borders[Word.WdBorderType.wdBorderRight].Color = Word.WdColor.wdColorGray30;
                    wTable1.Rows[1].Cells[6].Borders[Word.WdBorderType.wdBorderBottom].Color = Word.WdColor.wdColorGray30;
                    wTable1.Rows[1].Cells[6].Borders[Word.WdBorderType.wdBorderLeft].Color = Word.WdColor.wdColorGray30;

                    //нумерация страниц

                    int pagesInTitle = wDocument.ComputeStatistics(WdStatistic.wdStatisticPages, false);
                    //MessageBox.Show(pagesInTitle.ToString());
                    int countPages = (int)StartNumberNumeric.Value + (int)afterTitleNumeric.Value + (int)CountPagePZNumeric.Value + pagesInTitle - 2; // TODO  2 переделать

                    row = 3;

                    wTable1.Cell(row, 5).Range.Text = countPages.ToString();
                    countPages += (int)CountPagePZNumeric.Value;

                    row += 2;

                    foreach (var data in objectiveData) // объектные сметы
                    {
                        wTable1.Cell(row, 5).Range.Text = countPages.ToString();
                        countPages += data.PageCount;
                        row++;
                    }



                    //foreach (var lData in localData) 
                    //{
                    //    wTable1.Cell(row, 5).Range.Text = countPages.ToString();
                    //    countPages += lData.PageCount;
                    //    row++;
                    //}

                    row++;

                    foreach (var oData in setDict) // локальные сметы
                    {
                        row++;

                        foreach (var lData in localData)
                        {
                            if (lData.ShortCode == oData.Key)
                            {
                                wTable1.Cell(row, 5).Range.Text = countPages.ToString();
                                countPages += lData.PageCount;
                                row++;
                            }
                        }
                    }

                    //------------------


                    // ---
                    wDocument.SaveAs2($"{pdfFolder}\\Содержание.docx");
                    wDocument.ExportAsFixedFormat($"{pdfFolder}\\Содержание.pdf", Word.WdExportFormat.wdExportFormatPDF);
                    wDocument.Close(Word.WdSaveOptions.wdDoNotSaveChanges, Word.WdOriginalFormat.wdOriginalDocumentFormat, false);
                    //app.ActiveDocument.SaveAs2($@"{_path}\TEST.docx");
                    app.Quit();
                }
                else
                {

                    var wDocument = app.Documents.Add();

                    object oMissing = Type.Missing;
                    Object defaultTableBehavior = Word.WdDefaultTableBehavior.wdWord9TableBehavior;
                    Object autoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitWindow;

                    wDocument.Paragraphs.Add(ref oMissing);
                    wDocument.Paragraphs.Add(ref oMissing);

                    var wParagraph = wDocument.Paragraphs[1];
                    var wRange = wParagraph.Range;
                    wParagraph.SpaceAfter = 0; // межстрочный интервал
                    wRange.Text = "Содержание";
                    wRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wRange.Font.Name = "Times New Roman";
                    wRange.Font.Size = 12;
                    wRange.Font.Italic = 1;
                    wRange.Font.Bold = 1;
                    wRange.Font.Color = Word.WdColor.wdColorGray55;

                    wParagraph = wDocument.Paragraphs[2];
                    wRange = wParagraph.Range;
                    var wTable1 = wDocument.Tables.Add(wRange, 1, 6, ref defaultTableBehavior, ref autoFitBehavior);

                    int row = 1;

                    // Шапка
                    wTable1.Cell(row, 1).Range.Text = "N п/п";
                    wTable1.Cell(row, 2).Range.Text = "N сметы";
                    wTable1.Cell(row, 3).Range.Text = "Наименование";
                    wTable1.Cell(row, 4).Range.Text = "Всего тыс.руб.";
                    wTable1.Cell(row, 5).Range.Text = "Стр.";
                    wTable1.Cell(row, 6).Range.Text = "Часть";
                    // изменение параметров шапки
                    wTable1.Range.Rows[row].Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    wTable1.Range.Rows[row].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wTable1.Range.Rows[row].Range.Font.Name = "Times New Roman";
                    wTable1.Range.Rows[row].Cells.Borders.Enable = 0; // отключение границ в содержании
                    wTable1.Rows[row].Range.Font.Color = Word.WdColor.wdColorGray55;
                    wTable1.Rows[row].Range.Font.Bold = 1;
                    row++;



                    wTable1.Rows.Add();
                    row++;

                    int NumberDocument = 1; // номер документа

                    // пояснительная записка
                    wTable1.Rows.Add();
                    wTable1.Cell(row, 1).Range.Text = NumberDocument.ToString();
                    wTable1.Cell(row, 3).Range.Text = "Пояснительная записка" + "\n";

                    // изменение параметров строки
                    wTable1.Cell(row, 1).Range.Font.Size = 9;
                    wTable1.Cell(row, 3).Range.Font.Size = 10;
                    wTable1.Rows[row].Range.Font.Bold = 0;
                    wTable1.Rows[row].Range.Font.Color = Word.WdColor.wdColorBlack;
                    wTable1.Cell(row, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    wTable1.Rows[row].Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop;

                    //---
                    row++;




                    // шапка локальных смет
                    wTable1.Rows.Add();
                    wTable1.Cell(row, 3).Range.Text = "ЛОКАЛЬНЫЕ СМЕТЫ" + "\n";
                    // изменение параметров строки
                    wTable1.Cell(row, 3).Range.Font.Bold = 1;
                    wTable1.Cell(row, 3).Range.Font.Size = 14;
                    wTable1.Cell(row, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    row++;


                    // вывод соответствующих локальных смет
                    foreach (var lData in localData)
                    {

                        NumberDocument++;
                        wTable1.Rows.Add();
                        wTable1.Cell(row, 1).Range.Text = NumberDocument.ToString();
                        wTable1.Cell(row, 2).Range.Text = lData.Code;
                        wTable1.Cell(row, 3).Range.Text = lData.NameDate + "\n";
                        wTable1.Cell(row, 4).Range.Text = lData.Price;

                        // изменение параметров строки
                        wTable1.Cell(row, 3).Range.Font.Size = 10;
                        wTable1.Cell(row, 3).Range.Font.Italic = 0;
                        wTable1.Cell(row, 3).Range.Font.Bold = 0;
                        wTable1.Cell(row, 3).Range.Font.Underline = Word.WdUnderline.wdUnderlineNone;
                        wTable1.Cell(row, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                        wTable1.Cell(row, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                        wTable1.Cell(row, 4).Range.Font.Size = 10;

                        row++;

                    }

                    wTable1.Columns[1].PreferredWidth = 6f;
                    wTable1.Columns[2].PreferredWidth = 9f;
                    wTable1.Columns[3].PreferredWidth = 30f;
                    wTable1.Columns[4].PreferredWidth = 9f;
                    wTable1.Columns[5].PreferredWidth = 8f;
                    wTable1.Columns[6].PreferredWidth = 5f;

                    wTable1.Range.Rows[1].Cells.Borders.Enable = 1; // добавление границ в шапке
                    wTable1.Rows[1].Cells[1].Borders[Word.WdBorderType.wdBorderTop].Color = Word.WdColor.wdColorGray30;
                    wTable1.Rows[1].Cells[1].Borders[Word.WdBorderType.wdBorderRight].Color = Word.WdColor.wdColorGray30;
                    wTable1.Rows[1].Cells[1].Borders[Word.WdBorderType.wdBorderBottom].Color = Word.WdColor.wdColorGray30;
                    wTable1.Rows[1].Cells[1].Borders[Word.WdBorderType.wdBorderLeft].Color = Word.WdColor.wdColorGray30;

                    wTable1.Rows[1].Cells[2].Borders[Word.WdBorderType.wdBorderTop].Color = Word.WdColor.wdColorGray30;
                    wTable1.Rows[1].Cells[2].Borders[Word.WdBorderType.wdBorderRight].Color = Word.WdColor.wdColorGray30;
                    wTable1.Rows[1].Cells[2].Borders[Word.WdBorderType.wdBorderBottom].Color = Word.WdColor.wdColorGray30;
                    wTable1.Rows[1].Cells[2].Borders[Word.WdBorderType.wdBorderLeft].Color = Word.WdColor.wdColorGray30;

                    wTable1.Rows[1].Cells[3].Borders[Word.WdBorderType.wdBorderTop].Color = Word.WdColor.wdColorGray30;
                    wTable1.Rows[1].Cells[3].Borders[Word.WdBorderType.wdBorderRight].Color = Word.WdColor.wdColorGray30;
                    wTable1.Rows[1].Cells[3].Borders[Word.WdBorderType.wdBorderBottom].Color = Word.WdColor.wdColorGray30;
                    wTable1.Rows[1].Cells[3].Borders[Word.WdBorderType.wdBorderLeft].Color = Word.WdColor.wdColorGray30;

                    wTable1.Rows[1].Cells[4].Borders[Word.WdBorderType.wdBorderTop].Color = Word.WdColor.wdColorGray30;
                    wTable1.Rows[1].Cells[4].Borders[Word.WdBorderType.wdBorderRight].Color = Word.WdColor.wdColorGray30;
                    wTable1.Rows[1].Cells[4].Borders[Word.WdBorderType.wdBorderBottom].Color = Word.WdColor.wdColorGray30;
                    wTable1.Rows[1].Cells[4].Borders[Word.WdBorderType.wdBorderLeft].Color = Word.WdColor.wdColorGray30;

                    wTable1.Rows[1].Cells[5].Borders[Word.WdBorderType.wdBorderTop].Color = Word.WdColor.wdColorGray30;
                    wTable1.Rows[1].Cells[5].Borders[Word.WdBorderType.wdBorderRight].Color = Word.WdColor.wdColorGray30;
                    wTable1.Rows[1].Cells[5].Borders[Word.WdBorderType.wdBorderBottom].Color = Word.WdColor.wdColorGray30;
                    wTable1.Rows[1].Cells[5].Borders[Word.WdBorderType.wdBorderLeft].Color = Word.WdColor.wdColorGray30;

                    wTable1.Rows[1].Cells[6].Borders[Word.WdBorderType.wdBorderTop].Color = Word.WdColor.wdColorGray30;
                    wTable1.Rows[1].Cells[6].Borders[Word.WdBorderType.wdBorderRight].Color = Word.WdColor.wdColorGray30;
                    wTable1.Rows[1].Cells[6].Borders[Word.WdBorderType.wdBorderBottom].Color = Word.WdColor.wdColorGray30;
                    wTable1.Rows[1].Cells[6].Borders[Word.WdBorderType.wdBorderLeft].Color = Word.WdColor.wdColorGray30;


                    //нумерация страниц

                    int pagesInTitle = wDocument.ComputeStatistics(WdStatistic.wdStatisticPages, false);
                    //MessageBox.Show(pagesInTitle.ToString());
                    int countPages = (int)StartNumberNumeric.Value + (int)afterTitleNumeric.Value + (int)CountPagePZNumeric.Value + pagesInTitle - 1;

                    row = 3;

                    wTable1.Cell(row, 5).Range.Text = countPages.ToString();
                    countPages += (int)CountPagePZNumeric.Value;

                    row += 2;

                    foreach (var lData in localData) // локальные сметы
                    {
                        wTable1.Cell(row, 5).Range.Text = countPages.ToString();
                        countPages += lData.PageCount;
                        row++;
                    }

                    //------------------



                    // ---
                    //wDocument.SaveAs2($"{pdfFolder}\\Содержание.pdf", WdSaveFormat.wdFormatPDF);
                    wDocument.ExportAsFixedFormat($"{pdfFolder}\\Содержание.pdf", Word.WdExportFormat.wdExportFormatPDF);
                    wDocument.Close(Word.WdSaveOptions.wdDoNotSaveChanges, Word.WdOriginalFormat.wdOriginalDocumentFormat, false);

                    app.Quit();
                }
                return true;
            }
            catch (Exception)
            {
                DeleteTempFiles();
                DeleteTempVar();
                MessageBox.Show("Ошибка генерации содержания");
                backgroundWorker.CancelAsync();
                backgroundWorker.ReportProgress(1, "Сборка остановлена...");
                app.Quit();
                return false;
            }
        }

        protected void runBackgroundWorker_DoWork()
        {
            backgroundWorker.ReportProgress(1, "Сборка начата...");
            stopWatch.Start(); //Запуск секундомера (Время сборки)

            if (!ExcelParser()) return;
            if (!ExcelConverter()) return;
            if (!TitleGeneration()) return;
            if (!PdfMerge()) return;
            if (!CreateDesktopFolder()) return;
            if (!MoveFiles()) return;

            DeleteTempFiles();
            DeleteTempVar();

            stopWatch.Stop(); //Остановка секундомера
            TimeSpan ts = stopWatch.Elapsed;
            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);
            string Time = $"Время сборки: {elapsedTime}";
            backgroundWorker.ReportProgress(1, Time);
        }

        protected void pageBreaker() //Регулировка разрывов страниц
        {

            //var lastUsedRow = eWorksheet.Cells.Find("*", System.Reflection.Missing.Value,
            //           System.Reflection.Missing.Value, System.Reflection.Missing.Value,
            //           XlSearchOrder.xlByRows, XlSearchDirection.xlPrevious,
            //           false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
            //const int rowsCount = 45; // кол-во строк на странице
            //if ((lastUsedRow % rowsCount) < 13)
            //{
            //    eWorksheet.ResetAllPageBreaks();

            //    int nubmerPageBreaks = lastUsedRow / rowsCount;

            //    for (int ind = 1; ind < nubmerPageBreaks; ind++)
            //    {
            //        eWorksheet.HPageBreaks.Add(eWorksheet.Range[$"A{rowsCount * ind}"]);
            //    }
            //    eWorksheet.HPageBreaks.Add(eWorksheet.Range[$"A{lastUsedRow - 20}"]);
            //}
            //if(lastUsedRow > 35)
            //{
            //    eWorksheet.HPageBreaks.Add(eWorksheet.Range[$"A{lastUsedRow - 20}"]);
            //}
        }

        protected int fullBookPageCounter() //Счетчик общего количества страниц
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
                app.Quit();
                eWorkbook = null;
                GC.Collect();
                return numberPagesInBooks;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка!");
                Console.WriteLine(ex.StackTrace);
                Console.WriteLine(ex.Message.ToString());
                DeleteTempFiles();
                DeleteTempVar();
                app.Quit();
                eWorkbook = null;
                GC.Collect();
                return 0;
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
            localData = new List<SmetaFile>(); ;
            objectiveData = new List<SmetaFile>(); ;
        }
    }
}