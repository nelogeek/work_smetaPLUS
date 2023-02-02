using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
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
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using iTextSharp.text;

using System.Diagnostics;


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

        //int documentNumber = 1;
        //int countPages;

        Stopwatch stopWatch = new Stopwatch();

        string DesktopFolder = $@"{Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)}\Книга смет";

        public Form1()
        {
            InitializeComponent();
            backgroundWorker.WorkerReportsProgress = true;
            backgroundWorker.WorkerSupportsCancellation = true;
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
            backgroundWorker.ReportProgress(1, "Сборка начата...");
            stopWatch.Start(); //Запуск секундомера для отчета о времени выполнения программы
            BackgroundWorker worker = sender as BackgroundWorker;
            //DisableButton();
            if (localFiles != null)
            {
                if (CreateDesktopFolder())
                {
                    //countPages = (int)StartNumberNumeric.Value;
                    ExcelParser();
                    ExcelConverter();
                    TitleGeneration();
                    PdfMerge();
                    MoveFiles();
                }
                else
                {
                    MessageBox.Show("Папка 'Книга смет' уже существует");
                }
                
            }
            else
            {
                MessageBox.Show($"Ошибка! Вы не выбрали папку");
            }
            //EnabledButton();
            stopWatch.Stop(); //Остановка секундомера
            TimeSpan ts = stopWatch.Elapsed;
            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);
            string Time = $"Время сборки: {elapsedTime}";

            backgroundWorker.ReportProgress(1, Time);
        }
        protected void DisableButton()
        {
            this.StartNumberNumeric.Enabled = false;
            this.numericUpDown1.Enabled = false;
            this.numericUpDown2.Enabled = false;
            this.CountPagePZNumeric.Enabled = false;
            this.btnBuild.Enabled = false;
            this.btnSelectFolder.Enabled = false;
            this.TwoSidedPrintCheckBox.Enabled = false;
        }
        protected void EnabledButton()
        {
            this.StartNumberNumeric.Enabled = true;
            this.numericUpDown1.Enabled = true;
            this.numericUpDown2.Enabled = true;
            this.CountPagePZNumeric.Enabled = true;
            this.btnBuild.Enabled = true;
            this.btnSelectFolder.Enabled = true;
            this.TwoSidedPrintCheckBox.Enabled = true;
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
                    foreach (var file in localFiles)
                    {
                        string fileName = file.Name;

                        Regex regex = new Regex(@".*", RegexOptions.RightToLeft);
                        MatchCollection match = regex.Matches(fileName);
                        string fileNameStr = match[0].ToString();
                        string[] fileType = fileNameStr.Split('.');
                        if (fileType[fileType.Length - 1] != "xlsx")
                        {
                            MessageBox.Show($"В папке находится недопустимый файл");
                        }
                    }
                    dir = Directory.GetDirectories(_path);
                    if (dir.Length != 1)
                    {
                        MessageBox.Show("В сметах должна быть только одна папка, которая должна содержать объектные сметы!");
                        labelNameFolder.Text = "Добавьте папку со объектными сметами\"ОС\"";
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

                        //string[] fileNames = new string[localFiles.Length + objectiveFiles.Length];

                        Directory.GetFiles(_path, ".", SearchOption.TopDirectoryOnly).ToList()
                            .ForEach(f => infoTextBox.AppendText($"\n- {Path.GetFileName(f)}" + Environment.NewLine));

                        infoTextBox.AppendText(Environment.NewLine + $"\n\nПапка {dir[0]}" + Environment.NewLine);

                        Directory.GetFiles(dir[0], ".", SearchOption.TopDirectoryOnly).ToList()
                            .ForEach(f => infoTextBox.AppendText(Environment.NewLine + $"\n- {Path.GetFileName(f)}"));
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
                // Start the asynchronous operation.
                backgroundWorker.RunWorkerAsync();
            }
        }


        private void ExcelParser()
        {

            Excel.Application app = new Excel.Application
            {
                DisplayAlerts = false
            };

            Excel.Workbook eWorkbook;
            Excel.Worksheet eWorksheet;
            int pages;

            for (int i = 0; i < objectiveFiles.Length; i++) /// шаблон для объектных смет
            {
                string filePath = $"{childFolder}\\{objectiveFiles[i]}";
                eWorkbook = app.Workbooks.Open($@"{filePath}");
                eWorksheet = (Excel.Worksheet)eWorkbook.Sheets[1];

                Regex regex = new Regex(@"(\w*)-(\w*)-(\w*)");
                MatchCollection match = regex.Matches(eWorksheet.Range["B10"].Value.ToString());

                pages = eWorkbook.Sheets[1].PageSetup.Pages.Count; /// кол-во страниц на листе
                //countPages += pages;

                objectiveData.Add(new SmetaFile(
                    match[0].ToString(), // код сметы
                    eWorksheet.Range["B7"].Value.ToString(), // Наименование
                    eWorksheet.Range["F14"].Value.ToString(), // Сумма денег
                    eWorkbook.Sheets[1].PageSetup.Pages.Count, // кол-во страниц на листе
                    objectiveFiles[i],
                    match[0].ToString().Substring(3)));

                eWorkbook.Close(false);
                //documentNumber++;

            }

            for (int i = 0; i < localFiles.Length; i++) // шаблон для локальных смет
            {
                string filePath = $"{rootFolder}\\{localFiles[i]}";
                eWorkbook = app.Workbooks.Open($@"{filePath}");
                eWorksheet = (Excel.Worksheet)eWorkbook.Sheets[1];

                Regex regex = new Regex(@"(\w*)-(\w*)-(\w*)");
                MatchCollection match = regex.Matches(eWorksheet.Range["A18"].Value.ToString());

                pages = eWorkbook.Sheets[1].PageSetup.Pages.Count;
                //countPages += pages; // кол-во страниц на листе


                localData.Add(new SmetaFile(
                    match[0].ToString(), // код сметы
                    eWorksheet.Range["A20"].Value.ToString(), // Наименование
                    eWorksheet.Range["D28"].Value.ToString(), // Сумма денег
                    eWorkbook.Sheets[1].PageSetup.Pages.Count, // кол-во страниц на листе
                    localFiles[i],
                    match[0].ToString().Substring(3)));

                eWorkbook.Close(false);
                //documentNumber++;
            }



            localData = localData.OrderBy(x => x.Code).ThenBy(x => x.Name).ToList(); // Сортировка по коду и названию
            objectiveData = objectiveData.OrderBy(x => x.Code).ThenBy(x => x.Name).ToList(); // Сортировка по коду и названию


            app.Quit();
            eWorkbook = null;
            eWorksheet = null;
            pages = 0;
            GC.Collect();
        }

        protected void ExcelConverter()
        {
            Excel.Application app = new Excel.Application
            {
                DisplayAlerts = false
            };

            Excel.Workbook eWorkbook;


            /// конвертер Excel to PDF
            int countCompleted = 0;
            Directory.CreateDirectory($"{_path}\\TEMPdf");
            foreach (var file in objectiveData)
            {
                string filePath = $"{_path}\\ОС\\{file.FolderInfo}";
                eWorkbook = app.Workbooks.Open(filePath);
                string tempPDFPath = $"{_path}\\TEMPdf\\{file.FolderInfo}";
                eWorkbook.Sheets[1].PageSetup.RightFooter = ""; ///Удаление нумерации станиц в Excel
                app.ActiveWorkbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, tempPDFPath);
                eWorkbook.Close(false);
                countCompleted++;

                //labelCompleted.Text = $"Кол-во обработанных файлов: {countCompleted} из {localFiles.Length + objectiveFiles.Length}";

            }
            foreach (var file in localData)
            {
                string filePath = $"{_path}\\{file.FolderInfo}";
                eWorkbook = app.Workbooks.Open(filePath);
                string tempPDFPath = $"{_path}\\TEMPdf\\{file.FolderInfo}";
                eWorkbook.Sheets[1].PageSetup.RightFooter = ""; ///Удаление нумерации станиц в Excel
                app.ActiveWorkbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, tempPDFPath);
                eWorkbook.Close(false);
                countCompleted++;
                //labelCompleted.Text = $"Кол-во обработанных файлов: {countCompleted} из {localFiles.Length + objectiveFiles.Length}";
            }

            app.Quit();
            eWorkbook = null;
            GC.Collect();
        }



        protected void PdfMerge()
        {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string fileNameConcatPdf = $"{desktopPath}\\smetaBook.pdf";
            string fileNameSmetaPdf = $"{DesktopFolder}\\Сметы.pdf";
            string fileNameTitlePdf = $"{DesktopFolder}\\Содержание.pdf";

            List<SmetaFile> tempFilesArray = objectiveData;
            tempFilesArray.AddRange(localData);

            //Объединение PDF
            PdfDocument outputPdfDocument = new PdfDocument();
            int countCompleted = 0;

            // add title
            PdfDocument inputPdfDocument = PdfReader.Open(fileNameTitlePdf, PdfDocumentOpenMode.Import);
            for (int i = 0; i < inputPdfDocument.PageCount; i++)
            {
                PdfPage page = inputPdfDocument.Pages[i];
                outputPdfDocument.AddPage(page);
            }
            countCompleted++;
            //labelCompleted.Text = $"Кол-во обработанных файлов: {countCompleted} из {localFiles.Length + objectiveFiles.Length + 1}";

            if(SplitBookContentCheckBox.Checked)
            {
                PdfDocument outputSmetaPdfDocument = new PdfDocument();
                foreach (var file in tempFilesArray)
                {
                    inputPdfDocument = PdfReader.Open($"{pdfFolder}\\{file.FolderInfo}.pdf", PdfDocumentOpenMode.Import);

                    for (int i = 0; i < inputPdfDocument.PageCount; i++)
                    {
                        PdfPage page = inputPdfDocument.Pages[i];
                        outputPdfDocument.AddPage(page);
                        outputSmetaPdfDocument.AddPage(page);
                    }
                    countCompleted++;
                    //labelCompleted.Text = $"Кол-во обработанных файлов: {countCompleted} из {localFiles.Length + objectiveFiles.Length + 1}";
                }
                outputSmetaPdfDocument.Save(fileNameSmetaPdf);
                outputPdfDocument.Save(fileNameConcatPdf);
                outputSmetaPdfDocument.Close();
                outputPdfDocument.Close();
            }
            else
            {
                foreach (var file in tempFilesArray)
                {
                    inputPdfDocument = PdfReader.Open($"{pdfFolder}\\{file.FolderInfo}.pdf", PdfDocumentOpenMode.Import);

                    for (int i = 0; i < inputPdfDocument.PageCount; i++)
                    {
                        PdfPage page = inputPdfDocument.Pages[i];
                        outputPdfDocument.AddPage(page);
                    }
                    countCompleted++;
                    //labelCompleted.Text = $"Кол-во обработанных файлов: {countCompleted} из {localFiles.Length + objectiveFiles.Length + 1}";
                }
            }
            //Добавление правильной нумерации страниц книги по отдельности
            if (SplitBookContentCheckBox.Checked)
            {
                AddPageNumberTitleITextSharp(fileNameTitlePdf);
                AddPageNumberSmetaITextSharp(fileNameSmetaPdf);
            }
            //Добавление правильной нумерации страниц собранной книги
            AddPageNumberITextSharp(fileNameConcatPdf);
        }

        protected void AddPageNumberTitleITextSharp(string fileTitlePath)
        {
            byte[] bytesTitle = File.ReadAllBytes(fileTitlePath);

            Font blackFont = FontFactory.GetFont("Arial", 8, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            using (MemoryStream stream = new MemoryStream())
            {
                iTextSharp.text.pdf.PdfReader readerTitle = new iTextSharp.text.pdf.PdfReader(bytesTitle);
                int pages = readerTitle.NumberOfPages;

                using (iTextSharp.text.pdf.PdfStamper stamper = new iTextSharp.text.pdf.PdfStamper(readerTitle, stream))
                {
                    int startPageNumber = Convert.ToInt32(StartNumberNumeric.Value) - 1;
                    int pagesPzCount = Convert.ToInt32(CountPagePZNumeric.Value);

                    //Нумерация страниц содержания
                    for (int i = 1; i <= pages; i++)
                    {
                        iTextSharp.text.pdf.ColumnText.ShowTextAligned(stamper.GetUnderContent(i), Element.ALIGN_RIGHT, new Phrase((i + startPageNumber).ToString(), blackFont), 565f, 15f, 0);
                    }
                       
                }
                bytesTitle = stream.ToArray();
                readerTitle.Close();
            }
            File.WriteAllBytes(fileTitlePath, bytesTitle);
        }

        protected void AddPageNumberSmetaITextSharp(string filePath)
        {
            byte[] bytes = File.ReadAllBytes(filePath);
            byte[] bytesTitle = File.ReadAllBytes($"{DesktopFolder}\\Содержание.pdf");

            Font blackFont = FontFactory.GetFont("Arial", 8, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            using (MemoryStream stream = new MemoryStream())
            {
                iTextSharp.text.pdf.PdfReader reader = new iTextSharp.text.pdf.PdfReader(bytes);
                iTextSharp.text.pdf.PdfReader readerOnlyTitle = new iTextSharp.text.pdf.PdfReader(bytesTitle);
                int titlePages = readerOnlyTitle.NumberOfPages;
                int pages = reader.NumberOfPages;

                using (iTextSharp.text.pdf.PdfStamper stamper = new iTextSharp.text.pdf.PdfStamper(reader, stream))
                {
                    int startPageNumber = Convert.ToInt32(StartNumberNumeric.Value) - 1;
                    int pagesPzCount = Convert.ToInt32(CountPagePZNumeric.Value);

                    if (TwoSidedPrintCheckBox.Checked)
                    {
                        for (int i = 1; i <= pages; i++)
                        {
                            if (i % 2 == 0)
                            {
                                iTextSharp.text.pdf.ColumnText.ShowTextAligned(stamper.GetUnderContent(i), Element.ALIGN_RIGHT, new Phrase((i + startPageNumber + pagesPzCount + titlePages).ToString(), blackFont), 810f, 575f, 0);
                            }
                            else
                            {
                                iTextSharp.text.pdf.ColumnText.ShowTextAligned(stamper.GetUnderContent(i), Element.ALIGN_RIGHT, new Phrase((i + startPageNumber + pagesPzCount + titlePages).ToString(), blackFont), 810f, 15f, 0);
                            }
                        }
                    }
                    else
                    {
                        for (int i = titlePages + 1; i <= pages; i++)
                        {
                            iTextSharp.text.pdf.ColumnText.ShowTextAligned(stamper.GetUnderContent(i), Element.ALIGN_RIGHT, new Phrase((i + startPageNumber + pagesPzCount).ToString(), blackFont), 810f, 15f, 0);
                        }
                    }
                }
                bytes = stream.ToArray();
                reader.Close();

            }
            File.WriteAllBytes(filePath, bytes);
        }

        protected void AddPageNumberITextSharp(string filePath)
        {
            byte[] bytes = File.ReadAllBytes(filePath);
            byte[] bytesTitle = File.ReadAllBytes($"{DesktopFolder}\\Содержание.pdf");

            iTextSharp.text.Font blackFont = FontFactory.GetFont("Arial", 8, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            using (MemoryStream stream = new MemoryStream())
            {
                iTextSharp.text.pdf.PdfReader reader = new iTextSharp.text.pdf.PdfReader(bytes);
                iTextSharp.text.pdf.PdfReader readerOnlyTitle = new iTextSharp.text.pdf.PdfReader(bytesTitle);
                int titlePages = readerOnlyTitle.NumberOfPages;
                int pages = reader.NumberOfPages;

                using (iTextSharp.text.pdf.PdfStamper stamper = new iTextSharp.text.pdf.PdfStamper(reader, stream))
                {
                    int startPageNumber = Convert.ToInt32(StartNumberNumeric.Value) - 1;
                    int pagesPzCount = Convert.ToInt32(CountPagePZNumeric.Value);

                    if (TwoSidedPrintCheckBox.Checked)
                    {
                        //Нумерация страниц содержания
                        for (int i = 1; i <= titlePages; i++)
                        {
                            iTextSharp.text.pdf.ColumnText.ShowTextAligned(stamper.GetUnderContent(i), Element.ALIGN_RIGHT, new Phrase((i + startPageNumber).ToString(), blackFont), 565f, 15f, 0);
                        }
                        for (int i = 1 + titlePages; i <= pages; i++)
                        {
                            if (i % 2 == 0)
                            {
                                iTextSharp.text.pdf.ColumnText.ShowTextAligned(stamper.GetUnderContent(i), Element.ALIGN_RIGHT, new Phrase((i + startPageNumber + pagesPzCount).ToString(), blackFont), 810f, 575f, 0);
                            }
                            else
                            {
                                iTextSharp.text.pdf.ColumnText.ShowTextAligned(stamper.GetUnderContent(i), Element.ALIGN_RIGHT, new Phrase((i + startPageNumber + pagesPzCount).ToString(), blackFont), 810f, 15f, 0);
                            }
                        }
                    }
                    else
                    {
                        //Нумерация страниц содержания
                        for (int i = 1; i <= titlePages; i++)
                        {
                            iTextSharp.text.pdf.ColumnText.ShowTextAligned(stamper.GetUnderContent(i), Element.ALIGN_RIGHT, new Phrase((i + startPageNumber).ToString(), blackFont), 565f, 15f, 0);
                        }
                        for (int i = titlePages + 1; i <= pages; i++)
                        {
                            iTextSharp.text.pdf.ColumnText.ShowTextAligned(stamper.GetUnderContent(i), Element.ALIGN_RIGHT, new Phrase((i + startPageNumber + pagesPzCount).ToString(), blackFont), 810f, 15f, 0);
                        }
                    }
                }
                bytes = stream.ToArray();
                reader.Close();

            }
            File.WriteAllBytes(filePath, bytes);
        }



        protected void TitleGeneration()
        {

            // ---------------- Генерация содержания ----------------------------------------------------------------------------

            int countPages = (int)StartNumberNumeric.Value;

            if (localFiles != null)
            {
                Word.Application app = new Word.Application
                {
                    Visible = false
                    // или app.Visible = false;
                };
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
                wTable1.Cell(row, 5).Range.Text = countPages.ToString();
                // изменение параметров строки
                wTable1.Cell(row, 1).Range.Font.Size = 9;
                wTable1.Cell(row, 3).Range.Font.Size = 10;
                wTable1.Rows[row].Range.Font.Bold = 0;
                wTable1.Rows[row].Range.Font.Color = Word.WdColor.wdColorBlack;
                wTable1.Cell(row, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                wTable1.Rows[row].Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop;

                //---
                row++;
                countPages += (int)CountPagePZNumeric.Value;

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
                    wTable1.Cell(row, 3).Range.Text = data.Name + "\n";
                    wTable1.Cell(row, 4).Range.Text = data.Price;
                    wTable1.Cell(row, 5).Range.Text = countPages.ToString();
                    // изменение параметров строки
                    wTable1.Cell(row, 1).Range.Font.Size = 9;
                    wTable1.Cell(row, 2).Range.Font.Size = 10;
                    wTable1.Cell(row, 3).Range.Font.Bold = 0;
                    wTable1.Cell(row, 3).Range.Font.Size = 10;
                    wTable1.Cell(row, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    wTable1.Cell(row, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                    //---
                    countPages += data.PageCount;
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
                var setDict = pairs.GroupBy(x => x.Value.Trim()).Select(y => y.FirstOrDefault());

                // вывод локальных смет
                foreach (var oData in setDict)
                {
                    wTable1.Rows.Add();
                    // вывод объектной сметы
                    wTable1.Cell(row, 3).Range.Text = oData.Value + "\n";
                    // изменение параметров строки
                    wTable1.Cell(row, 3).Range.Font.Bold = 0;
                    wTable1.Cell(row, 3).Range.Font.Size = 10;
                    wTable1.Cell(row, 3).Range.Font.Italic = 1;
                    wTable1.Cell(row, 3).Range.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
                    //---
                    row++;

                    // вывод соответствующих локальных смет
                    foreach (var lData in localData)
                    {
                        if (lData.Code.Contains(oData.Key))
                        {
                            NumberDocument++;
                            wTable1.Rows.Add();
                            wTable1.Cell(row, 1).Range.Text = NumberDocument.ToString();
                            wTable1.Cell(row, 2).Range.Text = lData.Code;
                            wTable1.Cell(row, 3).Range.Text = lData.Name + "\n";
                            wTable1.Cell(row, 4).Range.Text = lData.Price;
                            wTable1.Cell(row, 5).Range.Text = countPages.ToString();
                            // изменение параметров строки
                            wTable1.Cell(row, 3).Range.Font.Size = 10;
                            wTable1.Cell(row, 3).Range.Font.Italic = 0;
                            wTable1.Cell(row, 3).Range.Font.Bold = 0;
                            wTable1.Cell(row, 3).Range.Font.Underline = Word.WdUnderline.wdUnderlineNone;
                            wTable1.Cell(row, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                            wTable1.Cell(row, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                            countPages += lData.PageCount;
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
                // ---
                //wDocument.SaveAs2($"{pdfFolder}\\Содержание.pdf", WdSaveFormat.wdFormatPDF);
                wDocument.ExportAsFixedFormat($"{pdfFolder}\\Содержание.pdf", Word.WdExportFormat.wdExportFormatPDF);
                wDocument.Close(Word.WdSaveOptions.wdDoNotSaveChanges, Word.WdOriginalFormat.wdOriginalDocumentFormat, false);
                //app.ActiveDocument.SaveAs2($@"{_path}\TEST.docx");
                app.Quit();
            }
            else
            {
                MessageBox.Show($"Ошибка! Вы не выбрали папку");
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

        protected void MoveFiles()
        {
            File.Move($@"{_path}\TEMPdf\Содержание.pdf", $@"{DesktopFolder}\Содержание.pdf");
        }


    }
}