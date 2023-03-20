using ExcelAPP;
using iTextSharp.text;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.ComponentModel;
using System.Diagnostics;

namespace ExcelApp.Functions
{
    partial class ProgramFunctions
    {
        public ProgramFunctions()
        {
            mf = MainForm.instance;
        }

        public MainForm mf;

        public string path;
        public string[] dirFolders;

        public DirectoryInfo pdfFolder;
        public DirectoryInfo finalSmetaFolder;
        public DirectoryInfo rootFolder;
        public FileInfo[] localFiles;

        public DirectoryInfo childFolder;
        public FileInfo[] objectiveFiles;

        public List<SmetaFile> localData = new List<SmetaFile>();
        public List<SmetaFile> objectiveData = new List<SmetaFile>();
        public readonly Stopwatch stopWatch = new Stopwatch();
        public int fullBookPageCount;

        public List<List<int>> firstPageNumbersList = new List<List<int>>();

        public int pagesInTitle = 0;
        public IEnumerable<Pair> setDict;
        public List<SmetaFile> allDataFilesList;
        
        public void SelectFolder() //Функция обработки выбора папки
        {
            try
            {
                mf.labelNameFolder.Text = path;

                if (Directory.Exists($"{path}\\ОС"))
                {
                    childFolder = new DirectoryInfo($@"{path}\ОС");
                    objectiveFiles = childFolder.GetFiles(".", SearchOption.TopDirectoryOnly); //Сбор объектных файлов файлов
                    foreach (var file in objectiveFiles) //Проверка расширения объектных файлов
                    {
                        Regex regex = new Regex(@".*", RegexOptions.RightToLeft);
                        MatchCollection match = regex.Matches(file.Name);
                        string fileNameStr = match[0].ToString();
                        string[] fileType = fileNameStr.Split('.');
                        if (fileType[fileType.Length - 1] != "xlsx" && fileType[fileType.Length - 1] != "xls")
                        {
                            MessageBox.Show($"В папке находится недопустимый файл");
                            return;
                        }
                    }
                }
                else
                {
                    childFolder = null;
                    objectiveFiles = null;
                }

                mf.infoTextBox.Clear();

                fullBookPageCount = FullBookPageCounter();
                mf.infoTextBox.Text = $"Общее количество страниц: {fullBookPageCount}";
                mf.infoTextBox.AppendText(
                    Environment.NewLine + $"Количество всех файлов: {localFiles.Length + objectiveFiles.Length}\n" +
                    Environment.NewLine + $"Количество папок: {dirFolders.Length}" + Environment.NewLine);

                if (childFolder != null)
                {
                    mf.infoTextBox.AppendText(Environment.NewLine +
                    $"Количество объектных файлов: {objectiveFiles.Length}\n" + Environment.NewLine);
                    mf.infoTextBox.AppendText(Environment.NewLine + $"Объектные файлы:" + Environment.NewLine);

                    Directory.GetFiles($"{path}\\ОС", ".", SearchOption.TopDirectoryOnly).ToList()
                    .ForEach(f => mf.infoTextBox.AppendText($"\n- {Path.GetFileName(f)}" + Environment.NewLine));
                }

                mf.infoTextBox.AppendText(Environment.NewLine + $"Количество локальных файлов: {localFiles.Length}\n" + Environment.NewLine);
                mf.infoTextBox.AppendText(Environment.NewLine + $"Локальные файлы:" + Environment.NewLine);

                Directory.GetFiles(path, ".", SearchOption.TopDirectoryOnly).ToList()
                    .ForEach(f => mf.infoTextBox.AppendText($"\n- {Path.GetFileName(f)}" + Environment.NewLine));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка функции выбора папки");
                Console.WriteLine(ex.Message.ToString());
                DeleteTempFiles();
                DeleteTempVar();
                GC.Collect();
                return;
            }
        }

        public int FullBookPageCounter() //Счетчик общего количества страниц
        {
            int numberPagesInBooks = 0;
            Excel.Application app = new Excel.Application { DisplayAlerts = false, Visible = false, ScreenUpdating = false };
            Workbook eWorkbook;
            try
            {
                mf.infoTextBox.Text = "Подсчёт страниц";
                if (childFolder != null)
                {
                    for (int i = 0; i < objectiveFiles.Length; i++)
                    {
                        eWorkbook = app.Workbooks.Open($"{childFolder}\\{objectiveFiles[i]}");
                        numberPagesInBooks += eWorkbook.Sheets[1].PageSetup.Pages.Count;
                        eWorkbook.Close();
                    }
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
                MessageBox.Show("Ошибка подсчета страниц");
                Console.WriteLine(ex.Message.ToString());
                mf.backgroundWorker.CancelAsync();
                app.Quit();
                eWorkbook = null;
                DeleteTempFiles();
                DeleteTempVar();
                GC.Collect();
                return 0;
            }
        }

        public void DisableButtons()
        {
            mf.StartNumberNumeric.Enabled = false;
            mf.CountPagePZNumeric.Enabled = false;
            mf.btnBuild.Enabled = false;
            mf.btnSelectFolder.Enabled = false;
            mf.TwoSidedPrintCheckBox.Enabled = false;
            mf.SplitBookContentCheckBox.Enabled = false;
            mf.RdPdToggle.Enabled = false;
            mf.settingsToolStripMenuItem.Enabled = false;
            mf.pagesInPartBookNumeric.Enabled = false;
            mf.partsBookCheckBox.Enabled = false;
            mf.dividerPassPagesCount.Enabled = false;
        }

        public void EnableButtons()
        {
            mf.StartNumberNumeric.Enabled = true;
            mf.CountPagePZNumeric.Enabled = true;
            mf.btnBuild.Enabled = true;
            mf.btnSelectFolder.Enabled = true;
            mf.TwoSidedPrintCheckBox.Enabled = true;
            mf.SplitBookContentCheckBox.Enabled = true;
            mf.RdPdToggle.Enabled = true;
            mf.settingsToolStripMenuItem.Enabled = true;
            mf.pagesInPartBookNumeric.Enabled = true;
            mf.partsBookCheckBox.Enabled = true;
            mf.dividerPassPagesCount.Enabled = true;
        }

        public bool ExcelParser() // Парсинг Excel файла
        {
            Excel.Application app = new Excel.Application { DisplayAlerts = false, Visible = false, ScreenUpdating = false };

            Workbook eWorkbook;
            Worksheet eWorksheet;
            try
            {
                if (childFolder != null) //Проверка наличия папки ОС
                {
                    for (int i = 0; i < objectiveFiles.Length; i++) //Шаблон для объектных смет
                    {
                        string filePath = $"{childFolder}\\{objectiveFiles[i]}";
                        eWorkbook = app.Workbooks.Open($@"{filePath}");
                        eWorksheet = (Worksheet)eWorkbook.Sheets[1];
                        eWorksheet.PageSetup.Orientation = XlPageOrientation.xlLandscape;
                        Regex regex = new Regex(@"(\w*)-(\w*)-(\w*)");
                        string code = regex.Matches(eWorksheet.Range["E8"].Value.ToString())[0].ToString();
                        string ShortCode = code.Replace("p", "").Replace("р", "").Replace("OC-", "").Replace("ОС-", "");
                        string money = eWorksheet.Range["G12"].Value.ToString();
                        string nameDate = eWorksheet.Range["C5"].Value.ToString();
                        string date = eWorksheet.Range["C18"].Value.ToString().Split(new string[] { " цен " }, StringSplitOptions.None)[1];
                        nameDate += $"\n(в ценах на {date})";

                        if (mf.RdPdToggle.Checked)
                        {
                            eWorksheet.Range["E8"].Replace("ОБЪЕКТНЫЙ СМЕТНЫЙ РАСЧЕТ (СМЕТА)", "ОБЪЕКТНАЯ СМЕТА");
                        }

                        int pages = eWorkbook.Sheets[1].PageSetup.Pages.Count; // кол-во страниц на листе

                        objectiveData.Add(new SmetaFile(
                            code, // код сметы
                            eWorksheet.Range["C5"].Value.ToString(), // наименование
                            nameDate, // Наименование
                            money, // Сумма денег
                            pages, // кол-во страниц на листе
                            objectiveFiles[i],
                            ShortCode));

                        if (mf.AutoPageBreakerToolStripMenuItem.Checked)
                        {
                            PageBreaker(eWorksheet);
                        }

                        money = null;
                        pages = 0;
                        nameDate = null;
                        date = null;
                        eWorkbook.Save();
                        eWorkbook.Close(true);
                    }
                }
                for (int j = 0; j < localFiles.Length; j++) //Шаблон для локальных смет
                {
                    string filePath = $"{rootFolder}\\{localFiles[j]}";
                    eWorkbook = app.Workbooks.Open($@"{filePath}");
                    eWorksheet = (Worksheet)eWorkbook.Sheets[1];
                    eWorksheet.PageSetup.Orientation = XlPageOrientation.xlLandscape;

                    Regex regex = new Regex(@"(\w*)-(\w*)-(\w*)");
                    MatchCollection match = regex.Matches(eWorksheet.Range["A18"].Value.ToString());

                    regex = new Regex(@"(\w*)-(\w*)");
                    string shortCode = regex.Matches(match[0].Value.ToString())[0].ToString();
                    string money = eWorksheet.Range["C28"].Value.ToString().Replace("(", "").Replace(")", "");
                    if (money == "0")
                    {
                        money = eWorksheet.Range["D28"].Value.ToString().Replace("(", "").Replace(")", "");
                    }
                    string nameDate = eWorksheet.Range["A20"].Value.ToString();
                    string date = eWorksheet.Range["D26"].Value.ToString();
                    nameDate += $"\n(в ценах на {date})";

                    if (mf.RdPdToggle.Checked)
                    {
                        eWorksheet.Range["A18"].Replace("ЛОКАЛЬНЫЙ СМЕТНЫЙ РАСЧЕТ (СМЕТА)", "ЛОКАЛЬНАЯ СМЕТА");
                    }

                    int pages = eWorksheet.PageSetup.Pages.Count; // кол-во страниц на листе

                    if (mf.AutoPageBreakerToolStripMenuItem.Checked)
                    {
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
                    eWorkbook.Close(true);
                }

                localData = localData.OrderBy(x => x.Code).ThenBy(x => x.NameDate).ToList(); // Сортировка по коду и названию
                objectiveData = objectiveData.OrderBy(x => x.Code).ThenBy(x => x.NameDate).ToList(); // Сортировка по коду и названию

                eWorkbook = null;
                eWorksheet = null;
                app.Quit();
                GC.Collect();

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при парсинге смет");
                Console.WriteLine(ex.Message.ToString());
                mf.backgroundWorker.CancelAsync();
                DeleteTempFiles();
                DeleteTempVar();

                eWorkbook = null;
                eWorksheet = null;
                app.Quit();
                GC.Collect();

                mf.backgroundWorker.ReportProgress(1, "Сборка остановлена");

                return false;
            }
        }

        public void PageBreaker(Excel.Worksheet eWorksheet) // Регулировка разрывов страниц
        {
            try
            {
                eWorksheet.Range[$"G7"].Value = "";
                eWorksheet.Rows[7].RowHeight = 11.25;
                int lastUsedRow = eWorksheet.Cells.Find("*", System.Reflection.Missing.Value,
                       System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                       XlSearchOrder.xlByRows, XlSearchDirection.xlPrevious,
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
                eWorksheet = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при регулировке разрывов страниц");
                Console.WriteLine(ex.StackTrace);
                Console.WriteLine(ex.Message.ToString());

                mf.backgroundWorker.CancelAsync();
                DeleteTempFiles();
                DeleteTempVar();

                GC.Collect();
            }
        }

        public bool ExcelConverter() //Конвертация Excel файлов в PDF
        {
            Excel.Application app = new Excel.Application { DisplayAlerts = false, Visible = false, ScreenUpdating = false };
            Workbook eWorkbook;
            Worksheet eWorksheet;

            try
            {
                Directory.CreateDirectory($"{path}\\TEMPdf");
                foreach (var file in objectiveData)
                {
                    string filePath = $"{path}\\ОС\\{file.FolderInfo}";
                    eWorkbook = app.Workbooks.Open(filePath);
                    eWorksheet = (Worksheet)eWorkbook.Sheets[1];
                    string tempPDFPath = $"{path}\\TEMPdf\\{file.FolderInfo}";
                    eWorksheet.PageSetup.RightFooter = ""; //Удаление нумерации станиц в Excel

                    app.ActiveWorkbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, tempPDFPath);
                    eWorkbook.Close(false);
                }
                foreach (var file in localData)
                {
                    string filePath = $"{path}\\{file.FolderInfo}";
                    string tempPDFPath = $"{path}\\TEMPdf\\{file.FolderInfo}";
                    eWorkbook = app.Workbooks.Open(filePath);
                    eWorksheet = (Worksheet)eWorkbook.Sheets[1];

                    eWorksheet.PageSetup.RightFooter = ""; //Удаление нумерации стpаниц в Excel
                    app.ActiveWorkbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, tempPDFPath);
                    eWorkbook.Close(false);
                }
                app.Quit();
                eWorkbook = null;
                eWorksheet = null;
                GC.Collect();

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка конвертации в PDF");
                MessageBox.Show(ex.Message.ToString());
                mf.backgroundWorker.CancelAsync();
                DeleteTempFiles();
                DeleteTempVar();

                app.Quit();
                eWorkbook = null;
                GC.Collect();

                mf.backgroundWorker.ReportProgress(10, "Сборка остановлена");

                return false;
            }
        }

        public bool CreateFinalSmetaFolder() //Создание финальной папки
        {
            try
            {
                if (!Directory.Exists(finalSmetaFolder.FullName))
                {
                    Directory.CreateDirectory(finalSmetaFolder.FullName);
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка создания финальной папки");
                MessageBox.Show(ex.Message.ToString());
                mf.backgroundWorker.CancelAsync();
                DeleteTempFiles();
                DeleteTempVar();

                GC.Collect();

                mf.backgroundWorker.ReportProgress(40, "Сборка остановлена");

                return false;
            }
        }

        public bool TitleGeneration()
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
                    wDocument.PageSetup.TopMargin = wordApp.InchesToPoints(0.4f);
                    wDocument.PageSetup.BottomMargin = wordApp.InchesToPoints(0.4f);
                    wDocument.PageSetup.LeftMargin = wordApp.InchesToPoints(0.4f);
                    wDocument.PageSetup.RightMargin = wordApp.InchesToPoints(0.4f);
                    wDocument.PageSetup.HeaderDistance = 20f;

                    if (mf.TwoSidedPrintCheckBox.Checked)
                    {
                        wDocument.Sections[1].PageSetup.OddAndEvenPagesHeaderFooter = -1; // -1 = true  -  настройка: четные-нечетные страницы

                        Word.Range headerRange = wDocument.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;

                        wDocument.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
                        wDocument.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.RestartNumberingAtSection = true;
                        wDocument.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.StartingNumber = (int)mf.StartNumberNumeric.Value; // номер первой страницы

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
                        headerTable.Cell(1, 6).Range.Font.Size = 10;

                        headerTable.Rows.Add();
                        headerTable.Cell(2, 3).Range.Text = "Содержание";
                        headerTable.Cell(2, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        headerTable.Cell(2, 3).Range.Font.Name = "Times New Roman";
                        headerTable.Cell(2, 3).Range.Font.Size = 12;
                        headerTable.Cell(2, 3).Range.Font.Italic = 1;
                        headerTable.Cell(2, 3).Range.Font.Bold = 1;
                        headerTable.Cell(2, 3).Range.Font.Color = Word.WdColor.wdColorBlack;
                        headerTable.Rows[2].Height = 0.93f;

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
                        headerTable.Columns[3].PreferredWidth = 32f;
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
                        headerTable.Cell(1, 1).Range.Font.Size = 10;

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
                        headerTable.Columns[3].PreferredWidth = 32f;
                        headerTable.Columns[4].PreferredWidth = 9f;
                        headerTable.Columns[5].PreferredWidth = 4f;
                        headerTable.Columns[6].PreferredWidth = 4f;
                    }
                    else
                    {

                        Word.HeaderFooter header = wDocument.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];

                        Word.Range headerRange = header.Range;

                        header.LinkToPrevious = false;
                        header.PageNumbers.RestartNumberingAtSection = true;
                        header.PageNumbers.StartingNumber = (int)mf.StartNumberNumeric.Value; // номер первой страницы

                        // колонтитул страницы
                        wDocument.Tables.Add(headerRange, 1, 6, ref defaultTableBehavior, ref autoFitBehavior);
                        Word.Table headerTable = headerRange.Tables[1];

                        headerTable.Borders.Enable = 0;
                        Word.Range rangePageNum = headerTable.Range.Cells[headerTable.Range.Cells.Count].Range;
                        rangePageNum.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                        Word.Field fld = rangePageNum.Document.Fields.Add(rangePageNum, oMissing, "Page", false);
                        Word.Range rangeFieldPageNum = fld.Result;
                        rangeFieldPageNum.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                        headerTable.Cell(1, 6).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                        headerTable.Cell(1, 6).Range.Font.Size = 10;

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
                        headerTable.Columns[3].PreferredWidth = 32f;
                        headerTable.Columns[4].PreferredWidth = 9f;
                        headerTable.Columns[5].PreferredWidth = 4f;
                        headerTable.Columns[6].PreferredWidth = 4f;
                    }

                    // основная часть
                    wDocument.Paragraphs.Add(ref oMissing);
                    var Paragraph = wDocument.Paragraphs[1];
                    var Range = Paragraph.Range;
                    var Table = wDocument.Tables.Add(Range, 1, 6, ref defaultTableBehavior, ref autoFitBehavior);
                    // ширина ячеек таблицы
                    Table.Columns[1].PreferredWidth = 6f;
                    Table.Columns[2].PreferredWidth = 9f;
                    Table.Columns[3].PreferredWidth = 32f;
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
                        if (mf.RdPdToggle.Checked)
                        {
                            Table.Cell(row, 2).Range.Text = $"{data.ShortCode}p"; //TODO иправить вывод кода объектных смет (буква 'p')
                        }
                        else
                        {
                            Table.Cell(row, 2).Range.Text = data.ShortCode;
                        }
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
                    setDict = pairs.GroupBy(x => x.Key.Trim()).Select(y => y.FirstOrDefault());

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
                    pagesInTitle = wDocument.ComputeStatistics(WdStatistic.wdStatisticPages, false); // кол-во страниц в содержании
                    int pageNumber = (int)mf.StartNumberNumeric.Value + pagesInTitle - 1; // номер страницы

                    row = 2;

                    if (mf.TwoSidedPrintCheckBox.Checked)
                    {
                        //TODO добавление страниц после содержания 

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
                        pageNumber += (int)mf.CountPagePZNumeric.Value - 1;
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
                        pageNumber += (int)mf.CountPagePZNumeric.Value - 1;
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
                    wDocument.PageSetup.TopMargin = wordApp.InchesToPoints(0.4f);
                    wDocument.PageSetup.BottomMargin = wordApp.InchesToPoints(0.4f);
                    wDocument.PageSetup.LeftMargin = wordApp.InchesToPoints(0.4f);
                    wDocument.PageSetup.RightMargin = wordApp.InchesToPoints(0.4f);
                    wDocument.PageSetup.HeaderDistance = 20f;

                    if (mf.TwoSidedPrintCheckBox.Checked)
                    {
                        wDocument.Sections[1].PageSetup.OddAndEvenPagesHeaderFooter = -1; // -1 = true  - настройка: четные-нечетные страницы

                        Word.Range headerRange = wDocument.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;

                        wDocument.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
                        wDocument.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.RestartNumberingAtSection = true;
                        wDocument.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.StartingNumber = (int)mf.StartNumberNumeric.Value; // номер первой страницы

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
                        headerTable.Cell(1, 6).Range.Font.Size = 10;

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
                        headerTable.Columns[3].PreferredWidth = 32f;
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
                        headerTable.Cell(1, 1).Range.Font.Size = 10;

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
                        headerTable.Columns[3].PreferredWidth = 32f;
                        headerTable.Columns[4].PreferredWidth = 9f;
                        headerTable.Columns[5].PreferredWidth = 4f;
                        headerTable.Columns[6].PreferredWidth = 4f;
                    }
                    else
                    {

                        Word.HeaderFooter header = wDocument.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];

                        Word.Range headerRange = header.Range;

                        header.LinkToPrevious = false;
                        header.PageNumbers.RestartNumberingAtSection = true;
                        header.PageNumbers.StartingNumber = (int)mf.StartNumberNumeric.Value; // номер первой страницы

                        // колонтитул страницы
                        wDocument.Tables.Add(headerRange, 1, 6, ref defaultTableBehavior, ref autoFitBehavior);
                        Word.Table headerTable = headerRange.Tables[1];

                        headerTable.Borders.Enable = 0;
                        Word.Range rangePageNum = headerTable.Range.Cells[headerTable.Range.Cells.Count].Range;
                        rangePageNum.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                        Word.Field fld = rangePageNum.Document.Fields.Add(rangePageNum, oMissing, "Page", false);
                        Word.Range rangeFieldPageNum = fld.Result;
                        rangeFieldPageNum.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                        headerTable.Cell(1, 6).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                        headerTable.Cell(1, 6).Range.Font.Size = 10;

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
                        headerTable.Columns[3].PreferredWidth = 32f;
                        headerTable.Columns[4].PreferredWidth = 9f;
                        headerTable.Columns[5].PreferredWidth = 4f;
                        headerTable.Columns[6].PreferredWidth = 4f;
                    }

                    // основная часть
                    wDocument.Paragraphs.Add(ref oMissing);
                    var Paragraph = wDocument.Paragraphs[1];
                    var Range = Paragraph.Range;
                    var Table = wDocument.Tables.Add(Range, 1, 6, ref defaultTableBehavior, ref autoFitBehavior);
                    // ширина ячеек таблицы
                    Table.Columns[1].PreferredWidth = 6f;
                    Table.Columns[2].PreferredWidth = 9f;
                    Table.Columns[3].PreferredWidth = 32f;
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
                    pagesInTitle = wDocument.ComputeStatistics(WdStatistic.wdStatisticPages, false);
                    int pageNumber = (int)mf.StartNumberNumeric.Value + pagesInTitle - 1;
                    row = 2;
                    if (mf.TwoSidedPrintCheckBox.Checked)
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
                        pageNumber += (int)mf.CountPagePZNumeric.Value - 1;
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
                        pageNumber += (int)mf.CountPagePZNumeric.Value - 1;
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
                mf.backgroundWorker.CancelAsync();
                mf.backgroundWorker.ReportProgress(45, "Сборка остановлена");

                return false;
            }
            finally
            {
                wordApp.Quit();
            }
        }

        public bool PdfMerge() // Соединение PDF файлов
        {
            try
            {
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                string fileNameConcatPdf = $"{finalSmetaFolder.FullName}\\TEMPdf\\smetaBook.pdf";
                string fileNameSmetaPdf = $"{finalSmetaFolder.FullName}\\Сметы.pdf";
                string fileNameTitlePdf = $"{path}\\TEMPdf\\Содержание.pdf";

                //TODO 1
                // тест сортировки смет по коду и имени
                var sortedObjData = objectiveData.OrderBy(ob => ob.Code).ThenBy(ob => ob.Name).ToList();
                var sortedLocData = localData.OrderBy(ob => ob.Code).ThenBy(ob => ob.Name).ToList();
                //----------------------

                allDataFilesList = sortedObjData;
                allDataFilesList.AddRange(sortedLocData);

                SmetaFile lastUsedDocument = null;

                PdfDocument inputPdfDocument;
                if (mf.partsBookCheckBox.Checked)
                {
                    int bookNumber = 1;
                    int i = 0;
                    bool changeBookCheck = true;


                    int tempFirstPageNubmer = 1;

                    while (lastUsedDocument != allDataFilesList[allDataFilesList.Count - 1])
                    {
                        PdfDocument outputSmetaPdfDocument = new PdfDocument();
                        for (; i < allDataFilesList.Count; i++) // TODO 3
                        {
                            var smetaFile = allDataFilesList[i];
                            inputPdfDocument = PdfReader.Open($"{pdfFolder}\\{smetaFile.FolderInfo}.pdf", PdfDocumentOpenMode.Import);
                            int pageCountInputDocument = inputPdfDocument.PageCount;
                            double dividerPass;

                            if (mf.AutoBooksPartPassCheckBox.Checked)
                            {
                                dividerPass = (double)mf.pagesInPartBookNumeric.Value * 12.5 / 100;
                            }
                            else
                            {
                                dividerPass = (double)mf.dividerPassPagesCount.Value;
                            }

                            if (outputSmetaPdfDocument.PageCount + pageCountInputDocument < (double)mf.pagesInPartBookNumeric.Value + dividerPass)
                            {
                                for (int j = 0; j < pageCountInputDocument; j++)
                                {
                                    PdfPage page = inputPdfDocument.Pages[j];
                                    outputSmetaPdfDocument.AddPage(page);
                                }
                                lastUsedDocument = smetaFile;
                                inputPdfDocument.Close();

                                allDataFilesList[i].Part = bookNumber; // тест

                                //Передача номера первой страницы каждого документа в сожержание
                                if (changeBookCheck)
                                {
                                    tempFirstPageNubmer = 1;
                                    changeBookCheck = false;
                                    firstPageNumbersList.Add(new List<int>());
                                }
                                else
                                {
                                    tempFirstPageNubmer += allDataFilesList[i - 1].PageCount;
                                }
                                firstPageNumbersList[bookNumber - 1].Add(tempFirstPageNubmer);
                            }
                            else
                            {
                                inputPdfDocument.Close();
                                tempFirstPageNubmer = 1;
                                break;
                            }
                        }

                        outputSmetaPdfDocument.Save($@"{finalSmetaFolder.FullName}\Сметы{bookNumber}.pdf");
                        outputSmetaPdfDocument.Close();

                        AddPageNumberSmetaITextSharp($@"{finalSmetaFolder.FullName}\Сметы{bookNumber}.pdf" , bookNumber);
                        bookNumber++;
                        changeBookCheck = true;
                    }
                }
                else
                {
                    int tempFirstPageNubmer = 1;
                    firstPageNumbersList.Add(new List<int>());
                    bool firstDocument = true;

                    if (mf.SplitBookContentCheckBox.Checked)
                    {
                        PdfDocument outputSmetaPdfDocument = new PdfDocument();
                        for (int i = 0; i < allDataFilesList.Count; i++)
                        {
                            var smetaFile = allDataFilesList[i];
                            inputPdfDocument = PdfReader.Open($"{pdfFolder}\\{smetaFile.FolderInfo}.pdf", PdfDocumentOpenMode.Import);
                            for (int j = 0; j < inputPdfDocument.PageCount; j++)
                            {
                                PdfPage page = inputPdfDocument.Pages[j];
                                outputSmetaPdfDocument.AddPage(page);
                            }
                            inputPdfDocument.Close();
                            //Передача номера первой страницы каждого документа в сожержание
                            if (firstDocument)
                            {
                                firstDocument = false;
                                firstPageNumbersList.Add(new List<int>());
                            }
                            else
                            {
                                tempFirstPageNubmer += allDataFilesList[i - 1].PageCount;
                            }
                            firstPageNumbersList[1].Add(tempFirstPageNubmer);
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
                        for (int i = 0; i < allDataFilesList.Count; i++)
                        {
                            var smetaFile = allDataFilesList[i];
                            inputPdfDocument = PdfReader.Open($"{pdfFolder}\\{smetaFile.FolderInfo}.pdf", PdfDocumentOpenMode.Import);
                            for (int j = 0; j < inputPdfDocument.PageCount; j++)
                            {
                                PdfPage page = inputPdfDocument.Pages[j];
                                outputPdfDocument.AddPage(page);
                            }
                            //Передача номера первой страницы каждого документа в сожержание
                            if (firstDocument)
                            {
                                firstDocument = false;
                            }
                            else
                            {
                                tempFirstPageNubmer += allDataFilesList[i - 1].PageCount;
                            }
                            firstPageNumbersList[1].Add(tempFirstPageNubmer);
                        }
                        outputPdfDocument.Save(fileNameConcatPdf);
                        inputPdfDocument.Close();
                        outputPdfDocument.Close();
                    }
                    if (mf.SplitBookContentCheckBox.Checked) //Нумерация страниц
                    {
                        AddPageNumberTitleITextSharp(fileNameTitlePdf);
                        AddPageNumberSmetaITextSharp(fileNameSmetaPdf, 1);
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
                mf.backgroundWorker.CancelAsync();
                DeleteTempFiles();
                DeleteTempVar();
                mf.backgroundWorker.ReportProgress(65, "Сборка остановлена");
                return false;
            }
        }

        public void AddPageNumberTitleITextSharp(string fileTitlePath) // Нумерация страниц содержания
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
                        int startPageNumber = Convert.ToInt32(mf.StartNumberNumeric.Value) - 1;

                        if (mf.TwoSidedPrintCheckBox.Checked)
                        {
                            for (int i = 1; i <= pagesTitle; i++)
                            {
                                if ((i + startPageNumber) % 2 == 0)
                                {
                                    iTextSharp.text.pdf.ColumnText.ShowTextAligned(stamper.GetUnderContent(i), Element.ALIGN_RIGHT, new Phrase((i + startPageNumber).ToString(), blackFont), 30f, 810f, 0);
                                }
                                else
                                {
                                    iTextSharp.text.pdf.ColumnText.ShowTextAligned(stamper.GetUnderContent(i), Element.ALIGN_RIGHT, new Phrase((i + startPageNumber).ToString(), blackFont), 565f, 810f, 0);
                                }
                            }
                        }
                        else
                        {
                            for (int i = 1; i <= pagesTitle; i++)
                            {
                                iTextSharp.text.pdf.ColumnText.ShowTextAligned(stamper.GetUnderContent(i), Element.ALIGN_RIGHT, new Phrase((i + startPageNumber).ToString(), blackFont), 565f, 15f, 0);
                            }
                        }
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
                mf.backgroundWorker.ReportProgress(65, "Сборка остановлена");
                mf.backgroundWorker.CancelAsync();
            }
        }

        public void AddPageNumberSmetaITextSharp(string filePath, int bookNumber) // Нумерация страниц книги смет
        {
            try
            {
                byte[] bytes = File.ReadAllBytes(filePath);

                iTextSharp.text.Font blackFont = FontFactory.GetFont("Arial", 8, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
                using (MemoryStream stream = new MemoryStream())
                {
                    iTextSharp.text.pdf.PdfReader reader = new iTextSharp.text.pdf.PdfReader(bytes);
                    int titlePages = pagesInTitle;
                    int pagesBook = reader.NumberOfPages;

                    using (iTextSharp.text.pdf.PdfStamper stamper = new iTextSharp.text.pdf.PdfStamper(reader, stream))
                    {
                        int startPageNumber = Convert.ToInt32(mf.StartNumberNumeric.Value) - 1;
                        int pagesPzCount = Convert.ToInt32(mf.CountPagePZNumeric.Value);

                        if (mf.TwoSidedPrintCheckBox.Checked)
                        {
                            if ((startPageNumber + titlePages) % 2 == 1)
                            {
                                titlePages++;
                            }
                            if (pagesPzCount % 2 == 1)
                            {
                                pagesPzCount++;
                            }
                            if(bookNumber != 1)
                            {
                                pagesPzCount = 0;
                            }

                            for (int i = 1; i <= pagesBook; i++)
                            {

                                if ((startPageNumber + titlePages + pagesPzCount + i) % 2 == 0)
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
                            if ((startPageNumber + titlePages) % 2 == 1)
                            {
                                titlePages++;
                            }
                            if (pagesPzCount % 2 == 1)
                            {
                                pagesPzCount++;
                            }

                            for (int i = 1; i <= pagesBook; i++)
                            {
                                iTextSharp.text.pdf.ColumnText.ShowTextAligned(stamper.GetUnderContent(i), Element.ALIGN_RIGHT, new Phrase((i + startPageNumber + pagesPzCount + titlePages).ToString(), blackFont), 810f, 15f, 0);

                            }
                        }
                        stamper.Close();
                        reader.Close();
                        bytes = stream.ToArray();
                    }
                    File.WriteAllBytes(filePath, bytes);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
                MessageBox.Show("Ошибка нумерации смет");
                DeleteTempFiles();
                DeleteTempVar();
                mf.backgroundWorker.ReportProgress(65, "Сборка остановлена");
                mf.backgroundWorker.CancelAsync();
            }
        }

        public void AddPageNumberITextSharp(string filePath) // Нумерация страниц содержания и книги смет
        {
            try
            {
                byte[] bytes = File.ReadAllBytes(filePath);

                iTextSharp.text.Font blackFont = FontFactory.GetFont("Arial", 8, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
                using (MemoryStream stream = new MemoryStream())
                {
                    iTextSharp.text.pdf.PdfReader reader = new iTextSharp.text.pdf.PdfReader(bytes);
                    int titlePages = pagesInTitle;
                    int pagesBook = reader.NumberOfPages;

                    using (iTextSharp.text.pdf.PdfStamper stamper = new iTextSharp.text.pdf.PdfStamper(reader, stream))
                    {
                        int startPageNumber = Convert.ToInt32(mf.StartNumberNumeric.Value) - 1;
                        int pagesPzCount = Convert.ToInt32(mf.CountPagePZNumeric.Value);

                        if (mf.TwoSidedPrintCheckBox.Checked)
                        {
                            if ((startPageNumber + titlePages) % 2 == 1)
                            {
                                titlePages++;
                            }
                            if (pagesPzCount % 2 == 1)
                            {
                                pagesPzCount++;
                            }

                            for (int i = 1; i <= pagesBook; i++)
                            {
                                if ((startPageNumber + titlePages + pagesPzCount + i) % 2 == 0)
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
                            if ((startPageNumber + titlePages) % 2 == 1)
                            {
                                titlePages++;
                            }
                            if (pagesPzCount % 2 == 1)
                            {
                                pagesPzCount++;
                            }
                            for (int i = 1; i <= pagesBook; i++)
                            {
                                iTextSharp.text.pdf.ColumnText.ShowTextAligned(stamper.GetUnderContent(i), Element.ALIGN_RIGHT, new Phrase((i + startPageNumber + pagesPzCount + titlePages).ToString(), blackFont), 810f, 15f, 0);
                            }
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
                mf.backgroundWorker.ReportProgress(65, "Сборка остановлена");
                mf.backgroundWorker.CancelAsync();
            }
        }

        public bool TitleNumOfPart() //Нумерация частей в содержании
        {
            Word.Application wordApp = new Word.Application
            {
                //Visible = true,
                //ScreenUpdating = true
                Visible = false,
                ScreenUpdating = false
            };

            try
            {

                var wDocument = wordApp.Documents.Open($@"{pdfFolder}\Содержание.docx");
                //var wDocument = wordApp.Documents.Open($@"C:\Users\lokot\Desktop\test2.docx");
                var table = wDocument.Tables[1];



                int startPageNumber = Convert.ToInt32(mf.StartNumberNumeric.Value) - 1; //TODO -1
                int pagesPzCount = Convert.ToInt32(mf.CountPagePZNumeric.Value);
                int titlePages = pagesInTitle;

                //-------------
                //pagesInTitle = wDocument.ComputeStatistics(WdStatistic.wdStatisticPages, false); // кол-во страниц в содержании
                int page = (int)mf.StartNumberNumeric.Value + pagesInTitle; // номер страницы

                //TODO добавление страниц после содержания 

                // нумерация ПЗ
                if ((page % 2) == 0)
                {
                    page += 1;
                }
                else
                {
                    page += 2;
                }
                //table.Cell(2, 5).Range.Text = page.ToString();
                table.Cell(2, 6).Range.Text = "1";
                page += (int)mf.CountPagePZNumeric.Value - 1; //TODO -1

                // нумерация сметы
                if ((page % 2) == 0)
                {
                    page += 1;
                }
                else
                {
                    page += 2;
                }
                int tempPage = page;

                // 1 вариант
                int i = 0;
                int temp = 0;
                if (mf.partsBookCheckBox.Checked)
                {
                    int rowInTable = table.Rows.Count;
                    for (var row = 1; row <= rowInTable; row++)
                    {
                        if (table.Cell(row, 2).Range.Text.Length > 3)
                        {
                            if (i != allDataFilesList.Count)
                            {
                                table.Cell(row, 6).Range.Text = allDataFilesList[i].Part.ToString();
                                //---------
                                if (temp != allDataFilesList[i].Part)
                                {
                                    page = tempPage;
                                    temp = allDataFilesList[i].Part;
                                }

                                table.Cell(row, 5).Range.Text = page.ToString();
                                page += allDataFilesList[i].PageCount;
                                //---------
                                i++;
                            }
                        }
                    }
                }


                wDocument.Save();
                if (Directory.Exists($"{pdfFolder}\\Содержание.pdf"))
                    Directory.Delete($"{pdfFolder}\\Содержание.pdf");
                wDocument.ExportAsFixedFormat($"{pdfFolder}\\Содержание.pdf", Word.WdExportFormat.wdExportFormatPDF);
                wDocument.Close(true);

                return true;
            }
            catch (Exception ex) 
            {
                MessageBox.Show(ex.Message.ToString());
                MessageBox.Show("Ошибка нумерации частей содержания");
                DeleteTempFiles();
                DeleteTempVar();
                mf.backgroundWorker.ReportProgress(80, "Сборка остановлена");
                mf.backgroundWorker.CancelAsync();
            }
            finally
            {
                wordApp.Quit();
                GC.Collect();
            }

            return false;
        }

        public bool MoveFiles() //Перемещение файлов в финальную папку
        {
            try
            {
                if (mf.SplitBookContentCheckBox.Checked)
                {
                    File.Move($@"{path}\TEMPdf\Содержание.pdf", $@"{finalSmetaFolder.FullName}\Содержание.pdf");
                    File.Move($@"{path}\TEMPdf\Содержание.docx", $@"{finalSmetaFolder.FullName}\Содержание.docx");
                }
                else
                {
                    File.Move($@"{path}\TEMPdf\smetaBook.pdf", $@"{finalSmetaFolder.FullName}\smetaBook.pdf");
                }

                mf.backgroundWorker.ReportProgress(77, "Сборка начата...");
                return true;
            }
            catch (Exception)
            {
                DeleteTempFiles();
                DeleteTempVar();
                MessageBox.Show("Ошибка перемещения файлов в финальную папку");
                mf.backgroundWorker.CancelAsync();
                mf.backgroundWorker.ReportProgress(85, "Сборка остановлена");
                return false;
            }
        }

        public void DeleteTempFiles() // Удаление временных файлов
        {
            if (Directory.Exists($"{path}\\TEMPdf"))
            {
                Directory.Delete($"{path}\\TEMPdf", true);
            }
        }

        public void DeleteTempVar()
        {
            path = null;
            dirFolders = null;
            pdfFolder = null;
            rootFolder = null;
            localFiles = null;
            childFolder = null;
            objectiveFiles = null;
            localData = new List<SmetaFile>();
            objectiveData = new List<SmetaFile>();
            allDataFilesList = new List<SmetaFile>();
            firstPageNumbersList = new List<List<int>>();
            GC.Collect();
        }
    }
}
