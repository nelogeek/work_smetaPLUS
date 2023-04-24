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
using Microsoft.Office.Core;
using iTextSharp.text.pdf.qrcode;
using System.Text;
using Org.BouncyCastle.Utilities;

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

        List<string> UpdatedFiles = new List<string>();

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
                            MessageBox.Show("В папке находится недопустимый файл", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
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
                if (objectiveFiles == null)
                {
                    mf.infoTextBox.AppendText(
                        Environment.NewLine + $"Количество всех файлов: {localFiles.Length}\n" +
                        Environment.NewLine + $"Количество папок: {dirFolders.Length}" + Environment.NewLine);
                }
                else
                {
                    mf.infoTextBox.AppendText(
                        Environment.NewLine + $"Количество всех файлов: {localFiles.Length + objectiveFiles.Length}\n" +
                        Environment.NewLine + $"Количество папок: {dirFolders.Length}" + Environment.NewLine);
                }

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
                MessageBox.Show("Ошибка во время выбора папки", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                Console.WriteLine(ex.Message.ToString());
                DeleteTempFiles();
                DeleteTempVar();
                GC.Collect();
                return;
            }
        }

        public int FullBookPageCounter() //Счетчик общего количества страниц
        {
            mf.btnReBuild.Enabled = false;
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
                MessageBox.Show("Ошибка подсчета страниц", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                Console.WriteLine(ex.Message.ToString());
                mf.backgroundWorker.CancelAsync();
                DeleteTempFiles();
                DeleteTempVar();
                GC.Collect();
                return 0;
            }
            finally
            {
                eWorkbook = null;
                app.Quit();
                ExcelKiller();

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
            mf.cbxType.Enabled = false;
        }

        public void EnableButtons()
        {
            mf.StartNumberNumeric.Enabled = true;
            mf.CountPagePZNumeric.Enabled = true;
            mf.btnBuild.Enabled = true;
            mf.btnSelectFolder.Enabled = true;
            mf.TwoSidedPrintCheckBox.Enabled = true;

            //mf.SplitBookContentCheckBox.Enabled = true;
            mf.SplitBookContentCheckBox.Checked = true;

            mf.RdPdToggle.Enabled = true;
            mf.settingsToolStripMenuItem.Enabled = true;
            mf.pagesInPartBookNumeric.Enabled = true;
            mf.partsBookCheckBox.Enabled = true;
            //mf.dividerPassPagesCount.Enabled = true;
            mf.cbxType.Enabled = true;
        }

        public bool Checker() // проверка на зависимость локальных смет от объектных (и наоборот)
        {
            Dictionary<string, int> temp = new Dictionary<string, int>();
            List<string> tempObj = new List<string>();
            List<string> tempLoc = new List<string>();

            try
            {
                if (mf.cbxIndex == 0)
                {
                    foreach (var obj in objectiveData) //для объектных смет
                    {

                        tempObj.Add(obj.ShortCode);
                    }
                    foreach (var loc in localData) //для локальных смет
                    {
                        if (loc.Type != FileType.Resource)
                        {
                            tempLoc.Add(loc.ShortCode);
                        }
                    }
                    foreach (var Obj in tempObj)
                    {
                        if (!temp.ContainsKey(Obj))
                        {
                            temp.Add(Obj, 0);
                        }

                    }
                    foreach (var Loc in tempLoc)
                    {
                        if (!temp.ContainsKey(Loc))
                        {
                            temp.Add(Loc, 0);
                        }
                    }

                    foreach (var Obj in tempObj)
                    {
                        foreach (var Loc in tempLoc)
                        {
                            if (Obj == Loc)
                            {
                                temp[Loc] += 1;
                                temp[Obj] += 1;
                            }
                        }
                    }
                    foreach (var T in temp)
                    {
                        if (T.Value == 0)
                        {

                            DialogResult dialogResult = MessageBox.Show("В исходной папке не соответствуют ОС и ЛС\n\n'ОК' - продолжить\n'Отмена' - выйти", "Внимание!", MessageBoxButtons.OKCancel);
                            if (dialogResult == DialogResult.Cancel)
                            {
                                return false;
                            }

                        }
                    }
                }

                else if (mf.cbxIndex == 1)
                {
                    foreach (var obj in objectiveData) //для объектных смет
                    {

                        tempObj.Add(obj.ShortCode);
                    }
                    foreach (var loc in localData) //для локальных смет
                    {
                        if (loc.Type != FileType.Resource)
                        {
                            tempLoc.Add(loc.ShortCode);
                        }
                    }
                    foreach (var Obj in tempObj)
                    {
                        if (!temp.ContainsKey(Obj))
                        {
                            temp.Add(Obj, 0);
                        }

                    }
                    foreach (var Loc in tempLoc)
                    {
                        if (!temp.ContainsKey(Loc))
                        {
                            temp.Add(Loc, 0);
                        }
                    }

                    foreach (var Obj in tempObj)
                    {
                        foreach (var Loc in tempLoc)
                        {
                            if (Obj == Loc)
                            {
                                temp[Loc] += 1;
                                temp[Obj] += 1;
                            }
                        }
                    }
                    foreach (var T in temp)
                    {
                        if (T.Value == 0)
                        {

                            DialogResult dialogResult = MessageBox.Show("В исходной папке не соответствуют ОС и ЛС\n\n'ОК' - продолжить\n'Отмена' - выйти", "Внимание!", MessageBoxButtons.OKCancel);
                            if (dialogResult == DialogResult.Cancel)
                            {
                                return false;
                            }

                        }
                    }
                }



                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                MessageBox.Show("Ошибка проверки смет", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                DeleteTempFiles();
                DeleteTempVar();
                return false;
            }
        }



        public bool FormatCheckerAndParser() // Проверка формата смет и парс данных
        {
            StringBuilder sb = new StringBuilder();
            bool sbBool = false;

            Excel.Application app = new Excel.Application
            {
                DisplayAlerts = false,
                Visible = false,
                ScreenUpdating = false
            };

            Excel.Workbook eWorkbook;
            Excel.Worksheet eWorksheet;

            string fileName = " "; //Для вывода информации о файле с ошибкой
            try
            {
                string code;
                string ShortCode;
                string money;
                string nameDate;
                string date;
                int pages;

                #region сметы Лукойл
                if (mf.cbxIndex == 0) // сметы Лукойла
                {


                    if (childFolder != null)
                    {
                        for (int i = 0; i < objectiveFiles.Length; i++) //Шаблон для объектных смет
                        {
                            try
                            {
                                string filePath = $"{childFolder}\\{objectiveFiles[i]}";
                                fileName = objectiveFiles[i].FullName;
                                eWorkbook = app.Workbooks.Open($@"{filePath}");
                                eWorksheet = (Excel.Worksheet)eWorkbook.Sheets[1];

                                eWorksheet.Application.ReferenceStyle = XlReferenceStyle.xlA1; // стиль ссылок
                                eWorksheet.PageSetup.Orientation = XlPageOrientation.xlLandscape; // альбомный вид
                                eWorksheet.PageSetup.PaperSize = XlPaperSize.xlPaperA4; // размер бумаги

                                Regex regex = new Regex(@"\b(\w*)-(\w*[.]?\w*)-(\w*[.]?\w*)\b");
                                code = regex.Matches(eWorksheet.Range["E8"].Value.ToString())[0].ToString().Replace("OC-", "").Replace("ОС-", "").Trim();
                                ShortCode = code.Replace("p", "").Replace("р", "").Replace("OC-", "").Replace("ОС-", "").Trim();
                                money = eWorksheet.Range["G12"].Value.ToString();
                                nameDate = eWorksheet.Range["C5"].Value.ToString();
                                date = eWorksheet.Range["C18"].Value.ToString().Split(new string[] { " цен " }, StringSplitOptions.None)[1];
                                nameDate += $"\n(в ценах на {date})";

                                if (mf.RdPdToggle.Checked)
                                {
                                    eWorksheet.Range["E8"].Replace("ОБЪЕКТНЫЙ СМЕТНЫЙ РАСЧЕТ (СМЕТА)", "ОБЪЕКТНАЯ СМЕТА");
                                }

                                if (mf.AutoPageBreakerToolStripMenuItem.Checked)
                                {
                                    PageBreaker(eWorksheet);
                                }

                                pages = eWorkbook.Sheets[1].PageSetup.Pages.Count;

                                objectiveData.Add(new SmetaFile(
                                    code, // код сметы
                                    eWorksheet.Range["C5"].Value.ToString(), // наименование
                                    nameDate, // Наименование c датой
                                    money, // Сумма денег
                                    pages, // Кол-во страниц на листе
                                    objectiveFiles[i], // путь файла
                                    ShortCode,
                                    FileType.Objective)); // короткий код для сравнения


                                eWorkbook.Save();
                                eWorkbook.Close(false);
                            }
                            catch (Exception)
                            {

                                sb.AppendLine(fileName + "\n");
                                sbBool = true;
                            }


                        }
                    }
                    for (int j = 0; j < localFiles.Length; j++) //Шаблон для локальных смет
                    {
                        try
                        {
                            string filePath = $"{rootFolder}\\{localFiles[j]}";
                            fileName = localFiles[j].FullName;
                            eWorkbook = app.Workbooks.Open($@"{filePath}");
                            eWorksheet = (Excel.Worksheet)eWorkbook.Sheets[1];

                            eWorksheet.Application.ReferenceStyle = XlReferenceStyle.xlA1; // стиль ссылок
                            eWorksheet.PageSetup.Orientation = XlPageOrientation.xlLandscape; // альбомный вид
                            eWorksheet.PageSetup.PaperSize = XlPaperSize.xlPaperA4; // размер бумаги

                            Regex regex = new Regex(@"\b(\w*)-(\w*[.]?\w*)-(\w*)\b");
                            MatchCollection match = regex.Matches(eWorksheet.Range["A18"].Value.ToString());
                            regex = new Regex(@"(\w*)-(\w*[.]?\w*)");
                            ShortCode = regex.Matches(match[0].Value.ToString())[0].ToString().Trim();

                            money = eWorksheet.Range["C28"].Value.ToString().Replace("(", "").Replace(")", "");
                            if (money == "0")
                            {
                                money = eWorksheet.Range["D28"].Value.ToString().Replace("(", "").Replace(")", "");
                            }
                            nameDate = eWorksheet.Range["A20"].Value.ToString();
                            date = eWorksheet.Range["D26"].Value.ToString();
                            nameDate += $"\n(в ценах на {date})";

                            if (mf.RdPdToggle.Checked)
                            {
                                eWorksheet.Range["A18"].Replace("ЛОКАЛЬНЫЙ СМЕТНЫЙ РАСЧЕТ (СМЕТА)", "ЛОКАЛЬНАЯ СМЕТА");
                            }

                            if (mf.AutoPageBreakerToolStripMenuItem.Checked)
                            {
                                PageBreaker(eWorksheet);
                            }

                            pages = eWorksheet.PageSetup.Pages.Count; // кол-во страниц на листе

                            localData.Add(new SmetaFile(
                                match[0].ToString(), // код сметы
                                eWorksheet.Range["A20"].Value.ToString(), // наименование
                                nameDate, // Наименование c датой
                                money, // Сумма денег
                                pages, // кол-во страниц на листе
                                localFiles[j], // путь файла
                                ShortCode,
                                FileType.Local)); // короткий код для сравнения

                            eWorkbook.Save();
                            eWorkbook.Close(false);

                        }
                        catch (Exception)
                        {
                            sb.AppendLine(fileName + "\n");
                            sbBool = true;
                        }
                    }

                    if (sbBool)
                    {
                        MessageBox.Show("Неправильный формат смет\n\nВся информация в log файле");
                        CreateFinalSmetaFolder();
                        File.AppendAllText($"{finalSmetaFolder.FullName}\\log.txt", sb.ToString());
                        sb.Clear();
                        return false;
                    }

                    localData = localData.OrderBy(x => x.Code).ThenBy(x => x.NameDate).ThenBy(x => x.Type).ToList(); // Сортировка по коду и названию
                    objectiveData = objectiveData.OrderBy(x => x.Code).ThenBy(x => x.NameDate).ThenBy(x => x.Type).ToList(); // Сортировка по коду и названию

                    code = null;
                    ShortCode = null;
                    money = null;
                    nameDate = null;
                    date = null;
                    pages = 0;
                }
                #endregion

                #region сметы Газпрома
                else if (mf.cbxIndex == 1) // сметы Газпрома
                {


                    if (childFolder != null)
                    {
                        for (int i = 0; i < objectiveFiles.Length; i++) //Шаблон для объектных смет
                        {
                            try
                            {
                                string filePath = $"{childFolder}\\{objectiveFiles[i]}";
                                fileName = objectiveFiles[i].FullName;
                                eWorkbook = app.Workbooks.Open($@"{filePath}");
                                eWorksheet = (Excel.Worksheet)eWorkbook.Sheets[1];
                                eWorksheet.Application.ReferenceStyle = XlReferenceStyle.xlA1; // стиль ссылок
                                eWorksheet.PageSetup.Orientation = XlPageOrientation.xlLandscape; // альбомный вид
                                eWorksheet.PageSetup.PaperSize = XlPaperSize.xlPaperA4; // размер бумаги

                                // удаление повторяющегося номера
                                eWorksheet.Range["B6"].Value = eWorksheet.Range["B6"].Value.ToString().Split('/')[0];

                                Regex regex = new Regex(@"\b(\w*[.]?\w*)-(\w*[.]?\w*)\b");
                                code = regex.Matches(eWorksheet.Range["B6"].Value.ToString())[0].ToString().Trim(); // .Replace("OC-", "").Replace("ОС-", "")
                                ShortCode = code.Replace("p", "").Replace("р", "").Replace("OC-", "").Replace("ОС-", "").Trim();
                                money = eWorksheet.Range["F13"].Value.ToString();
                                nameDate = eWorksheet.Range["C9"].Value.ToString();
                                date = eWorksheet.Range["B16"].Value.ToString().Split(new string[] { " цен " }, StringSplitOptions.None)[1];
                                nameDate += $"\n(в ценах на {date})";

                                if (mf.RdPdToggle.Checked)
                                {
                                    eWorksheet.Range["B6"].Replace("ОБЪЕКТНЫЙ СМЕТНЫЙ РАСЧЕТ (СМЕТА)", "ОБЪЕКТНАЯ СМЕТА");
                                }

                                if (mf.AutoPageBreakerToolStripMenuItem.Checked)
                                {
                                    PageBreaker(eWorksheet);
                                }

                                pages = eWorkbook.Sheets[1].PageSetup.Pages.Count;

                                objectiveData.Add(new SmetaFile(
                                    code, // код сметы
                                    eWorksheet.Range["C9"].Value.ToString(), // наименование
                                    nameDate, // Наименование c датой
                                    money, // Сумма денег
                                    pages, // Кол-во страниц на листе
                                    objectiveFiles[i], // путь файла
                                    ShortCode, // короткий код для сравнения
                                    FileType.Objective)); // тип документа


                                eWorkbook.Save();
                                eWorkbook.Close(false);
                            }
                            catch (Exception)
                            {

                                sb.AppendLine(fileName + "\n");
                                sbBool = true;
                            }


                        }
                    }
                    for (int j = 0; j < localFiles.Length; j++) //Шаблон для локальных и ресурсных смет
                    {
                        try
                        {
                            string filePath = $"{rootFolder}\\{localFiles[j]}";
                            fileName = localFiles[j].FullName;
                            eWorkbook = app.Workbooks.Open($@"{filePath}");
                            eWorksheet = (Excel.Worksheet)eWorkbook.Sheets[1];

                            eWorksheet.Application.ReferenceStyle = XlReferenceStyle.xlA1; // стиль ссылок
                            eWorksheet.PageSetup.Orientation = XlPageOrientation.xlLandscape; // альбомный вид
                            eWorksheet.PageSetup.PaperSize = XlPaperSize.xlPaperA4; // размер бумаги

                            string resType = null;
                            string locType = null;

                            try
                            {
                                resType = eWorksheet.Range["C4"].Value.ToString();
                            }
                            catch
                            {
                                locType = eWorksheet.Range["B12"].Value.ToString();
                            }

                            try
                            {
                                locType = eWorksheet.Range["B12"].Value.ToString();
                            }
                            catch
                            {
                                resType = eWorksheet.Range["C4"].Value.ToString();
                            }




                            if (locType.Contains("ЛОКАЛЬНАЯ СМЕТА")) // or
                            {
                                // удаление повторяющегося номера
                                eWorksheet.Range["B12"].Value = eWorksheet.Range["B12"].Value.ToString().Split('/')[0];

                                Regex regex = new Regex(@"\b(\w*)-(\w*[.]?\w*)-(\w*)\b");
                                MatchCollection match = regex.Matches(eWorksheet.Range["B12"].Value.ToString());
                                regex = new Regex(@"(\w*)-(\w*[.]?\w*)");
                                ShortCode = regex.Matches(match[0].Value.ToString())[0].ToString().Trim();

                                money = eWorksheet.Range["E19"].Value.ToString().Replace("(", "").Replace(")", "");
                                if (money == "0")
                                {
                                    money = eWorksheet.Range["D19"].Value.ToString().Replace("(", "").Replace(")", "");
                                }
                                nameDate = eWorksheet.Range["B9"].Value.ToString();
                                date = eWorksheet.Range["F25"].Value.ToString();
                                nameDate += $"\n(в ценах на {date})";

                                if (mf.RdPdToggle.Checked)
                                {
                                    eWorksheet.Range["B12"].Replace("ЛОКАЛЬНЫЙ СМЕТНЫЙ РАСЧЕТ (СМЕТА)", "ЛОКАЛЬНАЯ СМЕТА");
                                }

                                if (mf.AutoPageBreakerToolStripMenuItem.Checked)
                                {
                                    PageBreaker(eWorksheet);
                                }

                                pages = eWorksheet.PageSetup.Pages.Count; // кол-во страниц на листе

                                localData.Add(new SmetaFile(
                                    match[0].ToString(), // код сметы
                                    eWorksheet.Range["B9"].Value.ToString(), // наименование
                                    nameDate, // Наименование c датой
                                    money, // Сумма денег
                                    pages, // кол-во страниц на листе
                                    localFiles[j], // путь файла
                                    ShortCode, // короткий код для сравнения
                                    FileType.Local)); // тип документа

                                eWorkbook.Save();
                                eWorkbook.Close(false);
                            }
                            else if (resType.Contains("Локальный ресурсный сметный расчет")) // or
                            {
                                // сужение границ печати по столбцам
                                eWorksheet.PageSetup.PrintArea = "A:G";

                                // удаление повторяющегося номера
                                eWorksheet.Range["F4"].Value = eWorksheet.Range["F4"].Value.ToString().Split('/')[0];

                                Regex regex = new Regex(@"\b(\w*)-(\w*[.]?\w*)-(\w*)\b");
                                MatchCollection match = regex.Matches(eWorksheet.Range["F4"].Value.ToString());
                                regex = new Regex(@"(\w*)-(\w*[.]?\w*)");
                                ShortCode = regex.Matches(match[0].Value.ToString())[0].ToString().Trim();

                                nameDate = eWorksheet.Range["C8"].Value.ToString();

                                if (mf.AutoPageBreakerToolStripMenuItem.Checked)
                                {

                                    PageBreaker(eWorksheet);


                                }

                                pages = eWorksheet.PageSetup.Pages.Count; // кол-во страниц на листе

                                localData.Add(new SmetaFile(
                                    match[0].ToString(), // код сметы
                                    eWorksheet.Range["C8"].Value.ToString(), // наименование
                                    nameDate, // Наименование c датой
                                    null, // Сумма денег
                                    pages, // кол-во страниц на листе
                                    localFiles[j], // путь файла
                                    ShortCode, // короткий код для сравнения 
                                    FileType.Resource)); // тип документа

                                eWorkbook.Save();
                                eWorkbook.Close(false);
                            }
                        }
                        catch (Exception)
                        {
                            sb.AppendLine(fileName + "\n");
                            sbBool = true;
                        }
                    }

                    if (sbBool)
                    {
                        MessageBox.Show("Неправильный формат смет\n\nВся информация в log файле");
                        CreateFinalSmetaFolder();
                        File.AppendAllText($"{finalSmetaFolder.FullName}\\log.txt", sb.ToString());
                        sb.Clear();
                        return false;
                    }

                    localData = localData.OrderBy(x => x.Code).ThenBy(x => x.NameDate).ToList(); // Сортировка по коду и названию
                    objectiveData = objectiveData.OrderBy(x => x.Code).ThenBy(x => x.NameDate).ToList(); // Сортировка по коду и названию

                    code = null;
                    ShortCode = null;
                    money = null;
                    nameDate = null;
                    date = null;
                    pages = 0;
                }
                #endregion


                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при парсинге смет", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                Console.WriteLine(ex.StackTrace);
                Console.WriteLine(ex.Message.ToString());
                //MessageBox.Show(fileName, "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Console.WriteLine(ex.StackTrace);
                mf.backgroundWorker.CancelAsync();
                DeleteTempFiles();
                DeleteTempVar();
                mf.backgroundWorker.ReportProgress(1, "Сборка остановлена");

                return false;
            }
            finally
            {
                eWorksheet = null;
                eWorkbook = null;
                app.Quit();
                ExcelKiller();
                GC.Collect();
            }
        }


        public bool ExcelParser() // Парсинг Excel файла
        {
            Excel.Application app = new Excel.Application
            {
                DisplayAlerts = false,
                Visible = false,
                ScreenUpdating = false
            };

            Excel.Workbook eWorkbook;
            Excel.Worksheet eWorksheet;

            string fileName = " "; //Для вывода информации о файле с ошибкой
            try
            {
                string code;
                string ShortCode;
                string money;
                string nameDate;
                string date;
                int pages;

                if (mf.cbxIndex == 0) // Шаблоны Лукойл
                {


                    if (childFolder != null)
                    {
                        for (int i = 0; i < objectiveFiles.Length; i++) //Шаблон для объектных смет
                        {
                            string filePath = $"{childFolder}\\{objectiveFiles[i]}";
                            fileName = objectiveFiles[i].FullName;
                            eWorkbook = app.Workbooks.Open($@"{filePath}");
                            eWorksheet = (Excel.Worksheet)eWorkbook.Sheets[1];
                            eWorksheet.PageSetup.Orientation = XlPageOrientation.xlLandscape;
                            eWorksheet.PageSetup.PaperSize = XlPaperSize.xlPaperA4;

                            //Regex regex = new Regex(@"(\w*)-(\w*)-(\w*)");
                            Regex regex = new Regex(@"\b(\w*)-(\w*[.]?\w*)-(\w*[.]?\w*)\b");
                            code = regex.Matches(eWorksheet.Range["E8"].Value.ToString())[0].ToString().Replace("OC-", "").Replace("ОС-", "").Trim();
                            ShortCode = code.Replace("p", "").Replace("р", "").Replace("OC-", "").Replace("ОС-", "").Trim();
                            money = eWorksheet.Range["G12"].Value.ToString();
                            nameDate = eWorksheet.Range["C5"].Value.ToString();
                            date = eWorksheet.Range["C18"].Value.ToString().Split(new string[] { " цен " }, StringSplitOptions.None)[1];
                            nameDate += $"\n(в ценах на {date})";

                            if (mf.RdPdToggle.Checked)
                            {
                                eWorksheet.Range["E8"].Replace("ОБЪЕКТНЫЙ СМЕТНЫЙ РАСЧЕТ (СМЕТА)", "ОБЪЕКТНАЯ СМЕТА");
                            }

                            if (mf.AutoPageBreakerToolStripMenuItem.Checked)
                            {
                                PageBreaker(eWorksheet);
                            }

                            pages = eWorkbook.Sheets[1].PageSetup.Pages.Count;

                            objectiveData.Add(new SmetaFile(
                                code, // код сметы
                                eWorksheet.Range["C5"].Value.ToString(), // наименование
                                nameDate, // Наименование c датой
                                money, // Сумма денег
                                pages, // Кол-во страниц на листе
                                objectiveFiles[i], // путь файла
                                ShortCode,
                                FileType.Objective)); // короткий код для сравнения


                            eWorkbook.Save();
                            eWorkbook.Close(false);
                        }
                    }
                    for (int j = 0; j < localFiles.Length; j++) //Шаблон для локальных смет
                    {
                        string filePath = $"{rootFolder}\\{localFiles[j]}";
                        fileName = localFiles[j].FullName;
                        eWorkbook = app.Workbooks.Open($@"{filePath}");
                        eWorksheet = (Excel.Worksheet)eWorkbook.Sheets[1];
                        eWorksheet.PageSetup.Orientation = XlPageOrientation.xlLandscape;
                        eWorksheet.PageSetup.PaperSize = XlPaperSize.xlPaperA4;

                        Regex regex = new Regex(@"\b(\w*)-(\w*[.]?\w*)-(\w*)\b");
                        MatchCollection match = regex.Matches(eWorksheet.Range["A18"].Value.ToString());
                        regex = new Regex(@"(\w*)-(\w*[.]?\w*)");
                        ShortCode = regex.Matches(match[0].Value.ToString())[0].ToString().Trim();

                        money = eWorksheet.Range["C28"].Value.ToString().Replace("(", "").Replace(")", "");
                        if (money == "0")
                        {
                            money = eWorksheet.Range["D28"].Value.ToString().Replace("(", "").Replace(")", "");
                        }
                        nameDate = eWorksheet.Range["A20"].Value.ToString();
                        date = eWorksheet.Range["D26"].Value.ToString();
                        nameDate += $"\n(в ценах на {date})";

                        if (mf.RdPdToggle.Checked)
                        {
                            eWorksheet.Range["A18"].Replace("ЛОКАЛЬНЫЙ СМЕТНЫЙ РАСЧЕТ (СМЕТА)", "ЛОКАЛЬНАЯ СМЕТА");
                        }

                        if (mf.AutoPageBreakerToolStripMenuItem.Checked)
                        {
                            PageBreaker(eWorksheet);
                        }

                        pages = eWorksheet.PageSetup.Pages.Count; // кол-во страниц на листе

                        localData.Add(new SmetaFile(
                            match[0].ToString(), // код сметы
                            eWorksheet.Range["A20"].Value.ToString(), // наименование
                            nameDate, // Наименование c датой
                            money, // Сумма денег
                            pages, // кол-во страниц на листе
                            localFiles[j], // путь файла
                            ShortCode,
                            FileType.Local)); // короткий код для сравнения

                        eWorkbook.Save();
                        eWorkbook.Close(false);
                    }

                    localData = localData.OrderBy(x => x.Code).ThenBy(x => x.NameDate).ToList(); // Сортировка по коду и названию
                    objectiveData = objectiveData.OrderBy(x => x.Code).ThenBy(x => x.NameDate).ToList(); // Сортировка по коду и названию

                    code = null;
                    ShortCode = null;
                    money = null;
                    nameDate = null;
                    date = null;
                    pages = 0;
                }
                else if (mf.cbxIndex == 1)
                {
                    Console.WriteLine("Not");
                }


                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при парсинге смет", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                Console.WriteLine(ex.StackTrace);
                Console.WriteLine(ex.Message.ToString());
                MessageBox.Show(fileName, "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Console.WriteLine(ex.StackTrace);
                mf.backgroundWorker.CancelAsync();
                DeleteTempFiles();
                DeleteTempVar();
                mf.backgroundWorker.ReportProgress(1, "Сборка остановлена");

                return false;
            }
            finally
            {
                eWorksheet = null;
                eWorkbook = null;
                app.Quit();
                ExcelKiller();
                GC.Collect();
            }
        }

        private char[] chars = new char[] { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z' };
        string GetColumnName(int index)
        {
            index -= 1; //adjust so it matches 0-indexed array rather than 1-indexed column

            int quotient = index / 26;
            if (quotient > 0)
                return GetColumnName(quotient) + chars[index % 26].ToString();
            else
                return chars[index % 26].ToString();
        }


        protected void PageBreaker(Excel.Worksheet eWorksheet) // Регулировка разрывов страниц
        {
            try
            {
                eWorksheet.Range[$"G7"].Value = ""; //Удаление строки приказов
                eWorksheet.Rows[7].RowHeight = 11.25;
                int lastUsedRow = eWorksheet.Cells.Find("*", System.Reflection.Missing.Value,
                       System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                       Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                       false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
                var hPageBreaks = eWorksheet.HPageBreaks;
                eWorksheet.ResetAllPageBreaks();

                for (int p = 1; p <= hPageBreaks.Count; p++)
                {
                    int rowPageBreak = hPageBreaks[p].Location.Row;
                    hPageBreaks.Add(eWorksheet.Range[$"A{rowPageBreak}"]);
                }
                int lastPageBreak = hPageBreaks[hPageBreaks.Count].Location.Row;
                if (lastUsedRow - lastPageBreak < 13)
                {
                    hPageBreaks[hPageBreaks.Count].Delete();
                    hPageBreaks.Add(eWorksheet.Range[$"A{lastUsedRow - 13}"]);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.StackTrace);
                Console.WriteLine(ex.Message.ToString());
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
                    eWorksheet.PageSetup.RightFooter = ""; //Удаление нумерации страниц в Excel
                    app.ActiveWorkbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, $"{pdfFolder.FullName}\\{file.FolderInfo}");
                    eWorkbook.Close(false);
                }
                foreach (var file in localData)
                {
                    string filePath = $"{path}\\{file.FolderInfo}";
                    eWorkbook = app.Workbooks.Open(filePath);
                    eWorksheet = (Worksheet)eWorkbook.Sheets[1];
                    eWorksheet.PageSetup.RightFooter = ""; //Удаление нумерации страниц в Excel
                    app.ActiveWorkbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, $"{pdfFolder.FullName}\\{file.FolderInfo}");
                    eWorkbook.Close(false);
                }

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка конвертации в PDF", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                Console.WriteLine(ex.Message.ToString());
                mf.backgroundWorker.CancelAsync();
                DeleteTempFiles();
                DeleteTempVar();
                mf.backgroundWorker.ReportProgress(10, "Сборка остановлена");

                return false;
            }
            finally
            {
                eWorksheet = null;
                eWorkbook = null;
                app.Quit();
                ExcelKiller();
                GC.Collect();
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
                MessageBox.Show("Ошибка создания финальной папки", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                Console.WriteLine(ex.Message.ToString());
                mf.backgroundWorker.CancelAsync();
                mf.backgroundWorker.ReportProgress(45, "Сборка остановлена");
                DeleteTempFiles();
                DeleteTempVar();
                GC.Collect();
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
                    Table.Cell(row, 6).Range.Text = "1";
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
                        Table.Cell(row, 6).Range.Text = "1";
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
                            if (lData.ShortCode == oData.Key && lData.Type == FileType.Local)
                            {
                                NumberDocument++;
                                Table.Rows.Add();
                                Table.Cell(row, 1).Range.Text = NumberDocument.ToString();
                                Table.Cell(row, 2).Range.Text = lData.Code;
                                Table.Cell(row, 3).Range.Text = lData.NameDate + "\n";
                                Table.Cell(row, 4).Range.Text = lData.Price;
                                Table.Cell(row, 6).Range.Text = "1";
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
                    int pageNumber = (int)mf.StartNumberNumeric.Value + pagesInTitle - 1; // номера страниц

                    row = 2;

                    if (mf.TwoSidedPrintCheckBox.Checked)
                    {
                        if (mf.cbxIndex != 1)
                        {
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

                    }
                    else
                    {
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

                    if (File.Exists($"{pdfFolder}\\Содержание.docx"))
                    {
                        File.Delete($"{pdfFolder}\\Содержание.docx");
                    }
                    if (File.Exists($"{pdfFolder}\\Содержание.pdf"))
                    {
                        File.Delete($"{pdfFolder}\\Содержание.pdf");
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
                    Table.Cell(row, 6).Range.Text = "1";
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
                        if (lData.Type == FileType.Local)
                        {
                            NumberDocument++;
                            Table.Rows.Add();
                            Table.Cell(row, 1).Range.Text = NumberDocument.ToString();
                            Table.Cell(row, 2).Range.Text = lData.Code;
                            Table.Cell(row, 3).Range.Text = lData.NameDate + "\n";
                            Table.Cell(row, 4).Range.Text = lData.Price;
                            Table.Cell(row, 6).Range.Text = "1";
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
                    }

                    //нумерация страниц
                    pagesInTitle = wDocument.ComputeStatistics(WdStatistic.wdStatisticPages, false);
                    int pageNumber = (int)mf.StartNumberNumeric.Value + pagesInTitle - 1;
                    row = 2;
                    if (mf.TwoSidedPrintCheckBox.Checked)
                    {
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




                    if (File.Exists($"{pdfFolder}\\Содержание.docx"))
                    {
                        File.Delete($"{pdfFolder}\\Содержание.docx");
                    }
                    if (File.Exists($"{pdfFolder}\\Содержание.pdf"))
                    {
                        File.Delete($"{pdfFolder}\\Содержание.pdf");
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
                MessageBox.Show("Ошибка генерации содержания", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                Console.WriteLine(ex.Message.ToString());
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

                List<SmetaFile> sortedObjData = objectiveData.OrderBy(ob => ob.Code).ThenBy(ob => ob.Name).ThenBy(ob => ob.Type).ToList();
                List<SmetaFile> sortedLocData = localData.OrderBy(ob => ob.Code).ThenBy(ob => ob.Name).ThenBy(ob => ob.Type).ToList();

                allDataFilesList = sortedObjData;
                allDataFilesList.AddRange(sortedLocData);

                SmetaFile lastUsedDocument = null;

                PdfDocument inputPdfDocument;
                if (mf.partsBookCheckBox.Checked)
                {
                    int bookNumber = 1;
                    int index = 0;
                    bool changeBookCheck = true;
                    int tempFirstPageNubmer = 1;

                    while (lastUsedDocument != allDataFilesList[allDataFilesList.Count - 1])
                    {
                        PdfDocument outputSmetaPdfDocument = new PdfDocument();
                        for (; index < allDataFilesList.Count; index++)
                        {
                            var smetaFile = allDataFilesList[index];
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

                                allDataFilesList[index].Part = bookNumber;

                                //Передача номера первой страницы каждого документа в содержание
                                if (changeBookCheck)
                                {
                                    tempFirstPageNubmer = 1;
                                    changeBookCheck = false;
                                    firstPageNumbersList.Add(new List<int>());
                                }
                                else
                                {
                                    tempFirstPageNubmer += allDataFilesList[index - 1].PageCount;
                                }
                                if (smetaFile.Type != FileType.Resource)
                                {
                                    firstPageNumbersList[bookNumber - 1].Add(tempFirstPageNubmer);
                                }

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

                        AddPageNumberSmetaITextSharp($@"{finalSmetaFolder.FullName}\Сметы{bookNumber}.pdf", bookNumber);
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
                            //Передача номера первой страницы каждого документа в содержание
                            if (firstDocument)
                            {
                                firstDocument = false;
                                firstPageNumbersList.Add(new List<int>());
                            }
                            else
                            {
                                tempFirstPageNubmer += allDataFilesList[i - 1].PageCount;
                            }
                            if (smetaFile.Type != FileType.Resource)
                            {
                                firstPageNumbersList[1].Add(tempFirstPageNubmer);
                            }
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
                            //Передача номера первой страницы каждого документа в содержание
                            if (firstDocument)
                            {
                                firstDocument = false;
                            }
                            else
                            {
                                tempFirstPageNubmer += allDataFilesList[i - 1].PageCount;
                            }
                            if (smetaFile.Type != FileType.Resource)
                            {
                                firstPageNumbersList[1].Add(tempFirstPageNubmer);
                            }
                           
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
                MessageBox.Show("Ошибка сборки книги", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
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
                MessageBox.Show("Ошибка нумерации содержания", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                            if (bookNumber != 1)
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
                Console.WriteLine(ex.Message.ToString());
                MessageBox.Show("Ошибка нумерации смет", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                MessageBox.Show("Ошибка нумерации книги", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                mf.backgroundWorker.ReportProgress(65, "Сборка остановлена");
                mf.backgroundWorker.CancelAsync();
            }
        }


        public bool TitleNumOfPart() //Нумерация частей и страниц содержания
        {
            Word.Application wordApp = new Word.Application
            {
                Visible = false,
                ScreenUpdating = false
                //Visible = true,
                //ScreenUpdating = true
            };
            var wDocument = wordApp.Documents.Open($@"{pdfFolder}\Содержание.docx");

            try
            {
                var table = wDocument.Tables[1];

                int titlePages = pagesInTitle;
                int startPageNumber = Convert.ToInt32(mf.StartNumberNumeric.Value) - 1;
                int pagesPzCount = Convert.ToInt32(mf.CountPagePZNumeric.Value);

                if ((startPageNumber + titlePages) % 2 == 1)
                {
                    titlePages++;
                }
                int page = startPageNumber + pagesPzCount + titlePages;
                table.Cell(2, 5).Range.Text = page.ToString();

                if (pagesPzCount % 2 == 1)
                {
                    pagesPzCount++;
                }
                page = startPageNumber + pagesPzCount + titlePages;

                int DataFilesCount = 0;
                int FileIndex = 0;
                int NumberOfPart = 0;
                if (mf.partsBookCheckBox.Checked)
                {
                    int rowInTable = table.Rows.Count;
                    for (var row = 1; row <= rowInTable; row++)
                    {
                        if (table.Cell(row, 2).Range.Text.Length > 3)
                        {
                            if (DataFilesCount != allDataFilesList.Count)
                            {
                                table.Cell(row, 6).Range.Text = allDataFilesList[DataFilesCount].Part.ToString();

                                if (NumberOfPart != allDataFilesList[DataFilesCount].Part)
                                {
                                    NumberOfPart = allDataFilesList[DataFilesCount].Part;
                                    FileIndex = 0;
                                }
                                if (NumberOfPart != 1)
                                {
                                    pagesPzCount = 0;
                                }
                                page = startPageNumber + pagesPzCount + titlePages;

                                table.Cell(row, 5).Range.Text = (page + firstPageNumbersList[NumberOfPart - 1][FileIndex]).ToString();
                                DataFilesCount++;
                                FileIndex++;
                            }
                        }
                    }
                }
                else
                {
                    int rowInTable = table.Rows.Count;
                    for (var row = 1; row <= rowInTable; row++)
                    {
                        if (table.Cell(row, 2).Range.Text.Length > 3)
                        {
                            if (DataFilesCount != allDataFilesList.Count)
                            {
                                table.Cell(row, 6).Range.Text = allDataFilesList[DataFilesCount].Part.ToString();

                                if (NumberOfPart != allDataFilesList[DataFilesCount].Part)
                                {
                                    NumberOfPart = allDataFilesList[DataFilesCount].Part;
                                    FileIndex = 0;
                                }
                                if (NumberOfPart != 1)
                                {
                                    pagesPzCount = 0;
                                }
                                page = startPageNumber + pagesPzCount + titlePages;

                                table.Cell(row, 5).Range.Text = (page + firstPageNumbersList[0][FileIndex]).ToString();
                                DataFilesCount++;
                                FileIndex++;
                            }
                        }
                    }
                }

                wDocument.Save();
                if (Directory.Exists($"{pdfFolder}\\Содержание.pdf"))
                {
                    Directory.Delete($"{pdfFolder}\\Содержание.pdf");
                }
                wDocument.ExportAsFixedFormat($"{pdfFolder}\\Содержание.pdf", Word.WdExportFormat.wdExportFormatPDF);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message.ToString());
                MessageBox.Show("Ошибка нумерации частей содержания", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                DeleteTempFiles();
                DeleteTempVar();
                mf.backgroundWorker.ReportProgress(80, "Сборка остановлена");
                mf.backgroundWorker.CancelAsync();
                return false;
            }
            finally
            {
                wDocument.Close(false);
                wordApp.Quit();
                GC.Collect();
            }

        }

        public bool MoveFiles() //Перемещение файлов в финальную папку
        {
            try
            {
                if (mf.SplitBookContentCheckBox.Checked)
                {
                    File.Copy($@"{path}\TEMPdf\Содержание.pdf", $@"{finalSmetaFolder.FullName}\Содержание.pdf", true);
                    File.Copy($@"{path}\TEMPdf\Содержание.docx", $@"{finalSmetaFolder.FullName}\Содержание.docx", true);
                }
                else
                {
                    File.Copy($@"{path}\TEMPdf\smetaBook.pdf", $@"{finalSmetaFolder.FullName}\smetaBook.pdf", true);
                }

                mf.backgroundWorker.ReportProgress(77, "Сборка начата...");
                return true;
            }
            catch (Exception)
            {
                DeleteTempFiles();
                DeleteTempVar();
                MessageBox.Show("Ошибка перемещения файлов в финальную папку", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                mf.backgroundWorker.CancelAsync();
                mf.backgroundWorker.ReportProgress(85, "Сборка остановлена");
                return false;
            }
        }

        public void DeleteTempFiles() // Удаление временных файлов
        {
            try
            {

                if (Directory.Exists($"{path}\\TEMPdf"))
                {
                    Directory.Delete($"{path}\\TEMPdf", true);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                MessageBox.Show("Ошибка удаления временных файлов.\n\nЗакройте все файлы, принадлежащие сметам, и повторите попытку", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            List<string> UpdatedFiles = new List<string>();
            GC.Collect();
        }


        public bool UpdatedFindFiles()
        {
            try
            {
                if (childFolder != null)
                {
                    for (int i = 0; i < objectiveFiles.Length; i++)
                    {
                        var a = File.GetLastWriteTime($"{childFolder}\\{objectiveFiles[i]}");
                        var b = File.GetLastWriteTime($"{pdfFolder}\\{objectiveFiles[i]}.pdf");
                        //Console.WriteLine($"{objectiveFiles[i]} {a > b}");
                        if (a > b)
                        {
                            //UpdatedFiles.Append($"{childFolder}\\{objectiveFiles[i]}");
                            UpdatedFiles.Add($"{childFolder}\\{objectiveFiles[i]}");
                        }
                    }
                }
                Console.WriteLine("-----------------");
                for (int j = 0; j < localFiles.Length; j++)
                {
                    var a = File.GetLastWriteTime($"{rootFolder}\\{localFiles[j]}");
                    var b = File.GetLastWriteTime($"{pdfFolder}\\{localFiles[j]}.pdf");
                    //Console.WriteLine($"{localFiles[j]} {a > b}");
                    if (a > b)
                    {
                        //UpdatedFiles.Append($"{rootFolder}\\{localFiles[j]}");
                        UpdatedFiles.Add($"{rootFolder}\\{localFiles[j]}");
                    }

                }
                foreach (var file in UpdatedFiles)
                {
                    Console.WriteLine(file);
                }

                return true;
            }
            catch (Exception ex)
            {
                DeleteTempFiles();
                DeleteTempVar();
                Console.WriteLine(ex.Message.ToString());
                MessageBox.Show("Ошибка поиска измененных файлов", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                mf.backgroundWorker2.ReportProgress(20, "Сборка остановлена");
                mf.backgroundWorker2.CancelAsync();
                return false;
            }


        }

        public bool UpdatedConverter()
        {
            Excel.Application app = new Excel.Application { DisplayAlerts = false, Visible = false, ScreenUpdating = false };
            Workbook eWorkbook;
            Worksheet eWorksheet;

            try
            {
                //Directory.CreateDirectory($"{path}\\TEMPdf");
                foreach (var file in UpdatedFiles)
                {
                    var fileName = file.Split(new string[] { "\\" }, StringSplitOptions.None).Last();
                    eWorkbook = app.Workbooks.Open(file);
                    eWorksheet = (Worksheet)eWorkbook.Sheets[1];
                    eWorksheet.PageSetup.RightFooter = ""; //Удаление нумерации станиц в Excel
                    if (File.Exists($"{pdfFolder.FullName}\\{fileName}"))
                    {
                        File.Delete($"{pdfFolder.FullName}\\{fileName}");
                    }
                    app.ActiveWorkbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, $"{pdfFolder.FullName}\\{fileName}");
                    eWorkbook.Close(false);
                }

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка конвертации в PDF", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Console.WriteLine(ex.Message.ToString());
                mf.backgroundWorker.CancelAsync();
                DeleteTempFiles();
                DeleteTempVar();
                mf.backgroundWorker2.ReportProgress(45, "Сборка остановлена");

                return false;
            }
            finally
            {
                eWorksheet = null;
                eWorkbook = null;
                app.Quit();
                ExcelKiller();
                GC.Collect();
            }
        }

        public bool UpdatedCreateFinalSmetaFolder() //Создание финальной папки
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
                MessageBox.Show("Ошибка создания финальной папки", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Console.WriteLine(ex.Message.ToString());
                mf.backgroundWorker2.CancelAsync();
                mf.backgroundWorker2.ReportProgress(50, "Сборка остановлена");
                DeleteTempFiles();
                DeleteTempVar();
                GC.Collect();
                return false;
            }
        }

        public bool UpdatedMoveFiles() //Перемещение файлов в финальную папку
        {
            try
            {
                if (mf.SplitBookContentCheckBox.Checked)
                {
                    new FileInfo($@"{path}\TEMPdf\Содержание.pdf").CopyTo($@"{finalSmetaFolder.FullName}\Содержание.pdf", true);
                    new FileInfo($@"{path}\TEMPdf\Содержание.docx").CopyTo($@"{finalSmetaFolder.FullName}\Содержание.docx", true);
                }
                else
                {
                    File.Copy($@"{path}\TEMPdf\smetaBook.pdf", $@"{finalSmetaFolder.FullName}\smetaBook.pdf", true);
                }
                return true;
            }
            catch (Exception)
            {
                //TODO
                DeleteTempFiles();
                DeleteTempVar();
                MessageBox.Show("Ошибка перемещения файлов в финальную папку", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                mf.backgroundWorker.CancelAsync();
                mf.backgroundWorker.ReportProgress(95, "Сборка остановлена");
                return false;
            }
        }

        public bool UpdatedPdfMerge() // Соединение PDF файлов
        {
            try
            {
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                string fileNameConcatPdf = $"{finalSmetaFolder.FullName}\\TEMPdf\\smetaBook.pdf";
                string fileNameSmetaPdf = $"{finalSmetaFolder.FullName}\\Сметы.pdf";
                string fileNameTitlePdf = $"{path}\\TEMPdf\\Содержание.pdf";

                List<SmetaFile> sortedObjData = objectiveData.OrderBy(ob => ob.Code).ThenBy(ob => ob.Name).ThenBy(ob => ob.Type).ToList();
                List<SmetaFile> sortedLocData = localData.OrderBy(ob => ob.Code).ThenBy(ob => ob.Name).ThenBy(ob => ob.Type).ToList();

                allDataFilesList = sortedObjData;
                allDataFilesList.AddRange(sortedLocData);

                SmetaFile lastUsedDocument = null;

                PdfDocument inputPdfDocument;
                if (mf.partsBookCheckBox.Checked)
                {
                    int bookNumber = 1;
                    int index = 0;
                    bool changeBookCheck = true;
                    int tempFirstPageNubmer = 1;

                    while (lastUsedDocument != allDataFilesList[allDataFilesList.Count - 1])
                    {
                        PdfDocument outputSmetaPdfDocument = new PdfDocument();
                        for (; index < allDataFilesList.Count; index++)
                        {
                            var smetaFile = allDataFilesList[index];
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

                                allDataFilesList[index].Part = bookNumber;

                                //Передача номера первой страницы каждого документа в содержание
                                if (changeBookCheck)
                                {
                                    tempFirstPageNubmer = 1;
                                    changeBookCheck = false;
                                    firstPageNumbersList.Add(new List<int>());
                                }
                                else
                                {
                                    tempFirstPageNubmer += allDataFilesList[index - 1].PageCount;
                                }
                                if (smetaFile.Type != FileType.Resource)
                                {
                                    firstPageNumbersList[bookNumber - 1].Add(tempFirstPageNubmer);
                                }

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

                        AddPageNumberSmetaITextSharp($@"{finalSmetaFolder.FullName}\Сметы{bookNumber}.pdf", bookNumber);
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
                            //Передача номера первой страницы каждого документа в содержание
                            if (firstDocument)
                            {
                                firstDocument = false;
                                firstPageNumbersList.Add(new List<int>());
                            }
                            else
                            {
                                tempFirstPageNubmer += allDataFilesList[i - 1].PageCount;
                            }
                            if (smetaFile.Type != FileType.Resource)
                            {
                                firstPageNumbersList[1].Add(tempFirstPageNubmer);
                            }
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
                            //Передача номера первой страницы каждого документа в содержание
                            if (firstDocument)
                            {
                                firstDocument = false;
                            }
                            else
                            {
                                tempFirstPageNubmer += allDataFilesList[i - 1].PageCount;
                            }
                            if (smetaFile.Type != FileType.Resource)
                            {
                                firstPageNumbersList[1].Add(tempFirstPageNubmer);
                            }

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
                MessageBox.Show("Ошибка сборки книги", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                mf.backgroundWorker.CancelAsync();
                DeleteTempFiles();
                DeleteTempVar();
                mf.backgroundWorker2.ReportProgress(65, "Сборка остановлена");
                return false;
            }
        }
        public bool UpdatedTitleNumOfPart() //Нумерация частей и страниц содержания
        {
            Word.Application wordApp = new Word.Application
            {
                Visible = false,
                ScreenUpdating = false
                //Visible = true,
                //ScreenUpdating = true
            };
            var wDocument = wordApp.Documents.Open($@"{pdfFolder}\Содержание.docx");

            try
            {
                var table = wDocument.Tables[1];

                int titlePages = pagesInTitle;
                int startPageNumber = Convert.ToInt32(mf.StartNumberNumeric.Value) - 1;
                int pagesPzCount = Convert.ToInt32(mf.CountPagePZNumeric.Value);

                if ((startPageNumber + titlePages) % 2 == 1)
                {
                    titlePages++;
                }
                int page = startPageNumber + pagesPzCount + titlePages;
                table.Cell(2, 5).Range.Text = page.ToString();

                if (pagesPzCount % 2 == 1)
                {
                    pagesPzCount++;
                }
                page = startPageNumber + pagesPzCount + titlePages;

                int DataFilesCount = 0;
                int FileIndex = 0;
                int NumberOfPart = 0;
                if (mf.partsBookCheckBox.Checked)
                {
                    int rowInTable = table.Rows.Count;
                    for (var row = 1; row <= rowInTable; row++)
                    {
                        if (table.Cell(row, 2).Range.Text.Length > 3)
                        {
                            if (DataFilesCount != allDataFilesList.Count)
                            {
                                table.Cell(row, 6).Range.Text = allDataFilesList[DataFilesCount].Part.ToString();

                                if (NumberOfPart != allDataFilesList[DataFilesCount].Part)
                                {
                                    NumberOfPart = allDataFilesList[DataFilesCount].Part;
                                    FileIndex = 0;
                                }
                                if (NumberOfPart != 1)
                                {
                                    pagesPzCount = 0;
                                }
                                page = startPageNumber + pagesPzCount + titlePages;

                                table.Cell(row, 5).Range.Text = (page + firstPageNumbersList[NumberOfPart - 1][FileIndex]).ToString();
                                DataFilesCount++;
                                FileIndex++;
                            }
                        }
                    }
                }
                else
                {
                    int rowInTable = table.Rows.Count;
                    for (var row = 1; row <= rowInTable; row++)
                    {
                        if (table.Cell(row, 2).Range.Text.Length > 3)
                        {
                            if (DataFilesCount != allDataFilesList.Count)
                            {
                                table.Cell(row, 6).Range.Text = allDataFilesList[DataFilesCount].Part.ToString();

                                if (NumberOfPart != allDataFilesList[DataFilesCount].Part)
                                {
                                    NumberOfPart = allDataFilesList[DataFilesCount].Part;
                                    FileIndex = 0;
                                }
                                if (NumberOfPart != 1)
                                {
                                    pagesPzCount = 0;
                                }
                                page = startPageNumber + pagesPzCount + titlePages;

                                table.Cell(row, 5).Range.Text = (page + firstPageNumbersList[0][FileIndex]).ToString();
                                DataFilesCount++;
                                FileIndex++;
                            }
                        }
                    }
                }

                wDocument.Save();
                if (Directory.Exists($"{pdfFolder}\\Содержание.pdf"))
                {
                    Directory.Delete($"{pdfFolder}\\Содержание.pdf");
                }
                wDocument.ExportAsFixedFormat($"{pdfFolder}\\Содержание.pdf", Word.WdExportFormat.wdExportFormatPDF);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message.ToString());
                MessageBox.Show("Ошибка нумерации частей содержания", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                DeleteTempFiles();
                DeleteTempVar();
                mf.backgroundWorker2.ReportProgress(80, "Сборка остановлена");
                mf.backgroundWorker2.CancelAsync();
                return false;
            }
            finally
            {
                wDocument.Close(false);
                wordApp.Quit();
                GC.Collect();
            }

        }

        public void Attention()
        {
            foreach (Process clsProcess in Process.GetProcesses())
            {
                if (clsProcess.ProcessName.Equals("EXCEL"))
                {
                    DialogResult dialogResult = MessageBox.Show("Перед началом вам необходимо сохранить и закрыть все документы Excel.\n\n'ОК' - продолжить\n'Отмена' - выйти", "Отрыты документы Excel", MessageBoxButtons.OKCancel);
                    if (dialogResult == DialogResult.OK)
                    {
                        Attention();
                    }
                    else if (dialogResult == DialogResult.Cancel)
                    {
                        mf.Close();
                    }
                    break;
                }
            }

        }

        public void WordKiller()
        {
            foreach (Process p in System.Diagnostics.Process.GetProcessesByName("winword"))
            {
                try
                {
                    p.Kill();
                    p.WaitForExit(); // possibly with a timeout
                }
                catch (Win32Exception winException)
                {
                    // process was terminating or can't be terminated - deal with it
                }
                catch (InvalidOperationException invalidException)
                {
                    // process has already exited - might be able to let this one go
                }
            }
        }

        public void ExcelKiller()
        {
            foreach (Process clsProcess in Process.GetProcesses())
                if (clsProcess.ProcessName.Equals("EXCEL"))  //Process Excel?
                    clsProcess.Kill();
        }
    }
}