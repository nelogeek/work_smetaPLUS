using ExcelApp.Functions;
using System;
using System.ComponentModel;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ExcelAPP
{
    public partial class MainForm : Form
    {
        public static MainForm instance; //Singleton

        private readonly ProgramFunctions pf;

        public MainForm()
        {
            if (instance == null)
                instance = this;
            InitializeComponent();

            pf = new ProgramFunctions();

            backgroundWorker.WorkerReportsProgress = true;
            backgroundWorker.WorkerSupportsCancellation = true;
        }

        private void BtnSelectFolder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog selectedPatch = new FolderBrowserDialog();

            if (selectedPatch.ShowDialog() == DialogResult.OK)
            {
                pf.path = selectedPatch.SelectedPath; //Указание пути к корневой папке
                pf.rootFolder = new DirectoryInfo(pf.path);
                pf.localFiles = pf.rootFolder.GetFiles(".", SearchOption.TopDirectoryOnly); //Сбор локальных файлов

                pf.pdfFolder = new DirectoryInfo($"{pf.path}\\TEMPdf"); //Указание пути к папке с временными файлами
                pf.finalSmetaFolder = new DirectoryInfo($"{pf.path}\\Книга смет"); //Указание пути к итоговой папке

                pf.DeleteTempFiles();

                foreach (var file in pf.localFiles) //Проверка расширения файлов
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

                pf.dirFolders = Directory.GetDirectories(pf.path); //Сбор информации и обработка папок
                if (pf.dirFolders.Length == 0)
                {
                    MessageBox.Show("В корневой директории отсутствуют папки, книга будет сгенерирована без ОС");
                    pf.SelectFolder();
                }
                else if (pf.dirFolders.Length == 1)
                {
                    if (pf.dirFolders[0] == $"{pf.path}\\ОС" || pf.dirFolders[0] == $"{pf.path}\\OC")
                    {
                        pf.SelectFolder();
                    }
                    else if (pf.dirFolders[0] == $"{pf.path}\\Книга смет")
                    {
                        MessageBox.Show("Книга будет сгенерирована без ОС");
                        pf.SelectFolder();
                    }
                    else
                    {
                        MessageBox.Show("В корневом разделе неправльная папка, исправьте название или удалите");
                        return;
                    }
                }
                else if (pf.dirFolders.Length == 2)
                {
                    for (int i = 0; i < 2; i++)
                    {
                        if (!(pf.dirFolders[i] == $"{pf.path}\\ОС" || pf.dirFolders[i] == $"{pf.path}\\OC" || pf.dirFolders[i] == $"{pf.path}\\Книга смет"))
                        {
                            MessageBox.Show("В корневом разделе неправильная папка");
                            return;
                        }
                    }
                    pf.SelectFolder();
                }
                else if (pf.dirFolders.Length == 3) //Переделать для TEMP PDF
                {
                    MessageBox.Show("В корневом разделе находятся лишние папки");
                    return;
                }
                else
                {
                    MessageBox.Show("В корневом разделе находятся лишние папки");
                    return;
                }
            }
        }

        private void BtnBuild_Click(object sender, EventArgs e)
        {
            if (pf.fullBookPageCount > 400 && !partsBookCheckBox.Checked) //Проверка на слишком большое количество страниц
            {
                DialogResult dialogResult = MessageBox.Show("Вы точно хотите собрать одну книгу объемом более 400 страниц", "Подтверждение создания книги", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.No)
                {
                    return;
                }
            }
            if (backgroundWorker.IsBusy != true)
            {
                pf.DisableButtons();
                backgroundWorker.RunWorkerAsync();
                buildProgressBar.Visible = true;
            }
        }


        protected void RunBackgroundWorker_DoWork() //Запуск сборки
        {
            pf.stopWatch.Start();

            backgroundWorker.ReportProgress(1, "Парсинг");
            backgroundWorker.ProgressChanged += BackgroundWorker_ProgressChanged;
            if (!pf.ExcelParser()) return;

            backgroundWorker.ReportProgress(10, "Конвертация");
            backgroundWorker.ProgressChanged += BackgroundWorker_ProgressChanged;
            if (!pf.ExcelConverter()) return;

            backgroundWorker.ReportProgress(40, "Создание финальный папки");
            backgroundWorker.ProgressChanged += BackgroundWorker_ProgressChanged;
            if (!pf.CreateFinalSmetaFolder()) return;

            backgroundWorker.ReportProgress(45, "Генерация содержания");
            backgroundWorker.ProgressChanged += BackgroundWorker_ProgressChanged;
            if (!pf.TitleGeneration()) return;

            backgroundWorker.ReportProgress(65, "Сборка книги");
            backgroundWorker.ProgressChanged += BackgroundWorker_ProgressChanged;
            if (!pf.PdfMerge()) return;

            backgroundWorker.ReportProgress(80, "Нумерация частей содержания");
            backgroundWorker.ProgressChanged += BackgroundWorker_ProgressChanged;
            if (!pf.TitleNumOfPart()) return;

            backgroundWorker.ReportProgress(85, "Перемещение файлов");
            backgroundWorker.ProgressChanged += BackgroundWorker_ProgressChanged;
            if (!pf.MoveFiles()) return;

            backgroundWorker.ReportProgress(90, "Удаление временных файлов");
            backgroundWorker.ProgressChanged += BackgroundWorker_ProgressChanged;
            pf.DeleteTempFiles();

            backgroundWorker.ReportProgress(95, "Удаление временных переменных");
            backgroundWorker.ProgressChanged += BackgroundWorker_ProgressChanged;
            pf.DeleteTempVar();

            backgroundWorker.ReportProgress(100, "Сборка завершена");
            backgroundWorker.ProgressChanged += BackgroundWorker_ProgressChanged;

            pf.stopWatch.Stop();
            TimeSpan ts = pf.stopWatch.Elapsed;
            pf.stopWatch.Reset();
            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);
            backgroundWorker.ReportProgress(100, $"Время сборки: {elapsedTime}");
        }

        public void BackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled == true)
            {
                labelProgressStage.Text = "Отмена!";
                pf.EnableButtons();
            }
            else if (e.Error != null)
            {
                labelProgressStage.Text = "Ошибка: " + e.Error.Message;
                pf.EnableButtons();
            }
            else
            {
                pf.EnableButtons();
                infoTextBox.Clear();
            }
        }

        public void BackgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            labelProgressStage.Text = e.UserState.ToString();
            buildProgressBar.Value = e.ProgressPercentage;
        }

        public void BackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            if (pf.path != null)
            {
                if (Directory.Exists($"{pf.finalSmetaFolder.FullName}"))
                {
                    DialogResult dialogResult = MessageBox.Show("Вы хотите заменить папку 'Книга смет'?", "Подтверждение замены папки", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        Directory.Delete(pf.finalSmetaFolder.FullName, true);
                        RunBackgroundWorker_DoWork();
                    }
                    else
                    {
                        backgroundWorker.ReportProgress(0, "Сборка остановлена");
                        return;
                    }
                }
                else
                {
                    RunBackgroundWorker_DoWork();
                }
            }
            else
            {
                MessageBox.Show($"Ошибка! Вы не выбрали папку");
                backgroundWorker.ReportProgress(0, "Сборка остановлена");
                return;
            }
        }


        private void MainForm_FormClosing(object sender, FormClosingEventArgs e) // Закрытие программы
        {
            try
            {
                DialogResult dialogResult = MessageBox.Show("Вы точно хотите закрыть программу?", "Подтверждение закрытия программы", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    pf.DeleteTempVar();
                    e.Cancel = false;
                }
                else
                {
                    e.Cancel = true;
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка завершения программы");
            }
        }

        private void AutoBooksPartPassCheckBox_Click(object sender, EventArgs e)
        {
            if (AutoBooksPartPassCheckBox.Checked)
            {
                dividerPagesCountLabel.Enabled = false;
                dividerPassPagesCount.Enabled = false;
            }
            else
            {
                dividerPagesCountLabel.Enabled = true;
                dividerPassPagesCount.Enabled = true;
            }
        }

        private void TestBtn_Click(object sender, EventArgs e)
        {

        }
    }
}