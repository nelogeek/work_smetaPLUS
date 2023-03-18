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
        private readonly ProgramFunctions PF = new ProgramFunctions();
        
        public static MainForm instance; //Singleton

        public MainForm()
        {
            if(instance == null)
                instance = this;
            InitializeComponent();

            backgroundWorker.WorkerReportsProgress = true;
            backgroundWorker.WorkerSupportsCancellation = true;
        }

        private void BtnSelectFolder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog selectedPatch = new FolderBrowserDialog();

            if (selectedPatch.ShowDialog() == DialogResult.OK)
            {
                PF.path = selectedPatch.SelectedPath; //Указание пути к корневой папке
                PF.rootFolder = new DirectoryInfo(PF.path);
                PF.localFiles = PF.rootFolder.GetFiles(".", SearchOption.TopDirectoryOnly); //Сбор локальных файлов

                PF.pdfFolder = new DirectoryInfo($"{PF.path}\\TEMPdf"); //Указание пути к папке с временными файлами
                PF.finalSmetaFolder = new DirectoryInfo($"{PF.path}\\Книга смет"); //Указание пути к итоговой папке

                PF.DeleteTempFiles();

                foreach (var file in PF.localFiles) //Проверка расширения файлов
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

                PF.dirFolders = Directory.GetDirectories(PF.path); //Сбор информации и обработка папок
                if (PF.dirFolders.Length == 0)
                {
                    MessageBox.Show("В корневой директории отсутствуют папки, книга будет сгенерирована без ОС");
                    PF.SelectFolder();
                }
                else if (PF.dirFolders.Length == 1)
                {
                    if (PF.dirFolders[0] == $"{PF.path}\\ОС" || PF.dirFolders[0] == $"{PF.path}\\OC")
                    {
                        PF.SelectFolder();
                    }
                    else if (PF.dirFolders[0] == $"{PF.path}\\Книга смет")
                    {
                        MessageBox.Show("Книга будет сгенерирована без ОС");
                        PF.SelectFolder();
                    }
                    else
                    {
                        MessageBox.Show("В корневом разделе неправльная папка, исправьте название или удалите");
                        return;
                    }
                }
                else if (PF.dirFolders.Length == 2)
                {
                    for(int i = 0; i < 2; i++)
                    {
                        if (!(PF.dirFolders[i] == $"{PF.path}\\ОС" || PF.dirFolders[i] == $"{PF.path}\\OC" || PF.dirFolders[i] == $"{PF.path}\\Книга смет"))
                        {
                            MessageBox.Show("В корневом разделе неправильная папка");
                            return;
                        }
                    }
                    PF.SelectFolder();
                }
                else if (PF.dirFolders.Length == 3) //Переделать для TEMP PDF
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
            if (PF.fullBookPageCount > 400 && !partsBookCheckBox.Checked) //Проверка на слишком большое количество страниц
            {
                DialogResult dialogResult = MessageBox.Show("Вы точно хотите собрать одну книгу объемом более 400 страниц", "Подтверждение создания книги", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.No)
                {
                    return;
                }
            }
            if (backgroundWorker.IsBusy != true)
            {
                PF.DisableButtons();
                backgroundWorker.RunWorkerAsync();
                buildProgressBar.Visible = true;
            }
        }


        protected void RunBackgroundWorker_DoWork() //Запуск сборки
        {
            PF.stopWatch.Start();

            backgroundWorker.ReportProgress(1, "Парсинг");
            backgroundWorker.ProgressChanged += BackgroundWorker_ProgressChanged;
            if (!PF.ExcelParser()) return;

            backgroundWorker.ReportProgress(10, "Конвертация");
            backgroundWorker.ProgressChanged += BackgroundWorker_ProgressChanged;
            if (!PF.ExcelConverter()) return;

            backgroundWorker.ReportProgress(40, "Создание финальный папки");
            backgroundWorker.ProgressChanged += BackgroundWorker_ProgressChanged;
            if (!PF.CreateFinalSmetaFolder()) return;

            backgroundWorker.ReportProgress(45, "Генерация содержания");
            backgroundWorker.ProgressChanged += BackgroundWorker_ProgressChanged;
            if (!PF.TitleGeneration()) return;

            backgroundWorker.ReportProgress(65, "Сборка книги");
            backgroundWorker.ProgressChanged += BackgroundWorker_ProgressChanged;
            if (!PF.PdfMerge()) return;

            backgroundWorker.ReportProgress(80, "Нумерация частей содержания");
            backgroundWorker.ProgressChanged += BackgroundWorker_ProgressChanged;
            if (!PF.TitleNumOfPart()) return;

            backgroundWorker.ReportProgress(85, "Перемещение файлов");
            backgroundWorker.ProgressChanged += BackgroundWorker_ProgressChanged;
            if (!PF.MoveFiles()) return;

            backgroundWorker.ReportProgress(90, "Удаление временных файлов");
            backgroundWorker.ProgressChanged += BackgroundWorker_ProgressChanged;
            PF.DeleteTempFiles();

            backgroundWorker.ReportProgress(95, "Удаление временных переменных");
            backgroundWorker.ProgressChanged += BackgroundWorker_ProgressChanged;
            PF.DeleteTempVar();

            backgroundWorker.ReportProgress(100, "Сборка завершена");
            backgroundWorker.ProgressChanged += BackgroundWorker_ProgressChanged;

            PF.stopWatch.Stop();
            TimeSpan ts = PF.stopWatch.Elapsed;
            PF.stopWatch.Reset();
            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);
            backgroundWorker.ReportProgress(100, $"Время сборки: {elapsedTime}");
        }

        public void BackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled == true)
            {
                labelProgressStage.Text = "Отмена!";
                PF.EnableButtons();
            }
            else if (e.Error != null)
            {
                labelProgressStage.Text = "Ошибка: " + e.Error.Message;
                PF.EnableButtons();
            }
            else
            {
                PF.EnableButtons();
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
            if (PF.path != null)
            {
                if (Directory.Exists($"{PF.finalSmetaFolder.FullName}"))
                {
                    DialogResult dialogResult = MessageBox.Show("Вы хотите заменить папку 'Книга смет'?", "Подтверждение замены папки", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        Directory.Delete(PF.finalSmetaFolder.FullName, true);
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
                    PF.DeleteTempVar();
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