using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Text.Json;
using System.Text.Json.Serialization;
using Aspose.Pdf;
using Aspose.Pdf.Text;

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
                        MessageBox.Show("���������� ����� ��������� ���������� ��������");
                        labelNameFolder.Text = "";
                        return;
                    }
                    else
                    {
                        labelNameFolder.Text = _path;

                        childFolder = new DirectoryInfo(dir[0]);
                        objectiveFiles = childFolder.GetFiles(".", SearchOption.TopDirectoryOnly);

                        infoTextBox.AppendText($"���-�� ���� ������: {localFiles.Length + objectiveFiles.Length}\n" + Environment.NewLine +
                            $"���-�� ������ � �������� �����: {localFiles.Length}\n" + Environment.NewLine +
                            $"���-�� ������ � �������� �����: {objectiveFiles.Length}\n" + Environment.NewLine +
                            $"���-�� �����: {dir.Length}" + Environment.NewLine + Environment.NewLine);

                        string[] fileNames = new string[localFiles.Length + objectiveFiles.Length];

                        Directory.GetFiles(_path, ".", SearchOption.TopDirectoryOnly).ToList()
                            .ForEach(f => infoTextBox.AppendText($"\n- {Path.GetFileName(f)}" + Environment.NewLine));

                        infoTextBox.AppendText(Environment.NewLine + $"\n\n����� {dir[0]}" + Environment.NewLine);

                        Directory.GetFiles(dir[0], ".", SearchOption.TopDirectoryOnly).ToList()
                            .ForEach(f => infoTextBox.AppendText($"\n{Path.GetFileName(f)}"));



                    }


                }

            }
        }

        //TODO ���������� ������� ������� � PDF

        private void BtnBuild_Click(object sender, EventArgs e)
        {

            if (localFiles != null)
            {
                int countCompleted = 0;
                Excel.Application app = new Excel.Application();
                foreach (var fileExcel in localFiles)
                {

                    System.IO.Directory.CreateDirectory($"{_path}\\TEMPdf");
                    string filePath = $"{_path}\\{fileExcel}";
                    Excel.Workbook workbook = app.Workbooks.Open(filePath);
                    string tempPDFPath = $"{_path}\\TEMPdf\\{fileExcel}";

                    app.ActiveWorkbook.Sheets[1].PageSetup.FirstPageNumber = 5; // ��������� ��������� ��������
                    app.DisplayAlerts = false; // ��������� ������� �� ����������
                    app.ActiveWorkbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, tempPDFPath); /// ��������� Excel to PDF
                    workbook.Close(false); // ������ ���������� � ����� Excel
                    countCompleted++;
                    labelCompleted.Text = $"���-�� ������������ ������: {countCompleted} �� {localFiles.Length + objectiveFiles.Length}";
                }



                object[,] arrLocalData = new object[localFiles.Length, 5];
                object[,] arrObjectiveData = new object[objectiveFiles.Length, 5];

                Excel.Application ObjWorkExcel = new Excel.Application();

                int countPages = 1;
                int documentNumber = 1;

                if (startNumberTextBox.Text != "")
                {
                    countPages = Convert.ToInt32(startNumberTextBox.Text);
                }

                infoTextBox.Text += Environment.NewLine + Environment.NewLine;

                for (int i = 0; i < objectiveFiles.Length; i++) /// ������ ��� ��������� ����
                {
                    string filePath = $"{childFolder}\\{objectiveFiles[i]}";
                    Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open($@"{filePath}");
                    Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

                    int j = 0;
                    arrObjectiveData[i, j] = documentNumber;/// ����� ���������
                    j++;
                    Regex regex = new Regex(@"(\w*)-(\w*)-(\w*)");
                    MatchCollection match = regex.Matches(ObjWorkSheet.Range["B10"].Value.ToString());
                    arrObjectiveData[i, j] = match[0];/// ��� �����
                    infoTextBox.AppendText(arrObjectiveData[i, j] + Environment.NewLine);
                    j++;
                    arrObjectiveData[i, j] = ObjWorkSheet.Range["B7"].Value;/// ������������
                    infoTextBox.AppendText(arrObjectiveData[i, j] + Environment.NewLine);
                    j++;
                    arrObjectiveData[i, j] = ObjWorkSheet.Range["F14"].Value;/// ����� �����
                    infoTextBox.AppendText(arrObjectiveData[i, j] + Environment.NewLine);
                    j++;
                    arrObjectiveData[i, j] = countPages; /// ����� ������ �������� 
                    infoTextBox.AppendText("��������: " + arrObjectiveData[i, j] + Environment.NewLine);

                    countPages += ObjWorkBook.Sheets[1].PageSetup.Pages.Count; /// ���-�� ������� �� �����

                    ObjWorkBook.Close();

                    documentNumber++;
                }

                for (int i = 0; i < localFiles.Length; i++) // TODO ���������� ������ ��� ��������� ����
                {
                    string filePath = $"{rootFolder}\\{localFiles[i]}";
                    Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open($@"{filePath}");
                    Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

                    int j = 0;
                    arrLocalData[i, j] = i + 1;
                    j++;
                    Regex regex = new Regex(@"(\w*)-(\w*)-(\w*)");
                    MatchCollection match = regex.Matches(ObjWorkSheet.Range["A18"].Value.ToString());
                    arrLocalData[i, j] = match[0];
                    infoTextBox.Text += arrLocalData[i, j] + Environment.NewLine;
                    j++;
                    arrLocalData[i, j] = ObjWorkSheet.Range["A20"].Value;
                    infoTextBox.Text += arrLocalData[i, j] + Environment.NewLine;
                    j++;
                    arrLocalData[i, j] = ObjWorkSheet.Range["C28"].Value;
                    infoTextBox.Text += arrLocalData[i, j] + Environment.NewLine;
                    j++;
                    arrLocalData[i, j] = countPages; /// TODO ����� ��������(�� ������)
                    infoTextBox.Text += "��������: " + arrLocalData[i, j] + Environment.NewLine;

                    countPages += ObjWorkBook.Sheets[1].PageSetup.Pages.Count; // TODO ���-�� ������� �� �����
                    ObjWorkBook.Close();
                    documentNumber++;
                }
                ObjWorkExcel.Quit();
                app.Quit();

                //string newPDFPath = $"{_path}\\TEMPdf";
                //pdfFolder = new DirectoryInfo(newPDFPath);
                //FileInfo[] tempPDFFiles = pdfFolder.GetFiles(".", SearchOption.TopDirectoryOnly);

                /*License PdfMergeLicense = new License();
                PdfMergeLicense.SetLicense("Aspose.pdf.lic");

                PdfFileEditor pdfFileEditor = new PdfFileEditor();

                string[] pdffiles = new string[3];
                pdffiles[0] = $"{newPDFPath}\\{tempPDFFiles[0]}";
                pdffiles[1] = $"{newPDFPath}\\{tempPDFFiles[1]}";
                pdffiles[2] = $"{newPDFPath}\\{tempPDFFiles[2]}";

                pdfFileEditor.Concatenate(pdffiles, $"{newPDFPath}\\MergedPDF.pdf");*/

                /*Document pdfDocument1 = new Document($"{newPDFPath}\\{tempPDFFiles[0]}");
                Document pdfDocument2 = new Document(newPDFPath + "\\" + tempPDFFiles[1]);

                pdfDocument1.Pages.Add(pdfDocument2.Pages);

                pdfDocument1.Save(newPDFPath + "\\ConcatenatedPDF.pdf");*/

                /*Document pdfDocument1 = new Document(newPDFPath + "\\" + tempPDFFiles[0]);
                Document pdfDocument2 = new Document(newPDFPath + "\\" + tempPDFFiles[1]);
                pdfDocument1.Pages.Add(pdfDocument2.Pages);
                pdfDocument1.Save(newPDFPath + "ConcatenatedPDF.pdf");*/
                //Document document = new Document();
                //Page page = document.Pages.Add();
                //page.Paragraphs.Add(new Aspose.Pdf.Text.TextFragment("����� �� �����"));
                //document.Save(_path + "\\documentTest.pdf");


                /*Document pdfDocument1 = new Document(newPDFPath + "\\" + tempPDFFiles[0]);
                Document pdfDocument2 = new Document(newPDFPath + "\\" + tempPDFFiles[1]);
                pdfDocument1.Pages.Add(pdfDocument2.Pages);
                pdfDocument1.Save(newPDFPath + "ConcatenatedPDF.pdf");*/
                //Document document = new Document();
                //Page page = document.Pages.Add();
                //page.Paragraphs.Add(new Aspose.Pdf.Text.TextFragment("����� �� �����"));
                //document.Save(_path + "\\documentTest.pdf");

            }
            else
            {
                MessageBox.Show($"������! �� �� ������� �����");
            }


        }

        // TODO JSON
        //public class Local
        //{
        //    public string Code { get; set; }
        //    public string Name { get; set; }
        //    public string Price { get; set; }
        //    public int CountPages { get; set; }
        //}

        //public class Objective
        //{
        //    public string Code { get; set; }
        //    public string Name { get; set; }
        //    public string Price { get; set; }
        //    public int CountPages { get; set; }
        //}

        //public class Root
        //{
        //    public List<Objective> objective { get; set; }
        //    public List<Local> local { get; set; }
        //}

        // -------

        //var root = new Root
        //{
        //    objective = new List<Objective>
        //        {
        //            new Objective { Code = "1", Name = "2", Price = "2", CountPages = 1 },
        //            new Objective { Code = "1", Name = "2", Price = "2", CountPages = 1 },
        //        },

        //    local = new List<Local>
        //        {
        //            new Local { Code = "1", Name = "2", Price = "2", CountPages = 1 },
        //            new Local { Code = "1", Name = "2", Price = "2", CountPages = 1},
        //        }
        //};

        //var options = new JsonSerializerOptions { WriteIndented = true };
        //string jsonString = JsonSerializer.Serialize(root, options);

        //MessageBox.Show(jsonString);



        private void Read_Button_Click(object sender, EventArgs e) // TODO �������� ����������
        {
            Regex regex = new Regex(@"(\w*)-(\w*)-(\w*)");
            var s = localFiles.OrderBy(a => regex.Matches(a.ToString())[0]);
            //var sortedByCode = localFiles.OrderBy(a => regex.Matches(a.ToString())[0]).ToArray();
            //MessageBox.Show(String.Join(" ", sortedByCode.ToString()));
            MessageBox.Show("");

        }

        
    }
}



