using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace ExcelGrinder
{
    public partial class Form1 : Form
    {
        ExcelGrinderModel Model = new ExcelGrinderModel();

        private string selectedPath;
        private string[] files;
        private System.Data.DataTable surnameDT = new System.Data.DataTable();
        private System.Data.DataTable infoDT = new System.Data.DataTable();
        private string surnameFile = string.Empty;
        private List<string> NotFoundSurnames = new List<string>();
        
        private string newFileName = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\ExcelOutFile_" + DateTime.Now.ToString("yyyy_MM_dd") + ".xlsx";
        XSSFWorkbook destinationWb = new XSSFWorkbook();

        private bool CancelAction = false;

        public Form1()
        {
            InitializeComponent();
            ExcelGrinderModel Model = new ExcelGrinderModel();
        }
        #region SurnameList
        /* 1. Chose file with list
         * 2. read file and show it in the grid
         */
        private void ChoseFilebtn_Click(object sender, EventArgs e)
        {
            Stream myStream = null;
            surnameFile = string.Empty;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "Excel document (*.xls)|*.xlsx";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if ((myStream = openFileDialog1.OpenFile()) != null)
                    {
                        using (myStream)
                        {
                            XSSFWorkbook wb = new XSSFWorkbook(myStream);
                            ShowInfo("Файл: " + openFileDialog1.FileName);
                            surnameFile = openFileDialog1.FileName;
                            surnameDT = Model.GetDataFromRuleBook(wb);
                            ExcelRuleBookView.DataSource = surnameDT;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }
            }
        }
            
        #endregion

        #region ExcelGrind
        /* 1. Chose folder
         * 2. Create Destination WorkBook
         * 3. Go through list of names {
         * 4.    Go through excel files in folder {
         * 5.         Search name in files
         * 6.         After name found, find first and last line
         * 7.         Copy and past line range to a new file 
         * 8.         Show info in grid
         *       } 
         *       if name not found, write it in the file and inform user
         *    }    
         */

        private void ChoseFolderbtn_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {

                fbd.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    files = Directory.GetFiles(fbd.SelectedPath);

                    ShowInfo("В папке: " + fbd.SelectedPath + " найдено " + files.Length.ToString() + " файлов.");
                    selectedPath = fbd.SelectedPath;
                }
            }
        }

        private async void Grindbtn_ClickAsync(object sender, EventArgs e)
        {
            InfoLabel.Text = selectedPath;
            if (string.IsNullOrEmpty(selectedPath))
            {
                MessageBox.Show("Папка с файлами не выбрана");
                return;
            }
            if (files.Length == 0)
            {
                MessageBox.Show("Папка с файлами пуста");
                return;
            }
            if (surnameDT.Rows.Count == 0)
            {
                MessageBox.Show("Файл с фамилиями не выбран или пуст");
                return;
            }

            Model.ClearDestinationWb();
            
            foreach (DataRow row in surnameDT.Rows)
            {
                var name = row[2].ToString();
                if (string.IsNullOrEmpty(name))
                    continue;

                if (name.Contains(value: "Цикловая") || name.Contains(value: "Студклуб"))
                {
                    continue;
                }
                bool isFound = false;

                //TESTING STUB!!!
                /*if (name.Trim() == "Бібікова Оксана Юріївна")
                {
                    break;
                } */
                
                foreach (var fileName in files)
                {
                    if (CancelAction)
                    {
                        CancelAction = false;
                        break;
                    }

                    if (!Path.GetExtension(fileName).Contains("xls") || fileName == surnameFile)
                    {
                        continue;
                    }


                    ShowInfo("Ищу: " + name.Trim() + " в файле: " + fileName.Trim());
                    using (var fs = File.OpenRead(fileName))
                    {
                        XSSFWorkbook workBook = new XSSFWorkbook(fs);
                        ISheet sheet = workBook.GetSheet(workBook.GetSheetName(0));
                        int rowNumber = SearchNameInFile(sheet, name);
                        if (rowNumber >= 0)
                        {
                            isFound = true;

                            await Task.Run(() => Model.CopyPeople(workBook, sheet, rowNumber));
                            continue;
                        }
                    }
                }
                if (!isFound)
                {
                    AddNotFound(name);
                }
            }

            // Create file and write WorkBook          

            Model.WriteOutputFile();
            ExcelOutputView.DataSource = Model.ShowInfoInGrid();

            if (NotFoundSurnames.Count > 0)
            {
                Model.WriteNotFoundFile(NotFoundSurnames);
            }
        }

        private int SearchNameInFile(ISheet sheet, string name)
        {
            for (int i = 0; i < sheet.LastRowNum; i++)
            {
                try
                {
                    if (sheet.GetRow(i).GetCell(1).CellType == CellType.String)
                    {
                        if (sheet.GetRow(i).GetCell(1).StringCellValue.Trim().ToUpper() == name.Trim().ToUpper())
                        {
                            return i;
                        }
                    }
                }
                catch { }
            }
            return -1;
        }
  
       
        #endregion

        #region Helpers
        private void ShowInfo(string message)
        {
            InfoLabel.Text = message;
        }

        private void AddNotFound(string name)
        {
            NotFoundSurnames.Add(name);
        }

        private void Cancelbtn_Click(object sender, EventArgs e)
        {
            CancelAction = true;
        }
        #endregion

        private void TestExcelbtn_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            if (excel == null)
            {
                MessageBox.Show("Excel не установлен или версии не совпадают!");
                return;
            }

            if (files == null || files.Length == 0)
            {
                MessageBox.Show("Выберите папку с файлами Excel");
                return;
            }

            Workbook wb = excel.Workbooks.Open(files[0]);
            Worksheet sheet = wb.Worksheets.get_Item(1);
            Range sourceRange = sheet.Rows["1:100"];
            sourceRange.Copy();

            Workbook wbDest = excel.Workbooks.Add();
            Worksheet sheetDest = wbDest.Worksheets.get_Item(1);
            Range rangeDest = sheetDest.Rows["1:100"];
            rangeDest.PasteSpecial(XlPasteType.xlPasteAll);

            wbDest.SaveAs(newFileName);
            wb.Close(true);
            wbDest.Close(true);
            excel.Quit();

            Marshal.ReleaseComObject(wb);
            Marshal.ReleaseComObject(wb);
            Marshal.ReleaseComObject(wbDest);
            Marshal.ReleaseComObject(wbDest);
            Marshal.ReleaseComObject(excel);

            MessageBox.Show("Тест прошел успешно, тестовый файл сохранен в: " + newFileName);
        }
    }
}
