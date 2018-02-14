using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelGrinder
{
    public partial class Form1 : Form
    {

        string selectedPath;
        string[] files;
        DataTable DT = new DataTable();
        string surnameFile = string.Empty;

        string newFileName = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\ExcelOutFile_" + DateTime.Now.ToString("yyyy_MM_dd") + ".xlsx";

        public Form1()
        {
            InitializeComponent();
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
                            InfoLabel.Text = "Файл: " + openFileDialog1.FileName;
                            surnameFile = openFileDialog1.FileName;
                            GetDataFromRuleBook(wb);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }
            }
        }

        private void GetDataFromRuleBook(XSSFWorkbook wb)
        {
            var sheet = wb.GetSheet(wb.GetSheetName(0));
            var source = new BindingSource();

            DT.Rows.Clear();
            DT.Columns.Clear();

            for (int i = 0; i < sheet.LastRowNum; i++)
            {
                // add neccessary columns
                if (DT.Columns.Count < sheet.GetRow(i).Cells.Count)
                {
                    for (int j = 0; j < sheet.GetRow(i).Cells.Count; j++)
                    {
                        DT.Columns.Add("", typeof(string));
                    }
                }

                // add row
                DT.Rows.Add();

                // write row value
                for (int j = 0; j < sheet.GetRow(i).Cells.Count; j++)
                {
                    var cell = sheet.GetRow(i).GetCell(j);

                    if (cell != null)
                    {
                        // TODO: you can add more cell types capatibility, e. g. formula
                        switch (cell.CellType)
                        {
                            case NPOI.SS.UserModel.CellType.Numeric:
                                DT.Rows[i][j] = sheet.GetRow(i).GetCell(j).NumericCellValue;
                                //dataGridView1[j, i].Value = sh.GetRow(i).GetCell(j).NumericCellValue;

                                break;
                            case NPOI.SS.UserModel.CellType.String:
                                DT.Rows[i][j] = sheet.GetRow(i).GetCell(j).StringCellValue;

                                break;
                        }
                    }
                }
            }
            /*List<Object> rowList = new List<object>();
            for (int i = 0; i < sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                var cells = row.Cells.ToList();
                rowList.Add(cells);
            }

            source.DataSource = rowList;
            */
            ExcelRuleBookView.DataSource = DT;
        }
        #endregion
        #region ExcelGrind
        /* 1. Chose folder
         * 2. Go through list of names {
         * 3.    Go through excel files in folder {
         * 4.         Search name in files
         * 5.         After name found, find first and last line
         * 6.         Copy and past line range to a new file 
         * 7.         Show info in grid
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

        private void Grindbtn_Click(object sender, EventArgs e)
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
            if (DT.Rows.Count == 0)
            {
                MessageBox.Show("Файл с фамилиями не выбран или пуст");
                return;
            }
            


            foreach (DataRow row in DT.Rows)
            {
                if (row[2].ToString().Contains(value: "Цикловая") || row[2].ToString().Contains(value: "Студклуб"))
                {
                    continue;
                }

                foreach (var fileName in files)
                {
                    if (!Path.GetExtension(fileName).Contains("xls") || fileName == surnameFile)
                    {
                        continue;
                    }

                    ShowInfo("Ищу: " + row[2].ToString() + " в файле: " + fileName);
                    bool nameFound = SearchNameInFile(fileName, row[2].ToString());
                    if (nameFound)
                    {
                        FindRange(fileName);
                        CopyRange();
                        ShowInfoInGrid();
                    }                    
                }                
            }
        }

        private bool SearchNameInFile(string fileName, string name)
        {
            using (var fs = File.OpenRead(fileName))
            {
                XSSFWorkbook workBook = new XSSFWorkbook(fs);
                ISheet sheet = workBook.GetSheet(workBook.GetSheetName(0));
                for (int i = 0; i < sheet.LastRowNum; i++)
                {
                    if (sheet.GetRow(i).GetCell(1).StringCellValue.Trim().ToUpper() == name.Trim().ToUpper())
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        private void FindRange(string fileName)
        {
            throw new NotImplementedException();
        }

        private void CopyRange()
        {
            throw new NotImplementedException();
        }

        private void ShowInfoInGrid()
        {

        }
        #endregion

        #region Helpers
        private void ShowInfo(string message)
        {
            InfoLabel.Text = message;
        }

        #endregion
    }
}
