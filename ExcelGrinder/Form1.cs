using NPOI.HSSF.UserModel;
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
        private string selectedPath;
        private string[] files;
        private DataTable DT = new DataTable();
        private string surnameFile = string.Empty;
        private List<string> NotFoundSurnames = new List<string>();
        private Dictionary<string, int> Range = new Dictionary<string, int>();
        private int destinationRowNum = 0;

        private string newFileName = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\ExcelOutFile_" + DateTime.Now.ToString("yyyy_MM_dd") + ".xlsx";
        IWorkbook destinationWb = new XSSFWorkbook();

        private bool CancelAction = false;

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
                            ShowInfo("Файл: " + openFileDialog1.FileName);
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
            ISheet sheet = wb.GetSheet(wb.GetSheetName(0));
            BindingSource source = new BindingSource();

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
                            case CellType.Numeric:
                                DT.Rows[i][j] = sheet.GetRow(i).GetCell(j).NumericCellValue;

                                break;
                            case CellType.String:
                                DT.Rows[i][j] = sheet.GetRow(i).GetCell(j).StringCellValue;

                                break;
                        }
                    }
                }
            }
           
            ExcelRuleBookView.DataSource = DT;
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

            if (destinationWb.NumberOfSheets > 0)
            {
                destinationWb.RemoveSheetAt(0);
            }

            destinationWb.CreateSheet("OutPut");
            destinationRowNum = 0;

            foreach (DataRow row in DT.Rows)
            {
                var name = row[2].ToString();
                if (string.IsNullOrEmpty(name))
                    continue;

                if (name.Contains(value: "Цикловая") || name.Contains(value: "Студклуб"))
                {
                    continue;
                }
                bool isFound = false;
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

                    //TESTING STUB!!!
                    if (name.Trim() == "Бібікова Оксана Юріївна")
                    {
                        return;
                    }

                    ShowInfo("Ищу: " + name + " в файле: " + fileName);
                    using (var fs = File.OpenRead(fileName))
                    {
                        XSSFWorkbook workBook = new XSSFWorkbook(fs);
                        ISheet sheet = workBook.GetSheet(workBook.GetSheetName(0));
                        int rowNumber = SearchNameInFile(sheet, name);
                        if (rowNumber >= 0)
                        {
                            isFound = true;
                            FindRange(sheet, rowNumber);
                            CopyRange(sheet);
                            ShowInfoInGrid();
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
            WriteOutputFile();

            if (NotFoundSurnames.Count > 0)
            {
                WriteNotFoundFile(NotFoundSurnames);
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

        private void FindRange(ISheet sheet, int rowNumber)
        {
            Range.Clear();
            //Find start
            for (int i = rowNumber; i > sheet.FirstRowNum; i--)
            {
                //if (sheet.GetRow(i).GetCell(1).CellType == CellType.String)
                //{
                    var currentCell = sheet.GetRow(i).GetCell(0).StringCellValue.Trim();
                    if (sheet.GetRow(i).GetCell(0).StringCellValue.Trim().Contains("Розрахунковий"))
                    {
                        Range.Add("first", i);
                        break;
                    }
                //}
            }

            //Find finish
            for (int i = rowNumber; i < sheet.LastRowNum; i++)
            {
                if (sheet.GetRow(i).GetCell(1).CellType == CellType.String)
                {
                    if (sheet.GetRow(i).GetCell(1).StringCellValue.Trim().Contains("До видачі"))
                    {
                        Range.Add("last", i);
                        break;
                    }
                }
            }
        }

        private void CopyRange(ISheet sheet)
        {
            if (!Range.ContainsKey("first") || !Range.ContainsKey("last"))
            {
                MessageBox.Show("Не смог найти начало или конец диапазона копирования");
                return;
            }

            for (int sourceRowNum = Range["first"]; sourceRowNum <= Range["last"]; sourceRowNum++)
            {
                //read row
                IRow sourceRow = sheet.GetRow(sourceRowNum);
                IRow newRow = destinationWb.GetSheet(destinationWb.GetSheetName(0)).CreateRow(destinationRowNum);
                // Loop through source columns to add to new row
                for (int i = 0; i < sourceRow.LastCellNum; i++)
                {
                    // Grab a copy of the old/new cell
                    ICell oldCell = sourceRow.GetCell(i);
                    ICell newCell = newRow.CreateCell(i);

                    // If the old cell is null jump to next cell
                    if (oldCell == null)
                    {
                        newCell = null;
                        continue;
                    }
                    // Copy style from old cell and apply to new cell
                    ICellStyle newCellStyle = destinationWb.CreateCellStyle();
                    newCellStyle.CloneStyleFrom(oldCell.CellStyle); 
                    newCell.CellStyle = newCellStyle;

                    // If there is a cell comment, copy
                    if (newCell.CellComment != null) newCell.CellComment = oldCell.CellComment;

                    // If there is a cell hyperlink, copy
                    if (oldCell.Hyperlink != null) newCell.Hyperlink = oldCell.Hyperlink;

                    // Set the cell data type
                    newCell.SetCellType(oldCell.CellType);

                    // Set the cell data value
                    switch (oldCell.CellType)
                    {
                        case CellType.Blank:
                            newCell.SetCellValue(oldCell.StringCellValue);
                            break;
                        case CellType.Boolean:
                            newCell.SetCellValue(oldCell.BooleanCellValue);
                            break;
                        case CellType.Error:
                            newCell.SetCellErrorValue(oldCell.ErrorCellValue);
                            break;
                        case CellType.Formula:
                            newCell.SetCellFormula(oldCell.CellFormula);
                            break;
                        case CellType.Numeric:
                            newCell.SetCellValue(oldCell.NumericCellValue);
                            break;
                        case CellType.String:
                            newCell.SetCellValue(oldCell.RichStringCellValue);
                            break;
                        case CellType.Unknown:
                            newCell.SetCellValue(oldCell.StringCellValue);
                            break;
                    }

                }
                destinationRowNum++;
            }
                
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

        private void AddNotFound(string name)
        {
            NotFoundSurnames.Add(name);
        }

        private void WriteOutputFile()
        {
            using (FileStream stream = new FileStream(newFileName, FileMode.OpenOrCreate, FileAccess.Write))
            {
                destinationWb.Write(stream);
            }
        }

        private void WriteNotFoundFile(List<string> notFoundSurnames)
        {
            
        }

        private void Cancelbtn_Click(object sender, EventArgs e)
        {
            CancelAction = true;
        }


        #endregion
    }
}
