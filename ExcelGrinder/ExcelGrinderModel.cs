using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using NPOI.SS.Util;
using System.IO;

namespace ExcelGrinder
{
    class ExcelGrinderModel
    {
        private System.Data.DataTable surnameDT = new System.Data.DataTable();
        private System.Data.DataTable infoDT = new System.Data.DataTable();
        private string surnameFile = string.Empty;
        private List<string> NotFoundSurnames = new List<string>();
        private int destinationRowNum = 0;

        private string newFileName = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\ExcelOutFile_" + DateTime.Now.ToString("yyyy_MM_dd") + ".xlsx";
        XSSFWorkbook destinationWb = new XSSFWorkbook();

        Dictionary<string, int> Range = new Dictionary<string, int>();

        private bool CancelAction = false;

        public DataTable GetDataFromRuleBook(XSSFWorkbook wb)
        {
            ISheet sheet = wb.GetSheet(wb.GetSheetName(0));
            BindingSource source = new BindingSource();

            surnameDT.Rows.Clear();
            surnameDT.Columns.Clear();

            for (int i = 0; i < sheet.LastRowNum; i++)
            {
                // add neccessary columns
                if (surnameDT.Columns.Count < sheet.GetRow(i).Cells.Count)
                {
                    for (int j = 0; j < sheet.GetRow(i).Cells.Count; j++)
                    {
                        surnameDT.Columns.Add("", typeof(string));
                    }
                }

                // add row
                surnameDT.Rows.Add();

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
                                surnameDT.Rows[i][j] = sheet.GetRow(i).GetCell(j).NumericCellValue;

                                break;
                            case CellType.String:
                                surnameDT.Rows[i][j] = sheet.GetRow(i).GetCell(j).StringCellValue;

                                break;
                        }
                    }
                }
            }
            return surnameDT;            
        }


        public async Task CopyPeople(XSSFWorkbook workBook, ISheet sheet, int rowNumber)
        {
            Range = FindRange(sheet, rowNumber);
            SetColumnsWidth(sheet);
            CopyRange(sheet, workBook);
        }

        public Dictionary<string, int> FindRange(ISheet sheet, int rowNumber)
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

            return Range;
        }

        public void SetColumnsWidth(ISheet sourceSheet)
        {
            ISheet destSheet = destinationWb.GetSheet(destinationWb.GetSheetName(0));
            for (int i = 0; i < sourceSheet.GetRow(0).Cells.Count; i++)
            {
                destSheet.SetColumnWidth(i, sourceSheet.GetColumnWidth(i));
            }
            //Dirty hack
            destSheet.SetColumnWidth(4, 1);
            destSheet.SetColumnWidth(8, 1);
        }

        public void CopyRange(ISheet sheet, IWorkbook wb)
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
                    XSSFCell oldCell = (XSSFCell)sourceRow.GetCell(i);
                    XSSFCell newCell = (XSSFCell)newRow.CreateCell(i);

                    // If the old cell is null jump to next cell
                    if (oldCell == null)
                    {
                        newCell = null;
                        continue;
                    }
                    // Copy style from old cell and apply to new cell
                    XSSFCellStyle newCellStyle = (XSSFCellStyle)destinationWb.CreateCellStyle();
                    //newCellStyle.CloneStyleFrom(oldCell.CellStyle);


                    //Borders
                    CopyBordersStyle(oldCell, newCellStyle);


                    //Text Style
                    CopyTextStyle(oldCell, newCellStyle);

                    //Font
                    CopyFontStyle(wb, oldCell, newCellStyle);

                    //newCellStyle.CloneStyleFrom(oldCell.CellStyle);

                    newCell.CellStyle = newCellStyle;


                    //NPOI.SS.UserModel.IFont cellFont = oldCell.CellStyle.GetFont();
                    //newCell.CellStyle.SetFont(cellFont);


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

                // If there are are any merged regions in the source row, copy to new row
                for (int i = 0; i < sheet.NumMergedRegions; i++)
                {
                    CellRangeAddress cellRangeAddress = sheet.GetMergedRegion(i);
                    if (cellRangeAddress.FirstRow == sourceRow.RowNum)
                    {
                        CellRangeAddress newCellRangeAddress = new CellRangeAddress(newRow.RowNum,
                                                                                    (newRow.RowNum +
                                                                                     (cellRangeAddress.FirstRow -
                                                                                      cellRangeAddress.LastRow)),
                                                                                    cellRangeAddress.FirstColumn,
                                                                                    cellRangeAddress.LastColumn);
                        destinationWb.GetSheet(destinationWb.GetSheetName(0)).AddMergedRegion(newCellRangeAddress);
                    }
                }
                destinationRowNum++;
            }

            destinationRowNum++;

        }

        private void CopyFontStyle(IWorkbook wb, XSSFCell oldCell, XSSFCellStyle newCellStyle)
        {
            NPOI.SS.UserModel.IFont font = destinationWb.CreateFont();
            NPOI.SS.UserModel.IFont sourceFont = oldCell.CellStyle.GetFont(wb);
            font.FontName = sourceFont.FontName;
            font.FontHeightInPoints = sourceFont.FontHeightInPoints;
            font.Boldweight = sourceFont.Boldweight;
            newCellStyle.SetFont(font);
        }

        private static void CopyTextStyle(XSSFCell oldCell, XSSFCellStyle newCellStyle)
        {
            newCellStyle.WrapText = oldCell.CellStyle.WrapText;
            newCellStyle.ShrinkToFit = oldCell.CellStyle.ShrinkToFit;
            newCellStyle.Alignment = oldCell.CellStyle.Alignment;
            newCellStyle.VerticalAlignment = oldCell.CellStyle.VerticalAlignment;
        }

        private static void CopyBordersStyle(XSSFCell oldCell, XSSFCellStyle newCellStyle)
        {
            byte[] rgb = new byte[3] { 0, 0, 0 };
            newCellStyle.BorderBottom = oldCell.CellStyle.BorderBottom;
            newCellStyle.SetBottomBorderColor(new XSSFColor(rgb));
            newCellStyle.BorderLeft = oldCell.CellStyle.BorderLeft;
            newCellStyle.SetLeftBorderColor(new XSSFColor(rgb));
            newCellStyle.BorderTop = oldCell.CellStyle.BorderTop;
            newCellStyle.SetTopBorderColor(new XSSFColor(rgb));
            newCellStyle.BorderRight = oldCell.CellStyle.BorderRight;
            newCellStyle.SetRightBorderColor(new XSSFColor(rgb));
        }

        public void ClearDestinationWb()
        {
            if (destinationWb.NumberOfSheets > 0)
            {
                destinationWb.RemoveSheetAt(0);
            }

            destinationWb.CreateSheet("OutPut");
        }

        public void WriteOutputFile()
        {
            using (FileStream stream = new FileStream(newFileName, FileMode.OpenOrCreate, FileAccess.Write))
            {
                destinationWb.Write(stream);
            }
        }

        public DataTable ShowInfoInGrid()
        {
            ISheet sheet = destinationWb.GetSheet(destinationWb.GetSheetName(0));
            BindingSource source = new BindingSource();

            infoDT.Rows.Clear();
            infoDT.Columns.Clear();

            for (int i = 0; i < sheet.LastRowNum; i++)
            {
                // add neccessary columns
                try
                {
                    if (infoDT.Columns.Count < sheet.GetRow(i).Cells.Count)
                    {
                        for (int j = 0; j < sheet.GetRow(i).Cells.Count; j++)
                        {
                            infoDT.Columns.Add("", typeof(string));
                        }
                    }

                    // add row
                    infoDT.Rows.Add();

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
                                    infoDT.Rows[i][j] = sheet.GetRow(i).GetCell(j).NumericCellValue;

                                    break;
                                case CellType.String:
                                    infoDT.Rows[i][j] = sheet.GetRow(i).GetCell(j).StringCellValue;

                                    break;
                            }
                        }
                    }
                }
                catch { }
            }

            return infoDT;
        }

        public void WriteNotFoundFile(List<string> notFoundSurnames)
        {
            string notFoundFilePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\NotFound.txt";
            System.IO.File.WriteAllLines(notFoundFilePath, notFoundSurnames);
            MessageBox.Show("Файл с ненайденными фамилиями был записан: " + notFoundFilePath);
        }
    }
}
