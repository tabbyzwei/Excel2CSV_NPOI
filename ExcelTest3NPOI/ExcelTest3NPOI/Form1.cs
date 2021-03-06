﻿using GemBox.Spreadsheet;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace ExcelTest3NPOI
{
    public partial class Form1 : Form
    {
        String pathFile = "";
        String pathDirection = "";
        String pathOutputDirection = "";
        String fileName = "";
        String ErrorMessage = "";

        String[] headerTable = { "Status", "ODM", "Vendor","RMA No.","AcerP/N",
            "QTY","ODM U/P","Vendor C/N","Vendor U/P","MM#/ID"};

        IWorkbook wbSource = null;

        public Form1()
        {
            InitializeComponent();
            this.Text += " Ver0.2.20160712.2";
        }

        public void fileProcess(String filePath)
        {

            int srcSheetCount;

            FileInfo fiSource = new FileInfo(filePath);
            //if file exist
            if (fiSource.Exists)
            {
                //open file and get file stream
                FileStream fsSource = fiSource.OpenRead();
                if (fiSource.Extension.Equals(".xls"))
                {
                    wbSource = new HSSFWorkbook(fsSource);
                }
                else if (fiSource.Extension.Equals(".xlsx"))
                {
                    wbSource = new XSSFWorkbook(fsSource);
                }
                //close file stream
                fsSource.Close();

                //get sheet count
                srcSheetCount = wbSource.NumberOfSheets;
                //get Every Sheet
                for (int i = 0; i < srcSheetCount; i++)
                {
                    ISheet srcSheet = wbSource.GetSheetAt(i);
                    sheetProcess(srcSheet);
                }

                if (!ErrorMessage.Equals(""))
                {
                    MessageBox.Show(ErrorMessage);
                }

                //ISheet usingSheet = wbSource.GetSheetAt(0);
                //IRow rowHeader = usingSheet.GetRow(0);
                //if (rowHeader != null) {
                //    ICell cellHeader = rowHeader.GetCell(0);
                //    cellHeader.SetCellValue("Test");
                //}
                //usingSheet.CreateRow(0).CreateCell(0).SetCellValue("0");


                //using (FileStream fs = new FileStream(@"D:\123.xlsx", FileMode.Create, FileAccess.Write))
                //{
                //    wbSource.Write(fs);
                //}

                GC.Collect();
            }
        }
        public void sheetProcess(ISheet Sheet)
        {
            String strTemp = "";
            ICell srcCell;
            ICell tempCell;
            IRow tempRow;
            int tempRowIndex = 0;

            //test only
            int RowTotal = 0;
            int RowYellow = 0;

            //get sheet message
            sheetInfo si = sheetParse(Sheet);
            //create workbook
            XSSFWorkbook tempWorkbook = new XSSFWorkbook();
            ISheet tempSheet = tempWorkbook.CreateSheet(Sheet.SheetName + "_temp");

            //if error show error message
            if (!si.ErrorMessage.Equals(""))
            {
                ErrorMessage += Sheet.SheetName + ": " + si.ErrorMessage + "\r\n";
            }
            //no error,copy sheet to new sheet
            else
            {
                //Copy header message
                tempRow = tempSheet.CreateRow(tempRowIndex);
                tempRowIndex++;
                if (tempRow != null)
                {
                    for (int k = 0; k < si.RowIndex.Length; k++)
                    {
                        srcCell = Sheet.GetRow(0).GetCell(si.RowIndex[k]);
                        tempCell = tempRow.CreateCell(k);
                        if (srcCell != null && tempCell != null)
                        {
                            copyCellTo(tempCell, srcCell);
                        }
                    }
                }

                //find right Row
                for (int i = 1; i < si.RowMax; i++)
                {
                    //row = Sheet.GetRow(i);
                    srcCell = Sheet.GetRow(i).GetCell(si.RowIndex[0]);
                    if (srcCell != null)
                    {
                        RowTotal++;
                        strTemp = srcCell.StringCellValue;
                        ICellStyle style = srcCell.CellStyle;
                        IColor color = style.FillForegroundColorColor;
                        byte[] RGB = IndexedColors.Yellow.RGB;
                        if (color != null)
                        {
                            //if Cell color Equals Yellow
                            if (strTemp.Equals("open") && RGB[0] == color.RGB[0] && RGB[1] == color.RGB[1] && RGB[2] == color.RGB[2])
                            {
                                RowYellow++;
                                //if Vender CN not empty
                                srcCell = Sheet.GetRow(i).GetCell(si.RowIndex[7]);
                                if (srcCell != null)
                                {
                                    strTemp = srcCell.CellType.ToString();
                                    if (!strTemp.Equals("Blank"))
                                    {
                                        tempRow = tempSheet.CreateRow(tempRowIndex);
                                        tempRowIndex++;
                                        for (int k = 0; k < si.RowIndex.Length; k++)
                                        {
                                            srcCell = Sheet.GetRow(i).GetCell(si.RowIndex[k]);
                                            tempCell = tempRow.CreateCell(k);
                                            if (srcCell != null && tempCell != null)
                                            {
                                                copyCellTo(tempCell, srcCell);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            //auto fit
            for (int k = 0; k < si.RowIndex.Length; k++)
            {
                tempSheet.AutoSizeColumn(k);
            }
            //MessageBox.Show(Sheet.SheetName + "-- RowMax: " + si.RowMax + " RowFound: " + RowTotal + " RowYellow: " + RowYellow + "\r\n Row not lost: " + (RowTotal + 1 == si.RowMax));

            CellRangeAddress range = new CellRangeAddress(0,1,0,1);
            IAutoFilter filterTest =  tempSheet.SetAutoFilter(range);

            ISheet tempVenderSheet = tempWorkbook.CreateSheet(Sheet.SheetName + "_VenderCSV");



            var fileName = pathOutputDirection + "\\" + Sheet.SheetName + ".xlsx";
            var fileName2 = pathOutputDirection + "\\" + Sheet.SheetName + ".csv";
            //save xslx file
            using (FileStream fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                tempWorkbook.Write(fs);
            }

            //Only Convert first sheet(descending order)
            SpreadsheetInfo.SetLicense("EQU2-1000-0000-000U");
            ExcelFile ef = ExcelFile.Load(pathFile);

            ExcelFile ef2 = new ExcelFile();
            foreach (ExcelWorksheet item in ef.Worksheets)
            {
                if (item.Name.Equals(Sheet.SheetName))
                {
                    ExcelWorksheet ws = ef2.Worksheets.AddCopy(Sheet.SheetName, item);
                }
            }
            ef2.Save(fileName2, SaveOptions.CsvDefault);
        }
        public sheetInfo sheetParse(ISheet Sheet)
        {
            sheetInfo si = new sheetInfo();

            IRow row = Sheet.GetRow(0);
            si.RowMax = Sheet.LastRowNum;
            si.ListMax = row.LastCellNum;

            bool headerFoundFlag = false;

            String strTemp = "";
            for (int i = 0; i < headerTable.Length; i++)
            {
                //find header name in first row
                for (int j = 0; j < si.ListMax; j++)
                {
                    ICell cell = row.GetCell(j);
                    strTemp = cell.StringCellValue;
                    if (strTemp.Equals(headerTable[i]))
                    {
                        si.RowIndex[i] = j;
                        headerFoundFlag = true;
                        break;
                    }
                }
                //header name not find
                if (!headerFoundFlag)
                {
                    si.ErrorMessage += headerTable[i] + " No Found; ";
                }
                //reset flag
                headerFoundFlag = false;
            }
            return si;
        }
        //copy cell
        public void copyCellTo(ICell tarCell, ICell srcCell)
        {
            switch (srcCell.CellType)
            {
                case CellType.Numeric:
                    tarCell.SetCellValue(srcCell.NumericCellValue);
                    break;
                default:
                    tarCell.SetCellValue(srcCell.StringCellValue);
                    break;
            }
        }

        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            if (ofdOpenExcel.ShowDialog() == DialogResult.OK)
            {
                if (ofdOpenExcel.CheckFileExists)
                {
                    //get file path
                    pathFile = ofdOpenExcel.FileName;
                    //get Direction
                    pathDirection = System.IO.Path.GetDirectoryName(pathFile);
                    //get file name
                    fileName = System.IO.Path.GetFileNameWithoutExtension(pathFile);
                    //Creat new folder
                    pathOutputDirection = pathDirection + @"\Result_" + fileName;
                    if (!Directory.Exists(pathOutputDirection))
                    {
                        Directory.CreateDirectory(pathOutputDirection);
                    }
                    //debug output file
                    System.Diagnostics.Debug.WriteLine("\r\nDebug Output:\r\n" + pathOutputDirection + "\r\n");
                    //process the file
                    fileProcess(pathFile);
                }
                else
                {
                    MessageBox.Show("No File");
                }
            }
        }

        private void btnCreateFile_Click(object sender, EventArgs e)
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            MemoryStream ms = new MemoryStream();

            // 新增試算表。 
            ISheet sheet = workbook.CreateSheet("SheetA");
            //sheet.GetRow(0).GetCell(0);
            sheet.CreateRow(0).CreateCell(0).SetCellValue("想吃干馏抄手");
            workbook.CreateSheet("SheetB");
            workbook.CreateSheet("SheetC");

            workbook.Write(ms);
            using (FileStream fs = new FileStream(@"D:\想吃干馏抄手.xlsx", FileMode.Create, FileAccess.Write))
            {
                byte[] data = ms.ToArray();
                fs.Write(data, 0, data.Length);
                fs.Flush();
                data = null;
            }

            workbook = null;
            ms.Close();
            ms.Dispose();
            MessageBox.Show("想吃干馏抄手");
            //open explorer
            System.Diagnostics.Process.Start("explorer.exe", @"D:\");
        }
    }
}
