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
            IRow row;
            
            short indexColor;

            //get sheet message
            sheetInfo si = sheetParse(Sheet);
            //create workbook
            XSSFWorkbook workbook = new XSSFWorkbook();

            //if error show error message
            if (!si.ErrorMessage.Equals(""))
            {
                ErrorMessage += Sheet.SheetName + ": " + si.ErrorMessage + "\r\n";
            }
            //no error,copy sheet to new sheet
            else
            {
                for (int i = 1; i < si.RowMax; i++)
                {
                    //row = Sheet.GetRow(i);
                    ICell cell = Sheet.GetRow(i).GetCell(si.RowIndex[0]);
                    if (cell != null)
                    {
                        strTemp = cell.StringCellValue;
                        ICellStyle style = cell.CellStyle;
                        IColor color = style.FillForegroundColorColor;
                        byte[] RGB = IndexedColors.Yellow.RGB;
                        if (color != null)
                        {

                            if (strTemp.Equals("open") && RGB[0]==color.RGB[0] && RGB[1] == color.RGB[1] && RGB[2] == color.RGB[2])
                            {
                                //strTemp = row.GetCell(si.RowIndex[7]).StringCellValue;
                                if (!strTemp.Equals(""))
                                {

                                }
                            }
                        }
                    }
                }
            }
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
                    //output direction
                    pathOutputDirection = pathDirection + "\\" + fileName;
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
