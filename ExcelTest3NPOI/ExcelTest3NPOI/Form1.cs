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
using System.Windows.Forms;

namespace ExcelTest3NPOI
{
    public partial class Form1 : Form
    {
        String pathFile = "";
        String pathDirection = "";
        String pathOutputDirection = "";
        String fileName = "";

        IWorkbook wbSource = null;

        public Form1()
        {
            InitializeComponent();
            this.Text += " Ver0.2.20160712.1";
        }

        public void fileProcess(String filePath)
        {
            int srcSheetCount;

            MemoryStream ms = new MemoryStream();

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
                ISheet usingSheet = wbSource.GetSheetAt(0);
                IRow rowHeader = usingSheet.GetRow(0);
                if (rowHeader != null) {
                    ICell cellHeader = rowHeader.GetCell(0);
                    cellHeader.SetCellValue("Test");
                }
                //usingSheet.CreateRow(0).CreateCell(0).SetCellValue("0");

                wbSource.Write(ms);

                using (FileStream fs = new FileStream("123.xlsx", FileMode.Create, FileAccess.Write))
                {
                    byte[] data = ms.ToArray();
                    fs.Write(data, 0, data.Length);
                    fs.Flush();
                    data = null;
                }
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
                    //debug output file name
                    System.Diagnostics.Debug.WriteLine("\r\nDebug Output:\r\n" + pathFile + "\r\n");
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
