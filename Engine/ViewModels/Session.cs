using System;
using System.Collections.Generic;
using System.Linq;
using Engine.EventArgs;
using Engine.Models;
using System.IO;
using System.Diagnostics;
using System.Threading;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.Util;

namespace Engine.ViewModels
{
    public class Session : BaseNotificationClass
    {
        public event EventHandler<MessageEventArgs> OnMessageRaised;

        //local list of SO(sales order; jobs!!)
        static List<SO> jobsList = new List<SO>();

        // local path to _billsFolder
        public string _billsPath;

        private void RaiseMessage(string message)
        {
            OnMessageRaised?.Invoke(this, new MessageEventArgs(message));
        }

        public void GoButton()
        {
            List<string> facturas = GetExcelFilesInDirectory(_billsPath);
            
            ReadBillsInPath(facturas);
            //foreach(SO job in jobsList)
            //{
            //    RaiseMessage(job.PO);
            //}
            Facturas();
            RaiseMessage("Done!");
        }

        public static List<string> GetExcelFilesInDirectory(string targetDirectory)
        {
            List<string> filesList = new List<string>();
            string[] fileEntries = Directory.GetFiles(targetDirectory, "*.xlsx");
            foreach (string file in fileEntries)
            {
                filesList.Add(file);
            }
            return filesList;
        }

        public void ReadBillsInPath(List<string> bills)
        {
            int count = 1;
            XSSFWorkbook billwb;

            foreach (string bill in bills)
            {
                if (bill.Contains("~$"))
                {
                    RaiseMessage("Why?: " + bill);
                }
                else
                {
                    using (FileStream file = new FileStream(bill, FileMode.Open, FileAccess.Read))
                    {
                        billwb = new XSSFWorkbook(file);
                    }

                    ISheet current_sheet = billwb.GetSheetAt(0);

                    int sheet_count = 25;
                    int blanks = 0;
                    while (sheet_count < 500)
                    {
                        if (blanks == 3)
                        {
                            sheet_count += 1;
                            break;
                        }
                        else if(current_sheet.GetRow(sheet_count) == null)
                        {
                            sheet_count += 1;
                            blanks += 1;
                        }
                        else if (current_sheet.GetRow(sheet_count).GetCell(2).CellType != CellType.Numeric || current_sheet.GetRow(sheet_count).GetCell(3).CellType != CellType.Numeric)
                        {
                            sheet_count += 1;
                            blanks += 1;
                        }
                        else if (current_sheet.GetRow(sheet_count).GetCell(2) != null && current_sheet.GetRow(sheet_count).GetCell(3) != null)
                        {
                            blanks = 0;
                            string product = current_sheet.GetRow(sheet_count).GetCell(2).NumericCellValue.ToString() + '-' + current_sheet.GetRow(sheet_count).GetCell(3).NumericCellValue.ToString();
                            if( jobsList.FirstOrDefault(j => j.PO == current_sheet.GetRow(sheet_count).GetCell(3).NumericCellValue.ToString()) != null)
                            {
                                Console.WriteLine("PO: " + current_sheet.GetRow(sheet_count).GetCell(3).NumericCellValue.ToString() + " repeated! on bill: " + bill);
                            }
                            else
                            {
                                string job_num = current_sheet.GetRow(sheet_count).GetCell(2).NumericCellValue.ToString();
                                string po = current_sheet.GetRow(sheet_count).GetCell(3).NumericCellValue.ToString();
                                double cost = current_sheet.GetRow(sheet_count).GetCell(10).NumericCellValue;
                                double addedValue = current_sheet.GetRow(sheet_count).GetCell(9).NumericCellValue;
                                double weight = current_sheet.GetRow(sheet_count).GetCell(8).NumericCellValue;
                                string um = current_sheet.GetRow(sheet_count).GetCell(7).ToString();
                                string factura = bill;
                                jobsList.Add(new SO(product, job_num, po, cost, addedValue, weight, um, factura));
                                count += 1;
                            }
                            sheet_count += 1;
                        }
                        else
                        {
                            sheet_count += 1;
                            blanks += 1;
                        }
                    }
                }
            }
        }

        private static void Facturas()
        {
            var newFile = @"newbook.core.xlsx";

            using (var fs = new FileStream(newFile, FileMode.Create, FileAccess.Write))
            {
                IWorkbook workbook = new XSSFWorkbook();
                ISheet current_sheet = workbook.CreateSheet("Results");

                var headerStyle = workbook.CreateCellStyle();
                headerStyle.FillForegroundColor = HSSFColor.Grey80Percent.Index;
                headerStyle.FillPattern = FillPattern.SolidForeground;
                var headerFont = workbook.CreateFont();
                headerFont.Color = HSSFColor.White.Index;
                headerFont.IsBold = true;

                IRow headers = current_sheet.CreateRow(0);
                headers.CreateCell(0).SetCellValue("Producto");
                headers.CreateCell(1).SetCellValue("Fraccion");
                headers.CreateCell(2).SetCellValue("Costo");
                headers.CreateCell(3).SetCellValue("Valor Agregado");
                headers.CreateCell(4).SetCellValue("Peso");
                headers.CreateCell(5).SetCellValue("Medida");
                headers.CreateCell(6).SetCellValue("Po");
                headers.CreateCell(7).SetCellValue("Factura");

                int row_count = 1;
                foreach (SO job in jobsList)
                {
                    IRow current_row = current_sheet.CreateRow(row_count);
                    current_row.CreateCell(0).SetCellValue(job.Product); //producto (0)
                    current_row.CreateCell(2).SetCellValue(job.Cost.ToString()); // costo (2)
                    current_row.CreateCell(3).SetCellValue(job.AddedValue.ToString()); //valor agregado (3)
                    current_row.CreateCell(4).SetCellValue(job.Weight.ToString()); // peso (4)
                    current_row.CreateCell(5).SetCellValue(job.UM); //medida!!
                    current_row.CreateCell(6).SetCellValue(job.PO.ToString()); //po_only!!
                    current_row.CreateCell(7).SetCellValue(job.Factura); //factura!!
                    row_count += 1;
                }

                IRow headersRow = current_sheet.GetRow(0);
                for (int i=  0 ; i<8;i++)
                {
                    current_sheet.AutoSizeColumn(i);
                    var cellToFormat = headersRow.GetCell(i);
                    cellToFormat.CellStyle = headerStyle;
                    cellToFormat.CellStyle.SetFont(headerFont);
                }
                workbook.Write(fs);
            }
            Console.WriteLine("Excel  Done");
        }
    }
}
