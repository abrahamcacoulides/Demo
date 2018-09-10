using System;
using System.Collections.Generic;
using System.Linq;
using Engine.EventArgs;
using Engine.Models;
using System.IO;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.Util;
using System.Threading.Tasks;

namespace Engine.ViewModels
{
    public class Session : BaseNotificationClass
    {
        public event EventHandler<MessageEventArgs> OnMessageRaised;

        //local list of SO(sales order; jobs!!)
        static List<SO> jobsList = new List<SO>();

        static List<Record> Records = new List<Record>();

        static List<Record> Records_With_Added_PO = new List<Record>();

        static List<ToAddRecord> RecordsToAdd = new List<ToAddRecord>();

        // local path to _billsFolder
        public string _billsPath;

        // local path to S file
        public string _sPath;

        // local path to P file
        public string _pPath;

        // local path to Results File
        public string _resultsPath;

        // local path to Material Agregado Path
        public string _materialAgregadoPath;

        // local path to which the Results file would be saved
        public string _resultsSaveToPath;

        // local path to Final Results File
        public string _resultsSaveToPath1;

        public string _savedIN;

        private void RaiseMessage(string message)
        {
            OnMessageRaised?.Invoke(this, new MessageEventArgs(message));
        }

        public void GoButton()
        {
            if (Directory.Exists(_billsPath))
            {
                List<string> facturas = GetExcelFilesInDirectory(_billsPath);
                ReadBillsInPath(facturas,2,3,10,9,8,7);
                if (_resultsSaveToPath == "")
                {
                    _resultsSaveToPath = @"newbook.core.xlsx";
                }
                Facturas(_resultsSaveToPath);
                RaiseMessage("Results file saved to: " + _resultsSaveToPath);
            }
            else
            {
                RaiseMessage("Please select the path to the bills folder...");
            }
        }

        public void GoButton1()
        {
            if (File.Exists(_resultsPath) && Path.GetExtension(_resultsPath) == ".xlsx")
            {
                if (File.Exists(_pPath) && Path.GetExtension(_pPath) == ".xlsx")
                {
                    if (File.Exists(_sPath) && Path.GetExtension(_sPath) == ".xlsx")
                    {
                        if (File.Exists(_materialAgregadoPath) && Path.GetExtension(_materialAgregadoPath) == ".xlsx")
                        {
                            if (_resultsSaveToPath1 == "")
                            {
                                _resultsSaveToPath1 = @"newbookResults.core.xlsx";
                            }
                            ReadFromResultsFile(_resultsPath);
                            MCO(_sPath,3,14,1,10);
                            LoopThroughItemsToAdd(_pPath, 4, 1, 11, 10);
                            LoopThroughItemsToAdd(_materialAgregadoPath, 4, 1, 11, 10);
                            ItemsToAdd();
                            Results(_resultsSaveToPath1);
                            RaiseMessage("Results saved to: " + _resultsSaveToPath1);
                        }
                        else
                        {
                            RaiseMessage("There seems to be a problem with the 'Material Agregado' File, confirm your selection is an .xlsx file");
                        }
                    }
                    else
                    {
                        RaiseMessage("There seems to be a problem with the 'S' File, confirm your selection is an .xlsx file");
                    }
                }
                else
                {
                    RaiseMessage("There seems to be a problem with the 'P' File, confirm your selection is an .xlsx file");
                }
            }
            else
            {
                RaiseMessage("There seems to be a problem with the 'Results' File, confirm your selection is an .xlsx file");
            }
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

        public void ReadBillsInPath(List<string> bills, int jobNumCol, int poCol, int costCol, int addedValueCol, int weightCol, int umCol)
        {
            int count = 1;
            XSSFWorkbook billwb;

            foreach (string bill in bills)
            {
                if (bill.Contains("~$"))
                {
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
                        else if (current_sheet.GetRow(sheet_count).GetCell(jobNumCol).CellType != CellType.Numeric || current_sheet.GetRow(sheet_count).GetCell(poCol).CellType != CellType.Numeric)
                        {
                            sheet_count += 1;
                            blanks += 1;
                        }
                        else if (current_sheet.GetRow(sheet_count).GetCell(jobNumCol) != null && current_sheet.GetRow(sheet_count).GetCell(poCol) != null)
                        {
                            blanks = 0;
                            string product = current_sheet.GetRow(sheet_count).GetCell(jobNumCol).NumericCellValue.ToString() + '-' + current_sheet.GetRow(sheet_count).GetCell(poCol).NumericCellValue.ToString();
                            if( jobsList.FirstOrDefault(j => j.PO == current_sheet.GetRow(sheet_count).GetCell(poCol).NumericCellValue.ToString()) != null)
                            {
                            }
                            else
                            {
                                string job_num = current_sheet.GetRow(sheet_count).GetCell(jobNumCol).NumericCellValue.ToString();
                                string po = current_sheet.GetRow(sheet_count).GetCell(poCol).NumericCellValue.ToString();
                                double cost = current_sheet.GetRow(sheet_count).GetCell(costCol).NumericCellValue;
                                double addedValue = current_sheet.GetRow(sheet_count).GetCell(addedValueCol).NumericCellValue;
                                double weight = current_sheet.GetRow(sheet_count).GetCell(weightCol).NumericCellValue;
                                string um = current_sheet.GetRow(sheet_count).GetCell(umCol).ToString();
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

        private static void Facturas(string saveAs)
        {
            using (var fs = new FileStream(saveAs, FileMode.Create, FileAccess.Write))
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
                    current_row.CreateCell(2).SetCellValue(job.Cost); // costo (2)
                    current_row.GetCell(2).SetCellType(CellType.Numeric);
                    current_row.CreateCell(3).SetCellValue(job.AddedValue); //valor agregado (3)
                    current_row.GetCell(3).SetCellType(CellType.Numeric);
                    current_row.CreateCell(4).SetCellValue(job.Weight); // peso (4)
                    current_row.GetCell(4).SetCellType(CellType.Numeric);
                    current_row.CreateCell(5).SetCellValue(job.UM); //medida!!
                    current_row.CreateCell(6).SetCellValue(job.PO); //po_only!!
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
        }
        
        public static void MCO(string spreadsheet, int startingRow, int poCol, int materialCol, int qtyCol)
        {
            XSSFWorkbook mcoWB;

            using (FileStream file = new FileStream(spreadsheet, FileMode.Open, FileAccess.Read))
            {
                mcoWB = new XSSFWorkbook(file);
            }

            ISheet mco_sheet = mcoWB.GetSheetAt(0);
            int starting_row_current_sheet = startingRow;
            int starting_row_current_sheet_1 = startingRow;

            var positive_items = new Dictionary<string, double>();

            // this while goes to the MCO WB and records in the positive items dictionary the items with a positive value in the qty
            while (mco_sheet.GetRow(starting_row_current_sheet) != null)
            {
                {
                    if (mco_sheet.GetRow(starting_row_current_sheet).GetCell(qtyCol).NumericCellValue > 0)
                    {
                        // the dictionary key is po-material#
                        string key = mco_sheet.GetRow(starting_row_current_sheet).GetCell(poCol).NumericCellValue.ToString() + '-' + mco_sheet.GetRow(starting_row_current_sheet).GetCell(materialCol).StringCellValue;
                        double qty = mco_sheet.GetRow(starting_row_current_sheet).GetCell(qtyCol).NumericCellValue;
                        if (positive_items.ContainsKey(key))
                        {
                            positive_items[key] += qty;
                        }
                        else
                        {
                            positive_items.Add(key, qty);
                        }
                    }
                    starting_row_current_sheet++;
                }
            }

            // the following while loops through all the records in the MCO excel and if the po is in the bill po's list it checks if the qty value is negative, 
            // if it is it then proceeds to look for the item in the positive_items dict,
            // if it exist it compares the value in the dict and the qty, if they are equal it deletes the record and goes to the next record in the file
            // if the qty and value in the dict are not equal it adds the current record to the list of objects
            // in case it is not in the positive_items dict it proceeds to add the current record to the list of objects
            while (mco_sheet.GetRow(starting_row_current_sheet_1) != null)
            {
                string po = mco_sheet.GetRow(starting_row_current_sheet_1).GetCell(poCol).NumericCellValue.ToString();

                double qty = mco_sheet.GetRow(starting_row_current_sheet_1).GetCell(qtyCol).NumericCellValue;

                string material = mco_sheet.GetRow(starting_row_current_sheet_1).GetCell(materialCol).StringCellValue;

                // the dictionary key is po-material
                string key = po + '-' + material;
                SO job = jobsList.FirstOrDefault(j => j.PO == po);
                if (job != null)
                {
                    
                    if (qty < 0)
                    {
                        if (positive_items.ContainsKey(key))
                        {
                            if (positive_items[key] * -1 == qty)
                            {
                                positive_items.Remove(key);
                            }
                            else
                            {
                                //write to excel
                                Records.Add(
                                    new Record
                                    {
                                        Product = job.Product,
                                        Material = material,
                                        Qty = qty * -1,
                                        UM = mco_sheet.GetRow(starting_row_current_sheet_1).GetCell(11).StringCellValue
                                    });
                            }
                        }
                        else
                        {
                            //write to excel
                            Records.Add(
                                    new Record
                                    {
                                        Product = job.Product,
                                        Material = material,
                                        Qty = qty * -1,
                                        UM = mco_sheet.GetRow(starting_row_current_sheet_1).GetCell(11).StringCellValue
                                    });
                        }
                    }
                }
                starting_row_current_sheet_1++;
            }
        }

        public static void LoopThroughItemsToAdd(string spreadsheet, int startingRow, int materialCol, int uMCol, int qtyCol)
        {
            XSSFWorkbook itemsToAddWB;

            using (FileStream file = new FileStream(spreadsheet, FileMode.Open, FileAccess.Read))
            {
                itemsToAddWB = new XSSFWorkbook(file);
            }

            ISheet itemsToAddSheet = itemsToAddWB.GetSheetAt(0);

            List<ToAddRecord> _recordsToAdd = new List<ToAddRecord>();

            while (itemsToAddSheet.GetRow(startingRow) != null)
            {
                string material = itemsToAddSheet.GetRow(startingRow).GetCell(materialCol).StringCellValue;

                string um = itemsToAddSheet.GetRow(startingRow).GetCell(uMCol).StringCellValue;

                double qty = itemsToAddSheet.GetRow(startingRow).GetCell(qtyCol).NumericCellValue;

                if (_recordsToAdd.Exists(r => r.Material == material))
                {
                    if (_recordsToAdd.Where(r => r.Material == material).ToList().Exists(r => r.UM == um))
                    {
                        _recordsToAdd.FirstOrDefault(r => r.Material == material && r.UM == um).Qty += qty;
                    }
                    else
                    {
                        _recordsToAdd.Add(new ToAddRecord
                        {
                            Material = material,
                            UM = um,
                            Qty = qty
                        });
                    }
                }
                else
                {
                    _recordsToAdd.Add(new ToAddRecord
                    {
                        Material = material,
                        UM = um,
                        Qty = qty
                    });
                }
                startingRow++;
            }

            foreach (ToAddRecord record in _recordsToAdd.Where(r => r.Qty < 0))
            {
                RecordsToAdd.Add(new ToAddRecord
                {
                    Material = record.Material,
                    UM = record.UM,
                    Qty = record.Qty * -1
                });
            }
        }

        public static void ItemsToAdd()
        {
            int Record_index = 0;
            foreach (ToAddRecord itemToAdd in RecordsToAdd)
            {
                if ((itemToAdd.Qty) / jobsList.Count() >= 1)
                {
                    int qty_per_order = Convert.ToInt32(itemToAdd.Qty / jobsList.Count());
                    int remaining = Convert.ToInt32(itemToAdd.Qty % jobsList.Count());
                    double decimals = itemToAdd.Qty % 1;
                    int i = 0;
                    foreach (SO job in jobsList)
                    {
                        if (i == 0)
                        {
                            Records_With_Added_PO.Add(
                                   new Record
                                   {
                                       Product = job.Product,
                                       Material = itemToAdd.Material,
                                       Qty = qty_per_order + 1 + decimals,
                                       UM = itemToAdd.UM
                                   });
                        }
                        else if (i > 0 && i < remaining)
                        {
                            Records_With_Added_PO.Add(
                                new Record
                                {
                                    Product = job.Product,
                                    Material = itemToAdd.Material,
                                    Qty = qty_per_order + 1,
                                    UM = itemToAdd.UM
                                });
                        }
                        else
                        {
                            Records_With_Added_PO.Add(
                                new Record
                                {
                                    Product = job.Product,
                                    Material = itemToAdd.Material,
                                    Qty = qty_per_order,
                                    UM = itemToAdd.UM
                                });
                        }
                        i++;
                    }
                }
                else
                {
                    int remaining = Convert.ToInt32(itemToAdd.Qty % jobsList.Count());
                    double decimals = itemToAdd.Qty % 1;
                    int i = 0;
                    while (i < remaining)
                    {
                        if (i == 0)
                        {
                            Records_With_Added_PO.Add(
                                   new Record
                                   {
                                       Product = jobsList[i].Product,
                                       Material = itemToAdd.Material,
                                       Qty = 1 + decimals,
                                       UM = itemToAdd.UM
                                   });
                        }
                        else
                        {
                            Records_With_Added_PO.Add(
                                new Record
                                {
                                    Product = jobsList[i].Product,
                                    Material = itemToAdd.Material,
                                    Qty = 1,
                                    UM = itemToAdd.UM
                                });
                        }
                        i++;
                        if (Record_index + 1 == jobsList.Count)
                        {
                            Record_index = 0;
                        }
                        else
                        {
                            Record_index++;
                        }
                    }
                }
            }
        }

        private static void ReadFromResultsFile(string filePath)
        {
            XSSFWorkbook ResultsBook;

            using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                ResultsBook = new XSSFWorkbook(file);
            }

            ISheet current_sheet = ResultsBook.GetSheetAt(0);

            jobsList.Clear();

            int row_count = 1;

            while (current_sheet.GetRow(row_count) != null)
            {
                string product = current_sheet.GetRow(row_count).GetCell(0).StringCellValue;
                double cost = current_sheet.GetRow(row_count).GetCell(2).NumericCellValue;
                double addedValue = current_sheet.GetRow(row_count).GetCell(3).NumericCellValue;
                double weight = current_sheet.GetRow(row_count).GetCell(4).NumericCellValue;
                string um = current_sheet.GetRow(row_count).GetCell(5).StringCellValue;
                string po = current_sheet.GetRow(row_count).GetCell(6).StringCellValue;
                string factura = current_sheet.GetRow(row_count).GetCell(7).StringCellValue;
                string job_num = product.Split('-')[0];
                jobsList.Add(new SO(product, job_num, po, cost, addedValue, weight, um, factura));
                row_count += 1;
            }
        }

        private static void Results(string saveAs)
        {

            using (var fs = new FileStream(saveAs, FileMode.Create, FileAccess.Write))
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
                    current_row.CreateCell(2).SetCellValue(job.Cost); // costo (2)
                    current_row.GetCell(2).SetCellType(CellType.Numeric);
                    current_row.CreateCell(3).SetCellValue(job.AddedValue); //valor agregado (3)
                    current_row.GetCell(3).SetCellType(CellType.Numeric);
                    current_row.CreateCell(4).SetCellValue(job.Weight); // peso (4)
                    current_row.GetCell(4).SetCellType(CellType.Numeric);
                    current_row.CreateCell(5).SetCellValue(job.UM); //medida!!
                    current_row.CreateCell(6).SetCellValue(job.PO); //po_only!!
                    current_row.CreateCell(7).SetCellValue(job.Factura); //factura!!
                    row_count += 1;
                }

                ISheet results_sheet = workbook.CreateSheet("Structured_BOM");

                IRow result_headers = results_sheet.CreateRow(0);
                result_headers.CreateCell(0).SetCellValue("Product");
                result_headers.CreateCell(1).SetCellValue("Material");
                result_headers.CreateCell(2).SetCellValue("Qty");
                result_headers.CreateCell(3).SetCellValue("UM");
                result_headers.CreateCell(4).SetCellValue("Comment");

                int count = 1;
                foreach (Record record in Records)
                {
                    IRow current_row = results_sheet.CreateRow(count);
                    current_row.CreateCell(0).SetCellValue(record.Product);
                    current_row.CreateCell(1).SetCellValue(record.Material);
                    current_row.CreateCell(2).SetCellValue(record.Qty.ToString());
                    current_row.CreateCell(3).SetCellValue(record.UM);
                    current_row.CreateCell(4).SetCellValue("N/A");
                    count += 1;
                }

                foreach (Record record in Records_With_Added_PO)
                {
                    IRow current_row = results_sheet.CreateRow(count);
                    current_row.CreateCell(0).SetCellValue(record.Product);
                    current_row.CreateCell(1).SetCellValue(record.Material);
                    current_row.CreateCell(2).SetCellValue(record.Qty.ToString());
                    current_row.CreateCell(3).SetCellValue(record.UM);
                    current_row.CreateCell(4).SetCellValue("Added");
                    count += 1;
                }

                IRow headersRow = current_sheet.GetRow(0);
                for (int i = 0; i < 8; i++)
                {
                    current_sheet.AutoSizeColumn(i);
                    var cellToFormat = headersRow.GetCell(i);
                    cellToFormat.CellStyle = headerStyle;
                    cellToFormat.CellStyle.SetFont(headerFont);
                }
                IRow resultsheadersRow = results_sheet.GetRow(0);
                for (int i = 0; i < 4; i++)
                {
                    results_sheet.AutoSizeColumn(i);
                    var results_cellToFormat = resultsheadersRow.GetCell(i);
                    results_cellToFormat.CellStyle = headerStyle;
                    results_cellToFormat.CellStyle.SetFont(headerFont);
                }
                workbook.Write(fs);
            }
        }
    }
}
