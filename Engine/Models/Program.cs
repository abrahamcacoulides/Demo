using System;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;
using System.Collections.Generic;
using System.Text;
using NPOI.XSSF;
using NPOI.XSSF.UserModel;
using Example1.Models;
using System.Linq;

namespace Engine
{
    public class Program
    {
        static void Main(string[] args)
        {
            //datos a solicitar al user
            string facturas_path = @"Data\Facturas\2018 Facturas\2018-06 Packing List Junio 2018\";
            var material_con_orden_path = @"Data\S 2018-06.xlsx";
            string material_sin_orden_path = @"Data\P 2018-06.xlsx";
            string material_agregado_path = @"Data\Material Agregado 2018-06.xlsx";

            // Display a simple message to the user.
            Console.WriteLine("***** My First C# App *****");
            Console.WriteLine("Hello World!");
            Console.WriteLine();

            InitializeTemplateWorkbook();

            if (Directory.Exists(facturas_path))
            {
                // This path is a directory
                Console.WriteLine("El directorio seleccionado si existe!");
                List<string> facturas = GetExcelFilesInDirectory(facturas_path);
                WriteBillsToExcel(facturas);
            }
            else
            {
                Console.WriteLine("{ 0 } is not a valid directory.", facturas_path);
            }
            MCO(material_con_orden_path);

            Console.WriteLine("The number of items with po(before):" + Records.Count().ToString());

            //ToDo cuales son los valores??!!
            LoopThroughItemsToAdd(material_sin_orden_path, 4, 1, 11, 10);
            LoopThroughItemsToAdd(material_agregado_path, 4, 1, 11, 10);
            //
            ItemsToAdd();
            Console.WriteLine("The number of items with po:" + Records.Count().ToString());
            Console.WriteLine("The number of items without po:" + RecordsToAdd.Count().ToString());
            Console.WriteLine("The number of items from bills:" + jobpo_dict.Count().ToString());
            Console.WriteLine("The number of rows in Structured BOM's:" + Records_With_Added_PO.Count().ToString());
            Console.WriteLine("Finished!");
            WriteResultsToExcel();
            WriteToFile();
            // Wait for Enter key to be pressed before shutting down.
            Console.WriteLine("Finished!");
            Console.ReadLine();
        }

        static XSSFWorkbook hssfWorkbook;

        static List<Record> Records = new List<Record>();
        static List<Record> Records_With_Added_PO = new List<Record>();

        static List<ToAddRecord> RecordsToAdd = new List<ToAddRecord>();

        static void WriteToFile()
        {
            //Write the workbook’s data stream to the root directory **Might be a good a idea to give the user the option to how and where save it??
            FileStream file = new FileStream(@"results_weight_cost.xlsx", FileMode.Create);
            hssfWorkbook.Write(file);

            file.Close();
        }

        static void InitializeTemplateWorkbook()
        {
            using (var fs = File.OpenRead(@"Data\results_weight_cost.xlsx"))
            {
                hssfWorkbook = new XSSFWorkbook(fs);
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

        //local list of job-po found in Bills
        static Dictionary<string, string> jobpo_dict = new Dictionary<string, string>();

        //weight_cost_from_bill
        public static void WriteBillsToExcel(List<string> bills)
        {
            ISheet results_sheet = hssfWorkbook.GetSheet("Results");
            int count = 1;

            foreach (string bill in bills)
            {
                if(bill.Contains("~$"))
                {
                    Console.WriteLine("Why?: " + bill);
                }
                else
                {
                    var bill_to_read = File.OpenRead(bill);
                    XSSFWorkbook currentWB = new XSSFWorkbook(bill_to_read);
                    ISheet current_sheet = currentWB.GetSheetAt(0);

                    int sheet_count = 25;
                    int blanks = 0;
                    while (sheet_count < 5000)
                    {
                        if (blanks == 3)
                        {
                            sheet_count += 1;
                            break;
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
                            IRow current_row = results_sheet.CreateRow(count);
                            if (jobpo_dict.ContainsKey(current_sheet.GetRow(sheet_count).GetCell(3).NumericCellValue.ToString()))
                            {
                                Console.WriteLine("PO: " + current_sheet.GetRow(sheet_count).GetCell(3).NumericCellValue.ToString() + " repeated! on bill: " + bill);
                            }
                            else
                            {
                                jobpo_dict.Add(current_sheet.GetRow(sheet_count).GetCell(3).NumericCellValue.ToString(), current_sheet.GetRow(sheet_count).GetCell(2).NumericCellValue.ToString()); //adds the po as the key and the job num as the value found in the bill to the local dict
                                current_row.CreateCell(0).SetCellValue(product); //producto (0)
                                current_row.CreateCell(2).SetCellValue(current_sheet.GetRow(sheet_count).GetCell(10).NumericCellValue.ToString()); // costo (2)
                                current_row.CreateCell(3).SetCellValue(current_sheet.GetRow(sheet_count).GetCell(9).NumericCellValue.ToString()); //valor agregado (3)
                                current_row.CreateCell(4).SetCellValue(current_sheet.GetRow(sheet_count).GetCell(8).NumericCellValue.ToString()); // peso (4)
                                current_row.CreateCell(5).SetCellValue(current_sheet.GetRow(sheet_count).GetCell(7).ToString()); //medida!!
                                current_row.CreateCell(6).SetCellValue(current_sheet.GetRow(sheet_count).GetCell(3).NumericCellValue.ToString()); //po_only!!
                                current_row.CreateCell(7).SetCellValue(bill); //po_only!!
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

        //structured BOM with Added Items 
        public static void WriteResultsToExcel()
        {
            ISheet results_sheet = hssfWorkbook.GetSheet("Structured_BOM");
            int count = 1;
            foreach(Record record in Records)
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
        }

        //material with order;
        public static void MCO(string mco)
        {
            XSSFWorkbook mcoWB;

            using (var fs = File.OpenRead(mco))
            {
                mcoWB = new XSSFWorkbook(fs);
            }

            ISheet mco_sheet = mcoWB.GetSheetAt(0);
            int starting_row_current_sheet = 3;
            var positive_items = new Dictionary<string, double>();

            // this while goes to the MCO WB and records in the positive items dictionary the items with a positive value in the qty
            while (mco_sheet.GetRow(starting_row_current_sheet) != null)
            {
                {
                    if (mco_sheet.GetRow(starting_row_current_sheet).GetCell(10).NumericCellValue > 0)
                    {
                        // the dictionary key is po-material#
                        string key = mco_sheet.GetRow(starting_row_current_sheet).GetCell(14).NumericCellValue.ToString() + '-' + mco_sheet.GetRow(starting_row_current_sheet).GetCell(1).StringCellValue;
                        if (positive_items.ContainsKey(key))
                        {
                            positive_items[key] += mco_sheet.GetRow(starting_row_current_sheet).GetCell(10).NumericCellValue;
                        }
                        else
                        {
                            positive_items.Add(key, mco_sheet.GetRow(starting_row_current_sheet).GetCell(10).NumericCellValue);
                        }
                    }
                    starting_row_current_sheet++;
                }
            }

            starting_row_current_sheet = 3; //resetting counter

            // the following while loops through all the records in the MCO excel and if the po is in the bill po's list it checks if the qty value is negative, 
            // if it is it then proceeds to look for the item in the positive_items dict,
            // if it exist it compares the value in the dict and the qty, if they are equal it deletes the record and goes to the next record in the file
            // if the qty and value in the dict are not equal it adds the current record to the list of objects
            // in case it is not in the positive_items dict it proceeds to add the current record to the list of objects
            while (mco_sheet.GetRow(starting_row_current_sheet) != null)
            {
                string po = mco_sheet.GetRow(starting_row_current_sheet).GetCell(14).NumericCellValue.ToString();

                double qty = mco_sheet.GetRow(starting_row_current_sheet).GetCell(10).NumericCellValue;

                string material = mco_sheet.GetRow(starting_row_current_sheet).GetCell(1).StringCellValue;

                // the dictionary key is po-material
                string key = po + '-' + material;

                if (jobpo_dict.ContainsKey(po))
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
                                        Product = jobpo_dict[po] + '-' + po,
                                        Material = material,
                                        Qty = qty*-1,
                                        UM = mco_sheet.GetRow(starting_row_current_sheet).GetCell(11).StringCellValue
                                    });
                            }
                        }
                        else
                        {
                            //write to excel
                            Records.Add(
                                    new Record
                                    {
                                        Product = jobpo_dict[po] + '-' + po,
                                        Material = material,
                                        Qty = qty*-1,
                                        UM = mco_sheet.GetRow(starting_row_current_sheet).GetCell(11).StringCellValue
                                    });
                        }
                    }
                }
                starting_row_current_sheet++;
            }
        }

        public static void LoopThroughItemsToAdd(string spreadsheet, int startingRow, int materialCol, int uMCol, int qtyCol)
        {
            XSSFWorkbook itemsToAddWB;

            using (var fs = File.OpenRead(spreadsheet))
            {
                itemsToAddWB = new XSSFWorkbook(fs);
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
                    Qty = record.Qty*-1
                });
            }
        }

        public static void ItemsToAdd()
        {
            int Record_index = 0;
            foreach (ToAddRecord itemToAdd in RecordsToAdd)
            {
                if ((itemToAdd.Qty)/jobpo_dict.Count() >= 1)
                {
                    int qty_per_order = Convert.ToInt32(itemToAdd.Qty / jobpo_dict.Count());
                    int remaining = Convert.ToInt32(itemToAdd.Qty % jobpo_dict.Count());
                    double decimals = itemToAdd.Qty % 1;
                    int i = 0;
                    foreach (KeyValuePair<string, string> entry in jobpo_dict)
                    {
                        if(i == 0)
                        {
                            Records_With_Added_PO.Add(
                                   new Record
                                   {
                                       Product = entry.Value + '-' + entry.Key,
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
                                    Product = entry.Value + '-' + entry.Key,
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
                                    Product = entry.Value + '-' + entry.Key,
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
                    int remaining = Convert.ToInt32(itemToAdd.Qty % jobpo_dict.Count());
                    double decimals = itemToAdd.Qty % 1;
                    int i = 0;
                    while (i < remaining)
                    {
                        if (i == 0)
                        {
                            Records_With_Added_PO.Add(
                                   new Record
                                   {
                                       Product = jobpo_dict.ElementAt(Record_index).Value + '-' + jobpo_dict.ElementAt(Record_index).Key,
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
                                    Product = jobpo_dict.ElementAt(Record_index).Value + '-' + jobpo_dict.ElementAt(Record_index).Key,
                                    Material = itemToAdd.Material,
                                    Qty = 1,
                                    UM = itemToAdd.UM
                                });
                        }
                        i++;
                        if (Record_index + 1 == jobpo_dict.Count)
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
    }
}
