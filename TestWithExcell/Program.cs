using OfficeOpenXml;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace TestWithExcell
{
    class Program
    {
        public static void Main(string[] args)
        {
            string filePath = @"c:/adb/AndreyFile.xlsx";
            var resultBag = new List<List<string>>();
            var file = new FileInfo(filePath);

            using (ExcelPackage excelPackage = new ExcelPackage(file))
            {
                //create an instance of the the first sheet in the loaded file
                ExcelWorksheet worksheet1 = excelPackage.Workbook.Worksheets[0];
                var firstTable = ExcelTool.CreateDataTable(worksheet1);

                var firstSheet = new List<object[]>();

                foreach (var item in firstTable.Rows)
                {
                    var currentRow = (DataRow)item;
                    var cells = currentRow.ItemArray;
                    firstSheet.Add(cells);
                }

                ExcelWorksheet worksheet2 = excelPackage.Workbook.Worksheets[1];
                var secondTable = ExcelTool.CreateDataTable(worksheet2);

                var secondSheet = new List<object[]>();

                foreach (var item in secondTable.Rows)
                {
                    var currentRow = (DataRow)item;
                    var cells = currentRow.ItemArray;
                    secondSheet.Add(cells);
                }


                foreach (var item1 in firstSheet)
                {
                    var platec1 = item1[0].ToString();
                    var obekt = item1[1].ToString();


                    foreach (var item2 in secondSheet)
                    {
                        var currentRow = new List<string>();
                        var platec2 = item2[0].ToString();

                        if (platec1.Equals(platec2))
                        {
                            var table = "Table";
                            var product = item2[1].ToString();
                            var days = item2[2].ToString();
                            currentRow.Add(table);
                            currentRow.Add(obekt);
                            currentRow.Add(product);
                            currentRow.Add(days);

                            resultBag.Add(currentRow);
                        }
                    }
                }
            }

            using (ExcelPackage excelPackage = new ExcelPackage(file))
            {
                ExcelWorksheet resultSheet = excelPackage.Workbook.Worksheets.Add("Result");
                excelPackage.Save();
            }

            using (ExcelPackage excelPackage = new ExcelPackage(file))
            {
                ExcelWorksheet resultSheet = excelPackage.Workbook.Worksheets[2];

                for (int i = 0; i < resultBag.Count; i++)
                {
                    var currentRow = resultBag[i];

                    for (int j = 0; j < currentRow.Count; j++)
                    {
                        resultSheet.Cells[i + 1, j + 1].Value = currentRow[j];
                    }
                }
                excelPackage.Save();
            }
        }
    }
}