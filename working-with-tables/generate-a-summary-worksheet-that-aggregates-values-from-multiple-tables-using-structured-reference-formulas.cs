using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;   // Required for ListObject

namespace AsposeCellsSummaryExample
{
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Create a new workbook ----------
                Workbook workbook = new Workbook();

                // ---------- Populate first worksheet with Table1 ----------
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "SalesData";

                // Header row
                sheet1.Cells["A1"].PutValue("Region");
                sheet1.Cells["B1"].PutValue("Amount");

                // Sample data
                sheet1.Cells["A2"].PutValue("North");
                sheet1.Cells["B2"].PutValue(1200);
                sheet1.Cells["A3"].PutValue("South");
                sheet1.Cells["B3"].PutValue(850);
                sheet1.Cells["A4"].PutValue("East");
                sheet1.Cells["B4"].PutValue(970);
                sheet1.Cells["A5"].PutValue("West");
                sheet1.Cells["B5"].PutValue(660);

                // Convert range A1:B5 to a table (ListObject) named Table1
                int table1Index = sheet1.ListObjects.Add(0, 0, 4, 1, true);
                ListObject table1 = sheet1.ListObjects[table1Index];
                table1.DisplayName = "Table1";   // Set table name

                // ---------- Populate second worksheet with Table2 ----------
                Worksheet sheet2 = workbook.Worksheets.Add("ExpenseData");

                // Header row
                sheet2.Cells["A1"].PutValue("Category");
                sheet2.Cells["B1"].PutValue("Amount");

                // Sample data
                sheet2.Cells["A2"].PutValue("Travel");
                sheet2.Cells["B2"].PutValue(300);
                sheet2.Cells["A3"].PutValue("Supplies");
                sheet2.Cells["B3"].PutValue(150);
                sheet2.Cells["A4"].PutValue("Utilities");
                sheet2.Cells["B4"].PutValue(200);
                sheet2.Cells["A5"].PutValue("Misc");
                sheet2.Cells["B5"].PutValue(100);

                // Convert range A1:B5 to a table (ListObject) named Table2
                int table2Index = sheet2.ListObjects.Add(0, 0, 4, 1, true);
                ListObject table2 = sheet2.ListObjects[table2Index];
                table2.DisplayName = "Table2";   // Set table name

                // ---------- Create Summary worksheet ----------
                Worksheet summary = workbook.Worksheets.Add("Summary");

                // Labels
                summary.Cells["A1"].PutValue("Summary of Amounts");
                summary.Cells["A3"].PutValue("Total Sales (Table1)");
                summary.Cells["A4"].PutValue("Total Expenses (Table2)");
                summary.Cells["A5"].PutValue("Grand Total");

                // Structured reference formulas
                summary.Cells["B3"].Formula = "=SUM(Table1[Amount])";
                summary.Cells["B4"].Formula = "=SUM(Table2[Amount])";
                summary.Cells["B5"].Formula = "=B3+B4";

                // Calculate formulas so values are stored
                workbook.CalculateFormula();

                // ---------- Save the workbook ----------
                string outputPath = "SummaryWorkbook.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}