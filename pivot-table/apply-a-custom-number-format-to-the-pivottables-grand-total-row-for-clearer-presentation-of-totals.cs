using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class PivotGrandTotalNumberFormatDemo
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Region");
            sheet.Cells["C1"].PutValue("Sales");

            sheet.Cells["A2"].PutValue("Laptop");
            sheet.Cells["B2"].PutValue("North");
            sheet.Cells["C2"].PutValue(1200);

            sheet.Cells["A3"].PutValue("Laptop");
            sheet.Cells["B3"].PutValue("South");
            sheet.Cells["C3"].PutValue(1500);

            sheet.Cells["A4"].PutValue("Phone");
            sheet.Cells["B4"].PutValue("North");
            sheet.Cells["C4"].PutValue(800);

            sheet.Cells["A5"].PutValue("Phone");
            sheet.Cells["B5"].PutValue("South");
            sheet.Cells["C5"].PutValue(1100);

            // Add a pivot table
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Configure pivot fields
            pivot.AddFieldToArea(PivotFieldType.Row, "Product");
            int dataFieldPos = pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
            PivotField dataField = pivot.DataFields[dataFieldPos];

            // Set number format for the data field (applies to grand total as well)
            dataField.NumberFormat = "$#,##0.00";

            // Refresh and calculate the pivot table
            pivot.RefreshData();
            pivot.CalculateData();

            // Define output file path
            string outputPath = "PivotGrandTotalNumberFormatDemo.xlsx";

            // Ensure we can write the file (overwrite if exists)
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }

            // Save the workbook
            workbook.Save(outputPath);
        }
    }
}