using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class PivotTableShowRowHeadersDemo
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Pivot table created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Amount");
                sheet.Cells["A2"].PutValue("Fruit");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["A3"].PutValue("Vegetable");
                sheet.Cells["B3"].PutValue(80);
                sheet.Cells["A4"].PutValue("Fruit");
                sheet.Cells["B4"].PutValue(150);
                sheet.Cells["A5"].PutValue("Vegetable");
                sheet.Cells["B5"].PutValue(70);

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D2", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure the pivot table: Category as row field, Amount as data field
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

                // The ShowHeaders property does not exist; row headers are displayed by default.
                // If needed, you can customize the row header caption:
                // pivotTable.RowHeaderCaption = "Category";

                // Save the workbook to a file
                string outputPath = "PivotTableShowRowHeadersDemo.xlsx";
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                // Log any runtime errors that occur during pivot table creation
                Console.WriteLine($"Run error: {ex.Message}");
                throw; // Re‑throw to be caught by the outer handler if desired
            }
        }
    }
}