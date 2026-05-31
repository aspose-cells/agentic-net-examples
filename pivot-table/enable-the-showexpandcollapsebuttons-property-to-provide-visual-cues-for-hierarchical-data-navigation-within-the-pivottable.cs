using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class PivotTableShowExpandCollapseDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet and name it
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Data";

                // Populate sample data for the pivot table
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(100);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(200);
                sheet.Cells["A4"].PutValue("A");
                sheet.Cells["B4"].PutValue(150);
                sheet.Cells["A5"].PutValue("B");
                sheet.Cells["B5"].PutValue(250);

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure the pivot table fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Value");

                // Enable drill‑down and visual expand/collapse buttons
                pivotTable.EnableDrilldown = true;
                pivotTable.ShowDrill = true; // Shows the expand/collapse field buttons

                // Refresh and calculate the pivot table data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Define output file path
                string outputPath = "PivotTableShowExpandCollapseDemo.xlsx";

                // If the file already exists, delete it to avoid IOException
                if (File.Exists(outputPath))
                {
                    File.Delete(outputPath);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log any runtime errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point required for console application
    public class Program
    {
        public static void Main(string[] args)
        {
            PivotTableShowExpandCollapseDemo.Run();
        }
    }
}