using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class ProtectWorksheetDisablePivotEditingDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("Fruit");
                worksheet.Cells["A3"].PutValue("Vegetable");
                worksheet.Cells["B1"].PutValue("Quantity");
                worksheet.Cells["B2"].PutValue(120);
                worksheet.Cells["B3"].PutValue(80);

                // Add a pivot table based on the sample data
                int pivotIndex = worksheet.PivotTables.Add("A1:B3", "D5", "SamplePivot");
                PivotTable pivotTable = worksheet.PivotTables[pivotIndex];
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Quantity as data field

                // Access the worksheet protection settings
                Protection protection = worksheet.Protection;

                // Disallow manipulation of pivot tables on the protected sheet
                protection.AllowUsingPivotTable = false;

                // Protect the worksheet (all protection types, no password)
                worksheet.Protect(ProtectionType.All);

                // Refresh pivot tables – this operation is still allowed despite protection
                worksheet.RefreshPivotTables();

                // Save the workbook
                string outputPath = "ProtectWorksheetDisablePivotEditingDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            ProtectWorksheetDisablePivotEditingDemo.Run();
        }
    }
}