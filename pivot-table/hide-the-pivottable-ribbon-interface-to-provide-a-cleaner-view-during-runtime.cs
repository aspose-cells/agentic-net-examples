using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class HidePivotTableRibbonDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Amount");
                sheet.Cells["A2"].PutValue("Food");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["A3"].PutValue("Beverage");
                sheet.Cells["B3"].PutValue(80);
                sheet.Cells["A4"].PutValue("Food");
                sheet.Cells["B4"].PutValue(150);
                sheet.Cells["A5"].PutValue("Beverage");
                sheet.Cells["B5"].PutValue(70);

                // Add a new worksheet to host the pivot table
                int pivotSheetIndex = workbook.Worksheets.Add(SheetType.Worksheet);
                Worksheet pivotSheet = workbook.Worksheets[pivotSheetIndex];
                pivotSheet.Name = "PivotTable";

                // Create the pivot table
                int pivotIndex = pivotSheet.PivotTables.Add("A1:B5", "C3", "PivotTable1");
                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Amount

                // Hide the PivotTable field list (ribbon interface) for a cleaner runtime view
                workbook.Settings.HidePivotFieldList = true;

                // Save the workbook
                string outputPath = "HidePivotTableRibbonDemo.xlsx";
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
            HidePivotTableRibbonDemo.Run();
        }
    }
}