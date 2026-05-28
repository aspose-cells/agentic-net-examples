using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class PivotItemMoveTwoPositionsDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Alpha");
            sheet.Cells["A3"].PutValue("Beta");
            sheet.Cells["A4"].PutValue("Gamma");
            sheet.Cells["A5"].PutValue("Delta");

            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["B3"].PutValue(200);
            sheet.Cells["B4"].PutValue(300);
            sheet.Cells["B5"].PutValue(400);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add the "Category" field to the row area
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");

            // Refresh and calculate to populate the pivot items
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Get the collection of pivot items for the row field
            PivotItemCollection items = pivotTable.RowFields[0].PivotItems;

            // Ensure there are enough items to move
            if (items.Count > 2)
            {
                // Move the first item two positions down within the same parent node
                items[0].Move(2, true);
            }

            // Save the workbook
            string outputPath = "PivotItemMoveTwoPositionsDemo.xlsx";
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
    }
}