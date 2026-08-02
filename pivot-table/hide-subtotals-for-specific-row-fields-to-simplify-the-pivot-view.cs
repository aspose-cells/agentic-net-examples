// Title: Hide Row Field Subtotals in an Aspose.Cells Pivot Table (C#)
// Description: Demonstrates how to create a workbook, add sample sales data, build a pivot table, place the "Category" field in the row area, disable its automatic subtotals, turn off Sum, Average and Count subtotals, then refresh, calculate and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# pivot table hide subtotals | disable row field subtotals Aspose.Cells | IsAutoSubtotals false | SetSubtotals PivotFieldSubtotalType | Aspose.Cells pivot table subtotal settings | C# Excel pivot table without subtotals | Aspose.Cells for .NET
// Common Searches: Aspose.Cells hide subtotals for row field | C# disable automatic subtotals in pivot table | SetSubtotals false Aspose.Cells example | Remove sum average count subtotals Aspose.Cells | How to turn off pivot table subtotals in C#
// Developer Intent: Remove subtotal rows for a specific row field in an Aspose.Cells pivot table.
// Use Cases: Generate a sales report where categories appear without any subtotal rows for a cleaner view. | Create financial dashboards that omit automatic subtotal calculations to reduce visual clutter. | Export Excel files with pivot tables that show only distinct row items, improving readability for end users.
// AI Prompts: Write C# code with Aspose.Cells to add a pivot table and hide all subtotals for a given row field. | Show how to set IsAutoSubtotals to false and disable Sum, Average, and Count subtotals for a pivot row field using Aspose.Cells. | Explain the steps to refresh, calculate, and save a workbook after modifying pivot field subtotal settings in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add sample sales data, build a pivot table, place the "Category" field in the row area, disable its automatic subtotals, turn off Sum, Average and Count subtotals, then refresh, calculate and save the file using Aspose.Cells for .NET.
    public class HidePivotFieldSubtotalsDemo
    {
        public static void Main(string[] args)
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
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].Value = "Category";
            sheet.Cells["B1"].Value = "Product";
            sheet.Cells["C1"].Value = "Sales";

            sheet.Cells["A2"].Value = "Electronics";
            sheet.Cells["B2"].Value = "TV";
            sheet.Cells["C2"].Value = 1000;

            sheet.Cells["A3"].Value = "Electronics";
            sheet.Cells["B3"].Value = "Radio";
            sheet.Cells["C3"].Value = 500;

            sheet.Cells["A4"].Value = "Clothing";
            sheet.Cells["B4"].Value = "Shirt";
            sheet.Cells["C4"].Value = 300;

            sheet.Cells["A5"].Value = "Clothing";
            sheet.Cells["B5"].Value = "Pants";
            sheet.Cells["C5"].Value = 400;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("=A1:C5", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add the "Category" field to the row area
            int rowFieldPos = pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            PivotField categoryField = pivotTable.RowFields[rowFieldPos];

            // Disable automatic subtotals for the "Category" row field
            categoryField.IsAutoSubtotals = false;

            // Optionally hide specific subtotal types (e.g., Sum, Average, Count)
            categoryField.SetSubtotals(PivotFieldSubtotalType.Sum, false);
            categoryField.SetSubtotals(PivotFieldSubtotalType.Average, false);
            categoryField.SetSubtotals(PivotFieldSubtotalType.Count, false);

            // Add the data field to the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Refresh and calculate the pivot table to apply changes
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Ensure output directory exists
            string outputPath = "HidePivotFieldSubtotalsDemo.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
        }
    }
}
