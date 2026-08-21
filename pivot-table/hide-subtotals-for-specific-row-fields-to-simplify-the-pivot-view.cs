// Title: Hide Row Field Subtotals in an Aspose.Cells Pivot Table (C#)
// Description: Creates a workbook, adds a pivot table on sample sales data, and disables automatic subtotals for the Category row field using the IsAutoSubtotals property (or SetSubtotals). The pivot is refreshed, calculated, and saved as an Excel file.
// Keywords: Aspose.Cells | C# pivot table | hide subtotals | PivotField IsAutoSubtotals | SetSubtotals | .NET Excel | pivot table subtotal settings | Excel workbook generation
// Common Searches: Aspose.Cells hide subtotal row field C# | disable automatic subtotals Aspose.Cells .NET | set PivotField IsAutoSubtotals false | remove sum subtotal Aspose.Cells pivot | C# code to hide category subtotals in Excel pivot
// Developer Intent: Programmatically suppress subtotals for a specific row field in an Aspose.Cells pivot table while keeping other row fields' subtotals visible.
// Use Cases: Produce a sales report where category‑level totals are hidden for a cleaner hierarchy. | Generate a financial pivot that displays only item‑level subtotals, omitting higher‑level group totals. | Allow an application to toggle subtotal visibility per field based on user preferences.
// AI Prompts: Write C# code using Aspose.Cells that creates a pivot table and disables automatic subtotals for the 'Region' row field. | Show how to hide only the Sum subtotal for a pivot field while keeping Average and Count subtotals visible in Aspose.Cells .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a pivot table on sample sales data, and disables automatic subtotals for the Category row field using the IsAutoSubtotals property (or SetSubtotals). The pivot is refreshed, calculated, and saved as an Excel file.
    public class HidePivotFieldSubtotalsDemo
    {
        public static void Run()
        {
            try
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
                int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add two row fields: Category and Product
                int categoryRowIndex = pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                int productRowIndex = pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");

                // Add the data field (Sales)
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Retrieve the PivotField objects for the added row fields
                PivotField categoryField = pivotTable.RowFields[categoryRowIndex];
                PivotField productField = pivotTable.RowFields[productRowIndex];

                // Hide subtotals for the "Category" row field
                // Option 1: Disable automatic subtotals completely
                categoryField.IsAutoSubtotals = false;

                // Option 2: Explicitly hide specific subtotal types (e.g., Sum, Average)
                // categoryField.SetSubtotals(PivotFieldSubtotalType.Sum, false);
                // categoryField.SetSubtotals(PivotFieldSubtotalType.Average, false);
                // Add more SetSubtotals calls if other subtotal types are needed

                // Keep subtotals for the "Product" field (default behavior)

                // Refresh and calculate the pivot table to apply changes
                pivotTable.RefreshData();      // Correct API call
                pivotTable.CalculateData();

                // Save the workbook
                string outputPath = "HidePivotFieldSubtotalsDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            HidePivotFieldSubtotalsDemo.Run();
        }
    }
}
