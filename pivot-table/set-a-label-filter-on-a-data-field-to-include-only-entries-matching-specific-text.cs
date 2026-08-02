// Title: C# – Apply a PivotField Label Filter (CaptionEqual) in Aspose.Cells to Show Only "Fruit" Rows
// Description: Demonstrates how to create a workbook, add a pivot table, place the "Category" field in the row area, and use PivotField.FilterByLabel with PivotFilterType.CaptionEqual to keep only rows where the category equals "Fruit". The example refreshes the pivot, calculates data, and saves the file.
// Keywords: Aspose.Cells C# pivot table filter | PivotField.FilterByLabel example | CaptionEqual label filter Aspose.Cells | filter pivot rows by text C# | Aspose.Cells .NET label filter tutorial | GitHub Aspose.Cells pivot filter sample | Excel pivot table label filter code
// Common Searches: How to filter pivot table rows by exact text using Aspose.Cells C# | Aspose.Cells FilterByLabel CaptionEqual "Fruit" example | Set label filter on pivot row field in .NET | C# code to show only specific categories in a pivot table | Aspose.Cells pivot table label filter GitHub
// Developer Intent: Filter a pivot table so that only rows with the category "Fruit" are displayed.
// Use Cases: Create a sales report that lists amounts exclusively for the selected product category. | Build an interactive workbook where users can view pivot data limited to predefined labels such as "Fruit" or "Vegetable". | Export a filtered pivot view to Excel for downstream analysis, showing only rows that match a given caption.
// AI Prompts: Generate C# code with Aspose.Cells to add a label filter to a pivot table row field for the value "Fruit". | Explain the behavior of PivotField.FilterByLabel when using PivotFilterType.CaptionEqual in Aspose.Cells. | Show how to apply multiple label filters to different pivot fields in a single Aspose.Cells workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add a pivot table, place the "Category" field in the row area, and use PivotField.FilterByLabel with PivotFilterType.CaptionEqual to keep only rows where the category equals "Fruit". The example refreshes the pivot, calculates data, and saves the file.
    public class SetLabelFilterDemo
    {
        public static void Main(string[] args)
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
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Fruit");
            sheet.Cells["A3"].PutValue("Vegetable");
            sheet.Cells["A4"].PutValue("Fruit");
            sheet.Cells["A5"].PutValue("Grain");

            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(80);
            sheet.Cells["B4"].PutValue(150);
            sheet.Cells["B5"].PutValue(60);

            // Add a pivot table covering the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D1", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Add the "Category" field to the row area
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");

            // Add the "Amount" field to the data area
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Retrieve the row field (first row field) to apply a label filter
            PivotField rowField = pivot.RowFields[0];

            // Apply a label filter to include only rows where the category equals "Fruit"
            // PivotFilterType.CaptionEqual filters by exact caption match
            rowField.FilterByLabel(PivotFilterType.CaptionEqual, "Fruit", null);

            // Refresh and calculate the pivot table to reflect the filter
            pivot.RefreshData();
            pivot.CalculateData();

            // Save the workbook to the current directory
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "SetLabelFilterDemo.xlsx");
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
    }
}
