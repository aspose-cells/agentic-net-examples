// Title: Show Hidden PivotTable Column Field in Aspose.Cells (C#) by Toggling ShowAllItems
// Description: Creates a workbook, builds a PivotTable with Product rows, Region columns, and Sales values, hides the Region column using ShowAllItems = false, then reveals it by setting ShowAllItems = true and refreshing the PivotTable before saving the file.
// Keywords: Aspose.Cells PivotTable hide column C# | show hidden pivot column Aspose.Cells | ShowAllItems property C# | toggle pivot column visibility .NET | refresh pivot after visibility change | Aspose.Cells Excel automation | C# Excel PivotTable example
// Common Searches: Aspose.Cells unhide pivot column C# | ShowAllItems example for PivotTable column | C# toggle visibility of PivotTable field | refresh Aspose.Cells pivot after ShowAllItems change | how to display hidden pivot column Aspose.Cells
// Developer Intent: Make a previously hidden PivotTable column field visible by resetting its ShowAllItems flag and refreshing the table.
// Use Cases: Generate a summary report with hidden region columns, then expand to detailed view. | Implement a UI button that shows or hides specific PivotTable columns on demand. | Automate Excel exports that initially conceal certain pivot columns and later reveal them after calculations.
// AI Prompts: Write C# code using Aspose.Cells to hide a PivotTable column field and later show it by setting ShowAllItems to true. | Provide an Aspose.Cells example that toggles a PivotTable column's visibility based on a boolean variable. | Explain the steps to refresh and recalculate a PivotTable after changing the ShowAllItems property of a column field in C#.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    // Creates a workbook, builds a PivotTable with Product rows, Region columns, and Sales values, hides the Region column using ShowAllItems = false, then reveals it by setting ShowAllItems = true and refreshing the PivotTable before saving the file.
    public class ShowHiddenColumnField
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].Value = "Product";
                sheet.Cells["B1"].Value = "Region";
                sheet.Cells["C1"].Value = "Sales";

                sheet.Cells["A2"].Value = "Bike";
                sheet.Cells["B2"].Value = "North";
                sheet.Cells["C2"].Value = 1200;

                sheet.Cells["A3"].Value = "Bike";
                sheet.Cells["B3"].Value = "South";
                sheet.Cells["C3"].Value = 1500;

                sheet.Cells["A4"].Value = "Car";
                sheet.Cells["B4"].Value = "North";
                sheet.Cells["C4"].Value = 2000;

                sheet.Cells["A5"].Value = "Car";
                sheet.Cells["B5"].Value = "South";
                sheet.Cells["C5"].Value = 2500;

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Add fields: Product as row, Region as column, Sales as data
                pivot.AddFieldToArea(PivotFieldType.Row, "Product");
                pivot.AddFieldToArea(PivotFieldType.Column, "Region");
                pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Initially hide the column field (Region) by setting ShowAllItems to false
                PivotField columnField = pivot.ColumnFields[0];
                columnField.ShowAllItems = false; // Hide the column field

                // Refresh and calculate to apply the hide operation
                pivot.RefreshData();
                pivot.CalculateData();

                // Now show the previously hidden column field by toggling ShowAllItems back to true
                columnField.ShowAllItems = true; // Show the column field again

                // Refresh and calculate to reflect the change
                pivot.RefreshData();
                pivot.CalculateData();

                // Determine output path and ensure directory exists
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ShowHiddenColumnFieldDemo.xlsx");
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during pivot table processing: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ShowHiddenColumnField.Run();
        }
    }
}
