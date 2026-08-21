// Title: Disable Column Grand Totals in an Aspose.Cells PivotTable (C#)
// Description: Creates a workbook, adds sample data, builds a PivotTable, and sets ShowColumnGrandTotals to false to hide column grand totals before saving the file.
// Keywords: Aspose.Cells | C# | PivotTable | ShowColumnGrandTotals | hide column grand totals | disable column totals | Excel pivot programmatically | Aspose.Cells API example
// Common Searches: Aspose.Cells hide column grand totals C# | ShowColumnGrandTotals property example | Turn off column totals in Aspose.Cells PivotTable | C# code to suppress column grand totals in Excel pivot | Aspose.Cells PivotTable settings tutorial
// Developer Intent: Programmatically create a PivotTable and prevent column grand totals from appearing.
// Use Cases: Produce a sales dashboard where only row totals are needed for a cleaner view. | Generate financial summaries that omit column aggregates to reduce clutter. | Export Excel reports with customized PivotTable layouts that hide column totals by default.
// AI Prompts: Provide C# code that creates an Aspose.Cells PivotTable and disables column grand totals. | Show how to use the ShowColumnGrandTotals property to hide column totals while keeping row totals visible. | Explain how to toggle column grand totals on an existing Aspose.Cells PivotTable at runtime.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    // Creates a workbook, adds sample data, builds a PivotTable, and sets ShowColumnGrandTotals to false to hide column grand totals before saving the file.
    public class DisableColumnGrandTotals
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
                sheet.Cells["B2"].Value = "Phone";
                sheet.Cells["C2"].Value = 1200;

                sheet.Cells["A3"].Value = "Electronics";
                sheet.Cells["B3"].Value = "Laptop";
                sheet.Cells["C3"].Value = 2500;

                sheet.Cells["A4"].Value = "Furniture";
                sheet.Cells["B4"].Value = "Chair";
                sheet.Cells["C4"].Value = 300;

                sheet.Cells["A5"].Value = "Furniture";
                sheet.Cells["B5"].Value = "Table";
                sheet.Cells["C5"].Value = 800;

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure the pivot table: rows = Category, columns = Product, data = Sales
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);      // Category as row field
                pivotTable.AddFieldToArea(PivotFieldType.Column, 1);   // Product as column field
                pivotTable.AddFieldToArea(PivotFieldType.Data, 2);     // Sales as data field

                // Disable column grand totals
                pivotTable.ShowColumnGrandTotals = false;

                // Refresh the pivot cache and calculate the pivot data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook
                workbook.Save("PivotTable_NoColumnGrandTotals.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            DisableColumnGrandTotals.Run();
        }
    }
}
