// Title: Aspose.Cells .NET – Keep Column Widths Fixed When Refreshing Pivot Tables
// Description: Demonstrates how to set explicit column widths, add a pivot table, disable the AutofitColumnWidthOnUpdate property, refresh the pivot, and verify that column sizes stay unchanged before saving the workbook.
// Keywords: Aspose.Cells pivot column width | disable AutofitColumnWidthOnUpdate | fixed column width after pivot refresh | C# Aspose.Cells pivot formatting | prevent column auto‑resize Aspose
// Common Searches: Aspose.Cells keep column width after pivot refresh | disable column autofit for pivot tables .NET | Aspose.Cells pivot table column width stays same | C# prevent pivot table from changing column sizes
// Developer Intent: The developer needs to preserve predefined column widths when a pivot table is refreshed in an Aspose.Cells workbook.
// Use Cases: Create a report layout with precise column alignment and ensure it is not altered by pivot updates. | Refresh multiple pivots in a financial dashboard while maintaining the surrounding column formatting. | Automate Excel generation where column widths are part of a corporate style guide and must remain constant.
// AI Prompts: How do I stop Aspose.Cells from auto‑adjusting column widths when a pivot table is refreshed in C#? | Provide C# code that sets column widths, adds a pivot table, disables AutofitColumnWidthOnUpdate, refreshes the pivot, and saves the workbook. | Explain alternative techniques to preserve column widths after updating pivot tables with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsFixedColumnWidthDemo
{
    // Demonstrates how to set explicit column widths, add a pivot table, disable the AutofitColumnWidthOnUpdate property, refresh the pivot, and verify that column sizes stay unchanged before saving the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate data for the pivot table
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Amount");
                sheet.Cells["A2"].PutValue("Food");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["A3"].PutValue("Transport");
                sheet.Cells["B3"].PutValue(80);
                sheet.Cells["A4"].PutValue("Food");
                sheet.Cells["B4"].PutValue(150);
                sheet.Cells["A5"].PutValue("Utilities");
                sheet.Cells["B5"].PutValue(200);

                // Set fixed column widths (character units)
                sheet.Cells.SetColumnWidth(0, 20); // Column A
                sheet.Cells.SetColumnWidth(1, 15); // Column B
                sheet.Cells.SetColumnWidth(2, 25); // Column C (pivot location)

                // Define the source data range for the pivot table (A1:B5)
                string sourceDataRange = $"={sheet.Name}!A1:B5";

                // Add the pivot table; Add returns the index of the new pivot table
                int pivotIndex = sheet.PivotTables.Add(sourceDataRange, "C1", "PivotTable1");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Configure pivot fields
                pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
                pivot.AddFieldToArea(PivotFieldType.Data, 1); // Amount as data field

                // Disable automatic column width autofit on pivot refresh
                pivot.AutofitColumnWidthOnUpdate = false;

                // Refresh pivot tables (column widths stay fixed)
                sheet.RefreshPivotTables();

                // Output column widths to verify they remain unchanged
                Console.WriteLine("Column widths after pivot refresh (should remain fixed):");
                Console.WriteLine($"Column A width: {sheet.Cells.GetColumnWidth(0)}");
                Console.WriteLine($"Column B width: {sheet.Cells.GetColumnWidth(1)}");
                Console.WriteLine($"Column C width (pivot start): {sheet.Cells.GetColumnWidth(2)}");

                // Save the workbook
                string outputPath = "FixedColumnWidthDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
