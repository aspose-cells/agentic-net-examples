// Title: Refresh Pivot Table After Enabling Excel 2003 Compatibility (Text Truncation) – Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add a 300‑character string, set the pivot table's IsExcel2003Compatible flag, call RefreshPivotTables to enforce the 255‑character limit, and save the result as an Excel file using Aspose.Cells for C#.
// Keywords: Aspose.Cells refresh pivot table | Excel 2003 compatibility pivot | IsExcel2003Compatible C# | pivot table text truncation .NET | RefreshPivotTables Aspose | legacy Excel 2003 reports | Aspose.Cells pivot example
// Common Searches: how to refresh a pivot table after setting IsExcel2003Compatible | apply 255 character limit to pivot data with Aspose.Cells | enable Excel 2003 compatibility for pivot tables programmatically | Aspose.Cells refresh all pivot tables workbook | C# truncate long text in pivot table Excel 2003
// Developer Intent: Refresh a pivot table after activating Excel 2003 compatibility so that long text values are truncated to 255 characters.
// Use Cases: Generate legacy Excel 2003 reports where pivot fields must obey the 255‑character restriction. | Automate batch processing of workbooks to apply compatibility mode and refresh all pivots before distribution. | Validate that long description fields are correctly truncated when exporting to older Excel formats.
// AI Prompts: Write C# code with Aspose.Cells that sets IsExcel2003Compatible on a pivot table and refreshes it. | Explain why workbook.Worksheets.RefreshPivotTables() is necessary after enabling Excel 2003 compatibility. | Show how to handle text longer than 255 characters in a pivot table when saving to Excel 2003 format using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add a 300‑character string, set the pivot table's IsExcel2003Compatible flag, call RefreshPivotTables to enforce the 255‑character limit, and save the result as an Excel file using Aspose.Cells for C#.
    public class RefreshPivotTableExcel2003Compatibility
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "Data";

                // Populate source data with a long text (>255 characters) to demonstrate truncation
                string longText = new string('X', 300); // 300 characters
                dataSheet.Cells["A1"].PutValue("Product");
                dataSheet.Cells["B1"].PutValue("Description");
                dataSheet.Cells["A2"].PutValue("Item1");
                dataSheet.Cells["B2"].PutValue(longText); // Will be truncated when Excel2003 compatibility is on

                // Add a second row with normal length text
                dataSheet.Cells["A3"].PutValue("Item2");
                dataSheet.Cells["B3"].PutValue("Short description");

                // Add a pivot table on a new worksheet
                Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");
                int pivotIndex = pivotSheet.PivotTables.Add("Data!A1:B3", "A5", "PivotTable1");
                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

                // Configure pivot fields (Product as row, Description as data)
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Product
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Description

                // Ensure Excel 2003 compatibility is enabled (default is true, set explicitly)
                pivotTable.IsExcel2003Compatible = true;

                // Refresh all pivot tables in the workbook to apply truncation
                workbook.Worksheets.RefreshPivotTables();

                // Save the workbook
                workbook.Save("PivotTable_Excel2003Compatibility.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while creating the pivot table: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                RefreshPivotTableExcel2003Compatibility.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
