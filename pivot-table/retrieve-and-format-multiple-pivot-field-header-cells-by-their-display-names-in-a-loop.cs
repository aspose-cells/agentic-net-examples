// Title: C# – Loop PivotTable fields and format header cells by display name with Aspose.Cells
// Description: This example creates a workbook, builds a pivot table from sample data, refreshes the cache, and then iterates through Row, Column, and Data fields. For each field it retrieves the header cell using GetCellByDisplayName and applies a light‑yellow background with bold text before saving the file.
// Keywords: Aspose.Cells C# | PivotTable GetCellByDisplayName | format pivot header cells | loop pivot fields .NET | apply style to Excel pivot headers | Excel automation Aspose | pivot table formatting code | C# Excel pivot styling
// Common Searches: Aspose.Cells retrieve pivot header by display name | C# format pivot table header cells | loop through PivotTable.RowFields in .NET | apply background color to pivot headers Aspose | GetCellByDisplayName example C#
// Developer Intent: Automatically style all pivot table header cells (row, column, data) by looping through fields and using their display names.
// Use Cases: Consistently apply branding colors to pivot headers without hard‑coding cell addresses. | Dynamically highlight newly added fields after a pivot refresh. | Generate Excel reports where header styling adapts to any set of pivot fields.
// AI Prompts: Show C# code that loops through PivotTable.RowFields, ColumnFields, and DataFields and formats each header cell using GetCellByDisplayName in Aspose.Cells. | Provide an Aspose.Cells example that refreshes a pivot cache, retrieves header cells by display name, and applies a custom style.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // This example creates a workbook, builds a pivot table from sample data, refreshes the cache, and then iterates through Row, Column, and Data fields. For each field it retrieves the header cell using GetCellByDisplayName and applies a light‑yellow background with bold text before saving the file.
    public class RetrieveAndFormatPivotFieldHeaders
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // Populate sample data for the pivot table
                // -------------------------------------------------
                sheet.Cells["A1"].Value = "Category";
                sheet.Cells["B1"].Value = "Product";
                sheet.Cells["C1"].Value = "Sales";

                sheet.Cells["A2"].Value = "North";
                sheet.Cells["B2"].Value = "Apple";
                sheet.Cells["C2"].Value = 1200;

                sheet.Cells["A3"].Value = "North";
                sheet.Cells["B3"].Value = "Banana";
                sheet.Cells["C3"].Value = 800;

                sheet.Cells["A4"].Value = "South";
                sheet.Cells["B4"].Value = "Apple";
                sheet.Cells["C4"].Value = 1500;

                sheet.Cells["A5"].Value = "South";
                sheet.Cells["B5"].Value = "Banana";
                sheet.Cells["C5"].Value = 700;

                // -------------------------------------------------
                // Create a pivot table
                // -------------------------------------------------
                int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add fields to the pivot table
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);      // Category
                pivotTable.AddFieldToArea(PivotFieldType.Column, 1);   // Product
                pivotTable.AddFieldToArea(PivotFieldType.Data, 2);     // Sales

                // Refresh and calculate to build the pivot structure
                pivotTable.RefreshData();   // Correct API to refresh pivot cache
                pivotTable.CalculateData();

                // -------------------------------------------------
                // Helper method to format a cell
                // -------------------------------------------------
                void ApplyFormatting(Cell cell)
                {
                    if (cell == null) return;

                    Style style = cell.GetStyle();
                    style.ForegroundColor = Color.LightYellow;
                    style.Pattern = BackgroundType.Solid;
                    style.Font.IsBold = true;
                    cell.SetStyle(style);
                }

                // -------------------------------------------------
                // Process Row fields
                // -------------------------------------------------
                foreach (PivotField rowField in pivotTable.RowFields)
                {
                    string displayName = rowField.DisplayName;
                    Cell headerCell = pivotTable.GetCellByDisplayName(displayName);
                    ApplyFormatting(headerCell);
                }

                // -------------------------------------------------
                // Process Column fields
                // -------------------------------------------------
                foreach (PivotField colField in pivotTable.ColumnFields)
                {
                    string displayName = colField.DisplayName;
                    Cell headerCell = pivotTable.GetCellByDisplayName(displayName);
                    ApplyFormatting(headerCell);
                }

                // -------------------------------------------------
                // Process Data fields
                // -------------------------------------------------
                foreach (PivotField dataField in pivotTable.DataFields)
                {
                    string displayName = dataField.DisplayName;
                    Cell headerCell = pivotTable.GetCellByDisplayName(displayName);
                    ApplyFormatting(headerCell);
                }

                // -------------------------------------------------
                // Save the workbook
                // -------------------------------------------------
                workbook.Save("PivotFieldHeadersFormatted.xlsx");
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
            RetrieveAndFormatPivotFieldHeaders.Run();
        }
    }
}
