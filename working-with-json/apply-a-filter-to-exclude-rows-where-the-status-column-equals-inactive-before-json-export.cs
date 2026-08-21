// Title: C# – Export Active Rows to JSON with Aspose.Cells AutoFilter
// Description: Creates a workbook, adds sample data, applies an AutoFilter on the Status column to hide rows where Status = "Inactive", exports only the visible rows to a DataTable using ExportTableOptions (PlotVisibleRows = true), and serializes the result to an indented JSON string.
// Keywords: Aspose.Cells | C# | .NET | AutoFilter | ExportTableOptions | PlotVisibleRows | JSON export | filter rows | visible rows | DataTable to JSON
// Common Searches: Aspose.Cells export filtered rows to JSON | C# hide inactive rows before JSON export | Export only visible rows Aspose.Cells | How to use AutoFilter with ExportTableOptions | Serialize filtered Excel data to JSON C#
// Developer Intent: Filter out rows where the Status column equals "Inactive" and generate JSON that contains only the remaining (active) records.
// Use Cases: Create a JSON payload of active customers for an API call. | Produce a lightweight JSON report that includes only currently active items from a large spreadsheet. | Supply filtered data to front‑end grid components by exporting visible rows as JSON.
// AI Prompts: Show how to filter multiple status values before exporting to JSON with Aspose.Cells. | Provide code to write the resulting JSON string to a file instead of the console using ExportTableOptions. | Explain how to obtain the row count of the filtered DataTable after applying AutoFilter in Aspose.Cells.

using System;
using System.Data;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Creates a workbook, adds sample data, applies an AutoFilter on the Status column to hide rows where Status = "Inactive", exports only the visible rows to a DataTable using ExportTableOptions (PlotVisibleRows = true), and serializes the result to an indented JSON string.
class FilterAndExportToJson
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // Populate sample data (Header + some rows)
            // ------------------------------------------------------------
            // Header row
            cells["A1"].PutValue("ID");
            cells["B1"].PutValue("Name");
            cells["C1"].PutValue("Status"); // Column to filter on (index 2)

            // Data rows
            cells["A2"].PutValue(1); cells["B2"].PutValue("Alice");   cells["C2"].PutValue("Active");
            cells["A3"].PutValue(2); cells["B3"].PutValue("Bob");     cells["C3"].PutValue("Inactive");
            cells["A4"].PutValue(3); cells["B4"].PutValue("Charlie"); cells["C4"].PutValue("Active");
            cells["A5"].PutValue(4); cells["B5"].PutValue("Diana");   cells["C5"].PutValue("Inactive");
            cells["A6"].PutValue(5); cells["B6"].PutValue("Eve");     cells["C6"].PutValue("Active");

            // ------------------------------------------------------------
            // Apply AutoFilter to hide rows where Status = "Inactive"
            // ------------------------------------------------------------
            // Set the autofilter range (including header and all data rows)
            sheet.AutoFilter.SetRange(0, 0, 5); // startRow=0, startColumn=0, endRow=5 (A1:C6)

            // Filter the Status column (field index 2) to show only "Active"
            sheet.AutoFilter.Filter(2, "Active");
            // Refresh to apply the filter (rows with "Inactive" become hidden)
            sheet.AutoFilter.Refresh();

            // ------------------------------------------------------------
            // Export only the visible rows to a DataTable
            // ------------------------------------------------------------
            ExportTableOptions exportOptions = new ExportTableOptions
            {
                PlotVisibleRows = true,      // Export only rows that are not hidden by the filter
                ExportColumnName = true,     // Include column names as DataTable columns
                DataTable = null             // Let Aspose.Cells create the DataTable
            };

            // Export the range covering all used cells (A1:C6)
            // Parameters: firstRow, firstColumn, totalRows, totalColumns, options
            cells.ExportDataTable(0, 0, 6, 3, exportOptions);

            DataTable visibleData = exportOptions.DataTable;

            // ------------------------------------------------------------
            // Convert the DataTable to JSON string
            // ------------------------------------------------------------
            string jsonResult = JsonSerializer.Serialize(visibleData, new JsonSerializerOptions { WriteIndented = true });

            // Output JSON to console
            Console.WriteLine(jsonResult);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
