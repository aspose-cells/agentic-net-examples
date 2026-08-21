// Title: Export a Named Excel Table to an HTML Fragment with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, finds a ListObject by its name on the first worksheet, extracts the table's data range, and uses ExportTableOptions (ExportAsHtmlString = true) to return the table as a single HTML fragment stored in a DataTable cell.
// Keywords: Aspose.Cells C# export table HTML | Export ListObject to HTML string | Excel table to HTML fragment .NET | ExportTableOptions example | Convert Excel table to HTML snippet
// Common Searches: Aspose.Cells export specific table to HTML | C# get HTML from named ListObject | Export Excel table without headers as HTML | How to retrieve HTML fragment from Excel table using Aspose
// Developer Intent: Generate an HTML snippet from a named table (ListObject) inside an Excel workbook using Aspose.Cells.
// Use Cases: Create web‑ready table snippets for dashboards from user‑uploaded spreadsheets. | Insert Excel‑derived tables into email bodies or CMS content without full workbook conversion. | Dynamically render selected worksheet tables on a website while preserving styling.
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells, locates a ListObject called "MyTable", and returns its data as an HTML string without column headers. | Show how to handle the case where the specified table does not exist and log an appropriate message. | Explain how to modify the example to include column names in the HTML output and to export multiple named tables from the same workbook.

using System;
using System.Data;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Loads an Excel workbook, finds a ListObject by its name on the first worksheet, extracts the table's data range, and uses ExportTableOptions (ExportAsHtmlString = true) to return the table as a single HTML fragment stored in a DataTable cell.
class Program
{
    static void Main()
    {
        // Path to the workbook file
        string workbookPath = "input.xlsx";

        // Verify that the input file exists
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"File not found: {workbookPath}");
            return;
        }

        // Name of the table (ListObject) to export
        string tableName = "MyTable";

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Assume the table is in the first worksheet (adjust if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Locate the table (ListObject) by its name
            ListObject table = worksheet.ListObjects[tableName];
            if (table == null)
            {
                Console.WriteLine($"Table \"{tableName}\" not found.");
                return;
            }

            // Get the data range of the table (excluding header row)
            Aspose.Cells.Range dataRange = table.DataRange;

            // Set export options to get HTML string
            ExportTableOptions exportOptions = new ExportTableOptions
            {
                ExportAsHtmlString = true,
                ExportColumnName = false // optional: exclude column names from the HTML fragment
            };

            // Export the range to a DataTable where the cell contains HTML
            DataTable dt = dataRange.ExportDataTable(exportOptions);

            // The HTML fragment is stored in the first cell of the DataTable
            string htmlFragment = dt.Rows[0][0]?.ToString() ?? string.Empty;

            // Output the HTML fragment
            Console.WriteLine(htmlFragment);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
