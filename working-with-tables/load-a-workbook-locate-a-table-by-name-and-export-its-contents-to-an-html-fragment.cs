// Title: Export a Named Excel Table to an HTML Fragment with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, finds a ListObject by its name on the first worksheet, determines the table's data range, and uses HtmlSaveOptions (ExportDataOptions.Table) with a custom ExportArea to write only that table to a memory stream. The stream is returned as a UTF‑8 HTML fragment, with error handling for missing files or tables.
// Keywords: Aspose.Cells | C# export Excel table to HTML | HtmlSaveOptions ExportDataOptions.Table | ListObject to HTML string | export specific range as HTML | Excel table HTML fragment | .NET HTML export | memory stream HTML | UTF-8 HTML fragment | named table export
// Common Searches: Aspose.Cells export named table to HTML | C# convert Excel ListObject to HTML string | save only a table from Excel as HTML using Aspose | export Excel table range as HTML fragment .NET | HTML fragment from Excel table Aspose.Cells example
// Developer Intent: Generate an HTML snippet that contains only the data of a specific named table in an Excel workbook.
// Use Cases: Embed a table‑only HTML snippet in a web report without creating a full HTML page. | Create email body content by converting a workbook's table to a UTF‑8 HTML fragment. | Show a live preview of a user‑uploaded Excel table in a web application.
// AI Prompts: Write C# code that loads a workbook, locates a ListObject named 'MyTable', and returns its contents as an HTML string using Aspose.Cells. | Explain how to configure HtmlSaveOptions to keep the original table styles while exporting only the selected table. | Show how to handle a missing workbook file or table name and return a clear, user‑friendly error message.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Loads an Excel workbook, finds a ListObject by its name on the first worksheet, determines the table's data range, and uses HtmlSaveOptions (ExportDataOptions.Table) with a custom ExportArea to write only that table to a memory stream. The stream is returned as a UTF‑8 HTML fragment, with error handling for missing files or tables.
class ExportTableToHtmlFragment
{
    static void Main()
    {
        // Path to the source workbook and the name of the table to export
        string workbookPath = "input.xlsx";
        string targetTableName = "MyTable";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: The file \"{workbookPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from file
            Workbook workbook = new Workbook(workbookPath);

            // Locate the worksheet that contains the table (assumes first worksheet)
            Worksheet worksheet = workbook.Worksheets[0];

            // Find the table (ListObject) by its name
            ListObject table = worksheet.ListObjects[targetTableName];
            if (table == null)
            {
                Console.WriteLine($"Error: Table \"{targetTableName}\" not found in the worksheet.");
                return;
            }

            // Define the cell area that corresponds to the table's data range
            Aspose.Cells.Range dataRange = table.DataRange;
            int startRow = dataRange.FirstRow;
            int startColumn = dataRange.FirstColumn;
            int endRow = startRow + dataRange.RowCount - 1;
            int endColumn = startColumn + dataRange.ColumnCount - 1;

            CellArea exportArea = CellArea.CreateCellArea(startRow, startColumn, endRow, endColumn);

            // Configure HTML save options to export only the table part
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportDataOptions = HtmlExportDataOptions.Table, // export only tables
                ExportArea = exportArea,                         // limit to the specific table range
                ExportActiveWorksheetOnly = true                // export only the active sheet
            };

            // Save the workbook to a memory stream using the HTML options
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Save(ms, htmlOptions);

                // Convert the stream content to a UTF‑8 string (HTML fragment)
                string htmlFragment = Encoding.UTF8.GetString(ms.ToArray());

                // Output the HTML fragment
                Console.WriteLine(htmlFragment);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
