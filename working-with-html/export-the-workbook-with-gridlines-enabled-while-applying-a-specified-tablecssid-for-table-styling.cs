// Title: Export an Aspose.Cells workbook to HTML with visible gridlines and a custom TableCssId for table styling
// AI Prompts: Generate C# code that creates a workbook, adds a ListObject, assigns a TableCssId value for custom CSS, enables ExportGridLines in HtmlSaveOptions, and saves the result as an HTML file. | Demonstrate how to apply a custom CSS class to an Aspose.Cells ListObject and export the worksheet to HTML while preserving Excel gridlines.
// Common Searches: aspnet export excel to html with gridlines using aspose.cells | how to assign TableCssId to a ListObject in Aspose.Cells C# | enable gridlines in HTML output from Aspose.Cells workbook | custom CSS class for Aspose.Cells HTML table styling | Aspose.Cells HtmlSaveOptions ExportGridLines true example
// Tags: Aspose.Cells HTML export with gridlines | Set TableCssId on ListObject C# | HtmlSaveOptions ExportGridLines property | Aspose.Cells ListObject CSS styling | C# workbook to HTML using Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;

// The example creates a new workbook, fills cells A1:B3 with sample data, adds a ListObject covering that range, sets a custom TableCssId for CSS styling, configures HtmlSaveOptions to export gridlines, and saves the workbook as ExportedWithGridlines.html.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Fill some sample data
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");

            // Define the range for the table (A1:B3)
            int firstRow = 0;      // zero‑based index for row 1
            int firstColumn = 0;   // zero‑based index for column A
            int totalRows = 3;
            int totalColumns = 2;

            // Add a ListObject (table) to the worksheet; the last parameter indicates that the first row contains headers
            int tableIndex = sheet.ListObjects.Add(
                firstRow,
                firstColumn,
                firstRow + totalRows - 1,
                firstColumn + totalColumns - 1,
                true);

            ListObject table = sheet.ListObjects[tableIndex];
            // Optionally set a display name for the table
            table.DisplayName = "MyTable";

            // Configure HTML save options to show gridlines
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportGridLines = true
            };

            // Export the workbook to HTML with the specified options
            workbook.Save("ExportedWithGridlines.html", saveOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
