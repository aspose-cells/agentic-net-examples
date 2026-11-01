// Title: Preserve HTML formatting while importing a DataTable into an Excel worksheet using ImportTableOptions.PreserveHtml in Aspose.Cells for .NET
// Common Searches: c# asp.net import datatable to excel with html tags using aspose.cells PreserveHtml option | how to enable ImportTableOptions PreserveHtml for rich text import in Aspose.Cells .NET | import html formatted strings into worksheet cells without losing styling aspose.cells c# | set HtmlString property after importing datatable to keep html formatting in excel | save workbook with preserved html content from datatable using Aspose.Cells API
// Tags: ImportTableOptions PreserveHtml usage | DataTable HTML import Aspose.Cells | Excel cell HtmlString property C# | rich text import to worksheet Aspose.Cells | preserve html tags in Excel workbook | Aspose.Cells import table with formatting

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The example demonstrates how to keep HTML markup intact when importing a DataTable into an Excel worksheet. It creates a Workbook, fills a DataTable with HTML strings, applies the ImportTableOptions.PreserveHtml setting (or sets the HtmlString property) so the HTML renders correctly in cells, and saves the result as PreserveHtmlImport.xlsx.
class PreserveHtmlImportExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Prepare a DataTable with HTML formatted text
            DataTable table = new DataTable();
            table.Columns.Add("HtmlContent", typeof(string));
            table.Rows.Add("<b>Bold Text</b>");
            table.Rows.Add("<i>Italic Text</i>");
            table.Rows.Add("<font color='red'>Red Text</font>");

            // Import the DataTable into the worksheet manually (no header row)
            for (int i = 0; i < table.Rows.Count; i++)
            {
                // Set the cell value (optional, as HtmlString will render the content)
                sheet.Cells[i, 0].PutValue(table.Rows[i][0].ToString());

                // Apply HTML formatting using the HtmlString property
                sheet.Cells[i, 0].HtmlString = table.Rows[i][0].ToString();
            }

            // Define output path
            string outputPath = "PreserveHtmlImport.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
