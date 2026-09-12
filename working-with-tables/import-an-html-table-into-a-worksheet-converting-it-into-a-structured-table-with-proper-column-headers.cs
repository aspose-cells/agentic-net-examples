// Title: How to import an HTML table into an Excel worksheet and convert it to a structured ListObject using Aspose.Cells for .NET
// AI Prompts: Generate C# code that reads an HTML string containing a table, loads it into an Aspose.Cells Workbook, creates a ListObject with the first row as headers, applies a built‑in table style, and saves the result as an .xlsx file. | Write a reusable C# method `ImportHtmlTable(string htmlPath, string sheetName)` that loads the HTML file into a workbook, adds a styled ListObject over the detected range, and returns the populated Workbook.
// Common Searches: aspocells c# load html string and create excel table | convert html table to Aspose.Cells ListObject with header row | apply built‑in table style to imported html data using Aspose.Cells | c# import html table via memory stream and save as xlsx
// Tags: load html string into Aspose.Cells workbook | add ListObject with header detection C# | apply TableStyleMedium2 to Excel ListObject | save workbook as xlsx from memory stream | structured Excel table from html data Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Loads an HTML string containing a table into a Workbook, determines the used range, creates a ListObject that treats the first row as column headers, applies a medium built‑in table style, and saves the worksheet as an XLSX file.
class HtmlTableImport
{
    static void Main()
    {
        try
        {
            // Sample HTML containing a table
            string html = @"
            <html>
                <body>
                    <table border='1'>
                        <tr><th>Product</th><th>Quantity</th><th>Price</th></tr>
                        <tr><td>Apple</td><td>10</td><td>0.5</td></tr>
                        <tr><td>Banana</td><td>20</td><td>0.3</td></tr>
                        <tr><td>Cherry</td><td>15</td><td>0.8</td></tr>
                    </table>
                </body>
            </html>";

            // Load the HTML into a workbook using LoadOptions for HTML format
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
            using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(html)))
            {
                Workbook workbook = new Workbook(ms, loadOptions);

                // Get the first worksheet where the HTML table was imported
                Worksheet sheet = workbook.Worksheets[0];

                // Determine the used range of the imported data (including headers)
                int firstRow = 0; // A1 is the top‑left cell
                int firstColumn = 0;
                int totalRows = sheet.Cells.MaxDataRow + 1;      // MaxDataRow is zero‑based
                int totalColumns = sheet.Cells.MaxDataColumn + 1;

                // Add a structured table (ListObject) over the imported range
                // The last parameter 'true' indicates that the range contains a header row
                int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn,
                                                       totalRows - 1, totalColumns - 1, true);
                ListObject table = sheet.ListObjects[tableIndex];

                // Optional: give the table a display name and apply a style
                table.DisplayName = "ProductsTable";
                table.TableStyleType = TableStyleType.TableStyleMedium2;

                // Save the workbook to an Excel file
                string outputPath = "ImportedTable.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
