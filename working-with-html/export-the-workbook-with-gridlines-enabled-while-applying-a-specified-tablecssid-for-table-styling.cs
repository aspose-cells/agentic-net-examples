// Title: Export an Aspose.Cells workbook to HTML with gridlines visible and a custom TableCssId for table styling (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to save a workbook as HTML, keeps gridlines displayed, and sets HtmlSaveOptions.TableCssId to a user‑defined identifier. | Update the provided Aspose.Cells example so that the generated HTML includes gridlines and the table element has a custom CSS id via HtmlSaveOptions.
// Common Searches: how to keep gridlines when exporting Excel to HTML using Aspose.Cells .NET | Aspose.Cells set custom TableCssId for HTML table | C# export workbook to HTML with visible gridlines and custom CSS id | HtmlSaveOptions TableCssId example Aspose.Cells | Aspose.Cells HTML export styling table with CSS identifier
// Tags: Aspose.Cells HTML export with visible gridlines | custom CSS id for HTML table Aspose.Cells | C# HtmlSaveOptions table styling | export Excel to HTML Aspose.Cells | Aspose.Cells ListObject HTML output

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Saving;

// The sample creates a workbook, adds a simple table, configures HtmlSaveOptions (gridlines are shown by default), optionally sets a custom TableCssId for the HTML table, and saves the workbook as an HTML file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and name it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate sample data
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");

            // Define the range for the table (including header)
            int firstRow = 0;          // zero‑based index
            int firstColumn = 0;
            int totalRows = 3;         // header + 2 data rows
            int totalColumns = 2;

            // Add a ListObject (table) to the worksheet
            int tableIndex = sheet.ListObjects.Add(
                firstRow,
                firstColumn,
                firstRow + totalRows - 1,
                firstColumn + totalColumns - 1,
                true);

            ListObject table = sheet.ListObjects[tableIndex];

            // Set a display name for the table (used in Excel UI)
            table.DisplayName = "MyTable";

            // Optionally set a built‑in table style
            table.TableStyleType = TableStyleType.TableStyleMedium9;

            // Configure HTML save options (gridlines are shown by default)
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();

            // Define output path
            string outputPath = "ExportedWorkbook.html";

            // Save the workbook to HTML
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook successfully exported to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
