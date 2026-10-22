// Title: Remove extra spaces that follow line‑break characters in Excel cells before saving as HTML with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook using Aspose.Cells, iterates every worksheet, and trims whitespace that appears after newline characters in string cells, then saves the workbook as HTML. | Create a reusable method that replaces a line‑break followed by one or more spaces with a single line‑break in all cell values of a workbook via Aspose.Cells. | Demonstrate how to apply a regular expression to clean up cell text whitespace across a workbook before exporting to HTML with Aspose.Cells.
// Common Searches: aspocells how to eliminate spaces after line feeds when exporting to html | c# delete trailing whitespace in excel cell strings before html conversion | regex to clean up newline spacing in an Aspose.Cells workbook | export excel to html without extra blank spaces using Aspose.Cells
// Tags: remove newline trailing spaces Aspose.Cells | regex cleanup cell text C# | process all worksheets string cells Aspose.Cells | html export without extra line break spaces Aspose.Cells

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The program loads an Excel workbook, scans each string cell, uses a regular expression to replace line‑breaks followed by spaces with a single line‑break, updates modified cells, and saves the workbook as an HTML file using Aspose.Cells.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Define a regex that matches line breaks followed by one or more spaces
        Regex lineBreakSpaceRegex = new Regex(@"(\r?\n)\s+", RegexOptions.Compiled);

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the used range to limit iteration to cells that contain data
            var cells = sheet.Cells;
            var maxRow = cells.MaxDataRow;
            var maxColumn = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxColumn; col++)
                {
                    Cell cell = cells[row, col];
                    // Process only string cells
                    if (cell.Type == CellValueType.IsString && cell.Value != null)
                    {
                        string original = cell.StringValue;
                        // Replace line break + spaces with just line break
                        string cleaned = lineBreakSpaceRegex.Replace(original, "$1");
                        // Update the cell only if changes were made
                        if (!original.Equals(cleaned))
                        {
                            cell.PutValue(cleaned);
                        }
                    }
                }
            }
        }

        // Set HTML save options if needed (e.g., to embed CSS)
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.ExportActiveWorksheetOnly = false; // export all worksheets

        // Save the workbook as HTML
        workbook.Save("output.html", saveOptions);
    }
}
