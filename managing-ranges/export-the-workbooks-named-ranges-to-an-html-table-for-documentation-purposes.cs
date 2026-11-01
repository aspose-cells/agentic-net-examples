// Title: Generate an HTML documentation table of all named ranges in an Excel workbook using Aspose.Cells for .NET
// Common Searches: C# code to list Excel named ranges and export them to an HTML file with Aspose.Cells | How to create an HTML report of named ranges from an .xlsx using Aspose.Cells | Aspose.Cells .NET example for exporting workbook names to an HTML table | Retrieve the first cell value of each named range and save as HTML using C# | Document Excel named ranges as an HTML table programmatically with Aspose.Cells
// Tags: Aspose.Cells named range HTML export | C# HTML report of Excel named ranges | extract first cell value from named range Aspose.Cells | iterate workbook names .NET Aspose.Cells | generate documentation for Excel named ranges

using System;
using System.IO;
using System.Net;
using System.Text;
using Aspose.Cells;

// The sample loads an Excel workbook, iterates through all defined names, captures each name, its reference address, and the value of the first cell in the range, then builds a UTF‑8 HTML document containing a table with these details and writes it to a file.
class ExportNamedRangesToHtml
{
    static void Main()
    {
        const string inputFile = "input.xlsx";
        const string outputFile = "NamedRanges.html";

        // Verify that the input workbook exists
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: Input file \"{inputFile}\" not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook from the file
            workbook = new Workbook(inputFile);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // StringBuilder to build the HTML content
        StringBuilder html = new StringBuilder();

        // Start of the HTML document
        html.AppendLine("<!DOCTYPE html>");
        html.AppendLine("<html>");
        html.AppendLine("<head><meta charset='UTF-8'><title>Named Ranges Documentation</title></head>");
        html.AppendLine("<body>");
        html.AppendLine("<h2>Workbook Named Ranges</h2>");
        html.AppendLine("<table border='1' cellpadding='5' cellspacing='0'>");
        html.AppendLine("<tr><th>Name</th><th>Refers To</th><th>First Cell Value</th></tr>");

        // Iterate through all named ranges in the workbook
        foreach (Name namedRange in workbook.Worksheets.Names)
        {
            // Address the name refers to (e.g., Sheet1!$A$1:$B$2)
            string refersTo = namedRange.RefersTo;

            // Retrieve the value of the first cell in the range, if possible
            string firstValue = string.Empty;
            try
            {
                Aspose.Cells.Range range = namedRange.GetRange();
                if (range != null && range.RowCount > 0 && range.ColumnCount > 0)
                {
                    object val = range[0, 0].Value;
                    firstValue = val?.ToString() ?? string.Empty;
                }
            }
            catch
            {
                // If the name does not refer to a valid range, leave the value empty
                firstValue = string.Empty;
            }

            // Append a row for the current named range
            html.AppendLine("<tr>");
            html.AppendLine($"<td>{WebUtility.HtmlEncode(namedRange.Text)}</td>");
            html.AppendLine($"<td>{WebUtility.HtmlEncode(refersTo)}</td>");
            html.AppendLine($"<td>{WebUtility.HtmlEncode(firstValue)}</td>");
            html.AppendLine("</tr>");
        }

        // End of the HTML document
        html.AppendLine("</table>");
        html.AppendLine("</body>");
        html.AppendLine("</html>");

        try
        {
            // Write the HTML content to a file
            File.WriteAllText(outputFile, html.ToString(), Encoding.UTF8);
            Console.WriteLine($"Named ranges exported successfully to \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error writing HTML file: {ex.Message}");
        }
    }
}
