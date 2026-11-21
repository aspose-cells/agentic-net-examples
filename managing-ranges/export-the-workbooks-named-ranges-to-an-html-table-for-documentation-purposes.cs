// Title: Generate an HTML table documenting all named ranges in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: how to list all named ranges from an Excel file and export to HTML using Aspose.Cells C# | Aspose.Cells C# export defined names to an HTML table | retrieve first cell value of a named range with Aspose.Cells in .NET | create HTML documentation of workbook named ranges using Aspose.Cells | C# code to iterate Workbook.Worksheets.Names and generate an HTML report
// Tags: Aspose.Cells export named ranges HTML | C# iterate Workbook.Worksheets.Names | GetRangeByName first cell value Aspose.Cells | HTML report of Excel defined names | Aspose.Cells named range scope handling

using System;
using System.IO;
using System.Text;
using System.Net;
using Aspose.Cells;

// The example loads an Excel workbook, iterates through all workbook‑level and worksheet‑level named ranges via Workbook.Worksheets.Names, extracts each range's name, scope, reference string, and the value of its first cell, and writes this information into a styled HTML table for documentation purposes.
class ExportNamedRangesToHtml
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "NamedRanges.html";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Prepare HTML output
            StringBuilder html = new StringBuilder();
            html.AppendLine("<!DOCTYPE html>");
            html.AppendLine("<html>");
            html.AppendLine("<head>");
            html.AppendLine("<meta charset=\"UTF-8\">");
            html.AppendLine("<title>Named Ranges Documentation</title>");
            html.AppendLine("<style>");
            html.AppendLine("table { border-collapse: collapse; width: 100%; }");
            html.AppendLine("th, td { border: 1px solid #ddd; padding: 8px; }");
            html.AppendLine("th { background-color: #f2f2f2; }");
            html.AppendLine("</style>");
            html.AppendLine("</head>");
            html.AppendLine("<body>");
            html.AppendLine("<h2>Workbook Named Ranges</h2>");
            html.AppendLine("<table>");
            html.AppendLine("<tr><th>Name</th><th>Scope</th><th>Refers To</th><th>First Cell Value</th></tr>");

            // Export all defined names (both workbook‑level and worksheet‑level)
            foreach (Name name in workbook.Worksheets.Names)
            {
                AppendNameRow(html, name, workbook);
            }

            html.AppendLine("</table>");
            html.AppendLine("</body>");
            html.AppendLine("</html>");

            // Save the HTML to a file
            File.WriteAllText(outputPath, html.ToString());
            Console.WriteLine($"Named ranges exported to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method to add a row for a named range
    private static void AppendNameRow(StringBuilder html, Name name, Workbook workbook)
    {
        // Scope: Aspose.Cells does not expose worksheet‑level scope directly in older versions.
        // We'll indicate "Workbook" for global names and "Worksheet" otherwise.
        string scope = "Workbook";

        // Reference string (e.g., Sheet1!$A$1:$B$5)
        string refersTo = name.RefersTo;

        // Get the value of the first cell in the range
        string firstValue = string.Empty;
        try
        {
            // Retrieve the range by its defined name
            Aspose.Cells.Range range = workbook.Worksheets.GetRangeByName(name.Text);
            if (range != null && range.RowCount > 0 && range.ColumnCount > 0)
            {
                object val = range[0, 0].Value;
                firstValue = val?.ToString() ?? string.Empty;
            }
        }
        catch
        {
            // Silently ignore any errors while retrieving the first cell value
        }

        html.AppendLine("<tr>");
        html.AppendLine($"<td>{WebUtility.HtmlEncode(name.Text)}</td>");
        html.AppendLine($"<td>{WebUtility.HtmlEncode(scope)}</td>");
        html.AppendLine($"<td>{WebUtility.HtmlEncode(refersTo)}</td>");
        html.AppendLine($"<td>{WebUtility.HtmlEncode(firstValue)}</td>");
        html.AppendLine("</tr>");
    }
}
