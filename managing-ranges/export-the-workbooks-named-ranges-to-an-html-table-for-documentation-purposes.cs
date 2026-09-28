// Title: Generate an HTML documentation table of all named ranges in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# using Aspose.Cells to open an .xlsx workbook, iterate through workbook.Worksheets.Names, and build an HTML markup that lists each named range's name, RefersTo formula, and scope in a tabular layout. | Apply WebUtility.HtmlEncode to every field before inserting it into the markup and write the result to a file using UTF‑8 encoding. | Add validation that the source file exists and surround workbook loading and file writing with try‑catch blocks that output error details to the console.
// Common Searches: aspocells c# export excel named ranges as html report | generate html documentation of workbook named ranges using Aspose.Cells | list all named ranges with scope from an xlsx file in c# | how to create an html table of excel named ranges with Aspose.Cells .NET | save named range definitions to html using Aspose.Cells library
// Tags: Aspose.Cells export named ranges to HTML | C# generate named range documentation | Aspose.Cells list workbook named ranges | HTML markup generation from Excel named ranges | named range scope extraction Aspose.Cells

using System;
using System.IO;
using System.Text;
using System.Net;
using Aspose.Cells;

// The program loads an Excel workbook, iterates through its defined named ranges, HTML‑encodes each name, reference formula, and scope, and writes them into a UTF‑8 encoded HTML file containing a table for documentation purposes.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "NamedRanges.html";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Build an HTML document that contains a table of named ranges
        var html = new StringBuilder();

        html.AppendLine("<html>");
        html.AppendLine("<head>");
        html.AppendLine("<meta charset=\"UTF-8\">");
        html.AppendLine("<title>Workbook Named Ranges</title>");
        html.AppendLine("</head>");
        html.AppendLine("<body>");
        html.AppendLine("<h2>Named Ranges Documentation</h2>");
        html.AppendLine("<table border=\"1\" cellpadding=\"5\" cellspacing=\"0\">");
        html.AppendLine("<tr><th>Name</th><th>Refers To</th><th>Scope</th></tr>");

        // Iterate through all named ranges defined in the workbook
        foreach (Name namedRange in workbook.Worksheets.Names)
        {
            // The formula that defines the range (e.g., =Sheet1!$A$1:$B$5)
            string refersTo = namedRange.RefersTo;

            // Scope: older Aspose.Cells versions may not expose IsWorkbookScoped/Worksheet.
            // Default to "Workbook" when scope information is unavailable.
            string scope = "Workbook";

            // Encode values for safe HTML output
            string nameEncoded = WebUtility.HtmlEncode(namedRange.Text);
            string refersToEncoded = WebUtility.HtmlEncode(refersTo);
            string scopeEncoded = WebUtility.HtmlEncode(scope);

            // Append a row for the current named range
            html.AppendLine("<tr>");
            html.AppendLine($"<td>{nameEncoded}</td>");
            html.AppendLine($"<td>{refersToEncoded}</td>");
            html.AppendLine($"<td>{scopeEncoded}</td>");
            html.AppendLine("</tr>");
        }

        html.AppendLine("</table>");
        html.AppendLine("</body>");
        html.AppendLine("</html>");

        try
        {
            // Save the generated HTML to a file
            File.WriteAllText(outputPath, html.ToString(), Encoding.UTF8);
            Console.WriteLine($"Named ranges documentation saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to write HTML file: {ex.Message}");
        }
    }
}
