// Title: Convert an Excel worksheet to HTML and render the first row as <h1> headings using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export worksheet to HTML with column titles as h1 elements | How to map Excel first row to h1 tags when generating HTML with Aspose.Cells | Aspose.Cells generate HTML file from XLSX and include headings as top-level headings | Create HTML output from Excel in .NET preserving header row as <h1> | Encode worksheet header values and write them as h1 tags using Aspose.Cells C#
// Tags: Aspose.Cells HTML export with header h1 | C# convert XLSX to HTML using Aspose.Cells | write worksheet headings as top-level HTML headings | HTML table generation from Excel data Aspose.Cells | HTML encode worksheet cell values C#

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The example loads an XLSX workbook with Aspose.Cells, extracts the first row values and writes each non‑empty cell as an <h1> element, then builds an HTML table where the same first row becomes <th> headers and the remaining rows are rendered as <td> cells, finally saving the UTF‑8 HTML file.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        string inputPath = "Input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Determine the used range
        int maxRow = cells.MaxDataRow;
        int maxCol = cells.MaxDataColumn;

        // Build HTML output
        StringBuilder html = new StringBuilder();
        html.AppendLine("<!DOCTYPE html>");
        html.AppendLine("<html>");
        html.AppendLine("<head>");
        html.AppendLine("<meta charset=\"UTF-8\"/>");
        html.AppendLine("<title>Worksheet Export</title>");
        html.AppendLine("</head>");
        html.AppendLine("<body>");

        // Write worksheet headings (first row) as <h1> tags
        for (int col = 0; col <= maxCol; col++)
        {
            string heading = cells[0, col].StringValue;
            // If the cell is empty, skip it
            if (!string.IsNullOrEmpty(heading))
            {
                html.AppendLine($"<h1>{System.Web.HttpUtility.HtmlEncode(heading)}</h1>");
            }
        }

        // Write the rest of the data as an HTML table
        html.AppendLine("<table border=\"1\" cellspacing=\"0\" cellpadding=\"5\">");

        // Optional: write column headings inside the table as <th>
        html.AppendLine("<tr>");
        for (int col = 0; col <= maxCol; col++)
        {
            string colHeader = cells[0, col].StringValue;
            html.AppendLine($"<th>{System.Web.HttpUtility.HtmlEncode(colHeader)}</th>");
        }
        html.AppendLine("</tr>");

        // Write data rows starting from row index 1 (skip the heading row)
        for (int row = 1; row <= maxRow; row++)
        {
            html.AppendLine("<tr>");
            for (int col = 0; col <= maxCol; col++)
            {
                string cellValue = cells[row, col].StringValue;
                html.AppendLine($"<td>{System.Web.HttpUtility.HtmlEncode(cellValue)}</td>");
            }
            html.AppendLine("</tr>");
        }

        html.AppendLine("</table>");
        html.AppendLine("</body>");
        html.AppendLine("</html>");

        // Save the HTML to a file
        string outputPath = "Output.html";
        File.WriteAllText(outputPath, html.ToString(), Encoding.UTF8);

        Console.WriteLine($"HTML file generated at: {Path.GetFullPath(outputPath)}");
    }
}
