// Title: C# – Convert HTML to Excel with Aspose.Cells and map <hr> tags to worksheet page breaks
// Description: Loads an HTML file into an Aspose.Cells workbook, detects every <hr> element, calculates the corresponding row, inserts a horizontal page break at that row, and saves the result as an XLSX file.
// Keywords: Aspose.Cells HTML to Excel | C# convert HTML to XLSX | preserve <hr> as page break | horizontal page break Aspose.Cells | HTML pagination Excel | worksheet page break from HTML
// Common Searches: Aspose.Cells convert HTML to Excel with page breaks | C# add worksheet page break for <hr> tag | map HTML horizontal rule to Excel page break | how to keep HTML sections separate when exporting to XLSX | C# detect <hr> in HTML and insert Excel page break
// Developer Intent: Transform an HTML document into an Excel workbook while converting each <hr> element into a worksheet horizontal page break.
// Use Cases: Generate printable Excel reports from web pages where <hr> separates sections, ensuring each section starts on a new printed page. | Automate conversion of HTML‑based invoices to Excel while preserving visual separators defined by <hr> tags. | Create Excel versions of online tutorials or documentation that keep chapter breaks for proper pagination.
// AI Prompts: Improve the algorithm for locating <hr> tags and determining the exact worksheet row, handling different newline styles and nested HTML elements. | Explain how to manage consecutive or nested <hr> tags so Aspose.Cells adds appropriate page breaks without exceeding the data range. | Provide unit‑test examples that verify correct insertion of horizontal page breaks for various HTML inputs containing <hr> elements.

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// Loads an HTML file into an Aspose.Cells workbook, detects every <hr> element, calculates the corresponding row, inserts a horizontal page break at that row, and saves the result as an XLSX file.
class HtmlToExcelConverter
{
    static void Main()
    {
        string htmlFile = "input.html";
        string excelFile = "output.xlsx";

        try
        {
            // Verify that the input HTML file exists
            if (!File.Exists(htmlFile))
            {
                Console.WriteLine($"Error: The file '{htmlFile}' was not found.");
                return;
            }

            // Load the HTML file into a workbook (Aspose.Cells detects the format automatically)
            Workbook workbook = new Workbook(htmlFile);
            Worksheet sheet = workbook.Worksheets[0];

            // Read the raw HTML text
            string htmlContent = File.ReadAllText(htmlFile);

            // Find all <hr> tags (case‑insensitive)
            MatchCollection hrMatches = Regex.Matches(htmlContent, @"<hr\s*/?>", RegexOptions.IgnoreCase);

            foreach (Match match in hrMatches)
            {
                // Estimate the row where the <hr> appears by counting line‑feed characters before the match
                int precedingLineFeeds = htmlContent.Substring(0, match.Index).Split('\n').Length - 1;
                int rowIndex = Math.Max(0, precedingLineFeeds);
                if (rowIndex > sheet.Cells.MaxDataRow)
                    rowIndex = sheet.Cells.MaxDataRow;

                // Insert a horizontal page break after the identified row
                // Use HorizontalPageBreaks collection (available in recent Aspose.Cells versions)
                sheet.HorizontalPageBreaks.Add(rowIndex);
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(excelFile);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as an Excel file
            workbook.Save(excelFile, SaveFormat.Xlsx);
            Console.WriteLine($"Conversion succeeded. Excel file saved as '{excelFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
