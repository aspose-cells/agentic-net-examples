// Title: C# – Export Excel to a single HTML file and embed <!--PageBreak--> comments with Aspose.Cells
// Description: Demonstrates how to create a workbook, add a horizontal page break, save it as one HTML document using HtmlSaveOptions, and replace Aspose.Cells' <div class="pagebreak"></div> elements with custom <!--PageBreak--> markers for downstream processing.
// Keywords: Aspose.Cells HTML export | C# Excel to HTML | custom page break comment | HtmlSaveOptions single file | replace pagebreak div | horizontal page break worksheet
// Common Searches: Aspose.Cells export Excel to HTML single file C# | replace pagebreak div with comment Aspose.Cells | add manual page break before HTML export Aspose.Cells | how to get <!--PageBreak--> markers in exported HTML
// Developer Intent: Produce one HTML output from an Excel workbook and mark each page break with an HTML comment instead of the default div element.
// Use Cases: Post‑processing HTML to split content at page boundaries for PDF generation or pagination. | Embedding exported HTML into web pages that rely on comment markers for sectioning. | Automated batch conversion of workbooks where page‑break locations must be preserved for downstream scripts.
// AI Prompts: Generate C# code that uses Aspose.Cells to save a workbook as a single HTML file and swaps <div class="pagebreak"></div> with <!--PageBreak--> comments. | Show how to insert a horizontal page break at a specific row and configure HtmlSaveOptions for single‑file export. | Explain the steps to read the generated HTML, replace Aspose.Cells page‑break elements, and write the final file.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to create a workbook, add a horizontal page break, save it as one HTML document using HtmlSaveOptions, and replace Aspose.Cells' <div class="pagebreak"></div> elements with custom <!--PageBreak--> markers for downstream processing.
class ExportExcelToHtmlWithPageBreakComments
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with sample data
            for (int i = 0; i < 50; i++)
            {
                sheet.Cells[i, 0].PutValue($"Row {i + 1}");
            }

            // Add a manual horizontal page break after row 25
            // Use the HorizontalPageBreaks collection of the worksheet
            sheet.HorizontalPageBreaks.Add(25);

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                SaveAsSingleFile = true,   // Export as a single HTML file
                ExportPageHeaders = true   // Include page headers if any
            };

            // Save the workbook to a temporary HTML file
            string tempPath = Path.Combine(Path.GetTempPath(), "temp.html");
            workbook.Save(tempPath, saveOptions);

            // Ensure the temporary file exists before reading
            if (!File.Exists(tempPath))
                throw new FileNotFoundException("Temporary HTML file was not created.", tempPath);

            // Read the generated HTML content
            string html = File.ReadAllText(tempPath);

            // Aspose.Cells inserts a <div class="pagebreak"></div> for each page break.
            // Replace that element with an HTML comment marker for later processing.
            const string pageBreakDiv = "<div class=\"pagebreak\"></div>";
            const string commentMarker = "<!--PageBreak-->";
            html = html.Replace(pageBreakDiv, commentMarker);

            // Write the final HTML with page break comments to the desired output file
            string outputPath = "output_with_pagebreak_comments.html";
            File.WriteAllText(outputPath, html);

            Console.WriteLine($"HTML exported successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
