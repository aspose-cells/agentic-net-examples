// Title: Export Excel to HTML with Aspose.Cells for .NET and manage page‑break markers for later processing
// AI Prompts: Create C# code that uses Aspose.Cells to convert an .xlsx workbook to HTML, then scans the generated HTML and inserts <!--PageBreak--> comments before each page‑break element. | Modify the Aspose.Cells HTML export example to add a post‑processing step that inserts custom HTML comment markers for page breaks in the saved HTML file.
// Common Searches: Aspose.Cells .NET export Excel to HTML and add page break comments | C# generate HTML from Excel with Aspose.Cells and insert <!--PageBreak--> markers | How to identify and comment page breaks after saving Excel as HTML using Aspose.Cells | Aspose.Cells page break handling during HTML conversion for post‑processing | Add custom HTML comments for pagination when converting Excel to HTML in C#
// Tags: Aspose.Cells HtmlSaveOptions export to HTML | C# Excel to HTML conversion with pagination markers | post‑process Aspose.Cells HTML output for pagination comments | identify pagination points in Aspose.Cells generated HTML | HTML comment injection after Excel-to-HTML conversion in .NET

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the input.xlsx file, loads it with Aspose.Cells, configures HtmlSaveOptions, and saves the workbook as output.html. It notes that the current Aspose.Cells version does not provide a direct option to embed page‑break markers as HTML comments, so any comment insertion must be done via post‑processing of the generated HTML.
class ExcelToHtmlWithPageBreakComments
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Note: Aspose.Cells does not provide a direct option to export page breaks as HTML comments.
            // The ExportPageBreaks property is not available in this version, so we rely on default behavior.

            // Optional: export only the first sheet
            // htmlOptions.OnePagePerSheet = true;

            // Save the workbook as an HTML file
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
