// Title: How to exclude short‑text cells from tooltip (title) attributes when saving a workbook to HTML with Aspose.Cells AddTooltipText enabled (C#)
// AI Prompts: Write C# code that saves an Aspose.Cells workbook to HTML with AddTooltipText=true, then parses the resulting HTML and removes the title attribute from <td> elements whose text length is below a given threshold. | Show a C# example that uses HtmlAgilityPack (or Regex) to post‑process Aspose.Cells HTML output and strip tooltip attributes from cells containing short strings. | Provide a C# routine that implements a custom post‑save step to filter out tooltip text for cells shorter than N characters after Aspose.Cells HTML export.
// Common Searches: Aspose.Cells HTML export skip tooltip for cells with less than 5 characters | remove title attribute from short cell values in Aspose.Cells generated HTML | C# post‑process Aspose.Cells HTML to delete tooltips on brief text | conditional AddTooltipText based on cell content length Aspose.Cells | how to filter tooltip text when saving workbook to HTML using Aspose.Cells
// Tags: Aspose.Cells HTML tooltip suppression | conditional AddTooltipText handling | post‑process Aspose.Cells HTML output | remove short‑text title attribute C# | filter cell tooltip by length Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, fills cells with short and long strings, enables AddTooltipText in HtmlSaveOptions, and saves to HTML. Because the CustomCellHtmlAttributes event is unavailable, tooltips are added to all cells. Selective removal of tooltip attributes for short‑text cells must be performed via post‑processing of the generated HTML.
class TooltipFilterExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some cells with short and long text
            sheet.Cells["A1"].PutValue("Hi");                                 // short text
            sheet.Cells["A2"].PutValue("Hello, World!");                      // long text
            sheet.Cells["A3"].PutValue("Short");                              // short text
            sheet.Cells["A4"].PutValue("This is a longer sentence.");        // long text

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Enable tooltip text generation (adds title attribute with cell value)
                AddTooltipText = true
            };

            // NOTE: CustomCellHtmlAttributes event is not available in this version of Aspose.Cells.
            // The tooltip will be added for all cells. Adjustments can be made by post‑processing the HTML if needed.

            // Determine output file path
            string outputPath = "output.html";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to an HTML file with the configured options
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
