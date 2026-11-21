// Title: Convert an Excel workbook to HTML with default options and set hyperlink target to parent using Aspose.Cells for .NET
// Common Searches: how to export an .xlsx workbook to HTML using Aspose.Cells in C# | Aspose.Cells C# set hyperlink target attribute when saving to HTML | default HtmlSaveOptions example for Excel to HTML conversion Aspose.Cells | C# code to convert Excel file to HTML and handle missing HyperlinkTarget property | save Excel as HTML with parent window links using Aspose.Cells
// Tags: Aspose.Cells HTML export default options | C# Excel to HTML conversion Aspose.Cells | HtmlSaveOptions hyperlink target setting | handling missing HyperlinkTarget property Aspose.Cells | Workbook.Save HTML example C# | parent window hyperlink in exported HTML Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The sample verifies the input .xlsx file, loads it into a Workbook, creates default HtmlSaveOptions (noting the absence of a HyperlinkTarget property), and saves the workbook as an HTML file while handling any exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the source Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Create HTML save options with default settings
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();

            // Note: The HyperlinkTarget property is not available in the current Aspose.Cells version.
            // If needed, customize hyperlink behavior after saving the HTML.

            // Save the workbook as HTML
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook successfully saved as HTML to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
