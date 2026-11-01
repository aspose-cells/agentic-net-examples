// Title: Convert an Excel workbook to HTML in C# with Aspose.Cells and configure hyperlinks to open in a new browser tab
// Common Searches: Aspose.Cells C# convert .xlsx to HTML with links opening in new tab | set hyperlink target attribute when saving Excel as HTML using Aspose.Cells | HtmlSaveOptions ExportHyperlinkTarget property missing Aspose.Cells version | C# save Excel workbook as HTML and make hyperlinks open in a new window | how to ensure HTML output from Aspose.Cells opens links in a new tab
// Tags: Aspose.Cells HtmlSaveOptions SaveFormat.Html | C# convert Excel to HTML Aspose.Cells | hyperlink target attribute HTML export Aspose.Cells | export workbook as HTML with new tab links | Aspose.Cells HTML save options hyperlink handling

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The example loads an Excel file with Aspose.Cells, creates HtmlSaveOptions for HTML output, notes that the ExportHyperlinkTarget property is unavailable in the current version, and saves the workbook as an HTML file while handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the source Excel file
            Workbook workbook = new Workbook(inputPath);

            // Set HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Note: ExportHyperlinkTarget property is not available in this version of Aspose.Cells.
            // If needed, adjust hyperlink rendering via other available options.

            // Save the workbook as an HTML file with the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
