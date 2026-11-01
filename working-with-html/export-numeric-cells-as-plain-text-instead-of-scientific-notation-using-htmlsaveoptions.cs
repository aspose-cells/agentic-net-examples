// Title: How to export Excel numeric cells as plain text (avoid scientific notation) when saving to HTML with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export Excel numbers as plain text in HTML | how to stop scientific notation when saving workbook to HTML with Aspose.Cells | HtmlSaveOptions keep original number format during HTML conversion | save Excel to HTML without exponential format using Aspose.Cells .NET | display numeric cells as text instead of scientific notation in HTML output
// Tags: Aspose.Cells HTML numeric display | plain text numeric export C# | HtmlSaveOptions number format handling | Excel to HTML number formatting | prevent exponential notation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The sample loads an Excel workbook, configures HtmlSaveOptions, and saves the workbook as HTML. It notes that the ExportNumericDataAsString option is not available for HTML, so numeric cells are written using their displayed format, which may include scientific notation. To ensure numbers appear as plain text, format the cells in Excel before exporting.
class ExportNumericAsPlainText
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.html";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Note: In current Aspose.Cells versions, ExportNumericDataAsString is not available for HTML.
            // Numeric values will be saved using their displayed format.

            // Save the workbook as HTML
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
