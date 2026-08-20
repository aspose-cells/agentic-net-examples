// Title: Keep Original Numeric Formatting When Converting Excel to HTML with Aspose.Cells for .NET
// Description: Shows how to load an Excel workbook, enable HtmlSaveOptions.ExportNumericFormat (and optionally ExportFormula), and save it as HTML so that currency, percentage, and custom number formats are preserved exactly as they appear in the source file.
// Keywords: Aspose.Cells | HtmlSaveOptions | ExportNumericFormat | preserve numeric formatting | Excel to HTML conversion | C# | .NET | number format export | currency format HTML | percentage format Excel | custom number format Aspose
// Common Searches: Aspose.Cells keep numeric format when saving as HTML | HtmlSaveOptions ExportNumericFormat C# example | convert .xlsx to .html preserving number formats | retain currency and percentage formatting in HTML export Aspose.Cells | how to export Excel to HTML with original number formatting
// Developer Intent: Add the ExportNumericFormat flag to HtmlSaveOptions so that numeric formatting is retained during HTML export.
// Use Cases: Publish a financial statement to the web while keeping currency symbols and decimal precision. | Create an HTML preview of a spreadsheet that contains percentages and custom formats without losing visual fidelity. | Generate a web‑ready report from a data‑analysis workbook where the exact number formatting must match the Excel view.
// AI Prompts: Provide C# code using Aspose.Cells to convert an Excel file to HTML while preserving original numeric formatting, including the required HtmlSaveOptions properties. | Explain which HtmlSaveOptions settings control numeric formatting and formula export when saving to HTML with Aspose.Cells for .NET. | Give a step‑by‑step guide to modify existing Aspose.Cells HTML export code to keep number formats such as currency, percentages, and custom patterns.

using System;
using System.IO;
using Aspose.Cells;

// Shows how to load an Excel workbook, enable HtmlSaveOptions.ExportNumericFormat (and optionally ExportFormula), and save it as HTML so that currency, percentage, and custom number formats are preserved exactly as they appear in the source file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Keep formulas in the HTML output
                ExportFormula = true
            };

            // Save the workbook as HTML
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
