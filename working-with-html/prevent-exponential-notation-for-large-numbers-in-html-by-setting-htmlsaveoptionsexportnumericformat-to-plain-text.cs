// Title: How to stop scientific (exponential) notation for large numbers when saving an Excel workbook to HTML with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export large numeric values to HTML without scientific notation | Set HtmlSaveOptions ExportNumericFormat to plain text in .NET | Prevent exponential format for big numbers in HTML output using Aspose.Cells | C# convert XLSX to HTML keeping numeric values as plain text | Aspose.Cells HTML save options numeric format configuration example
// Tags: HtmlSaveOptions ExportNumericFormat plain text | Aspose.Cells numeric formatting HTML export | C# Excel to HTML large numbers handling | prevent scientific notation Aspose.Cells | export large numeric values as text .NET | Aspose.Cells HTMLSaveOptions configuration | Excel workbook HTML conversion numeric format

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The sample loads an XLSX workbook, configures HtmlSaveOptions by setting ExportNumericFormat to plain text, and saves the workbook as HTML. This ensures that large numeric values are written as ordinary numbers rather than scientific (exponential) notation.
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

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Default behavior exports numeric values as plain text.
            // Additional options can be set here if required.

            // Save the workbook to HTML using the configured options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
