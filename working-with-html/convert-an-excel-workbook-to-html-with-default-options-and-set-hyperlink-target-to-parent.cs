// Title: Convert an Excel workbook to HTML with Aspose.Cells in C# and set hyperlink targets to _parent using default HtmlSaveOptions
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, applies the default HtmlSaveOptions, and saves it as an .html file. | Modify the Aspose.Cells HTML export example to configure every hyperlink to use the "_parent" target attribute while keeping all other save options unchanged. | Add comprehensive file‑existence checking and exception handling to the C# routine that converts an Excel workbook to HTML with Aspose.Cells.
// Common Searches: how to export Excel to HTML with Aspose.Cells C# default settings | Aspose.Cells set hyperlink target _parent in HTML output | C# convert .xlsx to .html using HtmlSaveOptions default | Aspose.Cells HtmlSaveOptions hyperlink target attribute example | save workbook as HTML with Aspose.Cells and handle missing file
// Tags: Aspose.Cells HTML conversion default options | C# hyperlink target _parent Aspose.Cells | HtmlSaveOptions configuration Aspose.Cells | Excel to HTML export Aspose.Cells .NET | file existence validation Aspose.Cells conversion

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file, creates a HtmlSaveOptions object with default settings, configures hyperlinks to open in the parent frame, and saves the workbook as an HTML file while handling missing files and runtime exceptions.
class ExcelToHtmlConverter
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file not found at '{inputPath}'.");
            return;
        }

        try
        {
            // Load the Excel workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Create HTML save options with default settings
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Save the workbook as an HTML file using the specified options
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook successfully converted to HTML: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
