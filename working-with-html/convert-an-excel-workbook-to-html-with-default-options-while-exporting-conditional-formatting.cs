// Title: Convert an Excel .xlsx workbook to HTML in C# with Aspose.Cells default options while preserving conditional formatting
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells and saves it as an HTML file using the default HtmlSaveOptions, ensuring conditional formatting is retained. | Show how to export an Excel workbook to HTML in .NET without customizing any save options, relying on Aspose.Cells' built‑in behavior to include conditional formatting.
// Common Searches: asp.net convert xlsx to html using aspose.cells default settings | c# export excel workbook to html keeping conditional formatting | how to save workbook as html with Aspose.Cells without custom options | default HtmlSaveOptions export conditional formatting Aspose.Cells
// Tags: Aspose.Cells default HtmlSaveOptions export | C# convert Excel to HTML with conditional formatting | SaveFormat.Html usage in Aspose.Cells | Export .xlsx to .html using Aspose.Cells | HTML conversion preserving Excel conditional formatting

using Aspose.Cells;
using System;
using System.IO;

// The sample checks for an input .xlsx file, loads it into an Aspose.Cells Workbook, creates a HtmlSaveOptions instance (which exports conditional formatting by default), and saves the workbook as an .html file, handling any errors that may occur.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.html";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the Excel workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Create HTML save options (conditional formatting is exported by default)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Save the workbook as an HTML file using the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
