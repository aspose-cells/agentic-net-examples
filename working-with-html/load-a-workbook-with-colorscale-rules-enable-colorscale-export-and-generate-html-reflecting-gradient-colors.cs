// Title: Export an Excel workbook with ColorScale conditional formatting to HTML using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx file containing ColorScale conditional formatting and saves it as an HTML file with gradient colors using Aspose.Cells. | Demonstrate how to configure Aspose.Cells HtmlSaveOptions so that ColorScale rules are retained in the produced HTML output. | Provide a C# example that validates the input workbook, handles missing‑file errors, and converts the workbook with ColorScale rules to HTML via Aspose.Cells.
// Common Searches: Aspose.Cells C# export workbook with conditional formatting color scales to HTML | How to keep Excel gradient color scales when saving as HTML with Aspose.Cells | C# HtmlSaveOptions preserve ColorScale rules in generated HTML | Convert Excel file containing ColorScale CF to HTML using Aspose.Cells .NET | Saving Excel conditional formatting as HTML Aspose.Cells example
// Tags: Aspose.Cells HtmlSaveOptions for ColorScale preservation | C# generate HTML from Excel with gradient conditional formatting | export Excel conditional formatting to HTML using .NET API | load workbook and convert to HTML preserving color scales | conditional formatting color scale HTML export Aspose

using System;
using System.IO;
using Aspose.Cells;

// The sample checks for the input Excel file, loads it with Aspose.Cells, uses default HtmlSaveOptions (which include conditional formatting), and saves the workbook as HTML. The resulting HTML retains the gradient colors defined by ColorScale conditional formatting rules, with error handling for missing files.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook containing ColorScale conditional formatting rules
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options (conditional formatting is exported by default)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Save the workbook as HTML; the generated HTML reflects the gradient colors defined by ColorScale rules
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
