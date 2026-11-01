// Title: Enable custom CSS properties during HTML export of an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file, sets HtmlSaveOptions.EnableCustomProperties to true, and saves the workbook as an HTML file with Aspose.Cells. | Show a .NET example configuring Aspose.Cells HtmlSaveOptions to preserve user‑defined CSS variables while converting Excel to HTML. | Explain how to modify Aspose.Cells HTML export settings to include custom CSS properties in the generated HTML output.
// Common Searches: Aspose.Cells .NET enable custom CSS properties in HTML export | HtmlSaveOptions.EnableCustomProperties true example C# | Preserve custom style definitions when converting Excel to HTML with Aspose.Cells | How to keep custom CSS variables after saving workbook as HTML using Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions EnableCustomProperties | C# export Excel to HTML with custom CSS | preserve custom CSS Aspose.Cells | HTML export options Aspose.Cells .NET | custom CSS properties Excel to HTML conversion

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, creates HtmlSaveOptions, sets EnableCustomProperties to true so that any user‑defined CSS variables are retained, and saves the workbook as an HTML file while handling missing file and runtime errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML export options (default options are sufficient for basic export)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Export the workbook to HTML using the configured options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
