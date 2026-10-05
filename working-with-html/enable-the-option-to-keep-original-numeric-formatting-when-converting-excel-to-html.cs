// Title: How to preserve original numeric formatting when converting an Excel workbook to HTML with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file and saves it as HTML, configuring HtmlSaveOptions to keep the workbook's original numeric formats. | Show how to enable the numeric‑format export option in Aspose.Cells HtmlSaveOptions so that numbers retain their Excel formatting in the generated HTML.
// Common Searches: Aspose.Cells keep original number format when converting Excel to HTML in C# | C# HtmlSaveOptions preserve numeric formatting Aspose.Cells example | how to retain Excel numeric styles in HTML output using Aspose.Cells .NET | export Excel workbook to HTML without losing number formatting Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions numeric format | Excel to HTML numeric formatting C# | preserve number format Aspose.Cells HTML conversion | C# save workbook as HTML with original formatting

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the presence of input.xlsx, loads it into an Aspose.Cells Workbook, creates HtmlSaveOptions (which retain numeric formatting by default), and saves the workbook as output.html, handling missing files and runtime errors gracefully.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the source Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options (default settings preserve numeric formatting)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Save the workbook as an HTML file using the configured options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
