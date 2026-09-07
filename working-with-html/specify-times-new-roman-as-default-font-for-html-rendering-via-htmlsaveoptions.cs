// Title: How to set Times New Roman as the default font when exporting an Excel workbook to HTML using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file, sets HtmlSaveOptions.DefaultFontName to "Times New Roman", and saves the workbook as an HTML document with Aspose.Cells. | Show a .NET example that configures the HTML export options to apply a custom default font for Excel-to-HTML conversion using Aspose.Cells.
// Common Searches: Aspose.Cells C# export Excel to HTML with Times New Roman as default font | How to change the default font for HTML conversion in Aspose.Cells .NET | HtmlSaveOptions DefaultFontName usage example for Excel to HTML | Set default font for HTML output when saving a workbook with Aspose.Cells
// Tags: Aspose.Cells HTML export default font | C# HtmlSaveOptions custom font | Excel to HTML conversion with specific font | set default font Aspose.Cells .NET | HtmlSaveOptions default font property

using System;
using System.IO;
using Aspose.Cells;

// Loads input.xlsx, configures HtmlSaveOptions.DefaultFontName to "Times New Roman", and saves the workbook as output.html using Aspose.Cells.
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options to use "Times New Roman" as the default font
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.DefaultFontName = "Times New Roman";

            // Save the workbook as HTML with the specified default font
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook successfully saved as HTML to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
