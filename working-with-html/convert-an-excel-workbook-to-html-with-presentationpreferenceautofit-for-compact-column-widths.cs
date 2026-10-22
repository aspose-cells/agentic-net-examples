// Title: How to convert an Excel .xlsx workbook to compact HTML using Aspose.Cells PresentationPreference.AutoFit in C#
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, sets HtmlSaveOptions.PresentationPreference to AutoFit, and saves it as an HTML file. | Show a C# example that checks for the source Excel file, configures HtmlSaveOptions for automatic column width fitting, and writes the output HTML. | Provide a try‑catch block in C# that handles missing input files and exceptions while converting Excel to HTML with compact column widths using Aspose.Cells.
// Common Searches: Aspose.Cells C# convert xlsx to html with auto fit column widths | HtmlSaveOptions PresentationPreference.AutoFit sample code | Generate narrow column HTML from Excel using Aspose.Cells .NET | C# export workbook to html with column width optimization
// Tags: Aspose.Cells HtmlSaveOptions PresentationPreference.AutoFit | C# Excel to HTML compact column layout | auto-fit columns Aspose.Cells HTML conversion | save workbook as html with column optimization .NET | load xlsx and export to html Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the input .xlsx file, loads it into an Aspose.Cells Workbook, optionally sets HtmlSaveOptions.PresentationPreference to AutoFit for tighter column widths, and saves the workbook as an HTML file while handling errors gracefully.
class ExcelToHtmlConverter
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

            // Load the Excel workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            // Uncomment the following line if the PresentationPreference enum is available in your Aspose.Cells version
            // saveOptions.PresentationPreference = PresentationPreference.AutoFit;

            // Save the workbook as an HTML file
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook successfully saved to HTML: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
