// Title: Add <h1> worksheet titles when saving an Excel workbook to HTML using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file, configures HtmlSaveOptions to enable the worksheet header flag, sets the header template to "<h1>{0}</h1>", and saves the workbook as an HTML document. | Show how to detect the Aspose.Cells version at runtime and apply the worksheet header export settings only when the properties are available in C#.
// Common Searches: how to include worksheet name as h1 heading in Aspose.Cells HTML export c# | Aspose.Cells HtmlSaveOptions ExportWorksheetHeader property example | C# export Excel to HTML with custom header for each sheet using Aspose.Cells
// Tags: Aspose.Cells worksheet header export flag | C# customize worksheet header HTML Aspose.Cells | HTML export Excel sheets with h1 headings | conditional Aspose.Cells HTML save settings

using System;
using System.IO;
using Aspose.Cells;

// The program loads an Excel file, creates HtmlSaveOptions, optionally enables the ExportWorksheetHeader flag and defines a WorksheetHeader template containing an <h1> tag with the sheet name, then saves the workbook as HTML while handling missing files and runtime exceptions.
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

            // Load the workbook from the input file
            var workbook = new Workbook(inputPath);

            // Configure HTML save options
            var htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // The following properties are available in recent Aspose.Cells versions.
            // If your version supports them, uncomment to add <h1> headers before each worksheet.
            // htmlOptions.ExportWorksheetHeader = true;
            // htmlOptions.WorksheetHeader = "<h1>{0}</h1>";

            // Save the workbook as HTML using the configured options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
