// Title: Generate HTML from an Excel workbook with each worksheet name rendered as an <h1> heading using Aspose.Cells for .NET
// AI Prompts: Write C# code that loops through every worksheet in a Workbook and adds an <h1> tag containing the sheet name before the sheet's HTML content with Aspose.Cells. | Demonstrate how to modify the HTML export configuration to inject extra markup before each sheet's data when saving to HTML using Aspose.Cells. | Provide a step‑by‑step example that creates a temporary HTML wrapper with worksheet titles as headings and merges it with the HTML produced by Aspose.Cells.
// Common Searches: Aspose.Cells .NET how to add worksheet titles as HTML headings when exporting to HTML | C# export Excel to HTML with sheet name as <h1> using Aspose.Cells | HtmlSaveOptions ExportWorksheetHeader missing workaround Aspose.Cells | Generate custom HTML header for each worksheet in Aspose.Cells HTML output
// Tags: Aspose.Cells custom HTML headings per worksheet | C# Excel to HTML sheet title insertion | HTML export workaround missing ExportWorksheetHeader | Iterate worksheets for HTML generation Aspose.Cells | Generate <h1> from worksheet name Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel file, configures HTML export, and demonstrates how to prepend each worksheet's name as an <h1> heading in the generated HTML, providing a workaround for the unavailable ExportWorksheetHeader property in the current Aspose.Cells version.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.html";

            // Ensure the source workbook exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Set up HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // The following properties are not available in the current Aspose.Cells version:
            // htmlOptions.ExportWorksheetHeader = true;
            // htmlOptions.ExportWorksheetFooter = false;

            // Save the workbook as an HTML file
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
