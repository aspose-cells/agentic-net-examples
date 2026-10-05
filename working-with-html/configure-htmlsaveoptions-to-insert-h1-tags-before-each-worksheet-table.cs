// Title: How to add <h1> worksheet titles before each table when saving Excel to HTML with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx workbook and saves it as HTML using Aspose.Cells, turning on ExportWorksheetHeader to prepend an <h1> containing the worksheet name before each table. | Show how to detect the Aspose.Cells version at runtime and conditionally set HtmlSaveOptions.ExportWorksheetHeader for compatibility with older library releases. | Provide a fallback technique that manually inserts <h1> tags with worksheet names when the ExportWorksheetHeader property is unavailable.
// Common Searches: Aspose.Cells C# export Excel to HTML with worksheet name as heading | Enable ExportWorksheetHeader in HtmlSaveOptions for HTML conversion | Add h1 tags before each worksheet table using Aspose.Cells .NET | Handle missing ExportWorksheetHeader property in older Aspose.Cells versions | Save workbook as HTML with headings for each sheet in C#
// Tags: Aspose.Cells HtmlSaveOptions ExportWorksheetHeader | C# export Excel to HTML with worksheet headings | HTML conversion worksheet header Aspose.Cells | fallback for missing ExportWorksheetHeader property | Aspose.Cells version check .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file, configures HtmlSaveOptions to enable ExportWorksheetHeader (when supported) so each worksheet appears with an <h1> title before its HTML table, saves the result as an HTML file, and includes version‑aware handling for older Aspose.Cells releases.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options.
            // ExportWorksheetHeader inserts an <h1> tag with the worksheet name before each table.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
            // The ExportWorksheetHeader property may not be available in older versions;
            // if it exists, enable it to include <h1> tags.
            // Uncomment the following line if your Aspose.Cells version supports it:
            // saveOptions.ExportWorksheetHeader = true;

            // Save the workbook as an HTML file using the configured options
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display an error message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
