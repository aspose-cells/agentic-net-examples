// Title: Export an Excel workbook to HTML with Aspose.Cells C# while handling missing HtmlCrossType.MSExport
// AI Prompts: Write C# code that loads a .xlsx file, checks if HtmlCrossType.MSExport is supported, and saves the workbook as HTML using Aspose.Cells with a fallback when the property is unavailable. | Modify the Aspose.Cells example to detect the library version and conditionally apply HtmlCrossType for Excel‑compatible HTML output. | Describe strategies for achieving Excel‑style HTML export in Aspose.Cells when the HtmlCrossType property does not exist in the current SDK.
// Common Searches: how to export Excel to HTML with Aspose.Cells C# preserving formatting | Aspose.Cells HtmlCrossType MSExport not available in current version | C# save workbook as HTML using Aspose.Cells with fallback for missing HtmlCrossType | handle FileNotFoundException when converting .xlsx to HTML with Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions HTML export | C# export Excel to HTML Aspose.Cells | HtmlCrossType MSExport fallback | Excel native HTML style Aspose.Cells | Workbook.Save HTML Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies the existence of an input .xlsx file, loads it into an Aspose.Cells Workbook, configures HtmlSaveOptions for HTML output, notes that HtmlCrossType.MSExport is unavailable in the current SDK, and saves the workbook as an HTML file while handling exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputFile = "input.xlsx";
            const string outputFile = "output.html";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: The file '{inputFile}' was not found.");
                return;
            }

            // Load the existing Excel workbook
            Workbook workbook = new Workbook(inputFile);

            // Configure HTML save options (using default settings)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Note: HtmlCrossType property is not available in the current Aspose.Cells version,
            // so default cross‑type handling is used.

            // Export the workbook to an HTML file
            workbook.Save(outputFile, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
