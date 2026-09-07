// Title: Efficiently export a large Excel workbook to HTML with Aspose.Cells in C# (fallback when HtmlCrossType.Cross is unavailable)
// AI Prompts: Generate C# code that loads an .xlsx file, verifies its existence, and saves it as an HTML file using Aspose.Cells HtmlSaveOptions with performance considerations. | Show how to implement a graceful fallback to default HtmlSaveOptions when the HtmlCrossType.Cross property is missing in the current Aspose.Cells version. | Provide robust error‑handling that captures and logs any exceptions occurring during the workbook‑to‑HTML conversion.
// Common Searches: Aspose.Cells C# export large workbook to HTML with high performance | HtmlCrossType.Cross not found how to save Excel as HTML using Aspose.Cells | C# sample code for converting .xlsx to .html with Aspose.Cells HtmlSaveOptions | Best practices for handling missing HtmlCrossType in Aspose.Cells | How to check file existence before converting Excel to HTML in C#
// Tags: Aspose.Cells HTML export large workbook | HtmlSaveOptions performance tuning | handling missing HtmlCrossType | C# Excel to HTML conversion | exception handling Aspose.Cells save

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example verifies that the source .xlsx file exists, loads it into an Aspose.Cells Workbook, configures HtmlSaveOptions (since HtmlCrossType.Cross is unavailable), saves the workbook as an HTML file, and includes try‑catch logic to report any conversion errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        // Verify that the input workbook exists to avoid FileNotFoundException.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the source workbook.
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options. The HtmlCrossType property is not available
            // in the current Aspose.Cells version, so default rendering settings are used.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Export the workbook to HTML using the configured options.
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook successfully saved as HTML to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
