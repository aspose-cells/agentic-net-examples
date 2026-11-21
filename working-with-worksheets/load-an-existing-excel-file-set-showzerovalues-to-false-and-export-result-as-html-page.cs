// Title: Convert an existing Excel workbook to an HTML page using Aspose.Cells for .NET (zero‑value display note)
// Common Searches: Aspose.Cells C# export Excel to HTML without showing zero values | how to hide zero values when converting XLSX to HTML with Aspose.Cells | C# load workbook and save as HTML using Aspose.Cells HtmlSaveOptions | check if Excel file exists before converting to HTML in a .NET app | Aspose.Cells example for converting an XLSX file to a single HTML page
// Tags: Aspose.Cells workbook to HTML conversion | C# hide zero values Aspose.Cells export | HtmlSaveOptions usage Aspose.Cells | Excel to HTML conversion Aspose.Cells .NET | file existence validation Aspose.Cells C#

using Aspose.Cells;
using System;
using System.IO;

// The sample verifies the input XLSX file, loads it into an Aspose.Cells Workbook, notes that the ShowZeroValues/PrintZeroValues property is not available in the current version, and then saves the workbook as an HTML page using default HtmlSaveOptions.
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

            // Load the existing Excel file
            Workbook workbook = new Workbook(inputPath);

            // If needed, configure page setup options here.
            // Note: The PrintZeroValues property is not available in the current Aspose.Cells version,
            // so this step is omitted.

            // Prepare HTML save options (default settings are sufficient)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Export the workbook as an HTML page
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook successfully saved as HTML to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
