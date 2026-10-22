// Title: Export an Excel workbook to HTML while preserving locale‑specific date formats using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a .xlsx file with Aspose.Cells, applies the workbook's CultureInfo to match the system locale, and saves it as HTML so that dates keep their original formatting. | Adjust the HTML export to force a specific CultureInfo (e.g., "fr-FR") and verify that all dates in the generated HTML appear in the chosen language format. | Add robust error handling that checks whether the source Excel file exists before conversion and logs a clear message if the file is missing.
// Common Searches: aspnet convert excel to html with culture specific dates using Aspose.Cells | c# Aspose.Cells preserve workbook locale when saving as html | how to set CultureInfo for Aspose.Cells workbook before html export | export .xlsx to .html keeping date format based on system locale
// Tags: Aspose.Cells HTML export with locale-aware dates | Workbook.Settings.CultureInfo for HTML conversion | C# export Excel to HTML preserving date format | HTMLSaveOptions culture-specific date rendering | validate input Excel file before Aspose.Cells conversion

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// Loads an Excel file, applies the current CultureInfo to the workbook settings, and saves it as HTML using Aspose.Cells, ensuring that dates are formatted according to the workbook's locale.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Ensure the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the existing Excel file
            Workbook workbook = new Workbook(inputPath);

            // Apply the workbook's locale settings (or set a specific culture)
            workbook.Settings.CultureInfo = CultureInfo.CurrentCulture; // replace with a specific CultureInfo if needed

            // Configure HTML save options (date formatting follows the workbook's culture automatically)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Save the workbook as an HTML file with the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
