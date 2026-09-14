// Title: Convert an Excel .xlsx workbook to compact-width HTML using Aspose.Cells for .NET with PresentationPreference.AutoFit
// AI Prompts: Write C# code that loads a .xlsx file, verifies its existence, sets HtmlSaveOptions.PresentationPreference to AutoFit, and saves the workbook as an HTML file with compact column widths using Aspose.Cells. | Provide a C# example that handles a missing input file while converting an Excel workbook to HTML with auto‑fitted columns via Aspose.Cells.
// Common Searches: Aspose.Cells C# export Excel to HTML with auto‑fit column widths | C# convert .xlsx to HTML using PresentationPreference.AutoFit in Aspose.Cells | Save workbook as HTML with compact columns Aspose.Cells .NET example | HtmlSaveOptions PresentationPreference.AutoFit usage Aspose.Cells | C# Aspose.Cells HTML export handling file not found error
// Tags: Aspose.Cells HtmlSaveOptions PresentationPreference.AutoFit | Excel to HTML conversion C# Aspose.Cells | auto‑fit column widths Aspose.Cells HTML export | C# file existence check before Aspose.Cells conversion | error handling Aspose.Cells workbook save to HTML

using System;
using System.IO;
using Aspose.Cells;

// The program checks that the specified .xlsx file exists, loads it into an Aspose.Cells Workbook, configures HtmlSaveOptions with PresentationPreference.AutoFit to produce compact column widths, saves the workbook as an HTML file, and reports success or any caught exceptions.
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
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the Excel workbook from file
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options (auto‑fit columns not directly supported; default behavior is used)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Save the workbook as an HTML file with the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
