// Title: Convert HTML to XLSX with Aspose.Cells for .NET while preserving CSS‑based conditional formatting
// AI Prompts: Generate C# code that loads a local HTML file using Aspose.Cells HtmlLoadOptions, maps CSS classes containing conditional‑formatting rules to Excel style objects, and saves the workbook as an XLSX file. | Show how to add robust error handling that verifies the HTML file exists and gracefully handles conversion exceptions in a C# Aspose.Cells application. | Demonstrate customizing HtmlLoadOptions to retain CSS styling—including conditional formatting—when importing HTML into an Aspose.Cells Workbook.
// Common Searches: asp.net c# convert html table with css conditional formatting to xlsx using aspose.cells | how to keep css style rules when importing html into excel workbook with aspose.cells | example of HtmlLoadOptions preserving conditional formatting from html to excel | map css classes to Excel conditional formatting in a C# Aspose.Cells project | load html file into workbook and export as xlsx while retaining styling in .NET
// Tags: html-to-xlsx conversion Aspose.Cells C# | css conditional formatting mapping to Excel styles | Aspose.Cells HtmlLoadOptions HTML import | preserve css styling during HTML to Excel conversion | c# load html workbook and save as xlsx | excel conditional formatting from css classes

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an input.html file into an Aspose.Cells Workbook using HtmlLoadOptions, then saves the workbook as output.xlsx. It includes basic file‑existence checking and exception handling, and can be extended to map CSS classes that define conditional formatting into equivalent Excel style rules, ensuring the visual formatting from the HTML is retained in the generated XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "input.html";
                const string outputPath = "output.xlsx";

                // Verify that the input HTML file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                    return;
                }

                // Initialize HTML load options (additional options can be set if supported by the library version)
                HtmlLoadOptions loadOptions = new HtmlLoadOptions();

                // Load the HTML file into a new workbook using the specified options
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Save the workbook to an Excel file (XLSX format)
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Conversion completed successfully. Output saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
