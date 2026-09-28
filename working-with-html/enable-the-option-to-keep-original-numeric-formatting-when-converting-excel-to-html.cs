// Title: How to keep original numeric formatting when converting an Excel workbook to HTML with Aspose.Cells for .NET
// AI Prompts: Write C# code that configures HtmlSaveOptions to retain Excel number formats during HTML export using Aspose.Cells. | Update the sample program to explicitly enable numeric format preservation when saving a workbook as HTML. | Identify the HtmlSaveOptions property that controls number format export and state its default behavior.
// Common Searches: Aspose.Cells preserve number format Excel to HTML C# example | HtmlSaveOptions keep numeric display unchanged during conversion | C# export Excel workbook to HTML with original number formatting | how to retain Excel cell number formatting in HTML output using Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions numeric format | preserve Excel number formatting HTML conversion | C# export workbook to HTML with original formatting | keep cell numeric display Aspose.Cells | HTML save options retain number formats

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example checks for the source Excel file, loads it into a Workbook, creates an HtmlSaveOptions object (numeric formats are retained by default), and saves the workbook as HTML, handling any runtime exceptions.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputFile = "input.xlsx";
                string outputFile = "output.html";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    return;
                }

                // Load the source Excel workbook
                Workbook workbook = new Workbook(inputFile);

                // Set HTML save options (numeric formats are preserved by default)
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

                // Save the workbook as HTML using the configured options
                workbook.Save(outputFile, htmlOptions);
                Console.WriteLine($"Workbook successfully saved as HTML to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
