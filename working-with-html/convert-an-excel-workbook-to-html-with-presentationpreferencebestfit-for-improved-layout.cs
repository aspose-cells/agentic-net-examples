// Title: Convert an Excel .xlsx workbook to HTML with BestFit layout using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells convert .xlsx to HTML with best‑fit presentation | How to enable PresentationPreference in HtmlSaveOptions for Excel to HTML conversion | Export Excel workbook to HTML with improved layout using Aspose.Cells .NET | Aspose.Cells HTMLSaveOptions best fit layout example in C#
// Tags: Aspose.Cells HTML conversion PresentationPreference | C# HtmlSaveOptions best-fit layout | Excel to HTML export using Aspose.Cells | PresentationPreference enabled HTML output | BestFit HTML layout for Excel workbook

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlConversion
{
    // The sample loads an existing .xlsx file with Aspose.Cells, configures HtmlSaveOptions to set PresentationPreference = true for a BestFit layout, and saves the workbook as an HTML file, providing clearer rendering of the original Excel content.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths
            string inputFile = "input.xlsx";
            string outputFile = "output.html";

            // Validate input file existence
            if (!File.Exists(inputFile))
            {
                Console.Error.WriteLine($"Error: Input file \"{inputFile}\" does not exist.");
                return;
            }

            try
            {
                // Load the workbook from the input Excel file
                Workbook workbook = new Workbook(inputFile);
                Console.WriteLine($"Loaded workbook \"{inputFile}\" successfully.");

                // Create HTML save options and enable the best‑fit presentation preference
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
                htmlOptions.PresentationPreference = true; // Enables more beautiful layout (BestFit)

                // Save the workbook as an HTML file using the specified options
                workbook.Save(outputFile, htmlOptions);
                Console.WriteLine($"Workbook has been converted to HTML and saved as \"{outputFile}\".");
            }
            catch (Exception ex)
            {
                // Output any unexpected errors
                Console.Error.WriteLine("An error occurred during conversion:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
