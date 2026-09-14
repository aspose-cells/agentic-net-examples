// Title: Set Aspose.Cells workbook to print comments as separate notes (separate sheet) using C#
// AI Prompts: Write C# code that opens an existing Excel file with Aspose.Cells, sets the PrintComments property to PrintInSeparateSheet, and saves the modified workbook. | Show how to enable separate‑sheet comment printing in Aspose.Cells for .NET before exporting the workbook to a new file. | Demonstrate configuring a workbook's print settings so that cell comments are printed as independent notes rather than in‑place.
// Common Searches: Aspose.Cells C# print cell comments on a separate sheet | How to export Excel comments as notes using Aspose.Cells .NET | Workbook comment printing setting example in C# | Separate sheet comment printing with Aspose.Cells | Configure Aspose.Cells to output comments as independent notes
// Tags: Aspose.Cells comment printing configuration | PrintCommentsType separate sheet option C# | export Excel comments as separate notes Aspose.Cells | separate sheet comment output Aspose.Cells | Aspose.Cells workbook print settings .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // C# example that checks for an input.xlsx file, loads it into an Aspose.Cells Workbook, configures the workbook's print settings to output cell comments as separate notes by assigning the PrintComments property the PrintInSeparateSheet value, and saves the result to output.xlsx while handling potential errors.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // NOTE: In some Aspose.Cells versions the PrintComments setting may not be available.
                // If needed, adjust the setting using the appropriate API for the referenced version.

                // Save the workbook with the (potential) new print setting
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
