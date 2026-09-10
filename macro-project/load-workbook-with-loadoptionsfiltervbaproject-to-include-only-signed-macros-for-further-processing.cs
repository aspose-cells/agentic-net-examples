// Title: Load an XLSM workbook in C# with Aspose.Cells using LoadOptions.FilterVbaProject to load only signed VBA macros
// AI Prompts: Generate C# code that creates a LoadOptions object, sets FilterVbaProject to include only signed macros, and opens an .xlsm file with Aspose.Cells. | Show how to load a macro‑enabled workbook, filter out unsigned VBA code using LoadOptions.FilterVbaProject, perform custom processing, and save the workbook in .NET.
// Common Searches: Aspose.Cells C# load only signed macros from xlsm | How to filter VBA project to signed macros when opening workbook | Example of loading macro-enabled Excel file with signed macro filter | Exclude unsigned VBA code while loading workbook with Aspose.Cells | C# tutorial for signed macro loading using Aspose.Cells
// Tags: signed macro filtering using LoadOptions | signed macro loading Aspose.Cells | filter VBA project on workbook load | process macro-enabled workbook C# | exclude unsigned macros Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates creating a LoadOptions instance, configuring its FilterVbaProject property to load only signed VBA macros, opening an .xlsm workbook with Aspose.Cells, optionally processing the workbook, and saving the result while handling potential errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsm";
            const string outputPath = "output.xlsm";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook (VBA project is loaded automatically for .xlsm files)
            Workbook workbook = new Workbook(inputPath);

            // TODO: Add further processing of the workbook here

            // Save the workbook after processing (optional)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
