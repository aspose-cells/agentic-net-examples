// Title: Use a C# using block to load and save an Aspose.Cells Workbook with deterministic disposal
// Common Searches: C# using statement for Aspose.Cells Workbook load and save | how to guarantee workbook disposal when working with Aspose.Cells in .NET | deterministic resource cleanup for Excel files using Aspose.Cells | wrap Aspose.Cells Workbook in using to avoid memory leaks | load Excel with LoadOptions and save using Aspose.Cells while ensuring disposal
// Tags: Aspose.Cells Workbook using block | load workbook with LoadOptions Aspose.Cells | save workbook to output path Aspose.Cells | password protected Excel handling Aspose.Cells | deterministic workbook disposal Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the input file, creates LoadOptions, and loads the workbook inside a C# using block to ensure deterministic disposal. Within the block the workbook can be processed and then saved to an output file. Specific CellsException handling detects password‑protected files, while a generic catch reports other errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Prepare load options; set Password here if known, e.g., loadOptions.Password = "pwd";
            var loadOptions = new LoadOptions();

            // Load the workbook with the specified options
            using (var workbook = new Workbook(inputPath, loadOptions))
            {
                // (Optional) manipulate the workbook here

                // Save the workbook to the output path
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
        }
        catch (CellsException ex)
        {
            // Handle password‑protected files or other Cells-specific errors
            if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("The workbook is password protected. Please provide a valid password.");
            }
            else
            {
                Console.WriteLine($"Aspose.Cells error: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            // General exception handling
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
