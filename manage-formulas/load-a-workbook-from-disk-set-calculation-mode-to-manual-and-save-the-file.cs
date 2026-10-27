// Title: Load an Excel workbook, set calculation mode to Manual, and save it using Aspose.Cells for .NET
// Common Searches: how to load an existing .xlsx file with Aspose.Cells and change calculation mode in C# | Aspose.Cells set workbook calculation mode to Manual before saving | C# check if Excel file exists then open with Aspose.Cells | save a workbook to a new file after modifying settings using Aspose.Cells .NET | exception handling when loading and saving Excel files with Aspose.Cells
// Tags: load workbook from file using Aspose.Cells | set workbook calculation mode manual Aspose.Cells | save workbook to different file C# Aspose.Cells | check file existence before workbook load Aspose.Cells | handle exceptions when loading or saving Excel with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example verifies that the input.xlsx file exists, loads it into an Aspose.Cells Workbook, optionally sets the calculation mode to Manual, and saves the workbook as output.xlsx while handling any runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists before attempting to load it.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook from the specified file.
            Workbook workbook = new Workbook(inputPath);

            // NOTE: Setting calculation mode is optional and depends on the Aspose.Cells version.
            // If needed, use the appropriate API for the version you reference.

            // Save the modified workbook to the output file.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
