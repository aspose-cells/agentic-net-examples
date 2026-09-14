// Title: Load an Excel workbook from disk, verify the file exists, and save it as a new file with Aspose.Cells for .NET
// AI Prompts: Write C# code that uses Aspose.Cells to open a specified .xlsx file only after confirming the file is present, then saves the workbook to a different path with proper exception handling. | Show a .NET example that demonstrates safe loading and saving of an Excel workbook using Aspose.Cells, including checks for missing files and logging errors. | Provide a C# snippet that reads an existing workbook with Aspose.Cells, keeps the default calculation mode, and writes the workbook to a new filename while handling possible I/O exceptions.
// Common Searches: aspocells c# load workbook from existing file and save to another location | how to verify Excel file existence before opening with Aspose.Cells .NET | c# Aspose.Cells example for loading and saving workbook with try‑catch
// Tags: open workbook Aspose.Cells C# | save workbook to new file Aspose.Cells | file existence validation Aspose.Cells | error handling Aspose.Cells workbook | default calculation mode Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example checks whether the input.xlsx file exists, loads it into an Aspose.Cells Workbook, and saves the workbook as output.xlsx, handling any exceptions that may arise.
class Program
{
    static void Main()
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

            // Load the workbook from disk
            Workbook workbook = new Workbook(inputPath);

            // NOTE: Calculation mode setting removed because the property may not be available in some versions.
            // The workbook will use the default calculation mode.

            // Save the workbook to the specified output file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
