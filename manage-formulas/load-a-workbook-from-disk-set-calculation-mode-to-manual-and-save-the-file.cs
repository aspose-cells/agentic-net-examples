// Title: Load an Excel workbook from a file, verify its existence, and save it to a new location using Aspose.Cells for .NET
// Common Searches: how to load an existing .xlsx file with Aspose.Cells in C# and save it to another path | C# Aspose.Cells example for checking file existence before opening workbook | saving a workbook after loading it with Aspose.Cells while handling exceptions | Aspose.Cells load workbook from disk and write to new file in .NET | C# code to open an Excel file with Aspose.Cells and write it out to a different filename
// Tags: load workbook from disk using Aspose.Cells C# | save workbook to different file with Aspose.Cells | verify input file existence before workbook load | handle workbook save exceptions Aspose.Cells | Aspose.Cells workbook I/O handling

using System;
using System.IO;
using Aspose.Cells;

// The example checks that input.xlsx exists, loads it into an Aspose.Cells Workbook, and saves it as output.xlsx, using try‑catch blocks to manage file‑not‑found and save errors.
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

            // Save the workbook to the desired output path
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display an error message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
