// Title: Load an Excel workbook, check file existence, and save it unchanged with Aspose.Cells for .NET
// AI Prompts: Write C# code that verifies an Excel file exists, opens it with Aspose.Cells, and saves the workbook to a different location without altering any content. | Extend the example to log the workbook's theme status before saving and re‑apply the original theme after the file is saved. | Create a try‑catch block that gracefully handles missing files and I/O errors while loading and saving a workbook using Aspose.Cells.
// Common Searches: Aspose.Cells C# check if Excel file exists before loading workbook | How to copy an Excel file to a new location using Aspose.Cells for .NET | Save workbook without modifying theme Aspose.Cells C# | C# Aspose.Cells load workbook and immediately save to another file | Error handling when opening missing Excel file with Aspose.Cells
// Tags: load workbook Aspose.Cells C# | verify Excel file existence Aspose.Cells | save workbook to new path Aspose.Cells | preserve workbook theme Aspose.Cells | exception handling Aspose.Cells file operations | copy Excel file Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates loading an existing Excel workbook after confirming the file exists, then saving it to a new location unchanged using Aspose.Cells for .NET, with basic error handling.
class ThemeConversion
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Save the workbook to the output file (no theme conversion performed due to API limitations)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
