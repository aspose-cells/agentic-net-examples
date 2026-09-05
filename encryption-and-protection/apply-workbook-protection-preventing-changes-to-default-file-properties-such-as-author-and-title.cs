// Title: Use Aspose.Cells for .NET to password‑protect an Excel workbook and block editing of default file properties (author, title)
// AI Prompts: Write C# code that loads an existing .xlsx file, applies full workbook protection with a password using Aspose.Cells, and saves the result to a new file. | Demonstrate how to call Workbook.Protect with ProtectionType.All to prevent changes to workbook structure and built‑in document properties. | Add robust error handling that checks for the source file, catches exceptions, and logs meaningful messages when protecting a workbook with Aspose.Cells.
// Common Searches: aspnet protect Excel workbook metadata password Aspose.Cells | C# Aspose.Cells Workbook.Protect prevent editing author title | how to block changes to default file properties in .xlsx using Aspose.Cells | save protected Excel file with Aspose.Cells .NET example
// Tags: Aspose.Cells workbook protection password | prevent editing Excel file properties .NET | Workbook.Protect method example C# | protect default document properties Aspose.Cells | full workbook protection Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample checks that the source XLSX file exists, loads it into an Aspose.Cells Workbook, applies full protection with a password via Workbook.Protect(ProtectionType.All), saves the protected workbook to a new file, and handles any runtime exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output_protected.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Protect the workbook with a password (covers structure, windows, etc.)
                // This is the closest alternative to protecting default file properties
                workbook.Protect(ProtectionType.All, "StrongPassword123");

                // Save the protected workbook to a new file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions and display an error message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
