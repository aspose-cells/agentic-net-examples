// Title: Password‑protect Excel workbook structure and save as default XLSX with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an existing .xlsx file, applies structure protection with a password using Aspose.Cells, and saves the workbook in the default XLSX format. | Show how to implement robust error handling for missing input files and protection failures when using Aspose.Cells to protect and save an Excel workbook.
// Common Searches: how to protect Excel workbook structure with a password using Aspose.Cells C# | save a password‑protected workbook as XLSX default format Aspose.Cells .NET | C# check if Excel file exists before loading with Aspose.Cells | Aspose.Cells exception handling for workbook protection errors
// Tags: structure protection with password Aspose.Cells | save workbook as default XLSX Aspose.Cells | file existence validation Aspose.Cells C# | exception handling for workbook protection .NET | protect Excel file programmatically Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example checks that the source Excel file exists, loads it into an Aspose.Cells Workbook, applies structure protection using a password, and then saves the protected workbook in the default XLSX format, with comprehensive try‑catch blocks to handle missing files, protection failures, and other runtime errors.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputFile = "input.xlsx";
                const string outputFile = "protected_output.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Error: The file \"{inputFile}\" was not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputFile);

                // Protect the workbook structure with a password
                try
                {
                    workbook.Protect(ProtectionType.Structure, "MyPassword");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Unable to protect workbook. {ex.Message}");
                }

                // Save the protected workbook in XLSX format
                workbook.Save(outputFile, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to \"{outputFile}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
