// Title: Apply the same password to protect every worksheet in an Excel file using Aspose.Cells for .NET without altering formatting
// AI Prompts: Write C# code that opens an existing .xlsx file with Aspose.Cells, enumerates all worksheets, and calls Worksheet.Protect with ProtectionType.All and a supplied password while leaving cell styles unchanged. | Show how to verify the source workbook exists, apply uniform sheet protection across the workbook, and then save the result as an .xlsx file using Aspose.Cells SaveFormat. | Create a robust try‑catch block for loading, protecting, and saving a workbook, and log any exceptions that occur during the sheet‑protection process in C#.
// Common Searches: Aspose.Cells protect each worksheet with the same password in C# | How to loop through all sheets and set full protection using Aspose.Cells .NET | Preserve cell formatting while applying sheet protection in Aspose.Cells | Save a password‑protected Excel workbook as XLSX with Aspose.Cells | C# example for bulk worksheet protection using Aspose.Cells
// Tags: apply sheet password Aspose.Cells | bulk worksheet protection .NET | preserve cell styles during protection Aspose.Cells | save password‑protected workbook as xlsx Aspose.Cells | enumerate worksheets Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // // Loads an existing Excel workbook, enumerates each worksheet, applies full protection with a single password while keeping all formatting intact, and saves the workbook as an XLSX file using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Replace placeholders with actual values or pass them via args
            string inputPath = "{InputFilePath}";
            string outputPath = "{OutputFilePath}";
            string password = "{Password}";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                var workbook = new Workbook(inputPath);

                // Protect each worksheet with the specified password
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // oldPassword is null because the sheet is not previously protected
                    sheet.Protect(ProtectionType.All, password, null);
                }

                // Save the modified workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
