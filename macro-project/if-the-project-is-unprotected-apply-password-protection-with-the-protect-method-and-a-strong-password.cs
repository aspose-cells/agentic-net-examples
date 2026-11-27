// Title: Protect an Excel workbook with a strong password using Aspose.Cells Workbook.Protect in C#
// Common Searches: C# Aspose.Cells protect workbook with password only if not already protected | How to add full protection to an Excel file using Aspose.Cells Protect method | Apply strong password to Excel workbook programmatically with Aspose.Cells in .NET | Aspose.Cells Workbook.Protect example for protecting structure and windows | Save a password‑protected Excel file using Aspose.Cells C#
// Tags: Aspose.Cells Workbook.Protect password | C# protect Excel workbook structure | Aspose.Cells full workbook protection | Excel file password protection .NET | Aspose.Cells protect unprotected workbook

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The code loads an existing Excel file, applies full workbook protection (structure, windows, etc.) with a strong password via Workbook.Protect(ProtectionType.All, "Str0ngP@ssw0rd!2026"), and saves the protected file to a new location.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Protect the entire workbook (structure, windows, etc.) with a strong password
                // Aspose.Cells does not expose an IsProtected property for Workbook, so we protect unconditionally
                workbook.Protect(ProtectionType.All, "Str0ngP@ssw0rd!2026");

                // Save the protected workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
