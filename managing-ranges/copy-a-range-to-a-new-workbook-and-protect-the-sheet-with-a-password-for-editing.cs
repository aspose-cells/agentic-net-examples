// Title: Copy a Cell Range to a New Workbook and Password‑Protect the Sheet with Aspose.Cells (C#)
// Description: Loads a source workbook, copies a defined range (e.g., A1:C5) into a freshly created workbook, applies full sheet protection with a password, and saves the result. Includes file‑existence check and exception handling for robust automation.
// Keywords: Aspose.Cells copy range C# | new workbook from range | sheet protection password Aspose.Cells | ProtectionType.All example | Aspose.Cells range copy tutorial | C# Excel workbook automation | protect worksheet programmatically | Aspose.Cells file not found handling
// Common Searches: Aspose.Cells copy range to another workbook C# | How to protect an Aspose.Cells worksheet with a password | C# example for copying cells and setting sheet protection | Aspose.Cells create new workbook from selected cells | Password‑protect Excel sheet using Aspose.Cells .NET
// Developer Intent: Duplicate a specific cell block from an existing Excel file into a new file and lock the new sheet with a password to prevent edits.
// Use Cases: Extract a summary table from a master workbook and distribute it as a read‑only report. | Create a template file that contains only the required data range while restricting user modifications. | Automate data sharing workflows where confidential sections are omitted and the remaining sheet is secured.
// AI Prompts: Generate C# code that copies a range from one workbook to a new workbook and applies password protection using Aspose.Cells. | Show how to copy multiple non‑contiguous ranges into a new workbook and set different protection types with Aspose.Cells for .NET. | Explain how to verify sheet protection after saving a workbook with Aspose.Cells and how to change the password programmatically.

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeCopyAndProtect
{
    // Loads a source workbook, copies a defined range (e.g., A1:C5) into a freshly created workbook, applies full sheet protection with a password, and saves the result. Includes file‑existence check and exception handling for robust automation.
    class Program
    {
        static void Main()
        {
            try
            {
                const string sourcePath = "source.xlsx";
                const string outputPath = "output.xlsx";

                // Verify source file exists to avoid FileNotFoundException
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the source workbook
                using (Workbook sourceWorkbook = new Workbook(sourcePath))
                {
                    Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

                    // Define the source range to copy (e.g., A1:C5)
                    AsposeRange sourceRange = sourceSheet.Cells.CreateRange("A1:C5");

                    // Create a new workbook that will receive the copied range
                    using (Workbook destWorkbook = new Workbook())
                    {
                        Worksheet destSheet = destWorkbook.Worksheets[0];

                        // Define the destination range (same size as source) starting at A1
                        AsposeRange destRange = destSheet.Cells.CreateRange("A1:C5");

                        // Copy the source range into the destination range
                        sourceRange.Copy(destRange);

                        // Protect the destination worksheet with a password for editing
                        destSheet.Protect(ProtectionType.All, "EditPassword123", null);

                        // Save the new workbook
                        destWorkbook.Save(outputPath);
                    }
                }

                Console.WriteLine($"Workbook copied and saved to {outputPath}");
            }
            catch (Exception ex)
            {
                // Log or display the exception details for troubleshooting
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
