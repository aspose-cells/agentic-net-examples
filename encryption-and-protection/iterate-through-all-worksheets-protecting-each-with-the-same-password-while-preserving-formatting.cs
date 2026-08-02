// Title: Protect All Worksheets with One Password Using Aspose.Cells for .NET (C#)
// Description: Load an existing Excel file (or create a new workbook), loop through each worksheet, apply full protection with a single password, keep all cell styles intact, and save the protected workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# worksheet protection | Excel password protection .NET | protect multiple sheets programmatically | preserve Excel formatting | sheet-level security Aspose | protect all worksheets C#
// Common Searches: Aspose.Cells protect all sheets C# | C# protect Excel worksheets with same password | apply sheet protection without losing formatting .NET | programmatically set password for every worksheet | protect workbook worksheets using Aspose.Cells
// Developer Intent: Apply the same password to every worksheet in a workbook while preserving all existing formatting and styles.
// Use Cases: Secure a template before distribution by locking every sheet with a single password. | Automate protection of generated reports to prevent edits while keeping the original layout. | Batch‑process multiple workbooks to enforce sheet‑level security without altering cell formatting.
// AI Prompts: Write C# code with Aspose.Cells that opens an Excel file, iterates over all worksheets, protects each one with a given password, and saves the file without changing any formatting. | Show an example of using Aspose.Cells for .NET to apply identical password protection to every sheet in a workbook while keeping all styles and formulas unchanged.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Load an existing Excel file (or create a new workbook), loop through each worksheet, apply full protection with a single password, keep all cell styles intact, and save the protected workbook using Aspose.Cells for .NET.
    public class ProtectAllWorksheets
    {
        public static void Run()
        {
            string inputPath = "input.xlsx";
            string outputPath = "output_protected.xlsx";
            string password = "myPassword123";

            try
            {
                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a new one
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook(); // creates a workbook with a default sheet
                }

                // Protect each worksheet with the specified password
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    sheet.Protect(ProtectionType.All, password, null);
                }

                // Save the protected workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ProtectAllWorksheets.Run();
        }
    }
}
