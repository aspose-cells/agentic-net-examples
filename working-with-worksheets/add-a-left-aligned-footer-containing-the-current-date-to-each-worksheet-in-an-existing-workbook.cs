// Title: Add a left-aligned footer showing the current short date to every worksheet in an existing Excel workbook using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells set left footer with today's date for all sheets | How to add a date footer to each worksheet in an existing workbook using Aspose.Cells | Programmatically apply a left-aligned footer containing the short date to every worksheet in .NET | Aspose.Cells page setup SetFooter left footer current date loop worksheets | Add current short date to left footer of all worksheets in Excel file with Aspose.Cells
// Tags: set left footer date Aspose.Cells C# | page setup SetFooter all worksheets | add date footer to existing Excel workbook | iterate worksheets Aspose.Cells footer | left-aligned footer short date .NET | modify workbook footer programmatically | Aspose.Cells footer customization

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing Excel file (or creates a new workbook), iterates through each worksheet, and uses PageSetup.SetFooter with index 0 to place the current short date as a left-aligned footer, then saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "InputWorkbook.xlsx";
            const string outputPath = "OutputWorkbook.xlsx";

            try
            {
                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a new one.
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook(); // creates a workbook with a default worksheet
                }

                // Set left-aligned footer to current short date for each worksheet.
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    PageSetup pageSetup = worksheet.PageSetup;
                    // Use SetFooter to avoid version‑specific property issues.
                    pageSetup.SetFooter(0, DateTime.Now.ToString("d")); // 0 = left footer
                }

                // Ensure the output directory exists.
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
