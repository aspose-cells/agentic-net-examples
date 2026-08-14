// Title: Set 120% Zoom on All Worksheets and Export Each as a Separate XLS with Aspose.Cells for .NET
// Description: C# example that loads a workbook, iterates through every worksheet, copies each sheet into its own workbook, applies a 120 % zoom (valid 10‑400), and saves the result as an Excel 97‑2003 (XLS) file named after the original sheet.
// Keywords: Aspose.Cells C# zoom worksheet | set worksheet zoom percentage | export each sheet to separate XLS | copy worksheet to new workbook Aspose.Cells | Aspose.Cells SaveFormat.Excel97To2003 | split workbook into individual files | apply 120% zoom Aspose.Cells | Excel 97-2003 export C# | worksheet Zoom property Aspose | batch worksheet processing Aspose.Cells
// Common Searches: how to set worksheet zoom with Aspose.Cells C# | export each worksheet as separate XLS using Aspose.Cells | Aspose.Cells copy single sheet to new workbook and set zoom | C# split Excel file into individual sheets with specific zoom | Aspose.Cells save workbook as Excel97To2003 format
// Developer Intent: Apply a 120 % zoom to every worksheet in a workbook and save each sheet as an individual XLS file.
// Use Cases: Create legacy XLS reports that open at a predefined zoom for consistent viewing. | Generate per‑sheet files for downstream systems that require a fixed zoom level. | Automate splitting a multi‑sheet workbook into separate files while preserving layout scaling.
// AI Prompts: Generate C# code that reads a zoom value from a JSON config and applies it to each worksheet before exporting to XLS. | Show how to add robust logging (including file system permissions and invalid sheet names) to the worksheet export loop. | Provide a version of the example that saves the sheets as .xlsx files with a 150 % zoom.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsZoomExample
{
    // C# example that loads a workbook, iterates through every worksheet, copies each sheet into its own workbook, applies a 120 % zoom (valid 10‑400), and saves the result as an Excel 97‑2003 (XLS) file named after the original sheet.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the source workbook
                string sourcePath = "source.xlsx";

                // Verify that the source file exists to avoid FileNotFoundException
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the source workbook
                Workbook sourceWorkbook = new Workbook(sourcePath);

                // Iterate through all worksheets in the source workbook
                for (int i = 0; i < sourceWorkbook.Worksheets.Count; i++)
                {
                    try
                    {
                        // Get the current worksheet
                        Worksheet srcSheet = sourceWorkbook.Worksheets[i];

                        // Create a new workbook that will contain only this worksheet
                        Workbook singleSheetWorkbook = new Workbook();

                        // The new workbook contains one default sheet; copy the source sheet into it
                        Worksheet destSheet = singleSheetWorkbook.Worksheets[0];
                        srcSheet.Copy(destSheet);

                        // Rename the copied sheet to match the source name
                        destSheet.Name = srcSheet.Name;

                        // Apply a 120% zoom level to the worksheet (valid range 10‑400)
                        destSheet.Zoom = 120;

                        // Build a file name for the individual worksheet (e.g., "Sheet1.xls")
                        string outputFileName = $"{srcSheet.Name}.xls";

                        // Save the workbook as an Excel 97‑2003 file (XLS format)
                        singleSheetWorkbook.Save(outputFileName, SaveFormat.Excel97To2003);

                        Console.WriteLine($"Saved worksheet '{srcSheet.Name}' to '{outputFileName}'.");
                    }
                    catch (Exception ex)
                    {
                        // Log errors for individual worksheets but continue processing others
                        Console.WriteLine($"Error processing worksheet index {i}: {ex.Message}");
                    }
                }

                Console.WriteLine("All worksheets have been saved with 120% zoom.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
