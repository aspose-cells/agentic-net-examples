// Title: Convert an SXC workbook to CSV after renaming the active worksheet using Aspose.Cells for .NET
// Common Searches: C# code to load an .sxc file, rename its active sheet, and save as .csv with Aspose.Cells | How to change the name of the active worksheet before exporting to CSV in Aspose.Cells .NET | Aspose.Cells example for converting a StarOffice Calc (.sxc) workbook to CSV while renaming the sheet | Saving only the active worksheet to CSV from an SXC workbook using Aspose.Cells | Rename active worksheet programmatically and export to CSV using Aspose.Cells for .NET
// Tags: Aspose.Cells rename worksheet programmatically | convert SXC workbook to CSV C# | save only active sheet as CSV Aspose.Cells | auto-detect workbook format Aspose.Cells | set ActiveSheetIndex before saving Aspose.Cells

using System;
using Aspose.Cells;

namespace SxcToCsvExample
{
    // A C# console example that accepts an input .sxc file and an output .csv path, loads the workbook with Aspose.Cells (auto‑detecting the format), retrieves the currently active worksheet, renames it, ensures it remains the active sheet, and then saves only that sheet as a CSV file. Errors are caught and reported to the console.
    class Program
    {
        static void Main(string[] args)
        {
            // Expect two arguments: input SXC file path and output CSV file path
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: SxcToCsvExample <input.sxc> <output.csv>");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];

            try
            {
                // Load the SXC workbook. Aspose.Cells automatically detects the format.
                Workbook workbook = new Workbook(inputPath);

                // Get the active worksheet (the one currently selected in the workbook)
                int activeSheetIndex = workbook.Worksheets.ActiveSheetIndex;
                Worksheet activeSheet = workbook.Worksheets[activeSheetIndex];

                // Rename the active worksheet
                string newSheetName = "RenamedSheet";
                activeSheet.Name = newSheetName;
                Console.WriteLine($"Active worksheet renamed to \"{newSheetName}\".");

                // Ensure the renamed sheet is the active sheet before saving as CSV.
                // When saving to CSV, Aspose.Cells writes only the active sheet.
                workbook.Worksheets.ActiveSheetIndex = activeSheetIndex;

                // Export the active worksheet to CSV.
                // SaveFormat.Csv writes the active sheet only.
                workbook.Save(outputPath, SaveFormat.Csv);
                Console.WriteLine($"Worksheet \"{newSheetName}\" exported successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during processing:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
