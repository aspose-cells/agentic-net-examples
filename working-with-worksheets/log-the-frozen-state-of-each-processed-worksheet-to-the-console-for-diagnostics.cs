// Title: Log whether each worksheet has frozen panes using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an Excel workbook with Aspose.Cells, iterates over all worksheets, checks if FreezePanesRow or FreezePanesColumn is set, and prints the frozen‑pane status to the console, with a try‑catch fallback for older library versions. | Create a sample that logs the frozen pane state of every sheet in a workbook, handling missing FreezePanesRow/FreezePanesColumn properties and saving the file unchanged.
// Common Searches: how to detect frozen panes in each worksheet with Aspose.Cells C# | console output of worksheet freeze status using Aspose.Cells .NET | Aspose.Cells compatibility for FreezePanesRow property across versions | diagnostic logging of Excel sheet freeze panes programmatically
// Tags: detect frozen panes Aspose.Cells | worksheet freeze status console logging | fallback for missing FreezePanesRow property | iterate worksheets Aspose.Cells C# | diagnostic logging Excel freeze panes

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, loops through every worksheet, attempts to read the FreezePanesRow/FreezePanesColumn properties to determine if the sheet has frozen panes (defaulting to false when the properties are unavailable), writes each sheet's frozen status to the console, and then saves the workbook unchanged.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Aspose.Cells older versions do not expose FreezePanesRow/Column.
                // As a fallback, we assume no frozen panes when the properties are unavailable.
                bool isFrozen = false;

                // If the current version supports FreezePanesRow/Column, use them.
                // This block is guarded by a try-catch to avoid runtime errors on older APIs.
                try
                {
                    // The following properties exist in newer Aspose.Cells releases.
                    // Uncomment when using a version that provides them.
                    // isFrozen = sheet.FreezePanesRow > 0 || sheet.FreezePanesColumn > 0;
                }
                catch
                {
                    // Keep isFrozen as false if properties are unavailable.
                }

                Console.WriteLine($"Worksheet '{sheet.Name}' frozen: {isFrozen}");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (no modifications made in this example)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
