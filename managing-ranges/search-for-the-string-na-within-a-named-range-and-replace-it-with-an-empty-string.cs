// Title: Aspose.Cells for .NET – Replace “#N/A” in a Named Range with Blank (C#)
// Description: C# example that loads an Excel workbook, accesses a named range (e.g., "MyRange"), scans each cell for the literal "#N/A" error string, replaces matching cells with an empty value, and saves the updated file. Includes input validation and exception handling.
// Keywords: Aspose.Cells | C# | replace #N/A | named range | Excel error value | clear cell value | Workbook.Load | Workbook.Save | GitHub sample | Aspose.Cells .NET
// Common Searches: Aspose.Cells replace #N/A in named range C# | clear #N/A error values in Excel using Aspose.Cells | iterate cells of a named range Aspose.Cells .NET | how to remove #N/A from specific range in Excel programmatically | Aspose.Cells example for cleaning error strings
// Developer Intent: Remove every "#N/A" string from a specified named range and save the cleaned workbook.
// Use Cases: Sanitize imported datasets that contain placeholder "#N/A" entries before analysis. | Prepare a report section defined by a named range for publishing by stripping error values. | Automate batch processing of workbooks to ensure no "#N/A" strings appear in defined ranges.
// AI Prompts: Generate C# code with Aspose.Cells that opens a workbook, finds the named range "MyRange", replaces all "#N/A" cells with blanks, and writes the result to a new file. | Explain best practices for iterating over an Aspose.Cells named range while updating cell values safely. | Provide robust error‑handling patterns for loading a workbook, accessing a named range, and modifying cell contents using Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // C# example that loads an Excel workbook, accesses a named range (e.g., "MyRange"), scans each cell for the literal "#N/A" error string, replaces matching cells with an empty value, and saves the updated file. Includes input validation and exception handling.
    public class ReplaceNaInNamedRange
    {
        public static void Run()
        {
            const string inputPath = "InputWorkbook.xlsx";
            const string outputPath = "OutputWorkbook.xlsx";
            const string namedRange = "MyRange";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Retrieve the named range
                Name rangeName = workbook.Worksheets.Names[namedRange];
                if (rangeName == null)
                {
                    Console.WriteLine($"Named range '{namedRange}' not found.");
                    return;
                }

                // Get the actual range object (use fully qualified type to avoid ambiguity)
                Aspose.Cells.Range range = rangeName.GetRange();

                // Iterate through each cell in the range and replace "#N/A" with empty string
                foreach (Cell cell in range)
                {
                    if (cell.StringValue == "#N/A")
                    {
                        cell.PutValue(string.Empty);
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ReplaceNaInNamedRange.Run();
        }
    }
}
