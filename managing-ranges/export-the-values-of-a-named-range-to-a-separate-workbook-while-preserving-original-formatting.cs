// Title: Export a Named Range to a New Workbook with Full Formatting – Aspose.Cells for .NET (C#)
// Description: Loads a source workbook, locates a named range, parses its address, creates an identical range in a fresh workbook, copies all cell data, formulas and styles using PasteOptions.All, and saves the result as a separate file.
// Keywords: Aspose.Cells export named range | copy range to new workbook | preserve cell formatting C# | PasteOptions.All | .NET spreadsheet extraction | named range to separate file | Aspose.Cells range copy example
// Common Searches: Aspose.Cells copy named range to another workbook | export MyRange from Source.xlsx preserving styles | C# Aspose.Cells extract range and save as new file | how to keep formulas when copying a range with Aspose.Cells | export named range to separate workbook .NET
// Developer Intent: Create a new workbook that contains only the cells of a specified named range, keeping every value, formula and style intact.
// Use Cases: Generate a lightweight report that includes only a predefined data block from a master workbook. | Distribute a template containing a specific chart data range to external partners. | Automate extraction of user‑selected data for downstream processing while preserving formulas and visual formatting.
// AI Prompts: Write C# code with Aspose.Cells to export a single named range to a new workbook, retaining all formatting and formulas. | Show how to modify the sample to export multiple named ranges, each to its own worksheet in the destination workbook. | Explain handling of named ranges that reference external sheets or use absolute references when copying with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Loads a source workbook, locates a named range, parses its address, creates an identical range in a fresh workbook, copies all cell data, formulas and styles using PasteOptions.All, and saves the result as a separate file.
class ExportNamedRange
{
    static void Main()
    {
        try
        {
            const string sourcePath = "Source.xlsx";
            const string destinationPath = "ExportedRange.xlsx";
            const string rangeName = "MyRange";

            // Verify source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file \"{sourcePath}\" not found.");
                return;
            }

            // Load the workbook that contains the named range
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Retrieve the named range definition
            Name namedRange = sourceWorkbook.Worksheets.Names[rangeName];
            if (namedRange == null)
            {
                Console.WriteLine($"Named range \"{rangeName}\" not found.");
                return;
            }

            // Parse RefersTo to obtain sheet name and address
            string refersTo = namedRange.RefersTo.TrimStart('=');
            string sheetName;
            string address;

            if (refersTo.Contains("!"))
            {
                // Format: SheetName!A1:B2 (sheet name may be quoted)
                var parts = refersTo.Split('!');
                sheetName = parts[0].Trim('\'');
                address = parts[1];
            }
            else
            {
                // No sheet specified – use the first worksheet
                sheetName = sourceWorkbook.Worksheets[0].Name;
                address = refersTo;
            }

            // Get the worksheet where the named range resides
            Worksheet sourceSheet = sourceWorkbook.Worksheets[sheetName];
            if (sourceSheet == null)
            {
                Console.WriteLine($"Worksheet \"{sheetName}\" not found.");
                return;
            }

            // Create a Range object for the source named range
            Aspose.Cells.Range sourceRange = sourceSheet.Cells.CreateRange(address);

            // Prepare a new workbook for the exported range
            Workbook destinationWorkbook = new Workbook();
            Worksheet destSheet = destinationWorkbook.Worksheets[0];

            // Create a destination range with the same dimensions as the source range
            Aspose.Cells.Range destRange = destSheet.Cells.CreateRange(
                sourceRange.FirstRow,
                sourceRange.FirstColumn,
                sourceRange.RowCount,
                sourceRange.ColumnCount);

            // Copy the source range to the destination range, preserving all content and formatting
            PasteOptions pasteOptions = new PasteOptions
            {
                PasteType = PasteType.All
            };
            destRange.Copy(sourceRange, pasteOptions);

            // Save the result
            destinationWorkbook.Save(destinationPath);
            Console.WriteLine("Named range exported successfully with original formatting.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
