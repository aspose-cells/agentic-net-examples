// Title: Export a Named Range to a New Workbook with Full Formatting – Aspose.Cells for .NET (C#)
// Description: Loads a source workbook, extracts a named range (e.g., "MyRange") by parsing its RefersTo formula, creates a matching range in a fresh workbook, and copies the data using PasteOptions.PasteType.All to retain values, formulas, styles, and number formats before saving the result.
// Keywords: Aspose.Cells export named range | copy range to new workbook C# | preserve cell formatting Aspose | named range RefersTo parsing | PasteOptions All Aspose.Cells | C# Excel range extraction | Aspose.Cells named range example
// Common Searches: how to export a named range with Aspose.Cells | copy Excel range to another file preserving formatting C# | Aspose.Cells get RefersTo address of a named range | create new workbook from a named range Aspose | PasteOptions.All copy all cell attributes Aspose.Cells
// Developer Intent: Extract a specific named range from an existing workbook and save it as a separate Excel file while keeping all formatting and formulas intact.
// Use Cases: Generate a lightweight report that contains only the data defined by a named range, preserving its visual style. | Provide users with a downloadable slice of a larger spreadsheet (e.g., a financial model segment) without exposing the full workbook. | Automate template creation where a predefined named range is exported for downstream processing or archiving.
// AI Prompts: Write C# code using Aspose.Cells to copy a named range called 'MyRange' from one workbook to a new workbook, preserving formulas, styles, and number formats. | Explain how to parse the RefersTo property of an Aspose.Cells Name object to obtain the worksheet name and cell address for range extraction. | Show step‑by‑step how to configure PasteOptions in Aspose.Cells to copy all cell attributes when moving a range between workbooks.

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace ExportNamedRange
{
    // Loads a source workbook, extracts a named range (e.g., "MyRange") by parsing its RefersTo formula, creates a matching range in a fresh workbook, and copies the data using PasteOptions.PasteType.All to retain values, formulas, styles, and number formats before saving the result.
    class Program
    {
        static void Main()
        {
            const string sourcePath = "SourceWorkbook.xlsx";
            const string destPath = "ExportedRange.xlsx";

            // Verify source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file '{sourcePath}' not found.");
                return;
            }

            try
            {
                // Load the source workbook that contains the named range
                Workbook sourceWorkbook = new Workbook(sourcePath);

                // Retrieve the named range object (assumes the name is "MyRange")
                Name namedRange = sourceWorkbook.Worksheets.Names["MyRange"];
                if (namedRange == null)
                {
                    Console.WriteLine("Named range 'MyRange' not found.");
                    return;
                }

                // Create a Range instance from the RefersTo formula of the named range
                // RefersTo is like "=Sheet1!$A$1:$C$5"
                string refersTo = namedRange.RefersTo ?? string.Empty;
                if (refersTo.StartsWith("="))
                    refersTo = refersTo.Substring(1);

                // Split sheet name and address
                int exclPos = refersTo.IndexOf('!');
                if (exclPos < 0)
                {
                    Console.WriteLine("Invalid RefersTo format.");
                    return;
                }

                string sheetName = refersTo.Substring(0, exclPos);
                string address = refersTo.Substring(exclPos + 1);

                // Get the worksheet that holds the range
                Worksheet sourceSheet = sourceWorkbook.Worksheets[sheetName];
                AsposeRange sourceRange = sourceSheet.Cells.CreateRange(address);

                // Create a new workbook for the export
                Workbook destWorkbook = new Workbook();
                Worksheet destSheet = destWorkbook.Worksheets[0];

                // Create a destination range with the same dimensions as the source range
                AsposeRange destRange = destSheet.Cells.CreateRange(0, 0, sourceRange.RowCount, sourceRange.ColumnCount);

                // Set paste options to copy everything (values, formulas, formatting, etc.)
                PasteOptions pasteOptions = new PasteOptions
                {
                    PasteType = PasteType.All
                };

                // Copy the source range to the destination range preserving formatting
                destRange.Copy(sourceRange, pasteOptions);

                // Save the new workbook containing only the exported range
                destWorkbook.Save(destPath);

                Console.WriteLine("Named range exported successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
