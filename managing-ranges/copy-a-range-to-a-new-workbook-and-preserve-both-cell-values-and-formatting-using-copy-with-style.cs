// Title: Copy an Excel range with values and formatting to a new workbook using Aspose.Cells for .NET (C#)
// Description: Loads a source workbook, creates an empty destination workbook, defines matching ranges (e.g., A1:C5), copies both cell contents and style information with the Range.Copy method, and saves the result as a new file. Includes basic file‑existence checking and exception handling.
// Keywords: Aspose.Cells C# copy range | preserve Excel cell formatting | Range.Copy example | copy range to new workbook | .NET Excel style transfer | Aspose.Cells tutorial
// Common Searches: Aspose.Cells copy range with formatting | C# copy Excel cells to another workbook preserving styles | Range.Copy method Aspose.Cells .NET | how to duplicate a styled table in a new Excel file
// Developer Intent: Transfer a defined cell block from an existing workbook to a fresh workbook while retaining all formatting and values.
// Use Cases: Create a report workbook by reusing a styled template section. | Export a formatted data table from a master file for client distribution. | Clone a chart data range into a separate workbook for independent analysis.
// AI Prompts: Show C# code that copies a range with its styles to a new workbook using Aspose.Cells, including a check for a missing source file. | Provide an Aspose.Cells example that copies a range and then auto‑fits the destination columns to match the source widths. | Explain how to copy multiple non‑contiguous ranges with formatting into a new workbook in C#.

using System;
using System.IO;
using Aspose.Cells;

// Loads a source workbook, creates an empty destination workbook, defines matching ranges (e.g., A1:C5), copies both cell contents and style information with the Range.Copy method, and saves the result as a new file. Includes basic file‑existence checking and exception handling.
class CopyRangeWithStyleDemo
{
    static void Main()
    {
        try
        {
            string srcPath = "source.xlsx";
            string destPath = "copied.xlsx";

            // Verify source file exists
            if (!File.Exists(srcPath))
            {
                Console.WriteLine($"Source file '{srcPath}' not found.");
                return;
            }

            // Load the source workbook
            Workbook srcWorkbook = new Workbook(srcPath);

            // Create a new (empty) destination workbook
            Workbook destWorkbook = new Workbook();

            // Define the source range
            Worksheet srcSheet = srcWorkbook.Worksheets[0];
            Aspose.Cells.Range srcRange = srcSheet.Cells.CreateRange("A1:C5");

            // Define the destination range in the new workbook
            Worksheet destSheet = destWorkbook.Worksheets[0];
            Aspose.Cells.Range destRange = destSheet.Cells.CreateRange("A1:C5");

            // Copy both values and formatting from source to destination
            srcRange.Copy(destRange);

            // Save the destination workbook
            destWorkbook.Save(destPath);
            Console.WriteLine($"Range copied successfully to '{destPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
