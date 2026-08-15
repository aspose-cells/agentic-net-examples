// Title: Copy a cell range to a new workbook and save as XLSX with Aspose.Cells for .NET
// Description: Loads source.xlsx, extracts the A1:C3 range from the first worksheet, copies it into a fresh workbook, and saves the result as copied_range.xlsx in XLSX format using Aspose.Cells for C#.
// Keywords: Aspose.Cells copy range .NET | C# copy Excel cells to new workbook | Aspose.Cells create workbook from range | save copied range as XLSX | Excel range copy Aspose.Cells
// Common Searches: Aspose.Cells copy range to new workbook C# | How to copy cells A1:C3 to another Excel file using Aspose | C# Aspose.Cells save selected range as separate XLSX | Copy Excel range between workbooks with Aspose.Cells | Create new Excel file from a range using Aspose.Cells .NET
// Developer Intent: Extract a defined cell range from an existing Excel file, place it into a new workbook, and write the new file in XLSX format.
// Use Cases: Distribute a specific report section without exposing the full master workbook. | Generate a lightweight template that contains only the data needed for downstream processing. | Create a shareable snapshot of a data block while preserving the original workbook unchanged.
// AI Prompts: Write C# code with Aspose.Cells that copies a runtime‑determined range from a source workbook to a new workbook and saves it as XLSX. | Explain how to copy multiple non‑contiguous ranges into separate worksheets of a new workbook while keeping formatting using Aspose.Cells. | Show how to preserve formulas, styles, and merged cells when copying a range to a new workbook with Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;

// Loads source.xlsx, extracts the A1:C3 range from the first worksheet, copies it into a fresh workbook, and saves the result as copied_range.xlsx in XLSX format using Aspose.Cells for C#.
class CopyRangeToNewWorkbook
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string destPath = "copied_range.xlsx";

            // Verify source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook from a file
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Access the first worksheet in the source workbook
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

            // Define the range to be copied (e.g., cells A1:C3)
            Aspose.Cells.Range sourceRange = sourceSheet.Cells.CreateRange("A1:C3");

            // Create a new (empty) destination workbook
            Workbook destinationWorkbook = new Workbook();

            // Access the first worksheet in the destination workbook
            Worksheet destinationSheet = destinationWorkbook.Worksheets[0];

            // Define a destination range with the same dimensions starting at A1
            Aspose.Cells.Range destinationRange = destinationSheet.Cells.CreateRange("A1:C3");

            // Copy the source range into the destination range
            destinationRange.Copy(sourceRange);

            // Save the destination workbook as an XLSX file
            destinationWorkbook.Save(destPath, SaveFormat.Xlsx);
            Console.WriteLine($"Range copied successfully to {destPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
