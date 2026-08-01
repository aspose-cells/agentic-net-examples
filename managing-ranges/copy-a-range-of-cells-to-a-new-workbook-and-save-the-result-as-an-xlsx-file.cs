// Title: Copy a Cell Range to a New Workbook and Save as XLSX with Aspose.Cells for .NET (C#)
// Description: C# example that loads or creates a source.xlsx file, defines the range A1:C5, creates an empty destination workbook, copies the range preserving values and formatting, and saves the result as copied_range.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells copy range C# | copy Excel cells to new workbook | Aspose.Cells create range programmatically | save copied range as XLSX | C# Aspose.Cells example | Excel range duplication .NET
// Common Searches: how to copy a range of cells to a new Excel file using Aspose.Cells | Aspose.Cells C# copy A1:C5 to another workbook | create new workbook from selected cells Aspose.Cells | copy Excel range preserving formatting Aspose.Cells .NET | Aspose.Cells copy range and save as xlsx
// Developer Intent: Duplicate a specific cell block from an existing workbook into a fresh workbook and write the result to an XLSX file using Aspose.Cells for .NET.
// Use Cases: Extract a report section from a template and distribute it as an independent file. | Generate a lightweight workbook containing only the data needed for downstream processing. | Create a copy of chart source data for sharing without exposing the full original workbook.
// AI Prompts: Write C# code with Aspose.Cells that copies a runtime‑determined range from one workbook to a new workbook and saves it as XLSX. | Explain how to copy a range while keeping cell formatting, formulas, and comments using Aspose.Cells for .NET. | Show error handling for a missing source file and how to create a placeholder workbook before copying the range.

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// C# example that loads or creates a source.xlsx file, defines the range A1:C5, creates an empty destination workbook, copies the range preserving values and formatting, and saves the result as copied_range.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source workbook
            string sourcePath = "source.xlsx";

            // Ensure the source file exists; create a simple workbook if it does not
            if (!File.Exists(sourcePath))
            {
                var tempWb = new Workbook();
                var tempSheet = tempWb.Worksheets[0];
                var tempCells = tempSheet.Cells;
                // Fill A1:C5 with sample data
                for (int row = 0; row < 5; row++)
                {
                    for (int col = 0; col < 3; col++)
                    {
                        tempCells[row, col].PutValue($"R{row + 1}C{col + 1}");
                    }
                }
                tempWb.Save(sourcePath);
            }

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            Cells sourceCells = sourceSheet.Cells;

            // Define the source range (A1:C5)
            AsposeRange sourceRange = sourceCells.CreateRange("A1:C5");

            // Create a new (empty) destination workbook
            Workbook destinationWorkbook = new Workbook();
            Worksheet destinationSheet = destinationWorkbook.Worksheets[0];
            Cells destinationCells = destinationSheet.Cells;

            // Create a destination range with the same size as the source range, starting at A1
            AsposeRange destinationRange = destinationCells.CreateRange(
                0,                     // first row (0‑based)
                0,                     // first column (0‑based)
                sourceRange.RowCount, // number of rows
                sourceRange.ColumnCount // number of columns
            );

            // Copy the source range into the destination range
            sourceRange.Copy(destinationRange);

            // Save the destination workbook
            string destPath = "copied_range.xlsx";
            destinationWorkbook.Save(destPath);
            Console.WriteLine($"Range copied successfully to '{destPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
