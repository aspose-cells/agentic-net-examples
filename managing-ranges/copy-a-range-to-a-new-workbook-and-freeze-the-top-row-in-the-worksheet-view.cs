// Title: Copy a Range to a New Workbook and Freeze the Top Row with Aspose.Cells for .NET (C#)
// Description: C# example that loads a source Excel file, copies a defined range (e.g., A1:C5) into a newly created workbook, applies FreezePanes to keep the first row visible, and saves the result while handling missing files and exceptions.
// Keywords: Aspose.Cells for .NET | C# copy Excel range | copy range between workbooks | FreezePanes Aspose.Cells | freeze top row C# | Excel range copy example | Aspose.Cells sample code | GitHub Aspose.Cells copy range
// Common Searches: Aspose.Cells copy range to another workbook C# | How to freeze the first row with Aspose.Cells | Copy Excel range and preserve formatting Aspose.Cells | Freeze panes programmatically in C# | Aspose.Cells example for copying and freezing rows
// Developer Intent: Copy a specific cell block from an existing workbook into a new workbook and lock the header row in place.
// Use Cases: Create a lightweight report that extracts a table from a master file while keeping column headers fixed. | Generate a template workbook containing only the needed data range for downstream processing. | Distribute a data snapshot to partners with the header row frozen for easier navigation.
// AI Prompts: Generate C# code using Aspose.Cells to copy a dynamic range (based on used cells) from one workbook to a new workbook and freeze the top row. | Explain how to copy a range while preserving styles, formulas, and merged cells with Aspose.Cells, then apply FreezePanes to keep the header visible. | Show error‑handling patterns for missing source files and saving the destination workbook with a timestamped filename.

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// C# example that loads a source Excel file, copies a defined range (e.g., A1:C5) into a newly created workbook, applies FreezePanes to keep the first row visible, and saves the result while handling missing files and exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string outputPath = "output.xlsx";

            // Verify source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file \"{sourcePath}\" not found.");
                return;
            }

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Define the source range to copy (example: A1:C5)
            AsposeRange sourceRange = sourceWorkbook.Worksheets[0].Cells.CreateRange("A1:C5");

            // Create a new workbook that will receive the copied range
            Workbook destinationWorkbook = new Workbook();

            // Define the destination range with the same size as the source range
            AsposeRange destinationRange = destinationWorkbook.Worksheets[0].Cells.CreateRange("A1:C5");

            // Copy the source range into the destination range
            sourceRange.Copy(destinationRange);

            // Freeze the top row in the destination worksheet view
            Worksheet destSheet = destinationWorkbook.Worksheets[0];
            // Freeze at row index 1 (second row), column index 0, freezing 1 row and 0 columns
            destSheet.FreezePanes(1, 0, 1, 0);

            // Save the new workbook
            destinationWorkbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
