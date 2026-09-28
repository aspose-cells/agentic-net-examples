// Title: Cut a formula‑containing range and paste it to a new location while preserving references using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells to move the range B2:D5 to start at F2, keep all formulas intact, and then delete the original cells. | Write a C# method that creates a destination range matching the size of a source range with formulas, copies it with Aspose.Cells, and clears the source to emulate a cut operation.
// Common Searches: Aspose.Cells C# cut and paste cells with formulas preserving references | how to move a block of formula cells to another area in an Excel file using Aspose.Cells | copy range with formulas and clear original range Aspose.Cells .NET | preserve relative formula references when relocating cells with Aspose.Cells API | C# Aspose.Cells move range B2:D5 to F2 without breaking formulas
// Tags: cut range preserving formulas Aspose.Cells | copy range with formulas C# | move Excel block Aspose.Cells API | clear source after copy Aspose.Cells | create matching destination range Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// Alias to avoid conflict with System.Range introduced in C# 8.0
using AsposeRange = Aspose.Cells.Range;

// The example loads an Excel workbook, copies the formula‑filled range B2:D5 to a new location starting at F2 while preserving all formula references and styles, clears the original cells to simulate a cut, and saves the result as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Load an existing workbook if it exists; otherwise create a new one
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the source range that contains formulas (e.g., B2:D5)
            string sourceRangeAddress = "B2:D5";

            // Define the top‑left cell of the destination where the range will be moved
            string destinationCellAddress = "F2";

            // Create the source range object (using the alias to avoid ambiguity)
            AsposeRange sourceRange = sheet.Cells.CreateRange(sourceRangeAddress);

            // Convert destination cell address to row/column indices
            Cell destCell = sheet.Cells[destinationCellAddress];
            int destRow = destCell.Row;
            int destColumn = destCell.Column;

            // Create a destination range with the same size as the source range
            AsposeRange destRange = sheet.Cells.CreateRange(destRow, destColumn, sourceRange.RowCount, sourceRange.ColumnCount);

            // Copy the source range to the destination range (preserves formulas and styles)
            sourceRange.Copy(destRange);

            // Clear the original range to achieve a "cut" operation
            sourceRange.ClearContents();

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
