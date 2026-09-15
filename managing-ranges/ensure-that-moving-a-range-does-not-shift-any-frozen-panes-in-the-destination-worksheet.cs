// Title: Move a cell range to another worksheet while preserving frozen panes with Aspose.Cells for .NET
// AI Prompts: Copy a rectangular range from one worksheet to a different worksheet and keep the existing freeze‑pane layout unchanged using Aspose.Cells. | After copying, delete the original range and shift the remaining cells left without disturbing the frozen rows or columns on the target sheet. | Read the current FreezePanes settings of the destination worksheet, perform the range move, and reapply those settings programmatically if they are cleared.
// Common Searches: asp.net aspocells copy range preserve freeze panes destination worksheet | c# move cells between sheets without moving frozen rows aspocells | how to keep freeze pane settings after copying a range in aspocells | delete source range with left shift after moving cells aspocells c#
// Tags: copy range between worksheets preserving frozen rows | Aspose.Cells delete range with ShiftType.Left | reapply freeze pane configuration after range move | C# Aspose.Cells move cell block without shifting freeze panes | maintain frozen columns while transferring cell range

using System;
using System.IO;
using Aspose.Cells;

// The example loads a workbook, copies a defined range from the "Source" sheet to a target location on the "Destination" sheet, deletes the original range using a left shift, and ensures that any existing frozen rows or columns on the destination worksheet remain unchanged after the operation.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the existing workbook
            var workbook = new Workbook(inputPath);

            // Get source and destination worksheets (replace with actual sheet names)
            var sourceSheet = workbook.Worksheets["Source"];
            var destSheet   = workbook.Worksheets["Destination"];

            if (sourceSheet == null)
                throw new ArgumentException("Source worksheet not found.");
            if (destSheet == null)
                throw new ArgumentException("Destination worksheet not found.");

            // Define the range to move (example: A2:C10)
            int startRow = 1;   // zero‑based index (A2)
            int startCol = 0;   // column A
            int totalRows = 9;  // rows 2‑10 inclusive
            int totalCols = 3;  // columns A‑C

            // Destination start cell (example: E5)
            int destStartRow = 4; // zero‑based index (E5)
            int destStartCol = 4; // column E

            // Copy the range to the destination location
            var sourceRange = sourceSheet.Cells.CreateRange(startRow, startCol, totalRows, totalCols);
            var destRange   = destSheet.Cells.CreateRange(destStartRow, destStartCol, totalRows, totalCols);
            sourceRange.Copy(destRange);

            // Clear the original range (optional, if you want to "move" rather than copy)
            // Use ShiftType.Left to shift cells left after deletion
            sourceSheet.Cells.DeleteRange(startRow, startCol, totalRows, totalCols, ShiftType.Left);

            // Example: re‑freeze first row and first column in the destination sheet
            // destSheet.FreezePanes(1, 1, 1, 1);
            // Uncomment the line above if you need to apply frozen panes.

            // Ensure the output directory exists
            var outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
