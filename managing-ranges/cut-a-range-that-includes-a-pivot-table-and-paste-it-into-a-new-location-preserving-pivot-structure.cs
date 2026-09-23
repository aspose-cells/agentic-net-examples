// Title: Cut a range that includes a pivot table and paste it to a new location while preserving the pivot structure using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that cuts cells A1:D20 (containing a pivot table) and pastes them starting at F1, keeping the pivot table functional. | Show how to use Aspose.Cells Range.Copy and Range.ClearContents to relocate a pivot‑table range without losing its definition in a .NET workbook.
// Common Searches: Aspose.Cells C# move pivot table range without breaking the pivot | how to cut and paste cells that contain a pivot table using Aspose.Cells for .NET | preserve pivot table when copying a range in an Excel file with Aspose.Cells | C# example for relocating a pivot table block with Aspose.Cells | copy range with pivot table and clear original cells Aspose.Cells
// Tags: cut range with pivot table Aspose.Cells | paste range preserving pivot Aspose.Cells | Aspose.Cells move pivot table cells | C# copy range containing pivot table | Aspose.Cells range cut and paste example | preserve pivot definition Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads a workbook, defines a source range (A1:D20) that contains a pivot table, creates a destination range starting at F1 with the same size, copies the source to the destination, clears the original cells to simulate a cut, and saves the modified workbook as output.xlsx.
class PivotCutPasteExample
{
    static void Main()
    {
        try
        {
            // Input workbook path
            string inputPath = "input.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the source range address (e.g., A1:D20)
            string sourceAddress = "A1:D20";
            string[] srcParts = sourceAddress.Split(':');
            CellArea sourceArea = CellArea.CreateCellArea(srcParts[0], srcParts[1]);
            int sourceRows = sourceArea.EndRow - sourceArea.StartRow + 1;
            int sourceCols = sourceArea.EndColumn - sourceArea.StartColumn + 1;

            // Define the destination top‑left cell address (e.g., F1)
            string destinationAddress = "F1";
            CellArea destArea = CellArea.CreateCellArea(destinationAddress, destinationAddress);

            // Create source and destination ranges (use fully qualified Aspose.Cells.Range)
            Aspose.Cells.Range srcRange = sheet.Cells.CreateRange(sourceArea.StartRow, sourceArea.StartColumn, sourceRows, sourceCols);
            Aspose.Cells.Range destRange = sheet.Cells.CreateRange(destArea.StartRow, destArea.StartColumn, sourceRows, sourceCols);

            // Copy source range to destination
            srcRange.Copy(destRange);

            // Clear the original source range (simulate cut)
            srcRange.ClearContents();

            // Save the modified workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
