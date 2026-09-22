// Title: C# example: read cells from a worksheet's defined print area using Aspose.Cells LightCells
// AI Prompts: Provide C# sample code that opens an Excel file with Aspose.Cells, determines the PageSetup.PrintArea of the first worksheet, and uses the LightCells API to iterate through each cell in that area, outputting the address and value. | Create a reusable C# function that loads a workbook, identifies the defined print area (or falls back to the used range if none), and efficiently reads cell data with LightCells, including handling missing files and cell read exceptions.
// Common Searches: Aspose.Cells C# read only cells inside the defined print area | How to use LightCells to process a print area in an Excel workbook | C# extract cell values from worksheet print area with Aspose.Cells | Get cell addresses from a print area range using Aspose.Cells LightCells | Fallback to used range when print area is not set in Aspose.Cells
// Tags: Aspose.Cells LightCells read print area | C# extract worksheet print area cells | Aspose.Cells load workbook defined print area | Excel print area processing with LightCells | Efficient cell iteration Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The sample loads an Excel workbook, accesses the first worksheet, retrieves its PageSetup.PrintArea (or the used range if no print area is defined), calculates the boundaries, and uses Aspose.Cells LightCells to efficiently iterate over each cell in that region, printing the cell address and value while handling file‑not‑found and cell‑read errors.
class ExtractPrintArea
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            string workbookPath = "input.xlsx";

            // Ensure the file exists to avoid FileNotFoundException
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: File '{workbookPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the defined print area (e.g., "A1:C10")
            string printAreaRef = sheet.PageSetup.PrintArea;

            // Determine the area to process
            int startRow, startColumn, endRow, endColumn;

            if (string.IsNullOrEmpty(printAreaRef))
            {
                // No print area defined – use the worksheet's used range
                AsposeRange usedRange = sheet.Cells.MaxDisplayRange;
                startRow = usedRange.FirstRow;
                startColumn = usedRange.FirstColumn;
                endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                endColumn = usedRange.FirstColumn + usedRange.ColumnCount - 1;
            }
            else
            {
                // Use the print area reference to create a range and derive boundaries
                AsposeRange printRange = sheet.Cells.CreateRange(printAreaRef);
                startRow = printRange.FirstRow;
                startColumn = printRange.FirstColumn;
                endRow = printRange.FirstRow + printRange.RowCount - 1;
                endColumn = printRange.FirstColumn + printRange.ColumnCount - 1;
            }

            // Iterate through the cells within the determined area
            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startColumn; col <= endColumn; col++)
                {
                    try
                    {
                        Cell cell = sheet.Cells[row, col];
                        // Output cell address (e.g., A1) and its value
                        Console.WriteLine($"{cell.Name} = {cell.Value}");
                    }
                    catch (Exception cellEx)
                    {
                        Console.WriteLine($"Error reading cell R{row + 1}C{col + 1}: {cellEx.Message}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
