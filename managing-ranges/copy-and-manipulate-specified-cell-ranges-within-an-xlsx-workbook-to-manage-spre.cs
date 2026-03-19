using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeManipulation
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet
            Worksheet sourceSheet = workbook.Worksheets[0];
            Cells sourceCells = sourceSheet.Cells;

            // Define the source range A1:C5 (rows 0-4, columns 0-2)
            AsposeRange sourceRange = sourceCells.CreateRange("A1", "C5");

            // Define the destination range E1:G5 on the same worksheet
            AsposeRange destinationRange = sourceCells.CreateRange("E1", "G5");

            // Copy only the cell values (including formulas) from source to destination
            destinationRange.CopyData(sourceRange);

            // Clear the original source range (remove both contents and formatting)
            // ClearRange overload uses start row, start column, total rows, total columns
            sourceCells.ClearRange(0, 0, 5, 3); // A1:C5

            // Add a new worksheet to the workbook
            int newSheetIndex = workbook.Worksheets.Add();
            Worksheet newSheet = workbook.Worksheets[newSheetIndex];
            Cells newSheetCells = newSheet.Cells;

            // Create a range on the new sheet where the data will be copied
            AsposeRange newSheetRange = newSheetCells.CreateRange("A1", "C5");

            // Copy the same source data into the new worksheet
            newSheetRange.CopyData(sourceRange);

            // Save the modified workbook to a new file
            workbook.Save("output.xlsx");
        }
    }
}