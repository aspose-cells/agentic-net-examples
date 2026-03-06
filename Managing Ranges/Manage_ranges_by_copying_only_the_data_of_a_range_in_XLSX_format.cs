using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeCopyDemo
{
    class Program
    {
        static void Main()
        {
            // Load the source workbook (replace with your actual file path)
            Workbook sourceWorkbook = new Workbook("source.xlsx");

            // Create a new workbook that will hold the copied data
            Workbook destinationWorkbook = new Workbook();

            // Get the first worksheet from each workbook
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            Worksheet destinationSheet = destinationWorkbook.Worksheets[0];

            // Define the source range (e.g., A1:B5)
            AsposeRange sourceRange = sourceSheet.Cells.CreateRange("A1:B5");

            // Define the destination range where only the values will be copied (e.g., C1:D5)
            AsposeRange destinationRange = destinationSheet.Cells.CreateRange("C1:D5");

            // Copy only the cell values from the source range to the destination range
            destinationRange.CopyValue(sourceRange);

            // Save the result as an XLSX file
            destinationWorkbook.Save("output.xlsx");
        }
    }
}