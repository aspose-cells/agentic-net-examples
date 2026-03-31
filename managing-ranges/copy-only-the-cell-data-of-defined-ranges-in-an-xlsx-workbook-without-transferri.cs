using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        // Load the source workbook that contains the data to be copied
        Workbook sourceWorkbook = new Workbook("source.xlsx");

        // Create a new empty workbook that will receive the copied data
        Workbook destinationWorkbook = new Workbook();

        // Get the first worksheet from each workbook (you can adjust indices as needed)
        Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
        Worksheet destinationSheet = destinationWorkbook.Worksheets[0];

        // Define the source range you want to copy (example: A1:C5)
        AsposeRange sourceRange = sourceSheet.Cells.CreateRange("A1:C5");

        // Define the destination range where the data will be placed (same size as source)
        AsposeRange destinationRange = destinationSheet.Cells.CreateRange("A1:C5");

        // Copy only the cell data (values and formulas) without any formatting
        destinationRange.CopyData(sourceRange);

        // Save the resulting workbook
        destinationWorkbook.Save("output.xlsx");
    }
}