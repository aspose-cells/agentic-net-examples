using Aspose.Cells;
using System;

class CopyWorksheetAndFreezeHeader
{
    static void Main()
    {
        // Load the source workbook from a file
        Workbook sourceWorkbook = new Workbook("Source.xlsx");

        // Create a new (empty) destination workbook
        Workbook destinationWorkbook = new Workbook();

        // Get the worksheet to be copied (first worksheet in the source workbook)
        Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

        // Add a copy of the source worksheet to the destination workbook
        // The AddCopy method copies the worksheet data and formats
        int copiedIndex = destinationWorkbook.Worksheets.AddCopy(sourceSheet.Index);
        Worksheet copiedSheet = destinationWorkbook.Worksheets[copiedIndex];

        // Freeze the header row (first row) in the copied worksheet
        // Freeze at row index 1 (second row) with 1 frozen row and 0 frozen columns
        copiedSheet.FreezePanes(1, 0, 1, 0);

        // Save the destination workbook with the copied and frozen worksheet
        destinationWorkbook.Save("Result.xlsx");
    }
}