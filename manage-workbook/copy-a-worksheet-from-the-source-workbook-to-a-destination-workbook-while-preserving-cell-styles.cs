using System;
using Aspose.Cells;

class CopyWorksheetDemo
{
    static void Main()
    {
        // Load the source workbook from a file
        Workbook sourceWorkbook = new Workbook("source.xlsx");

        // Create a new (empty) destination workbook
        Workbook destWorkbook = new Workbook();

        // Get the source worksheet (e.g., the first sheet)
        Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

        // Get the destination worksheet (the first sheet in the new workbook)
        Worksheet destSheet = destWorkbook.Worksheets[0];

        // Copy the source worksheet into the destination worksheet.
        // This method copies both the cell values and their styles.
        destSheet.Copy(sourceSheet);

        // Save the destination workbook to a file
        destWorkbook.Save("destination.xlsx");
    }
}