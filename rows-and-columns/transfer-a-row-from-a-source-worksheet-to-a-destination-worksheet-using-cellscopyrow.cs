// Title: Copy a row from one workbook to another using Aspose.Cells Cells.CopyRow in C#
// AI Prompts: Copy row 3 from source.xlsx into row 6 of destination.xlsx with Aspose.Cells Cells.CopyRow in C#. | Transfer an entire row between worksheets of different workbooks using the Cells.CopyRow method in C#. | Duplicate a specific Excel row from a source workbook to a target workbook programmatically with Aspose.Cells.
// Common Searches: asp.net core copy row from source.xlsx to destination.xlsx using Aspose.Cells | C# Aspose.Cells example for copying rows across workbooks | how to use Cells.CopyRow to move a row between Excel files in C# | copy specific Excel row to another workbook Aspose.Cells tutorial
// Tags: copy row using Cells.CopyRow C# | transfer Excel row between workbooks Aspose.Cells | copy row across worksheets Aspose.Cells | Aspose.Cells row copy example C#

using Aspose.Cells;

// The sample loads source.xlsx and destination.xlsx, accesses their first worksheets, and uses Cells.CopyRow to copy the third row (index 2) from the source sheet to the sixth row (index 5) of the destination sheet, then saves the result as destination_updated.xlsx.
class Program
{
    static void Main()
    {
        // Load the source workbook
        Workbook sourceWorkbook = new Workbook("source.xlsx");
        // Load (or create) the destination workbook
        Workbook destinationWorkbook = new Workbook("destination.xlsx");

        // Access the first worksheet in each workbook (adjust index as needed)
        Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
        Worksheet destinationSheet = destinationWorkbook.Worksheets[0];

        // Define the zero‑based row index to copy from the source sheet
        int sourceRowIndex = 2; // e.g., third row

        // Define the zero‑based row index where the row will be placed in the destination sheet
        int destinationRowIndex = 5; // e.g., sixth row

        // Copy the entire row from the source sheet to the destination sheet
        destinationSheet.Cells.CopyRow(sourceSheet.Cells, sourceRowIndex, destinationRowIndex);

        // Save the updated destination workbook
        destinationWorkbook.Save("destination_updated.xlsx");
    }
}
