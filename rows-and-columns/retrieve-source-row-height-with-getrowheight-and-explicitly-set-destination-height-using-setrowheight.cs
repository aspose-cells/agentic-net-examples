// Title: Transfer a specific row’s height from one Excel workbook to another using Aspose.Cells GetRowHeight and SetRowHeight in C#
// AI Prompts: Write C# code that reads the height of row 5 from source.xlsx with Aspose.Cells and applies the same height to row 5 in target.xlsx using Cells.GetRowHeight and Cells.SetRowHeight. | Create a reusable C# method that accepts source and destination workbook paths, a row index, and transfers the row height from the source sheet to the destination sheet using Aspose.Cells.
// Common Searches: how to copy row height from one Excel file to another using Aspose.Cells C# | Aspose.Cells GetRowHeight example for transferring row height | C# set row height in destination workbook based on source workbook row | copying Excel row height programmatically with Aspose.Cells
// Tags: Aspose.Cells row height retrieval | Aspose.Cells row height assignment | replicate row height across Excel workbooks C# | transfer row height between workbooks Aspose | C# programmatic Excel row height setting

using Aspose.Cells;

// The example loads a source workbook, reads a specific row’s height with Cells.GetRowHeight, opens a destination workbook, and applies the retrieved height to the same row index using Cells.SetRowHeight before saving.
class RowHeightTransfer
{
    static void Main()
    {
        // Load the source workbook
        Workbook sourceWorkbook = new Workbook("source.xlsx");
        Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

        // Specify the row index (0‑based) to copy the height from
        int sourceRowIndex = 2; // example: third row

        // Retrieve the height of the source row
        double sourceRowHeight = sourceSheet.Cells.GetRowHeight(sourceRowIndex);

        // Load (or create) the destination workbook
        Workbook destinationWorkbook = new Workbook("destination.xlsx");
        Worksheet destinationSheet = destinationWorkbook.Worksheets[0];

        // Specify the row index (0‑based) to apply the height to
        int destinationRowIndex = 2; // example: third row

        // Explicitly set the destination row height using the retrieved value
        destinationSheet.Cells.SetRowHeight(destinationRowIndex, sourceRowHeight);

        // Save the changes to the destination file
        destinationWorkbook.Save("destination.xlsx");
    }
}
