// Title: How to copy a row height from one Excel worksheet to another using GetRowHeight and SetRowHeight in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# copy row height from source workbook to another workbook | GetRowHeight and SetRowHeight example for transferring row size in Excel files | Set row height in destination sheet based on source sheet row height Aspose.Cells | C# code to read row height and apply it to a different worksheet using Aspose.Cells | Transfer Excel row dimensions between workbooks with Aspose.Cells API
// Tags: Aspose.Cells GetRowHeight method | Aspose.Cells SetRowHeight method | replicate row height across Excel workbooks C# | transfer row dimensions Aspose.Cells | row height adjustment Aspose.Cells .NET

using Aspose.Cells;

// The example loads a source workbook and a destination workbook, reads the height of a specific row from the source worksheet using GetRowHeight, applies that height to a row in the destination worksheet with SetRowHeight, and saves the updated file.
class Program
{
    static void Main()
    {
        // Load the source workbook containing the row whose height we want to copy
        Workbook sourceWorkbook = new Workbook("source.xlsx");
        // Load the destination workbook where we will apply the row height
        Workbook destinationWorkbook = new Workbook("dest.xlsx");

        // Select the first worksheet in each workbook (adjust index as needed)
        Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
        Worksheet destinationSheet = destinationWorkbook.Worksheets[0];

        // Define the zero‑based row index in the source sheet
        int sourceRowIndex = 4; // e.g., 5th row in Excel

        // Retrieve the height of the source row using GetRowHeight
        double sourceRowHeight = sourceSheet.Cells.GetRowHeight(sourceRowIndex);

        // Define the zero‑based row index in the destination sheet where the height will be set
        int destinationRowIndex = 9; // e.g., 10th row in Excel

        // Explicitly set the destination row height using SetRowHeight
        destinationSheet.Cells.SetRowHeight(destinationRowIndex, sourceRowHeight);

        // Save the updated destination workbook
        destinationWorkbook.Save("dest_updated.xlsx");
    }
}
