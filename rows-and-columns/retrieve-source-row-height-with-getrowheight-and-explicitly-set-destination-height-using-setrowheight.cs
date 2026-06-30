using System;
using Aspose.Cells;

// Author: Aspose.Cells .NET example – copy row height from source to destination
class Program
{
    static void Main()
    {
        // Create a new workbook (use the provided create rule)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Example: set a specific height for the source row (row 0)
        int sourceRowIndex = 0;
        cells.SetRowHeight(sourceRowIndex, 30.0); // height in points

        // Retrieve the source row height using GetRowHeight
        double sourceHeight = cells.GetRowHeight(sourceRowIndex);
        // If GetRowHeight is obsolete in your version, you can use:
        // double sourceHeight = cells.GetViewRowHeight(sourceRowIndex);

        // Explicitly set the destination row height (row 1) to match the source
        int destinationRowIndex = 1;
        cells.SetRowHeight(destinationRowIndex, sourceHeight);

        // Save the workbook (use the provided save rule)
        workbook.Save("RowHeightCopy.xlsx");
    }
}