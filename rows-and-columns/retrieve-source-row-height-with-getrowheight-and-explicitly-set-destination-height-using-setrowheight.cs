// Title: How to copy row height from one row to another using GetRowHeight and SetRowHeight in Aspose.Cells for .NET (C#)
// AI Prompts: Read the height of row 2 with Cells.GetRowHeight and set row 5 to the same height using Cells.SetRowHeight in a C# Aspose.Cells workbook. | Write C# code that takes any source row index, retrieves its height via GetRowHeight, and applies that height to a target row with SetRowHeight in Aspose.Cells. | Show how to extract a row's height in points and assign it to another row before saving the workbook with Aspose.Cells.
// Common Searches: Aspose.Cells C# copy row height from one row to another | GetRowHeight example for setting another row height in .NET | How to duplicate row height using Aspose.Cells API | Set row height based on another row in Excel with Aspose.Cells | Retrieve and apply row height in points using Aspose.Cells C#
// Tags: Aspose.Cells GetRowHeight method | Aspose.Cells SetRowHeight method | copy row height between rows Aspose.Cells | row height in points Aspose.Cells | C# Aspose.Cells row formatting transfer

using System;
using Aspose.Cells;

// Creates a workbook, sets row 2 height to 30 points, reads that height with GetRowHeight, applies the same value to row 5 using SetRowHeight, prints both heights, and saves the file as RowHeightCopy.xlsx.
class RowHeightCopyDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Set a custom height for the source row (row index 2)
        int sourceRow = 2;
        cells.SetRowHeight(sourceRow, 30); // height in points

        // Retrieve the height of the source row
        double sourceHeight = cells.GetRowHeight(sourceRow);

        // Define the destination row index
        int destinationRow = 5;

        // Explicitly set the destination row height using the retrieved value
        cells.SetRowHeight(destinationRow, sourceHeight);

        // Output heights for verification
        Console.WriteLine($"Source row {sourceRow} height: {sourceHeight}");
        Console.WriteLine($"Destination row {destinationRow} height: {cells.GetRowHeight(destinationRow)}");

        // Save the workbook
        workbook.Save("RowHeightCopy.xlsx");
    }
}
