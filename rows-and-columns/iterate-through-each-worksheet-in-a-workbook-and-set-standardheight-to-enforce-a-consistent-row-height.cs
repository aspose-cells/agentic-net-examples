using System;
using Aspose.Cells;

// Author: Aspose.Cells .NET example – set consistent row height across all worksheets
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Define the desired standard row height (in points)
        double desiredHeight = 20.0;

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Set the default row height for the current worksheet
            sheet.Cells.StandardHeight = desiredHeight;
        }

        // Save the workbook to a file
        workbook.Save("ConsistentRowHeight.xlsx");
    }
}