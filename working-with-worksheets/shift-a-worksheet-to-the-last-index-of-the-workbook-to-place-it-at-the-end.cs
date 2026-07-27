using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (contains a default sheet)
        Workbook workbook = new Workbook();

        // Add additional worksheets
        workbook.Worksheets.Add("First");
        workbook.Worksheets.Add("Second");
        workbook.Worksheets.Add("Third");

        // Select the worksheet you want to move (e.g., "First")
        Worksheet sheetToMove = workbook.Worksheets["First"];

        // Destination index is the last position (zero‑based)
        int lastIndex = workbook.Worksheets.Count - 1;

        // Move the worksheet to the end of the workbook
        sheetToMove.MoveTo(lastIndex);

        // Save the workbook
        workbook.Save("ShiftedToEnd.xlsx");
    }
}

// Author: Example demonstrating how to shift a worksheet to the last index using Aspose.Cells for .NET.