using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add worksheets
        workbook.Worksheets.Add("First");
        workbook.Worksheets.Add("Second");
        workbook.Worksheets.Add("Third");

        // Move the "Third" worksheet to the first position (index 0)
        Worksheet sheetToMove = workbook.Worksheets["Third"];
        sheetToMove.MoveTo(0);

        // Save the workbook
        workbook.Save("MovedWorksheet.xlsx");
    }
}

// Author: Aspose.Cells .NET example code.