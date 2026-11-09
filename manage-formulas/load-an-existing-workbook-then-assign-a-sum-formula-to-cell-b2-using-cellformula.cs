// Title: Assign a SUM(A2:A10) formula to cell B2 in an existing Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# code to add a SUM formula to cell B2 after opening an Excel file with Aspose.Cells | How to set a formula for a specific cell in an existing workbook using Aspose.Cells .NET | Aspose.Cells example for writing =SUM(A2:A10) into B2 of a loaded workbook | Saving changes after inserting a formula into an Excel worksheet with Aspose.Cells C# | Update existing Excel file by assigning a SUM range formula to a cell using Aspose.Cells
// Tags: assign SUM formula to cell Aspose.Cells | write cell formula Excel C# | modify existing workbook formula Aspose.Cells | load workbook and set cell formula .NET | save workbook after adding formula Aspose

using Aspose.Cells;
using System;

// Loads an existing Excel workbook, assigns the formula =SUM(A2:A10) to cell B2 on the first worksheet using Aspose.Cells for .NET, and saves the updated file.
class Program
{
    static void Main()
    {
        // Load an existing workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (index 0)
        Worksheet worksheet = workbook.Worksheets[0];

        // Access cell B2
        Cell targetCell = worksheet.Cells["B2"];

        // Assign a SUM formula to B2 (example: sum of A2 through A10)
        targetCell.Formula = "=SUM(A2:A10)";

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
