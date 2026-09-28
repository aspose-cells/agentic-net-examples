// Title: Insert a SUM formula into cell G10 to calculate the total of B2:B9 with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells to place the formula =SUM(B2:B9) in cell G10 of a new workbook and save it as output.xlsx. | Write a C# snippet that creates an Excel file, accesses the first worksheet, sets a SUM range formula in G10, and persists the workbook using Aspose.Cells. | Provide a C# example showing how to assign a SUM(B2:B9) formula to cell G10 and write the workbook to disk with Aspose.Cells.
// Common Searches: Aspose.Cells C# how to add a SUM formula to a specific cell | set formula =SUM(B2:B9) in G10 using Aspose.Cells .NET | C# Aspose.Cells write formula to cell and save workbook
// Tags: Aspose.Cells set cell formula C# | SUM function insertion Excel Aspose.Cells | assign formula to G10 worksheet .NET | calculate range total Aspose.Cells C# | persist workbook with formula Aspose.Cells

using Aspose.Cells;
using System;

// Creates a new workbook, accesses the first worksheet, inserts the formula =SUM(B2:B9) into cell G10, and saves the file as output.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Insert formula in cell G10 to calculate the sum of range B2:B9
        Cell targetCell = sheet.Cells["G10"];
        targetCell.Formula = "=SUM(B2:B9)";

        // Save the workbook to a file
        workbook.Save("output.xlsx");
    }
}
