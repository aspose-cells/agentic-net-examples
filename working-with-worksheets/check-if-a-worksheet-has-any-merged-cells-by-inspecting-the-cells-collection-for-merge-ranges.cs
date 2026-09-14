// Title: How to check if a worksheet contains any merged cells using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells and returns a boolean indicating whether the first worksheet has any merged cell ranges. | Generate a method in C# using Aspose.Cells that prints a custom message based on the presence of merged cells in a given worksheet. | Create a C# snippet that counts the number of merged cell blocks in a worksheet via Aspose.Cells and outputs the count.
// Common Searches: asp.net aspose.cells detect merged cells in worksheet c# | c# check if excel sheet has merged cells using aspose.cells library | how to count merged cell ranges in an Excel file with aspose.cells c# | determine presence of merged cells in workbook using aspose.cells for .net
// Tags: detect merged cell ranges Aspose.Cells | worksheet merged cells check C# | Cells.MergedCells count Aspose.Cells | Excel merged cells detection C# | Aspose.Cells worksheet merge detection

using System;
using Aspose.Cells;

// Loads an Excel workbook, accesses a worksheet, uses worksheet.Cells.MergedCells.Count to determine if any merged cells exist, and writes a message indicating the result.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (or any specific worksheet by index/name)
        Worksheet worksheet = workbook.Worksheets[0];

        // Determine if there are any merged cells in the worksheet
        bool hasMergedCells = worksheet.Cells.MergedCells.Count > 0;

        // Output the result
        Console.WriteLine(hasMergedCells
            ? "The worksheet contains merged cells."
            : "The worksheet does not contain any merged cells.");
    }
}
