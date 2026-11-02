// Title: How to obtain the exact formula text of an Excel cell using Aspose.Cells Formula property in C#
// Common Searches: Aspose.Cells C# get formula string from a specific cell | retrieve Excel cell formula as plain text with Aspose.Cells .NET | display the formula entered in cell A1 using Aspose.Cells in C# | read a worksheet cell's formula without evaluating it via Aspose.Cells API | C# example for extracting raw Excel formula using Aspose.Cells
// Tags: Aspose.Cells Formula property example | C# extract raw Excel formula | Aspose.Cells get formula text programmatically | .NET retrieve Excel formula string | Excel cell formula extraction using Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a workbook, writes the formula "=SUM(B1:B5)" to cell A1, then accesses the cell's Formula property to obtain the exact textual representation of the formula and prints it to the console.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Put a formula into cell A1
        Cell cell = sheet.Cells["A1"];
        cell.Formula = "=SUM(B1:B5)";

        // Use FormulaText (the Formula property returns the exact textual representation)
        string formulaText = cell.Formula;

        // Output the formula text
        Console.WriteLine($"Formula in A1: {formulaText}");
    }
}
