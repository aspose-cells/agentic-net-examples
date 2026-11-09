// Title: Extract the formula from cell E10 and write it to the console using Aspose.Cells for .NET
// Common Searches: C# code to read the formula of a specific Excel cell with Aspose.Cells | How to get the formula string from cell E10 using the Aspose.Cells library | Logging Excel cell formulas for audit in a .NET application | Aspose.Cells retrieve cell formula and output to console example | Extracting an Excel formula from a workbook with Aspose.Cells in C#
// Tags: extract cell formula Aspose.Cells C# | read formula property Aspose.Cells | log Excel formula to console .NET | address based cell access Aspose.Cells | audit worksheet formulas programmatically

using Aspose.Cells;
using System;

// The example loads an Excel workbook with Aspose.Cells, accesses the first worksheet, retrieves the formula from cell E10 via the Cell.Formula property, and writes the formula text to the console for auditing purposes.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with the actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the desired worksheet (here we use the first worksheet)
        Worksheet worksheet = workbook.Worksheets[0];

        // Retrieve the cell at E10
        Cell cell = worksheet.Cells["E10"];

        // Get the formula string from the cell
        string formula = cell.Formula;

        // Log the formula for audit purposes
        Console.WriteLine($"Formula in E10: {formula}");
    }
}
