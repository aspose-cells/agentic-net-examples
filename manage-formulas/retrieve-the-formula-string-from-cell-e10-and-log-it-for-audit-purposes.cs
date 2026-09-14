// Title: How to retrieve and log the formula from cell E10 in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an .xlsx file with Aspose.Cells, reads the Formula property of cell E10, and writes the result to the console. | Show a snippet that extracts the formula string from a specific worksheet cell and logs it for audit tracking using the Aspose.Cells API.
// Common Searches: Aspose.Cells C# get formula text from a specific cell for audit logging | How to read the formula of cell E10 in an Excel file using Aspose.Cells .NET | Example code to output Excel cell formula to console with Aspose.Cells | Retrieve and print formula string from worksheet cell using Aspose.Cells library
// Tags: Aspose.Cells read cell formula .NET | extract formula string from Excel cell | audit Excel formulas with Aspose.Cells | log cell formula to console C# | retrieve Formula property Aspose.Cells

using Aspose.Cells;
using System;

// Loads an .xlsx workbook, accesses cell E10 on the first worksheet, reads its Formula property, and writes the formula string to the console for auditing.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with actual file path)
        var workbook = new Workbook("input.xlsx");

        // Access the first worksheet (or specify the worksheet name)
        var worksheet = workbook.Worksheets[0];

        // Retrieve cell E10
        var cell = worksheet.Cells["E10"];

        // Get the formula string from the cell
        string formula = cell.Formula;

        // Log the formula for audit purposes
        Console.WriteLine($"Formula in E10: {formula}");
    }
}
