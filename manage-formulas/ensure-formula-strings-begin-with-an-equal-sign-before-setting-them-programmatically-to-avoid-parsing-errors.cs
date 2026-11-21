// Title: Prepending a leading '=' to Excel formulas before assigning Cell.Formula with Aspose.Cells for .NET
// Common Searches: c# add missing leading equal sign to Excel formula using Aspose.Cells | avoid formula parsing exception in Aspose.Cells when setting cell formula | programmatically validate Excel formula string before assigning to Cell.Formula in .NET | how to ensure formulas start with '=' when using Aspose.Cells Workbook | prepend '=' to raw formula string Aspose.Cells C# example
// Tags: prepend leading equal sign to Excel formula Aspose.Cells | validate formula string before Cell.Formula assignment | prevent formula parsing errors Aspose.Cells C# | set cell formula with correct syntax Aspose.Cells | raw formula handling Aspose.Cells .NET

using Aspose.Cells;
using System;

// Demonstrates how to check a raw formula string for a leading '=', prepend it if missing, assign the corrected formula to a cell, populate sample data, and save the workbook using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Example formula that might be missing the leading '='
        string rawFormula = "SUM(A1:A10)";

        // Ensure the formula string starts with '=' to avoid parsing errors
        string formula = rawFormula.StartsWith("=") ? rawFormula : "=" + rawFormula;

        // Assign the formula to cell B1
        Cell cell = sheet.Cells["B1"];
        cell.Formula = formula;

        // Populate cells A1:A10 with sample numeric data
        for (int i = 0; i < 10; i++)
        {
            sheet.Cells[i, 0].PutValue(i + 1);
        }

        // Save the workbook to a file
        workbook.Save("output.xlsx");
    }
}
