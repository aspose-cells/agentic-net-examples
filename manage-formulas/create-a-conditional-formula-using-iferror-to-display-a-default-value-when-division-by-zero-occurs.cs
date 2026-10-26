// Title: Insert an IFERROR formula with Aspose.Cells for .NET (C#) to return a default value on division‑by‑zero
// Common Searches: Aspose.Cells C# how to use IFERROR to catch division by zero | set a default text for Excel formula errors using Aspose.Cells .NET | apply IFERROR to a worksheet cell programmatically with Aspose.Cells | calculate workbook after adding IFERROR formula in C# | avoid #DIV/0! error in Aspose.Cells by using IFERROR
// Tags: IFERROR formula insertion Aspose.Cells C# | default error value in Excel formula Aspose.Cells | programmatic formula evaluation Aspose.Cells .NET | save workbook after calculation Aspose.Cells | handling division errors with IFERROR Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a workbook, writes 10 to A1 and 0 to B1, inserts an IFERROR formula in C1 that returns "N/A" when A1/B1 causes a division‑by‑zero error, evaluates the formulas, and saves the file as ConditionalFormula.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data: A1 = 10, B1 = 0 (division by zero scenario)
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["B1"].PutValue(0);

        // Apply IFERROR formula in C1.
        // If A1/B1 results in an error (e.g., division by zero), display "N/A" instead.
        sheet.Cells["C1"].Formula = "=IFERROR(A1/B1, \"N/A\")";

        // Evaluate the formula so the result is stored in the cell
        workbook.CalculateFormula();

        // Save the workbook to a file
        workbook.Save("ConditionalFormula.xlsx");
    }
}
