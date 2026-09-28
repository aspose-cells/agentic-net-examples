// Title: Insert an IFERROR division formula with a default value into an Aspose.Cells worksheet using C#
// AI Prompts: Generate C# code that creates a new Aspose.Cells workbook, writes 10 to A1, 0 to B1, sets C1 formula to =IFERROR(A1/B1, "N/A"), calculates the sheet, and saves it as ConditionalFormula.xlsx. | Show how to apply an IFERROR expression to a cell in Aspose.Cells so that any division‑by‑zero error returns a custom text value. | Demonstrate programmatic evaluation of an IFERROR formula after inserting it into a worksheet with Aspose.Cells for .NET.
// Common Searches: how to use IFERROR in Aspose.Cells C# to avoid #DIV/0! errors | Aspose.Cells C# set default text for Excel formula errors | example of inserting conditional formula with IFERROR and saving workbook in .NET
// Tags: Aspose.Cells IFERROR formula C# | division by zero handling Aspose.Cells | calculate workbook formulas Aspose.Cells | save Excel file Aspose.Cells C# | set cell formula programmatically Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Creates a workbook with numerator and denominator values, applies an IFERROR formula that returns "N/A" on division by zero, forces calculation, and saves the file using Aspose.Cells for C#.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data
            cells["A1"].PutValue(10);   // Numerator
            cells["B1"].PutValue(0);    // Denominator (zero to trigger division by zero)

            // Set a formula using IFERROR to handle division by zero.
            // If A1/B1 results in an error, the cell will display the default value "N/A".
            cells["C1"].Formula = "=IFERROR(A1/B1, \"N/A\")";

            // Optionally calculate the formula immediately
            workbook.CalculateFormula();

            // Save the workbook to a file
            workbook.Save("ConditionalFormula.xlsx");
        }
    }
}
