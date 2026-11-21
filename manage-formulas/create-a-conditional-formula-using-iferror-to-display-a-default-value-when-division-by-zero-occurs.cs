// Title: Use IFERROR in Aspose.Cells for .NET to return a custom value when a division by zero occurs
// Common Searches: how to apply IFERROR in Aspose.Cells C# to avoid #DIV/0! errors | Aspose.Cells set formula with fallback text for division by zero in .NET | C# example of using IFERROR to display N/A when dividing cells with Aspose | calculate Excel formulas with Aspose.Cells and replace errors with custom string | programmatically handle division errors in Excel using Aspose.Cells API
// Tags: Aspose.Cells IFERROR formula | division by zero handling Aspose.Cells | set cell formula programmatically .NET | calculate workbook formulas Aspose.Cells | custom error value Excel formula C#

using System;
using Aspose.Cells;

// The sample creates a workbook, writes 10 to A1 and 0 to B1, assigns an IFERROR formula to C1 that returns "N/A" when A1/B1 triggers a division‑by‑zero error, forces formula calculation, and saves the file as ConditionalFormula.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate cells with sample data
        // A1 = numerator, B1 = denominator
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["B1"].PutValue(0); // This will cause division by zero

        // Set a formula in C1 that uses IFERROR to handle division by zero
        // If A1/B1 results in an error, display the default value "N/A"
        sheet.Cells["C1"].Formula = "IFERROR(A1/B1, \"N/A\")";

        // Optionally calculate the formulas immediately
        workbook.CalculateFormula();

        // Save the workbook to a file
        workbook.Save("ConditionalFormula.xlsx");
    }
}
