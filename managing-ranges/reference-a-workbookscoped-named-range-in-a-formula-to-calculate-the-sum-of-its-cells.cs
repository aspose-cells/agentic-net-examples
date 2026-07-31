// Title: Aspose.Cells C# – Create a Workbook‑Scoped Named Range and Sum It
// Description: Demonstrates how to add a workbook‑scoped named range in a new workbook, assign it to cells A1:A3, use =SUM(namedRange) in a formula, force calculation, retrieve the result, and save the file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# named range | workbook scoped name | SUM formula | calculate sum Aspose.Cells | named range RefersTo | Excel automation .NET | formula calculation
// Common Searches: Aspose.Cells create workbook scoped named range C# | How to sum a named range with Aspose.Cells | C# Aspose.Cells SUM formula using named range | Reference named range in formula Aspose.Cells
// Developer Intent: Create a workbook‑scoped named range and reference it in a SUM formula to obtain the total of its cells.
// Use Cases: Compute column totals with a single named range reference. | Apply the same named range across multiple sheets for consistent calculations. | Generate financial or reporting workbooks where summed values are needed before saving.
// AI Prompts: Show C# code using Aspose.Cells to add a workbook‑scoped named range and calculate its sum with =SUM(MyRange). | Explain how to set the RefersTo property for a named range and retrieve the calculated result in Aspose.Cells. | Provide step‑by‑step instructions to create, reference, and persist a named‑range sum in an Aspose.Cells workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsNamedRangeSum
{
    // Demonstrates how to add a workbook‑scoped named range in a new workbook, assign it to cells A1:A3, use =SUM(namedRange) in a formula, force calculation, retrieve the result, and save the file with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet (default name is "Sheet1")
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some numeric data that will be summed
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);

            // Add a workbook‑scoped named range called "MyRange"
            // The index returned by Names.Add is the position of the new name in the collection
            int nameIndex = workbook.Worksheets.Names.Add("MyRange");
            Name myRange = workbook.Worksheets.Names[nameIndex];

            // Define the range that the name refers to (absolute reference to A1:A3 on Sheet1)
            myRange.RefersTo = "=Sheet1!$A$1:$A$3";

            // Use the named range in a formula to calculate its sum
            sheet.Cells["B1"].Formula = "=SUM(MyRange)";

            // Force calculation of all formulas in the workbook
            workbook.CalculateFormula();

            // Output the result of the SUM formula
            Console.WriteLine("Sum of MyRange: " + sheet.Cells["B1"].Value);

            // Save the workbook (optional, demonstrates full lifecycle)
            workbook.Save("NamedRangeSum.xlsx");
        }
    }
}
