// Title: How to disable auto‑calculation, edit cells, and manually trigger Workbook.CalculateFormula in Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that turns off automatic calculation, updates multiple cell values, sets a formula, and then calls Workbook.CalculateFormula using Aspose.Cells. | Show an example of disabling auto‑calc in Aspose.Cells, modifying A1 and B1, assigning a dependent formula to C1, and manually recalculating the workbook before saving. | Provide a step‑by‑step C# snippet to prevent automatic formula evaluation, change cell contents, add a formula, and invoke manual calculation with Workbook.CalculateFormula.
// Common Searches: Aspose.Cells C# disable automatic calculation before editing cells | Manually recalculate formulas after setting cell values with Aspose.Cells .NET | Workbook.CalculateFormula example after updating dependent cells in C# | How to turn off auto‑calc in Aspose.Cells and trigger calculation later | Set formula programmatically and force calculation in Aspose.Cells for .NET
// Tags: auto‑calc off Aspose.Cells | manual formula recalculation Aspose.Cells | programmatic cell update Aspose.Cells C# | formula assignment Aspose.Cells | Workbook.CalculateFormula .NET

using Aspose.Cells;
using System;
using System.IO;

namespace AsposeCellsExample
{
    // Demonstrates disabling automatic calculation, updating cells A1 and B1, assigning a formula to C1 that references those cells, manually invoking Workbook.CalculateFormula to evaluate dependent formulas, and saving the workbook as output.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook();

                // Get the first worksheet and its cells collection
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Modify several cells
                cells["A1"].PutValue(10);          // Set numeric value
                cells["B1"].PutValue(20);          // Set another numeric value
                cells["C1"].Formula = "=A1+B1";    // Formula dependent on A1 and B1

                // Manually invoke calculation to update dependent formulas
                workbook.CalculateFormula();

                // Save the workbook to a file
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
