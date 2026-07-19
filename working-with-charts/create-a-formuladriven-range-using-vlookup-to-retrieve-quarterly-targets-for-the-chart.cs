// Title: C# – Apply Aspose.Cells SetSharedFormula with VLOOKUP to Generate Quarterly Target Data for Charts
// Description: This example shows how to create a new workbook, build a quarter‑to‑target lookup table, list quarters, apply a shared VLOOKUP formula (=VLOOKUP(C1,$A$1:$B$4,2,FALSE)) across D1:D4, calculate all formulas, output the results, and save the file as QuarterlyTargets.xlsx – ready to be used as a chart data source.
// Keywords: Aspose.Cells VLOOKUP C# | SetSharedFormula Aspose.Cells | quarterly targets Excel | chart data source Aspose.Cells | calculate formulas Aspose.Cells | save workbook C# | Excel lookup table programmatically | .NET Excel automation
// Common Searches: Aspose.Cells VLOOKUP across a range | C# SetSharedFormula example | How to create chart data with VLOOKUP in Aspose.Cells | Calculate formulas after setting VLOOKUP Aspose.Cells | Generate quarterly target series for Excel chart using Aspose.Cells
// Developer Intent: Create a column of quarterly target values using a shared VLOOKUP formula so the data can be bound to an Excel chart.
// Use Cases: Build a dynamic series for a column or line chart by looking up targets based on quarter identifiers. | Reuse a single VLOOKUP formula across multiple rows to populate chart data without manual entry. | Validate lookup results programmatically before exporting the workbook for reporting or dashboarding.
// AI Prompts: Show C# code that uses Aspose.Cells SetSharedFormula to apply VLOOKUP across D1:D4 and then calculates the workbook. | Generate an Aspose.Cells example that creates a quarter‑to‑target lookup table, retrieves values with VLOOKUP, and saves the workbook for charting. | Explain absolute vs. relative references in a VLOOKUP formula when using SetSharedFormula in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsVLookupExample
{
    // This example shows how to create a new workbook, build a quarter‑to‑target lookup table, list quarters, apply a shared VLOOKUP formula (=VLOOKUP(C1,$A$1:$B$4,2,FALSE)) across D1:D4, calculate all formulas, output the results, and save the file as QuarterlyTargets.xlsx – ready to be used as a chart data source.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // -----------------------------------------------------------------
            // 1. Prepare the lookup table (Quarter -> Target)
            // -----------------------------------------------------------------
            // Quarter identifiers
            cells["A1"].PutValue("Q1");
            cells["A2"].PutValue("Q2");
            cells["A3"].PutValue("Q3");
            cells["A4"].PutValue("Q4");

            // Corresponding quarterly targets
            cells["B1"].PutValue(120000);
            cells["B2"].PutValue(150000);
            cells["B3"].PutValue(130000);
            cells["B4"].PutValue(160000);

            // -----------------------------------------------------------------
            // 2. List of quarters for which we want to retrieve targets
            //    (this could be the source data for a chart)
            // -----------------------------------------------------------------
            cells["C1"].PutValue("Q1");
            cells["C2"].PutValue("Q2");
            cells["C3"].PutValue("Q3");
            cells["C4"].PutValue("Q4");

            // -----------------------------------------------------------------
            // 3. Apply VLOOKUP formula to retrieve the target for each quarter.
            //    Use SetSharedFormula to propagate the same relative formula
            //    across the range D1:D4 (4 rows, 1 column).
            // -----------------------------------------------------------------
            // Formula: =VLOOKUP(C1,$A$1:$B$4,2,FALSE)
            // C1 is relative; $A$1:$B$4 is absolute.
            cells["D1"].SetSharedFormula("=VLOOKUP(C1,$A$1:$B$4,2,FALSE)", 4, 1);

            // -----------------------------------------------------------------
            // 4. Calculate all formulas so that the VLOOKUP results are materialized.
            // -----------------------------------------------------------------
            workbook.CalculateFormula();

            // -----------------------------------------------------------------
            // 5. (Optional) Verify the results by printing them to the console.
            // -----------------------------------------------------------------
            for (int i = 0; i < 4; i++)
            {
                Console.WriteLine($"Quarter {cells[i, 2].StringValue} Target = {cells[i, 3].Value}");
            }

            // Save the workbook (lifecycle: save)
            workbook.Save("QuarterlyTargets.xlsx");
        }
    }
}
