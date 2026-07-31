// Title: Set a SEQUENCE dynamic array formula in column E across all worksheets with Aspose.Cells for .NET
// Description: C# example that creates a workbook, adds several sheets, fills column D with sample data, then loops through every worksheet to assign the dynamic array formula "=SEQUENCE(D1)" to cell E1 using SetDynamicArrayFormula, recalculates, refreshes spill ranges, and saves the file as XLSX.
// Keywords: Aspose.Cells | C# | .NET | dynamic array formula | SEQUENCE function | SetDynamicArrayFormula | loop worksheets | refresh spill range | calculate formulas | save workbook | Excel automation | multiple sheets
// Common Searches: Aspose.Cells set dynamic array formula on all sheets | C# apply SEQUENCE formula to each worksheet | Refresh spill ranges after inserting dynamic array in Aspose.Cells | Loop through worksheets and set formula Aspose.Cells .NET | Calculate and save workbook with dynamic arrays
// Developer Intent: Apply the same dynamic array formula to column E of every worksheet, ensure the spill results are updated, and persist the workbook.
// Use Cases: Generate a sequence in column E that depends on the value in D1 for each sheet. | Automate bulk formula insertion across many worksheets in an Excel file. | Guarantee that dynamic array spill ranges are refreshed before saving the workbook.
// AI Prompts: Write C# code that iterates over all worksheets in an Aspose.Cells workbook and sets the formula "=SEQUENCE(D1)" in cell E1 using SetDynamicArrayFormula, then calculates and refreshes the workbook. | Show how to use FormulaParseOptions with SetDynamicArrayFormula to apply a dynamic array formula to every sheet and save the result as an XLSX file.

using System;
using Aspose.Cells;

namespace AsposeCellsDynamicArrayExample
{
    // C# example that creates a workbook, adds several sheets, fills column D with sample data, then loops through every worksheet to assign the dynamic array formula "=SEQUENCE(D1)" to cell E1 using SetDynamicArrayFormula, recalculates, refreshes spill ranges, and saves the file as XLSX.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Add a few worksheets for demonstration
            Worksheet sheet1 = workbook.Worksheets[0]; // default sheet
            sheet1.Name = "FirstSheet";

            Worksheet sheet2 = workbook.Worksheets.Add("SecondSheet");
            Worksheet sheet3 = workbook.Worksheets.Add("ThirdSheet");

            // Sample data that the dynamic array formulas will reference
            // (optional, just to make the formulas produce visible results)
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // Populate column D with numbers 1..5
                for (int i = 0; i < 5; i++)
                {
                    ws.Cells[i, 3].PutValue(i + 1); // D1:D5
                }
            }

            // Iterate through all worksheets and set a dynamic array formula in column E
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // The formula will generate a sequence based on the value in D1 (e.g., =SEQUENCE(D1))
                // It will spill into the cells below E1 as needed.
                Cell targetCell = ws.Cells["E1"];
                string dynamicFormula = "=SEQUENCE(D1)";

                // Set the dynamic array formula; calculateValue = true to compute immediately
                targetCell.SetDynamicArrayFormula(dynamicFormula, new FormulaParseOptions(), true);
            }

            // Calculate all formulas in the workbook (optional, ensures values are up‑to‑date)
            workbook.CalculateFormula();

            // Refresh dynamic array formulas so that spill ranges are correctly updated
            workbook.RefreshDynamicArrayFormulas(true);

            // Save the workbook to a file (lifecycle: save)
            workbook.Save("DynamicArrayResult.xlsx", SaveFormat.Xlsx);
        }
    }
}
