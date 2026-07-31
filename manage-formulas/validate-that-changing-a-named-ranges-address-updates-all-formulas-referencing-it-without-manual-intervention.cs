// Title: Aspose.Cells for .NET – Verify formulas auto‑update after changing a named‑range address
// Description: This C# example creates a workbook, defines a named range "MyRange" (A1:A3), uses it in a SUM formula, expands the range to A1:A4, recalculates the workbook, and confirms that the formula result and the range address are refreshed automatically without editing the formula.
// Keywords: Aspose.Cells .NET named range | C# RefersTo property | dynamic named range formula | automatic formula recalculation | Name.GetRange example | update named range address | SUM(MyRange) Aspose.Cells | Workbook.CalculateFormula | Excel named range API | Aspose.Cells GitHub sample
// Common Searches: how to update a named range in Aspose.Cells | C# change named range address and recalculate formulas | Aspose.Cells automatic formula refresh after RefersTo change | verify SUM(MyRange) updates in .NET | Name.GetRange returns new address Aspose.Cells
// Developer Intent: Confirm that modifying a Name.RefersTo value instantly propagates to all dependent formulas.
// Use Cases: Create a named range, reference it in a formula, then extend the range and observe the updated result. | Retrieve the modified range with Name.GetRange() to validate the new address. | Save the workbook after the dynamic range change for downstream processing.
// AI Prompts: Generate C# code with Aspose.Cells that defines a named range, uses it in a formula, changes the range address, and automatically recalculates the workbook. | Explain why calling Workbook.CalculateFormula() after updating Name.RefersTo updates all formulas that reference the named range. | Provide a step‑by‑step verification that Name.GetRange() reflects the new address after modifying RefersTo.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsNamedRangeUpdateDemo
{
    // This C# example creates a workbook, defines a named range "MyRange" (A1:A3), uses it in a SUM formula, expands the range to A1:A4, recalculates the workbook, and confirms that the formula result and the range address are refreshed automatically without editing the formula.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate initial data in cells A1:A3
                sheet.Cells["A1"].PutValue(10);
                sheet.Cells["A2"].PutValue(20);
                sheet.Cells["A3"].PutValue(30);

                // Add a named range "MyRange" that refers to A1:A3
                int nameIndex = sheet.Workbook.Worksheets.Names.Add("MyRange");
                Name myRange = sheet.Workbook.Worksheets.Names[nameIndex];
                myRange.RefersTo = "=Sheet1!$A$1:$A$3";

                // Use the named range in a formula (SUM) placed in B1
                sheet.Cells["B1"].Formula = "=SUM(MyRange)";

                // Calculate formulas and display the result
                workbook.CalculateFormula();
                Console.WriteLine($"Initial SUM(MyRange) = {sheet.Cells["B1"].Value}"); // Expected 60

                // Extend the named range to include A4 and add a value to A4
                sheet.Cells["A4"].PutValue(40);
                myRange.RefersTo = "=Sheet1!$A$1:$A$4";

                // Recalculate formulas after the named range change
                workbook.CalculateFormula();

                // Display the updated result; it should now include the value from A4
                Console.WriteLine($"Updated SUM(MyRange) = {sheet.Cells["B1"].Value}"); // Expected 100

                // Verify that GetRange reflects the new address
                AsposeRange updatedRange = myRange.GetRange();
                Console.WriteLine($"Named range now refers to: {updatedRange.Address}"); // Expected A1:A4

                // Save the workbook (optional)
                workbook.Save("NamedRangeUpdateDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
