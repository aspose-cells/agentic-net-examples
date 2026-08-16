// Title: Aspose.Cells .NET: Verify formulas auto‑update when a named range address changes
// Description: This C# example creates a workbook, defines a named range "MyRange" for A1:A3, uses it in a SUM formula, extends the data to A4, updates the range to A1:A4, recalculates the workbook, and demonstrates that the formula result changes from 60 to 100 without editing the formula.
// Keywords: Aspose.Cells named range update | C# RefersTo property | automatic formula recalculation | SUM(MyRange) after range change | programmatic named range extension
// Common Searches: Aspose.Cells update named range formula automatically | C# change Name.RefersTo and recalc formulas | how to extend a named range in Aspose.Cells .NET | formula refresh after named range modification Aspose | Aspose.Cells dependent formula update
// Developer Intent: Confirm that changing a named range's RefersTo address triggers automatic updates of all formulas that reference the range.
// Use Cases: Define a named range for a data block, use it in calculations, then expand the range and rely on Aspose.Cells to adjust results. | Programmatically adjust a range after inserting new rows and let the library recalculate dependent formulas. | Save the workbook after modifying the range to ensure the updated formula values persist.
// AI Prompts: Show how to change a named range address in Aspose.Cells for .NET and have formulas recalculate automatically. | Provide a C# snippet that validates formulas referencing a named range reflect the updated range without manual edits. | Explain Aspose.Cells' handling of dependent formula updates when Name.RefersTo is modified.

using System;
using Aspose.Cells;

// This C# example creates a workbook, defines a named range "MyRange" for A1:A3, uses it in a SUM formula, extends the data to A4, updates the range to A1:A4, recalculates the workbook, and demonstrates that the formula result changes from 60 to 100 without editing the formula.
class NamedRangeUpdateDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Sheet1";

        // Populate initial data in cells A1:A3
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].PutValue(20);
        sheet.Cells["A3"].PutValue(30);

        // Create a named range "MyRange" that refers to A1:A3
        int nameIndex = workbook.Worksheets.Names.Add("MyRange");
        Name myRange = workbook.Worksheets.Names[nameIndex];
        myRange.RefersTo = "=Sheet1!$A$1:$A$3";

        // Use the named range in a formula: B1 = SUM(MyRange)
        sheet.Cells["B1"].Formula = "=SUM(MyRange)";
        workbook.CalculateFormula();

        Console.WriteLine("Initial SUM(MyRange) = " + sheet.Cells["B1"].Value); // Expected: 60

        // Extend the data range by adding a value to A4
        sheet.Cells["A4"].PutValue(40);

        // Update the named range to include the new cell (A1:A4)
        myRange.RefersTo = "=Sheet1!$A$1:$A$4";

        // Recalculate formulas; the change propagates automatically
        workbook.CalculateFormula();

        Console.WriteLine("After extending range, SUM(MyRange) = " + sheet.Cells["B1"].Value); // Expected: 100

        // Save the workbook (optional verification)
        workbook.Save("NamedRangeUpdateDemo.xlsx");
    }
}
