// Title: Set an absolute RefersTo address for a global named range in Aspose.Cells for .NET
// Description: Shows how to create a workbook with Aspose.Cells for .NET, add a global named range, assign an absolute address using SetRefersTo (or the RefersTo property) so the range stays fixed when rows or columns are inserted, use the range in a SUM formula, and save the workbook.
// Keywords: Aspose.Cells | .NET | C# | named range | absolute reference | SetRefersTo | RefersTo property | prevent range shift | global name | Excel formula | Workbook.Save | $A$1:$A$3
// Common Searches: Aspose.Cells set named range absolute address | SetRefersTo example C# | prevent named range from moving Aspose.Cells | global named range RefersTo property .NET | absolute cell reference in Aspose.Cells
// Developer Intent: Create a global named range with a fixed absolute reference that remains unchanged after worksheet modifications.
// Use Cases: Define a range that always points to A1:A3, regardless of inserted rows or columns, and use it in calculations. | Share the same absolute named range across multiple sheets for consistent reporting. | Export workbooks where the named range must stay static for downstream data processing.
// AI Prompts: Generate C# code that adds a global named range in Aspose.Cells and sets its RefersTo to an absolute address using SetRefersTo. | Explain how to keep a named range from shifting when rows are inserted in an Aspose.Cells workbook. | Show how to verify an absolute named range in a formula and save the workbook with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Shows how to create a workbook with Aspose.Cells for .NET, add a global named range, assign an absolute address using SetRefersTo (or the RefersTo property) so the range stays fixed when rows or columns are inserted, use the range in a SUM formula, and save the workbook.
class SetAbsoluteNamedRange
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Sheet1";

        // Populate some sample data
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].PutValue(20);
        sheet.Cells["A3"].PutValue(30);

        // Add a global named range
        int nameIndex = workbook.Worksheets.Names.Add("MyAbsoluteRange");
        Name namedRange = workbook.Worksheets.Names[nameIndex];

        // Set the RefersTo property using an absolute address (with $ signs) to prevent relative shifts
        // Using SetRefersTo method; the formula must start with '=', isR1C1 = false, isLocal = false
        namedRange.SetRefersTo("=Sheet1!$A$1:$A$3", false, false);
        // Alternatively, you could assign directly:
        // namedRange.RefersTo = "=Sheet1!$A$1:$A$3";

        // Verify the named range works in a formula
        sheet.Cells["B1"].Formula = "=SUM(MyAbsoluteRange)";
        workbook.CalculateFormula();

        Console.WriteLine("Named range RefersTo: " + namedRange.RefersTo);
        Console.WriteLine("SUM result: " + sheet.Cells["B1"].Value);

        // Save the workbook
        workbook.Save("AbsoluteNamedRange.xlsx");
    }
}
