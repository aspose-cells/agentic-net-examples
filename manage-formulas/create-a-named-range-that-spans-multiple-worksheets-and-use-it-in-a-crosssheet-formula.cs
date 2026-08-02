// Title: Define a Multi‑Sheet Named Range and Use It in a SUM Formula – Aspose.Cells for .NET
// Description: C# example that creates a workbook with two sheets, fills data, defines a named range spanning Sheet1!A1:B2 and Sheet2!C3:D4, inserts a SUM formula referencing the range, calculates the result, and saves the file.
// Keywords: Aspose.Cells | C# | .NET | named range | multi‑sheet range | cross‑sheet formula | SUM function | Excel automation | calculate formulas | workbook save
// Common Searches: Aspose.Cells create named range across worksheets | C# multi‑sheet named range example | how to use SUM with a range that spans multiple sheets in Aspose.Cells | calculate cross‑sheet formulas Aspose.Cells .NET | define and reference multi‑sheet range in Excel using Aspose
// Developer Intent: Create a named range that includes cells from several worksheets and reference it in a formula to aggregate the values.
// Use Cases: Consolidate sales figures from department sheets into a single total. | Build financial models that pull data from multiple worksheets with one range name. | Generate summary reports where a single formula updates automatically as source sheets change.
// AI Prompts: Write C# code with Aspose.Cells to define a named range covering Sheet1!A1:B2 and Sheet2!C3:D4 and use it in a SUM formula. | Show how to modify an existing multi‑sheet named range and recalculate dependent formulas in Aspose.Cells. | Explain how to reference a multi‑sheet named range in functions like AVERAGE, COUNT, or MAX using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// C# example that creates a workbook with two sheets, fills data, defines a named range spanning Sheet1!A1:B2 and Sheet2!C3:D4, inserts a SUM formula referencing the range, calculates the result, and saves the file.
class Program
{
    static void Main()
    {
        // Create a new workbook (contains a default Sheet1)
        Workbook wb = new Workbook();

        // Rename the default worksheet to "Sheet1" for clarity
        Worksheet sheet1 = wb.Worksheets[0];
        sheet1.Name = "Sheet1";

        // Add a second worksheet named "Sheet2"
        Worksheet sheet2 = wb.Worksheets.Add("Sheet2");

        // Populate data in Sheet1 range A1:B2
        sheet1.Cells["A1"].PutValue(1);
        sheet1.Cells["A2"].PutValue(2);
        sheet1.Cells["B1"].PutValue(3);
        sheet1.Cells["B2"].PutValue(4);

        // Populate data in Sheet2 range C3:D4
        sheet2.Cells["C3"].PutValue(5);
        sheet2.Cells["C4"].PutValue(6);
        sheet2.Cells["D3"].PutValue(7);
        sheet2.Cells["D4"].PutValue(8);

        // Create a named range that spans both worksheets
        // The RefersTo string must start with '=' and separate areas with commas
        int nameIdx = wb.Worksheets.Names.Add("MultiSheetRange");
        Name multiRange = wb.Worksheets.Names[nameIdx];
        multiRange.RefersTo = "=Sheet1!$A$1:$B$2,Sheet2!$C$3:$D$4";

        // Use the named range in a cross‑sheet formula (sum of all cells in the range)
        // Place the formula in Sheet1 cell E1
        sheet1.Cells["E1"].Formula = "=SUM(MultiSheetRange)";

        // Calculate all formulas in the workbook
        wb.CalculateFormula();

        // Display the calculated result
        Console.WriteLine("Sum of MultiSheetRange = " + sheet1.Cells["E1"].Value);

        // Save the workbook to a file
        wb.Save("MultiSheetNamedRange.xlsx");
    }
}
