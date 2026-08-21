// Title: Define a Multi‑Sheet Named Range and Sum It with Aspose.Cells for .NET
// Description: Creates a workbook with two worksheets, fills numeric data, defines a named range that references cells on both sheets via the RefersTo property, inserts a SUM formula that uses this range, recalculates formulas, outputs the result, and saves the file.
// Keywords: Aspose.Cells multi‑sheet named range | C# named range across worksheets | RefersTo multiple areas | SUM formula with named range | CalculateFormula Aspose.Cells | .NET workbook automation | global
// Common Searches: Aspose.Cells define named range on several sheets | C# sum values from multiple worksheets using a named range | How to use RefersTo for multi‑area named range in Aspose.Cells | Recalculate formulas after adding a named range in .NET | Save workbook with multi‑sheet named range Aspose
// Developer Intent: Create a named range that spans two worksheets and use it in a SUM formula.
// Use Cases: Aggregate totals from different sheets with a single named range. | Build a summary sheet that pulls values from multiple data sheets. | Standardize calculations across generated workbooks by reusing a multi‑sheet named range.
// AI Prompts: Generate C# code with Aspose.Cells to define a named range covering Sheet1!A1:A3 and Sheet2!B1:B3 and apply a SUM formula. | Explain the correct RefersTo syntax for a named range that includes multiple areas in Aspose.Cells. | Show how to recalculate formulas and retrieve the SUM result of a multi‑sheet named range using Aspose.Cells.

using System;
using Aspose.Cells;

// Creates a workbook with two worksheets, fills numeric data, defines a named range that references cells on both sheets via the RefersTo property, inserts a SUM formula that uses this range, recalculates formulas, outputs the result, and saves the file.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet and rename it
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "Sheet1";

        // Add a second worksheet
        Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");

        // Populate data in Sheet1 (A1:A3)
        sheet1.Cells["A1"].PutValue(10);
        sheet1.Cells["A2"].PutValue(20);
        sheet1.Cells["A3"].PutValue(30);

        // Populate data in Sheet2 (B1:B3)
        sheet2.Cells["B1"].PutValue(5);
        sheet2.Cells["B2"].PutValue(15);
        sheet2.Cells["B3"].PutValue(25);

        // Create a named range that spans both worksheets
        // The RefersTo string uses a comma to separate multiple areas
        int nameIndex = workbook.Worksheets.Names.Add("MultiSheetRange");
        Name multiSheetRange = workbook.Worksheets.Names[nameIndex];
        multiSheetRange.RefersTo = "=Sheet1!$A$1:$A$3,Sheet2!$B$1:$B$3";

        // Use the named range in a summary formula (SUM) on Sheet1 cell C1
        sheet1.Cells["C1"].Formula = "=SUM(MultiSheetRange)";

        // Calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Optional: display the calculated result
        Console.WriteLine("Sum of MultiSheetRange: " + sheet1.Cells["C1"].Value);

        // Save the workbook
        workbook.Save("MultiSheetNamedRange.xlsx");
    }
}
