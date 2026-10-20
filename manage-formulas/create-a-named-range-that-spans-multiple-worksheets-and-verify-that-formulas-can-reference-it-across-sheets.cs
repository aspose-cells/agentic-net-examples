// Title: Define a multi‑sheet named range and sum it with a formula using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# create named range that includes cells from multiple worksheets | C# Aspose.Cells sum values across sheets using a named range | Set RefersTo property for a multi‑sheet range in Aspose.Cells .NET | Calculate formula that references a named range spanning several worksheets with Aspose.Cells | Add multi‑sheet named range and use SUM function in Aspose.Cells workbook
// Tags: define multi‑sheet named range Aspose.Cells | SUM formula referencing named range C# Aspose.Cells | Workbook.Worksheets.Names.Add cross‑sheet range | RefersTo property multi‑sheet range Aspose.Cells | calculate formulas after adding named range Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a workbook with two worksheets, fills A1:A2 on each sheet, defines a named range that spans those cells across both sheets, inserts a SUM formula that references the multi‑sheet named range, calculates all formulas, outputs the result, and saves the file as MultiSheetNamedRange.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and rename it
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";

            // Add a second worksheet
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");

            // Populate some numeric data in both sheets
            sheet1.Cells["A1"].PutValue(10);
            sheet1.Cells["A2"].PutValue(20);
            sheet2.Cells["A1"].PutValue(30);
            sheet2.Cells["A2"].PutValue(40);

            // Define a named range that spans A1:A2 on both worksheets
            string refersTo = "Sheet1!$A$1:$A$2,Sheet2!$A$1:$A$2";

            // Add the named range (Names.Add returns the index of the new name)
            int nameIndex = workbook.Worksheets.Names.Add("MyMultiSheetRange");
            Name rangeName = workbook.Worksheets.Names[nameIndex];
            rangeName.RefersTo = refersTo;

            // Insert a formula that references the multi‑sheet named range
            sheet1.Cells["B1"].Formula = "SUM(MyMultiSheetRange)";

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Output the calculated result to the console
            Console.WriteLine("Result of SUM(MyMultiSheetRange): " + sheet1.Cells["B1"].StringValue);

            // Save the workbook
            string outputPath = "MultiSheetNamedRange.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
