// Title: Define a multi‑sheet named range and use it in a SUM formula with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# create named range that spans multiple worksheets | C# sum values across sheets using a named range in Aspose.Cells | Set RefersTo for a multi‑sheet named range in Aspose.Cells .NET | Apply SUM formula to a named range covering Sheet1 and Sheet2 with Aspose.Cells | Calculate workbook formulas after defining a cross‑sheet named range in Aspose.Cells C#
// Tags: Aspose.Cells multi-sheet named range | Aspose.Cells cross-sheet SUM formula | Aspose.Cells RefersTo multiple worksheets | Aspose.Cells CalculateFormula example | Aspose.Cells named range across worksheets

using System;
using Aspose.Cells;

// The example creates a new workbook, adds two worksheets (Sheet1 and Sheet2), fills cells A1:A2 on each sheet with numbers, defines a named range "MyRange" that references A1:A2 on both sheets via the RefersTo property, inserts a SUM formula that uses this named range into cell B1, calculates all formulas, and saves the file as CrossSheetNamedRange.xlsx.
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

            // Populate some sample data in both sheets
            sheet1.Cells["A1"].PutValue(10);
            sheet1.Cells["A2"].PutValue(20);
            sheet2.Cells["A1"].PutValue(30);
            sheet2.Cells["A2"].PutValue(40);

            // Create a named range that spans cells A1:A2 on both worksheets
            // Names.Add returns the index of the newly added name; retrieve the Name object using that index
            int rangeIndex = workbook.Worksheets.Names.Add("MyRange");
            Name multiSheetRange = workbook.Worksheets.Names[rangeIndex];
            multiSheetRange.RefersTo = "Sheet1!A1:A2,Sheet2!A1:A2";

            // Use the named range in a cross‑sheet formula (e.g., sum all values)
            sheet1.Cells["B1"].Formula = "SUM(MyRange)";

            // Calculate formulas so the result is stored in the cell
            workbook.CalculateFormula();

            // Save the workbook to a file
            string outputPath = "CrossSheetNamedRange.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
