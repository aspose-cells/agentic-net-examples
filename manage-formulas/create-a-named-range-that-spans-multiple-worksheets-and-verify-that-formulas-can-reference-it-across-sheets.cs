// Title: Aspose.Cells .NET – Create a Multi‑Sheet Named Range and Sum It
// Description: Demonstrates how to build a workbook with three worksheets, define a named range that spans Sheet1!A1:A2 and Sheet2!B1:B2, enumerate its separate areas with GetRanges(), apply =SUM(MultiSheetRange) on a Summary sheet, calculate formulas, and save the file.
// Keywords: Aspose.Cells | .NET | C# | named range | multi sheet range | cross‑sheet reference | SUM formula | GetRanges | Excel automation | formula calculation
// Common Searches: Aspose.Cells create named range across worksheets | reference multi‑sheet named range in formula .NET | GetRanges() Aspose.Cells example | sum values from different sheets using named range | C# Aspose.Cells multi‑sheet range tutorial
// Developer Intent: Define a named range that includes cells from multiple worksheets and use it in a formula on another sheet.
// Use Cases: Aggregate data from several detail sheets into a single total on a summary sheet. | Validate each area of a multi‑sheet named range by iterating over the Range objects returned by GetRanges(). | Generate a ready‑to‑export Excel file after formula evaluation for reporting pipelines.
// AI Prompts: Write C# code with Aspose.Cells that creates a named range covering Sheet1!A1:A2 and Sheet2!B1:B2, then places =SUM(MultiSheetRange) in Summary!A1. | Explain how to retrieve individual areas of a multi‑sheet named range using the GetRanges() method in Aspose.Cells. | Show error‑handling techniques when a named range references a missing worksheet while using Aspose.Cells.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsMultiSheetNamedRange
{
    // Demonstrates how to build a workbook with three worksheets, define a named range that spans Sheet1!A1:A2 and Sheet2!B1:B2, enumerate its separate areas with GetRanges(), apply =SUM(MultiSheetRange) on a Summary sheet, calculate formulas, and save the file.
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

                // Add a third worksheet to hold verification formulas
                Worksheet summary = workbook.Worksheets.Add("Summary");

                // Populate data in Sheet1 (A1:A2)
                sheet1.Cells["A1"].PutValue(10);
                sheet1.Cells["A2"].PutValue(20);

                // Populate data in Sheet2 (B1:B2)
                sheet2.Cells["B1"].PutValue(5);
                sheet2.Cells["B2"].PutValue(15);

                // Create a named range that spans both worksheets
                // The RefersTo string can contain multiple areas separated by commas
                int nameIndex = workbook.Worksheets.Names.Add("MultiSheetRange");
                Name multiSheetName = workbook.Worksheets.Names[nameIndex];
                multiSheetName.RefersTo = "=Sheet1!$A$1:$A$2,Sheet2!$B$1:$B$2";

                // Verify the named range consists of two separate areas
                AsposeRange[] areas = multiSheetName.GetRanges();
                Console.WriteLine($"Named range 'MultiSheetRange' consists of {areas.Length} area(s):");
                foreach (AsposeRange area in areas)
                {
                    // RefersTo of each area returns the absolute reference string
                    Console.WriteLine($"  Area RefersTo: {area.RefersTo}");
                }

                // Use the named range in a formula on the Summary sheet
                summary.Cells["A1"].Formula = "=SUM(MultiSheetRange)";

                // Calculate all formulas in the workbook
                workbook.CalculateFormula();

                // Output the result of the formula
                Console.WriteLine($"Result of SUM(MultiSheetRange) in Summary!A1: {summary.Cells["A1"].Value}");

                // Save the workbook
                workbook.Save("MultiSheetNamedRange.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
