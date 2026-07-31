// Title: Aspose.Cells .NET: Create and Use a Multi‑Sheet Named Range in C#
// Description: Demonstrates how to build a workbook, add data on two sheets, define a named range that spans both worksheets, retrieve its individual areas, place a SUM formula on a third sheet that references the multi‑sheet range, calculate the workbook, and save the result.
// Keywords: Aspose.Cells | C# named range multiple sheets | multi‑area named range | SUM formula across worksheets | calculate formulas Aspose.Cells | .NET spreadsheet automation | Excel named range spanning sheets
// Common Searches: Aspose.Cells create named range across worksheets | reference multi‑sheet named range in formula .NET | how to sum values from several sheets using a named range | retrieve areas of a multi‑area named range Aspose.Cells | calculate workbook after adding named range
// Developer Intent: Define a named range that includes cells from different worksheets and confirm that formulas can reference and sum it correctly.
// Use Cases: Consolidate regional sales figures stored on separate sheets into a single named range for summary calculations. | Programmatically iterate over each area of a multi‑sheet named range to generate custom reports. | Persist calculated totals in an Excel file that combines data from multiple sources.
// AI Prompts: Write C# code with Aspose.Cells to create a named range covering Sheet1!A1:A5 and Sheet2!B1:B5, then add a formula on Sheet3 that returns the average of the range. | Explain how Aspose.Cells parses a comma‑separated RefersTo string for a multi‑area named range and how to access each area via the API. | Provide troubleshooting steps when a SUM formula returns #NAME? or incorrect results while referencing a multi‑sheet named range.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Demonstrates how to build a workbook, add data on two sheets, define a named range that spans both worksheets, retrieve its individual areas, place a SUM formula on a third sheet that references the multi‑sheet range, calculate the workbook, and save the result.
class MultiSheetNamedRangeDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Add two worksheets that will contain the data for the named range
            Worksheet sheet1 = wb.Worksheets[0];
            sheet1.Name = "Data1";
            Worksheet sheet2 = wb.Worksheets.Add("Data2");

            // Populate Sheet1!A1:A2
            sheet1.Cells["A1"].PutValue(10);
            sheet1.Cells["A2"].PutValue(20);

            // Populate Sheet2!B1:B2
            sheet2.Cells["B1"].PutValue(30);
            sheet2.Cells["B2"].PutValue(40);

            // Create a named range that spans both worksheets
            int nameIdx = wb.Worksheets.Names.Add("MultiRange");
            Name multiName = wb.Worksheets.Names[nameIdx];
            // The RefersTo string can contain multiple areas separated by commas
            multiName.RefersTo = "=Data1!$A$1:$A$2,Data2!$B$1:$B$2";

            // Retrieve the individual ranges that compose the multi‑sheet named range
            AsposeRange[] ranges = multiName.GetRanges();
            Console.WriteLine("Named range 'MultiRange' consists of {0} areas:", ranges.Length);
            foreach (AsposeRange r in ranges)
            {
                Console.WriteLine(" - Sheet: {0}, Address: {1}", r.Worksheet.Name, r.Address);
            }

            // Add a third worksheet to test formulas that reference the multi‑sheet named range
            Worksheet calcSheet = wb.Worksheets.Add("Calc");
            // Use the named range in a SUM formula
            calcSheet.Cells["A1"].Formula = "=SUM(MultiRange)";

            // Calculate all formulas in the workbook
            wb.CalculateFormula();

            // Output the result of the formula
            Console.WriteLine("Result of SUM(MultiRange) on Calc!A1: " + calcSheet.Cells["A1"].Value);

            // Save the workbook to verify the result in Excel
            wb.Save("MultiSheetNamedRangeDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
