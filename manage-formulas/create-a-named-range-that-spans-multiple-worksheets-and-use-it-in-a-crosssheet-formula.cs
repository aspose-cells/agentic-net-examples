// Title: How to create a named range that spans multiple worksheets and use it in a cross‑sheet SUM formula with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells to define a named range covering Sheet1!A1:A3 and Sheet2!B1:B3, then place a =SUM(MyRange) formula on a Summary worksheet and evaluate the result. | Demonstrate the steps to add a multi‑sheet named range, assign its RefersTo string, recalculate all formulas, and save the workbook as an .xlsx file with Aspose.Cells.
// Common Searches: Aspose.Cells C# create named range across two worksheets | C# Aspose.Cells sum values from multiple sheets using a named range | How to set RefersTo for a named range that includes cells on different worksheets in Aspose.Cells | Calculate workbook formulas after adding a cross‑sheet named range with Aspose.Cells .NET | Save workbook after using multi‑sheet named range in Aspose.Cells
// Tags: Aspose.Cells multi‑worksheet named range | Aspose.Cells cross‑sheet SUM formula | Aspose.Cells RefersTo string | Aspose.Cells workbook calculation | Aspose.Cells .xlsx file saving

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, adds three worksheets (Sheet1, Sheet2, Summary), fills Sheet1!A1:A3 and Sheet2!B1:B3 with numbers, defines a named range called MyRange that references both areas via a comma‑separated RefersTo string, inserts a =SUM(MyRange) formula on the Summary sheet, calculates all formulas, prints the sum, and saves the file as NamedRangeCrossSheet.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Rename the default sheet to Sheet1
            Worksheet sheet1 = wb.Worksheets[0];
            sheet1.Name = "Sheet1";

            // Add a second worksheet named Sheet2
            Worksheet sheet2 = wb.Worksheets.Add("Sheet2");

            // Add a third worksheet for the summary formula
            Worksheet summary = wb.Worksheets.Add("Summary");

            // Populate Sheet1!A1:A3 with sample data
            sheet1.Cells["A1"].PutValue(10);
            sheet1.Cells["A2"].PutValue(20);
            sheet1.Cells["A3"].PutValue(30);

            // Populate Sheet2!B1:B3 with sample data
            sheet2.Cells["B1"].PutValue(5);
            sheet2.Cells["B2"].PutValue(15);
            sheet2.Cells["B3"].PutValue(25);

            // Define a named range that spans both worksheets
            // The reference string uses a comma to separate the two areas
            string refString = "Sheet1!A1:A3,Sheet2!B1:B3";

            // Add the named range and obtain the Name object
            int nameIndex = wb.Worksheets.Names.Add("MyRange");
            Name myRange = wb.Worksheets.Names[nameIndex];
            myRange.RefersTo = refString;

            // Use the named range in a cross‑sheet formula on the Summary sheet
            // Example: sum all values in the named range
            summary.Cells["A1"].Formula = "=SUM(MyRange)";

            // Calculate formulas in the workbook
            wb.CalculateFormula();

            // Output the result to the console
            Console.WriteLine("Sum of MyRange: " + summary.Cells["A1"].StringValue);

            // Save the workbook to a file (ensure the directory exists)
            string outputPath = "NamedRangeCrossSheet.xlsx";
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
