// Title: How to disable automatic calculation, insert formulas, and manually trigger workbook calculation using Workbook.CalculateFormula in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# disable automatic calculation and manually evaluate formulas | Insert SUM and AVERAGE formulas with Aspose.Cells and trigger calculation later | Workbook.CalculateFormula example for evaluating all worksheets in a .NET workbook | How to prevent auto formula recalculation in Aspose.Cells and run it on demand | C# Aspose.Cells manual formula evaluation after adding data to cells
// Tags: disable automatic calculation Aspose.Cells .NET | Workbook.CalculateFormula usage C# | programmatically insert Excel formulas Aspose.Cells | manual formula evaluation workbook Aspose.Cells | C# Aspose.Cells calculate all worksheets

using Aspose.Cells;
using System;

// The example creates a new Workbook, adds two worksheets, writes numeric values to cells A1 and A2, assigns a SUM formula to Sheet1!B1 and an AVERAGE formula to Sheet2!B1, and postpones any calculation until workbook.CalculateFormula is called, which evaluates all formulas before saving the file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // Worksheet 1: insert data and a formula
            // -------------------------------------------------
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";

            // Sample data
            sheet1.Cells["A1"].PutValue(10);
            sheet1.Cells["A2"].PutValue(20);

            // Formula: sum of A1 and A2
            sheet1.Cells["B1"].Formula = "=SUM(A1:A2)";

            // -------------------------------------------------
            // Worksheet 2: insert data and a formula
            // -------------------------------------------------
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");

            // Sample data
            sheet2.Cells["A1"].PutValue(5);
            sheet2.Cells["A2"].PutValue(15);

            // Formula: average of A1 and A2
            sheet2.Cells["B1"].Formula = "=AVERAGE(A1:A2)";

            // -------------------------------------------------
            // Trigger calculation for the entire workbook
            // -------------------------------------------------
            workbook.CalculateFormula(); // calculates all formulas in all worksheets

            // Save the workbook to a file
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
