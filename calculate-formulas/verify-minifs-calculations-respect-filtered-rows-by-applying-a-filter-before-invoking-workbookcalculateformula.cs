// Title: Ensuring MINIFS formula respects AutoFilter visibility when recalculating with Workbook.CalculateFormula in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# calculate MINIFS only on visible rows after applying AutoFilter | How to recalculate formulas on filtered data using Workbook.CalculateFormula in Aspose.Cells | C# MINIFS function returns hidden values with Aspose.Cells AutoFilter applied | Apply AutoFilter then evaluate MINIFS with Aspose.Cells .NET example | Workbook.CalculateFormula ignores AutoFilter rows Aspose.Cells tutorial
// Tags: MINIFS formula with AutoFilter Aspose.Cells | Workbook.CalculateFormula filtered rows | Aspose.Cells evaluate MINIFS on visible data | C# auto filter formula recalculation | Aspose.Cells filtered data calculation example

using System;
using Aspose.Cells;

// The sample creates a workbook, fills columns A and B with categories and numeric values, applies an AutoFilter to show only rows where the Category equals "A", inserts a MINIFS formula in column C to find the smallest value for that category, recalculates all formulas using Workbook.CalculateFormula, and prints the result, which should be 5, demonstrating that the MINIFS calculation respects the filtered (visible) rows.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Fill sample data
            // Headers
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            // Data rows
            sheet.Cells["A2"].PutValue("A"); sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B"); sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("A"); sheet.Cells["B4"].PutValue(5);
            sheet.Cells["A5"].PutValue("B"); sheet.Cells["B5"].PutValue(15);
            sheet.Cells["A6"].PutValue("A"); sheet.Cells["B6"].PutValue(8);

            // Apply an AutoFilter to the range A1:B6
            sheet.AutoFilter.Range = "A1:B6";

            // Filter to show only rows where Category = "A"
            // Add filter on column 0 (Category) for value "A"
            sheet.AutoFilter.AddFilter(0, "A");
            sheet.AutoFilter.Refresh();

            // Insert a MINIFS formula that calculates the minimum Value for Category "A"
            sheet.Cells["C2"].Formula = "=MINIFS(B:B, A:A, \"A\")";

            // Recalculate all formulas
            workbook.CalculateFormula();

            // Read and display the result; expected value is 5 (the smallest Value for Category "A")
            double minIfsResult = sheet.Cells["C2"].DoubleValue;
            Console.WriteLine("MINIFS result (should be 5): " + minIfsResult);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
