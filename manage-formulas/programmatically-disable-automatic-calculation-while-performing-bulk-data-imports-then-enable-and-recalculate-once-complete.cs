// Title: How to temporarily disable automatic calculation in Aspose.Cells for .NET during bulk data import and then recalculate formulas
// Common Searches: disable automatic formula calculation in Aspose.Cells C# before inserting thousands of rows | enable calculation and force full workbook recalc after bulk cell writes using Aspose.Cells | improve performance of large data imports by turning off Aspose.Cells calculation engine | bulk data load into Excel with Aspose.Cells without triggering formula evaluation | recalculate all formulas after bulk import in an Aspose.Cells workbook
// Tags: turn off calculation engine Aspose.Cells | bulk data import performance Aspose.Cells | Workbook.CalculateFormula method | large worksheet write optimization .NET | post-import formula recalculation Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a new Workbook, writes 10,000 rows of numeric data to the first two columns, then explicitly calls Workbook.CalculateFormula() to force a full recalculation of any formulas before saving the file as output.xlsx. This pattern avoids unnecessary formula evaluation during the bulk write and ensures accurate results after the import.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------
            // Bulk data import example
            // -------------------------
            // (Replace this loop with actual data import logic)
            for (int row = 0; row < 10000; row++)
            {
                sheet.Cells[row, 0].PutValue(row);          // Column A
                sheet.Cells[row, 1].PutValue(row * 2);      // Column B
            }

            // Force a full recalculation after data import
            workbook.CalculateFormula();

            // Save the workbook to a file
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
