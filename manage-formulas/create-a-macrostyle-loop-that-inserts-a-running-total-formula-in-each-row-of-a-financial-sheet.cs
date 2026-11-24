// Title: Create a macro‑style loop in C# with Aspose.Cells to add a running total formula to each row of a financial worksheet
// Common Searches: Aspose.Cells C# loop to insert cumulative SUM formula in column B | how to generate running total column programmatically using Aspose.Cells | C# example adding dynamic SUM(A1:Arow) formulas with Aspose.Cells | automate cumulative totals in Excel workbook using Aspose.Cells for .NET | macro style iteration for financial sheet running total in Aspose.Cells
// Tags: cumulative sum formula Aspose.Cells C# | dynamic range SUM formula workbook | loop insert formulas column B Aspose.Cells | financial worksheet running total automation | calculate running total programmatically Aspose.Cells

using Aspose.Cells;
using System;

// The sample creates a new workbook, populates column A with sequential values, then uses a macro‑style C# loop to set a running total SUM formula in column B for each row (e.g., =SUM($A$1:A5)). After recalculating formulas, the workbook is saved as FinancialRunningTotal.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet and name it "Financial"
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Financial";

        // -------------------------------------------------
        // Sample data: populate column A with amounts
        // (In a real scenario the data would already exist)
        // -------------------------------------------------
        int dataRows = 10; // number of rows with data
        for (int i = 0; i < dataRows; i++)
        {
            // Row i, column 0 (A) -> put a sample value
            sheet.Cells[i, 0].PutValue(i + 1); // values 1,2,3,...
        }

        // -------------------------------------------------
        // Insert a running total formula in column B
        // -------------------------------------------------
        for (int i = 0; i < dataRows; i++)
        {
            // Build the formula for the running total:
            // =SUM($A$1:A{rowNumber})
            // A1‑style formulas use 1‑based row numbers
            string formula = $"=SUM($A$1:A{i + 1})";

            // Set the formula in column B (index 1) of the current row
            sheet.Cells[i, 1].Formula = formula;
        }

        // Recalculate all formulas so the workbook contains the computed values
        workbook.CalculateFormula();

        // Save the workbook to a file
        workbook.Save("FinancialRunningTotal.xlsx");
    }
}
