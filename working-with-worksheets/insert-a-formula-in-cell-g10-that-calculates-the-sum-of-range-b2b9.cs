// Title: Insert SUM Formula in G10 (B2:B9) Using Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, fills cells B2‑B9 with sample numbers, assigns the formula "=SUM(B2:B9)" to cell G10 via the Cell.Formula property, forces a full recalculation, and saves the result as output.xlsx.
// Keywords: Aspose.Cells C# formula | set cell formula Aspose.Cells | SUM function .NET | calculate formulas Aspose.Cells | save workbook C# | programmatic Excel formula
// Common Searches: how to add a SUM formula with Aspose.Cells C# | set formula in a specific cell using Aspose.Cells | recalculate workbook after inserting formulas Aspose.Cells | Aspose.Cells example insert formula G10
// Developer Intent: Add a SUM formula to G10 that totals B2:B9, recalculate the workbook, and save the file.
// Use Cases: Generate a totals row for financial data after loading transaction values. | Create summary calculations automatically when importing CSV data into Excel. | Prepare aggregated metrics in a report workbook before distribution.
// AI Prompts: Write C# code that inserts an AVERAGE formula for range C2:C10 with Aspose.Cells and triggers calculation. | Explain how to apply the same formula to multiple cells dynamically using Aspose.Cells for .NET. | Show how to use absolute and relative references when setting formulas programmatically with Aspose.Cells.

using Aspose.Cells;

// Creates a new workbook, fills cells B2‑B9 with sample numbers, assigns the formula "=SUM(B2:B9)" to cell G10 via the Cell.Formula property, forces a full recalculation, and saves the result as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate B2:B9 with sample data (optional, just for demonstration)
        for (int i = 0; i < 8; i++)
        {
            // B2 is row index 1, column index 1 (zero‑based)
            worksheet.Cells[i + 1, 1].PutValue(i + 1);
        }

        // Insert the SUM formula into cell G10 (uses Cell.Formula property)
        worksheet.Cells["G10"].Formula = "=SUM(B2:B9)";

        // Recalculate all formulas in the workbook
        workbook.CalculateFormula();

        // Save the workbook to disk (lifecycle: save)
        workbook.Save("output.xlsx");
    }
}
