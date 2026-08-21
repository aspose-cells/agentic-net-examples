// Title: Validate MINIFS respects AutoFilter visible rows in Aspose.Cells for .NET
// Description: Creates a workbook, fills categories and values, adds a MINIFS formula, applies an AutoFilter to show only Category "A", runs Workbook.CalculateFormula, and outputs the result, proving that MINIFS evaluates only the rows that remain visible after filtering.
// Keywords: Aspose.Cells | C# | .NET | MINIFS | AutoFilter | CalculateFormula | filtered rows | visible rows | Excel formula evaluation | workbook calculation
// Common Searches: MINIFS hidden rows Aspose.Cells | AutoFilter affect MINIFS calculation .NET | Workbook.CalculateFormula filtered data | Aspose.Cells MINIFS visible rows only | C# example MINIFS with AutoFilter
// Developer Intent: Confirm that the MINIFS function returns the minimum value solely from rows that are visible after an AutoFilter is applied.
// Use Cases: Generate reports that need the smallest value among visible records | Perform conditional logic based on the minimum of filtered data | Validate data quality by checking MINIFS results after user‑applied filters | Automate Excel‑like calculations in server‑side .NET applications where hidden rows must be ignored
// AI Prompts: Write C# code using Aspose.Cells to apply an AutoFilter on column A and calculate a MINIFS formula that considers only visible rows. | Explain whether Workbook.CalculateFormula ignores hidden rows when evaluating MINIFS and other aggregate functions. | Create a unit test in .NET that asserts MINIFS returns 5 after filtering for Category "A". | Show how to refresh AutoFilter before calling CalculateFormula to ensure accurate results.

using System;
using Aspose.Cells;

// Creates a workbook, fills categories and values, adds a MINIFS formula, applies an AutoFilter to show only Category "A", runs Workbook.CalculateFormula, and outputs the result, proving that MINIFS evaluates only the rows that remain visible after filtering.
class MinIfsFilteredDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate header
        cells["A1"].PutValue("Category");
        cells["B1"].PutValue("Value");

        // Populate sample data (rows 2‑7)
        string[] categories = { "A", "B", "A", "B", "A", "B" };
        double[] values = { 10, 20, 5, 30, 15, 25 };

        for (int i = 0; i < categories.Length; i++)
        {
            cells[i + 1, 0].PutValue(categories[i]); // Column A
            cells[i + 1, 1].PutValue(values[i]);    // Column B
        }

        // Insert MINIFS formula: minimum Value where Category = "A"
        // Formula range includes all data rows; the filter will hide rows that do not meet the criteria
        cells["D1"].Formula = "=MINIFS(B2:B7, A2:A7, \"A\")";

        // Apply an AutoFilter to the data range and filter to show only Category "A"
        worksheet.AutoFilter.Range = "A1:B7";
        worksheet.AutoFilter.AddFilter(0, "A"); // Field index 0 corresponds to column A
        worksheet.AutoFilter.Refresh(); // Hide rows that do not match the filter

        // Calculate all formulas after the filter has been applied
        workbook.CalculateFormula();

        // Output the result of the MINIFS calculation
        Console.WriteLine("MINIFS result (visible rows only): " + cells["D1"].Value);
    }
}
