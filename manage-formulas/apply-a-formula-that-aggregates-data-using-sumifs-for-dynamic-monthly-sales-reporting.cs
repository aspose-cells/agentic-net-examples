// Title: C# – Add a Dynamic SUMIFS Formula for Monthly Sales with Aspose.Cells
// Description: Creates an Excel workbook in C# using Aspose.Cells, populates sample sales data, lets the user pick a month (cell E2), inserts a SUMIFS formula that totals the Sales column for the selected month of the current year, calculates the result, prints it to the console, and saves the file as MonthlySalesReport.xlsx.
// Keywords: Aspose.Cells C# SUMIFS | dynamic monthly sales total | Excel DATE TODAY formula | C# generate sales report | SUMIFS date range Aspose | monthly aggregation Excel | C# workbook automation
// Common Searches: Aspose.Cells SUMIFS month example | C# calculate monthly sales with SUMIFS | dynamic date range formula Aspose.Cells | how to sum sales by month in .NET | Excel SUMIFS with DATE and TODAY in code
// Developer Intent: Embed a SUMIFS formula that automatically sums sales for the month specified by the user, without hard‑coding dates.
// Use Cases: Provide a template where changing the month number instantly updates the total sales. | Automate monthly sales reporting in a .NET application that generates Excel files. | Create a reusable workbook that can be reused across different regions or fiscal years.
// AI Prompts: Show how to extend the formula to also filter by a region entered in another cell. | Generate code that adds a summary table with totals for all 12 months on the same sheet. | Explain how to format the result cell as currency and return zero when no data matches the selected month.

using System;
using Aspose.Cells;

namespace AsposeCellsSumIfsDemo
{
    // Creates an Excel workbook in C# using Aspose.Cells, populates sample sales data, lets the user pick a month (cell E2), inserts a SUMIFS formula that totals the Sales column for the selected month of the current year, calculates the result, prints it to the console, and saves the file as MonthlySalesReport.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Header row
            cells["A1"].PutValue("Date");
            cells["B1"].PutValue("Region");
            cells["C1"].PutValue("Sales");

            // Sample data spanning several months
            DateTime startDate = new DateTime(2023, 1, 1);
            string[] regions = { "North", "South", "East", "West" };
            Random rnd = new Random();

            for (int i = 0; i < 20; i++)
            {
                // Date increments by a few days
                DateTime date = startDate.AddDays(i * 5);
                cells[i + 1, 0].PutValue(date);
                cells[i + 1, 1].PutValue(regions[i % regions.Length]);
                cells[i + 1, 2].PutValue(rnd.Next(1000, 5000));
            }

            // Cell where user selects month (1 = Jan, 2 = Feb, etc.)
            cells["E1"].PutValue("Month");
            cells["E2"].PutValue(3); // Example: March

            // SUMIFS formula to aggregate sales for the selected month of the current year
            // Criteria 1: Date >= first day of the month
            // Criteria 2: Date < first day of the next month
            string sumIfsFormula = "=SUMIFS(C2:C21, A2:A21, \">=\"&DATE(YEAR(TODAY()),$E$2,1), " +
                                   "A2:A21, \"<\"&DATE(YEAR(TODAY()),$E$2+1,1))";

            // Place the formula in cell E3
            cells["E3"].Formula = sumIfsFormula;

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Output the result to console (optional)
            Console.WriteLine("Aggregated sales for month " + cells["E2"].IntValue + ": " + cells["E3"].Value);

            // Save the workbook
            workbook.Save("MonthlySalesReport.xlsx");
        }
    }
}
