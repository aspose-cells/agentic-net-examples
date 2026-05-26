using System;
using Aspose.Cells;

namespace AsposeCellsSumIfsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample sales data
            // Column A: Sale Date
            // Column B: Sale Amount
            cells["A1"].PutValue("Date");
            cells["B1"].PutValue("Sales");

            DateTime startDate = new DateTime(2023, 1, 1);
            Random rnd = new Random();

            // Generate 60 days of data (approx. 2 months)
            for (int i = 0; i < 60; i++)
            {
                cells[i + 1, 0].PutValue(startDate.AddDays(i)); // Date
                cells[i + 1, 1].PutValue(rnd.Next(100, 1000)); // Random sales amount
            }

            // Add a cell where the user can specify the month number (1‑12)
            cells["D1"].PutValue("Month (1‑12)");
            cells["D2"].PutValue(3); // Example: March

            // Build a SUMIFS formula that aggregates sales for the month specified in D2
            // Criteria:
            //   Date >= first day of the month
            //   Date < first day of the next month
            // The DATE function constructs the dates dynamically based on the year of TODAY()
            string sumIfsFormula = 
                "=SUMIFS(B:B, A:A, \">=\" & DATE(YEAR(TODAY()), D2, 1), " +
                "A:A, \"<\" & DATE(YEAR(TODAY()), D2+1, 1))";

            // Place the formula in E2 and calculate it
            cells["E1"].PutValue("Monthly Sales Total");
            cells["E2"].Formula = sumIfsFormula;

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Output the result to the console for verification
            Console.WriteLine($"Total sales for month {cells["D2"].IntValue}: {cells["E2"].Value}");

            // Save the workbook
            workbook.Save("MonthlySalesReport.xlsx");
        }
    }
}