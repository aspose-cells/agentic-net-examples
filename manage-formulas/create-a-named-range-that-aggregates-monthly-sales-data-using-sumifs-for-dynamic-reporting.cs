// Title: Aspose.Cells .NET: Create Named Ranges and Apply SUMIFS for Monthly Sales Aggregation
// Description: This example builds a new workbook, populates columns A (Date) and B (Sales) with sample data, defines the named ranges SalesDates and SalesValues, inserts a SUMIFS formula in cell D1 to total January 2023 sales, evaluates the formula, prints the result, and saves the file as MonthlySalesAggregation.xlsx.
// Keywords: Aspose.Cells | C# | .NET | named range | SUMIFS | monthly sales aggregation | dynamic reporting | Excel formula calculation | workbook automation | sales data analysis
// Common Searches: Aspose.Cells define named ranges C# | SUMIFS formula with Aspose.Cells .NET | how to total sales by month using Aspose.Cells | C# example for dynamic sales report in Excel | calculate monthly totals with named ranges in Aspose.Cells
// Developer Intent: Create named ranges for date and sales columns and compute a month‑specific total using a SUMIFS formula in Aspose.Cells.
// Use Cases: Generate a January sales total by applying SUMIFS to the SalesValues range filtered by the SalesDates range. | Reuse the same named ranges to calculate totals for any month by adjusting the date criteria in the formula. | Automate the creation of a ready‑to‑share sales workbook that updates automatically when source data changes.
// AI Prompts: Write C# code with Aspose.Cells that defines named ranges for columns A and B and adds a SUMIFS formula to sum sales for March 2023. | Show how to modify the SUMIFS formula so the start and end dates are read from cells instead of being hard‑coded. | Explain how to loop through a list of months and write each month’s total to consecutive cells using the defined named ranges.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This example builds a new workbook, populates columns A (Date) and B (Sales) with sample data, defines the named ranges SalesDates and SalesValues, inserts a SUMIFS formula in cell D1 to total January 2023 sales, evaluates the formula, prints the result, and saves the file as MonthlySalesAggregation.xlsx.
    public class MonthlySalesAggregation
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // ------------------------------------------------------------
                // Populate sample sales data
                // Column A: Date (yyyy-MM-dd)
                // Column B: Sales amount
                // ------------------------------------------------------------
                cells["A1"].PutValue("Date");
                cells["B1"].PutValue("Sales");

                // Sample data for two months
                DateTime startDate = new DateTime(2023, 1, 1);
                Random rnd = new Random();
                for (int i = 0; i < 60; i++) // 60 days of data
                {
                    cells[i + 1, 0].PutValue(startDate.AddDays(i));          // Date in column A
                    cells[i + 1, 1].PutValue(rnd.Next(100, 1000));          // Random sales in column B
                }

                // ------------------------------------------------------------
                // Create named ranges for the data columns
                // "SalesDates" -> Sheet1!$A$2:$A$61
                // "SalesValues" -> Sheet1!$B$2:$B$61
                // ------------------------------------------------------------
                int datesNameIndex = sheet.Workbook.Worksheets.Names.Add("SalesDates");
                sheet.Workbook.Worksheets.Names[datesNameIndex].RefersTo = $"={sheet.Name}!$A$2:$A$61";

                int valuesNameIndex = sheet.Workbook.Worksheets.Names.Add("SalesValues");
                sheet.Workbook.Worksheets.Names[valuesNameIndex].RefersTo = $"={sheet.Name}!$B$2:$B$61";

                // ------------------------------------------------------------
                // Insert a formula that aggregates sales for January 2023 using SUMIFS
                // ------------------------------------------------------------
                // Cell D1 will hold the result
                cells["D1"].Formula = "=SUMIFS(SalesValues, SalesDates, \">=DATE(2023,1,1)\", SalesDates, \"<=DATE(2023,1,31)\")";

                // Calculate all formulas in the workbook
                workbook.CalculateFormula();

                // Output the aggregated result to the console
                Console.WriteLine("Total sales for January 2023: " + cells["D1"].Value);

                // ------------------------------------------------------------
                // Save the workbook
                // ------------------------------------------------------------
                workbook.Save("MonthlySalesAggregation.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            MonthlySalesAggregation.Run();
        }
    }
}
