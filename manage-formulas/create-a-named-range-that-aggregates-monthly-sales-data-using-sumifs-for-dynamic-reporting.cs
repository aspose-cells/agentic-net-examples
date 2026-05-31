using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class NamedRangeSumIfsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle: create)
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample sales data
                // Column A: Date, Column B: Sales amount
                cells["A1"].PutValue("Date");
                cells["B1"].PutValue("Sales");

                DateTime startDate = new DateTime(2023, 1, 1);
                Random rnd = new Random();

                // Add 30 days of data
                for (int i = 0; i < 30; i++)
                {
                    cells[i + 1, 0].PutValue(startDate.AddDays(i)); // Date
                    cells[i + 1, 1].PutValue(rnd.Next(100, 1000)); // Sales
                }

                // Cell D1 will hold the month number for which we want the total
                cells["D1"].PutValue(1); // Default to January

                // Create a named range that points to the month input cell
                int monthInputIdx = workbook.Worksheets.Names.Add("MonthInput");
                workbook.Worksheets.Names[monthInputIdx].RefersTo = "=Sheet1!$D$1";

                // Define the range of dates and sales to be used in SUMIFS
                string dateRange = "Sheet1!$A$2:$A$31";
                string salesRange = "Sheet1!$B$2:$B$31";

                // Create a named range "MonthlySales" that aggregates sales for the month
                // specified in MonthInput using SUMIFS.
                int monthlySalesIdx = workbook.Worksheets.Names.Add("MonthlySales");
                workbook.Worksheets.Names[monthlySalesIdx].RefersTo =
                    $"=SUMIFS({salesRange}, {dateRange}, \">=\"&DATE(2023,MonthInput,1), {dateRange}, \"<\"&EDATE(DATE(2023,MonthInput,1),1))";

                // Use the named range in a cell to display the aggregated total
                cells["E1"].Formula = "=MonthlySales";

                // Calculate all formulas (including the dynamic SUMIFS)
                workbook.CalculateFormula();

                // Output the result to console (optional verification)
                Console.WriteLine($"Aggregated sales for month {cells["D1"].IntValue}: {cells["E1"].Value}");

                // Save the workbook (lifecycle: save)
                string outputPath = "MonthlySalesNamedRange.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            NamedRangeSumIfsDemo.Run();
        }
    }
}