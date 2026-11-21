// Title: Create a monthly sales summary workbook with SUMIFS formulas using Aspose.Cells for .NET (C#)
// Common Searches: how to use Aspose.Cells in C# to add a SUMIFS formula that totals sales per month | C# Aspose.Cells example for dynamic monthly sales aggregation with EOMONTH | generate Excel workbook with sales data and monthly summary using Aspose.Cells .NET | apply short date format to date columns in an Aspose.Cells C# workbook | auto‑fit columns and save Excel file after inserting formulas with Aspose.Cells
// Tags: aspocells sumifs monthly aggregation c# | c# create excel workbook with multiple worksheets aspocells | dynamic date range formula eomonth aspocells | apply short date style aspocells c# | auto fit columns aspocells workbook

using System;
using Aspose.Cells;

namespace AsposeCellsSumIfsExample
{
    // Demonstrates building an Excel workbook in C# with Aspose.Cells, populating a sales data sheet, adding a monthly summary sheet, inserting SUMIFS formulas that sum sales for each month using dynamic date ranges (>= month start and <= EOMONTH), formatting dates, auto‑fitting columns, and saving the file as MonthlySalesReport.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet (data sheet)
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "SalesData";

                // Add headers
                dataSheet.Cells["A1"].PutValue("Date");
                dataSheet.Cells["B1"].PutValue("Sales");

                // Sample sales data (Date, Sales)
                var salesData = new (DateTime date, double amount)[]
                {
                    (new DateTime(2023, 1, 5), 1200.50),
                    (new DateTime(2023, 1, 15), 850.75),
                    (new DateTime(2023, 2, 3), 970.00),
                    (new DateTime(2023, 2, 20), 1130.25),
                    (new DateTime(2023, 3, 10), 1420.40),
                    (new DateTime(2023, 3, 28), 760.60)
                };

                // Populate the data sheet
                for (int i = 0; i < salesData.Length; i++)
                {
                    int row = i + 2; // Data starts from row 2 (1‑based)
                    dataSheet.Cells[row - 1, 0].PutValue(salesData[i].date);
                    dataSheet.Cells[row - 1, 1].PutValue(salesData[i].amount);
                }

                // Format the Date column as short date
                Style dateStyle = workbook.CreateStyle();
                dateStyle.Number = 14; // Short date format
                StyleFlag dateFlag = new StyleFlag { NumberFormat = true };
                // CreateRange uses zero‑based indices: firstRow, firstColumn, totalRows, totalColumns
                dataSheet.Cells.CreateRange(1, 0, salesData.Length, 1).ApplyStyle(dateStyle, dateFlag);

                // Add a second worksheet for the monthly summary
                Worksheet summarySheet = workbook.Worksheets.Add("MonthlySummary");

                // Header for month and total sales
                summarySheet.Cells["A1"].PutValue("Month Start");
                summarySheet.Cells["B1"].PutValue("Total Sales");

                // Define the months you want to summarize (e.g., Jan to Mar 2023)
                DateTime[] monthStarts = new DateTime[]
                {
                    new DateTime(2023, 1, 1),
                    new DateTime(2023, 2, 1),
                    new DateTime(2023, 3, 1)
                };

                // Populate month start dates and set SUMIFS formulas
                for (int i = 0; i < monthStarts.Length; i++)
                {
                    int row = i + 2; // Data starts from row 2 (1‑based)

                    // Put the month start date
                    summarySheet.Cells[row - 1, 0].PutValue(monthStarts[i]);

                    // Build the SUMIFS formula:
                    // =SUMIFS(SalesData!B:B, SalesData!A:A, ">=MonthStart", SalesData!A:A, "<=EOMONTH(MonthStart,0)")
                    string monthStartCell = $"A{row}";
                    string formula = $"=SUMIFS(SalesData!B:B, SalesData!A:A, \">=\" & {monthStartCell}, SalesData!A:A, \"<=\" & EOMONTH({monthStartCell},0))";

                    summarySheet.Cells[row - 1, 1].Formula = formula;
                }

                // Format the Month Start column as short date
                summarySheet.Cells.CreateRange(1, 0, monthStarts.Length, 1).ApplyStyle(dateStyle, dateFlag);

                // Auto‑fit columns for better readability
                summarySheet.AutoFitColumn(0);
                summarySheet.AutoFitColumn(1);
                dataSheet.AutoFitColumn(0);
                dataSheet.AutoFitColumn(1);

                // Save the workbook to a file
                string outputPath = "MonthlySalesReport.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
