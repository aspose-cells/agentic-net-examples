// Title: Generate a monthly sales summary worksheet by grouping dates by month and year with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells group Excel rows by month and year for sales report | How to create a summary sheet with total sales per month using Aspose.Cells | Aspose.Cells .NET example to aggregate sales data by month | Export Excel range to DataTable and summarize sales in C# | Add a new worksheet with monthly totals using Aspose.Cells API
// Tags: group rows by month year Aspose.Cells C# | create summary worksheet from DataTable Aspose.Cells | aggregate sales totals per month Aspose.Cells | export Excel range to DataTable Aspose.Cells | add new worksheet with formatted month Aspose.Cells | pivot-like monthly sales summary Aspose.Cells

using Aspose.Cells;
using System;
using System.Data;
using System.IO;
using System.Linq;

// The example loads an existing workbook (or creates a sample one), exports the data range to a DataTable, groups rows by the Year and Month of the "Date" column, sums the "Sales" values for each group, adds a new "Summary" worksheet, writes the month/year label and total sales, and saves the workbook with the generated summary.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            Workbook workbook;

            if (File.Exists(inputPath))
            {
                // Load existing workbook
                workbook = new Workbook(inputPath);
            }
            else
            {
                // Create a sample workbook with required columns if input file is missing
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                Cells cells = ws.Cells;
                cells["A1"].PutValue("Date");
                cells["B1"].PutValue("Sales");
                // Sample data rows
                cells["A2"].PutValue(DateTime.Today);
                cells["B2"].PutValue(1000);
                cells["A3"].PutValue(DateTime.Today.AddMonths(-1));
                cells["B3"].PutValue(1500);
            }

            // Access first worksheet containing raw sales data
            Worksheet dataSheet = workbook.Worksheets[0];
            Cells dataCells = dataSheet.Cells;

            // Determine used range (including header row)
            int lastRow = dataCells.MaxDataRow;
            int lastCol = dataCells.MaxDataColumn;

            // Export the range to a DataTable (first row contains column names)
            DataTable dt = dataCells.ExportDataTable(0, 0, lastRow + 1, lastCol + 1, true);

            // Group rows by Year and Month of the "Date" column and sum the "Sales" column
            var grouped = dt.AsEnumerable()
                .GroupBy(r => new
                {
                    Year = r.Field<DateTime>("Date").Year,
                    Month = r.Field<DateTime>("Date").Month
                })
                .Select(g => new
                {
                    Year = g.Key.Year,
                    Month = g.Key.Month,
                    TotalSales = g.Sum(r => Convert.ToDouble(r["Sales"]))
                })
                .OrderBy(x => x.Year)
                .ThenBy(x => x.Month)
                .ToList();

            // Add a new worksheet to hold the summary
            Worksheet summarySheet = workbook.Worksheets.Add("Summary");
            Cells summaryCells = summarySheet.Cells;

            // Write header row
            summaryCells["A1"].PutValue("Month");
            summaryCells["B1"].PutValue("Year");
            summaryCells["C1"].PutValue("Total Sales");

            // Populate the summary data
            int row = 1; // zero‑based index; row 1 is the second row in the sheet
            foreach (var item in grouped)
            {
                // Format month as MM/yyyy for readability
                summaryCells[row, 0].PutValue($"{item.Month:D2}/{item.Year}");
                summaryCells[row, 1].PutValue(item.Year);
                summaryCells[row, 2].PutValue(item.TotalSales);
                row++;
            }

            // Save the workbook with the new summary sheet
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
