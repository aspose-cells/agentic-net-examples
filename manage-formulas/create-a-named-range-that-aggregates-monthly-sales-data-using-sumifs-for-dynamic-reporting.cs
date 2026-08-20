// Title: Aspose.Cells for .NET: Create a SUMIFS Named Range to Aggregate Monthly Sales (C#)
// Description: This C# example shows how to build a workbook with date, month, year, and sales columns, define column‑based named ranges (SalesData, MonthData, YearData), and create a fourth named range (MonthlySales) that uses the SUMIFS function to total sales for a month and year specified in cells F2 and G2. The result is displayed in I1, formulas are calculated, and the file is saved.
// Keywords: Aspose.Cells | C# | SUMIFS | named range | monthly sales aggregation | dynamic reporting | Excel formula automation | Aspose.Cells .NET example | GitHub sample | workbook calculation
// Common Searches: Aspose.Cells SUMIFS named range C# | how to aggregate monthly sales with Aspose.Cells | dynamic report using named ranges in .NET | C# create named range for SUMIFS | Aspose.Cells example for monthly totals
// Developer Intent: Define column‑level named ranges and a SUMIFS named range to compute month‑year sales totals that update automatically when criteria cells change.
// Use Cases: Show a running total for a selected month and year that refreshes when the user edits cells F2 or G2. | Reference the MonthlySales named range on other worksheets to build consolidated dashboards without duplicating the SUMIFS formula. | Retrieve the underlying SalesData range via GetRange for further processing or export.
// AI Prompts: Generate C# code using Aspose.Cells that creates named ranges for sales, month, and year columns and adds a SUMIFS named range to total sales based on month and year criteria cells. | Demonstrate how to calculate the SUMIFS named range, display the result in a cell, and save the workbook with Aspose.Cells for .NET. | Explain how changing the month or year values in the criteria cells updates the aggregated sales value automatically.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This C# example shows how to build a workbook with date, month, year, and sales columns, define column‑based named ranges (SalesData, MonthData, YearData), and create a fourth named range (MonthlySales) that uses the SUMIFS function to total sales for a month and year specified in cells F2 and G2. The result is displayed in I1, formulas are calculated, and the file is saved.
    public class NamedRangeSumIfsDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate header row
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Month");
            sheet.Cells["C1"].PutValue("Year");
            sheet.Cells["D1"].PutValue("Sales");

            // Sample data
            string[] dates = { "2023-01-05", "2023-01-12", "2023-01-20", "2023-02-03", "2023-02-15", "2023-02-28" };
            string[] months = { "Jan", "Jan", "Jan", "Feb", "Feb", "Feb" };
            int[] years = { 2023, 2023, 2023, 2023, 2023, 2023 };
            double[] sales = { 1500, 2300, 1800, 2100, 1900, 2500 };

            // Fill worksheet with sample rows
            for (int i = 0; i < dates.Length; i++)
            {
                int row = i + 1; // Row index (0‑based). Row 1 is the second row in the sheet.
                sheet.Cells[row, 0].PutValue(DateTime.Parse(dates[i])); // Date
                sheet.Cells[row, 1].PutValue(months[i]);               // Month
                sheet.Cells[row, 2].PutValue(years[i]);                // Year
                sheet.Cells[row, 3].PutValue(sales[i]);                // Sales
            }

            int lastDataRow = dates.Length; // Number of data rows

            // Create named ranges for each column
            int idxSales = sheet.Workbook.Worksheets.Names.Add("SalesData");
            sheet.Workbook.Worksheets.Names[idxSales].RefersTo = $"={sheet.Name}!$D$2:$D${lastDataRow + 1}";

            int idxMonth = sheet.Workbook.Worksheets.Names.Add("MonthData");
            sheet.Workbook.Worksheets.Names[idxMonth].RefersTo = $"={sheet.Name}!$B$2:$B${lastDataRow + 1}";

            int idxYear = sheet.Workbook.Worksheets.Names.Add("YearData");
            sheet.Workbook.Worksheets.Names[idxYear].RefersTo = $"={sheet.Name}!$C$2:$C${lastDataRow + 1}";

            // Cells that hold the criteria for month and year
            sheet.Cells["F1"].PutValue("Month");
            sheet.Cells["F2"].PutValue("Jan");   // Change to get different month totals
            sheet.Cells["G1"].PutValue("Year");
            sheet.Cells["G2"].PutValue(2023);    // Change to get different year totals

            // Create a named range that aggregates sales using SUMIFS
            int idxAgg = sheet.Workbook.Worksheets.Names.Add("MonthlySales");
            // Formula: =SUMIFS(SalesData, MonthData, F2, YearData, G2)
            sheet.Workbook.Worksheets.Names[idxAgg].RefersTo = "=SUMIFS(SalesData, MonthData, F2, YearData, G2)";

            // Use the named range in a cell to display the result
            sheet.Cells["I1"].Formula = "=MonthlySales";

            // Calculate all formulas
            workbook.CalculateFormula();

            // Demonstrate GetRange on the SalesData named range
            Name salesName = sheet.Workbook.Worksheets.Names["SalesData"];
            Aspose.Cells.Range salesRange = salesName.GetRange();
            Console.WriteLine($"SalesData refers to range: {salesRange.RefersTo}");

            // Output the aggregated result
            Console.WriteLine($"Aggregated sales for month '{sheet.Cells["F2"].StringValue}' and year {sheet.Cells["G2"].IntValue}: {sheet.Cells["I1"].Value}");

            // Save the workbook
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "NamedRangeSumIfsDemo.xlsx");
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
    }
}
