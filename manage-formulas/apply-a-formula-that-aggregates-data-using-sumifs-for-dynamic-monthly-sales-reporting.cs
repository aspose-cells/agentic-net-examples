// Title: Create a dynamic monthly sales report in C# using Aspose.Cells with a SUMIFS formula
// Common Searches: C# Aspose.Cells example of SUMIFS to total sales by month | How to build a monthly sales summary workbook with a dynamic month selector in Aspose.Cells | Using DATE and EDATE functions in an Aspose.Cells formula for monthly aggregation | Programmatically calculate total sales for a specific month with Aspose.Cells .NET | Generate an Excel report that sums sales per month based on a cell input using Aspose.Cells
// Tags: Aspose.Cells SUMIFS monthly aggregation | C# generate Excel sales report | dynamic month selector formula Aspose.Cells | DATE EDATE functions in Aspose.Cells | programmatic Excel formula calculation .NET

using Aspose.Cells;
using System;

// The sample creates a workbook with a 'Data' sheet containing dates and random sales figures, adds a 'MonthlyReport' sheet where a month number is entered, and applies a SUMIFS formula that uses DATE and EDATE to sum sales for the selected month. The formula is recalculated and the workbook is saved as MonthlySalesReport.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // -------------------------------------------------
        // Sheet 1: Raw sales data (Date, Sales)
        // -------------------------------------------------
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Name = "Data";

        // Headers
        dataSheet.Cells["A1"].PutValue("Date");
        dataSheet.Cells["B1"].PutValue("Sales");

        // Sample data (10 rows)
        DateTime startDate = new DateTime(DateTime.Today.Year, 1, 1);
        Random rnd = new Random();
        for (int i = 0; i < 10; i++)
        {
            // Date every 3 days
            dataSheet.Cells[i + 1, 0].PutValue(startDate.AddDays(i * 3));
            // Random sales amount
            dataSheet.Cells[i + 1, 1].PutValue(rnd.Next(100, 1000));
        }

        // -------------------------------------------------
        // Sheet 2: Monthly sales report using SUMIFS
        // -------------------------------------------------
        Worksheet reportSheet = workbook.Worksheets.Add("MonthlyReport");

        // Header for month selector and result
        reportSheet.Cells["A1"].PutValue("Month (1‑12)");
        reportSheet.Cells["B1"].PutValue("Total Sales");

        // Example month selector (e.g., March)
        reportSheet.Cells["A2"].PutValue(3);

        // SUMIFS formula:
        //   Sum Sales where Date >= first day of selected month
        //   and Date < first day of next month (dynamic using DATE & EDATE)
        string sumIfsFormula = @"=SUMIFS(Data!$B$2:$B$100,Data!$A$2:$A$100,"">=""&DATE(YEAR(TODAY()),A2,1),Data!$A$2:$A$100,""<""&EDATE(DATE(YEAR(TODAY()),A2,1),1))";

        reportSheet.Cells["B2"].Formula = sumIfsFormula;

        // Recalculate formulas so B2 contains the computed total
        workbook.CalculateFormula();

        // Save the workbook
        workbook.Save("MonthlySalesReport.xlsx");
    }
}
