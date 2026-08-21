// Title: Generate a Dynamic Monthly Sales Report with SUMIFS in C# using Aspose.Cells
// Description: This C# example creates an Excel workbook, fills column A with random 2023 dates and column B with random sales amounts, lists month numbers in column D, inserts a SUMIFS formula in column E to total sales for each month, evaluates all formulas, and saves the file as MonthlySalesReport.xlsx.
// Keywords: Aspose.Cells | C# SUMIFS | monthly sales aggregation | dynamic date range formula | Excel .NET report | calculate monthly totals | Aspose.Cells workbook creation | SUMIFS with DATE and EOMONTH | C# Excel automation | generate sales report
// Common Searches: Aspose.Cells SUMIFS example C# | How to calculate monthly totals with SUMIFS in .NET | C# code for dynamic monthly sales report Excel | Using DATE and EOMONTH in Aspose.Cells formulas | Create Excel workbook with random data C# Aspose.Cells
// Developer Intent: Programmatically build an Excel file that aggregates sales by month using a SUMIFS formula and saves the result.
// Use Cases: Generate a sample sales dataset with dates and amounts for testing analytics pipelines. | Automatically compute month‑by‑month totals without manual formula entry. | Export the calculated monthly totals to an .xlsx file for downstream reporting or visualization. | Adapt the same pattern to other time‑based aggregations such as quarterly or yearly summaries.
// AI Prompts: Show how to extend the SUMIFS formula to include a year parameter for multi‑year reports. | Provide code to format the Monthly Total column as currency and add a line chart of monthly sales. | Explain how to replace the random data generation with data imported from a CSV file while keeping the SUMIFS aggregation.

using System;
using Aspose.Cells;

namespace AsposeCellsSumIfsDemo
{
    // This C# example creates an Excel workbook, fills column A with random 2023 dates and column B with random sales amounts, lists month numbers in column D, inserts a SUMIFS formula in column E to total sales for each month, evaluates all formulas, and saves the file as MonthlySalesReport.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data: Column A = Date, Column B = Sales
            // Dates cover Jan to Dec 2023
            DateTime startDate = new DateTime(2023, 1, 1);
            Random rnd = new Random();
            for (int i = 0; i < 100; i++)
            {
                // Random date within the year 2023
                DateTime date = startDate.AddDays(rnd.Next(0, 365));
                double sales = rnd.Next(100, 1000); // Random sales amount

                cells[i, 0].PutValue(date); // A column
                cells[i, 1].PutValue(sales); // B column
            }

            // Header row
            cells[0, 0].PutValue("Date");
            cells[0, 1].PutValue("Sales");
            cells[0, 3].PutValue("Month");   // D column header
            cells[0, 4].PutValue("Monthly Total"); // E column header

            // List months 1..12 in column D (starting from row 2)
            for (int m = 1; m <= 12; m++)
            {
                cells[m, 3].PutValue(m); // D2:D13 contain month numbers
            }

            // Apply SUMIFS formula in column E to aggregate sales per month
            // Formula uses absolute references for data ranges and concatenates DATE/EOMONTH with month number from column D
            // Example for row 2 (month 1):
            // =SUMIFS($B$2:$B$101, $A$2:$A$101, ">="&DATE(2023, D2, 1), $A$2:$A$101, "<="&EOMONTH(DATE(2023, D2, 1),0))
            for (int row = 1; row <= 12; row++)
            {
                string formula = $"=SUMIFS($B$2:$B$101, $A$2:$A$101, \">=\"&DATE(2023, D{row + 1}, 1), $A$2:$A$101, \"<=\"&EOMONTH(DATE(2023, D{row + 1}, 1),0))";
                cells[row, 4].Formula = formula; // E column
            }

            // Calculate all formulas so that results are stored in the cells
            workbook.CalculateFormula();

            // Save the workbook
            workbook.Save("MonthlySalesReport.xlsx");
        }
    }
}
