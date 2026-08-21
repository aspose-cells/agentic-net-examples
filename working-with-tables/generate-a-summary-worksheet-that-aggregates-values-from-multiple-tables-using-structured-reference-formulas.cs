// Title: Create a Summary Sheet with Structured References from Multiple Tables in Aspose.Cells C#
// Description: This example builds a new workbook, adds two ListObject tables (SalesData and ExpenseData) on separate worksheets, then creates a Summary sheet that uses structured reference formulas (e.g., =SUM(SalesData[Amount])) to compute total sales, total expenses, and net profit. The formulas are evaluated and the workbook is saved as SummaryWorkbook.xlsx.
// Keywords: Aspose.Cells | C# structured references | ListObject table | summary worksheet | aggregate table data | SUM formula | net profit calculation | Excel automation .NET
// Common Searches: Aspose.Cells structured reference formula example | C# create summary sheet from multiple tables | How to sum ListObject column in Aspose.Cells | Calculate net profit using Aspose.Cells tables | Programmatic Excel summary worksheet C#
// Developer Intent: Generate a workbook with two data tables and a summary sheet that totals each table using structured reference formulas.
// Use Cases: Financial reporting workbook that updates totals automatically when source tables change. | Dynamic dashboard where key metrics are derived from separate sales and expense tables. | Automated profit‑and‑loss statement generation by aggregating table data into a single summary sheet.
// AI Prompts: Show C# code to add a ListObject table and reference its column with a structured formula in Aspose.Cells. | Provide an Aspose.Cells example that creates a summary worksheet calculating total sales, total expenses, and net profit. | Explain how to evaluate and save formulas after building tables in an Aspose.Cells workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace SummaryWorksheetExample
{
    // This example builds a new workbook, adds two ListObject tables (SalesData and ExpenseData) on separate worksheets, then creates a Summary sheet that uses structured reference formulas (e.g., =SUM(SalesData[Amount])) to compute total sales, total expenses, and net profit. The formulas are evaluated and the workbook is saved as SummaryWorkbook.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // 1. Create first data table (SalesData) on Sheet1
                // -------------------------------------------------
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "Sales";

                // Header row
                sheet1.Cells["A1"].PutValue("Region");
                sheet1.Cells["B1"].PutValue("Amount");

                // Data rows
                sheet1.Cells["A2"].PutValue("North");
                sheet1.Cells["B2"].PutValue(1200);
                sheet1.Cells["A3"].PutValue("South");
                sheet1.Cells["B3"].PutValue(850);
                sheet1.Cells["A4"].PutValue("East");
                sheet1.Cells["B4"].PutValue(950);
                sheet1.Cells["A5"].PutValue("West");
                sheet1.Cells["B5"].PutValue(1100);

                // Convert the range into a structured table
                ListObjectCollection tables1 = sheet1.ListObjects;
                int salesTableIndex = tables1.Add(0, 0, 5, 1, true); // A1:B5
                ListObject salesTable = tables1[salesTableIndex];
                // Set table name (used in structured references)
                salesTable.DisplayName = "SalesData";

                // -------------------------------------------------
                // 2. Create second data table (ExpenseData) on Sheet2
                // -------------------------------------------------
                Worksheet sheet2 = workbook.Worksheets.Add("Expenses");

                // Header row
                sheet2.Cells["A1"].PutValue("Category");
                sheet2.Cells["B1"].PutValue("Amount");

                // Data rows
                sheet2.Cells["A2"].PutValue("Rent");
                sheet2.Cells["B2"].PutValue(500);
                sheet2.Cells["A3"].PutValue("Utilities");
                sheet2.Cells["B3"].PutValue(200);
                sheet2.Cells["A4"].PutValue("Supplies");
                sheet2.Cells["B4"].PutValue(150);
                sheet2.Cells["A5"].PutValue("Travel");
                sheet2.Cells["B5"].PutValue(300);

                // Convert the range into a structured table
                ListObjectCollection tables2 = sheet2.ListObjects;
                int expenseTableIndex = tables2.Add(0, 0, 5, 1, true); // A1:B5
                ListObject expenseTable = tables2[expenseTableIndex];
                expenseTable.DisplayName = "ExpenseData";

                // -------------------------------------------------
                // 3. Create Summary worksheet with structured reference formulas
                // -------------------------------------------------
                Worksheet summary = workbook.Worksheets.Add("Summary");
                Cells sumCells = summary.Cells;

                // Labels
                sumCells["A1"].PutValue("Metric");
                sumCells["B1"].PutValue("Value");

                // Total Sales
                sumCells["A2"].PutValue("Total Sales");
                sumCells["B2"].Formula = "=SUM(SalesData[Amount])";

                // Total Expenses
                sumCells["A3"].PutValue("Total Expenses");
                sumCells["B3"].Formula = "=SUM(ExpenseData[Amount])";

                // Net Profit
                sumCells["A4"].PutValue("Net Profit");
                sumCells["B4"].Formula = "=B2-B3";

                // Calculate all formulas so that the workbook contains the computed values
                workbook.CalculateFormula();

                // -------------------------------------------------
                // 4. Save the workbook
                // -------------------------------------------------
                workbook.Save("SummaryWorkbook.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
