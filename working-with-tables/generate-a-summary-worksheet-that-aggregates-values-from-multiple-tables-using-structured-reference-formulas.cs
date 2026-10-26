// Title: Create a summary worksheet that totals sales and expenses using structured reference formulas with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# how to calculate total sales using a structured reference to a ListObject table | C# generate a summary sheet that sums columns from multiple tables with Aspose.Cells formulas | How to compute net revenue from SalesTable and ExpenseTable in Aspose.Cells .NET | Add a summary worksheet with SUM formulas referencing ListObject columns using Aspose.Cells | Aspose.Cells example for aggregating data across worksheets into a net revenue calculation
// Tags: structured reference SUM formula Aspose.Cells | aggregate table columns across worksheets C# | create summary worksheet Aspose.Cells .NET | ListObject table aggregation Aspose.Cells | net revenue calculation using Aspose.Cells formulas

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsSummaryExample
{
    // The example creates a workbook with SalesData and ExpenseData worksheets, converts each range into ListObject tables named SalesTable and ExpenseTable, then adds a Summary worksheet that uses structured reference formulas (e.g., =SUM(SalesTable[Amount])) to compute total sales, total expenses, and net revenue, auto‑fits columns, and saves the file as SummaryWorkbook.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // Worksheet 1: SalesData
                // -------------------------------------------------
                Worksheet wsSales = workbook.Worksheets[0];
                wsSales.Name = "SalesData";

                // Populate sample data
                wsSales.Cells["A1"].PutValue("Region");
                wsSales.Cells["B1"].PutValue("Product");
                wsSales.Cells["C1"].PutValue("Amount");

                wsSales.Cells["A2"].PutValue("North");
                wsSales.Cells["B2"].PutValue("Widget");
                wsSales.Cells["C2"].PutValue(1200);

                wsSales.Cells["A3"].PutValue("South");
                wsSales.Cells["B3"].PutValue("Gadget");
                wsSales.Cells["C3"].PutValue(850);

                wsSales.Cells["A4"].PutValue("East");
                wsSales.Cells["B4"].PutValue("Widget");
                wsSales.Cells["C4"].PutValue(970);

                // Convert range to a table (ListObject) and give it a name
                int firstRow = 0;          // zero‑based index
                int firstColumn = 0;
                int totalRows = 4;         // header + 3 data rows
                int totalColumns = 3;
                int salesTableIndex = wsSales.ListObjects.Add(
                    firstRow,
                    firstColumn,
                    firstRow + totalRows - 1,
                    firstColumn + totalColumns - 1,
                    true);
                ListObject salesTable = wsSales.ListObjects[salesTableIndex];
                salesTable.DisplayName = "SalesTable";

                // -------------------------------------------------
                // Worksheet 2: ExpenseData
                // -------------------------------------------------
                Worksheet wsExpense = workbook.Worksheets.Add("ExpenseData");

                // Populate sample data
                wsExpense.Cells["A1"].PutValue("Category");
                wsExpense.Cells["B1"].PutValue("Month");
                wsExpense.Cells["C1"].PutValue("Cost");

                wsExpense.Cells["A2"].PutValue("Travel");
                wsExpense.Cells["B2"].PutValue("Jan");
                wsExpense.Cells["C2"].PutValue(300);

                wsExpense.Cells["A3"].PutValue("Supplies");
                wsExpense.Cells["B3"].PutValue("Jan");
                wsExpense.Cells["C3"].PutValue(150);

                wsExpense.Cells["A4"].PutValue("Travel");
                wsExpense.Cells["B4"].PutValue("Feb");
                wsExpense.Cells["C4"].PutValue(420);

                // Convert range to a table and name it
                int expenseTableIndex = wsExpense.ListObjects.Add(0, 0, 3, 2, true);
                ListObject expenseTable = wsExpense.ListObjects[expenseTableIndex];
                expenseTable.DisplayName = "ExpenseTable";

                // -------------------------------------------------
                // Worksheet 3: Summary (aggregates using structured references)
                // -------------------------------------------------
                Worksheet wsSummary = workbook.Worksheets.Add("Summary");

                // Header for summary
                wsSummary.Cells["A1"].PutValue("Metric");
                wsSummary.Cells["B1"].PutValue("Value");

                // Row 2: Total Sales Amount (SUM of SalesTable[Amount])
                wsSummary.Cells["A2"].PutValue("Total Sales");
                wsSummary.Cells["B2"].Formula = "=SUM(SalesTable[Amount])";

                // Row 3: Total Expense Cost (SUM of ExpenseTable[Cost])
                wsSummary.Cells["A3"].PutValue("Total Expenses");
                wsSummary.Cells["B3"].Formula = "=SUM(ExpenseTable[Cost])";

                // Row 4: Net Revenue (Total Sales - Total Expenses)
                wsSummary.Cells["A4"].PutValue("Net Revenue");
                wsSummary.Cells["B4"].Formula = "=B2-B3";

                // Auto‑fit columns for better readability
                wsSummary.AutoFitColumns();

                // -------------------------------------------------
                // Save the workbook to a file
                // -------------------------------------------------
                string outputPath = "SummaryWorkbook.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
