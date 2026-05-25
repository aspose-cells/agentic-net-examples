using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsImportAndSummarize
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------- Import first table --------------------
            DataTable table1 = new DataTable("Sales");
            table1.Columns.Add("Product", typeof(string));
            table1.Columns.Add("Quantity", typeof(int));

            table1.Rows.Add("Apple", 120);
            table1.Rows.Add("Banana", 85);
            table1.Rows.Add("Cherry", 60);

            // Import options: show field names
            ImportTableOptions options1 = new ImportTableOptions
            {
                IsFieldNameShown = true,
                InsertRows = true
            };

            // Import into the first worksheet starting at A1 (row 0, column 0)
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "SalesData";
            sheet1.Cells.ImportData(table1, 0, 0, options1);

            // -------------------- Import second table --------------------
            DataTable table2 = new DataTable("Expenses");
            table2.Columns.Add("Category", typeof(string));
            table2.Columns.Add("Amount", typeof(double));

            table2.Rows.Add("Rent", 1500.00);
            table2.Rows.Add("Utilities", 300.50);
            table2.Rows.Add("Supplies", 120.75);

            ImportTableOptions options2 = new ImportTableOptions
            {
                IsFieldNameShown = true,
                InsertRows = true
            };

            // Add a new worksheet for the second table
            Worksheet sheet2 = workbook.Worksheets.Add("ExpenseData");
            sheet2.Cells.ImportData(table2, 0, 0, options2);

            // -------------------- Create Summary sheet --------------------
            Worksheet summarySheet = workbook.Worksheets.Add("Summary");

            // Header
            summarySheet.Cells["A1"].PutValue("Table");
            summarySheet.Cells["B1"].PutValue("Total");

            // Row for first table total
            summarySheet.Cells["A2"].PutValue("Sales Quantity Total");
            // SUM of Quantity column in SalesData (B column, rows 2 to 4)
            int salesDataRows = table1.Rows.Count + 1; // +1 for header row
            string salesSumFormula = $"=SUM('{sheet1.Name}'!B2:B{salesDataRows})";
            summarySheet.Cells["B2"].Formula = salesSumFormula;

            // Row for second table total
            summarySheet.Cells["A3"].PutValue("Expenses Amount Total");
            // SUM of Amount column in ExpenseData (B column, rows 2 to 4)
            int expenseDataRows = table2.Rows.Count + 1;
            string expenseSumFormula = $"=SUM('{sheet2.Name}'!B2:B{expenseDataRows})";
            summarySheet.Cells["B3"].Formula = expenseSumFormula;

            // Auto‑fit columns for better readability
            summarySheet.AutoFitColumns();

            // -------------------- Save the workbook --------------------
            workbook.Save("AggregatedSummary.xlsx", SaveFormat.Xlsx);
        }
    }
}