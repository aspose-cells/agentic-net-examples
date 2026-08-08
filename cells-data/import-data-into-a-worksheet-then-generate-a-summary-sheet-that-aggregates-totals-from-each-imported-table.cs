// Title: Import DataTables into worksheets and generate a totals summary with Aspose.Cells for .NET (C#)
// Description: Creates a new Workbook, builds two DataTables (Sales and Purchases), imports each table into its own worksheet with headers, adds a third "Summary" sheet, writes labels, and inserts SUM formulas that total the Amount or Cost columns from the source sheets. The formulas are calculated and the workbook is saved as an XLSX file.
// Keywords: Aspose.Cells | C# | Import DataTable | Worksheet import | Summary sheet | SUM formula across sheets | CalculateFormula | Export to XLSX | ImportTableOptions | Excel automation .NET
// Common Searches: Aspose.Cells import DataTable into specific worksheet | Create a summary worksheet that totals columns from multiple sheets in C# | How to add formulas that reference other worksheets with Aspose.Cells | Calculate formulas after importing data using Aspose.Cells | Save workbook as XLSX with Aspose.Cells .NET
// Developer Intent: Generate a workbook, import two DataTables into separate worksheets, and add a summary sheet that aggregates numeric totals via formulas.
// Use Cases: Load sales and purchase records into distinct sheets and automatically compute total sales amount and total purchase cost on a consolidated summary page. | Produce a departmental financial report where each department's data resides on its own worksheet and a summary sheet provides key aggregated metrics. | Combine multiple data sources into a single Excel file and create an overview sheet that sums important numeric columns for quick analysis.
// AI Prompts: Show how to add average quantity calculations for each source sheet in the summary worksheet. | Provide code to apply conditional formatting to the summary totals based on a configurable threshold. | Explain how to dynamically add more DataTables and automatically extend the summary sheet with additional total rows.

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsImportAndSummarize
{
    // Creates a new Workbook, builds two DataTables (Sales and Purchases), imports each table into its own worksheet with headers, adds a third "Summary" sheet, writes labels, and inserts SUM formulas that total the Amount or Cost columns from the source sheets. The formulas are calculated and the workbook is saved as an XLSX file.
    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();

            // ---------- Prepare first data table ----------
            DataTable salesTable = new DataTable("Sales");
            salesTable.Columns.Add("Product", typeof(string));
            salesTable.Columns.Add("Quantity", typeof(int));
            salesTable.Columns.Add("Amount", typeof(double));

            salesTable.Rows.Add("Apple", 10, 150.0);
            salesTable.Rows.Add("Banana", 20, 120.0);
            salesTable.Rows.Add("Orange", 15, 180.0);

            // ---------- Import first table into the first worksheet ----------
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "SalesData";
            ImportTableOptions options1 = new ImportTableOptions
            {
                IsFieldNameShown = true,   // import column headers
                InsertRows = true
            };
            sheet1.Cells.ImportData(salesTable, 0, 0, options1);

            // ---------- Prepare second data table ----------
            DataTable purchaseTable = new DataTable("Purchases");
            purchaseTable.Columns.Add("Supplier", typeof(string));
            purchaseTable.Columns.Add("Quantity", typeof(int));
            purchaseTable.Columns.Add("Cost", typeof(double));

            purchaseTable.Rows.Add("SupplierA", 30, 200.0);
            purchaseTable.Rows.Add("SupplierB", 25, 175.0);
            purchaseTable.Rows.Add("SupplierC", 40, 300.0);

            // ---------- Add a new worksheet and import second table ----------
            int sheetIndex = workbook.Worksheets.Add();
            Worksheet sheet2 = workbook.Worksheets[sheetIndex];
            sheet2.Name = "PurchaseData";
            ImportTableOptions options2 = new ImportTableOptions
            {
                IsFieldNameShown = true,
                InsertRows = true
            };
            sheet2.Cells.ImportData(purchaseTable, 0, 0, options2);

            // ---------- Create a summary worksheet ----------
            Worksheet summarySheet = workbook.Worksheets[workbook.Worksheets.Add()];
            summarySheet.Name = "Summary";

            // Header row
            summarySheet.Cells["A1"].PutValue("Source Sheet");
            summarySheet.Cells["B1"].PutValue("Total Amount/Cost");

            // Row for SalesData total (sum of Amount column, which is column C => index 2)
            summarySheet.Cells["A2"].PutValue("SalesData");
            // Formula sums the entire Amount column excluding header (starts at row 2)
            summarySheet.Cells["B2"].Formula = $"SUM('{sheet1.Name}'!C2:C{sheet1.Cells.MaxDataRow + 1})";

            // Row for PurchaseData total (sum of Cost column, column C => index 2)
            summarySheet.Cells["A3"].PutValue("PurchaseData");
            summarySheet.Cells["B3"].Formula = $"SUM('{sheet2.Name}'!C2:C{sheet2.Cells.MaxDataRow + 1})";

            // Optional: calculate the formulas now
            workbook.CalculateFormula();

            // ---------- Save the workbook ----------
            workbook.Save("ImportAndSummary.xlsx", SaveFormat.Xlsx);
        }
    }
}
