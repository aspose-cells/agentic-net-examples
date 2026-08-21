// Title: C# Aspose.Cells Smart Marker for Running Total Using a Previous‑Row Formula
// Description: Creates a workbook, adds Item, Amount, and Running Total headers, inserts smart markers &=$Item and &=$Amount, applies the formula =IFERROR(B2+C1,B2) to compute cumulative totals, binds a DataTable, processes the template with WorkbookDesigner, and saves the file as RunningTotalSmartMarker.xlsx.
// Keywords: Aspose.Cells | C# | Smart Marker | running total | previous row formula | IFERROR | cumulative sum | WorkbookDesigner | DataTable source | Excel automation
// Common Searches: Aspose.Cells running total smart marker example | C# formula referencing previous row in smart marker | How to calculate cumulative sum with Aspose.Cells | Smart marker IFERROR running total | Generate Excel report with running total using Aspose.Cells
// Developer Intent: Produce an Excel worksheet where a smart‑marker column automatically calculates a running total by adding each row’s amount to the total from the preceding row.
// Use Cases: Sales ledger that shows each transaction and a live cumulative total. | Inventory sheet that updates a running balance as stock entries are added. | Expense report where each line adds to a running expense total without manual formulas.
// AI Prompts: Generate C# code using Aspose.Cells that inserts smart markers for Item and Amount and adds a Running Total column with the formula =IFERROR(B2+C1,B2). | Explain the purpose of IFERROR(B2+C1,B2) in a smart‑marker template for cumulative totals. | Show how to adapt the template to start the running total from a different column or reset the total after a specific condition.

using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsSmartMarkerRunningTotal
{
    // Creates a workbook, adds Item, Amount, and Running Total headers, inserts smart markers &=$Item and &=$Amount, applies the formula =IFERROR(B2+C1,B2) to compute cumulative totals, binds a DataTable, processes the template with WorkbookDesigner, and saves the file as RunningTotalSmartMarker.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // -------------------------------------------------
                // 1. Create a new workbook (lifecycle rule: create)
                // -------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // -------------------------------------------------
                // 2. Build the template with smart markers
                // -------------------------------------------------
                // Header row
                cells["A1"].PutValue("Item");
                cells["B1"].PutValue("Amount");
                cells["C1"].PutValue("Running Total");

                // Data row (smart markers)
                // &=$Item and &=$Amount will be replaced by data source values
                cells["A2"].PutValue("&=$Item");
                cells["B2"].PutValue("&=$Amount");

                // Running total formula using regular cell references.
                // For the first data row, C1 is empty, so IFERROR returns the current amount.
                cells["C2"].Formula = "=IFERROR(B2 + C1, B2)";

                // -------------------------------------------------
                // 3. Prepare the data source (DataTable)
                // -------------------------------------------------
                DataTable dt = new DataTable("Sales");
                dt.Columns.Add("Item", typeof(string));
                dt.Columns.Add("Amount", typeof(double));

                dt.Rows.Add("Apple", 120);
                dt.Rows.Add("Banana", 80);
                dt.Rows.Add("Cherry", 150);
                dt.Rows.Add("Date", 60);

                // -------------------------------------------------
                // 4. Process the smart markers (lifecycle rule: process)
                // -------------------------------------------------
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource(dt);
                designer.Process();

                // -------------------------------------------------
                // 5. Save the result (lifecycle rule: save)
                // -------------------------------------------------
                string outputPath = "RunningTotalSmartMarker.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
