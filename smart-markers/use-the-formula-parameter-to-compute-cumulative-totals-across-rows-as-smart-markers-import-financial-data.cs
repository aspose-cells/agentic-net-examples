// Title: C# – Compute Cumulative Totals with Smart Markers & Formula Parameter in Aspose.Cells
// Description: Demonstrates how to import a financial DataTable into an Excel template using Aspose.Cells smart markers, apply a running‑total formula (=SUM($B$2:B2)) via the Formula parameter, let WorkbookDesigner copy the formula down automatically, evaluate the totals with CalculateFormula, and save the result as CumulativeTotals.xlsx.
// Keywords: Aspose.Cells | C# | .NET | Smart Markers | Formula Parameter | cumulative total | running total | WorkbookDesigner | Excel automation | financial report
// Common Searches: Aspose.Cells smart markers cumulative total example | C# running total column using formula parameter | WorkbookDesigner copy formulas automatically | How to calculate cumulative sum in Aspose.Cells | Smart markers import DataTable Excel
// Developer Intent: Create an Excel worksheet where each row from a DataTable is inserted via smart markers and a cumulative total column is generated automatically using a formula parameter.
// Use Cases: Generate a financial ledger with a running total for transaction amounts. | Build a monthly sales report that shows cumulative sales alongside each sale entry. | Produce a payroll sheet that lists employee earnings and accumulates total pay per employee.
// AI Prompts: Show how to reset the cumulative total at the start of each month using smart markers. | Provide code to handle multiple DataTables, each with its own cumulative total column, in a single workbook. | Explain how to add conditional formatting to highlight rows where the cumulative total exceeds a threshold after processing smart markers.

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// Demonstrates how to import a financial DataTable into an Excel template using Aspose.Cells smart markers, apply a running‑total formula (=SUM($B$2:B2)) via the Formula parameter, let WorkbookDesigner copy the formula down automatically, evaluate the totals with CalculateFormula, and save the result as CumulativeTotals.xlsx.
class CumulativeTotalsSmartMarkers
{
    static void Main()
    {
        try
        {
            // ---------- 1. Prepare source data (financial transactions) ----------
            DataTable financialData = new DataTable("Financial");
            financialData.Columns.Add("Date", typeof(DateTime));
            financialData.Columns.Add("Amount", typeof(double));

            financialData.Rows.Add(new DateTime(2023, 1, 1), 1500.0);
            financialData.Rows.Add(new DateTime(2023, 1, 5), 2300.0);
            financialData.Rows.Add(new DateTime(2023, 1, 10), -500.0);
            financialData.Rows.Add(new DateTime(2023, 1, 15), 1200.0);
            financialData.Rows.Add(new DateTime(2023, 1, 20), 800.0);

            // ---------- 2. Create a workbook and design the template ----------
            Workbook workbook = new Workbook();                     // create workbook
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Header row
            cells["A1"].PutValue("Date");
            cells["B1"].PutValue("Amount");
            cells["C1"].PutValue("Cumulative Total");

            // Row 2 contains smart markers that will be repeated for each DataTable row
            //   &=$Date   -> inserts the Date value
            //   &=$Amount -> inserts the Amount value
            //   =SUM($B$2:B2) -> running total: sum from first amount (absolute $B$2) to current row B
            cells["A2"].PutValue("&=$Date");
            cells["B2"].PutValue("&=$Amount");
            cells["C2"].PutValue("=SUM($B$2:B2)"); // formula parameter for cumulative total

            // ---------- 3. Use WorkbookDesigner to merge data with the template ----------
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(financialData);
            designer.Process(); // fills rows and copies the formula down automatically

            // ---------- 4. Calculate all formulas so cumulative totals are evaluated ----------
            workbook.CalculateFormula();

            // ---------- 5. Save the result ----------
            string outputPath = "CumulativeTotals.xlsx";
            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
