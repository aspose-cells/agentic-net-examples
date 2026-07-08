using System;
using System.Data;
using Aspose.Cells;

class SummarySheetExample
{
    static void Main()
    {
        // -------------------------------------------------
        // 1. Create a template workbook with smart markers
        // -------------------------------------------------
        Workbook template = new Workbook();

        // Worksheet 1 – Region1
        Worksheet ws1 = template.Worksheets[0];
        ws1.Name = "Region1";
        ws1.Cells["A1"].PutValue("Product");
        ws1.Cells["B1"].PutValue("Amount");
        ws1.Cells["A2"].PutValue("&=$Product");
        ws1.Cells["B2"].PutValue("&=$Amount");

        // Worksheet 2 – Region2
        int ws2Index = template.Worksheets.Add();
        Worksheet ws2 = template.Worksheets[ws2Index];
        ws2.Name = "Region2";
        ws2.Cells["A1"].PutValue("Product");
        ws2.Cells["B1"].PutValue("Amount");
        ws2.Cells["A2"].PutValue("&=$Product");
        ws2.Cells["B2"].PutValue("&=$Amount");

        // -------------------------------------------------
        // 2. Prepare a data source that contains data for both regions
        // -------------------------------------------------
        DataTable sales = new DataTable("Sales");
        sales.Columns.Add("Product", typeof(string));
        sales.Columns.Add("Amount", typeof(double));
        sales.Columns.Add("Region", typeof(string));

        sales.Rows.Add("A", 100.0, "Region1");
        sales.Rows.Add("B", 150.0, "Region1");
        sales.Rows.Add("C", 200.0, "Region2");
        sales.Rows.Add("D", 250.0, "Region2");

        // -------------------------------------------------
        // 3. Process smart markers – each worksheet will repeat rows for its region
        // -------------------------------------------------
        WorkbookDesigner designer = new WorkbookDesigner(template);
        designer.SetDataSource(sales);
        designer.Process();

        // -------------------------------------------------
        // 4. Add a summary worksheet that aggregates totals from the region sheets
        // -------------------------------------------------
        int summaryIdx = template.Worksheets.Add();
        Worksheet summary = template.Worksheets[summaryIdx];
        summary.Name = "Summary";

        // Header
        summary.Cells["A1"].PutValue("Region");
        summary.Cells["B1"].PutValue("Total Amount");

        // List of regions (matching the names of the data worksheets)
        summary.Cells["A2"].PutValue("Region1");
        summary.Cells["A3"].PutValue("Region2");

        // Formulas that sum the Amount column of each region sheet
        // Assuming data starts at row 2 and may extend downwards; using a large row number to cover all rows
        summary.Cells["B2"].Formula = $"=SUM('{ws1.Name}'!B2:B1000)";
        summary.Cells["B3"].Formula = $"=SUM('{ws2.Name}'!B2:B1000)";

        // Calculate the formulas so that the workbook contains the computed values
        template.CalculateFormula();

        // -------------------------------------------------
        // 5. Save the final workbook
        // -------------------------------------------------
        template.Save("SummaryReport.xlsx");
    }
}