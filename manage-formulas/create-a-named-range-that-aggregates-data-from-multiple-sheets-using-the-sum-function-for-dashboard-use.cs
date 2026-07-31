// Title: Aspose.Cells .NET: Create a Named Range that Sums Data Across Multiple Worksheets for Dashboard Reporting
// Description: Demonstrates how to build a new workbook, populate two sheets, define a named range called TotalSales that uses the SUM function to aggregate A1:A5 from both sheets, reference the range in a dashboard cell, trigger formula calculation, and save the file as DashboardNamedRange.xlsx using Aspose.Cells for C#.
// Keywords: Aspose.Cells | C# | .NET | named range | SUM across worksheets | RefersTo formula | dashboard reporting | Excel automation | calculate formulas | multiple sheets
// Common Searches: Aspose.Cells create named range that sums multiple sheets | C# SUM function across worksheets using Aspose.Cells | How to use RefersTo property for dashboard totals | Calculate formulas after adding a named range in Aspose.Cells | Save workbook with aggregated totals in Aspose.Cells .NET
// Developer Intent: Define a named range that aggregates values from several worksheets with SUM and display the result on a dashboard cell.
// Use Cases: Generate a consolidated total for sales data spread over several sheets. | Provide a single reference for KPI dashboards that updates automatically when source data changes. | Create reusable named ranges for financial summaries across quarterly worksheets. | Automate formula recalculation after adding or modifying named ranges before exporting the workbook.
// AI Prompts: Show C# code that creates a named range summing A1:A5 from Sheet1 and Sheet2 with Aspose.Cells. | Explain how to update the RefersTo formula of a named range dynamically based on a list of sheet names. | Give steps to force formula calculation after defining a named range so the dashboard cell shows the correct total.

using System;
using Aspose.Cells;

// Demonstrates how to build a new workbook, populate two sheets, define a named range called TotalSales that uses the SUM function to aggregate A1:A5 from both sheets, reference the range in a dashboard cell, trigger formula calculation, and save the file as DashboardNamedRange.xlsx using Aspose.Cells for C#.
class DashboardNamedRange
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // -------------------------------------------------
        // Sheet1: populate sample data in column A (A1:A5)
        // -------------------------------------------------
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "Sheet1";

        for (int i = 0; i < 5; i++)
        {
            // Values 1,2,3,4,5
            sheet1.Cells[i, 0].PutValue(i + 1);
        }

        // -------------------------------------------------
        // Sheet2: populate sample data in column A (A1:A5)
        // -------------------------------------------------
        Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");

        for (int i = 0; i < 5; i++)
        {
            // Values 10,20,30,40,50
            sheet2.Cells[i, 0].PutValue((i + 1) * 10);
        }

        // -------------------------------------------------
        // Create a named range that aggregates data from both sheets using SUM
        // -------------------------------------------------
        int nameIdx = workbook.Worksheets.Names.Add("TotalSales");
        Name totalSalesName = workbook.Worksheets.Names[nameIdx];
        // The RefersTo property holds a formula; SUM across the two ranges
        totalSalesName.RefersTo = "=SUM(Sheet1!$A$1:$A$5,Sheet2!$A$1:$A$5)";

        // -------------------------------------------------
        // Use the named range in a dashboard cell (e.g., B1 on Sheet1)
        // -------------------------------------------------
        sheet1.Cells["B1"].Formula = "=TotalSales";

        // Calculate formulas so the result appears in B1
        workbook.CalculateFormula();

        // -------------------------------------------------
        // Save the workbook
        // -------------------------------------------------
        workbook.Save("DashboardNamedRange.xlsx");
    }
}
