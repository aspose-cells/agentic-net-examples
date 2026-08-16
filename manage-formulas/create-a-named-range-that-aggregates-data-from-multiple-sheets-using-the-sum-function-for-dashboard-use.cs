// Title: Define a Multi‑Sheet Named Range and Sum It on a Dashboard with Aspose.Cells for .NET
// Description: Shows how to build a workbook, add two worksheets, create a named range that references the same cells on both sheets, insert a SUM(TotalData) formula on a Dashboard sheet, recalculate formulas, and save the result using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | .NET | C# | named range | multi‑sheet range | SUM formula | dashboard | aggregate data | calculate formulas | workbook save
// Common Searches: Aspose.Cells create named range across multiple worksheets | SUM formula with multi‑sheet named range C# | how to aggregate values from several sheets in Aspose.Cells | dashboard summary using named range Aspose.Cells | reference multiple areas in a named range Aspose.Cells
// Developer Intent: Create a named range that spans more than one worksheet and use it in a SUM formula to produce a consolidated total on a dashboard sheet.
// Use Cases: Combine sales totals from department sheets into a single KPI cell. | Summarize monthly performance metrics stored in separate month tabs. | Calculate total expenses across regional worksheets for a financial overview.
// AI Prompts: Generate C# code that defines a named range covering identical ranges on multiple worksheets and uses SUM(TotalData) on a dashboard sheet with Aspose.Cells. | Explain how to reference a multi‑area named range in a formula and trigger recalculation in Aspose.Cells for .NET. | Provide step‑by‑step instructions to aggregate data from several sheets into one cell using a named range and SUM in Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to build a workbook, add two worksheets, create a named range that references the same cells on both sheets, insert a SUM(TotalData) formula on a Dashboard sheet, recalculate formulas, and save the result using Aspose.Cells for .NET.
class DashboardNamedRange
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a second worksheet (default workbook already has one)
        workbook.Worksheets.Add();
        Worksheet sheet1 = workbook.Worksheets[0];
        Worksheet sheet2 = workbook.Worksheets[1];

        // Populate sample data in both sheets (A1:A2)
        sheet1.Cells["A1"].PutValue(10);
        sheet1.Cells["A2"].PutValue(20);
        sheet2.Cells["A1"].PutValue(30);
        sheet2.Cells["A2"].PutValue(40);

        // Create a named range that refers to the same area on both sheets
        int nameIndex = workbook.Worksheets.Names.Add("TotalData");
        Name totalDataName = workbook.Worksheets.Names[nameIndex];
        // The RefersTo string can contain multiple areas separated by commas
        totalDataName.RefersTo = $"={sheet1.Name}!$A$1:$A$2,{sheet2.Name}!$A$1:$A$2";

        // Add a dashboard sheet to display the aggregated result
        Worksheet dashboard = workbook.Worksheets.Add("Dashboard");
        dashboard.Cells["A1"].PutValue("Aggregated Sum");
        // Use the named range in a SUM formula
        dashboard.Cells["B1"].Formula = "=SUM(TotalData)";

        // Calculate formulas so the result is available
        workbook.CalculateFormula();

        // Output the calculated sum to the console (optional verification)
        Console.WriteLine("Aggregated Sum: " + dashboard.Cells["B1"].Value);

        // Save the workbook
        workbook.Save("DashboardNamedRange.xlsx");
    }
}
