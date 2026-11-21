// Title: Generate a summary worksheet that aggregates totals from multiple smart‑marker sheets using Aspose.Cells for .NET (C#)
// Common Searches: how to create a summary sheet that pulls totals from different worksheets using Aspose.Cells C# | Aspose.Cells formula to reference another sheet's total cell in C# code | C# example for adding a grand total row that sums values from several smart‑marker worksheets with Aspose.Cells | using Aspose.Cells .NET to build a summary worksheet that links to B10 totals on each month sheet | Aspose.Cells tutorial for aggregating smart‑marker sheet totals into a single summary workbook
// Tags: cross-sheet total aggregation Aspose.Cells C# | summary worksheet generation Aspose.Cells .NET | formula reference other worksheet Aspose.Cells | grand total calculation Aspose.Cells workbook | smart marker sheet total consolidation C#

using Aspose.Cells;
using System;

// The example creates a workbook, adds three month worksheets (January, February, March) each containing sample data and a total calculated in cell B10. It then adds a Summary worksheet that lists each month and uses formulas like ='January'!B10 to pull the totals. A Grand Total row sums the listed month totals with a SUM formula. Finally, the workbook is saved as 'SummaryWorkbook.xlsx'.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // -----------------------------------------------------------------
        // Sample data worksheets (in a real scenario these would be filled
        // by smart markers). Each sheet contains a total in cell B10.
        // -----------------------------------------------------------------
        string[] dataSheetNames = { "January", "February", "March" };
        foreach (string sheetName in dataSheetNames)
        {
            // Add a new worksheet and set its name
            Worksheet ws = workbook.Worksheets[workbook.Worksheets.Add()];
            ws.Name = sheetName;

            // Sample data
            ws.Cells["A1"].PutValue("Item");
            ws.Cells["B1"].PutValue("Amount");
            ws.Cells["A2"].PutValue("Product A");
            ws.Cells["B2"].PutValue(120);
            ws.Cells["A3"].PutValue("Product B");
            ws.Cells["B3"].PutValue(250);

            // Total for the sheet (placed in B10)
            ws.Cells["B10"].Formula = "SUM(B2:B3)";
        }

        // -----------------------------------------------------------------
        // Create the summary sheet that aggregates totals from the data sheets
        // -----------------------------------------------------------------
        Worksheet summarySheet = workbook.Worksheets[workbook.Worksheets.Add()];
        summarySheet.Name = "Summary";

        // Header row
        summarySheet.Cells["A1"].PutValue("Month");
        summarySheet.Cells["B1"].PutValue("Total");

        // Populate summary rows with formulas that reference each sheet's total cell (B10)
        for (int i = 0; i < dataSheetNames.Length; i++)
        {
            int rowIndex = i + 2; // Data starts at row 2 (A2, B2)

            // Month name
            summarySheet.Cells[rowIndex, 0].PutValue(dataSheetNames[i]);

            // Formula linking to the corresponding sheet's total
            // Example: ='January'!B10
            string sheetRef = dataSheetNames[i];
            summarySheet.Cells[rowIndex, 1].Formula = $"='{sheetRef}'!B10";
        }

        // Optional: Grand total of all months
        int grandTotalRow = dataSheetNames.Length + 2;
        summarySheet.Cells[grandTotalRow, 0].PutValue("Grand Total");
        // Sum the totals column (B2:B{last data row})
        summarySheet.Cells[grandTotalRow, 1].Formula = $"SUM(B2:B{grandTotalRow - 1})";

        // -----------------------------------------------------------------
        // Save the workbook to a file
        // -----------------------------------------------------------------
        workbook.Save("SummaryWorkbook.xlsx");
    }
}
