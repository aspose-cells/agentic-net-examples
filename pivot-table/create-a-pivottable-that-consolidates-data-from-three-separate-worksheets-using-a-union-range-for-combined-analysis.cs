// Title: Create a Consolidated PivotTable from Multiple Worksheets with Union Ranges – Aspose.Cells for .NET (C#)
// Description: A complete C# example that builds a new workbook, adds three worksheets with sample data, defines a union range (A1:B5 on each sheet), and creates a PivotTable on a fourth sheet using Aspose.Cells' multiple‑consolidation‑range feature. The pivot shows "Category" as rows and the summed "Value" as data, then refreshes, calculates, and saves the file as ConsolidatedPivot.xlsx.
// Keywords: Aspose.Cells | C# | .NET | PivotTable | multiple consolidation range | union range | combine worksheets | data consolidation | Excel automation | pivot example | sample code | GitHub | Aspose.Cells tutorial
// Common Searches: Aspose.Cells create pivot table from several sheets | multiple consolidation range pivot C# | union range pivot Aspose.Cells .NET | combine data from multiple worksheets into one pivot | Aspose.Cells sample code for consolidated pivot
// Developer Intent: Generate a PivotTable that aggregates data from three worksheets using a union (multiple‑consolidation) range.
// Use Cases: Merge regional sales sheets into a single executive‑level pivot report. | Consolidate monthly budget worksheets for a company‑wide financial overview. | Aggregate inventory lists from multiple warehouse tabs for unified stock analysis.
// AI Prompts: Add a column field to the consolidated pivot table in this Aspose.Cells example. | Generate code that builds the union range dynamically from an array of worksheet names. | Explain how to disable the auto‑page option and specify custom page fields for a multiple‑consolidation pivot.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace ConsolidatedPivotExample
{
    // A complete C# example that builds a new workbook, adds three worksheets with sample data, defines a union range (A1:B5 on each sheet), and creates a PivotTable on a fourth sheet using Aspose.Cells' multiple‑consolidation‑range feature. The pivot shows "Category" as rows and the summed "Value" as data, then refreshes, calculates, and saves the file as ConsolidatedPivot.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------- Prepare source worksheets --------------------
            // Worksheet 1
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";
            FillData(sheet1);

            // Worksheet 2
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
            FillData(sheet2);

            // Worksheet 3
            Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");
            FillData(sheet3);

            // -------------------- Create worksheet for the PivotTable --------------------
            Worksheet pivotSheet = workbook.Worksheets.Add("ConsolidatedPivot");
            PivotTableCollection pivotTables = pivotSheet.PivotTables;

            // Define the union (consolidation) ranges from the three source sheets
            string[] sourceRanges = new string[]
            {
                "=Sheet1!A1:B5",
                "=Sheet2!A1:B5",
                "=Sheet3!A1:B5"
            };

            // Add a PivotTable using the multiple consolidation ranges.
            // isAutoPage = true (auto creates a single page field, pageFields ignored)
            int pivotIndex = pivotTables.Add(sourceRanges, true, null, "A1", "CombinedPivot");
            PivotTable pivot = pivotTables[pivotIndex];

            // Configure the PivotTable fields
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");   // Row field
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");    // Data field (sum)

            // Refresh and calculate the PivotTable data
            pivot.RefreshData();
            pivot.CalculateData();

            // Save the workbook
            workbook.Save("ConsolidatedPivot.xlsx");
        }

        // Helper method to populate a worksheet with identical sample data
        private static void FillData(Worksheet sheet)
        {
            Cells cells = sheet.Cells;
            // Header
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Value");

            // Sample rows
            string[] categories = { "A", "B", "A", "C", "B" };
            int[] values = { 100, 200, 150, 300, 250 };

            for (int i = 0; i < categories.Length; i++)
            {
                cells[i + 2, 0].PutValue(categories[i]); // Column A
                cells[i + 2, 1].PutValue(values[i]);    // Column B
            }
        }
    }
}
