// Title: How to Disable Row Grand Totals in an Aspose.Cells Pivot Table (C#)
// Description: Creates a workbook, fills it with sample sales data, adds a pivot table, puts Product in the row area and Sales in the data area, sets ShowRowGrandTotals to false, recalculates the pivot, and saves the file without row grand totals.
// Keywords: Aspose.Cells C# hide row grand total | ShowRowGrandTotals false | disable row grand totals pivot table | Aspose.Cells pivot table grand total settings | remove row grand total Excel C# | pivot table formatting Aspose.Cells
// Common Searches: Aspose.Cells turn off row grand totals | C# pivot table hide row grand total | ShowRowGrandTotals property example | disable row grand totals in Excel using Aspose | pivot table without row totals Aspose.Cells
// Developer Intent: Create a pivot table and suppress its row grand total line.
// Use Cases: Produce a sales summary where row totals are unnecessary for a cleaner layout. | Design an Excel dashboard that shows product‑level figures without aggregated row totals. | Export pivot data for downstream processing while omitting row grand totals to simplify analysis.
// AI Prompts: Generate C# code that builds an Aspose.Cells pivot table and disables row grand totals. | Explain how the ShowRowGrandTotals property works and why recalculating the pivot is required. | Show how to hide both row and column grand totals in an Aspose.Cells pivot table.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Creates a workbook, fills it with sample sales data, adds a pivot table, puts Product in the row area and Sales in the data area, sets ShowRowGrandTotals to false, recalculates the pivot, and saves the file without row grand totals.
    public class DisableRowGrandTotalsDemo
    {
        // Entry point required for compilation
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].Value = "Product";
            sheet.Cells["B1"].Value = "Region";
            sheet.Cells["C1"].Value = "Sales";

            sheet.Cells["A2"].Value = "Product A";
            sheet.Cells["B2"].Value = "North";
            sheet.Cells["C2"].Value = 1000;

            sheet.Cells["A3"].Value = "Product B";
            sheet.Cells["B3"].Value = "South";
            sheet.Cells["C3"].Value = 2000;

            sheet.Cells["A4"].Value = "Product A";
            sheet.Cells["B4"].Value = "East";
            sheet.Cells["C4"].Value = 1500;

            sheet.Cells["A5"].Value = "Product B";
            sheet.Cells["B5"].Value = "West";
            sheet.Cells["C5"].Value = 2500;

            // Add a pivot table based on the data range
            PivotTableCollection pivotTables = sheet.PivotTables;
            int pivotIndex = pivotTables.Add("A1:C5", "E3", "PivotTable1");
            PivotTable pivotTable = pivotTables[pivotIndex];

            // Configure the pivot table: add row and data fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Disable row grand totals to simplify the summary view
            pivotTable.ShowRowGrandTotals = false;

            // Recalculate the pivot table after changing the setting
            pivotTable.CalculateData();

            // Save the workbook with the modified pivot table
            workbook.Save("PivotTable_NoRowGrandTotals.xlsx");
        }
    }
}
