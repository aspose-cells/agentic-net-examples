// Title: Hide Column Grand Totals in an Aspose.Cells PivotTable (C#) – ShowColumnGrandTotals Property
// Description: Learn how to create a workbook with sample data, add a PivotTable, assign fields, and disable column grand totals by setting the ShowColumnGrandTotals property to false. The example prints the setting to the console and saves the file as PivotTable_NoColumnGrandTotals.xlsx.
// Keywords: Aspose.Cells hide column grand totals | ShowColumnGrandTotals false C# | Aspose.Cells PivotTable settings | disable column grand totals Aspose.Cells | C# Excel pivot table without column totals | Aspose.Cells for .NET pivot table example
// Common Searches: how to hide column grand totals in Aspose.Cells PivotTable C# | ShowColumnGrandTotals property example | Aspose.Cells pivot table without column totals | C# code to disable column grand totals in Excel pivot | Aspose.Cells hide column grand totals tutorial
// Developer Intent: The developer needs to turn off column grand totals in a PivotTable generated with Aspose.Cells for .NET.
// Use Cases: Create a sales dashboard where column totals clutter the view. | Generate a financial summary that emphasizes individual item values without column aggregates. | Export an Excel report for presentation that omits column grand totals for a cleaner layout.
// AI Prompts: Write C# code using Aspose.Cells to build a PivotTable and hide its column grand totals. | Explain the effect of the ShowColumnGrandTotals property in Aspose.Cells and how to verify it at runtime. | Provide an example that disables both column and row grand totals in an Aspose.Cells PivotTable.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Learn how to create a workbook with sample data, add a PivotTable, assign fields, and disable column grand totals by setting the ShowColumnGrandTotals property to false. The example prints the setting to the console and saves the file as PivotTable_NoColumnGrandTotals.xlsx.
    public class DisableColumnGrandTotalsDemo
    {
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
            sheet.Cells["A1"].Value = "Category";
            sheet.Cells["B1"].Value = "Item";
            sheet.Cells["C1"].Value = "Amount";

            sheet.Cells["A2"].Value = "Fruit";
            sheet.Cells["B2"].Value = "Apple";
            sheet.Cells["C2"].Value = 120;

            sheet.Cells["A3"].Value = "Fruit";
            sheet.Cells["B3"].Value = "Banana";
            sheet.Cells["C3"].Value = 80;

            sheet.Cells["A4"].Value = "Vegetable";
            sheet.Cells["B4"].Value = "Carrot";
            sheet.Cells["C4"].Value = 50;

            sheet.Cells["A5"].Value = "Vegetable";
            sheet.Cells["B5"].Value = "Broccoli";
            sheet.Cells["C5"].Value = 70;

            // Add a pivot table based on the data range A1:C5, place it at E3, and name it "PivotTable1"
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table: Category as row, Item as column, Amount as data
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);      // Category
            pivotTable.AddFieldToArea(PivotFieldType.Column, 1);   // Item
            pivotTable.AddFieldToArea(PivotFieldType.Data, 2);     // Amount

            // Disable column grand totals
            pivotTable.ShowColumnGrandTotals = false;

            // Verify the setting (writes to console)
            Console.WriteLine("ShowColumnGrandTotals: " + pivotTable.ShowColumnGrandTotals);

            // Save the workbook to a file
            workbook.Save("PivotTable_NoColumnGrandTotals.xlsx");
        }
    }
}
