using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class TestGrandTotalAndSubtotalLabels
    {
        public static void Run()
        {
            try
            {
                // ---------- Create workbook and sample data ----------
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];
                var cells = sheet.Cells;

                // Header
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Sales");

                // Data rows
                cells["A2"].PutValue("North");
                cells["B2"].PutValue(1000);
                cells["A3"].PutValue("North");
                cells["B3"].PutValue(1500);
                cells["A4"].PutValue("South");
                cells["B4"].PutValue(2000);
                cells["A5"].PutValue("South");
                cells["B5"].PutValue(2500);

                // ---------- Create pivot table ----------
                // The pivot table will be placed starting at cell D1.
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D1", "PivotTable1");
                var pivot = sheet.PivotTables[pivotIndex];

                // Row field: Category
                pivot.AddFieldToArea(PivotFieldType.Row, 0);
                // Data field: Sales (default function is Sum)
                pivot.AddFieldToArea(PivotFieldType.Data, 1);

                // Refresh and calculate to generate the pivot data.
                pivot.RefreshData();
                pivot.CalculateData();

                // ---------- Output verification ----------
                Console.WriteLine("Pivot table created successfully.");

                // ---------- Save the workbook ----------
                string outputPath = "GrandTotalAndSubtotalTest.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point
    class Program
    {
        static void Main()
        {
            TestGrandTotalAndSubtotalLabels.Run();
        }
    }
}