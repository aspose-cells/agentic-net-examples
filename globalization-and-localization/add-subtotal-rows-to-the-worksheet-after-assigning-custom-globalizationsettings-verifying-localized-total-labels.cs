using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsSubtotalDemo
{
    // Custom globalization settings that changes the total label for the Sum function
    public class CustomGlobalizationSettings : SettableGlobalizationSettings
    {
        // Override the total name for the Sum function
        public override string GetTotalName(ConsolidationFunction functionType)
        {
            if (functionType == ConsolidationFunction.Sum)
                return "Custom Sum Total"; // localized label
            return base.GetTotalName(functionType);
        }
    }

    public class Program
    {
        public static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ---------- Populate sample data ----------
            // Header
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Amount");

            // Data rows
            string[] categories = { "Food", "Food", "Transport", "Transport", "Utilities" };
            double[] amounts = { 120.5, 80.0, 50.75, 60.25, 100.0 };

            for (int i = 0; i < categories.Length; i++)
            {
                cells[i + 1, 0].PutValue(categories[i]);   // Column A
                cells[i + 1, 1].PutValue(amounts[i]);     // Column B
            }

            // ---------- Assign custom globalization settings ----------
            workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

            // ---------- Add subtotal rows ----------
            // Define the range that contains the data (including header)
            CellArea dataArea = CellArea.CreateCellArea(0, 0, categories.Length, 1);
            // Group by the first column (Category), sum the second column (Amount)
            // Replace existing subtotals = true, add page breaks = false, summary below data = true
            cells.Subtotal(dataArea, 0, ConsolidationFunction.Sum, new int[] { 1 }, true, false, true);

            // ---------- Verify the localized total label ----------
            // After Subtotal, a total row is inserted after each group.
            // Find the first total row (it will be after the first group of "Food").
            // The total label appears in the first column of the total row.
            // Scan rows to locate a cell that contains the custom label.
            string expectedLabel = "Custom Sum Total";
            bool labelFound = false;
            for (int row = 0; row <= sheet.Cells.MaxDataRow; row++)
            {
                string cellValue = sheet.Cells[row, 0].StringValue;
                if (cellValue == expectedLabel)
                {
                    Console.WriteLine($"Localized total label found at row {row + 1}: {cellValue}");
                    labelFound = true;
                    break;
                }
            }

            if (!labelFound)
                Console.WriteLine("Localized total label was not found.");

            // ---------- Save the workbook ----------
            workbook.Save("SubtotalWithCustomGlobalization.xlsx");
        }
    }
}