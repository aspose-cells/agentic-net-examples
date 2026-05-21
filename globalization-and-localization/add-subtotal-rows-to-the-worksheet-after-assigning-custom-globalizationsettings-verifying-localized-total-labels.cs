using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsSubtotalDemo
{
    // Custom globalization settings that overrides the total name for the SUM function
    public class CustomGlobalizationSettings : SettableGlobalizationSettings
    {
        public override string GetTotalName(ConsolidationFunction functionType)
        {
            // Return a custom label for SUM totals; other functions use the base implementation
            if (functionType == ConsolidationFunction.Sum)
                return "Custom Sum Total";
            return base.GetTotalName(functionType);
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data (Category | Value)
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Value");
            cells["A2"].PutValue("A");
            cells["B2"].PutValue(10);
            cells["A3"].PutValue("A");
            cells["B3"].PutValue(20);
            cells["A4"].PutValue("B");
            cells["B4"].PutValue(30);
            cells["A5"].PutValue("B");
            cells["B5"].PutValue(40);
            cells["A6"].PutValue("C");
            cells["B6"].PutValue(50);

            // Assign the custom globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

            // Define the range that contains the data (including header)
            CellArea dataArea = CellArea.CreateCellArea(0, 0, 5, 1); // rows 0-5, columns 0-1

            // Add subtotal rows:
            // - Group by column 0 (Category)
            // - Use SUM function
            // - Apply subtotal to column 1 (Value)
            // - Replace existing subtotals, no page breaks, place summary below data
            cells.Subtotal(
                dataArea,
                groupBy: 0,
                function: ConsolidationFunction.Sum,
                totalList: new int[] { 1 },
                replace: true,
                pageBreaks: false,
                summaryBelowData: true);

            // Verify that the subtotal label uses the custom total name
            // Search for the custom label in the worksheet
            string customLabel = "Custom Sum Total";
            Cell foundCell = cells.Find(customLabel, null, new FindOptions() { LookInType = LookInType.Values });

            if (foundCell != null)
            {
                Console.WriteLine($"Verified custom total label found at {foundCell.Name}: \"{foundCell.StringValue}\"");
            }
            else
            {
                Console.WriteLine("Custom total label not found. Verification failed.");
            }

            // Save the workbook
            workbook.Save("SubtotalWithCustomGlobalization.xlsx");
        }
    }
}