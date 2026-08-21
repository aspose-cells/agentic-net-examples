// Title: Add Subtotal Rows with a Custom SUM Total Label Using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, define a custom SettableGlobalizationSettings class that overrides the SUM total label, apply the setting, insert subtotal rows grouped by a column, programmatically verify the custom label "Custom Sum Total", and save the file.
// Keywords: Aspose.Cells subtotal custom label | SettableGlobalizationSettings C# | override GetTotalName Aspose.Cells | custom SUM total name | globalization localization workbook | C# add subtotal rows | verify subtotal label | Aspose.Cells .NET example
// Common Searches: custom subtotal total label Aspose.Cells C# | how to change SUM total name in Aspose.Cells | add subtotal rows with globalization settings .NET | verify custom total label after subtotal Aspose.Cells | override GetTotalName for localization in Aspose.Cells
// Developer Intent: Insert subtotal rows while applying a custom globalization setting that changes the default SUM total label, then confirm the label appears in the worksheet.
// Use Cases: Replace the default "Sum Total" text with a localized label such as "Custom Sum Total" for multilingual reports. | Generate grouped subtotals in financial or sales data and ensure the custom total label complies with corporate terminology. | Automate validation of customized subtotal labels before distributing workbooks to end users.
// AI Prompts: Write C# code with Aspose.Cells that adds subtotal rows, uses a custom SettableGlobalizationSettings to rename the SUM total label to "Custom Sum Total", and checks for the label in the worksheet. | Explain how to override GetTotalName in SettableGlobalizationSettings to provide different total names for consolidation functions in Aspose.Cells. | Provide a step‑by‑step tutorial for adding subtotals, applying custom globalization, and programmatically verifying the custom total label in a .NET workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsSubtotalDemo
{
    // Custom globalization settings that changes the total label for the SUM function
    // Demonstrates how to create a workbook, define a custom SettableGlobalizationSettings class that overrides the SUM total label, apply the setting, insert subtotal rows grouped by a column, programmatically verify the custom label "Custom Sum Total", and save the file.
    public class CustomGlobalizationSettings : SettableGlobalizationSettings
    {
        public override string GetTotalName(ConsolidationFunction functionType)
        {
            // Return a custom label for SUM subtotals; other functions use the default label
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

            // Populate sample data (Header + 4 data rows)
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Amount");
            cells["A2"].PutValue("North");
            cells["B2"].PutValue(1000);
            cells["A3"].PutValue("North");
            cells["B3"].PutValue(1500);
            cells["A4"].PutValue("South");
            cells["B4"].PutValue(2000);
            cells["A5"].PutValue("South");
            cells["B5"].PutValue(2500);

            // Apply custom globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

            // Define the range that contains the data (including header)
            CellArea dataArea = CellArea.CreateCellArea(0, 0, 4, 1); // rows 0‑4, columns 0‑1

            // Add subtotal rows:
            // - Group by column 0 (Category)
            // - Use SUM function
            // - Apply subtotal to column 1 (Amount)
            // - Replace existing subtotals, add page breaks, place summary below data
            cells.Subtotal(
                dataArea,
                0,
                ConsolidationFunction.Sum,
                new int[] { 1 },
                true,
                true,
                true);

            // Verify that the custom total label appears in the worksheet
            bool labelFound = false;
            int maxRow = cells.MaxDataRow + 10; // safety margin
            for (int row = 0; row <= maxRow; row++)
            {
                if (cells[row, 0].StringValue == "Custom Sum Total")
                {
                    Console.WriteLine($"Custom total label found at row {row}");
                    labelFound = true;
                    break;
                }
            }

            if (!labelFound)
                Console.WriteLine("Custom total label not found.");

            // Save the workbook
            workbook.Save("SubtotalWithCustomGlobalization.xlsx");
        }
    }
}
