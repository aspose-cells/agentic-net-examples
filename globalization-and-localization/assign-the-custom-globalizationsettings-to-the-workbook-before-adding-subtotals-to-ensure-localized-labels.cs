// Title: C# – Apply Custom GlobalizationSettings Before Subtotal to Localize Total Labels in Aspose.Cells
// Description: Demonstrates how to create a CustomGlobalizationSettings class that overrides GetTotalName, assign it to workbook.Settings.GlobalizationSettings, and then call cells.Subtotal so the generated subtotal rows use custom or localized total names. The workbook is populated with sample data and saved as CustomGlobalization_Subtotal.xlsx.
// Keywords: Aspose.Cells | C# | CustomGlobalizationSettings | GlobalizationSettings | GetTotalName override | subtotal localized labels | Excel total name customization | multilingual Excel reports | cells.Subtotal | workbook.Settings.GlobalizationSettings
// Common Searches: Aspose.Cells customize subtotal total name | set GlobalizationSettings before Subtotal in .NET | override GetTotalName for localized labels | C# Aspose.Cells subtotal custom labels | how to localize Excel subtotal rows with Aspose
// Developer Intent: Assign a custom GlobalizationSettings object to a workbook prior to adding subtotals so that the subtotal rows display user‑defined or localized total names.
// Use Cases: Create financial summaries where subtotal rows show language‑specific labels such as "Custom Sum" or "Suma Personalizada". | Generate reports for multinational audiences by providing translated total names for Sum, Average, Count, Max, and Min. | Build pivot‑style groupings in a .NET application with custom total descriptors without modifying the data source.
// AI Prompts: Write C# code that defines a CustomGlobalizationSettings class overriding GetTotalName for Sum, Average, Count, Max, and Min, assigns it to workbook.Settings.GlobalizationSettings, and then adds subtotals using cells.Subtotal. | Explain why GlobalizationSettings must be set before calling Subtotal in Aspose.Cells and how it changes the text of generated subtotal rows. | Show how to extend GlobalizationSettings to return culture‑specific strings for total names and demonstrate its usage in a workbook that creates subtotals.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsGlobalizationDemo
{
    // Custom globalization settings that provide localized total/subtotal names
    // Demonstrates how to create a CustomGlobalizationSettings class that overrides GetTotalName, assign it to workbook.Settings.GlobalizationSettings, and then call cells.Subtotal so the generated subtotal rows use custom or localized total names. The workbook is populated with sample data and saved as CustomGlobalization_Subtotal.xlsx.
    public class CustomGlobalizationSettings : GlobalizationSettings
    {
        // Override the method that returns the name for a given consolidation function.
        // This name is used by the Subtotal operation to label the generated rows.
        public override string GetTotalName(ConsolidationFunction functionType)
        {
            return functionType switch
            {
                ConsolidationFunction.Sum => "Custom Sum",
                ConsolidationFunction.Average => "Custom Average",
                ConsolidationFunction.Count => "Custom Count",
                ConsolidationFunction.Max => "Custom Max",
                ConsolidationFunction.Min => "Custom Min",
                _ => base.GetTotalName(functionType)
            };
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Amount");
            cells["A2"].PutValue("Food");
            cells["B2"].PutValue(120);
            cells["A3"].PutValue("Food");
            cells["B3"].PutValue(80);
            cells["A4"].PutValue("Drink");
            cells["B4"].PutValue(150);
            cells["A5"].PutValue("Drink");
            cells["B5"].PutValue(200);
            cells["A6"].PutValue("Other");
            cells["B6"].PutValue(50);

            // Assign the custom globalization settings BEFORE adding subtotals
            workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

            // Define the range to subtotal (rows 0‑5, columns 0‑1)
            CellArea area = CellArea.CreateCellArea(0, 0, 5, 1);

            // Add subtotals:
            //   - Group by column 0 (Category)
            //   - Use Sum as the consolidation function
            //   - Replace existing subtotals if any, keep summary rows, and keep grand total
            cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 0 }, true, false, true);

            // Save the workbook
            workbook.Save("CustomGlobalization_Subtotal.xlsx");
        }
    }
}
