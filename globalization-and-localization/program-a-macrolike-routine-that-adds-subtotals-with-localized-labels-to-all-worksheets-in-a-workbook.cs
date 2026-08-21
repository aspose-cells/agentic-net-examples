// Title: C# – Add Subtotal Rows with Localized Labels to All Worksheets using Aspose.Cells
// Description: Demonstrates how to create or load a workbook, set a custom total name with SettableGlobalizationSettings, loop through every worksheet, define the data range, and call Cells.Subtotal to insert grouped summary rows that display the localized label, then save the file.
// Keywords: Aspose.Cells subtotal C# | localized subtotal label | SettableGlobalizationSettings | custom total name Excel | globalization settings Aspose.Cells | add subtotals to all worksheets | macro‑like subtotal routine | ConsolidationFunction.Sum custom label | Excel automation C# | internationalized subtotal rows
// Common Searches: Aspose.Cells change subtotal label | C# add subtotals to every sheet | SettableGlobalizationSettings example | custom total name for sum function Aspose.Cells | globalize Excel subtotal text C# | macro to add localized subtotals Aspose
// Developer Intent: Create a macro‑style routine that inserts subtotal rows with a custom, localized label into each worksheet of an Aspose.Cells workbook.
// Use Cases: Produce multilingual financial statements where the subtotal caption appears in the target language on every sheet. | Automate consolidation of sales data across dozens of worksheets while enforcing a consistent localized total label. | Prepare a master workbook for international distribution, ensuring all subtotal rows use the same translated wording.
// AI Prompts: Generate C# code that uses Aspose.Cells to set a custom total name for the Sum function and adds subtotal rows to all worksheets. | Explain the role of SettableGlobalizationSettings in overriding subtotal labels and show how to customize other consolidation functions. | Provide a step‑by‑step tutorial for loading an existing workbook, applying a localized subtotal label, adding grouped totals, and saving the result.

using System;
using Aspose.Cells;

namespace SubtotalWithLocalizedLabels
{
    // Demonstrates how to create or load a workbook, set a custom total name with SettableGlobalizationSettings, loop through every worksheet, define the data range, and call Cells.Subtotal to insert grouped summary rows that display the localized label, then save the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // replace with new Workbook("input.xlsx") if needed

            // ------------------------------------------------------------
            // 1. Define custom globalization settings to change the total label
            // ------------------------------------------------------------
            // SettableGlobalizationSettings allows us to override the default text
            // that appears for subtotal rows (e.g., "Sum").
            SettableGlobalizationSettings globalization = new SettableGlobalizationSettings();

            // Example: change the label for the Sum function to a localized version
            globalization.SetTotalName(ConsolidationFunction.Sum, "Σ Total");   // you can use any language string here
            // You can also customize other functions if required:
            // globalization.SetTotalName(ConsolidationFunction.Average, "Avg Total");

            // Apply the custom settings to the workbook
            workbook.Settings.GlobalizationSettings = globalization;

            // ------------------------------------------------------------
            // 2. Add subtotals to every worksheet in the workbook
            // ------------------------------------------------------------
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine the used range of the worksheet
                int maxRow = sheet.Cells.MaxDataRow;      // zero‑based index of the last row with data
                int maxCol = sheet.Cells.MaxDataColumn;   // zero‑based index of the last column with data

                // If the sheet is empty, skip it
                if (maxRow < 0 || maxCol < 0)
                    continue;

                // Define the cell area that contains the data (including header row)
                CellArea area = CellArea.CreateCellArea(0, 0, maxRow, maxCol);

                // Choose the column to group by (typically the first column, index 0)
                int groupByColumn = 0;

                // Choose which columns should receive the subtotal calculation.
                // Here we subtotal the second column (index 1). Adjust as needed.
                int[] totalColumns = new int[] { 1 };

                // Add subtotals:
                //   replace: true  – replace any existing subtotals
                //   pageBreaks: false – do not insert page breaks between groups
                //   summaryBelowData: true – place the subtotal row below each group
                sheet.Cells.Subtotal(
                    area,
                    groupByColumn,
                    ConsolidationFunction.Sum,
                    totalColumns,
                    true,   // replace existing subtotals
                    false,  // no page breaks
                    true    // summary below data
                );
            }

            // ------------------------------------------------------------
            // 3. Save the workbook
            // ------------------------------------------------------------
            workbook.Save("WorkbookWithLocalizedSubtotals.xlsx");
        }
    }
}
