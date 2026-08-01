// Title: Add Subtotals with Localized Labels to All Worksheets using Aspose.Cells C#
// Description: Creates a workbook with three sheets, fills each with sample data, sets a custom global label for the Sum subtotal via SettableGlobalizationSettings, then iterates all worksheets and applies the Subtotal method (group by first column, page breaks, summary below) and saves the file.
// Keywords: Aspose.Cells | C# | add subtotals | localized subtotal label | SettableGlobalizationSettings | globalization settings | Subtotal method | multiple worksheets | group by column | page breaks | Excel macro replacement | financial report automation
// Common Searches: Aspose.Cells set custom subtotal label .NET | How to add subtotals to every sheet in Aspose.Cells | Globalize subtotal function in C# workbook | Iterate worksheets and apply Subtotal in Aspose.Cells | Replace Excel subtotal macro with C# code
// Developer Intent: Generate a C# routine that applies subtotal rows with a custom localized label to each worksheet in an Aspose.Cells workbook.
// Use Cases: Create multi‑sheet financial statements where each sheet shows grouped totals with a language‑specific “Sum” label. | Consolidate regional sales data across worksheets, automatically inserting subtotals and page breaks for printed reports. | Migrate Excel VBA subtotal macros to a .NET solution that respects user‑locale settings.
// AI Prompts: Write a C# method that receives a label string and uses SettableGlobalizationSettings to rename the Sum total before calling Worksheet.Cells.Subtotal on all sheets. | Extend the subtotal routine to handle multiple subtotal columns and different functions (Average, Count) while keeping the localized total name. | Add detailed logging to the worksheet loop that records sheet names, success status, and any exceptions for troubleshooting.

using System;
using Aspose.Cells;
using Aspose.Cells.Settings;

namespace SubtotalMacroDemo
{
    // Creates a workbook with three sheets, fills each with sample data, sets a custom global label for the Sum subtotal via SettableGlobalizationSettings, then iterates all worksheets and applies the Subtotal method (group by first column, page breaks, summary below) and saves the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (contains one default worksheet)
                Workbook workbook = new Workbook();

                // Rename the default worksheet and populate it with sample data
                Worksheet firstSheet = workbook.Worksheets[0];
                firstSheet.Name = "Sheet1";
                PopulateSampleData(firstSheet);

                // Add two more worksheets and fill them with the same sample data
                for (int i = 1; i < 3; i++)
                {
                    // Worksheets.Add() returns the index of the newly added sheet
                    int newIndex = workbook.Worksheets.Add();
                    Worksheet ws = workbook.Worksheets[newIndex];
                    ws.Name = $"Sheet{i + 1}";
                    PopulateSampleData(ws);
                }

                // Apply custom localization for subtotal labels
                SettableGlobalizationSettings globalization = new SettableGlobalizationSettings();
                globalization.SetTotalName(ConsolidationFunction.Sum, "Localized Sum");
                workbook.Settings.GlobalizationSettings = globalization;

                // Add subtotals to every worksheet in the workbook
                AddSubtotalsToAllWorksheets(workbook);

                // Save the workbook
                workbook.Save("SubtotalsWithLocalizedLabels.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Adds subtotals to each worksheet in the provided workbook
        private static void AddSubtotalsToAllWorksheets(Workbook workbook)
        {
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Determine the used range of the worksheet
                    int maxRow = sheet.Cells.MaxDataRow;
                    int maxCol = sheet.Cells.MaxDataColumn;

                    // Skip empty sheets
                    if (maxRow < 0 || maxCol < 0)
                        continue;

                    // Define the cell area covering the data (including header row)
                    CellArea area = new CellArea
                    {
                        StartRow = 0,
                        StartColumn = 0,
                        EndRow = maxRow,
                        EndColumn = maxCol
                    };

                    // Subtotal the second column (index 1) if it exists; otherwise the first column
                    int[] totalList = maxCol >= 1 ? new int[] { 1 } : new int[] { 0 };

                    // Apply subtotal:
                    // - Group by the first column (index 0)
                    // - Use Sum function
                    // - Replace existing subtotals, add page breaks, place summary below data
                    sheet.Cells.Subtotal(
                        area,
                        0,                                 // groupBy column index
                        ConsolidationFunction.Sum,         // subtotal function
                        totalList,                         // columns to subtotal
                        true,                              // replace existing subtotals
                        true,                              // add page breaks between groups
                        true                               // place summary below data
                    );
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to add subtotals on sheet '{sheet.Name}': {ex.Message}");
                }
            }
        }

        // Helper method to fill a worksheet with simple sample data
        private static void PopulateSampleData(Worksheet sheet)
        {
            // Header
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Amount");

            // Sample rows
            object[,] data = new object[,]
            {
                { "North", 1200 },
                { "North", 800 },
                { "South", 1500 },
                { "South", 700 },
                { "East",  1100 },
                { "West",  900 }
            };

            for (int r = 0; r < data.GetLength(0); r++)
            {
                sheet.Cells[r + 1, 0].PutValue(data[r, 0]); // Category
                sheet.Cells[r + 1, 1].PutValue(data[r, 1]); // Amount
            }
        }
    }
}
