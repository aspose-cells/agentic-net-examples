// Title: C# Unit Test: Verify Custom Pivot Subtotal Labels with GlobalizationSettings in Aspose.Cells
// Description: Creates a workbook, builds a pivot table with automatic row subtotals, applies SettablePivotGlobalizationSettings to rename the Sum and Average subtotal labels, refreshes the pivot, and scans the worksheet to confirm that "My Custom Sum" and "My Custom Avg" appear. Throws an exception if the custom labels are missing and optionally saves the file for manual review.
// Keywords: Aspose.Cells | C# | .NET | pivot table | custom subtotal label | SettablePivotGlobalizationSettings | globalization settings | localization test | unit test | automated verification
// Common Searches: how to test custom subtotal text in Aspose.Cells pivot table | unit test for GlobalizationSettings in Aspose.Cells .NET | verify custom pivot subtotal labels programmatically | Aspose.Cells localization unit testing | C# test for custom sum/average subtotal names
// Developer Intent: Ensure that custom subtotal labels defined via GlobalizationSettings are correctly rendered in a generated pivot table.
// Use Cases: Automated regression test for localized pivot subtotal labels across multiple cultures. | CI/CD validation that changes to GlobalizationSettings do not break custom subtotal text. | Sample code for developers needing to assert pivot table label customizations in unit tests.
// AI Prompts: Generate an MSTest method that reproduces the example and asserts the presence of "My Custom Sum" and "My Custom Avg" in the pivot table. | Create an xUnit test for Aspose.Cells that sets SettablePivotGlobalizationSettings and verifies custom subtotal labels. | Write a NUnit test case that applies GlobalizationSettings to a pivot table and checks for custom subtotal text.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Settings;

namespace AsposeCellsExamples
{
    // Creates a workbook, builds a pivot table with automatic row subtotals, applies SettablePivotGlobalizationSettings to rename the Sum and Average subtotal labels, refreshes the pivot, and scans the worksheet to confirm that "My Custom Sum" and "My Custom Avg" appear. Throws an exception if the custom labels are missing and optionally saves the file for manual review.
    public class PivotSubtotalLabelDemo
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and add sample data
                var workbook = new Workbook();
                var ws = workbook.Worksheets[0];
                var cells = ws.Cells;

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

                // Create a pivot table based on the data range
                int pivotIdx = ws.PivotTables.Add("A1:B5", "D1", "PivotTable1");
                var pivot = ws.PivotTables[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Category field
                pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Value field

                // Enable automatic subtotals for the row field
                pivot.RowFields[0].IsAutoSubtotals = true;

                // Create custom globalization settings for pivot subtotals
                var pivotGSettings = new SettablePivotGlobalizationSettings();
                pivotGSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Sum, "My Custom Sum");
                pivotGSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Average, "My Custom Avg");

                // Apply the custom settings to the workbook
                workbook.Settings.GlobalizationSettings.PivotSettings = pivotGSettings;

                // Refresh pivot data and calculate results
                pivot.RefreshData();
                pivot.CalculateData();

                // Verify that the custom subtotal labels appear in the worksheet
                bool foundSum = false;
                bool foundAvg = false;

                // Scan the used range of the worksheet for the expected texts
                var usedRange = ws.Cells.MaxDisplayRange;
                for (int row = usedRange.FirstRow; row < usedRange.FirstRow + usedRange.RowCount; row++)
                {
                    for (int col = usedRange.FirstColumn; col < usedRange.FirstColumn + usedRange.ColumnCount; col++)
                    {
                        string cellText = ws.Cells[row, col].StringValue;
                        if (cellText == "My Custom Sum") foundSum = true;
                        if (cellText == "My Custom Avg") foundAvg = true;
                    }
                }

                if (!foundSum || !foundAvg)
                {
                    throw new Exception("Custom subtotal labels were not found in the pivot table.");
                }

                // Save the workbook (optional, for manual inspection)
                string outputPath = "CustomSubtotalLabels.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
