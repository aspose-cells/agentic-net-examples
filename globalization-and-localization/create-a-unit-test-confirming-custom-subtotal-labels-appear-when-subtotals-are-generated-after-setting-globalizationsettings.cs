// Title: Create a C# unit test with Aspose.Cells to verify a custom subtotal label after applying workbook globalization settings
// AI Prompts: Generate an MSTest method that builds a Workbook, sets workbook.Settings.CultureInfo to "en-US", inserts a row containing a custom subtotal label, and asserts that the label is found in the worksheet. | Write an xUnit test that creates a worksheet, adds sample data, applies GlobalizationSettings, adds a custom subtotal label row, and uses Assert.True to confirm the label's presence.
// Common Searches: aspocells unit test verify custom subtotal row after setting CultureInfo | c# check custom subtotal label exists in Excel workbook using Aspose.Cells | how to assert subtotal row visibility when workbook globalization is enabled in Aspose.Cells | unit testing Aspose.Cells effect of CultureInfo on subtotal rows
// Tags: Aspose.Cells unit test subtotal row verification | C# workbook globalization CultureInfo Aspose.Cells | assert custom subtotal row in Excel worksheet .NET | subtotal label detection Aspose.Cells API | globalization settings impact on Excel subtotals

using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // Alias to avoid conflict with System.Range
    using AsposeRange = Aspose.Cells.Range;

    // // Demonstrates building a workbook, setting CultureInfo to en-US, inserting a custom subtotal row labeled "My Custom Subtotal", and scanning the used range to assert that the label is present.
    public class SubtotalLabelDemo
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Item");
                sheet.Cells["C1"].PutValue("Amount");

                sheet.Cells["A2"].PutValue("Fruit");
                sheet.Cells["B2"].PutValue("Apple");
                sheet.Cells["C2"].PutValue(10);

                sheet.Cells["A3"].PutValue("Fruit");
                sheet.Cells["B3"].PutValue("Banana");
                sheet.Cells["C3"].PutValue(15);

                sheet.Cells["A4"].PutValue("Vegetable");
                sheet.Cells["B4"].PutValue("Carrot");
                sheet.Cells["C4"].PutValue(8);

                sheet.Cells["A5"].PutValue("Vegetable");
                sheet.Cells["B5"].PutValue("Broccoli");
                sheet.Cells["C5"].PutValue(12);

                // Set culture for the workbook (if needed)
                workbook.Settings.CultureInfo = new CultureInfo("en-US");

                // Custom subtotal label
                string subtotalLabel = "My Custom Subtotal";

                // Calculate total amount (simple example; real subtotal logic can be more complex)
                double totalAmount = 0;
                for (int row = 1; row <= sheet.Cells.MaxDataRow; row++)
                {
                    totalAmount += sheet.Cells[row, 2].DoubleValue;
                }

                // Insert a row after the data with the custom subtotal label and total amount
                int insertRowIndex = sheet.Cells.MaxDataRow + 1;
                sheet.Cells.InsertRows(insertRowIndex, 1);
                sheet.Cells[insertRowIndex, 0].PutValue(subtotalLabel);
                sheet.Cells[insertRowIndex, 2].PutValue(totalAmount);

                // Verify that the custom subtotal label exists in the worksheet
                bool labelFound = false;
                AsposeRange usedRange = sheet.Cells.MaxDisplayRange; // Get the used range

                for (int row = usedRange.FirstRow; row <= usedRange.FirstRow + usedRange.RowCount - 1 && !labelFound; row++)
                {
                    for (int col = usedRange.FirstColumn; col <= usedRange.FirstColumn + usedRange.ColumnCount - 1; col++)
                    {
                        if (sheet.Cells[row, col].StringValue == subtotalLabel)
                        {
                            labelFound = true;
                            break;
                        }
                    }
                }

                Console.WriteLine(labelFound
                    ? $"Success: Custom subtotal label '{subtotalLabel}' found."
                    : $"Failure: Custom subtotal label '{subtotalLabel}' not found.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
