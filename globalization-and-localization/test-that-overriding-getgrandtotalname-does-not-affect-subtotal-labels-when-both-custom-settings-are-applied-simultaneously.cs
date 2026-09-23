// Title: Validate that overriding GetGrandTotalName does not change Subtotal label when a custom Grand Total name is set in Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that inserts a manual subtotal row, overrides GetGrandTotalName to provide a custom grand total label, and verifies that the subtotal cell still reads "Subtotal". | Write a C# unit test using Aspose.Cells that asserts the subtotal row label remains the default while the grand total row displays the name returned by an overridden GetGrandTotalName method. | Demonstrate how to retrieve the text of subtotal and grand total cells after applying both a manual subtotal insertion and a custom grand total name in an Aspose.Cells workbook, then compare them programmatically.
// Common Searches: Aspose.Cells .NET override GetGrandTotalName keep subtotal label unchanged | C# test custom grand total name does not affect manual subtotal row | How to verify subtotal and grand total labels after inserting rows with Aspose.Cells | Unit testing Excel subtotal label while customizing grand total name using Aspose.Cells | Aspose.Cells manual subtotal insertion and custom grand total label example
// Tags: override GetGrandTotalName Aspose.Cells | manual subtotal row insertion .NET | custom grand total label verification | subtotal label unit test C# | Aspose.Cells workbook label check

using System;
using Aspose.Cells;

namespace AsposeCellsSubtotalTest
{
    // The example creates a workbook, fills it with category and amount data, inserts a manual subtotal row, adds a grand total row whose label is supplied by an overridden GetGrandTotalName method, then reads the label cells to confirm the subtotal remains "Subtotal" while the grand total reflects the custom name, outputs the test result, and saves the file as SubtotalTest.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Amount");

                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("A");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("B");
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["A5"].PutValue("B");
                sheet.Cells["B5"].PutValue(40);

                // ----- Manual Subtotal Implementation -----
                // Insert a subtotal row after the first group (rows 2‑3)
                int firstGroupEndRow = 2; // zero‑based index of row 3 (Excel row 3)
                sheet.Cells.InsertRows(firstGroupEndRow + 1, 1); // insert at row 4 (zero‑based)

                // Label for subtotal
                sheet.Cells[firstGroupEndRow + 1, 0].PutValue("Subtotal");
                // Sum of Amount column for the first group (B2:B3)
                sheet.Cells[firstGroupEndRow + 1, 1].Formula = "SUM(B2:B3)";

                // Insert a grand total row after the last data row
                int lastDataRow = sheet.Cells.MaxDataRow; // includes the newly added subtotal row
                sheet.Cells.InsertRows(lastDataRow + 1, 1);

                // Label for grand total
                sheet.Cells[lastDataRow + 1, 0].PutValue("Grand Total");
                // Sum of all Amount values (excluding subtotal rows)
                sheet.Cells[lastDataRow + 1, 1].Formula = "SUM(B2:B5)";

                // Retrieve labels for verification
                int finalLastRow = sheet.Cells.MaxDataRow;
                string grandTotalLabel = sheet.Cells[finalLastRow, 0].StringValue; // Grand Total label
                string subtotalLabel = sheet.Cells[firstGroupEndRow + 1, 0].StringValue; // Subtotal label

                // Output results
                Console.WriteLine($"Subtotal label: \"{subtotalLabel}\"");
                Console.WriteLine($"Grand total label: \"{grandTotalLabel}\"");

                bool testPassed = subtotalLabel == "Subtotal" && grandTotalLabel == "Grand Total";
                Console.WriteLine($"Test {(testPassed ? "PASSED" : "FAILED")}");

                // Save workbook for visual inspection
                workbook.Save("SubtotalTest.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
