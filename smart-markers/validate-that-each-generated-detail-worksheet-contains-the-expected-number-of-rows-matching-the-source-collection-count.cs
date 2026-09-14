// Title: Validate row count of each detail worksheet against a source collection using Aspose.Cells for .NET
// AI Prompts: Create a C# method that iterates through all worksheets in an Aspose.Cells workbook, compares each sheet's RowCollection.Count to a supplied expected count, and writes pass/fail messages to the console. | Write C# code that fills multiple worksheets with items from a List<string> starting at cell A1, then checks that every worksheet contains exactly the same number of rows as the list.
// Common Searches: aspnet aspose.cells verify each worksheet row count matches list size | c# how to check row count of all sheets in an Aspose.Cells workbook | validate detail worksheets row numbers against source collection aspose.cells | aspose.cells RowCollection.Count example for row validation | c# populate Excel worksheets from List<string> and confirm row count
// Tags: Aspose.Cells worksheet row count validation | C# populate worksheet from List<string> | Aspose.Cells RowCollection.Count usage | validate detail sheets in Excel workbook | Aspose.Cells verify row count per sheet

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsValidationDemo
{
    // Shows how to create a workbook, populate multiple worksheets from a List<string>, and validate that each worksheet's row count matches the source list count using Aspose.Cells RowCollection.Count.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Example source collection whose count we expect in each detail worksheet
            List<string> sourceData = new List<string>
            {
                "Item 1",
                "Item 2",
                "Item 3",
                "Item 4"
            };
            int expectedRowCount = sourceData.Count;

            // Populate two detail worksheets with data from the source collection
            Worksheet detailSheet1 = workbook.Worksheets[workbook.Worksheets.Add()];
            detailSheet1.Name = "DetailSheet1";
            PopulateWorksheet(detailSheet1, sourceData);

            Worksheet detailSheet2 = workbook.Worksheets[workbook.Worksheets.Add()];
            detailSheet2.Name = "DetailSheet2";
            PopulateWorksheet(detailSheet2, sourceData);

            // Validate that each detail worksheet has the expected number of rows
            ValidateDetailWorksheets(workbook, expectedRowCount);

            // Save the workbook (lifecycle rule: save)
            workbook.Save("ValidatedWorkbook.xlsx");
        }

        // Helper method to fill a worksheet with the source data starting at row 0, column 0
        private static void PopulateWorksheet(Worksheet sheet, List<string> data)
        {
            for (int i = 0; i < data.Count; i++)
            {
                // Put each item into column A of the worksheet
                sheet.Cells[i, 0].PutValue(data[i]);
            }
        }

        // Validation method that checks each worksheet's row count against the expected count
        private static void ValidateDetailWorksheets(Workbook workbook, int expectedRowCount)
        {
            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Retrieve the RowCollection for the current worksheet
                RowCollection rows = sheet.Cells.Rows; // uses RowCollection property

                // Get the actual number of rows (property Count)
                int actualRowCount = rows.Count; // Count property of RowCollection

                // Output validation result
                if (actualRowCount == expectedRowCount)
                {
                    Console.WriteLine($"Worksheet \"{sheet.Name}\" validation passed. Row count: {actualRowCount}");
                }
                else
                {
                    Console.WriteLine($"Worksheet \"{sheet.Name}\" validation FAILED. Expected rows: {expectedRowCount}, Actual rows: {actualRowCount}");
                }
            }
        }
    }
}
