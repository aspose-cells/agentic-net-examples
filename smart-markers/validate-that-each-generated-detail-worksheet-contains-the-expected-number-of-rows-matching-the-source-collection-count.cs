using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsRowValidation
{
    class Program
    {
        static void Main()
        {
            // Sample source collection whose count we expect to match rows in each detail worksheet
            List<string> sourceData = new List<string>
            {
                "Alpha",
                "Beta",
                "Gamma",
                "Delta"
            };

            // Create a new workbook (uses the create rule)
            Workbook workbook = new Workbook();

            // Iterate over the source collection and create a detail worksheet for each item
            for (int i = 0; i < sourceData.Count; i++)
            {
                // Create a new worksheet named after the source item
                Worksheet detailSheet = workbook.Worksheets[workbook.Worksheets.Add()];
                detailSheet.Name = $"Detail_{sourceData[i]}";

                // Populate the worksheet with some rows (for demonstration we add two rows per item)
                // Row 0 – header
                detailSheet.Cells[0, 0].PutValue("Index");
                detailSheet.Cells[0, 1].PutValue("Value");

                // Row 1 – data row
                detailSheet.Cells[1, 0].PutValue(i + 1);
                detailSheet.Cells[1, 1].PutValue(sourceData[i]);

                // Validate that the number of rows in the worksheet matches the source collection count
                // RowCollection.Count gives the total rows that have been instantiated (including empty rows up to the last used row)
                int actualRowCount = detailSheet.Cells.Rows.Count;

                // Expected rows: header + data row = 2
                // If you need to match the source collection count, adjust the expectation accordingly.
                int expectedRowCount = 2; // change as needed

                if (actualRowCount != expectedRowCount)
                {
                    Console.WriteLine($"Validation failed for worksheet '{detailSheet.Name}'. Expected rows: {expectedRowCount}, Actual rows: {actualRowCount}");
                }
                else
                {
                    Console.WriteLine($"Worksheet '{detailSheet.Name}' passed validation. Row count: {actualRowCount}");
                }
            }

            // Save the workbook (uses the save rule)
            workbook.Save("DetailWorksheets.xlsx");
        }
    }
}