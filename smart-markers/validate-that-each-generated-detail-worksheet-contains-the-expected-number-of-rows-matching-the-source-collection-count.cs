using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsValidationDemo
{
    class Program
    {
        static void Main()
        {
            // ----- Create source collection -----
            List<string> sourceData = new List<string>
            {
                "Alpha",
                "Beta",
                "Gamma",
                "Delta"
            };

            // ----- Create a new workbook -----
            Workbook workbook = new Workbook();

            // ----- Generate detail worksheets based on source collection -----
            // Each worksheet will contain one row per source item.
            for (int i = 0; i < sourceData.Count; i++)
            {
                // Add a new worksheet for the current item
                Worksheet detailSheet = workbook.Worksheets[workbook.Worksheets.Add()];
                detailSheet.Name = $"Detail_{i + 1}";

                // Populate rows: one row per source item
                for (int rowIndex = 0; rowIndex < sourceData.Count; rowIndex++)
                {
                    // Put the source value into column A of the current row
                    detailSheet.Cells[rowIndex, 0].PutValue(sourceData[rowIndex]);
                }
            }

            // ----- Validation: ensure each detail worksheet has the expected row count -----
            int expectedRowCount = sourceData.Count;
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Skip the default first sheet if it was not used for details
                if (!sheet.Name.StartsWith("Detail_"))
                    continue;

                // Get the actual number of rows that have been instantiated
                int actualRowCount = sheet.Cells.Rows.Count;

                // Compare with expected count and output the result
                if (actualRowCount == expectedRowCount)
                {
                    Console.WriteLine($"Worksheet '{sheet.Name}' validation passed. Row count = {actualRowCount}.");
                }
                else
                {
                    Console.WriteLine($"Worksheet '{sheet.Name}' validation FAILED. Expected {expectedRowCount} rows but found {actualRowCount} rows.");
                }
            }

            // ----- Save the workbook -----
            workbook.Save("DetailWorksheetsValidation.xlsx");
        }
    }
}