// Title: Validate that detail worksheets created with Aspose.Cells smart markers have the expected row count in C#
// Common Searches: C# Aspose.Cells validate row count of a generated detail worksheet | how to compare MaxDataRow to expected rows using Aspose.Cells | verify smart markers detail sheet row numbers match source collection in .NET | Aspose.Cells workbook row count validation with dictionary of expected values | throw exception when worksheet row count does not match expected in C# Excel report
// Tags: Aspose.Cells worksheet row count check | C# MaxDataRow usage for row verification | smart markers generated sheet validation | expected rows dictionary validation in Excel | exception handling for row mismatch Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The DetailWorksheetValidator class loads an Excel workbook with Aspose.Cells, iterates each worksheet, and for those listed in a supplied dictionary compares the actual data row count (Cells.MaxDataRow + 1) to the expected count. It throws an InvalidOperationException when a mismatch is found and writes a success message when all detail worksheets contain the expected number of rows.
    public class DetailWorksheetValidator
    {
        /// <param name="workbookPath">Full path to the Excel file to validate.</param>
        /// <param name="expectedRowCounts">
        /// A dictionary where the key is the worksheet name and the value is the expected number of data rows.
        /// </param>
        public void Validate(string workbookPath, Dictionary<string, int> expectedRowCounts)
        {
            if (!File.Exists(workbookPath))
                throw new FileNotFoundException($"The workbook file was not found: {workbookPath}");

            Workbook workbook;
            try
            {
                // Load the workbook using Aspose.Cells
                workbook = new Workbook(workbookPath);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to load workbook '{workbookPath}'.", ex);
            }

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Validate only worksheets that are present in the expected dictionary
                if (!expectedRowCounts.ContainsKey(sheet.Name))
                    continue;

                // MaxDataRow is zero‑based; add 1 to get the count of rows that contain data
                int actualRowCount = sheet.Cells.MaxDataRow + 1;
                int expectedCount = expectedRowCounts[sheet.Name];

                if (actualRowCount != expectedCount)
                {
                    throw new InvalidOperationException(
                        $"Worksheet \"{sheet.Name}\" contains {actualRowCount} rows, but the expected count is {expectedCount}.");
                }
            }

            Console.WriteLine("All detail worksheets contain the expected number of rows.");
        }
    }

    internal class Program
    {
        private static void Main(string[] args)
        {
            try
            {
                // Example usage
                var validator = new DetailWorksheetValidator();
                var expected = new Dictionary<string, int>
                {
                    { "DetailSales", 150 },
                    { "DetailInventory", 78 }
                };

                // Adjust the path to point to an existing workbook on your machine
                string workbookPath = @"C:\Reports\GeneratedReport.xlsx";

                validator.Validate(workbookPath, expected);
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.Error.WriteLine($"File error: {fnfEx.Message}");
            }
            catch (InvalidOperationException invEx)
            {
                Console.Error.WriteLine($"Validation error: {invEx.Message}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
