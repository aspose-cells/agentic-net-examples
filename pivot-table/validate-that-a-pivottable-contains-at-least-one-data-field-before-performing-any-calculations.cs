// Title: Validate PivotTable Data Fields Before Calculation with Aspose.Cells for .NET (C#)
// Description: C# example that creates a workbook, populates sample data, adds a PivotTable on range A1:B5, includes only a row field, checks PivotTable.DataFields.Count, throws an InvalidOperationException when no data fields exist, then refreshes, calculates, and saves the file.
// Keywords: Aspose.Cells | C# | .NET | PivotTable validation | DataFields count | CalculateData | RefreshData | InvalidOperationException | Excel pivot table | missing data field
// Common Searches: Aspose.Cells check PivotTable data fields before CalculateData | C# validate PivotTable has data field Aspose | Throw exception when PivotTable has no data fields Aspose.Cells | How to ensure PivotTable data field exists in .NET | PivotTable.DataFields count example Aspose
// Developer Intent: Verify that a PivotTable contains at least one data field before calling RefreshData or CalculateData to prevent runtime errors.
// Use Cases: Detect missing data fields in automated report generation and raise a clear error. | Guard Excel export pipelines by validating PivotTable configuration prior to calculation. | Provide user‑friendly feedback when a pivot table is defined without required data fields.
// AI Prompts: Generate C# code that adds a default data field to an Aspose.Cells PivotTable when DataFields.Count is zero, then refreshes and calculates the pivot. | Write a try‑catch block that validates a PivotTable's data fields and logs a custom warning instead of throwing an exception. | Show how to customize the InvalidOperationException message for a missing data field in an Aspose.Cells pivot table scenario.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // C# example that creates a workbook, populates sample data, adds a PivotTable on range A1:B5, includes only a row field, checks PivotTable.DataFields.Count, throws an InvalidOperationException when no data fields exist, then refreshes, calculates, and saves the file.
    public class PivotTableValidationDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Amount");
                sheet.Cells["A2"].PutValue("Food");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["A3"].PutValue("Drink");
                sheet.Cells["B3"].PutValue(80);
                sheet.Cells["A4"].PutValue("Food");
                sheet.Cells["B4"].PutValue(150);
                sheet.Cells["A5"].PutValue("Drink");
                sheet.Cells["B5"].PutValue(70);

                // Add a pivot table based on the data range A1:B5, placed at D3
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "SalesPivot");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add row field
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                // Data field intentionally omitted to demonstrate validation
                // pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Validate that at least one data field exists before calculations
                if (pivotTable.DataFields.Count == 0)
                {
                    throw new InvalidOperationException(
                        "PivotTable must contain at least one data field before calculating data.");
                }

                // Refresh and calculate the pivot data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook
                workbook.Save("ValidatedPivotTable.xlsx");
                Console.WriteLine("Workbook saved as 'ValidatedPivotTable.xlsx'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
