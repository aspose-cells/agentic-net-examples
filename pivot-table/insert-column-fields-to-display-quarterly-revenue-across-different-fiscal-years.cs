// Title: Insert a Notes column and build a quarterly revenue pivot table by fiscal year with Aspose.Cells for .NET
// Description: This example shows how to create a workbook, add FiscalYear, Quarter, and Revenue data, insert a blank "Notes" column using the InsertColumns API, set a header and placeholder values, define a source range that excludes the new column, and generate a pivot table that displays summed quarterly revenue per fiscal year with currency formatting.
// Keywords: Aspose.Cells insert column C# | Aspose.Cells InsertColumns API | Aspose.Cells pivot table example | quarterly revenue pivot Aspose.Cells | FiscalYear Quarter Revenue pivot | C# pivot table currency format | .NET spreadsheet pivot table | exclude column from pivot source Aspose.Cells | Aspose.Cells tutorial
// Common Searches: How to insert a column before creating a pivot table with Aspose.Cells for .NET | Create a pivot table that shows quarterly revenue per fiscal year using Aspose.Cells | Exclude an inserted column from the pivot source range in Aspose.Cells | Apply currency number format to a pivot table data field in C# | Aspose.Cells InsertColumns example
// Developer Intent: Add an extra column and generate a pivot table that summarizes quarterly revenue by fiscal year.
// Use Cases: Insert a blank "Notes" column after the Revenue column, add a header, and fill rows with a default value. | Define a source range that includes only FiscalYear, Quarter, and Revenue columns, then configure the pivot with FiscalYear as rows, Quarter as columns, and Revenue summed as the data field. | Apply a currency number format to the Revenue data field, refresh the pivot, and save the workbook.
// AI Prompts: Write C# code using Aspose.Cells to insert a column at a specific index, add a header, and populate it with placeholder values. | Generate C# code that creates a pivot table from a range, sets FiscalYear as row fields, Quarter as column fields, sums Revenue, and formats the data field as currency. | Explain how to exclude an inserted column from the pivot table source range so the pivot reflects only the intended data columns.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // This example shows how to create a workbook, add FiscalYear, Quarter, and Revenue data, insert a blank "Notes" column using the InsertColumns API, set a header and placeholder values, define a source range that excludes the new column, and generate a pivot table that displays summed quarterly revenue per fiscal year with currency formatting.
    public class QuarterlyRevenuePivotDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // 1. Prepare source data: FiscalYear, Quarter, Revenue
                // -------------------------------------------------
                sheet.Cells["A1"].PutValue("FiscalYear");
                sheet.Cells["B1"].PutValue("Quarter");
                sheet.Cells["C1"].PutValue("Revenue");

                // Sample data for two fiscal years
                string[] years = { "2022", "2023" };
                string[] quarters = { "Q1", "Q2", "Q3", "Q4" };
                double[,] revenues = {
                    { 120000, 150000, 130000, 160000 }, // 2022
                    { 140000, 170000, 150000, 180000 }  // 2023
                };

                int row = 1;
                for (int i = 0; i < years.Length; i++)
                {
                    for (int j = 0; j < quarters.Length; j++)
                    {
                        sheet.Cells[row, 0].PutValue(years[i]);          // FiscalYear
                        sheet.Cells[row, 1].PutValue(quarters[j]);      // Quarter
                        sheet.Cells[row, 2].PutValue(revenues[i, j]);   // Revenue
                        row++;
                    }
                }

                // -------------------------------------------------
                // 2. Insert an extra column (optional) to demonstrate InsertColumns API
                // -------------------------------------------------
                // Insert a blank column at index 3 (after the Revenue column)
                sheet.Cells.InsertColumns(3, 1, true);

                // Add a header for the new column
                sheet.Cells["D1"].PutValue("Notes");
                // Populate some placeholder notes
                for (int r = 2; r <= row; r++)
                {
                    sheet.Cells[r - 1, 3].PutValue("N/A");
                }

                // -------------------------------------------------
                // 3. Create a PivotTable to show quarterly revenue per fiscal year
                // -------------------------------------------------
                // Define the source range (excluding the Notes column)
                string sourceRange = "A1:C" + (row - 1).ToString();

                // Place the pivot table starting at cell E3
                int pivotIndex = sheet.PivotTables.Add(sourceRange, "E3", "QuarterlyRevenuePivot");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Add FiscalYear as Row field
                pivot.AddFieldToArea(PivotFieldType.Row, "FiscalYear");
                // Add Quarter as Column field
                pivot.AddFieldToArea(PivotFieldType.Column, "Quarter");
                // Add Revenue as Data field (Sum)
                int dataFieldPos = pivot.AddFieldToArea(PivotFieldType.Data, "Revenue");
                pivot.DataFields[dataFieldPos].Function = ConsolidationFunction.Sum;
                pivot.DataFields[dataFieldPos].NumberFormat = "$#,##0";

                // Refresh and calculate the pivot table using the correct API
                pivot.RefreshData();
                pivot.CalculateData();

                // -------------------------------------------------
                // 4. Save the workbook
                // -------------------------------------------------
                workbook.Save("QuarterlyRevenuePivot.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
