// Title: Insert a Column and Build a Quarterly Revenue Pivot Table with Aspose.Cells for .NET (C#)
// Description: The sample creates a workbook, populates FiscalYear, Quarter and random Revenue values, inserts a new column before the Revenue field, shifts the data, then generates a PivotTable that sums quarterly revenue per fiscal year with a currency format and saves the result as QuarterlyRevenuePivot.xlsx.
// Keywords: Aspose.Cells insert column C# | Aspose.Cells PivotTable example | quarterly revenue pivot Aspose | C# sum revenue pivot field | Aspose.Cells move column data | .NET financial reporting pivot | Excel pivot table automation
// Common Searches: How to insert a column before a data field using Aspose.Cells C# | Create a PivotTable that shows revenue by fiscal year and quarter in Aspose.Cells | Set sum aggregation and currency format for a PivotTable data field in .NET | Refresh and calculate a PivotTable after modifying its source range with Aspose.Cells
// Developer Intent: Generate a PivotTable that displays total quarterly revenue for each fiscal year after adding an extra column to the source worksheet.
// Use Cases: Add a placeholder column before the revenue column to accommodate future data without breaking existing calculations. | Produce a financial summary that aggregates revenue by fiscal year (rows) and quarter (columns). | Apply a $#,##0 number format to the revenue data field for clear monetary presentation.
// AI Prompts: Write C# code with Aspose.Cells to insert a column at index 2, shift existing columns, and move Revenue values to the new column. | Show how to configure a PivotTable with FiscalYear as rows, Quarter as columns, and Revenue as a summed data field using a currency format. | Explain the steps to refresh and recalculate a PivotTable after changing its source range in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // The sample creates a workbook, populates FiscalYear, Quarter and random Revenue values, inserts a new column before the Revenue field, shifts the data, then generates a PivotTable that sums quarterly revenue per fiscal year with a currency format and saves the result as QuarterlyRevenuePivot.xlsx.
    public class InsertColumnFieldsForQuarterlyRevenue
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // ------------------------------------------------------------
                // 1. Prepare source data: FiscalYear, Quarter, Revenue
                // ------------------------------------------------------------
                cells["A1"].PutValue("FiscalYear");
                cells["B1"].PutValue("Quarter");
                cells["C1"].PutValue("Revenue");

                // Sample data for two fiscal years
                string[] years = { "2022", "2023" };
                string[] quarters = { "Q1", "Q2", "Q3", "Q4" };
                Random rnd = new Random();

                int row = 1; // zero‑based index, row 1 is the second row (A2)
                foreach (string year in years)
                {
                    foreach (string quarter in quarters)
                    {
                        cells[row, 0].PutValue(year);          // FiscalYear
                        cells[row, 1].PutValue(quarter);       // Quarter
                        cells[row, 2].PutValue(rnd.Next(5000, 20000)); // Revenue
                        row++;
                    }
                }

                // ------------------------------------------------------------
                // 2. Insert an extra column before the Revenue column (optional)
                //    Demonstrates the InsertColumns method.
                // ------------------------------------------------------------
                // Insert one column at index 2 (C column). Existing columns D and beyond shift right.
                sheet.Cells.InsertColumns(2, 1);
                // After insertion, move the Revenue header/value to the new column (now D)
                sheet.Cells["D1"].PutValue("Revenue");
                for (int i = 2; i <= row; i++)
                {
                    sheet.Cells[i, 3].PutValue(sheet.Cells[i, 2].Value);
                    sheet.Cells[i, 2].PutValue(null); // clear old location
                }

                // ------------------------------------------------------------
                // 3. Create a PivotTable to show quarterly revenue per fiscal year
                // ------------------------------------------------------------
                // Define the source range (including the inserted column)
                string sourceRange = $"A1:D{row}";
                // Place the pivot table starting at cell F3
                int pivotIndex = sheet.PivotTables.Add(sourceRange, "F3", "QuarterlyRevenuePivot");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Add FiscalYear as Row field
                pivot.AddFieldToArea(PivotFieldType.Row, "FiscalYear");
                // Add Quarter as Column field
                pivot.AddFieldToArea(PivotFieldType.Column, "Quarter");
                // Add Revenue as Data field and set aggregation to Sum
                int dataFieldPos = pivot.AddFieldToArea(PivotFieldType.Data, "Revenue");
                pivot.DataFields[dataFieldPos].Function = ConsolidationFunction.Sum;
                pivot.DataFields[dataFieldPos].NumberFormat = "$#,##0";

                // Refresh and calculate the pivot table
                pivot.RefreshData();
                pivot.CalculateData();

                // ------------------------------------------------------------
                // 4. Save the workbook
                // ------------------------------------------------------------
                workbook.Save("QuarterlyRevenuePivot.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
