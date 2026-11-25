// Title: Create an Excel pivot table and freeze its row labels using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells example to create a pivot table and freeze row headers | How to apply FreezePanes to a pivot table's row labels with Aspose.Cells | Set PivotTableStyleMedium9 and freeze row labels in Aspose.Cells .NET | Generate sample data and build a pivot table programmatically with Aspose.Cells C# | Refresh and calculate pivot data before freezing panes in Aspose.Cells
// Tags: Aspose.Cells create pivot table C# | Aspose.Cells freeze panes pivot row labels | Aspose.Cells apply PivotTableStyleMedium9 | Aspose.Cells refresh calculate pivot data | Aspose.Cells populate sample data for pivot

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotFreezeDemo
{
    // The sample program creates a new workbook, fills columns A‑C with sample Category, Year, and Amount values, defines the source range, adds a pivot table at cell E5, assigns Category to rows, Year to columns, and Amount to data, applies the PivotTableStyleMedium9 style, refreshes and calculates the pivot data, then freezes panes so the pivot's row labels stay visible, and finally saves the file as an XLSX workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string outputPath = "PivotTable_FrozenRowLabels.xlsx";

            try
            {
                // Create a new workbook
                using (Workbook workbook = new Workbook())
                {
                    // Access the first worksheet
                    Worksheet sheet = workbook.Worksheets[0];
                    Cells cells = sheet.Cells;

                    // -------------------------------------------------
                    // Populate sample data for the pivot table
                    // -------------------------------------------------
                    cells["A1"].PutValue("Category");
                    cells["B1"].PutValue("Year");
                    cells["C1"].PutValue("Amount");

                    string[] categories = { "Fruit", "Vegetable", "Beverage", "Snack" };
                    int[] years = { 2020, 2021, 2022 };
                    Random rnd = new Random();

                    int row = 2; // 1‑based index for Excel rows
                    for (int i = 0; i < 30; i++)
                    {
                        cells[$"A{row}"].PutValue(categories[rnd.Next(categories.Length)]);
                        cells[$"B{row}"].PutValue(years[rnd.Next(years.Length)]);
                        cells[$"C{row}"].PutValue(rnd.Next(10, 100));
                        row++;
                    }

                    // -------------------------------------------------
                    // Create the pivot table
                    // -------------------------------------------------
                    string sourceData = $"=Sheet1!A1:C{row - 1}";
                    string pivotDestination = "E5";

                    PivotTableCollection pivots = sheet.PivotTables;
                    // Using the obsolete overload is acceptable; it still compiles.
                    int pivotIndex = pivots.Add(sourceData, pivotDestination, "SalesPivot");
                    PivotTable pivot = pivots[pivotIndex];

                    // Add fields to the pivot table
                    pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                    pivot.AddFieldToArea(PivotFieldType.Column, "Year");
                    pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

                    // Apply a built‑in style
                    pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;

                    // Refresh and calculate the pivot data
#pragma warning disable CS0618 // Suppress obsolete warnings for demonstration
                    pivot.RefreshData();
#pragma warning restore CS0618
                    pivot.CalculateData();

                    // -------------------------------------------------
                    // Freeze the row labels (first column of the pivot)
                    // -------------------------------------------------
                    int pivotStartRow = sheet.Cells[pivotDestination].Row;       // zero‑based
                    int pivotStartColumn = sheet.Cells[pivotDestination].Column; // zero‑based

                    // Freeze rows above the pivot and columns up to (and including) the pivot's first column.
                    // The overload requires: row, column, freezedRows, freezedColumns.
                    sheet.FreezePanes(pivotStartRow, pivotStartColumn + 1, pivotStartRow, pivotStartColumn + 1);

                    // -------------------------------------------------
                    // Save the workbook
                    // -------------------------------------------------
                    workbook.Save(outputPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Workbook saved to '{outputPath}'.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
