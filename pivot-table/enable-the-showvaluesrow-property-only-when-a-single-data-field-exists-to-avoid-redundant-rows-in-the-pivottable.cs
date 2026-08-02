// Title: Aspose.Cells C# Example: Conditionally Enable ShowValuesRow for PivotTables with a Single Data Field
// Description: This sample creates a workbook, populates it with sales data, adds a PivotTable, and sets the ShowValuesRow property only when the PivotTable contains exactly one data field, preventing an unnecessary values row. The workbook is then saved as an XLSX file.
// Keywords: Aspose.Cells | C# | PivotTable | ShowValuesRow | conditional ShowValuesRow | single data field | DataFields.Count | Excel pivot example | avoid extra values row | GitHub sample | Aspose.Cells tutorial
// Common Searches: Aspose.Cells enable ShowValuesRow for single data field | C# pivot table hide values row when multiple measures | conditional ShowValuesRow property Aspose.Cells | check DataFields count before setting ShowValuesRow | avoid redundant values row in Aspose.Cells PivotTable
// Developer Intent: Toggle the ShowValuesRow flag of a PivotTable based on whether it has exactly one data field, ensuring a clean layout.
// Use Cases: Generate a sales summary where the values row appears only for a single metric. | Build a dynamic reporting tool that adds or removes measures at runtime and automatically adjusts ShowValuesRow. | Export Excel workbooks with pivots that hide the values row when multiple data fields are present, eliminating duplicate total rows.
// AI Prompts: Write C# code using Aspose.Cells to add a PivotTable and enable ShowValuesRow only when there is exactly one data field. | Explain how to inspect the DataFields collection of an Aspose.Cells PivotTable and set ShowValuesRow conditionally. | Provide a complete Aspose.Cells example that disables ShowValuesRow for pivots with multiple data fields and saves the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // This sample creates a workbook, populates it with sales data, adds a PivotTable, and sets the ShowValuesRow property only when the PivotTable contains exactly one data field, preventing an unnecessary values row. The workbook is then saved as an XLSX file.
    public class PivotTableShowValuesRowConditionalDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].Value = "Category";
            cells["B1"].Value = "Product";
            cells["C1"].Value = "Sales";

            cells["A2"].Value = "Fruit";
            cells["B2"].Value = "Apple";
            cells["C2"].Value = 120;

            cells["A3"].Value = "Fruit";
            cells["B3"].Value = "Orange";
            cells["C3"].Value = 150;

            cells["A4"].Value = "Vegetable";
            cells["B4"].Value = "Carrot";
            cells["C4"].Value = 80;

            cells["A5"].Value = "Vegetable";
            cells["B5"].Value = "Potato";
            cells["C5"].Value = 90;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table: Category as row, Product as column, Sales as data
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Refresh data and calculate the pivot table before checking fields
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Enable ShowValuesRow only when there is a single data field
            if (pivotTable.DataFields.Count == 1)
            {
                pivotTable.ShowValuesRow = true;
                Console.WriteLine("ShowValuesRow enabled (single data field).");
            }
            else
            {
                pivotTable.ShowValuesRow = false;
                Console.WriteLine("ShowValuesRow disabled (multiple data fields).");
            }

            // Save the workbook
            workbook.Save("PivotTableShowValuesRowConditionalDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}
