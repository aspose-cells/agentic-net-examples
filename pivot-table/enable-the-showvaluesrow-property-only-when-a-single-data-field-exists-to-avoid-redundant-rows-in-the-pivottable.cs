// Title: Enable ShowValuesRow in an Aspose.Cells PivotTable only when a single data field is present (C#)
// AI Prompts: Write C# code using Aspose.Cells that creates a pivot table and sets PivotTable.ShowValuesRow to true only if the pivot contains exactly one data field. | Adapt an existing Aspose.Cells pivot table example to dynamically toggle the ShowValuesRow property based on the DataFields count at runtime.
// Common Searches: Aspose.Cells how to turn on ShowValuesRow for a pivot table with one data field | C# set ShowValuesRow property conditionally in Aspose.Cells pivot | Enable values row in Excel pivot only when there is a single data field using Aspose.Cells | Aspose.Cells pivot table hide values row when multiple data fields are added | Check DataFields count before enabling ShowValuesRow in Aspose.Cells .NET
// Tags: Aspose.Cells pivot ShowValuesRow conditional | C# pivot table data field count | Aspose.Cells enable values row dynamically | Excel pivot hide values row multiple fields | Aspose.Cells .NET pivot table settings

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // The example creates a workbook, fills it with sample data, adds a pivot table with Category rows, Product columns, and Sales as the data field (optionally a second field), refreshes and calculates the pivot, then enables ShowValuesRow only when exactly one data field exists, and finally saves the workbook as an .xlsx file.
    public class PivotTableShowValuesRowConditionalDemo
    {
        public static void Main(string[] args)
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

            // Populate sample data
            cells["A1"].Value = "Category";
            cells["B1"].Value = "Product";
            cells["C1"].Value = "Sales";
            cells["D1"].Value = "Quantity";

            cells["A2"].Value = "Fruit";
            cells["B2"].Value = "Apple";
            cells["C2"].Value = 1200;
            cells["D2"].Value = 10;

            cells["A3"].Value = "Fruit";
            cells["B3"].Value = "Orange";
            cells["C3"].Value = 1500;
            cells["D3"].Value = 15;

            cells["A4"].Value = "Vegetable";
            cells["B4"].Value = "Carrot";
            cells["C4"].Value = 800;
            cells["D4"].Value = 8;

            // Add a pivot table covering the data range
            int pivotIndex = sheet.PivotTables.Add("A1:D4", "F3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields: Category as row, Product as column, Sales as data
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // OPTIONAL: Add a second data field to demonstrate the conditional logic
            // Comment out the line below if you want only a single data field.
            // pivotTable.AddFieldToArea(PivotFieldType.Data, "Quantity");

            try
            {
                // Refresh the pivot cache and calculate data
                pivotTable.RefreshData();
                pivotTable.CalculateData();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Pivot refresh error: {ex.Message}");
                return;
            }

            // Enable ShowValuesRow only when there is exactly one data field
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
            string outputPath = "PivotTableShowValuesRowConditionalDemo.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Save error: {ex.Message}");
            }
        }
    }
}
