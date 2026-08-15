// Title: Aspose.Cells C# – Enable ShowValuesRow Only When a PivotTable Has One Data Field
// Description: Demonstrates how to create a workbook, add a pivot table, and programmatically turn on the ShowValuesRow flag only if the pivot contains a single data field, preventing an unnecessary values row in the generated Excel file.
// Keywords: Aspose.Cells ShowValuesRow | C# pivot table conditional formatting | single data field pivot | Hide extra values row Aspose | PivotTable.DataFields count | Excel report automation .NET
// Common Searches: set ShowValuesRow for one data field Aspose.Cells | C# pivot table hide values row when multiple fields | Aspose.Cells conditional ShowValuesRow example | how to enable ShowValuesRow based on DataFields count | avoid redundant values row in Excel pivot using Aspose
// Developer Intent: Activate the ShowValuesRow property of a PivotTable only when the table contains exactly one data field.
// Use Cases: Generate compact Excel reports where the values row appears only for a sole metric. | Automatically suppress the values row when additional measures (e.g., Quantity, Profit) are added to the pivot. | Create dynamic dashboards that adjust layout based on the number of data fields present.
// AI Prompts: Write C# code with Aspose.Cells that adds a pivot table and sets ShowValuesRow to true only when DataFields.Count equals 1. | Explain how to check the count of data fields in a PivotTable and toggle ShowValuesRow accordingly in .NET. | Provide a step‑by‑step guide to prevent duplicate values rows in an Aspose.Cells pivot table based on field count.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add a pivot table, and programmatically turn on the ShowValuesRow flag only if the pivot contains a single data field, preventing an unnecessary values row in the generated Excel file.
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
            try
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

                // Configure the pivot table: Category as row, Sales as data
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Refresh data and calculate the pivot table
                pivotTable.RefreshData();      // Correct API to refresh the cache
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
                string outputPath = "PivotTableShowValuesRowConditionalDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Runtime error: {ex.Message}");
            }
        }
    }
}
