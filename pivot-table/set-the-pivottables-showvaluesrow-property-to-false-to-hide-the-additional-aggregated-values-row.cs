// Title: Aspose.Cells C# – Hide PivotTable Values Row with ShowValuesRow = false
// Description: Demonstrates how to create a workbook, add sample fruit data, build a PivotTable on range A1:B4, assign row and data fields, and suppress the extra aggregated values row by setting PivotTable.ShowValuesRow to false before saving the file as XLSX.
// Keywords: Aspose.Cells | C# | PivotTable | ShowValuesRow | hide values row | remove aggregated row | Excel pivot table example | Aspose.Cells display options | Excel automation C# | pivot table layout
// Common Searches: Aspose.Cells hide values row pivot table C# | ShowValuesRow false example Aspose.Cells | remove extra row from PivotTable using Aspose.Cells | C# code to suppress values row in Excel pivot table | Aspose.Cells PivotTable display settings
// Developer Intent: Disable the aggregated values row in an Aspose.Cells PivotTable.
// Use Cases: Generate a sales report where the total values row is omitted for a cleaner layout. | Create a dashboard workbook that shows only category labels and data without the extra values row. | Export inventory summaries to Excel while matching a template that excludes the pivot table values row.
// AI Prompts: Write C# code with Aspose.Cells that builds a PivotTable and hides its values row by setting ShowValuesRow to false. | Explain the impact of the ShowValuesRow property on PivotTable appearance and when it should be used. | Provide a complete Aspose.Cells example that creates a PivotTable, disables the values row, and optionally configures ShowGrandTotals.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add sample fruit data, build a PivotTable on range A1:B4, assign row and data fields, and suppress the extra aggregated values row by setting PivotTable.ShowValuesRow to false before saving the file as XLSX.
    public class PivotTableShowValuesRowDemo
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

            // Populate sample data for the pivot table
            Cells cells = sheet.Cells;
            cells["A1"].Value = "Fruit";
            cells["B1"].Value = "Quantity";
            cells["A2"].Value = "Apple";
            cells["B2"].Value = 10;
            cells["A3"].Value = "Orange";
            cells["B3"].Value = 15;
            cells["A4"].Value = "Banana";
            cells["B4"].Value = 20;

            // Add a pivot table based on the data range A1:B4, place it at D1
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table: fruit as row field, quantity as data field
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Column 0 (Fruit)
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Column 1 (Quantity)

            // Hide the additional aggregated values row
            pivotTable.ShowValuesRow = false;

            // Save the workbook to a file
            string outputPath = "PivotTableShowValuesRowDemo.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
