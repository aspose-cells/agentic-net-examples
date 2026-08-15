// Title: Hide the Values Row in an Aspose.Cells PivotTable (C#) – Set ShowValuesRow = false
// Description: Demonstrates how to create a workbook, add sample data, build a PivotTable on range A1:B4, assign "Fruit" to rows and "Quantity" to data, and hide the extra aggregated values row by setting pivotTable.ShowValuesRow to false before saving the file.
// Keywords: Aspose.Cells PivotTable ShowValuesRow | C# hide values row Aspose.Cells | remove aggregated row PivotTable Aspose | Aspose.Cells pivot table formatting | Excel PivotTable hide total row C#
// Common Searches: Aspose.Cells hide values row C# example | ShowValuesRow false PivotTable Aspose | how to remove total row from Aspose.Cells pivot table | C# code to hide aggregated row in Excel pivot using Aspose | Aspose.Cells PivotTable formatting options
// Developer Intent: Set the PivotTable's ShowValuesRow property to false to suppress the aggregated values row.
// Use Cases: Generate a sales report where the pivot table should not display the values row for a cleaner layout. | Apply client‑specified formatting by programmatically removing the total row from a PivotTable before exporting. | Update an existing workbook's pivot table to hide the aggregated row as part of an automated Excel generation pipeline.
// AI Prompts: Write C# code with Aspose.Cells that creates a PivotTable from range A1:B10 and hides the values row using ShowValuesRow = false. | Explain the impact of the ShowValuesRow property on an Aspose.Cells PivotTable and show how to toggle it on and off. | Provide a step‑by‑step guide to hide the aggregated values row in an existing workbook's PivotTable using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add sample data, build a PivotTable on range A1:B4, assign "Fruit" to rows and "Quantity" to data, and hide the extra aggregated values row by setting pivotTable.ShowValuesRow to false before saving the file.
    public class PivotTableHideValuesRowDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Pivot table created and saved successfully.");
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

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "C3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table (Fruit as row, Quantity as data)
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);

            // Hide the additional aggregated values row
            pivotTable.ShowValuesRow = false;

            // Save the workbook to a file
            workbook.Save("PivotTableHideValuesRowDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}
