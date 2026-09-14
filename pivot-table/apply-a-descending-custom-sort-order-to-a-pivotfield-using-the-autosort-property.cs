// Title: Configure a PivotField to sort rows in descending order with Aspose.Cells for .NET
// AI Prompts: Create a C# program that builds a pivot table and sets the row PivotField to auto‑sort descending by the first data field using Aspose.Cells. | Write C# code that enables IsAutoSort, sets IsAscendSort to false, and assigns AutoSortField for a PivotField in an Aspose.Cells workbook.
// Common Searches: asp.net aspose.cells how to set descending auto sort on a pivot field | c# pivot table row field descending sort using Aspose.Cells | example of IsAutoSort and IsAscendSort properties in Aspose.Cells | sort pivot table rows by amount descending Aspose.Cells C#
// Tags: pivot field sort direction Aspose.Cells | IsAutoSort property usage C# | IsAscendSort false configuration Aspose.Cells | AutoSortField set to data field Aspose.Cells | custom pivot table sort order .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // The example creates a workbook, adds sample data, builds a pivot table, places 'Category' as a row field and 'Amount' as a data field, then enables auto‑sorting on the row field, configures it to sort descending, selects the first data field as the sort key, calculates the pivot data, and saves the workbook.
    public class PivotFieldDescendingCustomAutoSortDemo
    {
        public static void Main()
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
            sheet.Cells["A1"].Value = "Category";
            sheet.Cells["B1"].Value = "Amount";
            sheet.Cells["A2"].Value = "B";
            sheet.Cells["A3"].Value = "A";
            sheet.Cells["A4"].Value = "C";
            sheet.Cells["B2"].Value = 200;
            sheet.Cells["B3"].Value = 500;
            sheet.Cells["B4"].Value = 300;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add the row field (Category) and the data field (Amount)
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Configure descending auto‑sort for the row field
            PivotField rowField = pivotTable.RowFields[0];
            rowField.IsAutoSort = true;      // Enable auto sorting
            rowField.IsAscendSort = false;   // Set descending order
            rowField.AutoSortField = 0;      // Sort by the first data field (Amount)

            // Calculate the pivot table data
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotFieldDescendingCustomAutoSortDemo_out.xlsx");
        }
    }
}
