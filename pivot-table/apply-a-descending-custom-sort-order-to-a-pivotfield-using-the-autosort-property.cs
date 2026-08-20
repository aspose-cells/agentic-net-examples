// Title: Apply Descending AutoSort to a PivotField with Aspose.Cells in C#
// Description: Creates a workbook, adds sample data, builds a pivot table, places "Category" in rows and "Amount" in values, then enables AutoSort, sets IsAscendSort = false, selects the data field for sorting, calculates the pivot and saves the file.
// Keywords: Aspose.Cells | C# pivot table | PivotField AutoSort | descending sort pivot | .NET Excel automation | custom pivot sort order | IsAscendSort false | AutoSortField example
// Common Searches: Aspose.Cells descending sort for PivotField | How to use AutoSort in Aspose.Cells C# | Set row field sort order in Aspose.Cells pivot table | C# code for custom pivot table sorting | PivotTable AutoSortField usage Aspose.Cells
// Developer Intent: Programmatically configure a PivotTable row field to sort its items in descending order based on a data field using Aspose.Cells for .NET.
// Use Cases: Generate a report workbook where pivot rows are ordered from highest to lowest values automatically. | Update an existing pivot table in code to enforce descending sorting without opening Excel. | Create reusable C# utilities that apply custom AutoSort settings to any PivotField in Aspose.Cells.
// AI Prompts: Write C# code with Aspose.Cells that creates a pivot table and applies a descending AutoSort to the row field using IsAscendSort and AutoSortField. | Explain the relationship between AutoSort, IsAscendSort, and AutoSortField when sorting PivotField rows in Aspose.Cells. | Provide step‑by‑step instructions to modify a pivot table so its row field sorts descending by the first data field using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds sample data, builds a pivot table, places "Category" in rows and "Amount" in values, then enables AutoSort, sets IsAscendSort = false, selects the data field for sorting, calculates the pivot and saves the file.
    public class PivotFieldDescendingCustomSortDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].Value = "Category";
                sheet.Cells["B1"].Value = "Amount";
                sheet.Cells["A2"].Value = "B";
                sheet.Cells["A3"].Value = "C";
                sheet.Cells["A4"].Value = "A";
                sheet.Cells["B2"].Value = 150;
                sheet.Cells["B3"].Value = 300;
                sheet.Cells["B4"].Value = 200;

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

                // Calculate the pivot table results
                pivotTable.CalculateData();

                // Save the workbook
                workbook.Save("PivotFieldDescendingCustomSortDemo_out.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            PivotFieldDescendingCustomSortDemo.Run();
        }
    }
}
