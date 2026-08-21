// Title: C# – Collapse All Row PivotItems in an Aspose.Cells PivotTable (IsDetailHidden = true)
// Description: A .NET example that creates a workbook, adds sample data, builds a PivotTable, then iterates through each PivotItem of the first row field and sets IsDetailHidden to true, collapsing all row items. The pivot is recalculated and saved as CollapsedPivotItems.xlsx.
// Keywords: Aspose.Cells | C# | .NET | PivotTable | PivotItem | IsDetailHidden | collapse row items | hide pivot details | programmatic Excel report | sample code | GitHub example | US developers | UK developers | India developers
// Common Searches: How to collapse all row items in an Aspose.Cells PivotTable using C# | Set IsDetailHidden = true for PivotItems in Aspose.Cells | Iterate PivotItems to hide details in .NET | Programmatically collapse pivot rows with Aspose.Cells | Aspose.Cells PivotTable collapse example GitHub
// Developer Intent: Programmatically hide the detail rows of every row‑field item in an Aspose.Cells PivotTable using C#.
// Use Cases: Generate a compact Excel report by collapsing row items before exporting. | Automate dashboard refreshes where pivot details should start hidden for readability. | Apply a consistent collapsed view across multiple PivotTables in a workbook.
// AI Prompts: Write C# code with Aspose.Cells that iterates over a PivotField's PivotItems and sets IsDetailHidden = true, then recalculates and saves the workbook. | Create a reusable method that collapses all items of a specified row field in a PivotTable and returns the updated Workbook object. | Explain how IsDetailHidden, RefreshData, and CalculateData interact when toggling PivotItem expansion in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace MyAsposeDemo
{
    // A .NET example that creates a workbook, adds sample data, builds a PivotTable, then iterates through each PivotItem of the first row field and sets IsDetailHidden to true, collapsing all row items. The pivot is recalculated and saved as CollapsedPivotItems.xlsx.
    class CollapsePivotItemsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Amount");
                sheet.Cells["A2"].PutValue("Fruit");
                sheet.Cells["B2"].PutValue(100);
                sheet.Cells["A3"].PutValue("Fruit");
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["A4"].PutValue("Vegetable");
                sheet.Cells["B4"].PutValue(200);
                sheet.Cells["A5"].PutValue("Vegetable");
                sheet.Cells["B5"].PutValue(250);

                // Add a pivot table to the worksheet
                int ptIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[ptIndex];

                // Define row and data fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Refresh data and calculate the pivot table
                pivotTable.RefreshData();      // Correct API call
                pivotTable.CalculateData();

                // Collapse (hide detail) for each PivotItem in the row field
                PivotField rowField = pivotTable.RowFields[0];
                foreach (PivotItem item in rowField.PivotItems)
                {
                    item.IsDetailHidden = true;
                }

                // Recalculate after modifying item states
                pivotTable.CalculateData();

                // Save the workbook with collapsed pivot items
                string outputPath = "CollapsedPivotItems.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            CollapsePivotItemsDemo.Run();
        }
    }
}
