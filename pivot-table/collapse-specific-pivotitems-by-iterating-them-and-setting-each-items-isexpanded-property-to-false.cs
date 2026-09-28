// Title: How to collapse specific row items in an Aspose.Cells pivot table using C# by setting IsDetailHidden
// AI Prompts: Generate C# code that creates a pivot table with Aspose.Cells, refreshes the data, and collapses each row item by setting PivotItem.IsDetailHidden = true. | Show a C# snippet that iterates over a PivotField's PivotItems in Aspose.Cells and hides their details to collapse the items. | Explain how to programmatically collapse selected pivot items after refreshing the pivot cache in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# collapse row items in pivot table programmatically | Set IsDetailHidden for all PivotItems in Aspose.Cells example | How to hide details of pivot row fields using Aspose.Cells .NET | Iterate PivotItems to collapse them in Aspose.Cells workbook
// Tags: collapse pivot row items Aspose.Cells C# | PivotItem.IsDetailHidden Aspose.Cells | iterate PivotField items Aspose.Cells | refresh pivot cache Aspose.Cells before collapsing | Aspose.Cells pivot table row field manipulation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // The example creates a workbook, adds sample data, builds a pivot table, refreshes it, then iterates the row field's PivotItems and sets IsDetailHidden to true to collapse each item before saving the file.
    public class CollapseSpecificPivotItems
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

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add the "Category" field as a row field and "Amount" as a data field
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Refresh the pivot cache and calculate the pivot table so that items are generated
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Collapse each row item by hiding its detail
                PivotField rowField = pivotTable.RowFields[0];
                foreach (PivotItem item in rowField.PivotItems)
                {
                    item.IsDetailHidden = true;
                }

                // Recalculate after changing the collapse state
                pivotTable.CalculateData();

                // Save the workbook
                string outputPath = "CollapsedPivotItems.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            CollapseSpecificPivotItems.Run();
        }
    }
}
