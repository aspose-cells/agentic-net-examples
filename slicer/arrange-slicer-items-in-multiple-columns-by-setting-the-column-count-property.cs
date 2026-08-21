// Title: Aspose.Cells for .NET: Set Slicer NumberOfColumns to Arrange Items in Multiple Columns
// Description: This C# example creates a workbook, builds a pivot table from sample data, adds a slicer linked to the "Category" field, and uses the slicer’s NumberOfColumns property to display items across three columns before saving the file.
// Keywords: Aspose.Cells | .NET | C# | Excel slicer | NumberOfColumns | multiple columns | pivot table slicer | slicer layout | dashboard design | Excel automation
// Common Searches: Aspose.Cells set slicer column count | NumberOfColumns property C# example | display slicer items in several columns | multi‑column slicer layout Aspose.Cells | how to arrange slicer items horizontally in .NET
// Developer Intent: Configure a slicer’s NumberOfColumns property to show its items in a grid rather than a single vertical list.
// Use Cases: Compact dashboard where slicer items are shown in a three‑column grid to save vertical space. | Excel reports that need a readable slicer layout for many category values. | Dynamic reports that adjust slicer column count based on the number of distinct pivot field entries.
// AI Prompts: Generate a C# snippet that sets the slicer column count based on the distinct values of a pivot field using Aspose.Cells. | Show how to create multiple slicers with different NumberOfColumns settings in the same workbook. | Explain how to apply custom styles to slicer items after arranging them in multiple columns with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace SlicerColumnsDemo
{
    // This C# example creates a workbook, builds a pivot table from sample data, adds a slicer linked to the "Category" field, and uses the slicer’s NumberOfColumns property to display items across three columns before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data for a pivot table
                cells["A1"].Value = "Category";
                cells["B1"].Value = "Product";
                cells["A2"].Value = "Fruits";
                cells["B2"].Value = "Apple";
                cells["A3"].Value = "Fruits";
                cells["B3"].Value = "Banana";
                cells["A4"].Value = "Vegetables";
                cells["B4"].Value = "Carrot";

                // Add a pivot table based on the sample data
                int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "DemoPivot");
                PivotTable pivot = sheet.PivotTables[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Data, "Product");
                pivot.RefreshData();
                pivot.CalculateData();

                // Add a slicer linked to the pivot table field "Category"
                // Note: The correct parameter order is (pivot, destination cell, base field name)
                int slicerIdx = sheet.Slicers.Add(pivot, "E1", "Category");
                Slicer slicer = sheet.Slicers[slicerIdx];

                // Set slicer properties (size, caption, etc.)
                slicer.Caption = "Category Slicer";
                slicer.TopPixel = 50;
                slicer.LeftPixel = 50;
                slicer.HeightPixel = 150;
                slicer.WidthPixel = 200;

                // Arrange slicer items in multiple columns (e.g., 3 columns)
                slicer.NumberOfColumns = 3;

                // Save the workbook
                string outputPath = "SlicerMultipleColumns.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
