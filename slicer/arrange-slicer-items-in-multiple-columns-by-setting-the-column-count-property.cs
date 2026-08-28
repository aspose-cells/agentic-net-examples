// Title: Arrange slicer items into multiple columns with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a pivot table, adds a slicer for a field, and sets the slicer’s NumberOfColumns property to 3 using Aspose.Cells. | Show how to customize a slicer’s caption, size, and column layout programmatically in an Aspose.Cells workbook. | Provide a step‑by‑step example of saving an Excel file after configuring a slicer to display items in three columns with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# set slicer NumberOfColumns to display items in multiple columns | how to configure slicer layout columns in an Excel workbook using Aspose.Cells for .NET | C# example of adding a slicer to a pivot table and arranging items in three columns with Aspose.Cells | Aspose.Cells slicer column count property usage in .NET | programmatically change slicer item arrangement to multiple columns in C# Excel library
// Tags: Aspose.Cells slicer NumberOfColumns property | C# configure slicer multiple columns | Aspose.Cells pivot table slicer layout | Excel slicer column arrangement Aspose.Cells | Aspose.Cells workbook slicer customization

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace SlicerMultipleColumnsDemo
{
    // The sample creates a workbook, fills it with sample data, builds a pivot table, adds a slicer linked to the 'Category' field, sets the slicer's caption and dimensions, configures the slicer to show items in three columns via the NumberOfColumns property, and saves the result as an .xlsx file.
    class Program
    {
        static void Main()
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
                int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
                PivotTable pivot = sheet.PivotTables[pivotIndex];
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Data, "Product");
                pivot.RefreshData();
                pivot.CalculateData();

                // Add a slicer linked to the pivot table field "Category"
                // Correct argument order: destination cell, then field name
                int slicerIndex = sheet.Slicers.Add(pivot, "E1", "Category");
                Slicer slicer = sheet.Slicers[slicerIndex];

                // Set slicer properties (size, caption, etc.)
                slicer.Caption = "Product Categories";
                slicer.TopPixel = 50;
                slicer.LeftPixel = 50;
                slicer.HeightPixel = 150;
                slicer.WidthPixel = 200;

                // Arrange slicer items in multiple columns (e.g., 3 columns)
                slicer.NumberOfColumns = 3;

                // Save the workbook to a file
                workbook.Save("SlicerMultipleColumnsDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}
