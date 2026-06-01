using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace SlicerMultiColumnDemo
{
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

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Product");
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Add a slicer linked to the pivot table field "Category"
                // Correct argument order: destination cell name first, then field name
                int slicerIndex = sheet.Slicers.Add(pivotTable, "E1", "Category");
                Slicer slicer = sheet.Slicers[slicerIndex];

                // Set slicer properties
                slicer.Caption = "Product Categories";
                slicer.TopPixel = 50;
                slicer.LeftPixel = 50;
                slicer.HeightPixel = 150;
                slicer.WidthPixel = 200;

                // Arrange slicer items in multiple columns (e.g., 3 columns)
                slicer.NumberOfColumns = 3;

                // Define output file path
                string outputPath = "SlicerMultiColumnDemo.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}